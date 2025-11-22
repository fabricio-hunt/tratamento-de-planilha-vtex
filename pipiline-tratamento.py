import pandas as pd
import openpyxl
from pathlib import Path
import re

def limpar_caracteres_ilegais_excel(texto):
    """
    Remove APENAS caracteres que impedem o salvamento no Excel
    Preserva TODO o conteúdo válido (HTML, tags, etc)
    """
    if pd.isna(texto) or texto == "":
        return texto
    
    texto = str(texto)
    
    # Lista de caracteres de controle ilegais no Excel (0x00-0x1F exceto tab, newline, return)
    # Estes são os únicos que precisam ser removidos
    texto_limpo = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', texto)
    
    # Remove símbolos Unicode específicos que causam problemas
    caracteres_problematicos = ['♥', '♦', '♣', '♠', '★', '☆', '♡', '♤', '◆', '◇', '○', '●', '□', '■', '△', '▲', '▽', '▼']
    for char in caracteres_problematicos:
        texto_limpo = texto_limpo.replace(char, '')
    
    return texto_limpo

def processar_planilha_metatag():
    """
    Pipeline para tratamento da coluna _DescricaoMetaTag
    MODIFICA: Apenas coluna AE (conteúdo novo)
    LIMPA: Todas as colunas (remove apenas caracteres ilegais do Excel)
    PRESERVA: Todo o conteúdo válido (HTML, dados de e-commerce, etc)
    """
    
    # Configurações
    arquivo_entrada = r"C:\Users\PC\Downloads\pipiline bemol farma\dados-filtrados.xls"
    
    # Texto padrão para concatenação
    texto_padrao = " com os melhores preços e Produtos de Higiene Pessoal, Dermocosméticos e mais. Aproveite as vantagens e compre online!"
    
    print("=" * 80)
    print("PIPELINE DE TRATAMENTO - DESCRIÇÃO META TAG")
    print("=" * 80)
    
    # Verificar se o arquivo existe
    if not Path(arquivo_entrada).exists():
        print(f"❌ ERRO: Arquivo não encontrado: {arquivo_entrada}")
        return
    
    try:
        # Ler a planilha
        print(f"\n📂 Lendo arquivo: {Path(arquivo_entrada).name}")
        df = pd.read_excel(arquivo_entrada)
        
        print(f"✅ Planilha carregada com sucesso!")
        print(f"   Total de linhas: {len(df)}")
        print(f"   Total de colunas: {len(df.columns)}")
        
        # Identificar as colunas U e AE (índices 20 e 30)
        coluna_u_index = 20  # Coluna U (21ª coluna)
        coluna_ae_index = 30  # Coluna AE (31ª coluna)
        
        # Verificar se as colunas existem
        if len(df.columns) <= coluna_ae_index:
            print(f"❌ ERRO: A planilha não possui colunas suficientes.")
            print(f"   Colunas encontradas: {len(df.columns)}")
            print(f"   Necessário: pelo menos {coluna_ae_index + 1} colunas")
            return
        
        # Obter nomes das colunas
        nome_coluna_u = df.columns[coluna_u_index]
        nome_coluna_ae = df.columns[coluna_ae_index]
        
        print(f"\n📊 Colunas identificadas:")
        print(f"   Coluna U (índice {coluna_u_index}): '{nome_coluna_u}'")
        print(f"   Coluna AE (índice {coluna_ae_index}): '{nome_coluna_ae}'")
        
        # Contar registros antes do processamento
        total_registros = len(df)
        registros_com_nome = df[nome_coluna_u].notna().sum()
        
        print(f"\n📈 Análise dos dados:")
        print(f"   Total de registros: {total_registros}")
        print(f"   Registros com Nome do Produto: {registros_com_nome}")
        print(f"   Registros sem Nome do Produto: {total_registros - registros_com_nome}")
        
        # ETAPA 1: PROCESSAR APENAS A COLUNA AE
        print(f"\n⚙️  ETAPA 1: Processando coluna AE (_DescricaoMetaTag)...")
        
        # Criar a nova descrição meta tag APENAS na coluna AE
        df[nome_coluna_ae] = df.apply(
            lambda row: f"{row[nome_coluna_u]}{texto_padrao}" 
            if pd.notna(row[nome_coluna_u]) and str(row[nome_coluna_u]).strip() != "" 
            else "",
            axis=1
        )
        
        registros_processados = (df[nome_coluna_ae] != "").sum()
        print(f"✅ Coluna AE processada!")
        print(f"   Novos registros criados: {registros_processados}")
        
        # ETAPA 2: LIMPAR CARACTERES ILEGAIS DE TODAS AS COLUNAS
        print(f"\n🧹 ETAPA 2: Limpando caracteres ilegais de TODAS as colunas...")
        print(f"   (Remove apenas caracteres que impedem salvamento no Excel)")
        print(f"   (HTML, tags e dados válidos serão PRESERVADOS)")
        
        total_celulas_limpas = 0
        colunas_com_limpeza = []
        
        for idx, col in enumerate(df.columns):
            # Aplicar limpeza em colunas de texto
            if df[col].dtype == 'object':
                valores_antes = df[col].astype(str).copy()
                df[col] = df[col].apply(limpar_caracteres_ilegais_excel)
                
                # Contar células modificadas
                celulas_modificadas = (valores_antes != df[col].astype(str)).sum()
                
                if celulas_modificadas > 0:
                    total_celulas_limpas += celulas_modificadas
                    colunas_com_limpeza.append((col, celulas_modificadas))
                    
                    # Mostrar progresso para colunas problemáticas
                    if celulas_modificadas > 100:
                        letra_coluna = openpyxl.utils.get_column_letter(idx + 1)
                        print(f"   └─ Coluna {letra_coluna} ({col}): {celulas_modificadas} células limpas")
        
        print(f"\n✅ Limpeza concluída!")
        print(f"   Total de colunas processadas: {len([c for c in df.columns if df[c].dtype == 'object'])}")
        print(f"   Colunas com caracteres removidos: {len(colunas_com_limpeza)}")
        print(f"   Total de células limpas: {total_celulas_limpas}")
        
        if len(colunas_com_limpeza) > 0:
            print(f"\n📋 Top 5 colunas com mais limpezas:")
            for col_nome, qtd in sorted(colunas_com_limpeza, key=lambda x: x[1], reverse=True)[:5]:
                print(f"   • {col_nome}: {qtd} células")
        
        # ETAPA 3: SALVAR ARQUIVO
        arquivo_saida = arquivo_entrada.replace(".xls", "_PROCESSADO.xlsx")
        
        print(f"\n💾 ETAPA 3: Salvando arquivo processado...")
        print(f"   Destino: {Path(arquivo_saida).name}")
        
        # Verificar se arquivo de saída já existe e está aberto
        if Path(arquivo_saida).exists():
            print(f"   ⚠️  Arquivo já existe, será sobrescrito")
            try:
                # Tentar deletar para verificar se está aberto
                Path(arquivo_saida).unlink()
                print(f"   ✅ Arquivo anterior removido")
            except PermissionError:
                # Arquivo está aberto, criar com nome alternativo
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                arquivo_saida = arquivo_entrada.replace(".xls", f"_PROCESSADO_{timestamp}.xlsx")
                print(f"   ⚠️  Arquivo anterior está aberto!")
                print(f"   📝 Salvando com novo nome: {Path(arquivo_saida).name}")
        
        try:
            # Tentar salvar com openpyxl
            df.to_excel(arquivo_saida, index=False, engine='openpyxl')
            print(f"✅ Arquivo salvo com sucesso!")
            
        except PermissionError as e_perm:
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            arquivo_saida = arquivo_entrada.replace(".xls", f"_PROCESSADO_{timestamp}.xlsx")
            print(f"⚠️  Arquivo está aberto ou sem permissão")
            print(f"   Salvando com timestamp: {Path(arquivo_saida).name}")
            
            try:
                df.to_excel(arquivo_saida, index=False, engine='openpyxl')
                print(f"✅ Arquivo salvo com nome alternativo!")
            except Exception as e_retry:
                print(f"❌ Erro ao salvar: {type(e_retry).__name__}")
                print(f"   Detalhes: {str(e_retry)[:200]}")
                return
                
        except Exception as e_save:
            print(f"⚠️  Erro com openpyxl: {type(e_save).__name__}")
            
            # Verificar se xlsxwriter está disponível
            try:
                import xlsxwriter
                print(f"   Tentando com xlsxwriter...")
                
                try:
                    df.to_excel(arquivo_saida, index=False, engine='xlsxwriter')
                    print(f"✅ Arquivo salvo com xlsxwriter!")
                    
                except Exception as e_xlsx:
                    print(f"⚠️  Erro com xlsxwriter: {type(e_xlsx).__name__}")
                    print(f"   Detalhes: {str(e_xlsx)[:200]}")
                    return
                    
            except ImportError:
                print(f"❌ xlsxwriter não está instalado")
                print(f"\n💡 SOLUÇÃO:")
                print(f"   Execute: pip install xlsxwriter")
                print(f"   Depois execute o script novamente")
                return
        
        # ETAPA 4: VERIFICAÇÕES E RELATÓRIOS
        print(f"\n🔍 ETAPA 4: Verificações finais...")
        
        # Verificar se o arquivo foi criado
        if Path(arquivo_saida).exists():
            tamanho_mb = Path(arquivo_saida).stat().st_size / (1024 * 1024)
            print(f"   ✅ Arquivo criado: {tamanho_mb:.2f} MB")
        else:
            print(f"   ❌ Arquivo não foi criado")
            return
        
        # Verificar colunas vazias
        colunas_vazias = []
        for col in df.columns:
            if df[col].isna().all() or (df[col].astype(str) == "").all():
                colunas_vazias.append(col)
        
        if colunas_vazias:
            print(f"   ⚠️  Colunas vazias encontradas: {len(colunas_vazias)}")
            for col in colunas_vazias[:3]:
                print(f"      • {col}")
        else:
            print(f"   ✅ Todas as colunas contêm dados")
        
        # Exemplo de registros processados
        print(f"\n📋 Exemplo de registros processados (primeiras 3 linhas):")
        print("-" * 80)
        for idx, row in df.head(3).iterrows():
            if pd.notna(row[nome_coluna_u]):
                print(f"\nLinha {idx + 2}:")
                print(f"  Nome Produto: {str(row[nome_coluna_u])[:50]}...")
                meta_tag = str(row[nome_coluna_ae])[:80]
                print(f"  Nova Meta Tag: {meta_tag}...")
        
        print("\n" + "=" * 80)
        print("✨ PIPELINE CONCLUÍDO COM SUCESSO!")
        print("=" * 80)
        print(f"\n📊 RESUMO DA OPERAÇÃO:")
        print(f"   ✅ Arquivo original: PRESERVADO")
        print(f"   ✅ Novo arquivo: {Path(arquivo_saida).name}")
        print(f"   ✅ Total de colunas: {len(df.columns)}")
        print(f"   ✅ Total de linhas: {len(df)}")
        print(f"")
        print(f"   📝 MODIFICAÇÕES REALIZADAS:")
        print(f"   • Coluna AE: {registros_processados} registros atualizados")
        print(f"   • Limpeza geral: {total_celulas_limpas} células com caracteres ilegais removidos")
        print(f"   • Demais dados: PRESERVADOS (HTML, preços, SKUs, etc)")
        print(f"")
        print(f"   ⚠️  IMPORTANTE:")
        print(f"   • Apenas caracteres ilegais do Excel foram removidos")
        print(f"   • HTML e tags foram preservados")
        print(f"   • Dados de e-commerce permanecem intactos")
        print(f"   • Revise o arquivo antes de importar")
        
    except Exception as e:
        print(f"\n❌ ERRO durante o processamento:")
        print(f"   {type(e).__name__}: {str(e)}")
        import traceback
        print(f"\n🔍 Detalhes do erro:")
        traceback.print_exc()

# Executar o pipeline
if __name__ == "__main__":
    processar_planilha_metatag()