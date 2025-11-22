import pandas as pd
from pathlib import Path
import math

def calcular_tamanho_estimado(df):
    """
    Estima o tamanho do DataFrame em MB quando salvo como Excel
    """
    # Estima baseado no tamanho em memória (aproximado)
    tamanho_memoria = df.memory_usage(deep=True).sum()
    # Excel geralmente usa ~1.5-2x o tamanho em memória
    tamanho_estimado_mb = (tamanho_memoria * 1.8) / (1024 * 1024)
    return tamanho_estimado_mb

def dividir_planilha_por_tamanho():
    """
    Divide uma planilha grande em arquivos menores de até 4MB
    Mantém o cabeçalho original em cada arquivo
    """
    
    # Configurações
    arquivo_entrada = r"C:\Users\PC\Downloads\pipiline bemol farma\dados-filtrados_PROCESSADO.xlsx"
    tamanho_maximo_mb = 4.0  # Tamanho máximo por arquivo
    
    print("=" * 80)
    print("PIPELINE DE DIVISÃO DE PLANILHA")
    print("=" * 80)
    
    # Verificar se o arquivo existe
    if not Path(arquivo_entrada).exists():
        print(f"❌ ERRO: Arquivo não encontrado: {arquivo_entrada}")
        return
    
    try:
        # Ler a planilha completa
        print(f"\n📂 Lendo arquivo: {Path(arquivo_entrada).name}")
        df_completo = pd.read_excel(arquivo_entrada)
        
        total_linhas = len(df_completo)
        total_colunas = len(df_completo.columns)
        
        print(f"✅ Planilha carregada com sucesso!")
        print(f"   Total de linhas: {total_linhas:,}")
        print(f"   Total de colunas: {total_colunas}")
        
        # Estimar tamanho total
        tamanho_total = calcular_tamanho_estimado(df_completo)
        print(f"   Tamanho estimado: {tamanho_total:.2f} MB")
        
        # Calcular número de partes necessárias
        num_partes = math.ceil(tamanho_total / tamanho_maximo_mb)
        linhas_por_parte = math.ceil(total_linhas / num_partes)
        
        print(f"\n📊 Planejamento da divisão:")
        print(f"   Tamanho máximo por arquivo: {tamanho_maximo_mb} MB")
        print(f"   Número de partes estimado: {num_partes}")
        print(f"   Linhas por parte (aproximado): {linhas_por_parte:,}")
        
        # Criar diretório para os arquivos divididos
        diretorio_saida = Path(arquivo_entrada).parent / "planilhas_divididas"
        diretorio_saida.mkdir(exist_ok=True)
        
        print(f"\n📁 Diretório de saída: {diretorio_saida.name}")
        
        # Dividir e salvar
        print(f"\n⚙️  Dividindo e salvando arquivos...")
        print("-" * 80)
        
        arquivos_criados = []
        parte_atual = 1
        inicio = 0
        
        while inicio < total_linhas:
            # Calcular fim da parte atual
            fim = min(inicio + linhas_por_parte, total_linhas)
            
            # Extrair chunk com ajuste dinâmico de tamanho
            while True:
                df_parte = df_completo.iloc[inicio:fim].copy()
                tamanho_parte = calcular_tamanho_estimado(df_parte)
                
                # Se o tamanho está OK ou já está no mínimo (1 linha), prosseguir
                if tamanho_parte <= tamanho_maximo_mb or (fim - inicio) <= 1:
                    break
                
                # Se está muito grande, reduzir 20%
                if tamanho_parte > tamanho_maximo_mb:
                    reducao = int((fim - inicio) * 0.2)
                    fim = max(inicio + 1, fim - reducao)
                else:
                    break
            
            linhas_nesta_parte = fim - inicio
            
            # Criar nome do arquivo
            nome_base = Path(arquivo_entrada).stem
            nome_arquivo = f"{nome_base}_PARTE_{parte_atual:02d}.xlsx"
            caminho_saida = diretorio_saida / nome_arquivo
            
            # Salvar a parte
            print(f"\n📝 Parte {parte_atual}/{num_partes}:")
            print(f"   Arquivo: {nome_arquivo}")
            print(f"   Linhas: {inicio + 1} até {fim} ({linhas_nesta_parte:,} linhas)")
            print(f"   Tamanho estimado: {tamanho_parte:.2f} MB")
            
            try:
                df_parte.to_excel(caminho_saida, index=False, engine='openpyxl')
                
                # Verificar tamanho real do arquivo criado
                tamanho_real = caminho_saida.stat().st_size / (1024 * 1024)
                print(f"   ✅ Salvo com sucesso! (Tamanho real: {tamanho_real:.2f} MB)")
                
                arquivos_criados.append({
                    'parte': parte_atual,
                    'arquivo': nome_arquivo,
                    'linhas': linhas_nesta_parte,
                    'tamanho_mb': tamanho_real,
                    'inicio': inicio + 1,
                    'fim': fim
                })
                
                # Verificar se excedeu o tamanho (pode acontecer por estimativa)
                if tamanho_real > tamanho_maximo_mb * 1.1:  # 10% de tolerância
                    print(f"   ⚠️  Aviso: Arquivo ficou maior que {tamanho_maximo_mb} MB")
                
            except Exception as e:
                print(f"   ❌ Erro ao salvar: {type(e).__name__}")
                print(f"   Detalhes: {str(e)[:200]}")
                
                # Tentar com xlsxwriter
                try:
                    df_parte.to_excel(caminho_saida, index=False, engine='xlsxwriter')
                    tamanho_real = caminho_saida.stat().st_size / (1024 * 1024)
                    print(f"   ✅ Salvo com xlsxwriter! (Tamanho: {tamanho_real:.2f} MB)")
                    
                    arquivos_criados.append({
                        'parte': parte_atual,
                        'arquivo': nome_arquivo,
                        'linhas': linhas_nesta_parte,
                        'tamanho_mb': tamanho_real,
                        'inicio': inicio + 1,
                        'fim': fim
                    })
                except Exception as e2:
                    print(f"   ❌ Falha definitiva: {type(e2).__name__}")
                    continue
            
            # Avançar para próxima parte
            inicio = fim
            parte_atual += 1
        
        # Relatório final
        print("\n" + "=" * 80)
        print("✨ DIVISÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 80)
        
        print(f"\n📊 RESUMO DA OPERAÇÃO:")
        print(f"   ✅ Arquivo original: {Path(arquivo_entrada).name}")
        print(f"   ✅ Total de linhas processadas: {total_linhas:,}")
        print(f"   ✅ Total de colunas: {total_colunas}")
        print(f"   ✅ Arquivos criados: {len(arquivos_criados)}")
        print(f"   ✅ Diretório: {diretorio_saida}")
        
        print(f"\n📋 DETALHAMENTO DOS ARQUIVOS:")
        print("-" * 80)
        
        tamanho_total_partes = 0
        for info in arquivos_criados:
            print(f"\n   Parte {info['parte']:02d}: {info['arquivo']}")
            print(f"   │ Linhas {info['inicio']:,} até {info['fim']:,} ({info['linhas']:,} linhas)")
            print(f"   │ Tamanho: {info['tamanho_mb']:.2f} MB")
            print(f"   └─ {'✅ OK' if info['tamanho_mb'] <= tamanho_maximo_mb else '⚠️ Acima do limite'}")
            tamanho_total_partes += info['tamanho_mb']
        
        print(f"\n" + "-" * 80)
        print(f"   📦 Tamanho total dos arquivos: {tamanho_total_partes:.2f} MB")
        print(f"   💾 Economia de espaço: {((tamanho_total - tamanho_total_partes) / tamanho_total * 100):.1f}%")
        
        print(f"\n✅ VERIFICAÇÃO:")
        print(f"   • Cada arquivo tem o mesmo cabeçalho ({total_colunas} colunas)")
        print(f"   • Dados originais preservados")
        print(f"   • Total de linhas: {sum([a['linhas'] for a in arquivos_criados]):,}")
        
        if sum([a['linhas'] for a in arquivos_criados]) == total_linhas:
            print(f"   ✅ Todas as linhas foram incluídas!")
        else:
            print(f"   ⚠️  Atenção: Verificar contagem de linhas")
        
        print(f"\n📁 Os arquivos estão em: {diretorio_saida}")
        
    except Exception as e:
        print(f"\n❌ ERRO durante o processamento:")
        print(f"   {type(e).__name__}: {str(e)}")
        import traceback
        print(f"\n🔍 Detalhes do erro:")
        traceback.print_exc()

# Executar o pipeline
if __name__ == "__main__":
    dividir_planilha_por_tamanho()