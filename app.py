import pandas as pd

ARQUIVO_ENTRADA = "Imobilizado_Saneamento 1.xlsx"
ARQUIVO_SAIDA = "ativos_formatado.xlsx"

CABECALHOS_INTERMEDIARIOS = ["Cod Base Bem", "Descr. Sint.",
                             "Dt.Aquisicao", "Data Baixa", "Quantidade", "Num.Plaqueta"]


def limpar_valor(valor_str):
    """Converte valor monetário em string para float."""
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except (ValueError, AttributeError):
        return None


def processar_linhas(linhas):
    """Processa as linhas extraídas do Excel."""
    dados = []
    conta = ""
    descricao_conta = ""

    for linha in linhas:
        linha = linha.strip()
        if any(cabecalho in linha for cabecalho in CABECALHOS_INTERMEDIARIOS):
            continue  # pula cabeçalhos intermediários

        colunas = linha.split('\t')

        if len(colunas) == 2 and "." in colunas[0]:
            conta, descricao_conta = colunas

        elif colunas[0] == "R$":
            if dados:
                try:
                    valor_original = limpar_valor(colunas[2])
                    valor_atualizado = limpar_valor(colunas[3])
                    deprec_mes = limpar_valor(colunas[4])
                    deprec_acumulada = limpar_valor(colunas[6])
                    valor_contabil = None

                    if valor_atualizado is not None and deprec_acumulada is not None:
                        valor_contabil = valor_atualizado - deprec_acumulada

                    dados[-1].update({
                        "Valor original": valor_original,
                        "Valor atualizado": valor_atualizado,
                        "Deprec. Mês": deprec_mes,
                        "Deprec acumulada": deprec_acumulada,
                        "Valor contábil": valor_contabil
                    })
                except IndexError:
                    continue

        elif len(colunas) >= 13 and colunas[0] not in ["R$", "US$", "UFIR"]:
            dados.append({
                "Conta": conta,
                "Descrição conta": descricao_conta,
                "Cód. Base bem": colunas[2],
                "Desc. Sintética": colunas[5][:30],
                "Dt. aquisição": colunas[7],
                "Dt. Baixa": colunas[8],
                "Quantidade": colunas[9],
                "Nº placa": colunas[10],
            })

    return dados


def main():
    try:

        df_raw = pd.read_excel(ARQUIVO_ENTRADA, header=None, dtype=str)
        linhas = df_raw.fillna("").astype(str).agg('\t'.join, axis=1).tolist()

        dados_processados = processar_linhas(linhas)

        df_final = pd.DataFrame(dados_processados)
        df_final.to_excel(ARQUIVO_SAIDA, index=False)

        print(f"✅ Arquivo salvo como: {ARQUIVO_SAIDA}")

    except Exception as e:
        print(f"❌ Erro ao processar o arquivo: {e}")


if __name__ == "__main__":
    main()
