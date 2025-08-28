import pandas as pd
import streamlit as st
import traceback

# ... (o resto das suas funções como converter_valor, formatar_valor, etc. permanecem no topo)


def processar_planilha(file):
    try:
        xl = pd.ExcelFile(file)
        dados_processados = []

        for sheet_name in xl.sheet_names:
            # header=None é crucial aqui
            sheet_df = pd.read_excel(xl, sheet_name=sheet_name, header=None)

            conta_contabil_atual = "Não Identificado"
            descricao_conta_atual = "Não Identificado"
            item_temporario = None

            for idx, row in sheet_df.iterrows():
                # Pula linhas completamente vazias
                if row.isnull().all():
                    continue

                # Pega o valor da primeira célula para verificações
                primeira_celula = str(row.iloc[0]).strip()

                # IGNORAR LINHAS DE TOTALIZAÇÃO PARA NÃO DUPLICAR DADOS
                if primeira_celula.startswith('Filial :') or primeira_celula.startswith('* * *'):
                    item_temporario = None  # Reseta qualquer item pendente
                    continue

                # 1. IDENTIFICAR A LINHA DA CONTA CONTÁBIL
                if primeira_celula.startswith('1.2.3.'):
                    partes = primeira_celula.split(' ', 1)
                    conta_contabil_atual = partes[0]
                    descricao_conta_atual = partes[1].strip() if len(
                        partes) > 1 else "Sem Descrição"
                    item_temporario = None  # Reseta ao encontrar nova conta
                    continue

                # 2. IDENTIFICAR A LINHA PRINCIPAL DO ITEM INDIVIDUAL
                # A condição chave: coluna A é filial, coluna H (índice 7) é um número (a data)
                if len(row) > 7 and pd.api.types.is_numeric_dtype(row.iloc[0]) and pd.api.types.is_numeric_dtype(row.iloc[7]):

                    # Converte o número de série do Excel para data
                    # O 'coerce' trata erros se o número não for uma data válida
                    data_aquisicao = pd.to_datetime(
                        row.iloc[7], unit='d', origin='1899-12-30', errors='coerce')

                    # Se a conversão da data funcionou, temos um item válido
                    if pd.notna(data_aquisicao):
                        item_temporario = {
                            'Arquivo': file.name,
                            'Filial': str(int(row.iloc[0])).strip(),
                            'Conta contábil': conta_contabil_atual,
                            'Descrição da conta': descricao_conta_atual,
                            'Data de aquisição': data_aquisicao,
                            'Código do item': str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "",
                            'Descrição do item': str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else "",
                            # Adicionando colunas que podem estar faltando no seu layout final
                            'Código do sub item': ""
                        }
                        # Passa para a próxima linha para procurar os valores
                        continue

                # 3. IDENTIFICAR A LINHA DE VALORES DO ITEM
                # Esta linha deve vir logo após a linha do item
                if item_temporario and primeira_celula == 'R$':
                    # Preenche os valores no item que acabamos de encontrar
                    item_temporario['Valor original'] = converter_valor(
                        row.iloc[2])
                    item_temporario['Valor atualizado'] = converter_valor(
                        row.iloc[3])
                    item_temporario['Deprec. do mês'] = converter_valor(
                        row.iloc[4])
                    item_temporario['Deprec. do exercício'] = converter_valor(
                        row.iloc[5])
                    item_temporario['Deprec. Acumulada'] = converter_valor(
                        row.iloc[6])

                    dados_processados.append(item_temporario)

                    # Reseta para garantir que não será usado novamente
                    item_temporario = None

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            df_final['Valor Residual'] = df_final['Valor atualizado'] - \
                df_final['Deprec. Acumulada']

            # Use a lista de colunas completa novamente
            colunas_ordenadas = [
                'Filial', 'Conta contábil', 'Descrição da conta', 'Data de aquisição',
                'Código do item', 'Código do sub item', 'Descrição do item',
                'Valor original', 'Valor atualizado', 'Deprec. do mês',
                'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual', 'Arquivo'
            ]
            df_final = df_final.reindex(columns=colunas_ordenadas)
            return df_final, None

        return pd.DataFrame(), f"Nenhum item individual foi encontrado no formato esperado em '{file.name}'."
    except Exception as e:
        st.error(f"Erro crítico ao processar o arquivo '{file.name}': {e}")
        st.code(traceback.format_exc())
        return pd.DataFrame(), f"Erro crítico ao processar '{file.name}'."
