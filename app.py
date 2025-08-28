import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
from fpdf import FPDF
import traceback

# Ignorar avisos para uma interface mais limpa
warnings.filterwarnings('ignore')

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard de Ativos Contábeis",
    page_icon="📊",
    layout="wide"
)

# --- INICIALIZAÇÃO DO SESSION STATE ---
# Guarda a figura Plotly para exportação
if 'figura_plotly' not in st.session_state:
    st.session_state.figura_plotly = None

# --- CLASSE FPDF CUSTOMIZADA PARA SUPORTE A UTF-8 ---


class PDF(FPDF):
    def header(self):
        try:
            self.image("logo_GW.png", x=10, y=8, w=40)
        except FileNotFoundError:
            self.set_font("Arial", "B", 12)
            self.cell(40, 10, "General Water", 0, 0, 'L')
        self.set_font("Arial", "B", 20)
        self.cell(0, 10, "Relatório de Ativos Contábeis", 0, 1, 'C')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# --- FUNÇÕES DE LÓGICA ---


def converter_valor(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        # Converte para string, remove "R$", espaços e ajusta separadores decimais
        valor_str = str(valor).replace('R$', '').strip()
        if ',' in valor_str and '.' in valor_str:
            # Remove separador de milhar
            valor_str = valor_str.replace('.', '')
        # Converte vírgula em ponto decimal
        valor_str = valor_str.replace(',', '.')
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0


def formatar_valor(valor):
    try:
        # Formata o valor como moeda brasileira (BRL)
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"


def processar_planilha(file):
    try:
        xl = pd.ExcelFile(file)
        dados_processados = []

        # Adiciona um expansor para os logs de depuração de cada arquivo
        with st.expander(f"Logs de Depuração para '{file.name}'", expanded=False):
            st.write(f"Iniciando processamento da aba: {xl.sheet_names[0]}...")

            for sheet_name in xl.sheet_names:
                sheet_df = pd.read_excel(
                    xl, sheet_name=sheet_name, header=None)

                conta_contabil_atual = "Não Identificado"
                descricao_conta_atual = "Não Identificado"
                item_temporario = None

                for idx, row in sheet_df.iterrows():
                    # Pula linhas completamente vazias
                    if row.isnull().all():
                        continue

                    num_colunas = len(row)

                    # Log da linha que está sendo processada
                    # st.write(f"Linha {idx+1}: {row.to_list()}") # Descomente para depuração extrema

                    # 1. Identifica a linha da Conta Contábil
                    if num_colunas > 1 and pd.notna(row.iloc[0]) and str(row.iloc[0]).startswith('1.2.3.'):
                        conta_contabil_atual = str(row.iloc[0]).strip()
                        descricao_conta_atual = str(
                            row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                        item_temporario = None
                        st.write(
                            f"✅ Linha {idx+1}: Encontrada Conta Contábil - '{conta_contabil_atual}'")
                        continue

                    # 2. Identifica a linha principal do item
                    if num_colunas > 7:
                        # >>> AQUI ESTÁ A MUDANÇA PRINCIPAL <<<
                        # Verifica se a coluna A é um número OU uma string contendo apenas dígitos.
                        coluna_a_str = str(row.iloc[0]).strip()
                        is_filial_code = False
                        if pd.notna(row.iloc[0]):
                            if isinstance(row.iloc[0], (int, float)):
                                is_filial_code = True
                            elif isinstance(row.iloc[0], str) and coluna_a_str.isdigit():
                                is_filial_code = True

                        data_aquisicao = pd.to_datetime(
                            row.iloc[7], errors='coerce')

                        if is_filial_code and pd.notna(data_aquisicao):
                            item_temporario = {
                                'Arquivo': file.name,
                                'Filial': coluna_a_str,
                                'Conta contábil': conta_contabil_atual,
                                'Descrição da conta': descricao_conta_atual,
                                'Data de aquisição': data_aquisicao,
                                'Código do item': str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
                                'Código do sub item': str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "",
                                'Descrição do item': str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
                            }
                            st.write(
                                f"✅ Linha {idx+1}: Encontrada LINHA DE ITEM. Filial: '{coluna_a_str}'. Aguardando linha de valor...")
                            continue
                        # Log para falha na detecção de item
                        elif is_filial_code and not pd.notna(data_aquisicao):
                            st.write(
                                f"⚠️ Linha {idx+1}: Potencial item, mas data na coluna H ('{row.iloc[7]}') é inválida.")

                    # 3. Identifica a linha de valores
                    if item_temporario and num_colunas > 6 and pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == 'R$':
                        st.write(
                            f"✅ Linha {idx+1}: Encontrada LINHA DE VALOR. Associando ao item anterior e salvando.")
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
                        item_temporario = None  # Reseta para o próximo item
                        continue

        if dados_processados:
            st.success(
                f"Dados extraídos com sucesso de '{file.name}'. Total de {len(dados_processados)} registros.")
            df_final = pd.DataFrame(dados_processados)
            df_final['Valor Residual'] = df_final['Valor atualizado'] - \
                df_final['Deprec. Acumulada']

            colunas_ordenadas = [
                'Filial', 'Conta contábil', 'Descrição da conta', 'Data de aquisição',
                'Código do item', 'Código do sub item', 'Descrição do item',
                'Valor original', 'Valor atualizado', 'Deprec. do mês',
                'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual', 'Arquivo'
            ]
            df_final = df_final.reindex(columns=colunas_ordenadas)
            return df_final, None

        return pd.DataFrame(), f"Nenhum dado no formato esperado foi encontrado em '{file.name}'."
    except Exception as e:
        st.error(f"Erro crítico ao processar o arquivo '{file.name}': {e}")
        st.code(traceback.format_exc())
        return pd.DataFrame(), f"Erro crítico ao processar '{file.name}'."


def criar_pdf_completo(buffer, df_filtrado, fig_plotly):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    # Adiciona fonte com suporte a UTF-8
    pdf.add_font('Arial', '', 'arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', 'arialbd.ttf', uni=True)
    pdf.set_font('Arial', '', 12)
    pdf.add_page()

    # Seção do Gráfico
    if fig_plotly:
        try:
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Gráfico Analítico", 0, 1, 'L')

            # Converte a figura Plotly para imagem em bytes
            img_bytes = fig_plotly.to_image(
                format="png", width=1000, height=500, scale=2)
            img_buffer = BytesIO(img_bytes)

            # Largura máxima da página A4 paisagem
            pdf.image(img_buffer, x=10, y=pdf.get_y(), w=277)
            pdf.ln(135)  # Pula o espaço ocupado pela imagem
        except Exception as e:
            pdf.set_font("Arial", "", 10)
            pdf.multi_cell(
                0, 10, f"Não foi possível renderizar o gráfico no PDF: {e}", 0, 'L')

    # Seção da Tabela de Dados Agregados
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Dados Agregados por Filial e Descrição da Conta", 0, 1, 'L')
    pdf.ln(5)

    colunas_para_somar = ['Valor atualizado',
                          'Deprec. Acumulada', 'Valor Residual']
    df_agregado = df_filtrado.groupby(['Filial', 'Descrição da conta'])[
        colunas_para_somar].sum().reset_index()

    # Cabeçalho da tabela
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(230, 230, 230)
    col_widths = {'Filial': 60, 'Descrição da conta': 100,
                  'Valor atualizado': 38, 'Deprec. Acumulada': 40, 'Valor Residual': 39}
    for col_name in col_widths.keys():
        pdf.cell(col_widths[col_name], 10, col_name, 1, 0, 'C', fill=True)
    pdf.ln()

    # Corpo da tabela
    pdf.set_font("Arial", "", 8)
    for _, row in df_agregado.iterrows():
        pdf.cell(col_widths['Filial'], 10, str(row['Filial']), 1, 0, 'L')
        pdf.cell(col_widths['Descrição da conta'], 10,
                 str(row['Descrição da conta']), 1, 0, 'L')
        pdf.cell(col_widths['Valor atualizado'], 10,
                 formatar_valor(row['Valor atualizado']), 1, 0, 'R')
        pdf.cell(col_widths['Deprec. Acumulada'], 10,
                 formatar_valor(row['Deprec. Acumulada']), 1, 0, 'R')
        pdf.cell(col_widths['Valor Residual'], 10,
                 formatar_valor(row['Valor Residual']), 1, 0, 'R')
        pdf.ln()

    pdf.output(buffer)


# --- ESTRUTURA DA APLICAÇÃO ---
st.title("Dashboard de Ativos Contábeis")

with st.sidebar:
    try:
        st.image("logo_GW.png", width=200)
    except Exception:
        st.title("General Water")
    st.header("Instruções")
    st.info("1. **Carregue** os arquivos Excel.\n2. **Aguarde** o processamento.\n3. **Filtre** e analise os dados.\n4. **Explore** os gráficos interativos.\n5. **Baixe** os relatórios.")
    st.header("Ajuda & Suporte")
    email1 = "bruce@generalwater.com.br"
    email2 = "nathalia.vidal@generalwater.com.br"
    mensagem_inicial = "Olá, preciso de ajuda com o Dashboard de Ativos Contábeis."
    link_teams = f"https://teams.microsoft.com/l/chat/0/0?users={email1},{email2}&message={urllib.parse.quote(mensagem_inicial)}"
    st.markdown(f'<a href="{link_teams}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #4B53BC; color: white; text-align: center; text-decoration: none; border-radius: 5px; font-weight: bold;">Abrir Chat no Teams</a>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("Escolha os arquivos Excel de ativos", type=[
                                  'xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    all_data, errors = [], []
    progress_bar = st.progress(0, text="Iniciando processamento...")

    for i, file in enumerate(uploaded_files):
        progress_text = f"Processando arquivo {i+1}/{len(uploaded_files)}: {file.name}"
        progress_bar.progress((i) / len(uploaded_files), text=progress_text)
        dados, erro = processar_planilha(file)
        if dados is not None and not dados.empty:
            all_data.append(dados)
        if erro:
            errors.append(erro)

    progress_bar.progress(1.0, text="Processamento concluído!")
    progress_bar.empty()

    if all_data:
        dados_combinados = pd.concat(all_data, ignore_index=True)
        st.success(
            f"Processamento finalizado! {len(all_data)} arquivo(s) válidos carregados, totalizando {len(dados_combinados)} registros.")

        # --- Filtros ---
        st.markdown("### Filtros")
        col1, col2, col3 = st.columns(3)
        arquivos_options = sorted(dados_combinados['Arquivo'].unique())
        filiais_options = sorted(dados_combinados['Filial'].unique())
        categorias_options = sorted(
            dados_combinados['Descrição da conta'].unique())

        with col1:
            selecao_arquivo = st.multiselect(
                "Filtrar por Arquivo:", arquivos_options, default=arquivos_options)
        with col2:
            selecao_filial = st.multiselect(
                "Filtrar por Filial:", filiais_options, default=filiais_options)
        with col3:
            selecao_categoria = st.multiselect(
                "Filtrar por Descrição da Conta:", categorias_options, default=categorias_options)

        dados_filtrados = dados_combinados[
            (dados_combinados['Arquivo'].isin(selecao_arquivo)) &
            (dados_combinados['Filial'].isin(selecao_filial)) &
            (dados_combinados['Descrição da conta'].isin(selecao_categoria))
        ]

        # --- Métricas ---
        st.markdown("### Resumo dos Dados Filtrados")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Registros Filtrados", f"{len(dados_filtrados):,}")
        col2.metric("Valor Total Atualizado", formatar_valor(
            dados_filtrados["Valor atualizado"].sum()))
        col3.metric("Depreciação Acumulada", formatar_valor(
            dados_filtrados["Deprec. Acumulada"].sum()))
        col4.metric("Valor Residual Total", formatar_valor(
            dados_filtrados["Valor Residual"].sum()))

        # --- Abas de Análise ---
        tab1, tab2, tab3 = st.tabs(
            ["Dados Detalhados", "Análise por Filial", "Análise por Descrição da Conta"])
        with tab1:
            df_display = dados_filtrados.copy()
            colunas_formatar = ['Valor original', 'Valor atualizado', 'Deprec. do mês',
                                'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual']
            for col in colunas_formatar:
                if col in df_display.columns:
                    df_display[col] = df_display[col].apply(formatar_valor)
            st.dataframe(df_display, use_container_width=True, height=500)
        with tab2:
            analise_filial = dados_filtrados.groupby('Filial').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor atualizado', 'sum')).reset_index()
            st.dataframe(analise_filial, use_container_width=True, column_config={
                         "Valor_Total": st.column_config.NumberColumn(format="R$ %.2f")})
        with tab3:
            analise_categoria = dados_filtrados.groupby('Descrição da conta').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor atualizado', 'sum')).reset_index()
            st.dataframe(analise_categoria, use_container_width=True, column_config={
                         "Valor_Total": st.column_config.NumberColumn(format="R$ %.2f")})

        # --- Gráfico Interativo ---
        st.markdown("---")
        st.header("Gráfico Interativo")

        opcoes_eixo_y = ["Valor atualizado",
                         "Deprec. Acumulada", "Valor Residual"]
        col_graf1, col_graf2, col_graf3 = st.columns(3)
        with col_graf1:
            tipo_grafico = st.selectbox("Escolha o Tipo de Gráfico:", [
                                        "Barras", "Pizza", "Linhas"])
        with col_graf2:
            eixo_x = st.selectbox("Agrupar por (Eixo X):", [
                                  "Filial", "Descrição da conta", "Arquivo"])
        with col_graf3:
            if tipo_grafico == "Pizza":
                eixos_y = [st.selectbox(
                    "Analisar Valor:", opcoes_eixo_y, index=0)]
            else:
                eixos_y = st.multiselect("Analisar Valores (Eixo Y):", opcoes_eixo_y, default=[
                                         "Valor atualizado", "Valor Residual"])

        if not dados_filtrados.empty and eixo_x and eixos_y:
            dados_agrupados = dados_filtrados.groupby(
                eixo_x)[eixos_y].sum().reset_index()

            fig_plotly = None
            if tipo_grafico == "Barras":
                fig_plotly = px.bar(dados_agrupados, x=eixo_x, y=eixos_y, text_auto='.2s',
                                    barmode='group', title=f'Análise de {", ".join(eixos_y)} por {eixo_x}')
            elif tipo_grafico == "Linhas":
                fig_plotly = px.line(dados_agrupados, x=eixo_x, y=eixos_y, markers=True,
                                     title=f'Análise de {", ".join(eixos_y)} por {eixo_x}')
            elif tipo_grafico == "Pizza":
                fig_plotly = px.pie(dados_agrupados, names=eixo_x,
                                    values=eixos_y[0], hole=0.3, title=f'Distribuição de {eixos_y[0]} por {eixo_x}')
                fig_plotly.update_traces(
                    textposition='outside', textinfo='percent+label')

            if fig_plotly:
                fig_plotly.update_layout(
                    uniformtext_minsize=8, uniformtext_mode='hide', legend_title_text='')
                st.plotly_chart(fig_plotly, use_container_width=True)
                st.session_state.figura_plotly = fig_plotly  # Salva a figura para exportação
            else:
                st.session_state.figura_plotly = None
        else:
            st.info(
                "Não há dados para exibir no gráfico com os filtros selecionados.")
            st.session_state.figura_plotly = None

        # --- Exportação ---
        st.markdown("---")
        st.header("Exportar Relatório")

        col_download1, col_download2 = st.columns(2)
        with col_download1:
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                dados_filtrados.to_excel(
                    writer, sheet_name='Dados_Filtrados', index=False)
            st.download_button(
                label="📥 Baixar Dados (Excel)",
                data=output_excel.getvalue(),
                file_name="dados_ativos_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_download2:
            if st.session_state.figura_plotly:
                output_pdf = BytesIO()
                criar_pdf_completo(output_pdf, dados_filtrados,
                                   st.session_state.figura_plotly)
                st.download_button(
                    label="📄 Baixar Relatório (PDF)",
                    data=output_pdf.getvalue(),
                    file_name="relatorio_ativos.pdf",
                    mime="application/pdf"
                )
            else:
                st.warning(
                    "Gere um gráfico primeiro para poder exportar o relatório em PDF.")

    if errors:
        st.warning(
            "Alguns arquivos não puderam ser processados ou não continham dados válidos:")
        for error_msg in errors:
            st.error(error_msg)

    if not all_data and not errors:
        st.info("Nenhum dado válido foi encontrado nos arquivos carregados.")

else:
    st.info("Por favor, carregue um ou mais arquivos Excel para iniciar a análise.")
