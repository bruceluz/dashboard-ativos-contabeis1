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
if 'figura_plotly' not in st.session_state:
    st.session_state.figura_plotly = None

# --- CLASSE FPDF CUSTOMIZADA ---


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


def formatar_valor(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"

# --- NOVA LÓGICA DE PROCESSAMENTO (ADAPTADA DO SEU CÓDIGO) ---


CABECALHOS_INTERMEDIARIOS = ["Cod Base Bem", "Descr. Sint.",
                             "Dt.Aquisicao", "Data Baixa", "Quantidade", "Num.Plaqueta"]


def limpar_valor(valor_str):
    """Converte valor monetário em string para float."""
    try:
        return float(str(valor_str).replace('.', '').replace(',', '.'))
    except (ValueError, AttributeError):
        return 0.0  # Retorna 0.0 para consistência nos cálculos


def processar_linhas_do_arquivo(linhas_texto, nome_arquivo):
    """Processa as linhas extraídas do Excel e retorna uma lista de dicionários."""
    dados = []
    conta_atual = ""
    descricao_conta_atual = ""

    for linha in linhas_texto:
        linha = linha.strip()
        if any(cabecalho in linha for cabecalho in CABECALHOS_INTERMEDIARIOS) or not linha:
            continue

        colunas = linha.split('\t')

        # Tenta identificar a linha da conta contábil
        if len(colunas) >= 2 and colunas[0].startswith("1.2.3."):
            conta_atual = colunas[0]
            descricao_conta_atual = colunas[1]
            continue

        # Tenta identificar a linha de valores "R$"
        elif colunas[0] == "R$":
            if dados:  # Se houver um item anterior na lista para atualizar
                try:
                    item_anterior = dados[-1]
                    item_anterior["Valor original"] = limpar_valor(colunas[2])
                    item_anterior["Valor atualizado"] = limpar_valor(
                        colunas[3])
                    item_anterior["Deprec. do mês"] = limpar_valor(colunas[4])
                    item_anterior["Deprec. Acumulada"] = limpar_valor(
                        colunas[6])

                    # Calcula o valor residual
                    valor_atualizado = item_anterior["Valor atualizado"]
                    deprec_acumulada = item_anterior["Deprec. Acumulada"]
                    item_anterior["Valor Residual"] = valor_atualizado - \
                        deprec_acumulada

                except IndexError:
                    continue  # Pula se a linha R$ não tiver colunas suficientes

        # Tenta identificar uma linha de item de ativo
        elif len(colunas) >= 8 and colunas[0].isdigit():
            try:
                # Converte data numérica do Excel para datetime
                data_aquisicao = pd.to_datetime(
                    int(colunas[7]), unit='d', origin='1899-12-30', errors='coerce')

                dados.append({
                    "Arquivo": nome_arquivo,
                    "Filial": colunas[0],
                    "Conta contábil": conta_atual,
                    "Descrição da conta": descricao_conta_atual,
                    "Código do item": colunas[3],
                    "Descrição do item": colunas[5],
                    "Data de aquisição": data_aquisicao,
                    # Inicializa valores que serão preenchidos pela linha "R$"
                    "Valor original": 0.0,
                    "Valor atualizado": 0.0,
                    "Deprec. do mês": 0.0,
                    "Deprec. Acumulada": 0.0,
                    "Valor Residual": 0.0,
                })
            except (ValueError, IndexError):
                continue

    return dados


def processar_planilha(file):
    """Função principal que integra a nova lógica com o Streamlit."""
    try:
        # Lê o arquivo Excel, tratando todas as células como texto
        df_raw = pd.read_excel(file, header=None, dtype=str)
        # Concatena as células de cada linha com um tab, criando uma lista de strings
        linhas_texto = df_raw.fillna("").astype(
            str).agg('\t'.join, axis=1).tolist()

        dados_processados = processar_linhas_do_arquivo(
            linhas_texto, file.name)

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            # Garante que todas as colunas esperadas existam
            colunas_ordenadas = [
                'Filial', 'Conta contábil', 'Descrição da conta', 'Data de aquisição',
                'Código do item', 'Descrição do item', 'Valor original', 'Valor atualizado',
                'Deprec. do mês', 'Deprec. Acumulada', 'Valor Residual', 'Arquivo'
            ]
            df_final = df_final.reindex(
                columns=colunas_ordenadas, fill_value="")
            return df_final, None
        else:
            return pd.DataFrame(), f"Nenhum item individual foi encontrado no formato esperado em '{file.name}'."

    except Exception as e:
        st.error(f"Erro crítico ao processar o arquivo '{file.name}': {e}")
        st.code(traceback.format_exc())
        return pd.DataFrame(), f"Erro crítico ao processar '{file.name}'."


def criar_pdf_completo(buffer, df_filtrado, fig_plotly):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    try:
        pdf.add_font('Arial', '', 'fonts/arial.ttf', uni=True)
        pdf.add_font('Arial', 'B', 'fonts/arialbd.ttf', uni=True)
        pdf.set_font('Arial', '', 12)
    except RuntimeError:
        st.warning("Arquivos de fonte não encontrados. Usando fonte padrão.")
        pdf.set_font('Helvetica', '', 12)
    pdf.add_page()

    if fig_plotly:
        try:
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Gráfico Analítico", 0, 1, 'L')
            img_bytes = fig_plotly.to_image(
                format="png", width=1000, height=500, scale=2)
            img_buffer = BytesIO(img_bytes)
            pdf.image(img_buffer, x=10, y=pdf.get_y(), w=277)
            pdf.ln(135)
        except Exception as e:
            pdf.set_font("Arial", "", 10)
            pdf.multi_cell(
                0, 10, f"Não foi possível renderizar o gráfico no PDF: {e}", 0, 'L')

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Dados Agregados por Filial e Descrição da Conta", 0, 1, 'L')
    pdf.ln(5)

    colunas_para_somar = ['Valor atualizado',
                          'Deprec. Acumulada', 'Valor Residual']
    df_agregado = df_filtrado.groupby(['Filial', 'Descrição da conta'])[
        colunas_para_somar].sum().reset_index()

    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(230, 230, 230)
    col_widths = {'Filial': 60, 'Descrição da conta': 100,
                  'Valor atualizado': 38, 'Deprec. Acumulada': 40, 'Valor Residual': 39}
    for col_name in col_widths.keys():
        pdf.cell(col_widths[col_name], 10, col_name, 1, 0, 'C', fill=True)
    pdf.ln()

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


# --- ESTRUTURA DA APLICAÇÃO (INTERFACE DO USUÁRIO) ---
st.title("📊 Dashboard de Ativos Contábeis")

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

        st.markdown("### Resumo dos Dados Filtrados")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Registros Filtrados", f"{len(dados_filtrados):,}")
        col2.metric("Valor Total Atualizado", formatar_valor(
            dados_filtrados["Valor atualizado"].sum()))
        col3.metric("Depreciação Acumulada", formatar_valor(
            dados_filtrados["Deprec. Acumulada"].sum()))
        col4.metric("Valor Residual Total", formatar_valor(
            dados_filtrados["Valor Residual"].sum()))

        tab1, tab2, tab3 = st.tabs(
            ["Dados Detalhados", "Análise por Filial", "Análise por Descrição da Conta"])
        with tab1:
            df_display = dados_filtrados.copy()
            colunas_formatar = ['Valor original', 'Valor atualizado',
                                'Deprec. do mês', 'Deprec. Acumulada', 'Valor Residual']
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
                st.session_state.figura_plotly = fig_plotly
            else:
                st.session_state.figura_plotly = None
        else:
            st.info(
                "Não há dados para exibir no gráfico com os filtros selecionados.")
            st.session_state.figura_plotly = None

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
