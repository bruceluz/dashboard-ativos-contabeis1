import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
from fpdf import FPDF
import traceback
import os
import re  # Importa o módulo de expressões regulares

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
        font_path = 'LiberationSans-Regular.ttf'
        font_name = "LiberationSans"
        fallback_font = "Arial"

        if os.path.exists(font_path) and font_name not in self.font_families:
            try:
                self.add_font(font_name, "", font_path, uni=True)
                self.add_font(font_name, "B", font_path, uni=True)
                self.add_font(font_name, "I", font_path, uni=True)
            except Exception:
                st.warning(
                    f"Não foi possível carregar a fonte '{font_name}'. Usando '{fallback_font}'.")

        current_font = font_name if font_name in self.font_families else fallback_font

        try:
            self.image("logo_GW.png", x=10, y=8, w=40)
        except FileNotFoundError:
            self.set_font(current_font, "B", 12)
            self.cell(40, 10, "General Water", 0, 0, 'L')

        self.set_font(current_font, "B", 20)
        self.cell(0, 10, "Relatório de Ativos Contábeis", 0, 1, 'C')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        font_name = "LiberationSans"
        fallback_font = "Arial"
        current_font = font_name if font_name in self.font_families else fallback_font
        self.set_font(current_font, "I", 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# --- FUNÇÕES DE LÓGICA ---


def formatar_valor(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"


def limpar_valor(valor_str):
    try:
        s = str(valor_str).strip()
        if s == '-' or s == '':
            return 0.0
        return float(s.replace('.', '').replace(',', '.'))
    except (ValueError, AttributeError, TypeError):
        return 0.0

# --- LÓGICA DE PARSING APRIMORADA ---


def processar_linhas_do_arquivo(linhas_texto, nome_arquivo):
    dados_finais = []
    conta_atual, descricao_conta_atual = "", ""
    dados_item_atual = None

    TERMOS_A_IGNORAR = [
        "Filial", "C Custo", "Cod Base Bem", "Vl Ampliac.1", "US$", "UFIR",
        "TOTAL Conta", "BAIXAS", "T O T A L   G E R A L", "B A I X A S", "T O T A L"
    ]

    for linha in linhas_texto:
        linha_strip = linha.strip()

        if not linha_strip or any(termo in linha for termo in TERMOS_A_IGNORAR):
            continue

        # --- MUDANÇA CRÍTICA: Usa expressão regular para dividir por tabs ou múltiplos espaços ---
        colunas = [p.strip() for p in re.split(
            r'\s{2,}|[\t]', linha_strip) if p.strip()]
        if not colunas:
            continue

        if len(colunas) >= 2 and colunas[0].startswith("1.2.3."):
            conta_atual = colunas[0]
            descricao_conta_atual = " ".join(colunas[1:])
            dados_item_atual = None
            continue

        elif colunas[0] == "R$":
            if dados_item_atual:
                try:
                    # A lógica de índice aqui depende da separação correta.
                    # Ex: ['R$', '-', '1.119.850,39', ...]
                    dados_item_atual["Valor original"] = limpar_valor(
                        colunas[2])
                    dados_item_atual["Valor atualizado"] = limpar_valor(
                        colunas[3])
                    dados_item_atual["Deprec. do mês"] = limpar_valor(
                        colunas[4])
                    dados_item_atual["Deprec. Acumulada"] = limpar_valor(
                        colunas[6])
                    dados_item_atual["Valor Residual"] = dados_item_atual["Valor atualizado"] - \
                        dados_item_atual["Deprec. Acumulada"]
                    dados_finais.append(dados_item_atual)
                except (IndexError, ValueError):
                    pass
                finally:
                    dados_item_atual = None

        elif colunas[0].isdigit() and len(colunas) > 8:
            try:
                if dados_item_atual:
                    dados_item_atual = None

                data_aquisicao = pd.to_datetime(
                    colunas[7], format='%d/%m/%Y', errors='coerce')

                if pd.notna(data_aquisicao):
                    dados_item_atual = {
                        "Arquivo": nome_arquivo,
                        "Filial": colunas[0],
                        "Conta contábil": conta_atual,
                        "Descrição da conta": descricao_conta_atual,
                        "Código do item": colunas[2],
                        "Código do sub item": colunas[3],
                        "Descrição do item": colunas[5],
                        "Data de aquisição": data_aquisicao,
                        "Valor original": 0.0, "Valor atualizado": 0.0, "Deprec. do mês": 0.0,
                        "Deprec. Acumulada": 0.0, "Valor Residual": 0.0,
                    }
            except (ValueError, IndexError):
                dados_item_atual = None
                continue

    return dados_finais

# --- FUNÇÃO DE PROCESSAMENTO PRINCIPAL ---


@st.cache_data
def processar_planilha(file_content, file_name):
    try:
        # Tenta decodificar como UTF-8, com fallback para latin-1
        try:
            text_content = file_content.decode('utf-8')
        except UnicodeDecodeError:
            text_content = file_content.decode('latin-1', errors='replace')

        linhas_texto = text_content.splitlines()
        dados_processados = processar_linhas_do_arquivo(
            linhas_texto, file_name)

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            colunas_ordenadas = [
                'Filial', 'Conta contábil', 'Descrição da conta', 'Data de aquisição',
                'Código do item', 'Código do sub item', 'Descrição do item', 'Valor original', 'Valor atualizado',
                'Deprec. do mês', 'Deprec. Acumulada', 'Valor Residual', 'Arquivo'
            ]
            df_final = df_final.reindex(
                columns=colunas_ordenadas, fill_value="")
            return df_final, None
        else:
            return pd.DataFrame(), f"Nenhum item de ativo detalhado foi encontrado em '{file_name}'."
    except Exception as e:
        return pd.DataFrame(), f"Erro crítico ao processar '{file_name}': {e}\n{traceback.format_exc()}"

# --- O RESTANTE DO CÓDIGO DA INTERFACE PERMANECE O MESMO ---


st.title("📊 Dashboard de Ativos Contábeis")

with st.sidebar:
    if os.path.exists("logo_GW.png"):
        st.image("logo_GW.png", width=200)
    else:
        st.title("General Water")
    st.header("Instruções")
    st.info("1. **Carregue** um ou mais relatórios.\n2. **Aguarde** o processamento.\n3. **Use os filtros** para analisar.\n4. **Explore** os gráficos interativos.\n5. **Baixe** os dados ou o PDF.")
    st.header("Ajuda & Suporte")
    email1 = "bruce@generalwater.com.br"
    email2 = "nathalia.vidal@generalwater.com.br"
    mensagem_inicial = "Olá, preciso de ajuda com o Dashboard de Ativos Contábeis."
    link_teams = f"https://teams.microsoft.com/l/chat/0/0?users={email1},{email2}&message={urllib.parse.quote(mensagem_inicial)}"
    st.markdown(f'<a href="{link_teams}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #4B53BC; color: white; text-align: center; text-decoration: none; border-radius: 5px; font-weight: bold;">Abrir Chat no Teams</a>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("Escolha os arquivos de relatório de ativos", type=[
                                  'xlsx', 'xls', 'txt'], accept_multiple_files=True)

if uploaded_files:
    all_data, errors = [], []
    progress_bar = st.progress(0, text="Iniciando processamento...")

    for i, file in enumerate(uploaded_files):
        progress_text = f"Processando arquivo {i+1}/{len(uploaded_files)}: {file.name}"
        progress_bar.progress(
            (i + 1) / len(uploaded_files), text=progress_text)

        file_content = file.getvalue()
        dados, erro = processar_planilha(file_content, file.name)

        if dados is not None and not dados.empty:
            all_data.append(dados)
        if erro:
            errors.append(erro)

    progress_bar.empty()

    if all_data:
        dados_combinados = pd.concat(all_data, ignore_index=True)
        st.success(
            f"Processamento finalizado! {len(all_data)} arquivo(s) com dados detalhados carregados, totalizando {len(dados_combinados):,} registros.")

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
        st.warning("Avisos durante o processamento:")
        for error_msg in errors:
            if "Nenhum item de ativo detalhado foi encontrado" in error_msg:
                st.info(error_msg)
            else:
                st.error(error_msg)

    if not all_data and not any("crítico" in e for e in errors):
        st.info("Nenhum dado detalhado de ativo foi encontrado nos arquivos carregados. Os arquivos podem conter apenas totais.")

else:
    st.info(
        "Por favor, carregue um ou mais arquivos de relatório para iniciar a análise.")
