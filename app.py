import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
from fpdf import FPDF
import traceback
import os
import re
from datetime import datetime

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
        # Tenta usar uma fonte específica, mas recorre a uma padrão se não encontrada
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
    """Formata um número para o padrão monetário brasileiro (R$)."""
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"


def limpar_valor(valor_str):
    """Converte uma string monetária (ex: '1.234,56') para float."""
    try:
        s = str(valor_str).strip()
        if s in ('-', '', 'nan', 'None'):
            return 0.0
        # Remove caracteres não numéricos, exceto ponto, vírgula e sinal negativo
        s = re.sub(r'[^\d,.-]', '', s)
        # Padroniza para o formato americano (ponto como decimal) antes de converter
        s = s.replace('.', '').replace(',', '.')
        return float(s)
    except (ValueError, AttributeError, TypeError):
        return 0.0


def processar_linhas_detalhadas(linhas_texto, nome_arquivo):
    """Processa o conteúdo de um arquivo texto para extrair dados de ativos."""
    dados_finais = []
    conta_atual, descricao_conta_atual = "", ""
    dados_item_atual = None

    for linha in linhas_texto:
        linha_strip = linha.strip()
        if not linha_strip:
            continue

        try:
            # Identifica linha de conta contábil (ex: "1.2.3.02.010001 ESTACAO...")
            if '1.2.3.' in linha_strip and len(linha_strip.split()) >= 2:
                partes = linha_strip.split()
                for j, parte in enumerate(partes):
                    if '1.2.3.' in parte:
                        conta_atual = parte
                        descricao_conta_atual = ' '.join(partes[j+1:])
                        break

            # Identifica linha de dados detalhados (começa com código numérico de filial)
            elif linha_strip.split() and linha_strip.split()[0].isdigit() and len(linha_strip.split()) >= 8:
                partes = linha_strip.split()
                # Extrai os dados da linha do item
                filial = partes[0]
                cod_base_bem = partes[2]
                codigo_item = partes[3]
                # Junta a descrição que pode conter espaços
                descr_sint = ' '.join(partes[5:-5])
                data_aquisicao = partes[-5]

                dados_item_atual = {
                    "Arquivo": nome_arquivo, "Filial": filial,
                    "Conta contábil": conta_atual, "Descrição da conta": descricao_conta_atual,
                    "Data de aquisição": data_aquisicao, "Código do item": cod_base_bem,
                    "Código do sub item": codigo_item, "Descrição do item": descr_sint,
                    "Valor original": 0.0, "Valor atualizado": 0.0, "Deprec. do mês": 0.0,
                    "Deprec. do exercício": 0.0, "Deprec. Acumulada": 0.0, "Valor Residual": 0.0
                }

            # Identifica linha de valores em R$ (vinculada ao item anterior)
            elif linha_strip.startswith('R$') and dados_item_atual:
                partes = linha_strip.split()
                if len(partes) >= 8:
                    dados_item_atual["Valor original"] = limpar_valor(
                        partes[2])
                    dados_item_atual["Valor atualizado"] = limpar_valor(
                        partes[3])
                    dados_item_atual["Deprec. do mês"] = limpar_valor(
                        partes[4])
                    dados_item_atual["Deprec. do exercício"] = limpar_valor(
                        partes[5])
                    dados_item_atual["Deprec. Acumulada"] = limpar_valor(
                        partes[6])
                    dados_item_atual["Valor Residual"] = dados_item_atual["Valor atualizado"] - \
                        dados_item_atual["Deprec. Acumulada"]

                    dados_finais.append(dados_item_atual)
                # Reseta o item atual para evitar adicionar o mesmo item duas vezes
                dados_item_atual = None

        except (IndexError, ValueError):
            # Se ocorrer um erro ao processar uma linha, reseta e continua
            dados_item_atual = None
            continue

    return dados_finais

# --- FUNÇÃO DE PROCESSAMENTO PRINCIPAL ---


@st.cache_data
def processar_planilha(file_content, file_name):
    """Processa um arquivo (Excel ou Texto) e retorna um DataFrame."""
    try:
        # Lê o arquivo Excel, tratando todas as células como texto para evitar erros de tipo
        df_raw = pd.read_excel(BytesIO(file_content),
                               header=None, engine='openpyxl')
        df_raw = df_raw.astype(str).fillna('')

        # Concatena todas as colunas de cada linha em uma única string
        linhas_texto = df_raw.apply(
            lambda row: ' '.join(row.dropna()), axis=1).tolist()

        dados_processados = processar_linhas_detalhadas(
            linhas_texto, file_name)

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            # Garante a ordem correta das colunas
            colunas_ordenadas = [
                'Arquivo', 'Filial', 'Conta contábil', 'Descrição da conta', 'Data de aquisição',
                'Código do item', 'Código do sub item', 'Descrição do item',
                'Valor original', 'Valor atualizado', 'Deprec. do mês',
                'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual'
            ]
            df_final = df_final.reindex(
                columns=colunas_ordenadas, fill_value="")
            return df_final, None
        else:
            return pd.DataFrame(), f"Nenhum item de ativo detalhado foi encontrado em '{file_name}'."
    except Exception as e:
        return pd.DataFrame(), f"Erro crítico ao processar '{file_name}': {e}\n{traceback.format_exc()}"


# --- INTERFACE DO USUÁRIO ---
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
                                  'xlsx', 'xls'], accept_multiple_files=True)

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

    # --- CORREÇÃO PRINCIPAL: SÓ PROSSEGUIR SE HOUVER DADOS ---
    if all_data:
        dados_combinados = pd.concat(all_data, ignore_index=True)

        if not dados_combinados.empty:
            st.success(
                f"Processamento finalizado! {len(all_data)} arquivo(s) carregados, totalizando {len(dados_combinados):,} registros.")

            st.markdown("### Filtros")
            col1, col2, col3 = st.columns(3)

            # Criação segura das opções de filtro
            arquivos_options = sorted(dados_combinados['Arquivo'].unique(
            )) if 'Arquivo' in dados_combinados.columns else []
            filiais_options = sorted(dados_combinados['Filial'].unique(
            )) if 'Filial' in dados_combinados.columns else []
            categorias_options = sorted(dados_combinados['Descrição da conta'].unique(
            )) if 'Descrição da conta' in dados_combinados.columns else []

            with col1:
                selecao_arquivo = st.multiselect(
                    "Filtrar por Arquivo:", arquivos_options, default=arquivos_options)
            with col2:
                selecao_filial = st.multiselect(
                    "Filtrar por Filial:", filiais_options, default=filiais_options)
            with col3:
                selecao_categoria = st.multiselect(
                    "Filtrar por Descrição da Conta:", categorias_options, default=categorias_options)

            # Aplica os filtros ao DataFrame
            dados_filtrados = dados_combinados[
                (dados_combinados['Arquivo'].isin(selecao_arquivo)) &
                (dados_combinados['Filial'].isin(selecao_filial)) &
                (dados_combinados['Descrição da conta'].isin(
                    selecao_categoria))
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
                colunas_formatar = ['Valor original', 'Valor atualizado', 'Deprec. do mês',
                                    'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual']
                for col in colunas_formatar:
                    if col in df_display.columns:
                        df_display[col] = df_display[col].apply(formatar_valor)
                st.dataframe(df_display, use_container_width=True, height=500)

            with tab2:
                analise_filial = dados_filtrados.groupby('Filial').agg(
                    Contagem_de_Itens=('Arquivo', 'count'),
                    Valor_Total_Atualizado=('Valor atualizado', 'sum')
                ).reset_index().sort_values(by='Valor_Total_Atualizado', ascending=False)
                st.dataframe(analise_filial, use_container_width=True,
                             column_config={"Valor_Total_Atualizado": st.column_config.NumberColumn(format="R$ %.2f")})

            with tab3:
                analise_categoria = dados_filtrados.groupby('Descrição da conta').agg(
                    Contagem_de_Itens=('Arquivo', 'count'),
                    Valor_Total_Atualizado=('Valor atualizado', 'sum')
                ).reset_index().sort_values(by='Valor_Total_Atualizado', ascending=False)
                st.dataframe(analise_categoria, use_container_width=True,
                             column_config={"Valor_Total_Atualizado": st.column_config.NumberColumn(format="R$ %.2f")})

            # Adicione aqui o restante do seu código para gráficos e exportação...

        else:
            st.warning(
                "Os arquivos foram processados, mas não foi possível extrair nenhum dado de ativo contábil válido. Verifique o formato dos arquivos.")

    if errors:
        st.warning("Avisos e erros encontrados durante o processamento:")
        for error_msg in errors:
            st.error(error_msg)

else:
    st.info(
        "Por favor, carregue um ou mais arquivos de relatório para iniciar a análise.")
