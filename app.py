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
        if s == '-' or s == '' or s.lower() == 'nan' or s == 'None':
            return 0.0
        # Remove caracteres não numéricos, exceto ponto e vírgula
        s = re.sub(r'[^\d,.-]', '', s)
        # Substitui vírgula por ponto para conversão
        s = s.replace('.', '').replace(',', '.')
        return float(s)
    except (ValueError, AttributeError, TypeError):
        return 0.0


def extrair_quantidade(texto):
    """Extrai a quantidade do texto como 'QUANTIDADE XX'"""
    try:
        match = re.search(r'QUANTIDADE\s+(\d+)', texto, re.IGNORECASE)
        if match:
            return int(match.group(1))
        return 1  # Default para 1 se não encontrar
    except:
        return 1


def processar_planilha_sintetica(file_content, file_name):
    """Processa arquivos com dados sintéticos (totais por conta)"""
    try:
        # Lê todas as planilhas do arquivo
        xls = pd.ExcelFile(BytesIO(file_content))
        dados_finais = []

        for sheet_name in xls.sheet_names:
            if 'Parametros' in sheet_name:
                continue  # Pula a planilha de parâmetros

            df = pd.read_excel(BytesIO(file_content),
                               sheet_name=sheet_name, header=None)
            df = df.astype(str).fillna('')

            conta_atual = ""
            descricao_conta = ""
            filial = ""
            quantidade = 1

            for idx, row in df.iterrows():
                linha = ' '.join([str(cell)
                                 for cell in row if str(cell) != 'nan'])

                # Identifica linha de conta contábil
                if '1.2.3.' in linha and any(term in linha for term in ['ESTACAO', 'Veiculos', 'Maquinas', 'Computadores', 'SOFTWARE', 'Moveis', 'BENF', 'DIREITO']):
                    partes = linha.split()
                    for i, parte in enumerate(partes):
                        if '1.2.3.' in parte:
                            conta_atual = parte
                            descricao_conta = ' '.join(partes[i+1:])
                            break

                # Identifica linha de filial
                elif 'Filial :' in linha:
                    filial = linha.replace('Filial :', '').strip()

                # Identifica linha de quantidade
                elif 'QUANTIDADE' in linha.upper():
                    quantidade = extrair_quantidade(linha)

                # Identifica linha de valores em R$
                elif linha.startswith('R$') or (len(linha.split()) > 3 and linha.split()[0] == 'R$'):
                    partes = linha.split()
                    if len(partes) >= 8:  # Garante que temos valores suficientes
                        try:
                            dados_item = {
                                "Arquivo": file_name,
                                "Filial": filial,
                                "Conta contábil": conta_atual,
                                "Descrição da conta": descricao_conta,
                                "Quantidade": quantidade,
                                "Valor original": limpar_valor(partes[2]),
                                "Valor atualizado": limpar_valor(partes[3]),
                                "Deprec. do mês": limpar_valor(partes[4]),
                                "Deprec. Acumulada": limpar_valor(partes[6]),
                                "Valor Residual": limpar_valor(partes[3]) - limpar_valor(partes[6]),
                                "Planilha": sheet_name
                            }
                            dados_finais.append(dados_item)
                        except (IndexError, ValueError) as e:
                            continue

        if dados_finais:
            df_final = pd.DataFrame(dados_finais)
            return df_final, None
        else:
            return pd.DataFrame(), f"Nenhum dado sintético foi encontrado em '{file_name}'."

    except Exception as e:
        return pd.DataFrame(), f"Erro ao processar '{file_name}': {e}\n{traceback.format_exc()}"

# --- FUNÇÃO DE PROCESSAMENTO PRINCIPAL ---


@st.cache_data
def processar_planilha(file_content, file_name):
    return processar_planilha_sintetica(file_content, file_name)

# --- RESTANTE DO CÓDIGO PERMANECE IGUAL ---
# [O restante do código da interface permanece exatamente como estava]


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
            f"Processamento finalizado! {len(all_data)} arquivo(s) carregados, totalizando {len(dados_combinados):,} registros.")

        # [O restante do código de interface permanece igual...]
        # ... (filtros, métricas, tabelas, gráficos, etc.)

else:
    st.info(
        "Por favor, carregue um ou mais arquivos de relatório para iniciar a análise.")
