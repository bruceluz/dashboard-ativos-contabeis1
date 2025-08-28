import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
from fpdf import FPDF
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import re

warnings.filterwarnings('ignore')

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard de Ativos Contábeis",
    page_icon=" ",
    layout="wide"
)

# --- INICIALIZAÇÃO DO SESSION STATE ---
if 'figura_plotly' not in st.session_state:
    st.session_state.figura_plotly = None
if 'dados_grafico' not in st.session_state:
    st.session_state.dados_grafico = None
if 'tipo_grafico' not in st.session_state:
    st.session_state.tipo_grafico = None
if 'eixo_x' not in st.session_state:
    st.session_state.eixo_x = None
if 'eixos_y' not in st.session_state:
    st.session_state.eixos_y = None

# --- FUNÇÕES DE LÓGICA ---


def padronizar_nome_filial(nome_filial):
    if not isinstance(nome_filial, str):
        return "Não Identificado"
    nome_upper = nome_filial.upper().strip()
    mapa_nomes = {
        "GENERAL WATER": "General Water S/A", "GW S/A": "General Water S/A",
        "G W AGUAS": "GW Águas", "GW ÁGUAS": "GW Águas",
        "GW SANEAMENTO": "GW Saneamento", "GW SANEA": "GW Saneamento",
        "GW SISTEMAS": "GW Sistemas", "GW SISTEM": "GW Sistemas",
        "MATRIZ": "GW Sistemas Matriz"
    }
    return mapa_nomes.get(nome_upper, nome_filial)


def converter_valor(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        valor_str = str(valor).replace('R$', '').strip()
        if ',' in valor_str and '.' in valor_str:
            valor_str = valor_str.replace('.', '')
        valor_str = valor_str.replace(',', '.')
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0


def formatar_valor(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"


def corrigir_filiais_nao_identificadas(df_arquivo):
    if df_arquivo.empty:
        return df_arquivo
    contagem_filiais = df_arquivo[df_arquivo['Filial']
                                  != 'Não Identificado']['Filial'].mode()
    if not contagem_filiais.empty:
        filial_predominante = contagem_filiais[0]
        df_arquivo['Filial'] = df_arquivo['Filial'].replace(
            'Não Identificado', filial_predominante)
    return df_arquivo

# ### LÓGICA DE PROCESSAMENTO FINAL - LEITURA EM PARES DE LINHAS ###


def processar_planilha(file):
    try:
        xl = pd.ExcelFile(file)
        dados_processados = []

        for sheet_name in xl.sheet_names:
            sheet_df = pd.read_excel(
                xl, sheet_name=sheet_name, header=None, dtype=str)

            filial_atual = "Não Identificado"
            categoria_atual = "Não Identificado"

            # Usamos um iterador para poder avançar para a próxima linha manualmente
            rows_iter = sheet_df.iterrows()
            for index, row in rows_iter:
                if row.isnull().all():
                    continue

                # Captura a Filial
                row_str = ' '.join(str(x) for x in row if pd.notna(x))
                if 'Filial :' in row_str:
                    nome_extraido = row_str.split(
                        'Filial :')[-1].split(' - ')[-1].strip()
                    filial_atual = padronizar_nome_filial(nome_extraido)
                    continue

                # Captura a Categoria do ativo
                if str(row.iloc[0]).startswith('1.2.3.'):
                    categoria_atual = str(row.iloc[1]).strip() if pd.notna(
                        row.iloc[1]) else "Não Identificado"
                    continue

                # Identifica a LINHA DE DADOS do ativo
                # Critérios: coluna 2 (C Custo) e 3 (Cod Base Bem) não são vazias e a coluna 8 (Dt.Aquisicao) parece uma data
                c_custo = str(row.iloc[0]).strip() if pd.notna(
                    row.iloc[0]) else None
                cod_base_bem = str(row.iloc[2]).strip(
                ) if pd.notna(row.iloc[2]) else None
                dt_aquisicao = str(row.iloc[7]).strip(
                ) if pd.notna(row.iloc[7]) else None

                if c_custo and cod_base_bem and dt_aquisicao and '/' in dt_aquisicao:
                    dados_do_ativo = {
                        'Arquivo': file.name,
                        'Filial': filial_atual,
                        'Categoria': categoria_atual,
                        'C Custo': c_custo,
                        'Cod Base Bem': cod_base_bem,
                        'Codigo Item': str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None,
                        'Tipo Ativo': str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else None,
                        'Descr. Sint.': str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else None,
                        'Tipo Depr.': str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else None,
                        'Dt.Aquisicao': dt_aquisicao,
                        'Data Baixa': str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else None,
                        'Quantidade': converter_valor(str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else '0'),
                        'Num.Plaqueta': str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else None,
                        'Item Despesa': str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) else None,
                        'ClVl Despesa': str(row.iloc[12]).strip() if pd.notna(row.iloc[12]) else None,
                    }

                    # Tenta ler a PRÓXIMA linha, que deve ser a de valores (R$)
                    try:
                        _, next_row = next(rows_iter)
                        if pd.notna(next_row.iloc[0]) and 'R$' in str(next_row.iloc[0]):
                            valores = [converter_valor(v)
                                       for v in next_row.iloc[1:14]]

                            dados_do_ativo.update({
                                'Vl Ampliac.1': valores[0] if len(valores) > 0 else 0,
                                'Valor Original': valores[1] if len(valores) > 1 else 0,
                                'Valor Atualizado': valores[2] if len(valores) > 2 else 0,
                                'Deprec. no mes': valores[3] if len(valores) > 3 else 0,
                                'Deprec. no Exerc.': valores[4] if len(valores) > 4 else 0,
                                'Deprec. Acumulada': valores[5] if len(valores) > 5 else 0,
                                'Corre Mes M1': valores[7] if len(valores) > 7 else 0,
                                'Corre Bal M1': valores[8] if len(valores) > 8 else 0,
                                'Corr Acum M1': valores[9] if len(valores) > 9 else 0,
                                'Cor Dep Mes': valores[10] if len(valores) > 10 else 0,
                                'Cor Dep Exer': valores[11] if len(valores) > 11 else 0,
                                'Cor Dep Acum': valores[12] if len(valores) > 12 else 0,
                            })

                            valor_atualizado = dados_do_ativo.get(
                                'Valor Atualizado', 0)
                            deprec_acumulada = dados_do_ativo.get(
                                'Deprec. Acumulada', 0)
                            dados_do_ativo['Valor Residual'] = valor_atualizado - \
                                deprec_acumulada

                            dados_processados.append(dados_do_ativo)
                    except StopIteration:
                        # Chegou ao fim do arquivo
                        break

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            return corrigir_filiais_nao_identificadas(df_final), None

        return None, f"Nenhum registro de ativo válido encontrado em {file.name}."
    except Exception as e:
        return None, f"Erro crítico ao processar {file.name}: {e}"


def criar_pdf_completo(buffer, df_filtrado, dados_grafico, tipo_grafico, eixo_x, eixos_y):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()

    try:
        pdf.image("logo_GW.png", x=10, y=8, w=40)
    except Exception:
        pdf.set_font("Arial", "B", 12)
        pdf.cell(40, 10, "General Water", 0, 1, 'L')
    pdf.set_font("Arial", "B", 20)
    pdf.cell(0, 10, "Relatório de Ativos Contábeis", 0, 1, 'C')
    pdf.ln(15)

    if dados_grafico is not None:
        try:
            fig, ax = plt.subplots(figsize=(11, 5))

            if tipo_grafico in ['Barras', 'Linhas']:
                df_plot = dados_grafico.set_index(eixo_x)
                df_plot.plot(
                    kind='bar' if tipo_grafico == 'Barras' else 'line',
                    ax=ax,
                    rot=45,
                    grid=True
                )
                ax.set_title(f'Análise por {eixo_x}')
                ax.set_ylabel('Valores (R$)')
                ax.yaxis.set_major_formatter(
                    mticker.FuncFormatter(lambda x, p: f'R$ {x:,.0f}'))
                ax.legend(title='Métricas')

            elif tipo_grafico == 'Pizza':
                metrica_unica = eixos_y[0]
                ax.pie(
                    dados_grafico[metrica_unica],
                    labels=dados_grafico[eixo_x],
                    autopct='%1.1f%%',
                    startangle=90
                )
                ax.set_title(f'Distribuição de {metrica_unica} por {eixo_x}')
                ax.axis('equal')

            plt.tight_layout()

            img_buffer = BytesIO()
            fig.savefig(img_buffer, format='png', dpi=300)
            img_buffer.seek(0)

            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Gráfico Analítico", 0, 1, 'L')
            pdf.image(img_buffer, x=None, y=None, w=277)
            pdf.ln(5)

        except Exception as e:
            pdf.set_font("Arial", "", 10)
            pdf.cell(
                0, 10, f"Nao foi possivel renderizar o grafico no PDF: {e}", 0, 1, 'L')
        finally:
            plt.close(fig)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Dados Agregados por Filial e Categoria", 0, 1, 'L')
    pdf.ln(5)
    colunas_para_somar = ['Valor Atualizado',
                          'Deprec. Acumulada', 'Valor Residual']
    df_agregado = df_filtrado.groupby(['Filial', 'Categoria'])[
        colunas_para_somar].sum().reset_index()
    for col in colunas_para_somar:
        df_agregado[col] = df_agregado[col].apply(formatar_valor)
    col_widths = {'Filial': 60, 'Categoria': 100, 'Valor Atualizado': 35,
                  'Deprec. Acumulada': 40, 'Valor Residual': 35}
    pdf.set_font("Arial", "B", 9)
    for col_name in col_widths.keys():
        pdf.cell(col_widths[col_name], 10, col_name, 1, 0, 'C')
    pdf.ln()
    pdf.set_font("Arial", "", 8)
    for _, row in df_agregado.iterrows():
        for col_name in col_widths.keys():
            cell_text = str(row[col_name]).encode(
                'latin-1', 'replace').decode('latin-1')
            pdf.cell(col_widths[col_name], 10, cell_text, 1, 0, 'L')
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
    st.info("1. **Carregue** os arquivos.\n2. **Aguarde** o processamento.\n3. **Filtre** e analise os dados.\n4. **Explore** os gráficos interativos.\n5. **Baixe** o relatório.")
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
    progress_bar = st.progress(0, text="Iniciando...")
    for i, file in enumerate(uploaded_files):
        progress_bar.progress((i + 1) / len(uploaded_files),
                              text=f"Processando: {file.name}")
        dados, erro = processar_planilha(file)
        if dados is not None and not dados.empty:
            all_data.append(dados)
        if erro:
            errors.append(erro)

    if all_data:
        dados_combinados = pd.concat(all_data, ignore_index=True)
        st.success(
            f"Processamento concluído! {len(all_data)} arquivo(s) válidos e {len(dados_combinados)} registros encontrados.")

        col1, col2, col3 = st.columns(3)
        arquivos_options = sorted(dados_combinados['Arquivo'].unique())
        filiais_options = sorted(dados_combinados['Filial'].unique())
        categorias_options = sorted(dados_combinados['Categoria'].unique())
        with col1:
            selecao_arquivo = st.multiselect(
                "Arquivo:", ["Selecionar Todos"] + arquivos_options, default="Selecionar Todos")
        with col2:
            selecao_filial = st.multiselect(
                "Filial:", ["Selecionar Todos"] + filiais_options, default="Selecionar Todos")
        with col3:
            selecao_categoria = st.multiselect(
                "Categoria:", ["Selecionar Todos"] + categorias_options, default="Selecionar Todos")

        filtro_arquivo = arquivos_options if "Selecionar Todos" in selecao_arquivo else selecao_arquivo
        filtro_filial = filiais_options if "Selecionar Todos" in selecao_filial else selecao_filial
        filtro_categoria = categorias_options if "Selecionar Todos" in selecao_categoria else selecao_categoria
        dados_filtrados = dados_combinados[
            (dados_combinados['Arquivo'].isin(filtro_arquivo)) &
            (dados_combinados['Filial'].isin(filtro_filial)) &
            (dados_combinados['Categoria'].isin(filtro_categoria))
        ]

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Registros Filtrados", f"{len(dados_filtrados):,}")
        col2.metric("Valor Total Atualizado", formatar_valor(
            dados_filtrados["Valor Atualizado"].sum()))
        col3.metric("Depreciação Acumulada", formatar_valor(
            dados_filtrados["Deprec. Acumulada"].sum()))
        col4.metric("Valor Residual Total", formatar_valor(
            dados_filtrados["Valor Residual"].sum()))

        tab1, tab2, tab3 = st.tabs(
            ["Dados Detalhados", "Análise por Filial", "Análise por Categoria"])
        with tab1:
            df_display = dados_filtrados.copy()

            colunas_para_exibir = [
                'Arquivo', 'Filial', 'Categoria', 'C Custo', 'Cod Base Bem', 'Codigo Item', 'Tipo Ativo',
                'Descr. Sint.', 'Tipo Depr.', 'Dt.Aquisicao', 'Data Baixa',
                'Quantidade', 'Num.Plaqueta', 'Item Despesa', 'ClVl Despesa',
                'Vl Ampliac.1', 'Valor Original', 'Valor Atualizado', 'Deprec. no mes',
                'Deprec. no Exerc.', 'Deprec. Acumulada', 'Valor Residual', 'Corre Mes M1', 'Corre Bal M1',
                'Corr Acum M1', 'Cor Dep Mes', 'Cor Dep Exer', 'Cor Dep Acum'
            ]

            colunas_monetarias = [
                'Vl Ampliac.1', 'Valor Original', 'Valor Atualizado', 'Deprec. no mes',
                'Deprec. no Exerc.', 'Deprec. Acumulada', 'Valor Residual', 'Corre Mes M1', 'Corre Bal M1',
                'Corr Acum M1', 'Cor Dep Mes', 'Cor Dep Exer', 'Cor Dep Acum'
            ]
            for col in colunas_monetarias:
                if col in df_display.columns:
                    df_display[col] = df_display[col].apply(formatar_valor)

            colunas_existentes = [
                col for col in colunas_para_exibir if col in df_display.columns]
            st.dataframe(df_display[colunas_existentes],
                         use_container_width=True, height=500)

        with tab2:
            analise_filial = dados_filtrados.groupby('Filial').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor Atualizado', 'sum')).reset_index()
            analise_filial['Valor_Total'] = analise_filial['Valor_Total'].apply(
                formatar_valor)
            st.dataframe(analise_filial, use_container_width=True)
        with tab3:
            analise_categoria = dados_filtrados.groupby('Categoria').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor Atualizado', 'sum')).reset_index()
            analise_categoria['Valor_Total'] = analise_categoria['Valor_Total'].apply(
                formatar_valor)
            st.dataframe(analise_categoria, use_container_width=True)

        st.markdown("---")
        st.header("Gráfico Interativo")

        opcoes_eixo_y = ["Valor Atualizado",
                         "Deprec. Acumulada", "Valor Residual"]
        col_graf1, col_graf2, col_graf3 = st.columns(3)
        with col_graf1:
            tipo_grafico = st.selectbox("Escolha o Tipo de Gráfico:", [
                                        "Barras", "Pizza", "Linhas"])
        with col_graf2:
            eixo_x = st.selectbox("Agrupar por (Eixo X):", [
                                  "Filial", "Categoria", "Arquivo"], key="eixo_x_selectbox")
        with col_graf3:
            if tipo_grafico == "Pizza":
                eixos_y = st.selectbox(
                    "Analisar Valor (Eixo Y):", opcoes_eixo_y, index=0)
                eixos_y = [eixos_y]
            else:
                eixos_y = st.multiselect("Analisar Valores (Eixo Y):", opcoes_eixo_y, default=[
                                         "Valor Atualizado", "Valor Residual"])

        if not dados_filtrados.empty and eixo_x and eixos_y:
            dados_agrupados = dados_filtrados.groupby(
                eixo_x)[eixos_y].sum().reset_index()

            fig_plotly = None
            if tipo_grafico == "Barras":
                dados_grafico_melted = pd.melt(dados_agrupados, id_vars=[
                                               eixo_x], value_vars=eixos_y, var_name='Métrica', value_name='Valor')
                fig_plotly = px.bar(dados_grafico_melted, x=eixo_x, y='Valor',
                                    color='Métrica', text_auto='.2s', barmode='group')
                fig_plotly.update_traces(textposition='outside')
            elif tipo_grafico == "Linhas":
                dados_grafico_melted = pd.melt(dados_agrupados, id_vars=[
                                               eixo_x], value_vars=eixos_y, var_name='Métrica', value_name='Valor')
                fig_plotly = px.line(
                    dados_grafico_melted, x=eixo_x, y='Valor', color='Métrica', markers=True)
            elif tipo_grafico == "Pizza":
                metrica_unica = eixos_y[0]
                fig_plotly = px.pie(
                    dados_agrupados, names=eixo_x, values=metrica_unica, hole=0.3)
                fig_plotly.update_traces(
                    textposition='outside', textinfo='percent+label')

            if fig_plotly:
                fig_plotly.update_layout(title=f'Análise de {", ".join(eixos_y)} por {eixo_x}', uniformtext_minsize=8, uniformtext_mode='hide', margin=dict(
                    t=80, b=50), plot_bgcolor='rgba(0,0,0,0)', legend_title_text='')
                st.plotly_chart(fig_plotly, use_container_width=True)

                st.session_state.dados_grafico = dados_agrupados
                st.session_state.tipo_grafico = tipo_grafico
                st.session_state.eixo_x = eixo_x
                st.session_state.eixos_y = eixos_y
            else:
                st.session_state.dados_grafico = None
        else:
            st.session_state.dados_grafico = None

        st.markdown("---")
        st.header("Exportar Relatório")

        col_download1, col_download2 = st.columns(2)

        with col_download1:
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                colunas_para_exportar = [
                    col for col in colunas_para_exibir if col in dados_filtrados.columns]
                dados_filtrados[colunas_para_exportar].to_excel(
                    writer, sheet_name='Dados_Filtrados', index=False)
            st.download_button(
                label="📥 Baixar Relatório em Excel",
                data=output_excel.getvalue(),
                file_name="relatorio_ativos_filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_download2:
            if st.session_state.dados_grafico is not None:
                pdf_buffer = BytesIO()
                criar_pdf_completo(
                    pdf_buffer,
                    dados_filtrados,
                    st.session_state.dados_grafico,
                    st.session_state.tipo_grafico,
                    st.session_state.eixo_x,
                    st.session_state.eixos_y
                )
                st.download_button(
                    label="📄 Baixar Relatório Completo (PDF)",
                    data=pdf_buffer.getvalue(),
                    file_name="relatorio_completo.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.download_button(
                    label="📄 Baixar Relatório Completo (PDF)",
                    data=b'',
                    disabled=True,
                    use_container_width=True,
                    help="Gere um gráfico na tela para habilitar o download do PDF."
                )

    if errors:
        st.warning("Alguns arquivos apresentaram problemas:", icon="❗")
        for error in errors:
            st.error(error)
else:
    st.info("Aguardando o upload dos arquivos para iniciar o processamento.")

st.markdown("---")
st.caption("Desenvolvido para General Water | v35.0 (Final) - Suporte via Teams")
