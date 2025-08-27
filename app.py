import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
from PIL import Image
import plotly.express as px
import urllib.parse  # Importa a biblioteca para formatar a URL

warnings.filterwarnings('ignore')

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard de Ativos Contábeis",
    page_icon=" ",
    layout="wide"
)

# --- FUNÇÕES DE LÓGICA (sem alteração) ---


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


def processar_planilha(file):
    try:
        xl = pd.ExcelFile(file)
        dados_processados = []
        for sheet_name in xl.sheet_names:
            sheet_df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            categoria_atual, filial_atual = "Não Identificado", "Não Identificado"
            for _, row in sheet_df.iterrows():
                if pd.notna(row.iloc[0]) and str(row.iloc[0]).startswith('1.2.3.'):
                    categoria_atual = str(row.iloc[1]).strip() if pd.notna(
                        row.iloc[1]) else str(row.iloc[0]).strip()
                elif pd.notna(row.iloc[0]) and 'Filial :' in str(row.iloc[0]):
                    nome_extraido = str(row.iloc[0]).split(
                        'Filial :')[-1].split(' - ')[-1].strip()
                    filial_atual = padronizar_nome_filial(nome_extraido)
                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == 'R$':
                    valores = [converter_valor(v) for v in row.iloc[1:8]]
                    valor_atualizado = valores[2] if len(valores) > 2 else 0
                    deprec_acumulada = valores[5] if len(valores) > 5 else 0
                    if valor_atualizado > 0:
                        dados_processados.append({
                            'Arquivo': file.name, 'Filial': filial_atual, 'Categoria': categoria_atual,
                            'Valor Original': valores[1] if len(valores) > 1 else 0,
                            'Valor Atualizado': valor_atualizado,
                            'Deprec. no mês': valores[3] if len(valores) > 3 else 0,
                            'Deprec. no Exercício': valores[4] if len(valores) > 4 else 0,
                            'Deprec. Acumulada': deprec_acumulada,
                            'Valor Residual': valor_atualizado - deprec_acumulada
                        })
        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            return corrigir_filiais_nao_identificadas(df_final), None
        return None, f"Nenhum dado relevante encontrado em {file.name}."
    except Exception as e:
        return None, f"Erro crítico ao processar {file.name}: {e}"


# --- ESTRUTURA DA APLICAÇÃO ---
st.title("Dashboard de Ativos Contábeis")

with st.sidebar:
    try:
        st.image("logo_GW.png", width=200)
    except Exception:
        st.title("General Water")

    st.header("ℹ️ Instruções")
    st.info("1. **Carregue** os arquivos.\n2. **Aguarde** o processamento.\n3. **Filtre** e analise os dados.\n4. **Explore** os gráficos interativos.\n5. **Baixe** o relatório.")

    # --- BOTÃO DE AJUDA DO TEAMS ---
    st.header("💬 Ajuda & Suporte")

    # **IMPORTANTE**: Substitua pelos e-mails reais
    email1 = "bruce@generalwater.com.br"
    email2 = "nathalia.vidal@generalwater.com.br"

    # Mensagem opcional que aparecerá no chat
    mensagem_inicial = "Olá, preciso de ajuda com o Dashboard de Ativos Contábeis."

    # Formata a URL para ser segura (substitui espaços por %20, etc.)
    link_teams = f"https://teams.microsoft.com/l/chat/0/0?users={email1},{email2}&message={urllib.parse.quote(mensagem_inicial)}"

    # Cria o link clicável usando Markdown
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
            f"✅ Processamento concluído! {len(all_data)} arquivo(s) válidos.")

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
            for col in ['Valor Original', 'Valor Atualizado', 'Deprec. no mês', 'Deprec. no Exercício', 'Deprec. Acumulada', 'Valor Residual']:
                df_display[col] = df_display[col].apply(formatar_valor)
            st.dataframe(df_display, use_container_width=True, height=500)
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

        opcoes_eixo_y = ["Valor Atualizado",
                         "Deprec. Acumulada", "Valor Residual"]

        col_graf1, col_graf2 = st.columns(2)
        with col_graf1:
            eixo_x = st.selectbox("Agrupar por (Eixo X):", [
                                  "Filial", "Categoria", "Arquivo"], key="eixo_x_selectbox")
        with col_graf2:
            eixos_y = st.multiselect("Analisar Valores (Eixo Y):", opcoes_eixo_y, default=[
                                     "Valor Atualizado", "Valor Residual"])

        if not dados_filtrados.empty and eixo_x:
            opcoes_foco = ["Mostrar Todos"] + \
                sorted(dados_filtrados[eixo_x].unique().tolist())
            foco_selecionado = st.selectbox(
                f"Focar em um(a) {eixo_x} específico(a) (opcional):", opcoes_foco)

        if not dados_filtrados.empty and eixo_x and eixos_y:
            dados_para_grafico = dados_filtrados.copy()
            if foco_selecionado != "Mostrar Todos":
                dados_para_grafico = dados_para_grafico[dados_para_grafico[eixo_x]
                                                        == foco_selecionado]

            dados_agrupados = dados_para_grafico.groupby(
                eixo_x)[eixos_y].sum().reset_index()
            dados_grafico_melted = pd.melt(dados_agrupados, id_vars=[
                                           eixo_x], value_vars=eixos_y, var_name='Métrica', value_name='Valor')

            titulo = f"Comparativo de Métricas por {eixo_x}"
            if foco_selecionado != "Mostrar Todos":
                titulo = f"Análise Focada em: {foco_selecionado}"

            fig = px.bar(
                dados_grafico_melted, x=eixo_x, y='Valor', color='Métrica', title=titulo,
                labels={eixo_x: eixo_x, 'Valor': "Soma dos Valores",
                        'Métrica': "Métrica Financeira"},
                text_auto='.2s', barmode='group'
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(
                uniformtext_minsize=8, uniformtext_mode='hide',
                margin=dict(t=80, b=50), plot_bgcolor='rgba(0,0,0,0)',
                legend_title_text=''
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning(
                "Selecione uma opção para 'Agrupar por' e pelo menos uma 'Métrica' para gerar o gráfico.")

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            dados_filtrados.to_excel(
                writer, sheet_name='Dados_Filtrados', index=False)
        st.download_button(label="📥 Baixar Relatório Filtrado (Excel)", data=output.getvalue(
        ), file_name="relatorio_ativos_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if errors:
        st.warning("⚠️ Alguns arquivos apresentaram problemas:", icon="❗")
        for error in errors:
            st.error(error)
else:
    st.info("👆 Aguardando o upload dos arquivos para iniciar o processamento.")

st.markdown("---")
st.caption("Desenvolvido para General Water | v16.0 - Suporte via Teams")
