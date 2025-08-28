
import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
from fpdf import FPDF
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

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
            filial_atual = "Não Identificado"
            conta_contabil_atual = "Não Identificado"
            descricao_conta_atual = "Não Identificado"

            for idx, row in sheet_df.iterrows():
                # Identificar Filial
                if pd.notna(row.iloc[0]) and 'Filial :' in str(row.iloc[0]):
                    nome_extraido = str(row.iloc[0]).split(
                        'Filial :')[-1].split(' - ')[-1].strip()
                    filial_atual = padronizar_nome_filial(nome_extraido)
                
                # Identificar Conta Contábil e Descrição da Conta
                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).startswith('1.2.3.'):
                    conta_contabil_atual = str(row.iloc[0]).strip()
                    descricao_conta_atual = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

                # Identificar linhas de dados (onde a Filial e o Código do Item estão presentes)
                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip().isdigit() and len(str(row.iloc[0]).strip()) == 6 and \
                     pd.notna(row.iloc[2]) and str(row.iloc[2]).strip().isdigit():
                    
                    # Extrair dados das colunas solicitadas
                    filial_item = str(row.iloc[0]).strip() # Coluna A
                    codigo_item = str(row.iloc[2]).strip() # Coluna C
                    data_aquisicao = pd.to_datetime(row.iloc[7], errors='coerce') # Coluna H
                    
                    # Novos campos solicitados
                    codigo_sub_item = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else "" # Coluna J
                    descricao_item = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "" # Coluna D

                    # Valores numéricos
                    # A planilha de exemplo mostra que os valores estão em linhas separadas, logo abaixo da linha de dados principais.
                    # Precisamos ajustar a lógica para capturar esses valores corretamente.
                    # Por enquanto, vamos usar valores padrão ou tentar inferir se possível.
                    # A estrutura atual do código já tenta pegar valores de uma linha 'R$'.
                    # Vou manter a lógica existente para 'R$' e tentar mapear para as novas colunas.
                    
                    # A lógica original pegava valores da linha 'R$' e os mapeava para 'Valor Original', 'Valor Atualizado', etc.
                    # A solicitação do usuário é para ter 'Valor original', 'Valor atualizado', 'Deprec. do mês', 'Deprec. do exercício', 'Deprec. Acumulada'
                    # A linha 'R$' na planilha de exemplo tem os valores em row.iloc[1] a row.iloc[7]
                    # Valor Original: row.iloc[1] da linha 'R$'
                    # Valor Atualizado: row.iloc[2] da linha 'R$'
                    # Deprec. do mês: row.iloc[3] da linha 'R$'
                    # Deprec. do exercício: row.iloc[4] da linha 'R$'
                    # Deprec. Acumulada: row.iloc[5] da linha 'R$'

                    # Filtrar por data de aquisição
                    if pd.notna(data_aquisicao) and data_aquisicao >= pd.to_datetime('1990-01-01'):
                        # A lógica para extrair os valores da linha 'R$' já existe e está logo abaixo.
                        # Vou adicionar um placeholder para os valores e preenchê-los na próxima iteração da linha 'R$'.
                        # Isso significa que a linha de dados e a linha de valores 'R$' precisam ser processadas em conjunto.
                        # Para simplificar, vou assumir que a linha 'R$' sempre segue a linha de dados.
                        # Isso pode ser um problema se a estrutura da planilha variar.
                        
                        # A abordagem atual do código é processar a linha 'R$' separadamente.
                        # Para as novas colunas, precisamos associar os valores da linha 'R$' à linha de dados.
                        # A melhor forma é capturar os dados da linha de dados e esperar pela linha 'R$'.
                        # Ou, se a linha 'R$' estiver sempre na próxima linha, podemos tentar ler a próxima linha.

                        # Para o propósito desta modificação, vou ajustar a lógica para extrair os valores da linha 'R$' e associá-los.
                        # A lógica atual do `processar_planilha` já faz isso, mas os nomes das colunas precisam ser ajustados.
                        # O problema é que a linha 'R$' é processada como um item separado no `dados_processados`.
                        # Precisamos garantir que os valores da linha 'R$' sejam associados à linha de dados correta.

                        # Vou modificar a estrutura para que a linha 'R$' preencha os valores da última linha de dados processada.
                        # Isso exigirá um ajuste na forma como `dados_processados` é construído.
                        
                        # Temporariamente, vou extrair os valores da linha 'R$' na mesma iteração, assumindo que eles estão na mesma linha de dados.
                        # Isso é uma simplificação e pode precisar de ajuste se a estrutura for mais complexa.
                        
                        # Revertendo para a lógica original de como os valores são extraídos, mas ajustando os nomes das colunas.
                        # A lógica atual do `processar_planilha` já extrai os valores da linha 'R$' e os mapeia.
                        # O que precisamos é garantir que as colunas solicitadas sejam preenchidas corretamente.
                        
                        # A solicitação é para ter as colunas na ordem:
                        # Filial, Conta contábil, Descrição da conta, Data de aquisição, Código do item, Código do sub item, Descrição do item, Valor original, Valor atualizado, Deprec. do mês, Deprec. do exercício, Deprec. Acumulada

                        # A lógica atual já extrai 'Filial', 'Categoria' (que é a Descrição da Conta), 'Valor Original', 'Valor Atualizado', 'Deprec. no mês', 'Deprec. no Exercício', 'Deprec. Acumulada'.
                        # Precisamos adicionar 'Conta contábil', 'Data de aquisição', 'Código do item', 'Código do sub item', 'Descrição do item'.
                        
                        # A 'Categoria' atual é a 'Descrição da conta'.
                        # 'Valor Original' e 'Valor Atualizado' já são extraídos.
                        # 'Deprec. no mês', 'Deprec. no Exercício', 'Deprec. Acumulada' também.

                        # O desafio é que a 'Conta contábil' e 'Descrição da conta' são extraídas de uma linha diferente ('1.2.3.').
                        # E os valores ('R$') são extraídos de outra linha.
                        # A 'Filial' também é extraída de uma linha diferente.

                        # A melhor abordagem é coletar todos os dados relevantes de cada tipo de linha e, em seguida, combiná-los.
                        # Vou reestruturar a função `processar_planilha` para ser mais robusta.

                        # Nova abordagem: Coletar informações de cabeçalho (Filial, Conta Contábil, Descrição da Conta) e depois processar as linhas de dados.
                        # A linha 'R$' contém os valores que precisam ser associados à última linha de dados encontrada.

                        # Vamos manter a estrutura de `dados_processados` como uma lista de dicionários.
                        # Cada dicionário representará uma linha final de dados.
                        
                        # A lógica atual já tenta fazer isso, mas a forma como os dados são agrupados é o problema.
                        # A 'categoria_atual' no código original é a 'Descrição da conta'.
                        # A 'filial_atual' é a 'Filial'.

                        # Vamos ajustar o dicionário `dados_processados.append` para incluir as novas colunas e reordenar.
                        # E garantir que os valores da linha 'R$' sejam corretamente associados.

                        # A lógica atual para a linha 'R$' é:
                        # elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == 'R$':
                        #     valores = [converter_valor(v) for v in row.iloc[1:8]]
                        #     valor_atualizado = valores[2] if len(valores) > 2 else 0
                        #     deprec_acumulada = valores[5] if len(valores) > 5 else 0
                        #     if valor_atualizado > 0:
                        #         dados_processados.append({
                        #             'Arquivo': file.name, 'Filial': filial_atual, 'Categoria': categoria_atual,
                        #             'Valor Original': valores[1] if len(valores) > 1 else 0,
                        #             'Valor Atualizado': valor_atualizado,
                        #             'Deprec. no mês': valores[3] if len(valores) > 3 else 0,
                        #             'Deprec. no Exercício': valores[4] if len(valores) > 4 else 0,
                        #             'Deprec. Acumulada': deprec_acumulada,
                        #             'Valor Residual': valor_atualizado - deprec_acumulada
                        #         })

                        # Isso significa que a linha 'R$' cria uma nova entrada em `dados_processados`.
                        # O que queremos é que a linha 'R$' *atualize* a última entrada de dados.

                        # Reestruturando `processar_planilha` para lidar com isso.
                        # Vamos manter um dicionário temporário para a linha de dados atual.
                        
                        # Resetando a lógica de processamento da planilha para ser mais sequencial e robusta.
                        # A ideia é que cada linha de dados (com Filial e Código do Item) seja uma entrada.
                        # E a linha 'R$' subsequente preencha os valores financeiros para essa entrada.

                        # Novo `processar_planilha`:
                        dados_processados = []
                        last_data_row_info = {}

                        for sheet_name in xl.sheet_names:
                            sheet_df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                            filial_atual = "Não Identificado"
                            conta_contabil_atual = "Não Identificado"
                            descricao_conta_atual = "Não Identificado"

                            for idx, row in sheet_df.iterrows():
                                # Identificar Filial
                                if pd.notna(row.iloc[0]) and 'Filial :' in str(row.iloc[0]):
                                    nome_extraido = str(row.iloc[0]).split(
                                        'Filial :')[-1].split(' - ')[-1].strip()
                                    filial_atual = padronizar_nome_filial(nome_extraido)
                                
                                # Identificar Conta Contábil e Descrição da Conta
                                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).startswith('1.2.3.'):
                                    conta_contabil_atual = str(row.iloc[0]).strip()
                                    descricao_conta_atual = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

                                # Identificar linhas de dados (onde a Filial e o Código do Item estão presentes)
                                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip().isdigit() and len(str(row.iloc[0]).strip()) == 6 and \
                                     pd.notna(row.iloc[2]) and str(row.iloc[2]).strip().isdigit():
                                    
                                    data_aquisicao = pd.to_datetime(row.iloc[7], errors='coerce') # Coluna H
                                    
                                    if pd.notna(data_aquisicao) and data_aquisicao >= pd.to_datetime('1990-01-01'):
                                        last_data_row_info = {
                                            'Arquivo': file.name,
                                            'Filial': filial_atual, # Coluna A
                                            'Conta contábil': conta_contabil_atual, # Coluna A
                                            'Descrição da conta': descricao_conta_atual, # Coluna B
                                            'Data de aquisição': data_aquisicao, # Coluna H
                                            'Código do item': str(row.iloc[2]).strip(), # Coluna C
                                            'Código do sub item': str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else "", # Coluna J
                                            'Descrição do item': str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "", # Coluna D
                                            'Valor original': 0.0,
                                            'Valor atualizado': 0.0,
                                            'Deprec. do mês': 0.0,
                                            'Deprec. do exercício': 0.0,
                                            'Deprec. Acumulada': 0.0,
                                            'Valor Residual': 0.0 # Será calculado depois
                                        }
                                        # Adiciona a linha de dados, os valores financeiros serão preenchidos pela próxima linha 'R$'
                                        dados_processados.append(last_data_row_info)
                                
                                # Identificar linha de valores 'R$'
                                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == 'R$':
                                    if last_data_row_info: # Se houver uma linha de dados anterior para associar
                                        valores = [converter_valor(v) for v in row.iloc[1:8]]
                                        
                                        # Atualiza a última entrada em dados_processados
                                        # Acessa o último dicionário adicionado e o modifica
                                        if dados_processados:
                                            dados_processados[-1]['Valor original'] = valores[1] if len(valores) > 1 else 0
                                            dados_processados[-1]['Valor atualizado'] = valores[2] if len(valores) > 2 else 0
                                            dados_processados[-1]['Deprec. do mês'] = valores[3] if len(valores) > 3 else 0
                                            dados_processados[-1]['Deprec. do exercício'] = valores[4] if len(valores) > 4 else 0
                                            dados_processados[-1]['Deprec. Acumulada'] = valores[5] if len(valores) > 5 else 0
                                            dados_processados[-1]['Valor Residual'] = dados_processados[-1]['Valor atualizado'] - dados_processados[-1]['Deprec. Acumulada']
                                        last_data_row_info = {} # Reset para a próxima linha de dados

        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            # Reordenar as colunas conforme solicitado
            colunas_ordenadas = [
                'Filial',
                'Conta contábil',
                'Descrição da conta',
                'Data de aquisição',
                'Código do item',
                'Código do sub item',
                'Descrição do item',
                'Valor original',
                'Valor atualizado',
                'Deprec. do mês',
                'Deprec. do exercício',
                'Deprec. Acumulada',
                'Valor Residual' # Manter o valor residual para cálculos internos
            ]
            df_final = df_final[colunas_ordenadas]
            return corrigir_filiais_nao_identificadas(df_final), None
        return None, f"Nenhum dado relevante encontrado em {file.name}."
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
                # Para Matplotlib, é melhor ter o eixo X como índice para plotagem de múltiplas colunas
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
    colunas_para_somar = ['Valor atualizado',
                          'Deprec. Acumulada', 'Valor Residual']
    df_agregado = df_filtrado.groupby(['Filial', 'Descrição da conta'])[
        colunas_para_somar].sum().reset_index()
    for col in colunas_para_somar:
        df_agregado[col] = df_agregado[col].apply(formatar_valor)
    col_widths = {'Filial': 60, 'Descrição da conta': 100, 'Valor atualizado': 35,
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
            f"Processamento concluído! {len(all_data)} arquivo(s) válidos.")

        col1, col2, col3 = st.columns(3)
        arquivos_options = sorted(dados_combinados['Arquivo'].unique())
        filiais_options = sorted(dados_combinados['Filial'].unique())
        # A categoria agora é 'Descrição da conta'
        categorias_options = sorted(dados_combinados['Descrição da conta'].unique())
        with col1:
            selecao_arquivo = st.multiselect(
                "Arquivo:", ["Selecionar Todos"] + arquivos_options, default="Selecionar Todos")
        with col2:
            selecao_filial = st.multiselect(
                "Filial:", ["Selecionar Todos"] + filiais_options, default="Selecionar Todos")
        with col3:
            selecao_categoria = st.multiselect(
                "Descrição da Conta:", ["Selecionar Todos"] + categorias_options, default="Selecionar Todos")

        filtro_arquivo = arquivos_options if "Selecionar Todos" in selecao_arquivo else selecao_arquivo
        filtro_filial = filiais_options if "Selecionar Todos" in selecao_filial else selecao_filial
        filtro_categoria = categorias_options if "Selecionar Todos" in selecao_categoria else selecao_categoria
        dados_filtrados = dados_combinados[
            (dados_combinados['Arquivo'].isin(filtro_arquivo)) &
            (dados_combinados['Filial'].isin(filtro_filial)) &
            (dados_combinados['Descrição da conta'].isin(filtro_categoria))
        ]

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Registros Filtrados", f"{len(dados_filtrados):,}")
        col2.metric("Valor Total Atualizado", formatar_valor(
            dados_filtrados["Valor atualizado"].sum()))
        col3.metric("Depreciação Acumulada", formatar_valor(
            dados_filtrados["Deprec. Acumulada"].sum()))
        col4.metric("Valor Residual Total", formatar_valor(
            dados_filtrados["Valor Residual"].sum()))

        # ### CÓDIGO DAS ABAS RESTAURADO ###
        tab1, tab2, tab3 = st.tabs(
            ["Dados Detalhados", "Análise por Filial", "Análise por Descrição da Conta"])
        with tab1:
            df_display = dados_filtrados.copy()
            for col in ['Valor original', 'Valor atualizado', 'Deprec. do mês', 'Deprec. do exercício', 'Deprec. Acumulada', 'Valor Residual']:
                df_display[col] = df_display[col].apply(formatar_valor)
            st.dataframe(df_display, use_container_width=True, height=500)
        with tab2:
            analise_filial = dados_filtrados.groupby('Filial').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor atualizado', 'sum')).reset_index()
            analise_filial['Valor_Total'] = analise_filial['Valor_Total'].apply(
                formatar_valor)
            st.dataframe(analise_filial, use_container_width=True)
        with tab3:
            analise_categoria = dados_filtrados.groupby('Descrição da conta').agg(Contagem=(
                'Arquivo', 'count'), Valor_Total=('Valor atualizado', 'sum')).reset_index()
            analise_categoria['Valor_Total'] = analise_categoria['Valor_Total'].apply(
                formatar_valor)
            st.dataframe(analise_categoria, use_container_width=True)
        # ### FIM DO BLOCO RESTAURADO ###

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
                                  "Filial", "Descrição da conta", "Arquivo"], key="eixo_x_selectbox")
        with col_graf3:
            if tipo_grafico == "Pizza":
                eixos_y = st.selectbox(
                    "Analisar Valor (Eixo Y):", opcoes_eixo_y, index=0)
                eixos_y = [eixos_y]
            else:
                eixos_y = st.multiselect("Analisar Valores (Eixo Y):", opcoes_eixo_y, default=[
                                         "Valor atualizado", "Valor Residual"])

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
                dados_filtrados.to_excel(
                    writer, sheet_name='Dados_Filtrados', index=False)
            st.download_button(
                label="Baixar Dados Filtrados (Excel)",
                data=output_excel.getvalue(),
                file_name="dados_ativos_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_download2:
            output_pdf = BytesIO()
            criar_pdf_completo(output_pdf, dados_filtrados, st.session_state.dados_grafico, st.session_state.tipo_grafico, st.session_state.eixo_x, st.session_state.eixos_y)
            st.download_button(
                label="Baixar Relatório (PDF)",
                data=output_pdf.getvalue(),
                file_name="relatorio_ativos.pdf",
                mime="application/pdf"
            )

    elif errors:
        for error_msg in errors:
            st.error(error_msg)
    else:
        st.info("Nenhum arquivo carregado ou dados processados.")

else:
    st.info("Por favor, carregue seus arquivos Excel para começar.")


