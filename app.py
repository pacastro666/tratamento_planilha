import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

# ----------------------------
# Utilidades
# ----------------------------
def extract_numbers_from_text(text: str):
    if pd.isna(text):
        return []
    return [int(n) for n in re.findall(r"\d+", str(text))]

def build_number_to_consultant_map(df_reat10_raw: pd.DataFrame):
    # Espera-se:
    # Coluna A (index 0): CONSULTOR
    # Coluna B (index 1): PEDIDOS (sequências com números)
    consultor_col = df_reat10_raw.iloc[:, 0].astype(str)
    pedidos_col = df_reat10_raw.iloc[:, 1]

    num_to_consultants = {}
    for cons, pedidos in zip(consultor_col, pedidos_col):
        if pd.isna(pedidos):
            continue
        for n in extract_numbers_from_text(pedidos):
            s = num_to_consultants.setdefault(n, set())
            s.add(cons)
    return num_to_consultants

def create_or_update_extracted_sheet(wb: Workbook, sheet_source: str = "REAT-10", sheet_out: str = "EXTRAIDOS_REAT10"):
    # Ler REAT-10 para mapear números -> consultores
    ws_src = wb[sheet_source]
    last_row_src = ws_src.max_row

    # Carregar REAT-10 em DataFrame bruto (sem header)
    # Como estamos em memória com openpyxl, vamos montar um DF manualmente
    data = []
    for r in ws_src.iter_rows(values_only=True):
        data.append(list(r))
    df_raw = pd.DataFrame(data)

    # Mapear número -> consultores
    num_to_cons = build_number_to_consultant_map(df_raw)
    # Números únicos em ordem de primeira aparição
    unique_numbers = list(num_to_cons.keys())

    # (Re)criar aba de saída
    if sheet_out in wb.sheetnames:
        ws_out = wb[sheet_out]
        wb.remove(ws_out)
    ws_out = wb.create_sheet(title=sheet_out)

    # Cabeçalhos
    ws_out.cell(row=1, column=1, value="NUMERO")
    ws_out.cell(row=1, column=2, value="ENCONTRADO_REAT10")
    ws_out.cell(row=1, column=3, value="OCORRENCIAS_REAT10")
    ws_out.cell(row=1, column=4, value="CONSULTORES")

    # Intervalo de busca para a fórmula
    lookup_range = f"'{sheet_source}'!$B$1:$B${last_row_src}"

    # Preencher linhas
    for i, num in enumerate(unique_numbers, start=2):
        ws_out.cell(row=i, column=1, value=num)

        # Fórmula que considera números dentro de sequências com '-'
        # =SUMPRODUCT(--ISNUMBER(SEARCH("-"&A2&"-","-"&SUBSTITUTE('REAT-10'!$B$1:$B$28," ","")&"-")))>0
        formula = f'=SUMPRODUCT(--ISNUMBER(SEARCH("-"&A{i}&"-","-"&SUBSTITUTE({lookup_range}," ","")&"-")))>0'
        ws_out.cell(row=i, column=2, value=formula)

        consultants = sorted(
            x for x in num_to_cons[num]
            if x and x.lower() not in ("nan", "consultor")
        )
        ws_out.cell(row=i, column=3, value=len(consultants))
        ws_out.cell(row=i, column=4, value="; ".join(consultants))

    # Estética básica
    ws_out.auto_filter.ref = f"A1:D{len(unique_numbers)+1}"
    ws_out.column_dimensions['A'].width = 16
    ws_out.column_dimensions['B'].width = 22
    ws_out.column_dimensions['C'].width = 22
    ws_out.column_dimensions['D'].width = 50

    return unique_numbers, last_row_src  # para exibição no app

def get_dataframe_from_sheet(wb: Workbook, sheet_name: str):
    ws = wb[sheet_name]
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))
    # Detectar header da primeira linha
    if len(data) == 0:
        return pd.DataFrame()
    # Se primeira linha parece cabeçalho (strings), usar como header
    header = data[0]
    df = pd.DataFrame(data[1:], columns=header)
    return df

# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="REAT-10 Extrator", layout="wide")
st.title("Extrator de Números da REAT-10")

with st.sidebar:
    st.header("Importar Planilha")
    uploaded = st.file_uploader("Selecione o arquivo .xlsx", type=["xlsx"])
    process_btn = st.button("Processar e atualizar aba EXTRAIDOS_REAT10")

# Estado
if "workbook_bytes" not in st.session_state:
    st.session_state.workbook_bytes = None
if "report" not in st.session_state:
    st.session_state.report = None

tab1, tab2, tab3 = st.tabs(["Visualizar REAT-10", "EXTRAIDOS_REAT10 (visualização)", "Baixar arquivo modificado"])

wb = None

if uploaded is not None:
    # Carregar workbook na memória
    content = uploaded.read()
    st.session_state.workbook_bytes = content
    bio = io.BytesIO(content)
    wb = load_workbook(bio)

    with tab1:
        st.subheader("Dados da aba REAT-10")
        if "REAT-10" in wb.sheetnames:
            df_reat10 = get_dataframe_from_sheet(wb, "REAT-10")
            st.dataframe(df_reat10, use_container_width=True, height=500)
        else:
            st.warning("A aba 'REAT-10' não foi encontrada no arquivo.")

    if process_btn:
        if wb is None:
            st.error("Erro ao carregar workbook.")
        elif "REAT-10" not in wb.sheetnames:
            st.error("A aba 'REAT-10' não foi encontrada. Verifique o arquivo.")
        else:
            # Processar e atualizar aba
            unique_numbers, last_row_src = create_or_update_extracted_sheet(wb, "REAT-10", "EXTRAIDOS_REAT10")

            # Salvar workbook modificado em memória
            out_buf = io.BytesIO()
            wb.save(out_buf)
            st.session_state.workbook_bytes = out_buf.getvalue()

            st.session_state.report = {
                "unique_count": len(unique_numbers),
                "last_row_src": last_row_src
            }
            st.success(f"Aba EXTRAIDOS_REAT10 atualizada. Números únicos: {len(unique_numbers)} | Última linha da REAT-10: {last_row_src}")

    # Visualização da aba gerada (se já foi processada nesta sessão)
    with tab2:
        st.subheader("Prévia da aba EXTRAIDOS_REAT10")
        if wb is not None and "EXTRAIDOS_REAT10" in wb.sheetnames:
            df_out = get_dataframe_from_sheet(wb, "EXTRAIDOS_REAT10")
            st.dataframe(df_out, use_container_width=True, height=500)
        else:
            st.info("Ainda não há aba EXTRAIDOS_REAT10 para visualizar. Clique em 'Processar e atualizar aba EXTRAIDOS_REAT10'.")

    # Download do arquivo atualizado
    with tab3:
        st.subheader("Baixar arquivo Excel com a aba EXTRAIDOS_REAT10")
        if st.session_state.workbook_bytes:
            st.download_button(
                label="Baixar .xlsx atualizado",
                data=st.session_state.workbook_bytes,
                file_name="Resultados2025_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            if st.session_state.report:
                st.caption(f"Números únicos: {st.session_state.report['unique_count']} | Última linha REAT-10: {st.session_state.report['last_row_src']}")
        else:
            st.info("Após processar, o arquivo atualizado ficará disponível para download aqui.")
else:
    st.info("Faça o upload de uma planilha .xlsx na barra lateral para começar.")
