import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from segmentation import (
    gerar_publico, load_spot_pedidos, load_leadscore, load_rd_station, FAMILIAS,
)

st.set_page_config(page_title="Publico Segmentado - Clube do Malte", layout="wide")

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)
SPOT_PATH = DATA_DIR / "spot_pedidos.xlsx"
LEAD_PATH = DATA_DIR / "leadscore.xlsx"
RD_PATH = DATA_DIR / "rd_station.csv"

st.title("Gerador de Publico Segmentado")
st.caption("Processo padrao: Fase 1 (produto exato) + fallback por similaridade (Fase 2), "
           "com prioridade extra para quem navegou recentemente no site, ranqueado por Score RFV equilibrado.")


def _fmt_ts(path: Path) -> str:
    if not path.exists():
        return "nunca atualizada"
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M")


@st.cache_data(show_spinner="Carregando base de pedidos (pode levar ~1 min na primeira vez)...")
def _load_spot_cached(path_str, _mtime):
    return load_spot_pedidos(path_str)


@st.cache_data(show_spinner="Carregando Lead Score (pode levar ~1 min na primeira vez)...")
def _load_lead_cached(path_str, _mtime):
    return load_leadscore(path_str)


@st.cache_data(show_spinner="Carregando dados da RD Station...")
def _load_rd_cached(path_str, _mtime):
    return load_rd_station(path_str)


def _base_card(label, path: Path, key_prefix, help_text):
    exists = path.exists()
    if exists:
        st.success(f"{label}: OK — atualizada em {_fmt_ts(path)}")
    else:
        st.warning(f"{label}: base ainda nao carregada")
    with st.expander(f"Atualizar {label.lower()}"):
        st.caption(help_text)
        up = st.file_uploader(f"Nova versao ({label})", type=["xlsx"], key=f"{key_prefix}_up")
        if up is not None:
            path.write_bytes(up.getbuffer())
            st.cache_data.clear()
            st.success("Base atualizada nesta sessao. Para tornar permanente (sobreviver a "
                       "reinicios do app), suba este mesmo arquivo na pasta data/ do repositorio no GitHub.")
            st.rerun()


with st.sidebar:
    st.header("1. Bases de dados")

    st.subheader("Pedidos (10 anos)")
    _base_card("Base de pedidos", SPOT_PATH, "spot",
               "Fica fixa no app entre uma geracao e outra. So suba de novo quando tiver uma exportacao mais recente.")

    st.subheader("Lead Score")
    _base_card("Lead Score", LEAD_PATH, "lead",
               "Fica fixa no app entre uma geracao e outra. So suba de novo quando tiver uma exportacao mais recente.")

    st.subheader("RD Station - Navegacao (semanal)")
    if RD_PATH.exists():
        st.success(f"RD Station: OK — atualizada em {_fmt_ts(RD_PATH)}")
    else:
        st.info("RD Station: opcional. Sem ela, o app roda normalmente sem o criterio de navegacao.")
    rd_up = st.file_uploader("Subir planilha/CSV de navegacao da RD Station", type=["csv", "xlsx"], key="rd_up")
    if rd_up is not None:
        suffix = ".xlsx" if rd_up.name.lower().endswith("xlsx") else ".csv"
        rd_target = DATA_DIR / f"rd_station{suffix}"
        rd_target.write_bytes(rd_up.getbuffer())
        if rd_target != RD_PATH and RD_PATH.exists():
            RD_PATH.unlink()
        st.cache_data.clear()
        st.success("Dados da RD Station atualizados.")
        st.rerun()

    st.header("2. Parametros do processo")
    tamanho = st.number_input("Tamanho do publico", min_value=50, max_value=2000, value=400, step=50)
    min_pedidos = st.number_input("Min. pedidos totais (criterio 3)", min_value=1, max_value=10, value=3)
    dias_exclusao = st.number_input("Dias sem comprar o produto exato (criterio 4)", min_value=0, max_value=180, value=30)
    limite_fallback = st.number_input("Gatilho do fallback (elegiveis min. na Fase 1)", min_value=10, max_value=1000, value=100)
    dias_navegacao = st.number_input("Dias p/ considerar navegacao recente (RD Station)", min_value=30, max_value=720, value=360)

st.header("Oferta")
col1, col2 = st.columns(2)
with col1:
    nome_oferta = st.text_input("Nome da oferta", placeholder="Kit de Cervejas Bitburger - Compre 4 e Leve 7")
    link_oferta = st.text_input("Link do produto (opcional, so pra referencia)", placeholder="https://www.clubedomalte.com.br/produto/...")
    exact_kw = st.text_input("Palavra(s)-chave do PRODUTO EXATO", placeholder="bitburger",
                              help="Como o produto aparece no historico de pedidos. Separe varias palavras por virgula.")
with col2:
    similar_kw = st.text_input("Palavra(s)-chave de ESTILO SIMILAR", placeholder="pils, pilsen, lager, helles",
                                help="Estilos parecidos, usados no criterio 2 e na Camada A do fallback.")
    familia = st.selectbox("Familia do produto (Camada B do fallback)", FAMILIAS)

rodar = st.button("Gerar publico", type="primary", use_container_width=True)

if rodar:
    if not SPOT_PATH.exists() or not LEAD_PATH.exists():
        st.error("Faltam bases obrigatorias. Suba a base de Pedidos e o Lead Score na barra lateral antes de gerar.")
    elif not exact_kw:
        st.error("Informe ao menos uma palavra-chave do produto exato (mesmo em lancamentos - "
                  "use o nome que o produto vai ter).")
    else:
        spot_df = _load_spot_cached(str(SPOT_PATH), SPOT_PATH.stat().st_mtime)
        lead_master = _load_lead_cached(str(LEAD_PATH), LEAD_PATH.stat().st_mtime)
        rd_df = None
        if RD_PATH.exists():
            rd_df = _load_rd_cached(str(RD_PATH), RD_PATH.stat().st_mtime)

        with st.spinner("Calculando o publico..."):
            resultado = gerar_publico(
                spot_df, lead_master,
                exact_keywords_raw=exact_kw,
                similar_keywords_raw=similar_kw,
                familia=familia,
                tamanho=int(tamanho),
                min_pedidos=int(min_pedidos),
                dias_exclusao=int(dias_exclusao),
                limite_fallback=int(limite_fallback),
                rd_df=rd_df,
                dias_navegacao=int(dias_navegacao),
            )

        tabela = resultado['tabela']
        funil = resultado['funil']
        modo = resultado['modo']

        modo_label = {
            'fase1': 'Fase 1 (produto com historico suficiente)',
            'fallback_camada_a': 'Fallback acionado - Camada A (estilo similar)',
            'fallback_camada_a_b': 'Fallback acionado - Camadas A + B (estilo + familia)',
        }[modo]
        aviso_rd = "" if rd_df is not None else " (sem dado de navegacao da RD Station nesta rodada)"
        st.success(f"Publico gerado: {len(tabela)} contatos. Modo: {modo_label}{aviso_rd}")

        st.subheader("Funil de aplicacao dos criterios")
        funil_df = pd.DataFrame(funil, columns=["Etapa", "Qtd. de clientes"])
        st.dataframe(funil_df, hide_index=True, use_container_width=True)

        if 'Origem' in tabela.columns:
            st.subheader("Composicao por origem")
            st.dataframe(tabela['Origem'].value_counts().rename_axis('Origem').reset_index(name='Qtd.'),
                         hide_index=True, use_container_width=True)

        st.subheader("Publico final")
        st.dataframe(tabela, hide_index=True, use_container_width=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            tabela.to_excel(writer, sheet_name="Publico", index=False)
            funil_df.to_excel(writer, sheet_name="Funil", index=False)
        buf.seek(0)

        fname = f"Publico_{(nome_oferta or 'oferta').strip().replace(' ', '_')[:40]}.xlsx"
        st.download_button("Baixar Excel", data=buf, file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)
else:
    st.info("As bases de Pedidos e Lead Score ficam salvas no app — so precisa subir de novo quando quiser "
            "atualizar. RD Station e opcional, pensada pra subir toda semana. Preencha os dados da oferta e "
            "clique em 'Gerar publico'.")
