import io
import pandas as pd
import streamlit as st
from segmentation import gerar_publico, load_spot_pedidos, load_leadscore, FAMILIAS

st.set_page_config(page_title="Publico Segmentado - Clube do Malte", layout="wide")

st.title("Gerador de Publico Segmentado")
st.caption("Processo padrao: Fase 1 (produto exato) + fallback por similaridade (Fase 2), "
           "ranqueado por Score RFV equilibrado.")

with st.sidebar:
    st.header("1. Bases de dados")
    spot_file = st.file_uploader("SPOT_Pedidos_Clientes (.xlsx)", type=["xlsx"])
    lead_file = st.file_uploader("LeadScore_Final_Segmentos (.xlsx)", type=["xlsx"])
    st.caption("Suba a versao mais recente exportada. O app usa sempre o que estiver aqui.")

    st.header("2. Parametros do processo")
    tamanho = st.number_input("Tamanho do publico", min_value=50, max_value=2000, value=400, step=50)
    min_pedidos = st.number_input("Min. pedidos totais (criterio 3)", min_value=1, max_value=10, value=3)
    dias_exclusao = st.number_input("Dias sem comprar o produto exato (criterio 4)", min_value=0, max_value=180, value=30)
    limite_fallback = st.number_input("Gatilho do fallback (elegiveis min. na Fase 1)", min_value=10, max_value=1000, value=100)

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
    if not spot_file or not lead_file:
        st.error("Suba os dois arquivos (Pedidos e LeadScore) antes de gerar.")
    elif not exact_kw:
        st.error("Informe ao menos uma palavra-chave do produto exato (mesmo em lancamentos - "
                  "use o nome que o produto vai ter).")
    else:
        with st.spinner("Processando pedidos e calculando o publico..."):
            spot_df = load_spot_pedidos(spot_file)
            lead_master = load_leadscore(lead_file)
            resultado = gerar_publico(
                spot_df, lead_master,
                exact_keywords_raw=exact_kw,
                similar_keywords_raw=similar_kw,
                familia=familia,
                tamanho=int(tamanho),
                min_pedidos=int(min_pedidos),
                dias_exclusao=int(dias_exclusao),
                limite_fallback=int(limite_fallback),
            )

        tabela = resultado['tabela']
        funil = resultado['funil']
        modo = resultado['modo']

        modo_label = {
            'fase1': 'Fase 1 (produto com historico suficiente)',
            'fallback_camada_a': 'Fallback acionado - Camada A (estilo similar)',
            'fallback_camada_a_b': 'Fallback acionado - Camadas A + B (estilo + familia)',
        }[modo]
        st.success(f"Publico gerado: {len(tabela)} contatos. Modo: {modo_label}")

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
    st.info("Suba os dois arquivos, preencha os dados da oferta e clique em 'Gerar publico'.")
