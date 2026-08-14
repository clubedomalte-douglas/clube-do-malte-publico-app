# Gerador de Público Segmentado — Clube do Malte

App web (Streamlit) que roda o processo padrão de segmentação combinado com o Douglas:

**Fase 1** — comprou o produto exato + estilo similar + 3 ou mais pedidos totais + não comprou
nos últimos 30 dias → cruza com opt-in WhatsApp → ranqueia por Score RFV equilibrado.

**Fase 2 (fallback automático)** — se a Fase 1 render menos de 100 elegíveis (comum em
lançamentos), o app completa a base com Camada A (estilo similar, sem exigir o produto exato)
e, se ainda faltar gente, Camada B (família genérica: Leves, Lupuladas, Frutadas, Maltadas,
Complexas, Torradas, Azedas, Sem Álcool, Kits).

Toda vez que o app roda, ele usa os dois arquivos que você subir na hora (Pedidos + Lead Score) —
não guarda nada de uma rodada pra outra, então pra ficar "diário" basta subir a versão mais
recente dos arquivos antes de gerar o público.

## Arquivos deste pacote

- `app.py` — interface (upload dos arquivos, campos da oferta, botão gerar, download do Excel)
- `segmentation.py` — o motor de cálculo (a lógica dos 4 critérios + fallback + RFV)
- `requirements.txt` — dependências Python

## Opção 1 — Rodar no seu computador (mais simples, sem hospedagem)

Só funciona no computador onde você rodar o comando — ainda não é acessível de outro lugar.
Bom pra testar antes de publicar de verdade.

1. Instale o [Python 3.10+](https://www.python.org/downloads/) se ainda não tiver.
2. Abra o terminal (PowerShell/CMD) nesta pasta e rode:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Abre sozinho no navegador em `http://localhost:8501`.

## Opção 2 — Publicar num link (acessível por você e o time, de qualquer computador)

**Streamlit Community Cloud** é a forma mais rápida e gratuita de publicar isso como um site
com endereço próprio, sem precisar mexer em servidor.

Passo a passo:

1. Crie uma conta gratuita no GitHub (se não tiver): https://github.com/signup
2. Crie um repositório novo (pode ser privado) e suba estes 3 arquivos (`app.py`,
   `segmentation.py`, `requirements.txt`) — dá pra fazer isso direto pelo site do GitHub,
   arrastando os arquivos, sem precisar de linha de comando.
3. Crie uma conta gratuita em https://streamlit.io/cloud (pode entrar direto com o GitHub).
4. Clique em "New app", escolha o repositório que você criou, aponte o arquivo principal
   como `app.py` e clique em "Deploy".
5. Em alguns minutos o Streamlit te dá um link (tipo
   `https://seu-app.streamlit.app`) — esse é o link que você e o time vão usar, de qualquer
   computador com internet.
6. Se o repositório for privado, o Streamlit Cloud te deixa convidar só quem você quiser (por
   e-mail) pra acessar o app — assim fica restrito ao time.

Como os dados continuam sendo upload manual (você escolheu não automatizar isso por agora),
não tem nenhuma informação sensível guardada no repositório do GitHub — só o código.

### Alternativas ao Streamlit Community Cloud

Se no futuro quiser algo com marca própria, sem o domínio `.streamlit.app`, dá pra hospedar o
mesmo `app.py` em serviços como Render.com ou Railway.app (ambos têm plano gratuito/baixo custo
e aceitam apps Streamlit com pouquíssima configuração extra). O código não muda.

## Quando migrar para atualização automática dos dados

Hoje o fluxo é: você exporta os arquivos de Pedidos e Lead Score e sobe no app na hora de gerar
o público. Quando fizer sentido automatizar (puxar direto da RD Station / plataforma da loja
todo dia, sem upload manual), o `segmentation.py` já está pronto pra receber os dados de outra
fonte — só troca a forma como `spot_df` e `lead_master` chegam até a função `gerar_publico`.
