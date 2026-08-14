# Gerador de Público Segmentado — Clube do Malte

App web (Streamlit) que roda o processo padrão de segmentação combinado com o Douglas:

**Fase 1** — comprou o produto exato + estilo similar + 3 ou mais pedidos totais + não comprou
nos últimos 30 dias → cruza com opt-in WhatsApp → ranqueia por Score RFV equilibrado.

**Fase 2 (fallback automático)** — se a Fase 1 render menos de 100 elegíveis (comum em
lançamentos), o app completa a base com Camada A (estilo similar, sem exigir o produto exato)
e, se ainda faltar gente, Camada B (família genérica: Leves, Lupuladas, Frutadas, Maltadas,
Complexas, Torradas, Azedas, Sem Álcool, Kits).

**Critério extra — navegação recente (RD Station)** — se você subir o export de navegação/
engajamento da RD Station, o app prioriza quem navegou no site nos últimos N dias (360 por
padrão, ajustável na barra lateral) dentro de cada camada, antes de aplicar o Score RFV.

## Bases de dados: fixas, com botão de atualizar

- **Base de Pedidos** e **Lead Score** ficam salvas dentro do app (pasta `data/`). Você só
  precisa subir de novo quando tiver uma exportação mais nova — não é preciso subir a cada
  oferta.
- **RD Station** é pensada para atualização semanal: tem upload próprio na barra lateral,
  separado das outras duas, e é opcional (o app funciona sem ela, só sem o critério de
  navegação).
- Toda vez que você sobe uma base nova pela barra lateral, ela vale a partir daquele momento
  nesta sessão do app. Para essa atualização sobreviver a um reinício do app (ex: depois de
  ficar um tempo sem uso), suba esse mesmo arquivo também na pasta `data/` do repositório no
  GitHub — é o mesmo passo de arrastar-e-soltar que fizemos a primeira vez.

## Arquivos deste pacote

- `app.py` — interface (bases fixas + botão de atualizar, upload semanal da RD Station, campos
  da oferta, botão gerar, download do Excel)
- `segmentation.py` — o motor de cálculo (4 critérios + fallback + RFV + critério de navegação)
- `requirements.txt` — dependências Python
- `data/` — onde ficam as bases (Pedidos e Lead Score); a RD Station não fica versionada aqui
  por ser grande demais para o GitHub e por ser atualizada toda semana

## Rodar no seu computador

```
pip install -r requirements.txt
streamlit run app.py
```
Abre em `http://localhost:8501`.

## Publicado em

https://clube-do-malte-publico-app-e7ux5xegmoezuhmtnys3sx.streamlit.app

Qualquer commit novo na branch `main` do repositório republica o app automaticamente em
alguns minutos.
