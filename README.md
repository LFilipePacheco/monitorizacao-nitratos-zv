# Monitorização periódica dos teores de nitratos
### Zona Vulnerável Esposende – Vila do Conde (Diretiva Nitratos, 91/676/CEE)

> Sistema integrado de recolha, gestão e análise de dados de qualidade da água
> subterrânea em poços agrícolas — do registo de campo ao dashboard interactivo,
> com pipeline de dados automatizado e governação documentada.

---

## Porquê monitorizar

Para compreender a dinâmica temporal da contaminação por nitratos e avaliar se
as medidas em vigor estão efetivamente a produzir resultados, é indispensável
dispor de séries de dados contínuas e espacialmente representativas. É esse o
papel da monitorização periódica, que constitui o elemento central do sistema
de acompanhamento da Zona Vulnerável.

A Zona Vulnerável Esposende – Vila do Conde (ZV) é uma das zonas designadas em
Portugal ao abrigo da Diretiva Nitratos, abrangendo territórios agrícolas de
elevada intensidade produtiva nos municípios de Barcelos, Esposende, Póvoa de
Varzim e Vila do Conde. O Programa de Ação em vigor impõe restrições às
práticas de fertilização; a monitorização é o instrumento que permite verificar
o seu efeito real sobre a água subterrânea, poço a poço, campanha a campanha.

A rede de monitorização assenta em poços agrícolas distribuídos pela área da ZV,
com registo periódico das concentrações de NO₃ (mg/L) e
avaliação face ao Valor Máximo Admissível (VMA = 50 mg/L) e a um limiar de
alerta interno (25 mg/L). Algumas séries remontam a 2008 — um património de
dados de quase duas décadas.

## O problema

O processo herdado apresentava fragilidades típicas de sistemas crescidos
organicamente ao longo de anos:

- **Dados presos num Excel partilhado** — o registo de campo funcionava bem,
  mas a análise exigia extração e tratamento manual a cada campanha;
- **Análise pontual, não contínua** — gráficos estáticos produzidos
  manualmente para relatórios; tendências emergentes podiam passar
  despercebidas entre relatórios;
- **Sem partilha operacional** — as equipas que registam os dados não tinham
  acesso fácil aos resultados analíticos do seu próprio trabalho;
- **Qualidade de dados sem verificação sistemática** — valores anómalos
  (coordenadas GPS registadas em células de medições, valores impossíveis)
  permaneciam na série e contaminavam estatísticas;
- **Conhecimento não documentado** — regras de gestão (suspensão de poços,
  correções, poços novos) dependiam de memória individual.

## A solução

Um pipeline de dados em camadas, com o princípio de **fonte única**: o Excel
partilhado permanece o ponto de registo (sem alterar o fluxo de trabalho das
equipas de campo); tudo o resto deriva dele de forma automática e idempotente.

```
Equipas de campo registam no Excel partilhado (OneDrive)
        │
        ├──► PostgreSQL (fonte de verdade analítica)
        │        └──► Relatório Word automatizado (estatísticas, gráficos, conformidade)
        │
        ├──► ArcGIS (feature class de medições — análise espacial)
        │
        └──► GitHub (publicação mensal agendada, com log de auditoria)
                 └──► Dashboard Streamlit na cloud
                          └──► equipas e decisores, via browser, autenticado
```

### Componentes principais

**1. Ingestão validada (Excel → PostgreSQL)** — leitura estruturada do formato
de campo (uma folha por município, códigos de poço em linha de cabeçalho),
com validação na entrada e carregamento idempotente: correções no Excel
propagam-se; nada é apagado automaticamente (proteção contra perdas
acidentais).

**2. Sincronização espacial (Excel → GDB)** — as medições alimentam uma
feature class no ArcGIS, permitindo análise espacial (interpolação IDW e
kriging da distribuição de nitratos) e cartografia de apoio à decisão.

**3. Relatório automatizado** — documento Word institucional gerado por
script: estatísticas descritivas por poço, séries temporais, tendências
lineares com significância, heatmaps anuais, conformidade com o VMA e
conclusões calculadas — o que antes exigia dias passa a minutos.

**4. Dashboard interactivo (Streamlit)** — publicado na cloud com
autenticação, atualizado automaticamente todos os meses por tarefa agendada
(commit/push para GitHub com log de auditoria). Sete vistas analíticas:
séries temporais, distribuições, heatmap anual, tendências com regressão
(declive anualizado, R², p-value), conformidade, qualidade dos dados e
exportação. Detalhes de desenho com relevância metodológica:

- *Separação estrita entre medido e imputado* — todas as estatísticas usam
  exclusivamente observações reais; valores imputados (lacunas interiores,
  mínimo de 3 observações) servem apenas a continuidade visual e são
  assinalados graficamente;
- *Metadados lidos da convenção de campo* — poços suspensos são detetados
  pela cor de preenchimento no cabeçalho do Excel, eliminando listas
  paralelas de configuração;
- *Filtro de sanidade na leitura* (0–1000 mg/L) — rejeita erros
  (coordenadas, datas serializadas) sem excluir medições extremas genuínas,
  com aviso do número de valores rejeitados.

**5. Governação documentada** — guia formal de organização e gestão:
arquitetura, exceções justificadas à regra da fonte única, procedimentos
(suspender poço, poço novo, correções), rotina mensal, cópias de segurança
(OneDrive + GitHub offsite + dumps PostgreSQL agendados).

## Benefícios

- **Do registo à análise sem intervenção manual** — a atualização mensal é
  uma tarefa agendada; o esforço analítico recorrente aproximou-se de zero;
- **Deteção atempada** — tendências crescentes e ultrapassagens ao VMA
  tornam-se visíveis no próprio mês, não no relatório seguinte; poços
  problemáticos são identificáveis de imediato para investigação dirigida;
- **Qualidade de dados auditável** — a verificação sistemática na leitura
  já identificou e isolou anomalias reais (coordenadas GPS em células de
  medições; recuperação de medições extremas genuínas de 501–600 mg/L que
  um filtro ingénuo excluiria);
- **Democratização dos resultados** — a equipa que produzem os dados
  passaram a ver o resultado analítico do seu trabalho num browser, sem
  instalar nada;
- **Resiliência institucional** — processos, regras e arquitetura
  documentados; histórico protegido por três camadas de backup; o sistema
  não depende da memória de uma pessoa.

## Stack

Python (pandas · openpyxl · psycopg2 · SciPy · Plotly) · PostgreSQL ·
ArcGIS Pro (arcpy) · Streamlit · GitHub (+ Actions de deploy da Streamlit
Cloud) · Agendador de Tarefas do Windows · python-docx

## Nota sobre os dados

Os dados de monitorização são propriedade institucional da CCDR-Norte, I.P. e
não são publicados neste repositório. A demonstração pública do dashboard
utiliza **dados sintéticos** com estrutura e comportamento estatístico
plausíveis (sazonalidade, tendências, lacunas), gerados por script incluído
no repositório.

---

**Luís Filipe Pacheco** — Senior Agricultural Engineer & Data Scientist,
CCDR-Norte, I.P. · [Perfil GitHub](https://github.com/LFilipePacheco) ·
[LinkedIn](https://www.linkedin.com/in/lu%C3%ADs-filipe-pacheco-471495b/) ·
[ORCID](https://orcid.org/0009-0001-7676-6542)
