# Histórico de versões do ViabilizaLab

Programa gratuito de plano de negócio e análise de viabilidade para o ensino de Administração (antigo PNBOX Acadêmico).
Cada versão publicada tem uma etiqueta (*tag*) e uma página em **Releases** no GitHub. A numeração segue o padrão
`maior.menor.correção`: o primeiro número muda quando entra um conjunto grande de ferramentas, o segundo quando entram
funções novas e o terceiro quando há só correções.

---

## v8.2 (25/11/2026): mix de marketing integrado: Praça e Promoção

**Plano de marketing (4Ps): Praça e Promoção**
- **Praça:** intensidade da distribuição (intensiva, seletiva ou exclusiva) e tabela de canais de venda com a parte das vendas e a comissão de cada canal. Um botão aplica a comissão média aos custos variáveis, sem contar duas vezes a comissão de aplicativos que já existia.
- Ligação com o plano operacional para a localização.
- **Promoção:** objetivo principal da comunicação (tornar conhecido, fazer experimentar, converter ou fidelizar) com as ações típicas de cada um.
- **Volume do plano × promoção:** clientes ativos que o primeiro mês exige (receita ÷ gasto mensal por cliente) comparados aos clientes novos que o orçamento de marketing traz, com o tempo para formar a base.
- **CAC e retorno:** custo por cliente novo, meses para recuperar esse custo em margem de contribuição e verba de promoção sobre a receita.
- O diagnóstico do mix ganha alertas de canais que não somam 100%, comissão não aplicada, dependência de um canal de terceiros, posição premium em canais de comparação de preço e volume sem promoção que o sustente.

---

## v8.1 (25/11/2026): mix de marketing integrado: Produto e Preço

**Plano de marketing (4Ps): Produto e Preço**
- A página do 4Ps passa a ser um painel com abas por P e um diagnóstico do mix. Os textos já escritos continuam como justificativa.
- **Produto:** níveis de produto (benefício central, produto em si e produto ampliado), fase do ciclo de vida com a orientação de cada fase, concentração do portfólio (itens que fazem 80% da receita) e aviso quando o diferencial não aparece na proposta de valor nem no posicionamento.
- **Preço:** estratégia de preço (penetração, desnatação, valor percebido, custo mais margem, referência na concorrência) e posição de preço (econômico, intermediário, premium).
- **Preços dos concorrentes** por produto, com o índice de preço (seu preço ÷ média dos concorrentes).
- **Margem de contribuição de cada item** e a conta de equilíbrio: quanto o volume precisa subir para compensar um desconto e quanto pode cair depois de um aumento.
- **Condições de pagamento** (Pix, dinheiro, débito, crédito, parcelado) com taxa e prazo de cada forma. Um botão aplica a taxa média aos custos variáveis e os prazos ao capital de giro.
- **Diagnóstico do mix:** alertas de margem negativa ou estreita, preço incoerente com a estratégia ou com a posição declarada, preço fora da faixa da pesquisa Van Westendorp e dependência de um único item.

---

## v8.0 (25/11/2026): dados públicos, ABNT, financiamento e modo professor completo

**Funções básicas**
- **Trilha guiada e o que falta:** ordem recomendada de preenchimento, próximo passo e lista de pendências com atalho.
- **Glossário e leitura dos números:** 35 termos explicados e a interpretação, em linguagem simples, dos indicadores do projeto aberto.
- **Importar e exportar planilha:** produtos, custos, investimentos, materiais, equipe, entrevistas, teste de preço,
  concorrentes e riscos vindos de Excel ou CSV, com planilha-modelo.
- **Relatório acadêmico ABNT:** capa, folha de rosto, sumário com páginas (NBR 6027), seções numeradas (NBR 6024),
  tabelas com título e fonte, citações autor-data (NBR 10520) e referências (NBR 6023), margens de 3 e 2 cm e fonte 12.
- **Financiamento:** sistemas SAC e Price, carência, custo efetivo total e comparação entre propostas de crédito.
- **Histórico de versões do projeto:** comparação dos indicadores entre as versões salvas.

**Funções completas**
- **Dados públicos:** população, domicílios, PIB e empresas do município pela API do IBGE, e gastos das famílias por
  categoria da POF 2017-2018, com botões para usar os valores no tamanho do mercado.
- **Referências do setor:** comparação com a média nacional da atividade (IBGE, Pesquisa Anual de Comércio e Pesquisa
  Anual de Serviços 2024) e taxas de sobrevivência das empresas (Sebrae).
- **SWOT cruzada:** estratégias ofensivas, de confronto, de reforço e defensivas.
- **Lean Canvas**, montado a partir das outras ferramentas.
- **Valuation e rodada de investimento:** fluxo de caixa descontado com perpetuidade, múltiplos de EBITDA e receita,
  tabela de participação dos sócios antes e depois do aporte.
- **Impacto social e ODS:** teoria da mudança, indicadores de impacto e os 17 Objetivos de Desenvolvimento Sustentável.
- **Editais de fomento:** rascunho da proposta no formato usual do Programa Centelha, com limite de caracteres e
  orçamento por rubrica.
- **Modo professor:** comentários do professor em cada ferramenta (modo de correção), painel de prontidão e evidências
  da turma, detecção de planos parecidos e rodada de investimento da turma.

**Controle de versões**
- Este arquivo, etiquetas e páginas de *release* para todas as versões e a tela **Novidades** dentro do programa.

---

## v7.2.1 (25/11/2026): correção
- O rascunho automático do sumário executivo aparece na hora também na versão on-line.

## v7.2 (25/11/2026): novo visual das matrizes e modelo de serviços
- Matriz GE/McKinsey 3x3 colorida, BCG com quadrantes descritos e bolhas, BSC em ciclo e Canvas com blocos coloridos, na
  tela e no PDF.
- Quarto exemplo: **TecVale Soluções**, empresa de serviços de TI (link `?versao=servico`).
- Correção do texto "None" que aparecia ao gerar o PDF.

## v7.1 (25/11/2026): validação, marketing estratégico, incerteza e defesa
- Placar de evidências, entrevistas com clientes, teste de preço de Van Westendorp e intenção de compra ajustada.
- Canvas da proposta de valor, curva de valor e matriz ERRC, posicionamento e mapa perceptual, orçamento por canal e funil AARRR.
- Simulação de Monte Carlo, perda aceitável dos sócios e curva de adoção de Bass.
- Régua de prontidão, banca simulada e checklist de formalização.
- Teste de viabilidade real no prompt de IA.

## v7.0 (25/11/2026): ViabilizaLab
- Novo nome (antes PNBOX Acadêmico); arquivos `.pnbox` continuam compatíveis.
- Horizonte de projeção de 1 a 10 anos, plano de contratações mês a mês e CAC pela verba de aquisição.
- Missão, visão, valores e objetivos SMART; Business Model Canvas, cinco forças de Porter, matrizes GE e BCG e Balanced Scorecard.
- Simulação da reforma tributária (Simples puro x híbrido).

## v6.0 (24/11/2026): serviços e assinaturas
- Custo da hora de serviço e ocupação da equipe; receita recorrente com safras, churn, MRR, CAC, LTV e receita antecipada.
- Calculadora do Simples Nacional (anexos I a V e fator R).
- Exemplo **Conecta Vale** (aplicativo por assinatura) e layout adaptado para celular.

## v5.0.1 (24/11/2026): ajustes
- Cartões clicáveis com descrição e selo de andamento; organograma legível; instalador sem a janela do Python.

## v5.0 (24/11/2026): plano completo
- Interface em cartões, banco de dados e motor financeiro ampliados.
- Análise com IA, sumário executivo, modo apresentação, relatório em PDF, pitch e Word, modo professor com rubrica,
  configurações, pontos de restauração e o exemplo **Café Mucuri**.

## v4.1 (24/11/2026): primeira versão publicada no GitHub
- Segmentação, proposta de valor, personas, jornada do cliente, concorrência, matriz SWOT, canais e funil, investimentos,
  materiais e fichas técnicas, custos, DRE de 3 anos, indicadores (VPL, TIR, payback, pontos de equilíbrio), cenários,
  gráficos, checagens de consistência e relatório em PDF e Excel.
- Versão on-line (GitHub Pages) e instaladores para Windows (modelo FitBox e em branco).

## v1 a v4.0: desenvolvimento local
- Versões construídas no computador do autor, antes da criação do repositório. Não há registro detalhado de cada etapa;
  o conjunto de funções que elas reuniram é o descrito na v4.1.
