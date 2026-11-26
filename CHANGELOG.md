# Histórico de versões do ViabilizaLab

Programa gratuito de plano de negócio e análise de viabilidade para o ensino de Administração (antigo PNBOX Acadêmico).
Cada versão publicada tem uma etiqueta (*tag*) e uma página em **Releases** no GitHub. A numeração segue o padrão
`maior.menor.correção`: o primeiro número muda quando entra um conjunto grande de ferramentas, o segundo quando entram
funções novas e o terceiro quando há só correções.

---

## v9.0.3 (26/11/2026): instalador com a versão correta

**Instalador mostra a versão certa**
- O instalador para Windows exibia "versão 7.0" na tela de instalação e na lista de programas do Windows, porque esse número estava fixo no arquivo do instalador. O programa instalado já era o atual. Agora a versão é lida do próprio programa a cada geração, e a página dos instaladores também informa a versão.
- Endereços do README atualizados para o novo nome do repositório (Viabiliza_Lab).

---

## v9.0.2 (26/11/2026): correção do botão Voltar do navegador

**Correção: botão Voltar do navegador**
- O botão Voltar do navegador não levava ao painel do plano. O comando que escutava o Voltar ficava preso a um quadro interno da página que o Streamlit remove a cada atualização, e parava de funcionar logo depois de criado. Agora ele é registrado na própria página e continua ativo durante todo o uso.

---

## v9.0.1 (25/11/2026): ajustes de navegação

**Ajustes de navegação**
- O botão de voltar ao painel do plano passa a ficar no **fim de cada página**, abaixo dos botões de ferramenta anterior e seguinte.
- Ao abrir uma ferramenta pelo cartão, a página **começa no topo**. Antes, o comando de rolar a tela era ignorado quando se repetia, e a página podia abrir no fim.
- O **botão Voltar do navegador** (e o gesto de voltar no celular) leva de volta ao painel do plano. Passar de uma ferramenta para a seguinte não acumula etapas no histórico: um único Voltar retorna ao painel.

---

## v9.0 (25/11/2026): mapa de concorrentes, árvore de decisão e acessibilidade

**Mapa de concorrentes, árvore de decisão e acessibilidade**
- **Mapa de concorrentes** (OpenStreetMap, gratuito): busca no município do projeto os estabelecimentos do mesmo tipo (14 tipos, como cafeteria, padaria, salão, oficina e loja de roupas) e os lugares que geram movimento (escolas, faculdades, pontos de ônibus, hospitais e supermercados).
- O mapa mostra o ponto do negócio e o raio de atendimento, a distância de cada concorrente e a densidade por mil habitantes, usando a população do IBGE. Os nomes encontrados podem ser levados para o cadastro de concorrentes. Os pontos ficam salvos no projeto.
- **Árvore de decisão:** abrir ou não abrir em três cenários de demanda com probabilidades escolhidas pelo grupo. Inclui a **opção de abandono** (encerrar depois do Ano 1, vender os equipamentos e recuperar o capital de giro) e o **valor da informação perfeita**, que indica quanto vale, no máximo, uma pesquisa ou um piloto antes de decidir.
- **Acessibilidade** (menu lateral): letra maior, alto contraste com foco visível no teclado e botão para ouvir a leitura dos resultados com a voz do próprio navegador.
- Com esta versão, o programa chega a 73 ferramentas. Os projetos antigos continuam abrindo com os mesmos números da v8.0.

---

## v8.9 (25/11/2026): pitch em PowerPoint e pontos de partida por setor

**Pitch em PowerPoint e pontos de partida por setor**
- **Pitch em PowerPoint (.pptx) editável**, que abre no PowerPoint, no LibreOffice e no Google Apresentações. Tem gráficos nativos (receita e resultado por ano; composição do investimento) e **notas do apresentador** com o roteiro de fala e o tempo sugerido de cada slide.
- **Três roteiros:** banca acadêmica (com sensibilidade e riscos), investidor (com validação, uso do dinheiro e equipe) e edital de fomento (com impacto). Os slides são montados com os dados do plano.
- **Slide de reserva** com as perguntas prováveis da banca, geradas a partir das fragilidades do próprio plano.
- **Ponto de partida por setor**, uma nova opção ao criar um projeto. Há seis setores: padaria e lanchonete, salão de beleza, oficina mecânica, loja de roupas, produção rural e agricultura familiar, e aplicativo ou serviço digital.
- Cada setor traz itens de venda, investimentos, custos fixos, licenças, riscos típicos, referência do setor e perfil de sazonalidade, **sem valores**: preços e custos são preenchidos pelo grupo com orçamentos reais.
- A biblioteca gratuita python-pptx entra na lista de instalação do programa e da versão online.

---

## v8.8 (25/11/2026): Selic, inflação e TMA em degraus

**Selic, inflação e TMA**
- Nova ferramenta **Selic, inflação e TMA**: busca a meta da taxa Selic (série 432) e o IPCA acumulado em 12 meses (série 13522) no Sistema Gerenciador de Séries Temporais do Banco Central, com o juro real calculado pela equação de Fisher. Sem internet, usa os últimos valores guardados no programa (Selic de 13,75% em 04/11/2026; IPCA de 4,63% até outubro de 2026).
- **TMA em degraus:** taxa livre de risco (Selic), prêmio pelo risco do negócio (com quatro níveis de referência) e prêmio por ser pequeno negócio e de difícil revenda, com gráfico e botões para usar a TMA sugerida e o IPCA como inflação de custos no plano.
- Leitura dos resultados: TMA abaixo da Selic, risco possivelmente subestimado, inflação de custos distante do IPCA e empréstimos com taxa abaixo da Selic (só realistas em linhas subsidiadas).
- Os valores consultados ficam salvos no projeto e entram nas referências do relatório ABNT.

---

## v8.7 (25/11/2026): sensibilidade do VPL (tornado e ponto de ruptura)

**Sensibilidade do VPL (tornado)**
- Nova ferramenta **Sensibilidade do VPL**: gráfico de tornado com o VPL no pior e no melhor valor de cada variável (preço, volume, custo de materiais, salários, demais custos fixos e crescimento), usando as faixas definidas no Monte Carlo. As barras mais compridas mostram as premissas que mais merecem pesquisa.
- **Ponto de ruptura:** quanto cada variável pode mudar, sozinha, até o VPL chegar a zero.
- **Valide isto primeiro:** cruza as três variáveis que mais mexem no VPL com o placar de evidências. Premissa decisiva que ainda é só estimativa recebe alerta vermelho.
- **Mapa preço × volume:** VPL para combinações de variação de preço e de volume, com as combinações em que o projeto não se paga em vermelho.
- Leitura dos resultados com a variável mais sensível e a margem de segurança mais estreita.
- O cálculo usa a mesma lógica do Monte Carlo: sem variação, o VPL é exatamente o do plano.

---

## v8.6 (25/11/2026): sazonalidade das vendas

**Sazonalidade das vendas**
- Nova ferramenta **Sazonalidade das vendas**: índice de cada mês do calendário (100 = mês médio) e o mês em que o negócio abre.
- **Perfis de partida ilustrativos** para ajustar: varejo de presentes, alimentação, sorvetes e bebidas geladas, cursos, serviços a empresas e turismo.
- O motor financeiro distribui a receita e os custos diretos do Ano 1 pelos meses conforme os índices. O total do ano é preservado: a DRE anual, o VPL e a TIR não mudam; mudam a projeção mensal e o **fluxo de caixa**, que passa a mostrar os meses apertados.
- Gráfico de receita e resultado mês a mês e leitura dos resultados com o mês mais forte, o mais fraco, os meses com prejuízo operacional e o ponto mais baixo do caixa.
- Nova seção no relatório PDF com fator, receita, resultado e saldo de caixa de cada mês.
- Projetos sem sazonalidade informada continuam com os mesmos números.

---

## v8.5 (25/11/2026): leitura automática dos resultados e modo exercício

**Leitura automática dos resultados**
- Ao final das tabelas do programa aparece a caixa **📌 Leitura dos resultados**, com duas a quatro frases que dizem o que os números mostram.
- Há regras para 37 tabelas: produtos, custos, equipe, investimentos, empréstimos, sócios, prazos, segmentos, SWOT, Porter, riscos, pesquisas de preço e entrevistas, canais, funil, orçamento de marketing, fornecedores, cronograma, BCG, objetivos, mix de marketing e outras. Também há leitura para a DRE e para o fluxo de caixa.
- As regras são fixas e não usam IA: o resultado é o mesmo em qualquer computador e o professor sabe de onde veio cada frase.
- **Modo exercício** (menu lateral): o aluno escreve a própria leitura de cada tabela e só depois vê a do programa. As leituras ficam salvas no projeto e saem no relatório PDF, numa seção própria, para o professor avaliar.
- A leitura automática pode ser desligada no menu lateral.

---

## v8.4 (25/11/2026): plano de marketing estruturado

**Plano de marketing estruturado**
- Nova ferramenta **Plano de marketing**, que amarra as ferramentas de marketing na sequência da literatura: análise da situação, objetivos, estratégia, táticas, ações e controle (Smith, 2011; McDonald, 2007; Las Casas, 2011).
- **Situação de cada etapa** (completa, em andamento ou vazia), com atalhos para a ferramenta que alimenta cada uma.
- **Situação:** síntese do diagnóstico escrita pelo grupo, somada ao tamanho do mercado e à comparação de preços com os concorrentes.
- **Objetivos:** os objetivos SMART da perspectiva Clientes.
- **Estratégia:** declaração de posicionamento, direção de crescimento pela matriz de Ansoff e estratégia competitiva de Porter, com alertas quando contradizem a estratégia de preço.
- **Táticas:** resumo do mix montado a partir das escolhas estruturadas, com a coerência do mix.
- **Ações:** o calendário de marketing.
- **Controle:** indicadores com meta, frequência, fonte do dado e responsável, com sugestões prontas.
- O relatório PDF e o **relatório ABNT** ganham o plano de marketing completo, com citações e referências (Smith, McDonald, Las Casas, Kotler e Keller, Booms e Bitner, Ansoff).
- A antiga página "Plano de marketing (4Ps)" passa a se chamar **Mix de marketing (4Ps e 7Ps)**.
- Os modelos Café Mucuri e TecVale trazem o plano de marketing preenchido como exemplo.

---

## v8.3 (25/11/2026): 7Ps de serviços, coerência do mix e 4Cs

**Mix de marketing: 7Ps de serviços, coerência do mix e 4Cs**
- **Pessoas, processos e evidência física** (Booms; Bitner, 1981), nova aba do mix: quem atende e como o atendimento é padronizado, o que o cliente vê e toca, e o **mapa do serviço** (Shostack, 1984), com etapas visíveis ao cliente, bastidores e apoio, tempo de cada etapa e pontos de falha.
- **Coerência entre os Ps:** tabela de combinações que se reforçam ou se contradizem (preço premium com produto ampliado e distribuição seletiva, penetração com comunicação de alcance, desnatação no lançamento, objetivo da comunicação conforme a fase do ciclo de vida, canais cadastrados conforme a intensidade de distribuição).
- **Autoavaliação pelos 4Cs** (Lauterborn, 1990): cliente, custo para o cliente, conveniência e comunicação, com nota, evidência e indicação do ponto mais fraco.
- Os modelos TecVale e Café Mucuri trazem exemplos do mapa do serviço.

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
