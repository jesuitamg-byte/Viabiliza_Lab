"""
================================================================================
VIABILIZALAB - ESTUDO DE CASO FITBOX (pnbox_modelo.py)
================================================================================
Dados de exemplo inspirados no relatório PNBOX/Sebrae do projeto FitBox
(Teófilo Otoni - MG). Servem de referência de preenchimento para os alunos.
================================================================================
"""


def popular_modelo_fitbox(conn) -> None:
    ins = conn.executemany

    conn.execute("""
        UPDATE projeto SET nome_negocio='FitBox', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Pequena Empresa (EPP)',
            setor_atuacao='Comércio; Serviço; Alimentos e bebidas; Bares e restaurantes; Varejo; Moda',
            cidade_uf='Teófilo Otoni - MG', contatos='@fitbox.br',
            aliquota_imposto_pct=10.0, crescimento_mensal_ano1_pct=1.0, crescimento_anual_pct=8.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=12.0, pme_dias=15, reserva_caixa_extra=150000,
            recuperar_giro=1, cen_otim_receita_pct=10, cen_otim_custo_pct=5,
            cen_pess_receita_pct=15, cen_pess_custo_pct=5
        WHERE id=1""")

    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Alessandra Ribeiro Gonçalves", "Sócia-administradora", "",
         "Graduanda em Administração; planejamento operacional e financeiro, controle de custos e fornecedores.", 100000.0),
        ("Paulo Henrique Batista", "Sócio-administrador", "",
         "Gestão estratégica, marketing, posicionamento da marca e relacionamento com clientes.", 80000.0),
        ("Camila Gomes de Souza", "Sócia comercial", "",
         "Desenvolvimento comercial, parcerias com academias e varejo fitness.", 60000.0),
        ("Cauan Ribeiro de Souza", "Sócio operacional", "",
         "Acompanhamento de indicadores, expansão e consolidação da empresa no mercado.", 43975.0),
    ])

    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Alimentação saudável / Conveniência",
         "Trabalhadores, estudantes e consumidores que buscam refeições equilibradas, saborosas e práticas.", 10, 9, 10, 6, 9),
        ("Praticantes de atividades físicas",
         "Pessoas que praticam academia ou esportes e buscam alimentação equilibrada, suplementos e roupas fitness.", 8, 9, 9, 7, 10),
        ("Iniciantes no estilo de vida saudável",
         "Pessoas que desejam melhorar hábitos sem dietas restritivas.", 9, 7, 9, 7, 8),
        ("Consumidores de lifestyle fitness",
         "Público engajado em bem-estar, moda esportiva e experiências.", 7, 8, 8, 6, 9),
    ])

    conn.execute("""UPDATE proposta_valor SET
        segmento_alvo='trabalhadores, estudantes e praticantes de atividades físicas',
        beneficios_desejados='conforto, conveniência, economia de tempo, transparência nutricional e alimentação equilibrada',
        meios_entrega='refeições saborosas e personalizáveis, planos semanais, delivery ágil e loja integrada de suplementos e moda fitness',
        diferencial_competitivo='estabelecimentos focados apenas em marmitas tradicionais ou somente em artigos esportivos',
        entrega_produtos_servicos='praticidade, variedade, sabor e controle nutricional em uma experiência acessível e acolhedora.'
        WHERE id=1""")

    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, caracteristicas, fluencia_digital, redes_sociais, quem_influencia,
           rotina_manha, rotina_tarde, rotina_noite, ganhos, desafios, duvidas_objecoes, aborrecimentos,
           experiencia_ideal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Mariana", 24, "Analista Administrativa", "Solteira", 3500.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Trabalha em horário comercial e busca conciliar carreira, vida pessoal e alimentação equilibrada. "
         "Valoriza praticidade, sabor, variedade e informações nutricionais confiáveis.",
         "Quero cuidar da minha saúde, mas minha rotina nem sempre permite cozinhar.",
         "Pratica exercícios algumas vezes por semana; pesquisa preços, cardápios e avaliações pelo celular.",
         "#Organizada #Pragmática #Responsável #Sociável #Consciente #Exigente",
         "Computador: alta; Smartphone: alta; Aplicativos: alta; Compras on-line: média",
         "Instagram, WhatsApp, TikTok",
         "Amigos, colegas de trabalho, especialistas, avaliações de outros clientes",
         "Acorda às 6h30, toma café rápido, verifica mensagens e inicia o expediente às 8h.",
         "Reuniões e tarefas administrativas; no almoço prefere refeições equilibradas e sem longa espera.",
         "Vai à academia; nem sempre tem disposição para cozinhar e considera pedir refeições prontas.",
         "Economizar tempo; manter alimentação saudável; conhecer calorias e nutrientes; personalizar refeições.",
         "Falta de tempo; preços elevados; poucas opções práticas; demora na entrega; falta de informação sobre ingredientes.",
         "As refeições têm boa qualidade? Os preços cabem no orçamento? O delivery chega no horário?",
         "Informações nutricionais incorretas, cobranças inesperadas, atrasos sem aviso, cardápio confuso.",
         "Acessa um cardápio intuitivo com preços e macros, personaliza o pedido e recebe a refeição fresca no horário."),
        ("Rafael", 26, "Técnico em Informática", "Solteiro", 3200.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Pratica musculação diariamente e busca refeições proteicas, porções adequadas e suplementos confiáveis.",
         "Treinar faz parte da minha rotina; preciso de proteína na medida certa.",
         "Treina após o trabalho; acompanha influenciadores fitness; compra suplementos on-line.",
         "#Disciplinado #Focado #Curioso #Exigente",
         "Computador: alta; Smartphone: alta; Aplicativos: alta; Compras on-line: alta",
         "Instagram, YouTube, WhatsApp",
         "Personal trainer, nutricionista, amigos da academia",
         "Acorda às 6h30, organiza lanches proteicos e acompanha conteúdos de treino.",
         "Suporte técnico; busca refeições com alto teor proteico e porções bem dimensionadas.",
         "Musculação; compra refeição pós-treino e pesquisa suplementos.",
         "Aumentar performance; saber gramatura de proteínas; comprar suplementos originais.",
         "Porções pequenas; falta de transparência nutricional; suplementos de origem duvidosa.",
         "Os suplementos são originais? As porções atendem à minha meta de proteína?",
         "Produtos falsificados, porções menores que o anunciado, falta de tabela nutricional.",
         "Encontra em um só lugar refeição pós-treino, suplementos originais e roupas para treinar."),
    ])

    etapas = [
        ("Antes", "Descoberta", "Busca alternativas práticas para almoçar bem no trabalho.", "Onde como bem perto do trabalho?", "Instagram, Google e indicações", "Neutro", "Conteúdos sobre rotina saudável e fotos reais dos pratos."),
        ("Antes", "Consideração", "Compara preços, porções e informações nutricionais.", "Vale a pena pelo preço?", "Cardápio digital e WhatsApp", "Neutro", "Exibir macros, ingredientes e descontos de pacotes semanais."),
        ("Antes", "Decisão", "Verifica prazo de entrega e escolhe a refeição.", "Vai chegar a tempo?", "Canal de pedidos e Google Avaliações", "Satisfeito", "Pedidos agendados para o horário exato do intervalo."),
        ("Durante", "Compra e Pagamento", "Personaliza acompanhamentos e paga via Pix/cartão.", "Foi fácil pedir.", "Cardápio digital e checkout", "Satisfeito", "Checkout rápido sem taxas ocultas."),
        ("Durante", "Entrega", "Recebe a marmita lacrada e identificada no trabalho.", "Chegou quentinha.", "Entregador e embalagem térmica", "Satisfeito", "Embalagem sustentável com lacre e etiqueta nutricional."),
        ("Durante", "Uso", "Consome a refeição e avalia sabor, temperatura e saciedade.", "Está muito bom!", "Produto e apresentação", "Entusiasmado", "Padronização rigorosa de temperos e gramaturas."),
        ("Depois", "Avaliação", "Responde pesquisa rápida e avalia no Google.", "Vou recomendar.", "WhatsApp Business e Google", "Satisfeito", "Cupom de retorno para quem avalia o pedido."),
        ("Depois", "Suporte", "Tira dúvidas sobre ingredientes ou conservação.", "Posso congelar?", "Atendimento WhatsApp", "Neutro", "Atendimento ágil e política clara de troca."),
        ("Depois", "Lealdade e advocacia", "Assina plano mensal e indica colegas de trabalho.", "Virou parte da rotina.", "Programa de fidelidade", "Entusiasmado", "Benefícios progressivos para assinantes."),
    ]
    ins("""INSERT INTO jornada_cliente (persona_nome, macro_fase, etapa, acoes_cliente, pensamentos,
           pontos_contato, estado_emocional, oportunidades) VALUES ('Mariana',?,?,?,?,?,?,?)""", etapas)

    concs = ["Concorrente A (marmitaria fitness)", "Concorrente B (restaurante natural)",
             "Concorrente C (loja de suplementos)"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {
        "Atendimento": (9, 8, 7, 8), "Canais de venda": (9, 7, 8, 6),
        "Conveniência": (9, 8, 7, 6), "Design": (8, 7, 8, 6),
        "Experiência do cliente": (9, 8, 7, 7), "Inovação": (9, 6, 7, 6),
        "Preço": (7, 8, 7, 8), "Presença digital": (9, 7, 6, 8),
        "Qualidade": (9, 8, 9, 8), "Variedade": (10, 7, 7, 5),
        "Personalização": (9, 6, 6, 3), "Transparência nutricional": (10, 5, 4, 6),
        "Alimentação equilibrada": (10, 9, 8, 4), "Sabor dos produtos": (9, 8, 8, 6),
    }
    linhas = []
    for atr, vals in notas.items():
        for emp, v in zip(["Nossa Empresa"] + concs, vals):
            linhas.append((atr, emp, v))
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)", linhas)

    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Pessoas", "Treinamento da equipe", "Força", "Alto", "Programa de capacitação em higiene, atendimento e vendas.", "Calendário trimestral de capacitação.", 0),
        ("Interno", "Pessoas", "Liderança visionária", "Força", "Alto", "Proposta estratégica definida integrando alimentação e produtos fitness.", "Metas, indicadores e reuniões periódicas entre os sócios.", 1),
        ("Interno", "Marketing", "Linhas de produtos e serviços", "Força", "Alto", "Refeições, suplementos, roupas e acessórios em um só lugar.", "Ampliar portfólio conforme a demanda e criar combos.", 1),
        ("Interno", "Marketing", "Presença digital", "Força", "Alto", "Redes sociais e canais de atendimento estruturados.", "Calendário editorial e anúncios segmentados.", 0),
        ("Interno", "Finanças", "Controles financeiros", "Força", "Alto", "Controle de receitas, despesas, fluxo de caixa e indicadores.", "Auditar CMV semanalmente e revisar margens.", 0),
        ("Interno", "Operação", "Controle de qualidade", "Força", "Alto", "Procedimentos de higiene, segurança alimentar e padronização.", "Checklists diários de temperatura e validade.", 1),
        ("Interno", "Operação", "Nível de automação", "Força", "Baixo", "Automação gradual de pedidos e pagamentos.", "Implantar sistema de pedidos integrado ao PDV.", 0),
        ("Interno", "Marketing", "Reputação e imagem", "Fraqueza", "Alto", "Empreendimento novo, sem avaliações de clientes.", "Ações de lançamento e incentivo a avaliações autênticas.", 1),
        ("Interno", "Marketing", "Participação de mercado", "Fraqueza", "Alto", "Início sem participação consolidada.", "Parcerias com academias e empresas locais.", 1),
        ("Interno", "Operação", "Instalações e equipamentos", "Fraqueza", "Alto", "Estrutura física ainda a viabilizar.", "Planejar estrutura conforme capacidade produtiva.", 1),
        ("Interno", "Operação", "Relação com fornecedores", "Fraqueza", "Alto", "Fornecedores ainda não consolidados.", "Homologar ao menos 2 fornecedores por insumo crítico.", 0),
        ("Externo", "Social", "Preocupação com saúde", "Oportunidade", "Alto", "Maior procura por alimentação equilibrada.", "Divulgar benefícios com informações nutricionais confiáveis.", 1),
        ("Externo", "Social", "Busca por qualidade de vida", "Oportunidade", "Alto", "Interesse por bem-estar e hábitos saudáveis.", "Planos de alimentação e campanhas de bem-estar.", 1),
        ("Externo", "Tecnológico", "Hábitos de consumo digitais", "Oportunidade", "Alto", "Valorização da praticidade e das compras digitais.", "Pedidos pelo WhatsApp e programa de fidelidade.", 1),
        ("Externo", "Econômico", "Inflação", "Ameaça", "Alto", "Alta de alimentos, embalagens e energia pressiona custos.", "Negociar com fornecedores e revisar preços periodicamente.", 1),
        ("Externo", "Econômico", "Renda dos consumidores", "Ameaça", "Alto", "Renda limitada reduz compras recorrentes.", "Refeições econômicas e combos promocionais.", 1),
        ("Externo", "Político-Legal", "Fiscalização e regulação", "Ameaça", "Alto", "Exigências sanitárias e risco de penalidades.", "Licenças atualizadas e consultoria técnica.", 0),
        ("Externo", "Econômico", "Concorrência de aplicativos", "Ameaça", "Baixo", "Restaurantes em apps de delivery com promoções.", "Diferenciar pela transparência nutricional.", 0),
    ])

    ins("""INSERT INTO canais (canal, atratividade, alcance, cliente, relacionamento, viabilidade)
           VALUES (?,?,?,?,?,?)""", [
        ("Marketing de Conteúdo", 9, 9, 9, 8, 10), ("Publicidade em Redes Sociais", 9, 9, 10, 8, 10),
        ("Marketing Viral", 9, 9, 9, 8, 9), ("Comunidade", 8, 8, 8, 8, 7),
        ("Plataformas Existentes", 8, 7, 7, 7, 8), ("Acordos e Parcerias", 7, 6, 6, 7, 7),
        ("Eventos", 7, 6, 5, 7, 7),
    ])
    ins("INSERT INTO funil_vendas (etapa, acao, orcamento, alcancados, convertidos) VALUES (?,?,?,?,?)", [
        ("Topo do Funil", "Anúncios Instagram/Meta Ads + parcerias com academias", 1200.0, 25000, 3200),
        ("Meio do Funil", "Cardápio digital interativo + cupom de primeira compra", 500.0, 3200, 950),
        ("Fundo do Funil", "Atendimento consultivo no WhatsApp + oferta de assinatura", 300.0, 950, 410),
    ])

    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,10)""", [
        ("Forno combinado profissional 6 GN", 1, 34000, 10), ("Fogão industrial inox 6 bocas", 2, 4000, 10),
        ("Chapa profissional para grelhados", 1, 4500, 10), ("Processador industrial de alimentos", 1, 5500, 10),
        ("Liquidificador industrial 8 litros", 2, 1800, 10), ("Extrator profissional de sucos", 1, 3600, 10),
        ("Seladora a vácuo profissional", 1, 5500, 10), ("Balança comercial 30 kg", 3, 750, 10),
        ("Refrigerador industrial inox 4 portas", 2, 8500, 10), ("Freezer vertical comercial", 2, 6500, 10),
        ("Balcão refrigerado de preparo", 1, 8500, 10), ("Expositor refrigerado de bebidas", 1, 11000, 10),
        ("Máquina de gelo profissional", 1, 7000, 10), ("Coifa e exaustão industrial", 1, 11000, 10),
        ("Lava-louças profissional", 1, 18000, 10), ("Máquina de café expresso", 1, 7500, 10),
        ("Bancada inox de manipulação", 3, 3000, 10), ("Panelas e utensílios profissionais (conjunto)", 2, 4000, 10),
        ("Mesas de atendimento para 4 lugares", 12, 900, 10), ("Cadeiras de atendimento", 48, 320, 10),
        ("Balcão de atendimento e caixa", 1, 7000, 10), ("Móveis expositores para suplementos e roupas", 5, 1700, 10),
        ("Sistema de climatização instalado", 1, 18000, 10), ("Instalações elétrica, hidráulica e de gás", 1, 35000, 10),
        ("Computadores, PDV e tablets", 1, 17800, 5), ("Impressoras, nobreak e rede", 1, 9500, 5),
        ("Motocicleta para entregas", 2, 16000, 5),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Reformas, obras e adequação sanitária", 75000), ("Legalização, alvarás e licenças sanitárias", 12000),
        ("Projeto arquitetônico e de cozinha", 15000), ("Marca e identidade visual", 8000),
        ("Softwares de gestão e cardápio digital", 9500), ("Marketing de inauguração", 25000),
        ("Treinamento inicial da equipe", 10000), ("Taxas, cauções e despesas de implantação", 10000),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Capital de giro (Pronampe)", 120000, 14.0, 36),
        ("Financiamento de equipamentos (BDMG)", 280000, 12.0, 60),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 55.0), ("Recebimento", 15, 25.0), ("Recebimento", 30, 15.0), ("Recebimento", 60, 5.0),
        ("Pagamento", 0, 15.0), ("Pagamento", 15, 20.0), ("Pagamento", 30, 40.0),
        ("Pagamento", 45, 15.0), ("Pagamento", 60, 10.0),
    ])

    ins("INSERT INTO materiais (nome, categoria, unidade, valor_unitario, qtd_estoque_inicial) VALUES (?,?,?,?,?)", [
        ("Peito de frango", "Matéria-prima", "kg", 22.00, 120), ("Carne bovina (patinho)", "Matéria-prima", "kg", 42.00, 60),
        ("Salmão", "Matéria-prima", "kg", 85.00, 15), ("Ovos", "Matéria-prima", "dúzia", 12.00, 60),
        ("Arroz integral", "Matéria-prima", "kg", 8.00, 80), ("Feijão carioca", "Matéria-prima", "kg", 9.00, 50),
        ("Batata-doce", "Matéria-prima", "kg", 6.00, 80), ("Brócolis", "Matéria-prima", "kg", 14.00, 40),
        ("Cenoura", "Matéria-prima", "kg", 6.00, 45), ("Abobrinha", "Matéria-prima", "kg", 7.00, 35),
        ("Rúcula", "Matéria-prima", "maço", 4.00, 60), ("Tomate", "Matéria-prima", "kg", 8.00, 65),
        ("Cebola", "Matéria-prima", "kg", 6.00, 50), ("Alho", "Matéria-prima", "kg", 30.00, 12),
        ("Azeite extra virgem", "Matéria-prima", "L", 55.00, 25), ("Sal", "Matéria-prima", "kg", 4.00, 15),
        ("Wrap integral", "Matéria-prima", "pacote", 12.00, 65), ("Banana", "Matéria-prima", "kg", 6.00, 75),
        ("Morango", "Matéria-prima", "kg", 22.00, 25), ("Leite desnatado", "Matéria-prima", "L", 6.00, 110),
        ("Iogurte natural", "Matéria-prima", "kg", 13.00, 65), ("Whey protein para preparo", "Matéria-prima", "kg", 95.00, 15),
        ("Polpa de fruta", "Matéria-prima", "kg", 17.00, 90), ("Marmita térmica 750 ml", "Embalagem", "unidade", 0.85, 2500),
        ("Copo 500 ml com tampa", "Embalagem", "unidade", 0.85, 1200), ("Etiqueta lacre", "Embalagem", "unidade", 0.10, 4000),
        ("Kit talher descartável", "Embalagem", "unidade", 0.45, 2000),
        ("Detergente e sanitizante", "Limpeza e higiene", "galão", 36.00, 20), ("Luva e touca descartáveis", "Limpeza e higiene", "caixa", 30.00, 30),
    ])

    F, S, R = "Fabricação Própria", "Serviço", "Revenda"
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal,
           qtd_estoque_inicial) VALUES (?,?,?,?,?,?)""", [
        (S, "Assinatura semanal - 5 marmitas tradicionais", 129.90, 42.20, 78, 0),
        (S, "Assinatura mensal - 20 marmitas tradicionais", 499.90, 168.80, 59, 0),
        (S, "Personalização de pedido (adicional)", 5.90, 1.70, 244, 0),
        (S, "Coffee break saudável para eventos (por pessoa)", 34.90, 16.98, 116, 0),
        (S, "Delivery até 3 km", 6.90, 4.20, 438, 0),
        (S, "Delivery de 3 a 10 km", 9.90, 6.50, 1168, 0),
        (F, "Marmita fitness tradicional", 29.90, 0, 876, 0),
        (F, "Marmita fitness premium", 39.90, 12.80, 291, 0),
        (F, "Marmita fitness econômica", 24.90, 6.90, 487, 0),
        (F, "Marmita low carb", 34.90, 11.20, 340, 0),
        (F, "Marmita hiperproteica", 39.90, 0, 291, 0),
        (F, "Marmita vegetariana", 29.90, 8.90, 98, 0),
        (F, "Prato executivo saudável", 36.90, 11.90, 389, 0),
        (F, "Bowl de frango e vegetais", 34.90, 10.80, 175, 0),
        (F, "Bowl de salmão", 54.90, 0, 59, 0),
        (F, "Salada completa com frango", 32.90, 9.80, 155, 0),
        (F, "Wrap integral de frango", 24.90, 0, 210, 0),
        (F, "Tapioca de frango", 18.90, 5.20, 224, 0),
        (F, "Açaí proteico 300 ml", 26.90, 9.20, 168, 0),
        (F, "Suco natural 500 ml", 15.90, 0, 560, 0),
        (F, "Smoothie de morango", 19.90, 0, 252, 0),
        (F, "Shake proteico com whey", 24.90, 0, 280, 0),
        (F, "Café expresso", 6.90, 1.40, 700, 0),
        (R, "Whey protein concentrado 900 g", 149.90, 90.00, 42, 40),
        (R, "Whey protein isolado 900 g", 229.90, 145.00, 17, 15),
        (R, "Creatina monohidratada 300 g", 79.90, 45.00, 49, 30),
        (R, "Pré-treino 300 g", 119.90, 70.00, 17, 15),
        (R, "Barra de proteína", 12.90, 6.00, 308, 200),
        (R, "Pasta de amendoim 500 g", 24.90, 14.00, 63, 40),
        (R, "Mix de castanhas 200 g", 29.90, 17.00, 56, 40),
        (R, "Água mineral 500 ml", 4.90, 1.30, 560, 300),
        (R, "Legging fitness feminina", 119.90, 60.00, 25, 20),
        (R, "Camiseta dry fit masculina", 79.90, 38.00, 28, 25),
        (R, "Coqueteleira 600 ml", 34.90, 15.00, 35, 30),
    ])

    ficha = {
        "Marmita fitness tradicional": [("Peito de frango", 0.15), ("Arroz integral", 0.10), ("Feijão carioca", 0.08),
                                        ("Brócolis", 0.06), ("Cenoura", 0.05), ("Cebola", 0.02), ("Alho", 0.005),
                                        ("Azeite extra virgem", 0.01), ("Sal", 0.003), ("Marmita térmica 750 ml", 1),
                                        ("Etiqueta lacre", 1), ("Kit talher descartável", 1)],
        "Marmita hiperproteica": [("Peito de frango", 0.25), ("Batata-doce", 0.15), ("Brócolis", 0.08),
                                  ("Ovos", 0.17), ("Azeite extra virgem", 0.01), ("Sal", 0.003),
                                  ("Marmita térmica 750 ml", 1), ("Etiqueta lacre", 1), ("Kit talher descartável", 1)],
        "Bowl de salmão": [("Salmão", 0.15), ("Arroz integral", 0.10), ("Abobrinha", 0.05), ("Rúcula", 0.1),
                           ("Azeite extra virgem", 0.01), ("Marmita térmica 750 ml", 1), ("Etiqueta lacre", 1),
                           ("Kit talher descartável", 1)],
        "Wrap integral de frango": [("Wrap integral", 0.1), ("Peito de frango", 0.10), ("Rúcula", 0.1),
                                    ("Tomate", 0.05), ("Iogurte natural", 0.03), ("Etiqueta lacre", 1)],
        "Suco natural 500 ml": [("Polpa de fruta", 0.15), ("Copo 500 ml com tampa", 1)],
        "Smoothie de morango": [("Morango", 0.15), ("Iogurte natural", 0.15), ("Banana", 0.10),
                                ("Copo 500 ml com tampa", 1)],
        "Shake proteico com whey": [("Whey protein para preparo", 0.03), ("Leite desnatado", 0.30),
                                    ("Banana", 0.10), ("Copo 500 ml com tampa", 1)],
    }
    ins("INSERT INTO ficha_tecnica (produto, material, quantidade) VALUES (?,?,?)",
        [(p, m, q) for p, itens in ficha.items() for m, q in itens])

    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [
        ("Comissões de vendas", 0.5), ("Taxas de cartões", 1.5), ("Propaganda variável", 0.8),
        ("Previsão de inadimplência", 0.3), ("Comissões de aplicativos de delivery", 0.5),
        ("Descontos e cupons promocionais", 0.5), ("Frete subsidiado em entregas", 0.5),
    ])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Gerente operacional", A, 1, 3800, 20, 200, "Coordena equipe, compras, escalas e indicadores."),
        ("Chef de cozinha / líder", P, 1, 3200, 20, 200, "Cardápio, fichas técnicas e padrão de qualidade."),
        ("Cozinheiros", P, 2, 2400, 20, 200, "Preparo das refeições conforme fichas técnicas."),
        ("Auxiliares de cozinha", P, 4, 1850, 20, 200, "Pré-preparo, porcionamento, embalagem e limpeza."),
        ("Atendentes / barista / caixa", C, 6, 1870, 20, 200, "Atendimento, bebidas, caixa e pedidos on-line."),
        ("Vendedores / estoquista / limpeza / entregadores", C, 8, 1880, 20, 200, "Loja, estoque, limpeza e entregas."),
    ])
    CIF, DA, DC, DF = ("Custo indireto de produção", "Despesa administrativa", "Despesa comercial",
                       "Despesa financeira")
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Aluguel da área de produção (cozinha)", CIF, 4200), ("Aluguel da loja e escritório", DA, 2800),
        ("Pró-labore", DA, 6000), ("Energia elétrica da cozinha", CIF, 3600), ("Energia elétrica da loja", DA, 900),
        ("Água", CIF, 2000), ("Honorários contábeis", DA, 1800), ("Manutenção de equipamentos", CIF, 1300),
        ("Seguros", DA, 1200), ("Serviços de terceiros", DA, 750), ("Marketing e propaganda (fixo)", DC, 600),
        ("Materiais de limpeza", CIF, 600), ("Telefone e internet", DA, 500), ("Material de escritório", DA, 450),
        ("Tarifas bancárias", DF, 250), ("Mensalidade de máquinas de cartão", DC, 150),
    ])

    conn.execute("""UPDATE projeto SET
        autores='Alessandra Ribeiro Gonçalves; Paulo Henrique Batista; Camila Gomes de Souza; Cauan Ribeiro de Souza',
        turma='Administração - Empreendedorismo (exemplo)', professor='', data_entrega='',
        tam_desc='Moradores de Teófilo Otoni e região que consomem refeições fora de casa (estimativa ilustrativa)',
        tam_clientes=60000,
        sam_desc='Trabalhadores, estudantes e praticantes de atividade física no raio de entrega de 10 km',
        sam_clientes=18000, som_pct=12, gasto_anual_cliente=1600,
        mercado_fontes='Estimativas ilustrativas para fins didáticos; substitua por dados do IBGE, Sebrae e pesquisa de campo.',
        mkt_produto='Refeições saudáveis com informação nutricional, planos semanais/mensais, suplementos e moda fitness em um só lugar.',
        mkt_preco='Preço intermediário: marmitas de R$ 24,90 a R$ 54,90; descontos progressivos em assinaturas.',
        mkt_praca='Loja física com consumo no local, retirada, delivery próprio até 10 km e pedidos via WhatsApp e aplicativo.',
        mkt_promocao='Instagram e anúncios segmentados, parcerias com academias, programa de fidelidade e degustações.',
        loc_endereco='Região central de Teófilo Otoni - MG (próximo a academias e centros comerciais)',
        loc_justificativa='Alto fluxo de trabalhadores e estudantes, proximidade de academias e facilidade para entregas.',
        loc_area_m2=180, loc_fluxo='Fluxo intenso no almoço (11h-14h) e no fim da tarde (17h-20h).',
        op_layout='Cozinha industrial com áreas separadas (pré-preparo, cocção, porcionamento, expedição), salão com 48 lugares e área de loja.',
        op_processos='Compras semanais -> recebimento e conferência -> pré-preparo -> cocção -> porcionamento e etiquetagem -> venda/entrega -> pós-venda.',
        op_capacidade_txt='Capacidade estimada pela equipe de cozinha e pelos equipamentos de cocção e refrigeração.',
        op_horario='Segunda a sábado, das 7h às 21h.',
        rateio_criterio='Custo direto', margem_desejada_pct=15
        WHERE id=1""")

    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Mariana", "Conforto e conveniência; Economia de tempo; Alimentação equilibrada; Transparência nutricional; Rapidez",
         "manter uma alimentação equilibrada mesmo com a rotina corrida",
         "refeições prontas com informações nutricionais, planos semanais e entrega no horário do almoço",
         "restaurantes por quilo e marmitarias tradicionais sem informação nutricional",
         "praticidade, sabor e controle nutricional por um preço acessível"),
        ("Rafael", "Desempenho; Confiança; Integração de produtos; Personalização",
         "atingir sua meta de proteína e comprar suplementos confiáveis",
         "marmitas hiperproteicas com gramatura informada e loja de suplementos originais",
         "lojas que vendem só suplementos ou só refeições",
         "refeição pós-treino, suplementos e roupas de treino em um só lugar"),
    ])
    ins("INSERT INTO capacidade (recurso, produto_ou_grupo, capacidade_mensal, observacao) VALUES (?,?,?,?)", [
        ("Cozinha e bar (equipe + equipamentos)", "Fabricação Própria", 7500, "≈ 290 itens por dia em 26 dias"),
        ("Equipe de atendimento e serviços", "Serviço", 2600, "Assinaturas, eventos e entregas"),
        ("Entregadores (2 motos)", "Delivery de 3 a 10 km", 1400, "≈ 54 entregas longas por dia"),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Distribuidora de carnes (exemplo)", "Frango, carne bovina, peixe", 28, 2, "Boleto 28 dias; pedido mínimo R$ 800", 1),
        ("Ceasa / hortifrúti local (exemplo)", "Verduras, legumes e frutas", 7, 1, "Pagamento semanal", 1),
        ("Atacadista de alimentos (exemplo)", "Arroz, feijão, grãos, óleos", 30, 3, "Boleto 30 dias", 0),
        ("Fornecedor de embalagens (exemplo)", "Marmitas, copos, etiquetas", 30, 7, "Boleto 30/60 dias", 1),
        ("Distribuidor de suplementos (exemplo)", "Whey, creatina, barras", 45, 5, "Nota fiscal com rastreabilidade", 0),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Alta no preço das proteínas", 4, 4, "Contratos de compra programada e cardápio com proteínas alternativas", "Sócio operacional"),
        ("Vendas abaixo do previsto nos primeiros meses", 3, 5, "Campanha de lançamento, parcerias e reserva de caixa", "Sócia comercial"),
        ("Interdição sanitária", 1, 5, "Boas práticas, checklists diários e consultoria técnica", "Gerente operacional"),
        ("Rotatividade de funcionários", 3, 3, "Treinamento, plano de carreira e bom clima de trabalho", "Sócia-administradora"),
        ("Atraso nas entregas em horário de pico", 3, 3, "Pedidos agendados e terceira moto em períodos de pico", "Gerente operacional"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Trabalhadores do centro pagariam R$ 29,90 por marmita com informação nutricional",
         "Venda-teste de 100 marmitas em 2 semanas via Instagram", "Unidades vendidas e recompra",
         "Vender 80 unidades e 30% de recompra", "", "Planejado"),
        ("Alunos de academia assinariam um plano mensal", "Oferta de plano em 3 academias parceiras",
         "Número de assinaturas", "15 assinaturas no primeiro mês", "", "Planejado"),
    ])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Abertura da empresa e registro de marca", 1, 1, "Sócia-administradora", 12000),
        ("Reforma e adequação sanitária", 1, 3, "Sócio operacional", 75000),
        ("Compra e instalação de equipamentos", 2, 2, "Sócio operacional", 330910),
        ("Contratação e treinamento da equipe", 3, 1, "Sócia-administradora", 10000),
        ("Campanha de pré-lançamento", 3, 1, "Sócia comercial", 25000),
        ("Inauguração", 4, 1, "Todos os sócios", 0),
    ])


# ==============================================================================
# SEGUNDO ESTUDO DE CASO: CAFÉ MUCURI (cafeteria e padaria artesanal)
# ==============================================================================
def popular_modelo_cafe(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='Café Mucuri', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Alimentos e bebidas; Padaria e confeitaria; Bares e restaurantes',
            cidade_uf='Teófilo Otoni - MG', contatos='@cafemucuri',
            aliquota_imposto_pct=8.0, crescimento_mensal_ano1_pct=1.5, crescimento_anual_pct=7.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=12.0, pme_dias=7, reserva_caixa_extra=25000,
            recuperar_giro=1, cen_otim_receita_pct=10, cen_otim_custo_pct=5,
            cen_pess_receita_pct=15, cen_pess_custo_pct=5,
            autores='Integrante 1; Integrante 2; Integrante 3', turma='Exemplo didático - setor de alimentação',
            tam_desc='Moradores de Teófilo Otoni que consomem café e lanches fora de casa (estimativa ilustrativa)',
            tam_clientes=45000, sam_desc='Público do centro e bairros vizinhos num raio de 2 km', sam_clientes=9000,
            som_pct=10, gasto_anual_cliente=900,
            mercado_fontes='Estimativas ilustrativas; substitua por IBGE, Sebrae, ABIC e pesquisa de campo.',
            mkt_produto='Cafés especiais de produtores do Vale do Mucuri e do Jequitinhonha, pães de fermentação natural e quitandas mineiras.',
            mkt_preco='Preço intermediário; combos de café da manhã e cartão fidelidade (10º café grátis).',
            mkt_praca='Loja de esquina no centro, retirada no balcão e encomendas pelo WhatsApp.',
            mkt_promocao='Instagram com a história dos produtores, degustações, parcerias com coworkings e eventos culturais.',
            loc_endereco='Centro de Teófilo Otoni - MG, próximo a escritórios, bancos e escolas',
            loc_justificativa='Fluxo de pedestres pela manhã e no fim da tarde, poucos cafés especiais na região.',
            loc_area_m2=90, loc_fluxo='Picos das 7h às 9h e das 15h às 18h.',
            op_layout='Padaria compacta nos fundos (forno de lastro, masseira, bancada), balcão de café e salão com 28 lugares.',
            op_processos='Fermentação noturna -> forneamento às 5h -> abertura às 6h30 -> produção de quitandas -> vendas -> fechamento e limpeza.',
            op_horario='Segunda a sábado, das 6h30 às 19h.', rateio_criterio='Receita', margem_desejada_pct=12
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Administração e finanças", "", "Gestão financeira, compras e pessoas.", 70000.0),
        ("Sócio(a) 2", "Produção e barista", "", "Panificação artesanal e preparo de cafés especiais.", 50000.0),
    ])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Trabalhadores do centro", "Pessoas que tomam café da manhã ou lanche perto do trabalho.", 9, 7, 9, 6, 9),
        ("Estudantes e professores", "Público das escolas e faculdades próximas.", 8, 6, 8, 7, 8),
        ("Apreciadores de cafés especiais", "Consumidores que valorizam origem e método de preparo.", 5, 9, 8, 8, 10),
        ("Famílias no fim de semana", "Compra de pães e quitandas para casa.", 7, 7, 9, 5, 7),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Luciana", 34, "Bancária", "Casada", 5200.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Chega ao centro às 7h30 e quer um café bem feito antes do expediente.",
         "Um bom café muda o meu dia, mas não tenho tempo de esperar.",
         "Usa Instagram, paga por aproximação, valoriza produtos locais.",
         "Rapidez; qualidade constante; ambiente agradável.", "Filas; café requentado; pouco lugar para sentar.",
         "O café especial cabe no meu orçamento diário?", "Pede pelo WhatsApp e retira pronto no balcão em 3 minutos."),
    ])
    ins("""INSERT INTO jornada_cliente (persona_nome, macro_fase, etapa, acoes_cliente, pensamentos,
           pontos_contato, estado_emocional, oportunidades) VALUES ('Luciana',?,?,?,?,?,?,?)""", [
        ("Antes", "Descoberta", "Vê post de colega no Instagram.", "Que lugar é esse?", "Instagram", "Neutro", "Conteúdo sobre os produtores."),
        ("Antes", "Consideração", "Compara com a padaria do bairro.", "Será caro?", "Cardápio na vitrine", "Neutro", "Combo café + pão de queijo."),
        ("Antes", "Decisão", "Passa na porta e sente o cheiro do pão.", "Vou experimentar.", "Fachada e vitrine", "Satisfeito", "Degustação na calçada."),
        ("Durante", "Compra e Pagamento", "Pede no balcão e paga por aproximação.", "Foi rápido.", "Caixa", "Satisfeito", "Pedido antecipado pelo WhatsApp."),
        ("Durante", "Entrega", "Recebe o café em xícara ou copo para viagem.", "Está quentinho.", "Balcão", "Satisfeito", "Copo reutilizável com desconto."),
        ("Durante", "Uso", "Consome no salão ou a caminho do trabalho.", "Que sabor!", "Produto", "Entusiasmado", "Cartão com a origem do grão."),
        ("Depois", "Avaliação", "Posta foto nos stories.", "Vou indicar.", "Instagram", "Entusiasmado", "Repostar clientes e marcar produtores."),
        ("Depois", "Suporte", "Pergunta sobre encomendas para festa.", "Fazem para 30 pessoas?", "WhatsApp", "Neutro", "Catálogo de encomendas."),
        ("Depois", "Lealdade e advocacia", "Usa o cartão fidelidade.", "Já é minha parada diária.", "Cartão fidelidade", "Entusiasmado", "Assinatura mensal de café."),
    ])
    concs = ["Padaria tradicional do centro", "Rede de cafeteria"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Qualidade do café": (9, 5, 7), "Pães e quitandas": (9, 8, 5), "Preço": (6, 9, 5),
             "Atendimento": (8, 6, 7), "Ambiente": (9, 5, 8), "Rapidez": (7, 8, 7), "Presença digital": (8, 3, 9)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Operação", "Panificação artesanal", "Força", "Alto", "Fermentação natural e receitas mineiras.", "Padronizar fichas técnicas.", 1),
        ("Interno", "Marketing", "Café de origem regional", "Força", "Alto", "Parceria direta com produtores locais.", "Contar a história dos produtores.", 1),
        ("Interno", "Finanças", "Capital limitado", "Fraqueza", "Alto", "Pouca reserva para os primeiros meses.", "Controle semanal de caixa.", 1),
        ("Interno", "Operação", "Espaço pequeno", "Fraqueza", "Baixo", "Salão de 28 lugares.", "Estimular o para viagem.", 0),
        ("Externo", "Social", "Valorização do consumo local", "Oportunidade", "Alto", "Clientes buscam produtos da região.", "Selo de origem nos produtos.", 1),
        ("Externo", "Econômico", "Alta do preço do café", "Ameaça", "Alto", "Café arábica subiu fortemente.", "Contratos com produtores e revisão de preços.", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,10)""", [
        ("Forno de lastro elétrico", 1, 22000, 10), ("Masseira espiral 15 kg", 1, 9500, 10),
        ("Máquina de café expresso 2 grupos", 1, 28000, 10), ("Moedor de café profissional", 2, 4500, 10),
        ("Vitrine refrigerada", 1, 7800, 10), ("Geladeira e freezer comerciais", 2, 5200, 10),
        ("Mesas e cadeiras", 7, 1400, 10), ("Balcão e mobiliário", 1, 15000, 10),
        ("Computador, PDV e impressora", 1, 6500, 5), ("Utensílios de padaria e café", 1, 6000, 5),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Reforma e adequação sanitária", 30000), ("Legalização e alvarás", 4000),
        ("Identidade visual e fachada", 6000), ("Marketing de inauguração", 5000), ("Treinamento da equipe", 2500),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 60000, 14.0, 36)])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 75.0), ("Recebimento", 30, 25.0), ("Pagamento", 0, 40.0), ("Pagamento", 30, 60.0)])
    ins("INSERT INTO materiais (nome, categoria, unidade, valor_unitario, qtd_estoque_inicial) VALUES (?,?,?,?,?)", [
        ("Café especial em grão", "Matéria-prima", "kg", 95.00, 15), ("Leite integral", "Matéria-prima", "L", 5.50, 60),
        ("Farinha de trigo", "Matéria-prima", "kg", 5.00, 100), ("Polvilho azedo", "Matéria-prima", "kg", 12.00, 30),
        ("Queijo minas meia cura", "Matéria-prima", "kg", 48.00, 10), ("Ovos", "Matéria-prima", "dúzia", 12.00, 20),
        ("Manteiga", "Matéria-prima", "kg", 45.00, 8), ("Açúcar", "Matéria-prima", "kg", 4.50, 30),
        ("Copo para viagem 200 ml", "Embalagem", "unidade", 0.40, 1000), ("Saco de papel kraft", "Embalagem", "unidade", 0.25, 1000),
    ])
    F, R = "Fabricação Própria", "Revenda"
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial)
           VALUES (?,?,?,?,?,?)""", [
        (F, "Café expresso", 6.50, 0, 2200, 0), (F, "Cappuccino", 11.00, 0, 900, 0),
        (F, "Café coado (método)", 12.00, 0, 450, 0), (F, "Pão de queijo (unidade)", 4.50, 0, 3000, 0),
        (F, "Pão de fermentação natural (500 g)", 22.00, 0, 520, 0), (F, "Bolo caseiro (fatia)", 9.00, 2.40, 800, 0),
        (F, "Biscoito de polvilho (pacote)", 8.00, 2.10, 400, 0),
        (R, "Café em grão 250 g (revenda)", 42.00, 26.00, 120, 30), (R, "Suco engarrafado", 9.00, 4.50, 300, 60),
    ])
    ficha = {
        "Café expresso": [("Café especial em grão", 0.009), ("Copo para viagem 200 ml", 0.4)],
        "Cappuccino": [("Café especial em grão", 0.009), ("Leite integral", 0.15), ("Copo para viagem 200 ml", 0.5)],
        "Café coado (método)": [("Café especial em grão", 0.018), ("Copo para viagem 200 ml", 0.3)],
        "Pão de queijo (unidade)": [("Polvilho azedo", 0.025), ("Queijo minas meia cura", 0.012), ("Leite integral", 0.01),
                                    ("Ovos", 0.008), ("Saco de papel kraft", 0.2)],
        "Pão de fermentação natural (500 g)": [("Farinha de trigo", 0.32), ("Saco de papel kraft", 1)],
    }
    ins("INSERT INTO ficha_tecnica (produto, material, quantidade) VALUES (?,?,?)",
        [(p, m, q) for p, itens in ficha.items() for m, q in itens])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [
        ("Taxas de cartões", 2.0), ("Perdas e quebras", 1.5), ("Propaganda variável", 0.5)])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Padeiro(a)", P, 1, 2300, 20, 200, "Massas, fermentação e forneamento."),
        ("Auxiliar de produção", P, 1, 1700, 20, 200, "Quitandas, limpeza e apoio."),
        ("Barista / atendente", C, 2, 1750, 20, 200, "Preparo de cafés, atendimento e caixa."),
    ])
    CIF, DA, DC, DF = "Custo indireto de produção", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Aluguel", DA, 3200), ("Energia elétrica (forno e máquinas)", CIF, 1500), ("Água e gás", CIF, 450),
        ("Pró-labore", DA, 4000), ("Contador", DA, 700), ("Internet e telefone", DA, 180),
        ("Manutenção de equipamentos", CIF, 300), ("Marketing fixo", DC, 400), ("Tarifas bancárias", DF, 90),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Luciana", "Rapidez; Qualidade; Origem regional; Ambiente acolhedor",
         "começar o dia com um café de qualidade sem perder tempo",
         "cafés especiais da região, pão de queijo saindo do forno e pedido antecipado pelo WhatsApp",
         "padarias tradicionais com café comum e redes sem identidade local",
         "sabor mineiro autêntico com rapidez e preço justo"),
    ])
    ins("INSERT INTO capacidade (recurso, produto_ou_grupo, capacidade_mensal, observacao) VALUES (?,?,?,?)", [
        ("Máquina de expresso 2 grupos", "Café expresso", 6000, "≈ 230 bebidas por dia"),
        ("Forno de lastro", "Fabricação Própria", 7000, "Fornadas das 5h às 11h"),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Produtor de café do Vale (exemplo)", "Café especial em grão", 30, 7, "Contrato trimestral de preço", 1),
        ("Laticínio regional (exemplo)", "Leite, queijo e manteiga", 14, 1, "Entrega 3x por semana", 1),
        ("Atacadista (exemplo)", "Farinha, açúcar, polvilho", 28, 3, "Boleto 28 dias", 0),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Alta do preço do café", 4, 4, "Contrato com produtor e ajuste trimestral de preços", "Sócio(a) 1"),
        ("Quebra do forno", 2, 5, "Manutenção preventiva e parceria com padaria vizinha", "Sócio(a) 2"),
        ("Vendas abaixo do previsto", 3, 4, "Degustações, combos e reserva de caixa", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Clientes do centro pagariam R$ 12 por café coado especial", "Barraca de degustação em feira por 2 sábados",
         "Cafés vendidos", "Vender 150 cafés", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Abertura da empresa e alvarás", 1, 1, "Sócio(a) 1", 4000), ("Reforma", 1, 2, "Sócio(a) 2", 30000),
        ("Equipamentos", 2, 1, "Sócio(a) 2", 120000), ("Contratação e treinamento", 3, 1, "Sócio(a) 1", 2500),
        ("Inauguração", 3, 1, "Todos", 5000),
    ])


# ==============================================================================
# TERCEIRO ESTUDO DE CASO: CONECTA VALE (aplicativo com planos de assinatura)
# Negócio fictício, para ensinar receita recorrente, churn, CAC/LTV e custos de tecnologia.
# ==============================================================================
def popular_modelo_assinatura(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='Conecta Vale', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Tecnologia e aplicativos; Turismo, cultura e lazer', cidade_uf='Teófilo Otoni - MG',
            contatos='@conectavale', tipo_negocio='Assinatura (receita recorrente)',
            horas_mes_pessoa=168, utilizacao_pct=70, anexo_simples='Anexo III - Serviços',
            aliquota_imposto_pct=6.0, crescimento_mensal_ano1_pct=3.0, crescimento_anual_pct=10.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=14.0, pme_dias=0, reserva_caixa_extra=50000,
            recuperar_giro=1, cen_otim_receita_pct=15, cen_otim_custo_pct=5, cen_pess_receita_pct=25,
            cen_pess_custo_pct=10, autores='Integrante 1; Integrante 2', turma='Exemplo didático - negócio por assinatura',
            tam_desc='Empresas de comércio, alimentação, cultura e turismo do Vale do Mucuri (estimativa ilustrativa)',
            tam_clientes=6000, sam_desc='Empresas de Teófilo Otoni com presença digital', sam_clientes=2000, som_pct=8,
            gasto_anual_cliente=2400,
            mercado_fontes='Estimativas ilustrativas; substitua por dados da Junta Comercial, IBGE e pesquisa com empresas.',
            mkt_produto='Aplicativo que reúne eventos, cultura, gastronomia e lazer da região. Empresas pagam planos para aparecer em  destaque; usuários usam grátis ou assinam o Premium com descontos.',
            mkt_preco='Três planos para empresas (mensal, trimestral e anual, com desconto para quem paga adiantado) e um plano  Premium barato para usuários. Primeiro mês grátis para empresas.',
            mkt_praca='Aplicativo nas lojas Google Play e App Store; venda consultiva às empresas, de porta em porta e pelo WhatsApp.',
            mkt_promocao='Parcerias com secretarias de cultura e turismo, influenciadores locais e semana de descontos mensal.',
            loc_endereco='Coworking no centro de Teófilo Otoni - MG (trabalho principalmente remoto)',
            loc_justificativa='Custo baixo e perto das empresas clientes para reuniões presenciais.',
            op_processos='Cadastro da empresa -> criação do perfil e das publicações -> aprovação -> divulgação no app ->  relatório mensal de visualizações -> renovação.',
            op_horario='Atendimento às empresas em horário comercial; aplicativo disponível 24 horas.',
            rateio_criterio='Receita', margem_desejada_pct=15
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Produto e tecnologia", "", "Desenvolvimento de aplicativos e gestão do produto.", 50000.0),
        ("Sócio(a) 2", "Comercial e marketing", "", "Vendas para empresas, parcerias e conteúdo.", 50000.0),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 21200, 14.0, 36)])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Bares, restaurantes e cafés", "Querem atrair clientes em dias de pouco movimento.", 9, 7, 8, 6, 9),
        ("Produtores de eventos e cultura", "Precisam divulgar shows, peças e festas.", 6, 6, 7, 7, 10),
        ("Hotéis e pousadas", "Buscam turistas e visitantes de negócios.", 4, 8, 8, 5, 7),
        ("Moradores e turistas (usuários)", "Querem saber o que fazer na cidade.", 10, 3, 8, 6, 8),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Marcos", 41, "Dono de restaurante", "Casado", 9000.0, "Ensino Médio", "Teófilo Otoni - MG",
         "Tem um restaurante com música ao vivo e faz a própria divulgação no Instagram.",
         "Eu pago impulsionamento, mas não sei se vem cliente por causa disso.",
         "Usa WhatsApp o dia todo; pouco tempo para marketing.", "Mesas cheias às quartas e quintas.",
         "Não sabe medir o retorno do marketing; caixa apertado.", "Vale a pena pagar todo mês?",
         "Ver no relatório quantas pessoas abriram a oferta e usaram o cupom."),
    ])
    concs = ["Plataforma nacional de eventos", "Perfis locais no Instagram"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Conteúdo local": (9, 4, 8), "Preço para pequenas empresas": (8, 5, 9), "Relatórios de resultado": (8, 7, 2),
             "Alcance de público": (5, 9, 7), "Facilidade de uso": (8, 8, 7)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Operação", "Equipe técnica própria", "Força", "Alto", "Sócio desenvolvedor e equipe enxuta.", "Lançar melhorias a cada mês.", 1),
        ("Interno", "Finanças", "Receita depende de renovações", "Fraqueza", "Alto", "Churn alto derruba a receita.", "Programa de sucesso do cliente.", 1),
        ("Externo", "Social", "Crescimento do turismo regional", "Oportunidade", "Alto", "Mais visitantes buscando o que fazer.", "Parcerias com o poder público.", 1),
        ("Externo", "Tecnológico", "Mudança nas taxas das lojas de aplicativos", "Ameaça", "Médio", "Comissões e regras mudam com frequência.", "Cobrar empresas fora da loja (boleto/Pix).", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,?)""", [
        ("Notebooks para desenvolvimento", 4, 6500, 5, 10), ("Celulares para testes (Android e iPhone)", 2, 3000, 3, 10),
        ("Mobiliário e equipamentos de escritório", 1, 3500, 10, 10),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Desenvolvimento da primeira versão do app (MVP)", 25000), ("Legalização e contratos", 2500),
        ("Registro de marca", 1500), ("Marketing de lançamento", 6000), ("Contas de desenvolvedor nas lojas", 700),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 100.0), ("Pagamento", 0, 100.0)])
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial, horas_unit)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Serviço", "Cobertura de evento patrocinado", 900, 120, 2, 0, 16),
        ("Serviço", "Produção de conteúdo avulso (post e vídeo)", 180, 10, 8, 0, 3),
    ])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [("Comissões de vendas", 5.0)])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Desenvolvedor(a)", P, 1, 5500, 28, 300, "Manter e evoluir o aplicativo, corrigir falhas."),
        ("Atendimento e conteúdo", P, 1, 2300, 28, 300, "Montar publicações, cobrir eventos, suporte às empresas."),
        ("Vendedor(a) externo(a)", C, 1, 2200, 28, 300, "Visitar empresas e fechar planos."),
    ])
    TEC, DA, DC, DF = "Custo de tecnologia / infraestrutura", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Servidores em nuvem (base fixa)", TEC, 700), ("Ferramentas e licenças de software", TEC, 450),
        ("Contas de desenvolvedor (rateio mensal)", TEC, 60), ("Coworking", DA, 900), ("Pró-labore", DA, 4000),
        ("Contador", DA, 500), ("Internet e telefone", DA, 200), ("Marketing digital (anúncios)", DC, 2000),
        ("Tarifas bancárias", DF, 80),
    ])
    ins("""INSERT INTO assinaturas (plano, publico, preco_ciclo, ciclo_meses, novos_mes1, crescimento_novos_pct,
           churn_mensal_pct, custo_entrega_mensal, taxa_cobranca_pct) VALUES (?,?,?,?,?,?,?,?,?)""", [
        ("Parceiro Essencial (mensal)", "Empresas (B2B)", 149, 1, 8, 4, 6, 10, 3.5),
        ("Parceiro Destaque (trimestral)", "Empresas (B2B)", 600, 3, 5, 4, 4, 14, 3.5),
        ("Parceiro Premium (anual)", "Empresas (B2B)", 3000, 12, 2, 4, 2, 18, 3.5),
        ("Usuário Premium (no app)", "Consumidores (B2C)", 6.90, 1, 50, 5, 8, 0.6, 15),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Marcos", "Clientes novos; Resultado medido; Preço acessível",
         "lotar o restaurante nos dias fracos sem perder tempo com marketing",
         "divulgação no aplicativo da cidade, cupons e relatório mensal de resultados",
         "impulsionamentos genéricos em redes sociais", "divulgação local com resultado que dá para medir"),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Churn maior que o previsto", 4, 5, "Acompanhamento mensal das empresas e relatórios de resultado", "Sócio(a) 2"),
        ("Falha ou queda do aplicativo", 2, 4, "Monitoramento, cópias de segurança e plano de contingência", "Sócio(a) 1"),
        ("Aumento do custo de nuvem com o crescimento", 3, 3, "Otimizar o sistema e negociar créditos de nuvem", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Restaurantes pagariam R$ 149/mês por divulgação no app", "Oferecer 30 dias grátis a 20 restaurantes",
         "Conversão para pago", "Pelo menos 6 assinam após o teste", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Desenvolvimento do MVP", 1, 3, "Sócio(a) 1", 25000), ("Teste com empresas piloto", 3, 1, "Sócio(a) 2", 0),
        ("Publicação nas lojas", 4, 1, "Sócio(a) 1", 700), ("Lançamento", 4, 1, "Todos", 6000),
    ])


# ==============================================================================
# QUARTO EXEMPLO: TECVALE SOLUÇÕES (SERVIÇOS DE TI POR HORA E POR PROJETO) - fictício
# ==============================================================================
def popular_modelo_servico(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='TecVale Soluções', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Tecnologia e aplicativos; Serviço', cidade_uf='Teófilo Otoni - MG',
            contatos='@tecvalesolucoes', tipo_negocio='Serviços (hora ou projeto)',
            horas_mes_pessoa=168, utilizacao_pct=70, anexo_simples='Anexo III - Serviços',
            aliquota_imposto_pct=6.0, crescimento_mensal_ano1_pct=2.0, crescimento_anual_pct=10.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=14.0, pme_dias=0, reserva_caixa_extra=25000,
            recuperar_giro=1, cen_otim_receita_pct=15, cen_otim_custo_pct=5, cen_pess_receita_pct=20,
            cen_pess_custo_pct=10, autores='Integrante 1; Integrante 2', turma='Exemplo didático - empresa de serviços',
            tam_desc='Micro e pequenas empresas do Vale do Mucuri (estimativa ilustrativa)',
            tam_clientes=9000, sam_desc='Micro e pequenas empresas de Teófilo Otoni com computador e presença digital',
            sam_clientes=3500, som_pct=5, gasto_anual_cliente=2400,
            mercado_fontes='Estimativas ilustrativas; substitua por dados do Sebrae (Data MPE), da Junta Comercial e de pesquisa com empresas.',
            mkt_produto='Suporte de informática, manutenção de computadores, criação de sites e lojas virtuais e treinamento para pequenas empresas. O carro-chefe é o contrato mensal de manutenção, que gera receita previsível.',
            mkt_preco='Hora técnica avulsa, pacotes fechados por projeto (sites e lojas virtuais) e contrato mensal com horas incluídas, mais barato que a hora avulsa.',
            mkt_praca='Atendimento no cliente, na bancada da empresa e remoto. Venda direta pelo WhatsApp, indicação e parcerias com contadores.',
            mkt_promocao='Diagnóstico gratuito de segurança e backup, conteúdo educativo no Instagram e programa de indicação.',
            loc_endereco='Sala comercial no centro de Teófilo Otoni - MG', loc_area_m2=45,
            loc_justificativa='Perto dos clientes do comércio e com espaço para a bancada de manutenção.',
            op_processos='Chamado pelo WhatsApp -> triagem e orçamento -> atendimento remoto ou no local -> registro das horas -> faturamento mensal.',
            op_horario='Segunda a sexta, 8h às 18h; plantão para clientes com contrato.',
            rateio_criterio='Horas', margem_desejada_pct=15
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Técnico(a) responsável", "", "Técnico em informática, 6 anos de experiência em suporte.", 30000.0),
        ("Sócio(a) 2", "Comercial e projetos web", "", "Desenvolvimento de sites e atendimento a clientes.", 30000.0),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 25000, 14.0, 36)])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Comércio varejista", "Lojas com sistema de vendas, impressoras e redes que não podem parar.", 9, 6, 8, 5, 9),
        ("Escritórios e clínicas", "Contabilidade, advocacia e saúde: dados sensíveis e muitos computadores.", 7, 8, 9, 6, 8),
        ("Novos negócios sem site", "Precisam de presença digital rápida e barata.", 8, 5, 6, 4, 7),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Cláudia", 47, "Dona de loja de roupas", "Casada", 8000.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Tem uma loja com duas filiais e depende do sistema de vendas e da maquininha.",
         "Quando o computador para, eu paro de vender.", "Resolve tudo pelo WhatsApp; chama o sobrinho para consertar o PC.",
         "Sistema funcionando e site que traga clientes.", "Atendimento demorado e sem garantia; medo de perder dados.",
         "Contrato mensal compensa?", "Um técnico que atende no mesmo dia e avisa antes de o problema acontecer."),
    ])
    concs = ["Técnicos autônomos", "Loja de informática do centro", "Agência digital da capital"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Rapidez no atendimento": (9, 6, 5, 4), "Preço": (7, 9, 6, 3), "Garantia do serviço": (8, 3, 6, 7),
             "Conhecimento técnico": (8, 5, 6, 9), "Atendimento presencial": (9, 8, 8, 2), "Sites e lojas virtuais": (7, 3, 2, 9)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Pessoas", "Sócios com experiência técnica e em sites", "Força", "Alto", "Atende do computador ao site.", "Pacotes que combinam os dois serviços.", 1),
        ("Interno", "Operação", "Receita depende das horas da equipe", "Fraqueza", "Alto", "Faturamento limitado pela capacidade.", "Contratos mensais e serviços remotos.", 1),
        ("Externo", "Tecnológico", "Digitalização das pequenas empresas", "Oportunidade", "Alto", "Pix, nota fiscal eletrônica e vendas on-line exigem estrutura.", "Pacote \"empresa digital\".", 1),
        ("Externo", "Legal", "LGPD exige cuidado com dados de clientes", "Oportunidade", "Médio", "Clínicas e escritórios precisam se adequar.", "Diagnóstico de segurança e backup.", 1),
        ("Externo", "Mercado", "Técnicos informais cobram muito pouco", "Ameaça", "Médio", "Concorrência por preço.", "Vender garantia, prazo e contrato.", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,?)""", [
        ("Notebooks para a equipe", 3, 4500, 5, 10), ("Bancada, ferramentas e estação de solda", 1, 4500, 5, 10),
        ("Mobiliário da sala", 1, 4000, 10, 10), ("Moto para atendimento externo", 1, 14000, 5, 40),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Legalização e contrato social", 2500), ("Licenças de software (primeiro ano)", 2500),
        ("Reforma e adequação da sala", 3000), ("Marketing de lançamento", 3000),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 50.0), ("Recebimento", 30, 50.0), ("Pagamento", 0, 100.0)])
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial, horas_unit)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Serviço", "Contrato mensal de manutenção (6 h incluídas)", 690, 20, 14, 0, 6),
        ("Serviço", "Hora técnica avulsa", 130, 0, 40, 0, 1),
        ("Serviço", "Formatação e configuração de computador", 190, 15, 16, 0, 2.5),
        ("Serviço", "Site institucional", 2900, 250, 2, 0, 40),
        ("Serviço", "Loja virtual", 5800, 400, 0.5, 0, 80),
        ("Serviço", "Treinamento em ferramentas digitais (turma)", 1300, 80, 1, 0, 12),
    ])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [("Taxas de cartão e Pix", 2.0)])
    P = "Produção / operação"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes, mes_inicio)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Técnico(a) de suporte", P, 2, 2400, 32, 350, "Atendimento remoto e presencial, manutenção de computadores e redes.", 1),
        ("Desenvolvedor(a) web", P, 1, 3500, 32, 350, "Criação de sites e lojas virtuais.", 1),
        ("Técnico(a) de suporte júnior", P, 1, 2000, 32, 350, "Contratação prevista para quando a ocupação passar de 90%.", 25),
    ])
    TEC, DA, DC, DF = "Custo de tecnologia / infraestrutura", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Softwares de acesso remoto e gestão de chamados", TEC, 350), ("Hospedagem e domínios dos clientes (base)", TEC, 150),
        ("Combustível e manutenção da moto", TEC, 450), ("Aluguel da sala", DA, 1500), ("Pró-labore", DA, 5000),
        ("Contador", DA, 500), ("Internet e telefone", DA, 300), ("Energia elétrica", DA, 300),
        ("Anúncios e conteúdo", DC, 700), ("Tarifas bancárias", DF, 80),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Distribuidora de peças de informática", "Memórias, SSDs, fontes, cabos", 28, 5, "Boleto 28 dias", 1),
        ("Provedor de hospedagem", "Hospedagem e domínios", 0, 0, "Cartão, mensal", 1),
        ("Software de acesso remoto", "Licenças por técnico", 0, 0, "Anual", 0),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Cláudia", "Rapidez; Garantia; Prevenção", "manter o sistema de vendas funcionando sem sustos",
         "contrato mensal com atendimento no mesmo dia, backup e manutenção preventiva",
         "técnicos avulsos que só aparecem quando o problema já aconteceu", "tranquilidade para vender todos os dias"),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Equipe sem horas livres para novos clientes", 3, 4, "Acompanhar a ocupação mensal e contratar no mês 25", "Sócio(a) 1"),
        ("Inadimplência dos contratos mensais", 3, 3, "Cobrança por boleto recorrente e suspensão após 30 dias", "Sócio(a) 2"),
        ("Perda de dados de um cliente durante o serviço", 2, 5, "Backup antes de cada intervenção e termo de responsabilidade", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Lojas fecham contrato mensal após um diagnóstico gratuito", "Oferecer diagnóstico a 25 lojas do centro",
         "Contratos fechados", "Pelo menos 5 contratos", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Formalização e sala", 1, 1, "Sócio(a) 2", 5500), ("Diagnósticos gratuitos", 1, 2, "Sócio(a) 1", 0),
        ("Lançamento do site e das redes", 1, 1, "Sócio(a) 2", 3000), ("Contratação do técnico júnior", 25, 1, "Sócio(a) 1", 0),
    ])



# ==============================================================================
# IDENTIDADE, OBJETIVOS SMART E MATRIZES ESTRATÉGICAS DOS TRÊS EXEMPLOS
# ==============================================================================
_ESTRATEGIA = {
    "fitbox": {
        "missao": "Oferecer refeições saudáveis, saborosas e práticas a quem tem pouco tempo, com ingredientes frescos e entrega rápida.",
        "visao": "Ser a principal referência em alimentação saudável de Teófilo Otoni até 2030.",
        "valores": "Qualidade; Respeito ao cliente; Transparência; Sustentabilidade; Valorização das pessoas",
        "rel": "Atendimento personalizado pelo WhatsApp\nCardápio semanal enviado aos clientes\nPrograma de fidelidade",
        "par": "Nutricionista parceira\nAcademias do bairro\nFornecedores de hortifrúti locais",
        "obj": [
            ("Atingir a receita prevista para o primeiro ano", "Financeira", "Receita bruta acumulada", 680000, "R$", 12, "Sócia administradora",
             "Meta baseada no volume e nos preços calculados no plano.", "Sustenta a viabilidade projetada.", "Acompanhamento semanal das vendas", ""),
            ("Fidelizar clientes com o programa de pontos", "Clientes", "Clientes que compram toda semana", 150, "clientes", 9, "Sócio comercial",
             "Base de clientes prevista e programa de baixo custo.", "Recompra reduz o custo de aquisição.", "Cartão fidelidade digital",
             "Atingir a receita prevista para o primeiro ano"),
            ("Reduzir o desperdício na produção", "Processos internos", "Perdas sobre a produção", 3, "% (máximo)", 6, "Chefe de cozinha",
             "Fichas técnicas já padronizadas.", "Desperdício corrói a margem de contribuição.", "Controle diário de sobras",
             "Atingir a receita prevista para o primeiro ano"),
            ("Treinar toda a equipe em boas práticas de manipulação", "Aprendizado e crescimento", "Funcionários certificados", 100, "%", 3,
             "Sócia administradora", "Curso gratuito do Sebrae/Anvisa disponível.", "Garante a segurança alimentar prometida.",
             "Curso de boas práticas", "Reduzir o desperdício na produção"),
        ],
        "bcg": [("Marmita fitness tradicional", 6, 1.4), ("Bowl de salmão", 15, 0.3), ("Suco natural 500 ml", 4, 0.6),
                ("Shake proteico com whey", 14, 1.1)],
        "porter": [(4, "Muitos restaurantes e marmitarias no centro.", "Diferenciar pela nutrição balanceada e pela entrega pontual."),
                   (4, "Investimento inicial baixo e receitas fáceis de copiar.", "Construir marca e fidelizar antes da concorrência."),
                   (3, "Cliente pode cozinhar em casa ou comprar congelados.", "Mostrar economia de tempo e variedade semanal."),
                   (3, "Cliente compara preço pelo aplicativo.", "Combos e planos semanais com desconto."),
                   (2, "Vários fornecedores de alimentos na região.", "Manter dois fornecedores para cada insumo crítico.")],
    },
    "cafe": {
        "missao": "Servir cafés especiais e quitandas mineiras feitas na hora, valorizando os produtores do Vale do Mucuri.",
        "visao": "Ser o café de referência do centro de Teófilo Otoni até 2029.",
        "valores": "Compromisso com a região; Qualidade; Excelência no atendimento; Confiança",
        "rel": "Atendimento pelo nome no balcão\nCartão fidelidade (10º café grátis)\nEventos de degustação",
        "par": "Produtores de café do Vale\nLaticínio regional\nCoworkings do centro",
        "obj": [
            ("Vender a média diária de cafés prevista no plano", "Financeira", "Cafés vendidos por dia", 140, "cafés/dia", 6, "Sócio(a) 2",
             "Fluxo de pedestres do centro e capacidade da máquina.", "O café é o produto de maior margem.", "Combos de café da manhã", ""),
            ("Conquistar clientes recorrentes do centro", "Clientes", "Clientes com cartão fidelidade ativo", 400, "clientes", 12, "Sócio(a) 1",
             "Cerca de 9.000 pessoas no público alcançável.", "Recompra diária sustenta a receita.", "Parcerias com escritórios",
             "Vender a média diária de cafés prevista no plano"),
            ("Padronizar o preparo de todas as bebidas", "Processos internos", "Bebidas com ficha de preparo", 100, "%", 3, "Sócio(a) 2",
             "Cardápio enxuto e barista experiente.", "Qualidade constante é a promessa da marca.", "Fichas de preparo e treinamento",
             "Conquistar clientes recorrentes do centro"),
            ("Formar a equipe em atendimento e barismo", "Aprendizado e crescimento", "Horas de treinamento por atendente", 16, "horas", 4,
             "Sócio(a) 2", "Treinamento interno com o sócio barista.", "Atendimento é diferencial frente às redes.",
             "Treinamentos mensais", "Padronizar o preparo de todas as bebidas"),
        ],
        "canais": ["Loja física (balcão)", "Instagram", "WhatsApp (encomendas)", "Indicação (boca a boca)"],
        "bcg": [("Café expresso", 8, 1.1), ("Pão de queijo (unidade)", 5, 1.5), ("Café coado (método)", 14, 0.5)],
        "porter": [(3, "Padarias tradicionais e uma rede de cafeteria por perto.", "Foco em café especial e origem regional."),
                   (3, "Abrir uma cafeteria exige investimento moderado em equipamentos.", "Fidelizar e criar marca local forte."),
                   (3, "Café de máquina no trabalho e cafés de padaria.", "Experiência e qualidade que justificam o preço."),
                   (3, "Cliente individual, compra pequena, mas compara com a padaria.", "Combos e programa de fidelidade."),
                   (4, "Café especial depende de poucos produtores e o preço do grão oscila.", "Contratos trimestrais e mais de um produtor.")],
    },
    "assinatura": {
        "missao": "Conectar moradores e visitantes ao que acontece no Vale do Mucuri e ajudar pequenas empresas locais a atrair clientes.",
        "visao": "Ser o principal aplicativo de cultura, lazer e consumo local do Vale do Mucuri até 2030.",
        "valores": "Compromisso com a região; Transparência; Inovação; Simplicidade",
        "rel": "Gerente de conta para empresas\nRelatório mensal de resultados\nSuporte pelo WhatsApp\nComunidade de usuários no app",
        "par": "Secretarias de cultura e turismo\nProdutores de eventos\nInfluenciadores locais\nProvedor de nuvem",
        "obj": [
            ("Atingir a receita recorrente mensal prevista para o fim do primeiro ano", "Financeira", "MRR", 37000, "R$/mês", 12,
             "Sócio(a) 2", "Projeção por safras com churn e captação realistas.", "MRR sustenta a viabilidade do aplicativo.",
             "Metas mensais de vendas por plano", ""),
            ("Manter o cancelamento de empresas abaixo do previsto", "Clientes", "Churn mensal das empresas", 5, "% (máximo)", 12,
             "Sócio(a) 2", "Relatórios mensais mostram resultado ao cliente.", "Cada ponto de churn reduz o LTV.",
             "Programa de sucesso do cliente", "Atingir a receita recorrente mensal prevista para o fim do primeiro ano"),
            ("Publicar as divulgações das empresas em até 24 horas", "Processos internos", "Publicações no prazo", 95, "%", 6,
             "Atendimento e conteúdo", "Fluxo de aprovação simples e equipe dedicada.", "Agilidade é o que as empresas cobram.",
             "Fila de publicação com prazos", "Manter o cancelamento de empresas abaixo do previsto"),
            ("Lançar uma melhoria do aplicativo por mês", "Aprendizado e crescimento", "Versões publicadas", 12, "versões/ano", 12,
             "Sócio(a) 1", "Equipe técnica própria e ciclo curto de desenvolvimento.", "Produto melhor reduz o churn.",
             "Roteiro mensal de melhorias", "Publicar as divulgações das empresas em até 24 horas"),
        ],
        "canais": ["Aplicativo (Google Play e App Store)", "Venda consultiva às empresas", "Instagram", "Acordos e Parcerias"],
        "bcg": [("Cobertura de evento patrocinado", 18, 0.3), ("Produção de conteúdo avulso (post e vídeo)", 12, 0.2),
                ("Assinaturas (todos os planos)", 25, 0.6)],
        "porter": [(3, "Plataformas nacionais e perfis locais de Instagram disputam a atenção.", "Conteúdo hiperlocal e relatórios de resultado."),
                   (4, "Criar um aplicativo é relativamente barato; a barreira é a base de usuários.", "Crescer rápido a comunidade e as parcerias."),
                   (4, "Redes sociais gratuitas cumprem parte da função.", "Mostrar resultado medido (cupons usados, visitas)."),
                   (3, "Empresas pequenas cancelam com facilidade.", "Plano anual com desconto e acompanhamento mensal."),
                   (3, "Dependência de lojas de aplicativos e provedor de nuvem.", "Cobrar empresas fora da loja e evitar dependência de um só provedor.")],
    },
    "servico": {
        "missao": "Manter a tecnologia das pequenas empresas do Vale do Mucuri funcionando, com atendimento rápido, garantia e prevenção.",
        "visao": "Ser a empresa de suporte de TI mais lembrada pelo comércio de Teófilo Otoni até 2030.",
        "valores": "Agilidade; Honestidade no diagnóstico; Segurança dos dados; Aprendizado contínuo",
        "rel": "Atendimento pelo WhatsApp com prazo de resposta\nVisita preventiva mensal aos clientes com contrato\nRelatório de chamados",
        "par": "Escritórios de contabilidade\nDistribuidora de peças\nProvedor de hospedagem\nSebrae (capacitação de clientes)",
        "obj": [
            ("Chegar a 25 contratos mensais de manutenção", "Financeira", "Contratos ativos", 25, "contratos", 12, "Sócio(a) 2",
             "Começa com 14 e o diagnóstico gratuito gera novos contratos.", "Contrato dá receita previsível e reduz ociosidade.",
             "Diagnóstico gratuito e indicação de contadores", ""),
            ("Atender os chamados no mesmo dia", "Clientes", "Chamados resolvidos em até 8 horas", 90, "%", 6, "Sócio(a) 1",
             "Equipe de dois técnicos e acesso remoto.", "Rapidez é o principal motivo de escolha.", "Fila de chamados com prazo",
             "Chegar a 25 contratos mensais de manutenção"),
            ("Manter a ocupação da equipe entre 70% e 90%", "Processos internos", "Horas vendidas ÷ horas cobráveis", 85, "%", 12,
             "Sócio(a) 1", "Ocupação calculada na ferramenta de serviços.", "Ocupação baixa é custo; alta demais derruba a qualidade.",
             "Registro de horas por chamado", "Atender os chamados no mesmo dia"),
            ("Certificar os técnicos em redes e segurança", "Aprendizado e crescimento", "Técnicos certificados", 2, "técnicos", 9,
             "Sócio(a) 1", "Cursos gratuitos e de baixo custo disponíveis.", "Abre o serviço de adequação à LGPD.",
             "Plano de cursos", "Manter a ocupação da equipe entre 70% e 90%"),
        ],
        "canais": ["Indicação (boca a boca)", "WhatsApp Business", "Instagram", "Acordos e Parcerias"],
        "bcg": [("Contrato mensal de manutenção (6 h incluídas)", 12, 1.5), ("Hora técnica avulsa", 3, 0.8),
                ("Site institucional", 9, 0.4), ("Loja virtual", 16, 0.2), ("Formatação e configuração de computador", 1, 1.1)],
        "porter": [(4, "Muitos técnicos autônomos e lojas de informática.", "Vender contrato, garantia e prazo, não só a hora."),
                   (4, "Basta conhecimento e um notebook para abrir.", "Construir carteira de contratos e reputação."),
                   (3, "Suporte do próprio fabricante e sobrinho que entende de computador.", "Mostrar o custo de uma parada do sistema."),
                   (3, "Pequenas empresas comparam preço e negociam.", "Planos com horas incluídas e preço fechado."),
                   (2, "Peças e softwares têm vários fornecedores.", "Dois fornecedores para peças críticas.")],
    },
}


def popular_estrategia(conn, qual: str) -> None:
    e = _ESTRATEGIA[qual]
    conn.execute("UPDATE projeto SET missao=?, visao=?, valores=?, canvas_relacionamento=?, canvas_parcerias=? WHERE id=1",
                 (e["missao"], e["visao"], e["valores"], e["rel"], e["par"]))
    conn.executemany("""INSERT INTO objetivos (objetivo, perspectiva, indicador, meta, unidade, prazo_meses, responsavel,
                        por_que_atingivel, relevancia, iniciativa, contribui_para) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", e["obj"])
    conn.executemany("INSERT INTO bcg (produto, crescimento_mercado_pct, participacao_relativa) VALUES (?,?,?)", e["bcg"])
    if e.get("canais") and conn.execute("SELECT COUNT(*) FROM canais").fetchone()[0] == 0:
        conn.executemany("INSERT INTO canais (canal, atratividade, alcance, cliente, relacionamento, viabilidade) VALUES (?,?,?,?,?,?)",
                         [(c, 8, 7, 8, 7, 8) for c in e["canais"]])
    forcas = ["Rivalidade entre concorrentes", "Ameaça de novos entrantes", "Ameaça de produtos substitutos",
              "Poder de barganha dos clientes", "Poder de barganha dos fornecedores"]
    conn.executemany("INSERT INTO porter (forca, intensidade, justificativa, estrategia) VALUES (?,?,?,?)",
                     [(f, *v) for f, v in zip(forcas, e["porter"])])
    popular_validacao(conn, qual)


# ------------------------------------------------------------------------------
# DADOS DE EXEMPLO DAS FERRAMENTAS DE VALIDAÇÃO E MARKETING (v7.1) - fictícios
# ------------------------------------------------------------------------------
def _respostas_vw(preco: float, n: int, semente: int) -> list:
    """Respostas fictícias e coerentes do teste de Van Westendorp em torno de um preço."""
    import random
    rnd = random.Random(semente)
    linhas = []
    for i in range(n):
        base = preco * rnd.uniform(0.75, 1.25)
        mb = round(base * rnd.uniform(0.45, 0.65), 0)
        ba = round(base * rnd.uniform(0.75, 0.9), 0)
        ca = round(base * rnd.uniform(1.05, 1.25), 0)
        mc = round(base * rnd.uniform(1.35, 1.7), 0)
        linhas.append((f"R{i + 1:02d}", mb, max(ba, mb), max(ca, ba), max(mc, ca)))
    return linhas


_VALIDACAO = {
    "fitbox": {
        "vw_produto": "Assinatura semanal - 5 marmitas tradicionais", "vw_preco": 129.9, "vw_n": 32,
        "intencao": (18.0, 31.0),
        "entrevistas": [
            ("03/02", "E01", "Trabalhadora de escritório", "falta de tempo; comida pouco saudável perto do trabalho", "marmita de casa", "Sim", 130, "Quer variar o cardápio na semana"),
            ("03/02", "E02", "Praticante de academia", "dificuldade de controlar macros; falta de tempo", "prepara no domingo", "Sim", 150, "Valoriza tabela nutricional"),
            ("04/02", "E03", "Estudante universitário", "preço alto; comida pouco saudável perto do trabalho", "restaurante a quilo", "Talvez", 90, "Sensível a preço"),
            ("04/02", "E04", "Profissional liberal", "falta de tempo; enjoa da mesma comida", "aplicativo de delivery", "Sim", 140, "Pediu entrega até 11h"),
            ("05/02", "E05", "Mãe que trabalha fora", "falta de tempo; dificuldade de controlar macros", "congela comida", "Sim", 120, "Quer opção para a família"),
            ("05/02", "E06", "Praticante de corrida", "dificuldade de controlar macros", "marmitaria fitness", "Talvez", 110, "Já usa concorrente"),
            ("06/02", "E07", "Servidor público", "comida pouco saudável perto do trabalho; falta de tempo", "restaurante a quilo", "Sim", 130, "Almoça no trabalho"),
            ("06/02", "E08", "Estudante universitário", "preço alto", "lanche", "Não", 60, "Não é o público prioritário"),
            ("07/02", "E09", "Trabalhadora de escritório", "falta de tempo; enjoa da mesma comida", "marmita de casa", "Sim", 125, "Indicaria a colegas"),
            ("07/02", "E10", "Personal trainer", "dificuldade de controlar macros", "prepara em casa", "Talvez", 140, "Possível parceiro de indicação"),
            ("08/02", "E11", "Profissional de saúde", "falta de tempo; comida pouco saudável perto do trabalho", "aplicativo de delivery", "Sim", 135, "Plantões longos"),
            ("08/02", "E12", "Praticante de academia", "falta de tempo; dificuldade de controlar macros", "marmitaria fitness", "Sim", 145, "Quer suplementos junto"),
        ],
        "evidencias": {"Preço de venda": ("Testada com clientes (entrevista, pré-venda, piloto)", "32 respostas do teste de preço e 12 entrevistas (fev.)"),
                       "Custo de materiais e mercadorias": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Orçamentos de 3 fornecedores"),
                       "Salários e encargos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Piso da categoria e tabela do contador"),
                       "Demais custos fixos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Aluguel e contas cotados")},
        "vpc": ("Quando tenho pouco tempo no almoço, quero comer algo saudável e saboroso, para manter a dieta sem perder o horário.",
                "Almoçar em pouco tempo\nSeguir a dieta da nutricionista\nPlanejar a alimentação da semana",
                "Falta de tempo\nComida pouco saudável perto do trabalho\nDificuldade de controlar macros\nEnjoa da mesma comida",
                "Sentir mais disposição\nVer resultado no treino\nEconomizar com delivery caro",
                "Marmitas fitness\nAssinatura semanal e mensal\nCoffee break saudável",
                "Entrega no trabalho até 11h\nTabela nutricional em cada marmita\nCardápio que muda toda semana",
                "Acompanhamento com nutricionista parceira\nPreço de assinatura menor que o delivery"),
        "encaixe": [("Falta de tempo", "Dor", "Entrega no trabalho até 11h"),
                    ("Comida pouco saudável perto do trabalho", "Dor", "Tabela nutricional em cada marmita"),
                    ("Dificuldade de controlar macros", "Dor", "Tabela nutricional em cada marmita"),
                    ("Enjoa da mesma comida", "Dor", "Cardápio que muda toda semana"),
                    ("Ver resultado no treino", "Ganho", "Acompanhamento com nutricionista parceira"),
                    ("Economizar com delivery caro", "Ganho", "Preço de assinatura menor que o delivery")],
        "errc": [("Variedade", "Reduzir", 6, "Cardápio enxuto e rotativo reduz custo e desperdício"),
                 ("Design", "Reduzir", 5, "Embalagem simples e reciclável"),
                 ("Transparência nutricional", "Elevar", 10, "Tabela completa e macros em cada marmita"),
                 ("Conveniência", "Elevar", 9, "Entrega programada no trabalho"),
                 ("Personalização", "Criar", 8, "Ajuste de porções conforme o objetivo do cliente")],
        "posic": ("trabalhadores e praticantes de atividade física de Teófilo Otoni", "não têm tempo de preparar refeições saudáveis",
                  "FitBox", "uma assinatura de refeições saudáveis", "entrega no trabalho refeições balanceadas com controle de macros",
                  "marmitarias comuns e aplicativos de delivery", "cada marmita traz a tabela nutricional e o cardápio é planejado por nutricionista"),
        "mapa": ("Preço (baixo → alto)", "Saudabilidade (baixa → alta)",
                 [("FitBox", 6, 9), ("Concorrente A (marmitaria fitness)", 7, 7), ("Concorrente B (restaurante natural)", 8, 8),
                  ("Concorrente C (loja de suplementos)", 5, 6), ("Restaurante a quilo", 4, 4)]),
        "orc": [("Publicidade em Redes Sociais", 1500, 45, "Cupom por anúncio"), ("Acordos e Parcerias", 600, 25, "Código de academia parceira"),
                ("Marketing de Conteúdo", 400, 12, "Link na bio"), ("Eventos", 800, 10, "Inscrições no estande")],
        "aarrr": [("Aquisição", "Visitas ao perfil e cadastros", "300 cadastros/mês", "Anúncios e parcerias com academias"),
                  ("Ativação", "Primeiro pedido em 7 dias", "40% dos cadastros", "Cupom de primeira semana"),
                  ("Retenção", "Assinantes que renovam no mês", "75%", "Cardápio rotativo e avaliação semanal"),
                  ("Receita", "Ticket médio mensal", "R$ 450", "Upgrade para plano mensal"),
                  ("Indicação", "Clientes vindos de indicação", "20%", "Uma marmita grátis por indicação")],
        "cal": [(1, "Lançamento com degustação em academias", "Eventos", 1200), (1, "Campanha de anúncios de abertura", "Redes sociais", 1500),
                (3, "Parceria com nutricionistas", "Parcerias", 600), (6, "Campanha de indicação", "Indicação", 500)],
        "perda": [40000, 30000, 25000, 15000],
        "formal": {"Contratar contador ou escritório contábil": "Concluído",
                   "Consultar a viabilidade do endereço na prefeitura (zoneamento)": "Em andamento",
                   "Escolher a natureza jurídica (MEI, empresário individual, SLU ou LTDA)": "Concluído"},
        "banca": [("Por que este é o momento certo para abrir este negócio?",
                   "A procura por alimentação saudável cresce na cidade e nenhum concorrente entrega no local de trabalho com controle de macros. "
                   "As 12 entrevistas mostraram que falta de tempo é a dor mais citada.")],
    },
    "cafe": {
        "vw_produto": "Cappuccino", "vw_preco": 11.0, "vw_n": 20, "intencao": (12.0, 25.0), "entrevistas": [],
        "evidencias": {"Demais custos fixos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Aluguel cotado")},
        "vpc": ("Quando quero uma pausa no meio da tarde, quero um café de qualidade num lugar agradável, para descansar e encontrar amigos.",
                "Fazer uma pausa no trabalho\nEncontrar amigos\nComprar pão para casa",
                "Café fraco ou requentado\nFila na padaria\nPoucos lugares para sentar", "Ambiente acolhedor\nProdutos artesanais",
                "Cafés especiais\nPães de fermentação natural\nBolos caseiros", "Café moído na hora\nAtendimento rápido no balcão",
                "Ambiente com mesas e Wi-Fi\nProdutos da região"),
        "errc": [("Rapidez", "Manter", 0, ""), ("Ambiente", "Elevar", 9, "Mesas, tomadas e Wi-Fi"),
                 ("Presença digital", "Elevar", 8, "Encomendas pelo WhatsApp")],
        "posic": ("moradores e trabalhadores do centro", "querem um café de qualidade sem pagar preço de rede", "Café Mucuri",
                  "uma cafeteria de bairro", "serve café especial e pães artesanais com preço justo",
                  "padarias tradicionais e redes de cafeteria", "usamos grãos e farinhas de produtores da região"),
        "mapa": ("Preço (baixo → alto)", "Qualidade do café (baixa → alta)",
                 [("Café Mucuri", 5, 8), ("Padaria tradicional do centro", 3, 4), ("Rede de cafeteria", 8, 8)]),
        "orc": [("Instagram", 600, 30, "Cupom do perfil"), ("Indicação (boca a boca)", 200, 20, "Cartão fidelidade")],
        "aarrr": [], "cal": [], "perda": [35000, 25000], "formal": {}, "banca": [],
    },
    "assinatura": {
        "vw_produto": "Assinatura (ticket médio)", "vw_preco": 0, "vw_n": 0, "intencao": (8.0, 20.0), "entrevistas": [],
        "evidencias": {"Cancelamento mensal (churn)": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Referências de SaaS para pequenas empresas")},
        "vpc": ("Quando organizo um evento na cidade, quero que o público certo saiba dele, para encher o evento sem gastar muito com anúncios.",
                "Divulgar eventos\nAtrair público local\nMedir o resultado da divulgação",
                "Anúncio caro e pouco segmentado\nNão sabe se a divulgação deu resultado", "Público certo no evento\nRelatório para os patrocinadores",
                "Aplicativo de agenda local\nPlanos de parceiro\nCobertura de eventos", "Público segmentado por cidade e interesse\nRelatório mensal de visualizações",
                "Selo de parceiro destaque\nCobertura em vídeo"),
        "errc": [("Conteúdo local", "Elevar", 10, "Agenda completa do Vale"), ("Relatórios de resultado", "Criar", 8, "Relatório mensal ao parceiro")],
        "posic": ("produtores de eventos e pequenos comércios do Vale do Mucuri", "querem divulgar para o público local", "Conecta Vale",
                  "um aplicativo de agenda e divulgação local", "leva o evento ao público certo com relatório de resultado",
                  "perfis locais no Instagram", "o parceiro recebe todo mês o relatório de visualizações e cliques"),
        "mapa": ("Preço (baixo → alto)", "Alcance local (baixo → alto)",
                 [("Conecta Vale", 4, 8), ("Plataforma nacional de eventos", 7, 4), ("Perfis locais no Instagram", 3, 6)]),
        "orc": [("Venda consultiva às empresas", 1200, 8, "Contratos fechados"), ("Instagram", 500, 4, "Cadastros pelo link")],
        "aarrr": [], "cal": [], "perda": [30000, 30000], "formal": {}, "banca": [],
    },
    "servico": {
        "vw_produto": "Contrato mensal de manutenção (6 h incluídas)", "vw_preco": 690, "vw_n": 24, "intencao": (10.0, 22.0),
        "entrevistas": [
            ("10/03", "S01", "Loja de roupas", "computador para de funcionar; técnico demora", "sobrinho", "Sim", 600, "Perde vendas quando o sistema para"),
            ("10/03", "S02", "Escritório de contabilidade", "medo de perder dados; técnico demora", "técnico autônomo", "Sim", 900, "Quer backup diário"),
            ("11/03", "S03", "Clínica odontológica", "medo de perder dados; LGPD", "empresa da capital", "Talvez", 800, "Pede contrato com garantia"),
            ("11/03", "S04", "Mercearia", "computador para de funcionar", "loja de informática", "Não", 250, "Prefere pagar só quando quebra"),
            ("12/03", "S05", "Loja de calçados", "técnico demora; computador para de funcionar", "técnico autônomo", "Sim", 650, "Indica outros lojistas"),
            ("12/03", "S06", "Imobiliária", "não tem site; medo de perder dados", "nenhuma", "Talvez", 500, "Interesse maior no site"),
        ],
        "encaixe": [],
        "evidencias": {"Salários e encargos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Piso da categoria e tabela do contador"),
                       "Preço de venda": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Preços de 4 concorrentes locais")},
        "vpc": ("Quando o computador ou o sistema da loja para, quero que alguém resolva no mesmo dia, para não perder vendas.",
                "Manter o sistema de vendas funcionando\nGuardar os dados com segurança\nTer presença na internet",
                "Técnico demora\nComputador para de funcionar\nMedo de perder dados\nNão tem site",
                "Previsibilidade de gastos\nSaber que os dados estão seguros",
                "Contrato mensal de manutenção\nHora técnica\nSites e lojas virtuais",
                "Atendimento no mesmo dia\nManutenção preventiva\nBackup automático\nSite em pacote fechado",
                "Preço fixo mensal\nRelatório de backup"),
        "errc": [("Preço", "Reduzir", 6, "Não competir com o técnico informal pelo menor preço"),
                 ("Garantia do serviço", "Elevar", 10, "Garantia por escrito de 90 dias"),
                 ("Rapidez no atendimento", "Elevar", 10, "Atendimento no mesmo dia para contratos"),
                 ("Manutenção preventiva", "Criar", 8, "Visita mensal antes de o problema acontecer")],
        "posic": ("pequenas empresas do comércio e de serviços de Teófilo Otoni", "não podem parar quando o computador falha",
                  "TecVale Soluções", "um serviço de suporte de TI por contrato", "resolve no mesmo dia e evita que o problema aconteça",
                  "técnicos avulsos e lojas de informática", "o contrato inclui visita preventiva mensal, backup e garantia por escrito"),
        "mapa": ("Preço (baixo → alto)", "Rapidez e garantia (baixa → alta)",
                 [("TecVale Soluções", 6, 9), ("Técnicos autônomos", 3, 4), ("Loja de informática do centro", 5, 5),
                  ("Agência digital da capital", 9, 5)]),
        "orc": [("Indicação (boca a boca)", 200, 3, "Desconto de um mês por indicação"), ("Instagram", 500, 2, "Mensagens pelo link"),
                ("Acordos e Parcerias", 300, 2, "Contadores parceiros")],
        "aarrr": [], "cal": [], "perda": [20000, 20000], "formal": {}, "banca": [],
    },
}


def popular_validacao(conn, qual: str) -> None:
    v = _VALIDACAO[qual]
    vpc = v["vpc"]
    conn.execute("""UPDATE projeto SET vw_produto=?, intencao_certamente_pct=?, intencao_provavel_pct=?, vpc_job=?, vpc_tarefas=?,
                    vpc_dores=?, vpc_ganhos=?, vpc_produtos=?, vpc_aliviadores=?, vpc_criadores=?, posic_segmento=?,
                    posic_necessidade=?, posic_marca=?, posic_categoria=?, posic_beneficio=?, posic_concorrente=?,
                    posic_diferencial=?, mapa_eixo_x=?, mapa_eixo_y=? WHERE id=1""",
                 (v["vw_produto"], *v["intencao"], *vpc, *v["posic"], v["mapa"][0], v["mapa"][1]))
    if v["vw_n"]:
        conn.executemany("INSERT INTO precos_vw (respondente, muito_barato, barato, caro, muito_caro) VALUES (?,?,?,?,?)",
                         _respostas_vw(v["vw_preco"], v["vw_n"], 7))
    conn.executemany("""INSERT INTO entrevistas (data, entrevistado, perfil, dores, solucao_atual, pagaria, valor_pagaria, aprendizado)
                        VALUES (?,?,?,?,?,?,?,?)""", v["entrevistas"])
    conn.executemany("INSERT INTO evidencias (premissa, status, fonte, observacao) VALUES (?,?,?,'')",
                     [(k, s, f) for k, (s, f) in v["evidencias"].items()])
    conn.executemany("INSERT INTO vpc_encaixe (item, tipo, resposta) VALUES (?,?,?)", v.get("encaixe", []))
    conn.executemany("INSERT INTO errc (atributo, acao, nova_nota, justificativa) VALUES (?,?,?,?)", v["errc"])
    conn.executemany("INSERT INTO mapa_perceptual (marca, eixo_x, eixo_y) VALUES (?,?,?)", v["mapa"][2])
    conn.executemany("INSERT INTO orcamento_canais (canal, gasto_mensal, clientes_mes, observacao) VALUES (?,?,?,?)", v["orc"])
    conn.executemany("INSERT INTO aarrr (etapa, metrica, meta, acao) VALUES (?,?,?,?)", v["aarrr"])
    conn.executemany("INSERT INTO calendario_mkt (mes, acao, canal, custo) VALUES (?,?,?,?)", v["cal"])
    ids = [r[0] for r in conn.execute("SELECT id FROM socios ORDER BY id")]
    conn.executemany("UPDATE socios SET perda_aceitavel=? WHERE id=?", list(zip(v["perda"], ids)))
    conn.executemany("INSERT INTO formalizacao (item, status, observacao) VALUES (?,?,'')", list(v["formal"].items()))
    conn.executemany("INSERT INTO banca_respostas (pergunta, resposta) VALUES (?,?)", v["banca"])
