"""
================================================================================
PNBOX ACADÊMICO - ESTUDO DE CASO FITBOX (pnbox_modelo.py)
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
