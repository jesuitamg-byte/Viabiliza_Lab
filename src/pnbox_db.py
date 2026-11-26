"""
================================================================================
VIABILIZALAB - CAMADA DE DADOS (pnbox_db.py)
================================================================================
Este módulo cuida de tudo o que envolve o banco de dados SQLite (.pnbox):
    - estrutura das tabelas (SCHEMA) e atualização automática de arquivos antigos;
    - criação de projetos novos (em branco ou a partir do modelo FitBox);
    - leitura e gravação das tabelas.

Cada projeto é um arquivo .pnbox independente, guardado na pasta "projetos",
ao lado do programa. O aluno pode baixar esse arquivo e enviá-lo ao professor.
================================================================================
"""

import sqlite3
from pathlib import Path

import pandas as pd

# ------------------------------------------------------------------------------
# 1. ESTRUTURA DAS TABELAS
# ------------------------------------------------------------------------------
# Tabelas "únicas" têm apenas uma linha (id = 1). As demais são listas.
TABELAS_UNICAS = {"projeto", "proposta_valor"}
# Tabelas que não são editadas como planilha comum
TABELAS_INTERNAS = {"projeto", "proposta_valor", "versoes"}

SCHEMA = {
    "projeto": [
        ("nome_negocio", "TEXT"), ("objetivo", "TEXT"), ("enquadramento", "TEXT"),
        ("periodo_anos", "INTEGER"), ("porte", "TEXT"), ("setor_atuacao", "TEXT"),
        ("cidade_uf", "TEXT"), ("contatos", "TEXT"),
        ("aliquota_imposto_pct", "REAL"),          # Impostos sobre vendas (ex.: Simples)
        ("crescimento_mensal_ano1_pct", "REAL"),
        ("crescimento_anual_pct", "REAL"),
        ("inflacao_custos_anual_pct", "REAL"),
        ("tma_anual_pct", "REAL"),
        ("pme_dias", "REAL"),                      # Prazo médio de estoques
        ("reserva_caixa_extra", "REAL"),
        ("recuperar_giro", "INTEGER"),             # Recupera capital de giro no último ano do VPL
        ("cen_otim_receita_pct", "REAL"), ("cen_otim_custo_pct", "REAL"),
        ("cen_pess_receita_pct", "REAL"), ("cen_pess_custo_pct", "REAL"),
        # Identificação acadêmica
        ("autores", "TEXT"), ("turma", "TEXT"), ("professor", "TEXT"), ("data_entrega", "TEXT"),
        # Mercado (TAM / SAM / SOM)
        ("tam_desc", "TEXT"), ("tam_clientes", "REAL"), ("sam_desc", "TEXT"), ("sam_clientes", "REAL"),
        ("som_pct", "REAL"), ("gasto_anual_cliente", "REAL"), ("mercado_fontes", "TEXT"),
        # Marketing (4Ps) e localização
        ("mkt_produto", "TEXT"), ("mkt_preco", "TEXT"), ("mkt_praca", "TEXT"), ("mkt_promocao", "TEXT"),
        ("loc_endereco", "TEXT"), ("loc_justificativa", "TEXT"), ("loc_area_m2", "REAL"),
        ("loc_fluxo", "TEXT"),
        # Plano operacional
        ("op_layout", "TEXT"), ("op_processos", "TEXT"), ("op_capacidade_txt", "TEXT"),
        ("op_horario", "TEXT"),
        # Custos e preço
        ("rateio_criterio", "TEXT"), ("margem_desejada_pct", "REAL"),
        # Textos finais
        ("sumario_executivo", "TEXT"), ("conclusao", "TEXT"),
        # Serviços e assinaturas (v6)
        ("tipo_negocio", "TEXT"),                  # Produtos | Serviços | Assinatura | Misto
        ("horas_mes_pessoa", "REAL"),              # horas contratadas por pessoa da operação
        ("utilizacao_pct", "REAL"),                # % dessas horas que viram trabalho cobrável
        ("anexo_simples", "TEXT"),
        ("verba_aquisicao_mensal", "REAL"),         # parte do gasto comercial dedicada a conquistar assinantes
        # Identidade e estratégia (v7)
        ("missao", "TEXT"), ("visao", "TEXT"), ("valores", "TEXT"),
        ("canvas_relacionamento", "TEXT"), ("canvas_parcerias", "TEXT"),
        ("bcg_limite_crescimento_pct", "REAL"),
        # Reforma tributária (simulação)
        ("reforma_b2b_pct", "REAL"), ("reforma_cbs_pct", "REAL"), ("reforma_ibs_pct", "REAL"),
        ("reforma_ano", "REAL"), ("reforma_compras_credito", "REAL"),
        # Validação, marketing e incerteza (v7.1)
        ("intencao_certamente_pct", "REAL"), ("intencao_provavel_pct", "REAL"),
        ("fator_certamente_pct", "REAL"), ("fator_provavel_pct", "REAL"), ("vw_produto", "TEXT"),
        ("vpc_job", "TEXT"), ("vpc_tarefas", "TEXT"), ("vpc_dores", "TEXT"), ("vpc_ganhos", "TEXT"),
        ("vpc_produtos", "TEXT"), ("vpc_aliviadores", "TEXT"), ("vpc_criadores", "TEXT"),
        ("posic_segmento", "TEXT"), ("posic_necessidade", "TEXT"), ("posic_marca", "TEXT"), ("posic_categoria", "TEXT"),
        ("posic_beneficio", "TEXT"), ("posic_concorrente", "TEXT"), ("posic_diferencial", "TEXT"),
        ("mapa_eixo_x", "TEXT"), ("mapa_eixo_y", "TEXT"), ("nps_meta", "REAL"),
        ("bass_mercado", "REAL"), ("bass_p", "REAL"), ("bass_q", "REAL"), ("revenda_ativos_pct", "REAL"),
        # v8: relatório ABNT, dados públicos, Lean Canvas, valuation, impacto e edital
        ("abnt_instituicao", "TEXT"), ("abnt_curso", "TEXT"), ("abnt_disciplina", "TEXT"), ("abnt_orientador", "TEXT"),
        ("abnt_cidade", "TEXT"), ("abnt_ano", "TEXT"), ("municipio_ibge", "TEXT"), ("municipio_nome", "TEXT"),
        ("uf_sigla", "TEXT"), ("setor_referencia", "TEXT"), ("lean_vantagem", "TEXT"), ("lean_alternativas", "TEXT"),
        ("lean_metricas", "TEXT"), ("val_crescimento_perp_pct", "REAL"), ("val_multiplo_ebitda", "REAL"),
        ("val_multiplo_receita", "REAL"), ("rodada_investimento", "REAL"), ("rodada_pre_money", "REAL"),
        ("imp_problema", "TEXT"), ("imp_publico", "TEXT"), ("imp_atividades", "TEXT"), ("imp_produtos", "TEXT"),
        ("imp_resultados", "TEXT"), ("imp_impacto", "TEXT"), ("ods", "TEXT"), ("edital_trl", "REAL"),
        # v8.1: mix de marketing integrado (Produto e Preço)
        ("mix_nucleo", "TEXT"), ("mix_real", "TEXT"), ("mix_ampliado", "TEXT"), ("mix_ciclo", "TEXT"),
        ("mix_estrategia_preco", "TEXT"), ("mix_posicao_preco", "TEXT"),
        # v8.2: Praça e Promoção
        ("mix_distribuicao", "TEXT"), ("mix_objetivo_com", "TEXT"), ("mix_ticket", "REAL"), ("mix_freq", "REAL"),
    ],
    "socios": [("nome", "TEXT"), ("funcao", "TEXT"), ("contato", "TEXT"),
               ("perfil", "TEXT"), ("capital_integralizado", "REAL"), ("perda_aceitavel", "REAL")],
    "segmentos": [("nome", "TEXT"), ("descricao", "TEXT"), ("nota_tamanho", "REAL"),
                  ("nota_lucro", "REAL"), ("nota_longevidade", "REAL"),
                  ("nota_concorrencia", "REAL"), ("nota_aderencia", "REAL")],
    "proposta_valor": [("segmento_alvo", "TEXT"), ("beneficios_desejados", "TEXT"),
                       ("meios_entrega", "TEXT"), ("diferencial_competitivo", "TEXT"),
                       ("entrega_produtos_servicos", "TEXT")],
    "personas": [
        ("nome", "TEXT"), ("idade", "INTEGER"), ("profissao", "TEXT"), ("estado_civil", "TEXT"),
        ("renda_mensal", "REAL"), ("escolaridade", "TEXT"), ("cidade", "TEXT"),
        ("descricao", "TEXT"), ("citacao", "TEXT"), ("habitos", "TEXT"),
        ("caracteristicas", "TEXT"), ("fluencia_digital", "TEXT"), ("redes_sociais", "TEXT"),
        ("quem_influencia", "TEXT"), ("rotina_manha", "TEXT"), ("rotina_tarde", "TEXT"),
        ("rotina_noite", "TEXT"), ("ganhos", "TEXT"), ("desafios", "TEXT"),
        ("duvidas_objecoes", "TEXT"), ("aborrecimentos", "TEXT"), ("experiencia_ideal", "TEXT"),
    ],
    "jornada_cliente": [("persona_nome", "TEXT"), ("macro_fase", "TEXT"), ("etapa", "TEXT"),
                        ("acoes_cliente", "TEXT"), ("pensamentos", "TEXT"),
                        ("pontos_contato", "TEXT"), ("estado_emocional", "TEXT"),
                        ("oportunidades", "TEXT")],
    "concorrentes": [("nome", "TEXT"), ("observacoes", "TEXT")],
    "concorrencia_notas": [("atributo", "TEXT"), ("empresa", "TEXT"), ("nota", "REAL")],
    "swot": [("ambiente", "TEXT"), ("dimensao", "TEXT"), ("fator", "TEXT"),
             ("natureza", "TEXT"), ("impacto", "TEXT"), ("diagnostico", "TEXT"),
             ("plano_acao", "TEXT"), ("priorizado", "INTEGER")],
    "canais": [("canal", "TEXT"), ("atratividade", "REAL"), ("alcance", "REAL"),
               ("cliente", "REAL"), ("relacionamento", "REAL"), ("viabilidade", "REAL")],
    "funil_vendas": [("etapa", "TEXT"), ("acao", "TEXT"), ("orcamento", "REAL"),
                     ("alcancados", "REAL"), ("convertidos", "REAL")],
    "inv_fixos": [("descricao", "TEXT"), ("quantidade", "REAL"), ("valor_unitario", "REAL"),
                  ("vida_util_anos", "REAL"), ("valor_residual_pct", "REAL")],
    "inv_pre": [("descricao", "TEXT"), ("valor", "REAL")],
    "emprestimos": [("descricao", "TEXT"), ("valor", "REAL"), ("taxa_anual_pct", "REAL"),
                    ("prazo_meses", "REAL"), ("carencia_meses", "REAL"), ("sistema", "TEXT")],
    "capital_giro_prazos": [("tipo", "TEXT"), ("prazo_dias", "REAL"), ("percentual", "REAL")],
    "materiais": [("nome", "TEXT"), ("categoria", "TEXT"), ("unidade", "TEXT"), ("valor_unitario", "REAL"),
                  ("qtd_estoque_inicial", "REAL")],
    "catalogo": [("categoria", "TEXT"), ("descricao", "TEXT"), ("preco_venda", "REAL"),
                 ("custo_aquisicao", "REAL"), ("qtd_mensal", "REAL"),
                 ("qtd_estoque_inicial", "REAL"), ("horas_unit", "REAL")],
    "ficha_tecnica": [("produto", "TEXT"), ("material", "TEXT"), ("quantidade", "REAL")],
    "custos_variaveis": [("descricao", "TEXT"), ("percentual", "REAL")],
    "pessoal": [("cargo", "TEXT"), ("area", "TEXT"), ("num_empregados", "REAL"), ("salario_mensal", "REAL"),
                ("encargos_pct", "REAL"), ("beneficios_unit", "REAL"), ("atribuicoes", "TEXT"),
                ("mes_inicio", "REAL")],                 # mês do plano em que o cargo começa (contratações futuras)
    "custos_fixos": [("descricao", "TEXT"), ("classificacao", "TEXT"), ("valor_mensal", "REAL")],
    # --- novas tabelas (v5) ---
    "propostas": [("persona_nome", "TEXT"), ("valores", "TEXT"), ("que_desejam", "TEXT"),
                  ("por_meio", "TEXT"), ("diferentemente", "TEXT"), ("oferecem", "TEXT")],
    "capacidade": [("recurso", "TEXT"), ("produto_ou_grupo", "TEXT"), ("capacidade_mensal", "REAL"),
                   ("observacao", "TEXT")],
    "fornecedores": [("nome", "TEXT"), ("itens", "TEXT"), ("prazo_pagamento_dias", "REAL"),
                     ("prazo_entrega_dias", "REAL"), ("condicoes", "TEXT"), ("critico", "INTEGER")],
    "riscos": [("risco", "TEXT"), ("probabilidade", "REAL"), ("impacto", "REAL"), ("mitigacao", "TEXT"),
               ("responsavel", "TEXT")],
    "experimentos": [("hipotese", "TEXT"), ("experimento", "TEXT"), ("metrica", "TEXT"),
                     ("criterio_sucesso", "TEXT"), ("resultado", "TEXT"), ("status", "TEXT")],
    "cronograma": [("atividade", "TEXT"), ("mes_inicio", "REAL"), ("duracao_meses", "REAL"),
                   ("responsavel", "TEXT"), ("custo", "REAL")],
    "realizado": [("mes", "REAL"), ("receita_real", "REAL"), ("custos_variaveis_real", "REAL"),
                  ("custos_fixos_real", "REAL")],
    "listas": [("tipo", "TEXT"), ("valor", "TEXT")],
    # --- v6: receita recorrente (planos de assinatura) ---
    "assinaturas": [("plano", "TEXT"), ("publico", "TEXT"), ("preco_ciclo", "REAL"), ("ciclo_meses", "REAL"),
                    ("novos_mes1", "REAL"), ("crescimento_novos_pct", "REAL"), ("churn_mensal_pct", "REAL"),
                    ("custo_entrega_mensal", "REAL"), ("taxa_cobranca_pct", "REAL")],
    "versoes": [("nome", "TEXT"), ("data", "TEXT"), ("dados", "TEXT")],
    # --- v7: objetivos SMART, matrizes estratégicas ---
    "objetivos": [("objetivo", "TEXT"), ("perspectiva", "TEXT"), ("indicador", "TEXT"), ("meta", "REAL"),
                  ("unidade", "TEXT"), ("prazo_meses", "REAL"), ("responsavel", "TEXT"), ("por_que_atingivel", "TEXT"),
                  ("relevancia", "TEXT"), ("iniciativa", "TEXT"), ("contribui_para", "TEXT")],
    "bcg": [("produto", "TEXT"), ("crescimento_mercado_pct", "REAL"), ("participacao_relativa", "REAL")],
    "ge": [("unidade", "TEXT"), ("forca", "REAL"), ("justificativa", "TEXT")],
    "porter": [("forca", "TEXT"), ("intensidade", "REAL"), ("justificativa", "TEXT"), ("estrategia", "TEXT")],
    # --- v7.1: validação, marketing, incerteza e defesa ---
    "evidencias": [("premissa", "TEXT"), ("status", "TEXT"), ("fonte", "TEXT"), ("observacao", "TEXT")],
    "entrevistas": [("data", "TEXT"), ("entrevistado", "TEXT"), ("perfil", "TEXT"), ("dores", "TEXT"),
                    ("solucao_atual", "TEXT"), ("pagaria", "TEXT"), ("valor_pagaria", "REAL"), ("aprendizado", "TEXT")],
    "precos_vw": [("respondente", "TEXT"), ("muito_barato", "REAL"), ("barato", "REAL"), ("caro", "REAL"),
                  ("muito_caro", "REAL")],
    "vpc_encaixe": [("item", "TEXT"), ("tipo", "TEXT"), ("resposta", "TEXT")],
    "errc": [("atributo", "TEXT"), ("acao", "TEXT"), ("nova_nota", "REAL"), ("justificativa", "TEXT")],
    "mapa_perceptual": [("marca", "TEXT"), ("eixo_x", "REAL"), ("eixo_y", "REAL")],
    "orcamento_canais": [("canal", "TEXT"), ("gasto_mensal", "REAL"), ("clientes_mes", "REAL"), ("observacao", "TEXT")],
    "aarrr": [("etapa", "TEXT"), ("metrica", "TEXT"), ("meta", "TEXT"), ("acao", "TEXT")],
    "calendario_mkt": [("mes", "REAL"), ("acao", "TEXT"), ("canal", "TEXT"), ("custo", "REAL")],
    "mc_faixas": [("variavel", "TEXT"), ("minimo_pct", "REAL"), ("maximo_pct", "REAL")],
    "formalizacao": [("item", "TEXT"), ("status", "TEXT"), ("observacao", "TEXT")],
    "banca_respostas": [("pergunta", "TEXT"), ("resposta", "TEXT")],
    # --- v8 ---
    "referencias": [("referencia", "TEXT")],
    "swot_estrategias": [("quadrante", "TEXT"), ("estrategia", "TEXT"), ("acao", "TEXT"), ("prazo_meses", "REAL")],
    "impacto_indicadores": [("indicador", "TEXT"), ("tipo", "TEXT"), ("linha_base", "TEXT"), ("meta", "TEXT"), ("ods", "TEXT")],
    "edital_textos": [("secao", "TEXT"), ("texto", "TEXT"), ("limite", "REAL")],
    "edital_orcamento": [("rubrica", "TEXT"), ("item", "TEXT"), ("valor", "REAL")],
    "comentarios": [("pagina", "TEXT"), ("comentario", "TEXT"), ("autor", "TEXT"), ("data", "TEXT"), ("resolvido", "INTEGER")],
    # --- v8.1: mix de marketing integrado ---
    "precos_concorrentes": [("concorrente", "TEXT"), ("produto", "TEXT"), ("preco", "REAL"), ("fonte", "TEXT")],
    "meios_pagamento": [("forma", "TEXT"), ("pct_vendas", "REAL"), ("taxa_pct", "REAL"), ("prazo_dias", "REAL")],
    "canais_venda": [("canal", "TEXT"), ("pct_vendas", "REAL"), ("comissao_pct", "REAL"), ("observacao", "TEXT")],
    "dados_publicos": [("chave", "TEXT"), ("rotulo", "TEXT"), ("valor", "REAL"), ("periodo", "TEXT"), ("fonte", "TEXT")],
}

# Valores-padrão de um projeto em branco
PROJETO_PADRAO = {
    "nome_negocio": "Novo Projeto", "objetivo": "Criar um novo negócio",
    "enquadramento": "Simples Nacional", "periodo_anos": 3, "porte": "Microempresa (ME)",
    "setor_atuacao": "", "cidade_uf": "", "contatos": "",
    "aliquota_imposto_pct": 6.0, "crescimento_mensal_ano1_pct": 1.0,
    "crescimento_anual_pct": 8.0, "inflacao_custos_anual_pct": 4.5, "tma_anual_pct": 12.0,
    "pme_dias": 15.0, "reserva_caixa_extra": 0.0, "recuperar_giro": 1,
    "cen_otim_receita_pct": 10.0, "cen_otim_custo_pct": 5.0,
    "cen_pess_receita_pct": 15.0, "cen_pess_custo_pct": 5.0,
    "rateio_criterio": "Custo direto", "margem_desejada_pct": 15.0,
    "tipo_negocio": "Produtos", "horas_mes_pessoa": 168.0, "utilizacao_pct": 75.0, "anexo_simples": "",
    "verba_aquisicao_mensal": 0.0, "bcg_limite_crescimento_pct": 10.0,
    "reforma_b2b_pct": 0.0, "reforma_cbs_pct": 11.3, "reforma_ibs_pct": 0.1, "reforma_ano": 2027.0,
    "reforma_compras_credito": -1.0,
    "intencao_certamente_pct": 0.0, "intencao_provavel_pct": 0.0, "fator_certamente_pct": 70.0, "fator_provavel_pct": 30.0,
    "nps_meta": 50.0, "bass_mercado": 0.0, "bass_p": 0.03, "bass_q": 0.38, "revenda_ativos_pct": 40.0,
    "mapa_eixo_x": "Preço (baixo → alto)", "mapa_eixo_y": "Conveniência (baixa → alta)",
    "val_crescimento_perp_pct": 3.0, "val_multiplo_ebitda": 4.0, "val_multiplo_receita": 1.0,
    "rodada_investimento": 0.0, "rodada_pre_money": 0.0, "edital_trl": 3.0,
}

# Custos variáveis sugeridos para um projeto em branco (percentuais zerados)
CUSTOS_VARIAVEIS_PADRAO = [
    "Comissões de vendas", "Taxas de cartões", "Propaganda variável",
    "Previsão de inadimplência", "Comissões de aplicativos de delivery",
    "Descontos e cupons promocionais", "Frete subsidiado em entregas",
]


# ------------------------------------------------------------------------------
# 2. CONEXÃO E CRIAÇÃO / ATUALIZAÇÃO DA ESTRUTURA
# ------------------------------------------------------------------------------
def conectar(caminho: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(caminho), check_same_thread=False)
    garantir_estrutura(conn)
    return conn


def _colunas(conn, tabela):
    return [r[1] for r in conn.execute(f"PRAGMA table_info({tabela})")]


def _existe(conn, tabela):
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tabela,)
    ).fetchone() is not None


def garantir_estrutura(conn: sqlite3.Connection) -> None:
    """Cria as tabelas que faltam e acrescenta colunas novas em arquivos antigos."""
    for tabela, cols in SCHEMA.items():
        pk = "id INTEGER PRIMARY KEY CHECK (id = 1)" if tabela in TABELAS_UNICAS \
            else "id INTEGER PRIMARY KEY AUTOINCREMENT"
        defs = ", ".join(f"{c} {t}" for c, t in cols)
        conn.execute(f"CREATE TABLE IF NOT EXISTS {tabela} ({pk}, {defs})")
        existentes = _colunas(conn, tabela)
        for c, t in cols:
            if c not in existentes:
                conn.execute(f"ALTER TABLE {tabela} ADD COLUMN {c} {t}")

    # Linha única do projeto
    if conn.execute("SELECT COUNT(*) FROM projeto").fetchone()[0] == 0:
        cols = list(PROJETO_PADRAO)
        conn.execute(
            f"INSERT INTO projeto (id, {', '.join(cols)}) VALUES (1, {', '.join('?' * len(cols))})",
            [PROJETO_PADRAO[c] for c in cols],
        )
    else:  # preenche campos novos que vieram vazios de versões anteriores
        for c, v in PROJETO_PADRAO.items():
            conn.execute(f"UPDATE projeto SET {c}=? WHERE id=1 AND {c} IS NULL", (v,))
    if conn.execute("SELECT COUNT(*) FROM proposta_valor").fetchone()[0] == 0:
        conn.execute("INSERT INTO proposta_valor (id) VALUES (1)")

    _migrar_versao_anterior(conn)
    conn.commit()


def _migrar_versao_anterior(conn):
    """Converte dados de arquivos .pnbox gerados por versões anteriores do programa."""
    conn.execute("UPDATE custos_fixos SET classificacao='Despesa administrativa' "
                 "WHERE classificacao IS NULL OR classificacao=''")
    conn.execute("UPDATE pessoal SET area='Produção / operação' WHERE area IS NULL OR area=''")
    conn.execute("UPDATE catalogo SET horas_unit=0 WHERE horas_unit IS NULL")
    conn.execute("UPDATE pessoal SET mes_inicio=1 WHERE mes_inicio IS NULL OR mes_inicio<1")
    conn.execute("UPDATE emprestimos SET carencia_meses=0 WHERE carencia_meses IS NULL")
    conn.execute("UPDATE emprestimos SET sistema='Price' WHERE sistema IS NULL OR sistema=''")
    conn.execute("UPDATE assinaturas SET ciclo_meses=1 WHERE ciclo_meses IS NULL OR ciclo_meses<1")
    # Proposta de valor única (v1-v4) -> lista de propostas
    if conn.execute("SELECT COUNT(*) FROM propostas").fetchone()[0] == 0:
        pv = conn.execute("SELECT segmento_alvo, beneficios_desejados, meios_entrega, diferencial_competitivo, "
                          "entrega_produtos_servicos FROM proposta_valor WHERE id=1").fetchone()
        if pv and any(pv):
            conn.execute("INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, "
                         "oferecem) VALUES (?,?,?,?,?,?)", (pv[0], "", pv[1], pv[2], pv[3], pv[4]))
    # Catálogo: custo_unitario -> custo_aquisicao
    if "custo_unitario" in _colunas(conn, "catalogo"):
        conn.execute("UPDATE catalogo SET custo_aquisicao = custo_unitario "
                     "WHERE custo_aquisicao IS NULL")
    # Investimentos em tabela única -> tabelas separadas
    if _existe(conn, "investimentos") and \
            conn.execute("SELECT COUNT(*) FROM inv_fixos").fetchone()[0] == 0 and \
            conn.execute("SELECT COUNT(*) FROM inv_pre").fetchone()[0] == 0:
        for tipo, desc, qtd, vu, vida in conn.execute(
                "SELECT tipo, descricao, quantidade, valor_unitario, vida_util_anos FROM investimentos"):
            if tipo == "Fixo":
                conn.execute("INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, "
                             "vida_util_anos, valor_residual_pct) VALUES (?,?,?,?,0)",
                             (desc, qtd, vu, vida))
            elif tipo == "Pré-Operacional":
                conn.execute("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)",
                             (desc, (qtd or 0) * (vu or 0)))
            else:
                conn.execute("INSERT INTO materiais (nome, unidade, valor_unitario, "
                             "qtd_estoque_inicial) VALUES (?, 'unidade', ?, ?)", (desc, vu, qtd))
    # Dívidas da v2 -> tabela de empréstimos
    cols_proj = _colunas(conn, "projeto")
    if "divida_cp_inicial" in cols_proj and \
            conn.execute("SELECT COUNT(*) FROM emprestimos").fetchone()[0] == 0:
        cp, lp, tx = conn.execute(
            "SELECT divida_cp_inicial, divida_lp_inicial, taxa_juros_divida_pct FROM projeto").fetchone()
        if cp:
            conn.execute("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) "
                         "VALUES ('Empréstimo de curto prazo', ?, ?, 12)", (cp, tx or 0))
        if lp:
            conn.execute("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) "
                         "VALUES ('Empréstimo de longo prazo', ?, ?, 48)", (lp, tx or 0))
    # Concorrência em formato largo -> formato longo
    if "concorrente_1" in _colunas(conn, "concorrencia") if _existe(conn, "concorrencia") else False:
        if conn.execute("SELECT COUNT(*) FROM concorrencia_notas").fetchone()[0] == 0:
            nomes = {"nossa_empresa": "Nossa Empresa", "concorrente_1": "Concorrente 1",
                     "concorrente_2": "Concorrente 2", "concorrente_3": "Concorrente 3"}
            for n in list(nomes.values())[1:]:
                conn.execute("INSERT INTO concorrentes (nome) VALUES (?)", (n,))
            for row in conn.execute("SELECT atributo, nossa_empresa, concorrente_1, "
                                    "concorrente_2, concorrente_3 FROM concorrencia").fetchall():
                for (col, nome), nota in zip(nomes.items(), row[1:]):
                    conn.execute("INSERT INTO concorrencia_notas (atributo, empresa, nota) "
                                 "VALUES (?,?,?)", (row[0], nome, nota))


# ------------------------------------------------------------------------------
# 3. LEITURA E GRAVAÇÃO
# ------------------------------------------------------------------------------
def colunas_de(tabela):
    return [c for c, _ in SCHEMA[tabela]]


def ler(conn, tabela) -> pd.DataFrame:
    cols = colunas_de(tabela)
    df = pd.read_sql(f"SELECT {', '.join(cols)} FROM {tabela} ORDER BY id", conn)
    return limpar(df, tabela)


def limpar(df: pd.DataFrame, tabela) -> pd.DataFrame:
    """Garante colunas, troca vazios numéricos por 0 e textos vazios por ''."""
    df = df.copy()
    for c, t in SCHEMA[tabela]:
        if c not in df.columns:
            df[c] = None
        if t in ("REAL", "INTEGER"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
            if c in ("priorizado", "critico"):
                df[c] = df[c].astype(bool)
        else:
            df[c] = df[c].astype(object).where(df[c].notna(), "")
    return df[colunas_de(tabela)].reset_index(drop=True)


def salvar(conn, tabela, df: pd.DataFrame, chave: str | None = None) -> None:
    """Substitui o conteúdo da tabela pelo DataFrame editado."""
    df = df.copy()
    if chave:
        df = df[df[chave].astype(str).str.strip() != ""]
        df = df[df[chave].notna()]
    df = limpar(df, tabela)
    conn.execute(f"DELETE FROM {tabela}")
    if not df.empty:
        df.to_sql(tabela, conn, if_exists="append", index=False)
    conn.commit()


def ler_unica(conn, tabela) -> dict:
    cols = colunas_de(tabela)
    row = conn.execute(f"SELECT {', '.join(cols)} FROM {tabela} WHERE id=1").fetchone()
    d = dict(zip(cols, row))
    for c, t in SCHEMA[tabela]:
        if t in ("REAL", "INTEGER"):
            d[c] = float(d[c] or 0)
        else:
            d[c] = d[c] or ""
    return d


def salvar_unica(conn, tabela, dados: dict) -> None:
    cols = [c for c in colunas_de(tabela) if c in dados]
    conn.execute(f"UPDATE {tabela} SET {', '.join(c + '=?' for c in cols)} WHERE id=1",
                 [dados[c] for c in cols])
    conn.commit()


# ------------------------------------------------------------------------------
# 4. CRIAÇÃO DE PROJETOS
# ------------------------------------------------------------------------------
def criar_projeto_em_branco(caminho: Path, nome: str) -> None:
    conn = conectar(caminho)
    salvar_unica(conn, "projeto", {"nome_negocio": nome})
    conn.executemany("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?, 0)",
                     [(d,) for d in CUSTOS_VARIAVEIS_PADRAO])
    conn.executemany("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)",
                     [("Recebimento", 0, 100.0), ("Pagamento", 0, 100.0)])
    conn.commit()
    conn.close()


def criar_projeto_modelo(caminho: Path) -> None:
    """Cria o estudo de caso FitBox (inspirado no relatório PNBOX do Sebrae)."""
    conn = conectar(caminho)
    popular_modelo_fitbox(conn)
    popular_estrategia(conn, "fitbox")
    conn.commit()
    conn.close()


def criar_projeto_cafe(caminho: Path) -> None:
    """Cria o segundo estudo de caso: Café Mucuri (cafeteria e padaria artesanal)."""
    conn = conectar(caminho)
    popular_modelo_cafe(conn)
    popular_estrategia(conn, "cafe")
    conn.commit()
    conn.close()


def criar_projeto_servico(caminho: Path) -> None:
    """Cria o quarto estudo de caso: TecVale Soluções (serviços de TI por hora, por projeto e por contrato)."""
    conn = conectar(caminho)
    popular_modelo_servico(conn)
    popular_estrategia(conn, "servico")
    conn.commit()
    conn.close()


def criar_projeto_assinatura(caminho: Path) -> None:
    """Cria o terceiro estudo de caso: Conecta Vale (aplicativo com planos de assinatura)."""
    conn = conectar(caminho)
    popular_modelo_assinatura(conn)
    popular_estrategia(conn, "assinatura")
    conn.commit()
    conn.close()
