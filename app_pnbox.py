"""
================================================================================
PNBOX ACADÊMICO v5 - PLANO DE NEGÓCIO E VIABILIDADE ECONÔMICO-FINANCEIRA
================================================================================
Programa único (um só arquivo). Formas de abrir:
    - dois cliques no arquivo "Abrir PNBOX.bat" (na mesma pasta); ou
    - dois cliques neste arquivo .py (se o Windows abrir .py com o Python); ou
    - no terminal:  py app_pnbox.py
Na primeira vez ele instala sozinho as bibliotecas que faltarem
(streamlit, pandas, openpyxl, fpdf2, matplotlib) e já abre o navegador,
sem pedir e-mail, numa porta livre escolhida na hora.
Ao fechar a aba do navegador, o programa se encerra sozinho (após ~15 s).

Os projetos ficam na pasta "projetos", criada ao lado deste arquivo.
Partes do código: (1) banco de dados, (2) modelo FitBox, (3) motor financeiro,
(4) relatório PDF, (5) extras (IA, pitch, Word, professor), (6) interface.
================================================================================
"""
from __future__ import annotations

# ==============================================================================
# INICIALIZAÇÃO AUTOMÁTICA (executada antes de tudo)
# ==============================================================================
import importlib.util as _ilu
import subprocess as _sp
import sys as _sys
from pathlib import Path as _Path


def _no_streamlit() -> bool:
    """True quando o arquivo já está rodando dentro do Streamlit."""
    try:
        from streamlit.runtime import exists
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return exists() or get_script_run_ctx(suppress_warning=True) is not None
    except Exception:
        return False


def _preparar_ambiente() -> None:
    pacotes = {"streamlit": "streamlit", "pandas": "pandas", "openpyxl": "openpyxl",
               "matplotlib": "matplotlib", "altair": "altair", "docx": "python-docx"}
    faltando = [pip for mod, pip in pacotes.items() if _ilu.find_spec(mod) is None]
    try:
        import fpdf
        if int(fpdf.__version__.split(".")[0]) < 2 or tuple(map(int, fpdf.__version__.split(".")[:2])) < (2, 7):
            faltando.append("fpdf2")
    except Exception:
        faltando.append("fpdf2")
    if faltando:
        print("Instalando bibliotecas necessárias:", ", ".join(faltando))
        _sp.check_call([_sys.executable, "-m", "pip", "install", "-U", *faltando])
    # Evita a pergunta "Email:" do Streamlit na primeira execução
    cred = _Path.home() / ".streamlit" / "credentials.toml"
    if not cred.exists():
        cred.parent.mkdir(parents=True, exist_ok=True)
        cred.write_text('[general]\nemail = ""\n', encoding="utf-8")


if __name__ == "__main__" and not _no_streamlit():
    _preparar_ambiente()
    import os as _os
    import socket as _socket
    # Porta livre escolhida pelo sistema (evita conflito com outro programa)
    with _socket.socket() as _s:
        _s.bind(("127.0.0.1", 0))
        _porta = _s.getsockname()[1]
    _os.environ["PNBOX_AUTOFECHAR"] = "1"
    print(f"Abrindo o PNBOX em http://localhost:{_porta}  (feche a aba do navegador para encerrar)")
    from streamlit.web import cli as _stcli
    _sys.argv = ["streamlit", "run", str(_Path(__file__).resolve()),
                 f"--server.port={_porta}", "--server.address=localhost",
                 "--browser.gatherUsageStats=false", "--server.headless=false"]
    _sys.exit(_stcli.main())
from datetime import date
from datetime import datetime
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace
from pathlib import Path
import altair as alt
import hashlib
import io
import json
import math
import os
import pandas as pd
import re
import shutil
import sqlite3
import streamlit as st
import sys
import tempfile
import threading
import time


# ##############################################################################
# PARTE: pnbox_db.py
# ##############################################################################



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
    ],
    "socios": [("nome", "TEXT"), ("funcao", "TEXT"), ("contato", "TEXT"),
               ("perfil", "TEXT"), ("capital_integralizado", "REAL")],
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
                    ("prazo_meses", "REAL")],
    "capital_giro_prazos": [("tipo", "TEXT"), ("prazo_dias", "REAL"), ("percentual", "REAL")],
    "materiais": [("nome", "TEXT"), ("categoria", "TEXT"), ("unidade", "TEXT"), ("valor_unitario", "REAL"),
                  ("qtd_estoque_inicial", "REAL")],
    "catalogo": [("categoria", "TEXT"), ("descricao", "TEXT"), ("preco_venda", "REAL"),
                 ("custo_aquisicao", "REAL"), ("qtd_mensal", "REAL"),
                 ("qtd_estoque_inicial", "REAL")],
    "ficha_tecnica": [("produto", "TEXT"), ("material", "TEXT"), ("quantidade", "REAL")],
    "custos_variaveis": [("descricao", "TEXT"), ("percentual", "REAL")],
    "pessoal": [("cargo", "TEXT"), ("area", "TEXT"), ("num_empregados", "REAL"), ("salario_mensal", "REAL"),
                ("encargos_pct", "REAL"), ("beneficios_unit", "REAL"), ("atribuicoes", "TEXT")],
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
    "versoes": [("nome", "TEXT"), ("data", "TEXT"), ("dados", "TEXT")],
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
    conn.commit()
    conn.close()


def criar_projeto_cafe(caminho: Path) -> None:
    """Cria o segundo estudo de caso: Café Mucuri (cafeteria e padaria artesanal)."""
    conn = conectar(caminho)
    popular_modelo_cafe(conn)
    conn.commit()
    conn.close()

# ##############################################################################
# PARTE: pnbox_modelo.py
# ##############################################################################


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

# ##############################################################################
# PARTE: pnbox_motor.py
# ##############################################################################



CATEGORIAS = ["Fabricação Própria", "Revenda", "Serviço"]


# ------------------------------------------------------------------------------
# Formatação no padrão brasileiro
# ------------------------------------------------------------------------------
def brl(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def num_br(v, casas=2) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, casas=2) -> str:
    return "n.a." if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))) \
        else f"{num_br(v * 100, casas)}%"


def div(a, b):
    return a / b if b not in (0, 0.0) and b is not None else float("nan")


# ------------------------------------------------------------------------------
# 1. CUSTO UNITÁRIO DOS PRODUTOS (FICHA TÉCNICA x CUSTO INFORMADO)
# ------------------------------------------------------------------------------
def custos_unitarios(cat: pd.DataFrame, mat: pd.DataFrame, ficha: pd.DataFrame) -> pd.DataFrame:
    precos = dict(zip(mat["nome"], mat["valor_unitario"]))
    f = ficha.copy()
    f["valor_unitario"] = f["material"].map(precos).fillna(0.0)
    f["custo"] = f["quantidade"] * f["valor_unitario"]
    custo_ficha = f.groupby("produto")["custo"].sum().to_dict()

    df = cat.copy()
    custos, origens = [], []
    for _, r in df.iterrows():
        if r["categoria"] != "Revenda" and r["descricao"] in custo_ficha:
            custos.append(custo_ficha[r["descricao"]])
            origens.append("Ficha técnica")
        else:
            custos.append(float(r["custo_aquisicao"]))
            origens.append("Valor de aquisição" if r["categoria"] == "Revenda" else "Custo informado")
    df["custo_unit"] = custos
    df["origem_custo"] = origens
    df["margem_unit"] = df["preco_venda"] - df["custo_unit"]
    df["margem_pct"] = [div(m, p) for m, p in zip(df["margem_unit"], df["preco_venda"])]
    df["markup"] = [div(p, c) for p, c in zip(df["preco_venda"], df["custo_unit"])]
    df["receita_mensal"] = df["preco_venda"] * df["qtd_mensal"]
    df["custo_mensal"] = df["custo_unit"] * df["qtd_mensal"]
    return df


# ------------------------------------------------------------------------------
# 2. CRONOGRAMA DOS EMPRÉSTIMOS (TABELA PRICE, PARCELAS MENSAIS)
# ------------------------------------------------------------------------------
def cronograma_emprestimos(emp: pd.DataFrame, anos: int = 4) -> dict:
    juros = [0.0] * anos
    amort = [0.0] * anos
    saldo_fim = [0.0] * anos
    juros_m = [0.0] * (anos * 12)
    amort_m = [0.0] * (anos * 12)
    for _, r in emp.iterrows():
        pv, n = float(r["valor"]), int(round(r["prazo_meses"]))
        if pv <= 0:
            continue
        n = max(n, 1)
        i = (1 + r["taxa_anual_pct"] / 100.0) ** (1 / 12) - 1
        pmt = pv / n if i == 0 else pv * i / (1 - (1 + i) ** (-n))
        saldo = pv
        for m in range(anos * 12):
            ano = m // 12
            if m < n and saldo > 1e-9:
                j = saldo * i
                a = min(pmt - j, saldo)
                saldo -= a
                juros[ano] += j
                amort[ano] += a
                juros_m[m] += j
                amort_m[m] += a
            if m % 12 == 11:
                saldo_fim[ano] += max(saldo, 0.0)
    return {"juros": juros, "amort": amort, "saldo_fim": saldo_fim, "juros_m": juros_m, "amort_m": amort_m}


# ------------------------------------------------------------------------------
# 3. TIR (bisseção)
# ------------------------------------------------------------------------------
def calcular_tir(fluxos):
    if not fluxos or all(f >= 0 for f in fluxos) or all(f <= 0 for f in fluxos):
        return None
    vpl = lambda r: sum(f / (1 + r) ** t for t, f in enumerate(fluxos))
    lo, hi = -0.99, 10.0
    if vpl(lo) * vpl(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if vpl(lo) * vpl(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


# ------------------------------------------------------------------------------
# 4. MOTOR PRINCIPAL
# ------------------------------------------------------------------------------
def executar_motor(proj: dict, d: dict) -> dict:
    """d = dicionário com os DataFrames do projeto (nomes das tabelas)."""
    cat = custos_unitarios(d["catalogo"], d["materiais"], d["ficha_tecnica"])
    aliq = proj["aliquota_imposto_pct"] / 100.0
    var_pct = float(d["custos_variaveis"]["percentual"].sum()) / 100.0

    # --- Receita e custos diretos do Mês 1, por categoria
    receita_m1 = float(cat["receita_mensal"].sum())
    cdir_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "custo_mensal"].sum()) for c in CATEGORIAS}
    rec_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "receita_mensal"].sum()) for c in CATEGORIAS}
    cdir_m1 = sum(cdir_cat_m1.values())

    # --- Pessoal, custos fixos e depreciação (mensais)
    pes = d["pessoal"]
    sal = pes["num_empregados"] * pes["salario_mensal"]
    salarios_m = float(sal.sum())
    encargos_m = float((sal * pes["encargos_pct"] / 100).sum())
    beneficios_m = float((pes["num_empregados"] * pes["beneficios_unit"]).sum())
    pessoal_m = salarios_m + encargos_m + beneficios_m
    fixos_m = float(d["custos_fixos"]["valor_mensal"].sum())
    pes_total_linha = sal * (1 + pes["encargos_pct"] / 100) + pes["num_empregados"] * pes["beneficios_unit"]
    area = pes["area"].fillna("") if "area" in pes else pd.Series([""] * len(pes))
    mod_m = float(pes_total_linha[area.str.startswith("Produ")].sum())          # mão de obra direta (produção)
    pessoal_adm_m = float(pes_total_linha[area.str.startswith("Admin")].sum())
    pessoal_com_m = pessoal_m - mod_m - pessoal_adm_m
    cfx = d["custos_fixos"]
    cls = cfx["classificacao"].fillna("") if "classificacao" in cfx else pd.Series([""] * len(cfx))
    cif_m = float(cfx.loc[cls.str.startswith("Custo indireto"), "valor_mensal"].sum())
    desp_adm_m = float(cfx.loc[cls.str.contains("administrativa") | (cls == ""), "valor_mensal"].sum())
    desp_com_m = float(cfx.loc[cls.str.contains("comercial"), "valor_mensal"].sum())
    desp_fin_m = float(cfx.loc[cls.str.contains("financeira"), "valor_mensal"].sum())

    inv = d["inv_fixos"].copy()
    inv["total"] = inv["quantidade"] * inv["valor_unitario"]
    inv["dep_anual"] = [
        (t * (1 - res / 100) / v) if v > 0 else 0.0
        for t, res, v in zip(inv["total"], inv["valor_residual_pct"], inv["vida_util_anos"])
    ]
    inv["dep_mensal"] = inv["dep_anual"] / 12
    dep_anual = float(inv["dep_anual"].sum())
    dep_m = dep_anual / 12
    cf_desemb_m = pessoal_m + fixos_m

    # --- Investimentos
    inv_fixo = float(inv["total"].sum())
    inv_pre = float(d["inv_pre"]["valor"].sum())
    mat = d["materiais"]
    est_mat = float((mat["qtd_estoque_inicial"] * mat["valor_unitario"]).sum())
    rev = cat[cat["categoria"] == "Revenda"]
    est_rev = float((rev["qtd_estoque_inicial"] * rev["custo_unit"]).sum())
    inv_estoque = est_mat + est_rev

    # --- Prazos médios e capital de giro
    pz = d["capital_giro_prazos"]
    rec_pz, pag_pz = pz[pz["tipo"] == "Recebimento"], pz[pz["tipo"] == "Pagamento"]
    pmr = float((rec_pz["prazo_dias"] * rec_pz["percentual"] / 100).sum())
    pmp = float((pag_pz["prazo_dias"] * pag_pz["percentual"] / 100).sum())
    pme = float(proj["pme_dias"])
    ciclo_op = pme + pmr
    ciclo_fin = ciclo_op - pmp
    desembolso_m1 = cdir_m1 + receita_m1 * (aliq + var_pct) + cf_desemb_m
    ncg = max(0.0, desembolso_m1 / 30.0 * ciclo_fin)
    reserva = float(proj["reserva_caixa_extra"])
    capital_giro = ncg + reserva
    investimento_total = inv_fixo + inv_pre + inv_estoque + capital_giro

    # --- Fontes
    capital_proprio = float(d["socios"]["capital_integralizado"].sum())
    capital_terceiros = float(d["emprestimos"]["valor"].sum())
    fontes = capital_proprio + capital_terceiros
    crono = cronograma_emprestimos(d["emprestimos"], anos=5)

    # --- Projeção mensal do Ano 1
    g_m = proj["crescimento_mensal_ano1_pct"] / 100
    g_a = proj["crescimento_anual_pct"] / 100
    inf = proj["inflacao_custos_anual_pct"] / 100
    tma = proj["tma_anual_pct"] / 100
    meses = []
    for m in range(12):
        f = (1 + g_m) ** m
        rec = receita_m1 * f
        cv = cdir_m1 * f + rec * (aliq + var_pct)
        cf = cf_desemb_m + dep_m
        meses.append({"Mês": f"Mês {m + 1}", "Receita": rec, "Custos variáveis": cv,
                      "Margem de contribuição": rec - cv, "Custos fixos": cf,
                      "Resultado operacional": rec - cv - cf})
    df_meses = pd.DataFrame(meses)
    fator_a1 = sum((1 + g_m) ** m for m in range(12))

    # --- DRE e balanço dos Anos 1 a 3
    anos = {}
    lucros_acum = 0.0
    for k in (1, 2, 3):
        fr = fator_a1 * (1 + g_a) ** (k - 1)
        fi = (1 + inf) ** (k - 1)
        rec_bruta = receita_m1 * fr
        rec_cat = {c: v * fr for c, v in rec_cat_m1.items()}
        cmv, cpv, csv = (cdir_cat_m1["Revenda"] * fr, cdir_cat_m1["Fabricação Própria"] * fr,
                         cdir_cat_m1["Serviço"] * fr)
        custo_dir = cmv + cpv + csv
        impostos = rec_bruta * aliq
        outras_var = rec_bruta * var_pct
        rec_liquida = rec_bruta - impostos
        cv_tot = custo_dir + impostos + outras_var
        mc = rec_bruta - cv_tot
        imc = div(mc, rec_bruta)
        pessoal = pessoal_m * 12 * fi
        fixos = fixos_m * 12 * fi
        cf_desemb = pessoal + fixos
        ebitda = mc - cf_desemb
        ebit = ebitda - dep_anual
        juros = crono["juros"][k - 1]
        resultado = ebit - juros
        lucros_acum += resultado
        fco = resultado + dep_anual
        amort = crono["amort"][k - 1]
        fluxo_socio = fco - amort

        # Balanço simplificado (caixa calculado por diferença: Ativo = Passivo + PL)
        receber = rec_bruta / 360 * pmr
        estoques = inv_estoque * (rec_bruta / (receita_m1 * fator_a1)) if receita_m1 > 0 else inv_estoque
        imobilizado = max(0.0, inv_fixo - dep_anual * k)
        fornecedores = custo_dir / 360 * pmp
        obrigacoes = (pessoal + impostos) / 12
        saldo_div = crono["saldo_fim"][k - 1]
        div_cp = crono["amort"][k] if k < len(crono["amort"]) else 0.0
        div_cp = min(div_cp, saldo_div)
        div_lp = saldo_div - div_cp
        pc = fornecedores + obrigacoes + div_cp
        pnc = div_lp
        pl = capital_proprio + lucros_acum
        caixa = pc + pnc + pl - (receber + estoques + imobilizado + inv_pre)
        ac = caixa + receber + estoques
        ativo_total = ac + imobilizado + inv_pre

        cf_total = cf_desemb + dep_anual
        pe_cont = div(cf_total, imc) if imc and imc > 0 else float("nan")
        pe_fin = div(cf_desemb, imc) if imc and imc > 0 else float("nan")
        pe_eco = div(cf_total + investimento_total * tma, imc) if imc and imc > 0 else float("nan")

        anos[k] = dict(
            rec_bruta=rec_bruta, rec_cat=rec_cat, impostos=impostos, rec_liquida=rec_liquida,
            cmv=cmv, cpv=cpv, csv=csv, custo_dir=custo_dir, outras_var=outras_var,
            cv_tot=cv_tot, mc=mc, imc=imc, pessoal=pessoal, fixos=fixos, cf_desemb=cf_desemb,
            dep=dep_anual, cf_total=cf_total, ebitda=ebitda, ebit=ebit, juros=juros,
            resultado=resultado, fco=fco, amort=amort, fluxo_socio=fluxo_socio,
            caixa=caixa, receber=receber, estoques=estoques, ac=ac, imobilizado=imobilizado,
            ativo_total=ativo_total, fornecedores=fornecedores, obrigacoes=obrigacoes,
            div_cp=div_cp, div_lp=div_lp, pc=pc, pnc=pnc, pl=pl, divida_bruta=saldo_div,
            divida_liquida=saldo_div - caixa, pe_contabil=pe_cont, pe_financeiro=pe_fin,
            pe_economico=pe_eco,
        )

    # --- Avaliação do projeto (fluxo do projeto = EBITDA, investimento no ano 0)
    fluxos = [-investimento_total] + [anos[k]["ebitda"] for k in (1, 2, 3)]
    if proj.get("recuperar_giro"):
        fluxos[3] += capital_giro
    vpl = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos))
    tir = calcular_tir(fluxos)
    vp_entradas = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos) if t > 0)
    il = div(vp_entradas, investimento_total)

    lucro_mensal_a1 = anos[1]["resultado"] / 12
    payback_simples = div(investimento_total, lucro_mensal_a1) if lucro_mensal_a1 > 0 else float("nan")
    acum, payback_desc = -investimento_total, float("nan")
    for t in (1, 2, 3):
        fd = fluxos[t] / (1 + tma) ** t
        if fd > 0 and acum + fd >= 0:
            payback_desc = ((t - 1) + (-acum) / fd) * 12
            break
        acum += fd

    a1 = anos[1]
    resumo = {
        "receita_anual": a1["rec_bruta"], "custos_variaveis": a1["cv_tot"], "margem_contrib": a1["mc"],
        "custos_fixos": a1["cf_total"], "resultado_operacional": a1["mc"] - a1["cf_total"],
        "lucratividade": div(a1["mc"] - a1["cf_total"], a1["rec_bruta"]),
        "rentabilidade": div(a1["mc"] - a1["cf_total"], investimento_total),
        "payback_meses": div(investimento_total, (a1["mc"] - a1["cf_total"]) / 12)
        if a1["mc"] - a1["cf_total"] > 0 else float("nan"),
    }

    # --- Custeio por absorção: MOD + custos indiretos de produção + depreciação rateados
    overhead_m = mod_m + cif_m + dep_m
    criterio = proj.get("rateio_criterio") or "Custo direto"
    prod_mask = cat["categoria"] != "Revenda"
    base = {"Custo direto": cat["custo_mensal"], "Receita": cat["receita_mensal"],
            "Quantidade": cat["qtd_mensal"]}.get(criterio, cat["custo_mensal"]).where(prod_mask, 0.0)
    tot_base = float(base.sum())
    cat["rateio_mensal"] = base / tot_base * overhead_m if tot_base > 0 else 0.0
    cat["custo_absorcao_unit"] = [cu + (r / q if q > 0 else 0.0) for cu, r, q in
                                  zip(cat["custo_unit"], cat["rateio_mensal"], cat["qtd_mensal"])]
    cat["margem_absorcao_pct"] = [div(p - c, p) for p, c in zip(cat["preco_venda"], cat["custo_absorcao_unit"])]

    # --- Formação de preço (markup divisor)
    margem_desej = float(proj.get("margem_desejada_pct") or 0) / 100
    df_pct = div(cf_desemb_m + dep_m, receita_m1) if receita_m1 > 0 else 0.0
    df_pct = 0.0 if df_pct != df_pct else df_pct
    divisor = 1 - (aliq + var_pct + df_pct + margem_desej)
    markup_mult = div(1, divisor) if divisor > 0 else float("nan")
    cat["preco_sugerido"] = [c * markup_mult if divisor > 0 else float("nan") for c in cat["custo_unit"]]
    cat["dif_preco_pct"] = [div(p - s_, s_) for p, s_ in zip(cat["preco_venda"], cat["preco_sugerido"])]

    # --- Ponto de equilíbrio em unidades por produto (mix atual de vendas)
    pe_m = anos[1]["pe_contabil"] / 12 if anos[1]["pe_contabil"] == anos[1]["pe_contabil"] else float("nan")
    cat["pe_unidades_mes"] = [div(pe_m * div(r, receita_m1), p) if receita_m1 > 0 else float("nan")
                              for r, p in zip(cat["receita_mensal"], cat["preco_venda"])]

    # --- Fluxo de caixa mensal do Ano 1
    def distribuir(valores, prazos_df):
        """Distribui cada valor mensal no tempo conforme os prazos (dias) e percentuais."""
        saida = [0.0] * 14
        regras = [(float(p), float(q) / 100) for p, q in zip(prazos_df["prazo_dias"], prazos_df["percentual"])] \
            or [(0.0, 1.0)]
        for m, v in enumerate(valores):
            for dias, frac in regras:
                lag, resto = int(dias // 30), (dias % 30) / 30
                if m + lag < 14:
                    saida[m + lag] += v * frac * (1 - resto)
                if m + lag + 1 < 14:
                    saida[m + lag + 1] += v * frac * resto
        return saida[:12]

    rec_mes = [receita_m1 * (1 + g_m) ** m for m in range(12)]
    cdir_mes = [cdir_m1 * (1 + g_m) ** m for m in range(12)]
    entradas = distribuir(rec_mes, rec_pz)
    pag_forn = distribuir(cdir_mes, pag_pz)
    saldo = fontes - inv_fixo - inv_pre - inv_estoque
    fluxo_m = []
    for m in range(12):
        sai_var = rec_mes[m] * (aliq + var_pct)
        parcela = crono["juros_m"][m] + crono["amort_m"][m]
        sai = pag_forn[m] + sai_var + cf_desemb_m + parcela
        ini = saldo
        saldo = saldo + entradas[m] - sai
        fluxo_m.append({"Mês": f"Mês {m + 1}", "Saldo inicial": ini, "Recebimentos": entradas[m],
                        "Fornecedores": pag_forn[m], "Impostos e desp. variáveis": sai_var,
                        "Pessoal e custos fixos": cf_desemb_m, "Parcelas de empréstimos": parcela,
                        "Saldo final": saldo})
    df_fluxo_m = pd.DataFrame(fluxo_m)

    return dict(
        mod_m=mod_m, pessoal_adm_m=pessoal_adm_m, pessoal_com_m=pessoal_com_m, cif_m=cif_m,
        desp_adm_m=desp_adm_m, desp_com_m=desp_com_m, desp_fin_m=desp_fin_m, overhead_m=overhead_m,
        criterio_rateio=criterio, margem_desejada=margem_desej, df_pct=df_pct, markup_mult=markup_mult,
        divisor_markup=divisor, fluxo_mensal=df_fluxo_m, g_m=g_m,
        cat=cat, inv=inv, receita_m1=receita_m1, cdir_m1=cdir_m1, var_pct=var_pct, aliq=aliq,
        salarios_m=salarios_m, encargos_m=encargos_m, beneficios_m=beneficios_m, pessoal_m=pessoal_m,
        fixos_m=fixos_m, dep_m=dep_m, dep_anual=dep_anual, inv_fixo=inv_fixo, inv_pre=inv_pre,
        est_mat=est_mat, est_rev=est_rev, inv_estoque=inv_estoque, pmr=pmr, pmp=pmp, pme=pme,
        ciclo_op=ciclo_op, ciclo_fin=ciclo_fin, desembolso_m1=desembolso_m1, ncg=ncg,
        reserva=reserva, capital_giro=capital_giro, investimento_total=investimento_total,
        capital_proprio=capital_proprio, capital_terceiros=capital_terceiros, fontes=fontes,
        crono=crono, meses=df_meses, anos=anos, fluxos=fluxos, vpl=vpl, tir=tir, il=il,
        payback_simples=payback_simples, payback_desc=payback_desc, resumo=resumo, tma=tma,
    )


# ------------------------------------------------------------------------------
# 5. CENÁRIOS (estilo "Simulador de Resultados" do PNBOX)
# ------------------------------------------------------------------------------
def cenarios(res: dict, proj: dict) -> pd.DataFrame:
    a1 = res["anos"][1]
    inv, tma = res["investimento_total"], res["tma"]

    def calc(fr, fc):
        rec, cv, cf = a1["rec_bruta"] * fr, a1["cv_tot"] * fc, a1["cf_total"] * fc
        mc = rec - cv
        r = mc - cf
        imc = div(mc, rec)
        ok = imc == imc and imc > 0
        return [brl(rec), brl(cv), brl(mc), brl(cf), brl(r), pct(div(r, rec)), pct(div(r, inv)),
                f"{num_br(div(inv, r / 12), 1)} meses" if r > 0 else "Não recupera",
                brl(cf / imc) if ok else "n.a.", brl((cf - a1["dep"] * fc) / imc) if ok else "n.a.",
                brl((cf + inv * tma) / imc) if ok else "n.a."]

    rotulos = ["Receita anual", "Custos variáveis", "Margem de contribuição", "Custos fixos",
               "Resultado operacional", "Lucratividade", "Rentabilidade", "Payback simples",
               "PE contábil", "PE financeiro", "PE econômico"]
    ro, co = proj["cen_otim_receita_pct"] / 100, proj["cen_otim_custo_pct"] / 100
    rp, cp = proj["cen_pess_receita_pct"] / 100, proj["cen_pess_custo_pct"] / 100
    return pd.DataFrame({
        "Indicador (Ano 1)": rotulos,
        f"Pessimista (-{num_br(rp*100,0)}% rec. / +{num_br(cp*100,0)}% custos)": calc(1 - rp, 1 + cp),
        "Provável": calc(1, 1),
        f"Otimista (+{num_br(ro*100,0)}% rec. / -{num_br(co*100,0)}% custos)": calc(1 + ro, 1 - co),
    })


# ------------------------------------------------------------------------------
# 6. DRE EM TABELA
# ------------------------------------------------------------------------------
def tabela_dre(res: dict) -> pd.DataFrame:
    A = res["anos"]
    linhas = [
        ("1. Receita operacional bruta", "rec_bruta", 1),
        ("   (-) Impostos sobre vendas", "impostos", -1),
        ("2. (=) Receita líquida", "rec_liquida", 1),
        ("   (-) CMV - mercadorias de revenda", "cmv", -1),
        ("   (-) CPV - produtos de fabricação própria", "cpv", -1),
        ("   (-) CSV - serviços vendidos", "csv", -1),
        ("   (-) Outras despesas variáveis", "outras_var", -1),
        ("3. (=) Margem de contribuição", "mc", 1),
        ("   (-) Pessoal (salários, encargos e benefícios)", "pessoal", -1),
        ("   (-) Demais custos fixos", "fixos", -1),
        ("4. (=) EBITDA", "ebitda", 1),
        ("   (-) Depreciação", "dep", -1),
        ("5. (=) Resultado operacional (EBIT)", "ebit", 1),
        ("   (-) Juros dos empréstimos", "juros", -1),
        ("6. (=) Resultado líquido do exercício", "resultado", 1),
    ]
    rows = []
    for nome, k, s in linhas:
        v = [A[a][k] * s for a in (1, 2, 3)]
        rows.append({"Conta": nome, "Ano 1": brl(v[0]), "Ano 2": brl(v[1]), "Ano 3": brl(v[2]),
                     "% Receita (Ano 1)": pct(div(A[1][k], A[1]["rec_bruta"]))})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------------
# 7. INDICADORES EM 5 BLOCOS
# ------------------------------------------------------------------------------
def tabela_indicadores(res: dict) -> pd.DataFrame:
    A = res["anos"]
    inv = res["investimento_total"]
    linhas = []

    def add(bloco, nome, formula, vals, fmt, leitura, var="pct"):
        f = {"brl": brl, "pct": pct, "x": lambda v: "n.a." if v != v else f"{num_br(v)}x"}[fmt]
        v1, v2 = vals[0], vals[1]
        if var == "pp":
            vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br((v2 - v1) * 100):s} p.p."
        elif var == "dif":
            vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br(v2 - v1)}x"
        else:
            vr = "n.a." if not v1 or v1 != v1 else f"{num_br((v2 - v1) / abs(v1) * 100, 1)}%"
        linhas.append({"Bloco": bloco, "Indicador": nome, "Fórmula": formula,
                       "Ano 1": f(vals[0]), "Ano 2": f(vals[1]), "Ano 3": f(vals[2]),
                       "Var. Ano 2/Ano 1": vr, "Leitura": leitura})

    g = lambda k: [A[a][k] for a in (1, 2, 3)]
    r = lambda n, d_: [div(A[a][n], A[a][d_]) for a in (1, 2, 3)]

    b = "1. Rentabilidade e resultado"
    add(b, "Receita operacional bruta", "Σ preço × quantidade", g("rec_bruta"), "brl", "Faturamento antes das deduções.")
    add(b, "Margem de contribuição", "Receita − custos e despesas variáveis", g("mc"), "brl", "Valor que sobra para pagar os custos fixos.")
    add(b, "Índice de margem de contribuição", "MC ÷ Receita bruta", g("imc"), "pct", "Parcela de cada R$ 1,00 vendido que cobre a estrutura fixa.", "pp")
    add(b, "EBITDA", "MC − custos fixos desembolsáveis", g("ebitda"), "brl", "Geração operacional de caixa antes de juros e depreciação.")
    add(b, "Margem EBITDA", "EBITDA ÷ Receita líquida", r("ebitda", "rec_liquida"), "pct", "Eficiência operacional.", "pp")
    add(b, "Resultado líquido", "EBIT − juros", g("resultado"), "brl", "Lucro (ou prejuízo) do exercício.")
    add(b, "Lucratividade líquida", "Resultado líquido ÷ Receita bruta", r("resultado", "rec_bruta"), "pct", "Ganho para cada real faturado.", "pp")
    add(b, "Rentabilidade (ROI)", "Resultado líquido ÷ Investimento total", [div(A[a]["resultado"], inv) for a in (1, 2, 3)], "pct", "Retorno contábil sobre o capital investido.", "pp")
    add(b, "ROE", "Resultado líquido ÷ Patrimônio líquido", r("resultado", "pl"), "pct", "Retorno sobre o capital dos sócios.", "pp")

    b = "2. Endividamento e estrutura de capital"
    add(b, "Dívida bruta", "Saldo dos empréstimos no fim do ano", g("divida_bruta"), "brl", "Obrigações financeiras onerosas.")
    add(b, "Dívida líquida", "Dívida bruta − caixa", g("divida_liquida"), "brl", "Negativa indica caixa maior que a dívida.")
    add(b, "Endividamento geral", "(PC + PNC) ÷ Ativo total", [div(A[a]["pc"] + A[a]["pnc"], A[a]["ativo_total"]) for a in (1, 2, 3)], "pct", "Parcela do ativo financiada por terceiros.", "pp")
    add(b, "Dívida líquida ÷ EBITDA", "Dívida líquida ÷ EBITDA", [div(A[a]["divida_liquida"], A[a]["ebitda"]) if A[a]["ebitda"] > 0 else float("nan") for a in (1, 2, 3)], "x", "Anos de geração operacional para quitar a dívida.", "dif")
    add(b, "Cobertura de juros", "EBITDA ÷ juros", [div(A[a]["ebitda"], A[a]["juros"]) if A[a]["juros"] > 0 else float("nan") for a in (1, 2, 3)], "x", "Abaixo de 1,0x a operação não paga os juros.", "dif")

    b = "3. Liquidez e capital de giro"
    add(b, "Liquidez corrente", "AC ÷ PC", r("ac", "pc"), "x", "Acima de 1,0x indica folga de curto prazo.", "dif")
    add(b, "Liquidez seca", "(AC − estoques) ÷ PC", [div(A[a]["ac"] - A[a]["estoques"], A[a]["pc"]) for a in (1, 2, 3)], "x", "Capacidade de pagamento sem vender estoques.", "dif")
    add(b, "Liquidez imediata", "Caixa ÷ PC", r("caixa", "pc"), "x", "Disponibilidade imediata.", "dif")
    add(b, "Capital circulante líquido", "AC − PC", [A[a]["ac"] - A[a]["pc"] for a in (1, 2, 3)], "brl", "Folga (ou falta) de recursos de curto prazo.")
    add(b, "Caixa projetado", "Ativo − demais contas (por diferença)", g("caixa"), "brl", "Negativo indica necessidade de mais recursos.")

    b = "4. Viabilidade do investimento"
    add(b, "Custo com pessoal ÷ Receita líquida", "Folha total ÷ Receita líquida", r("pessoal", "rec_liquida"), "pct", "Peso da mão de obra.", "pp")
    add(b, "Fluxo de caixa operacional (FCO)", "Resultado líquido + depreciação", g("fco"), "brl", "Caixa gerado pelas operações.")
    add(b, "Fluxo de caixa do sócio", "FCO − amortização dos empréstimos", g("fluxo_socio"), "brl", "Caixa que sobra após pagar o principal das dívidas.")

    b = "5. Pontos de equilíbrio"
    add(b, "PE contábil (anual)", "Custos fixos totais ÷ IMC", g("pe_contabil"), "brl", "Receita para lucro zero.")
    add(b, "PE financeiro (anual)", "(Custos fixos − depreciação) ÷ IMC", g("pe_financeiro"), "brl", "Receita para cobrir os desembolsos.")
    add(b, "PE econômico (anual)", "(Custos fixos + TMA × investimento) ÷ IMC", g("pe_economico"), "brl", "Receita para remunerar o capital à TMA.")
    add(b, "Margem de segurança", "(Receita − PE contábil) ÷ Receita", [div(A[a]["rec_bruta"] - A[a]["pe_contabil"], A[a]["rec_bruta"]) for a in (1, 2, 3)], "pct", "Quanto a receita pode cair antes do prejuízo.", "pp")
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 8. CHECAGENS DE CONSISTÊNCIA
# ------------------------------------------------------------------------------
def tabela_checagens(res: dict, d: dict) -> pd.DataFrame:
    cat = res["cat"]
    A = res["anos"]
    chk = []

    def add(t, valor, ok, obs):
        chk.append({"Teste": t, "Valor apurado": valor, "Situação": "OK" if ok else "VERIFICAR", "Observação": obs})

    dif = res["fontes"] - res["investimento_total"]
    add("Fontes de recursos x investimento total", f"Diferença: {brl(dif)}", abs(dif) <= 1,
        "Capital dos sócios + empréstimos deve igualar o investimento total.")
    neg = cat[cat["margem_unit"] <= 0]["descricao"].tolist()
    add("Margem unitária positiva", f"{len(neg)} item(ns) com margem <= 0", not neg,
        ("Itens: " + ", ".join(neg[:5])) if neg else "Todos os itens têm preço maior que o custo.")
    sem = cat[cat["custo_unit"] <= 0]["descricao"].tolist()
    add("Produtos com custo cadastrado", f"{len(sem)} item(ns) sem custo", not sem,
        ("Sem ficha técnica/custo: " + ", ".join(sem[:5])) if sem else "Todos os itens têm custo.")
    mats = set(d["materiais"]["nome"])
    orf = sorted(set(d["ficha_tecnica"]["material"]) - mats)
    add("Fichas técnicas com materiais válidos", f"{len(orf)} material(is) não cadastrado(s)", not orf,
        ", ".join(orf[:5]) if orf else "Todos os insumos existem no cadastro de materiais.")
    for tipo in ("Recebimento", "Pagamento"):
        s = d["capital_giro_prazos"].query("tipo == @tipo")["percentual"].sum()
        add(f"Prazos de {tipo.lower()} somam 100%", f"{num_br(s)}%", abs(s - 100) < 0.01,
            "A soma dos percentuais de cada tipo de prazo deve ser 100%.")
    sw = d["swot"]
    incoer = sw[((sw["ambiente"] == "Interno") & sw["natureza"].isin(["Oportunidade", "Ameaça"])) |
                ((sw["ambiente"] == "Externo") & sw["natureza"].isin(["Força", "Fraqueza"]))]
    add("SWOT: ambiente coerente com a classificação", f"{len(incoer)} fator(es) incoerente(s)", incoer.empty,
        "Interno = Força/Fraqueza; Externo = Oportunidade/Ameaça.")
    neg_cx = [k for k in (1, 2, 3) if A[k]["caixa"] < 0]
    add("Caixa projetado não negativo", "Anos com caixa negativo: " + (", ".join(map(str, neg_cx)) or "nenhum"),
        not neg_cx, "Caixa negativo indica que as fontes não cobrem a operação.")
    lc = div(A[1]["ac"], A[1]["pc"])
    add("Liquidez corrente do Ano 1 >= 1,00", f"{num_br(lc)}x", lc == lc and lc >= 1, "Abaixo de 1,00 há aperto de curto prazo.")
    add("VPL positivo", brl(res["vpl"]), res["vpl"] > 0, "VPL positivo indica criação de valor acima da TMA.")
    return pd.DataFrame(chk)


# ------------------------------------------------------------------------------
# 9. DIAGNÓSTICO SWOT (matriz de estratégias)
# ------------------------------------------------------------------------------
QUADRANTES = {
    "FO": ("Estratégia de Desenvolvimento", "Forças + Oportunidades",
           "Cenário mais favorável: aproveitar oportunidades e expandir (novos mercados, produtos e parcerias)."),
    "FrO": ("Estratégia de Crescimento", "Fraquezas + Oportunidades",
            "Há oportunidades, mas fraquezas internas limitam o aproveitamento: corrigir pontos fracos para crescer."),
    "FA": ("Estratégia de Manutenção", "Forças + Ameaças",
           "Usar as forças para enfrentar as ameaças e manter a posição conquistada."),
    "FrA": ("Estratégia de Sobrevivência", "Fraquezas + Ameaças",
            "Situação mais delicada: reduzir custos, rever o modelo e proteger o caixa."),
}


def pontuacao_swot(sw: pd.DataFrame) -> dict:
    peso = lambda nat: int(sum(2 if i == "Alto" else 1 for i in sw.loc[sw["natureza"] == nat, "impacto"]))
    F, Fr, O, A = peso("Força"), peso("Fraqueza"), peso("Oportunidade"), peso("Ameaça")
    q = {"FO": F + O, "FrO": Fr + O, "FA": F + A, "FrA": Fr + A}
    pred = max(q, key=q.get) if any(q.values()) else None
    return {"F": F, "Fr": Fr, "O": O, "A": A, "quadrantes": q, "predominante": pred}


# ------------------------------------------------------------------------------
# 10. COMPOSIÇÕES PARA GRÁFICOS
# ------------------------------------------------------------------------------
def _serie(dic: dict) -> pd.DataFrame:
    df = pd.DataFrame({"Item": list(dic.keys()), "Valor": [float(v) for v in dic.values()]})
    df = df[df["Valor"] > 0.005].sort_values("Valor", ascending=False).reset_index(drop=True)
    tot = df["Valor"].sum()
    df["%"] = df["Valor"] / tot if tot else 0.0
    return df


def composicoes(res: dict, d: dict) -> dict:
    mat = d["materiais"]
    cat_mat = {n: (c or "Não classificado") for n, c in zip(mat["nome"], mat["categoria"])}
    preco_mat = dict(zip(mat["nome"], mat["valor_unitario"]))
    out = {}

    out["investimento"] = _serie({
        "Ativo imobilizado (investimentos fixos)": res["inv_fixo"], "Investimentos pré-operacionais": res["inv_pre"],
        "Estoque inicial": res["inv_estoque"], "Capital de giro": res["capital_giro"]})

    est = {}
    for _, r in mat.iterrows():
        k = r["categoria"] or "Não classificado"
        est[k] = est.get(k, 0) + r["qtd_estoque_inicial"] * r["valor_unitario"]
    est["Mercadorias para revenda"] = res["est_rev"]
    out["estoque"] = _serie(est)

    ficha = d["ficha_tecnica"]
    custo = {}
    for _, p in res["cat"].iterrows():
        if p["qtd_mensal"] <= 0:
            continue
        if p["origem_custo"] == "Ficha técnica":
            for _, f in ficha[ficha["produto"] == p["descricao"]].iterrows():
                k = cat_mat.get(f["material"], "Não classificado")
                custo[k] = custo.get(k, 0) + p["qtd_mensal"] * f["quantidade"] * preco_mat.get(f["material"], 0)
        elif p["categoria"] == "Revenda":
            custo["Mercadorias para revenda"] = custo.get("Mercadorias para revenda", 0) + p["custo_mensal"]
        else:
            custo["Custo informado (sem ficha técnica)"] = custo.get("Custo informado (sem ficha técnica)", 0) + p["custo_mensal"]
    out["custo_direto"] = _serie(custo)

    out["fontes"] = _serie({"Capital próprio (sócios)": res["capital_proprio"],
                            "Capital de terceiros (empréstimos)": res["capital_terceiros"]})
    out["receita_categoria"] = _serie(res["anos"][1]["rec_cat"])
    top = res["cat"].sort_values("receita_mensal", ascending=False).head(10)
    out["top_produtos"] = _serie(dict(zip(top["descricao"], top["receita_mensal"])))
    inv = res["inv"].sort_values("total", ascending=False)
    fixos = dict(zip(inv["descricao"].head(10), inv["total"].head(10)))
    resto = inv["total"].iloc[10:].sum()
    if resto > 0:
        fixos["Demais itens"] = resto
    out["inv_fixos"] = _serie(fixos)

    a1 = res["anos"][1]
    estrutura = {"CMV (revenda)": a1["cmv"], "CPV (fabricação)": a1["cpv"], "CSV (serviços)": a1["csv"],
                 "Impostos sobre vendas": a1["impostos"], "Outras despesas variáveis": a1["outras_var"],
                 "Pessoal": a1["pessoal"], "Demais custos fixos": a1["fixos"], "Depreciação": a1["dep"],
                 "Juros": a1["juros"], "Lucro líquido": max(a1["resultado"], 0)}
    df = pd.DataFrame({"Item": list(estrutura), "Valor": list(estrutura.values())})
    df["%"] = df["Valor"] / a1["rec_bruta"] if a1["rec_bruta"] else 0.0
    out["estrutura_receita"] = df[df["Valor"] > 0.005].reset_index(drop=True)

    # Gráfico do ponto de equilíbrio (receita x custo total, Ano 1)
    imc = a1["imc"] if a1["imc"] == a1["imc"] else 0.0
    pe = a1["pe_contabil"] if a1["pe_contabil"] == a1["pe_contabil"] else a1["rec_bruta"]
    xmax = max(a1["rec_bruta"], pe) * 1.3 or 1.0
    xs = [xmax * i / 20 for i in range(21)]
    out["equilibrio"] = pd.DataFrame({"Receita": xs, "Receita total": xs,
                                      "Custo total": [a1["cf_total"] + x * (1 - imc) for x in xs],
                                      "Custos fixos": [a1["cf_total"]] * len(xs)})
    out["pe_ponto"] = (pe, a1["rec_bruta"])

    fl = res["fluxos"]
    acum, acum_d, s, sd = [], [], 0.0, 0.0
    for t, f in enumerate(fl):
        s += f
        sd += f / (1 + res["tma"]) ** t
        acum.append(s)
        acum_d.append(sd)
    out["fluxo_acumulado"] = pd.DataFrame({"Ano": [f"Ano {t}" for t in range(len(fl))], "Fluxo do ano": fl,
                                           "Acumulado": acum, "Acumulado descontado (TMA)": acum_d})
    return out


# ------------------------------------------------------------------------------
# 11. SIMULADOR "E SE?" E ANÁLISE DE SENSIBILIDADE (Ano 1)
# ------------------------------------------------------------------------------
def simular(res: dict, preco=0.0, volume=0.0, materiais=0.0, salarios=0.0, fixos=0.0) -> dict:
    """Recalcula o resultado do Ano 1 aplicando variações percentuais (0.10 = +10%)."""
    a1 = res["anos"][1]
    fp, fv, fm, fs, ff = 1 + preco, 1 + volume, 1 + materiais, 1 + salarios, 1 + fixos
    rec = a1["rec_bruta"] * fp * fv
    cdir = a1["custo_dir"] * fv * fm
    var = rec * (res["aliq"] + res["var_pct"])
    mc = rec - cdir - var
    cf = a1["pessoal"] * fs + a1["fixos"] * ff + a1["dep"]
    resultado = mc - cf - a1["juros"]
    imc = div(mc, rec)
    return {"receita": rec, "custos_variaveis": cdir + var, "mc": mc, "custos_fixos": cf,
            "resultado": resultado, "lucratividade": div(resultado, rec),
            "pe": div(cf, imc) if imc == imc and imc > 0 else float("nan")}


def sensibilidade(res: dict, delta: float = 0.10) -> pd.DataFrame:
    base = simular(res)["resultado"]
    linhas = []
    for nome, chave_, sinal in [("Preço de venda", "preco", 1), ("Volume de vendas", "volume", 1),
                                ("Custo de materiais e mercadorias", "materiais", -1),
                                ("Salários e encargos", "salarios", -1), ("Demais custos fixos", "fixos", -1)]:
        up = simular(res, **{chave_: delta})["resultado"] - base
        dn = simular(res, **{chave_: -delta})["resultado"] - base
        linhas.append({"Variável": nome, f"+{int(delta*100)}%": up, f"-{int(delta*100)}%": dn,
                       "Impacto absoluto": max(abs(up), abs(dn))})
    return pd.DataFrame(linhas).sort_values("Impacto absoluto", ascending=False).reset_index(drop=True)


# ------------------------------------------------------------------------------
# 12. MERCADO (TAM / SAM / SOM) E CAPACIDADE x DEMANDA
# ------------------------------------------------------------------------------
def mercado(proj: dict, res: dict) -> dict:
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    tam_c, sam_c = float(proj.get("tam_clientes") or 0), float(proj.get("sam_clientes") or 0)
    som_c = sam_c * float(proj.get("som_pct") or 0) / 100
    tam, sam, som = tam_c * gasto, sam_c * gasto, som_c * gasto
    rec1 = res["anos"][1]["rec_bruta"]
    return {"tam": tam, "sam": sam, "som": som, "tam_c": tam_c, "sam_c": sam_c, "som_c": som_c,
            "receita_ano1": rec1, "pct_som": div(rec1, som), "pct_sam": div(rec1, sam),
            "preenchido": tam > 0 and sam > 0 and som > 0}


def capacidade_demanda(d: dict, res: dict) -> pd.DataFrame:
    cat = res["cat"]
    fator12 = (1 + res["g_m"]) ** 11
    linhas = []
    for _, r in d["capacidade"].iterrows():
        alvo = str(r["produto_ou_grupo"])
        if alvo in ("Fabricação Própria", "Serviço", "Revenda"):
            q = cat.loc[cat["categoria"] == alvo, "qtd_mensal"].sum()
        elif alvo == "Todos os produtos":
            q = cat["qtd_mensal"].sum()
        else:
            q = cat.loc[cat["descricao"] == alvo, "qtd_mensal"].sum()
        dem1, dem12 = float(q), float(q) * fator12
        cap = float(r["capacidade_mensal"])
        linhas.append({"Recurso": r["recurso"], "Atende": alvo, "Capacidade/mês": cap,
                       "Demanda Mês 1": dem1, "Demanda Mês 12": dem12,
                       "Uso Mês 1": div(dem1, cap), "Uso Mês 12": div(dem12, cap)})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 13. SEMÁFORO DOS INDICADORES COM "EXPLIQUE ESTE NÚMERO"
# ------------------------------------------------------------------------------
def semaforo(res: dict) -> list:
    a1 = res["anos"][1]
    it, tma = res["investimento_total"], res["tma"]
    out = []

    def add(nome, valor, txt, status, leitura, conta):
        out.append({"Indicador": nome, "Valor": txt, "Status": status, "Leitura": leitura,
                    "Como foi calculado": conta, "valor_num": valor})

    luc = div(a1["resultado"], a1["rec_bruta"])
    add("Lucratividade líquida (Ano 1)", luc, pct(luc),
        "🟢" if luc >= 0.10 else ("🟡" if luc > 0 else "🔴"),
        "Quanto sobra de lucro para cada R$ 1,00 vendido. Acima de 10% é bom para pequenos negócios de alimentação e varejo; "
        "negativo significa prejuízo.",
        f"{brl(a1['resultado'])} ÷ {brl(a1['rec_bruta'])} = {pct(luc)}")
    imc = a1["imc"]
    add("Margem de contribuição", imc, pct(imc),
        "🟢" if imc >= 0.40 else ("🟡" if imc >= 0.25 else "🔴"),
        "Parte de cada venda que sobra, depois dos custos variáveis, para pagar a estrutura fixa. Margens baixas "
        "exigem volume muito alto.",
        f"{brl(a1['mc'])} ÷ {brl(a1['rec_bruta'])} = {pct(imc)}")
    ms = div(a1["rec_bruta"] - a1["pe_contabil"], a1["rec_bruta"])
    add("Margem de segurança", ms, pct(ms),
        "🟢" if ms >= 0.20 else ("🟡" if ms > 0 else "🔴"),
        "Quanto as vendas podem cair antes de o negócio ter prejuízo. Abaixo de 0% significa que a receita prevista "
        "não cobre o ponto de equilíbrio.",
        f"({brl(a1['rec_bruta'])} − PE {brl(a1['pe_contabil'])}) ÷ {brl(a1['rec_bruta'])} = {pct(ms)}")
    add("VPL (3 anos)", res["vpl"], brl(res["vpl"]),
        "🟢" if res["vpl"] > 0 else "🔴",
        "Valor criado pelo projeto acima do retorno mínimo exigido (TMA). Positivo: o investimento compensa.",
        "Soma dos fluxos descontados: " + " + ".join(f"{brl(f)}/(1+{num_br(tma*100,1)}%)^{t}" for t, f in enumerate(res["fluxos"])))
    tir = res["tir"]
    add("TIR", tir if tir is not None else float("nan"), pct(tir) if tir is not None else "n.a.",
        "🟢" if tir is not None and tir > tma else "🔴",
        f"Rentabilidade anual do projeto. Deve superar a TMA de {pct(tma)}.",
        "Taxa que faz o VPL ser zero nos fluxos: " + "; ".join(brl(f) for f in res["fluxos"]))
    pb = res["payback_desc"]
    add("Payback descontado", pb, f"{num_br(pb,1)} meses" if pb == pb else "> 36 meses",
        "🟢" if pb == pb and pb <= 24 else ("🟡" if pb == pb else "🔴"),
        "Tempo para recuperar o investimento considerando o valor do dinheiro no tempo.",
        f"Investimento de {brl(it)} recuperado pelos fluxos descontados à TMA de {pct(tma)}")
    lc = div(a1["ac"], a1["pc"])
    add("Liquidez corrente (Ano 1)", lc, f"{num_br(lc)}x",
        "🟢" if lc >= 1.5 else ("🟡" if lc >= 1 else "🔴"),
        "Quantos reais de bens e direitos de curto prazo existem para cada R$ 1,00 de dívida de curto prazo.",
        f"{brl(a1['ac'])} ÷ {brl(a1['pc'])} = {num_br(lc)}x")
    pes = div(a1["pessoal"], a1["rec_liquida"])
    add("Peso da folha de pagamento", pes, pct(pes),
        "🟢" if pes <= 0.25 else ("🟡" if pes <= 0.35 else "🔴"),
        "Quanto da receita líquida vai para salários, encargos e benefícios. Acima de 35% costuma apertar o resultado.",
        f"{brl(a1['pessoal'])} ÷ {brl(a1['rec_liquida'])} = {pct(pes)}")
    fm = res["fluxo_mensal"]
    minimo = float(fm["Saldo final"].min()) if not fm.empty else 0.0
    mes_min = fm.loc[fm["Saldo final"].idxmin(), "Mês"] if not fm.empty else "-"
    add("Menor saldo de caixa no Ano 1", minimo, brl(minimo),
        "🟢" if minimo > res["desembolso_m1"] * 0.5 else ("🟡" if minimo >= 0 else "🔴"),
        f"O caixa fica mais apertado em {mes_min}. Negativo indica que faltará dinheiro nesse mês.",
        f"Saldo inicial + recebimentos − pagamentos, mês a mês (menor valor em {mes_min})")
    return out


# ------------------------------------------------------------------------------
# 14. CHECAGENS COM INDICAÇÃO DE ONDE CORRIGIR
# ------------------------------------------------------------------------------
def tabela_checagens_v5(res: dict, d: dict, proj: dict) -> pd.DataFrame:
    base = tabela_checagens(res, d)
    destinos = {
        "Fontes de recursos x investimento total": ("investimentos", "Investimentos → Fontes de recursos",
            "Ajuste o capital dos sócios (aba Projeto) ou os empréstimos para que a soma seja igual ao investimento total."),
        "Margem unitária positiva": ("produtos", "Produtos e materiais → Produtos e serviços",
            "Revise o preço de venda ou o custo dos itens listados."),
        "Produtos com custo cadastrado": ("produtos", "Produtos e materiais → Ficha técnica",
            "Monte a ficha técnica ou informe o custo dos itens listados."),
        "Fichas técnicas com materiais válidos": ("produtos", "Produtos e materiais → Ficha técnica",
            "Cadastre os materiais que faltam ou troque o insumo na ficha técnica."),
        "Prazos de recebimento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de recebimento para somarem 100%."),
        "Prazos de pagamento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de pagamento para somarem 100%."),
        "SWOT: ambiente coerente com a classificação": ("forcas", "Estratégia → Fatores internos e externos",
            "Fatores internos são Forças/Fraquezas; externos são Oportunidades/Ameaças."),
        "Caixa projetado não negativo": ("fluxo", "Finanças → Fluxo de caixa",
            "Aumente a reserva/capital de giro, renegocie prazos ou reduza custos."),
        "Liquidez corrente do Ano 1 >= 1,00": ("investimentos", "Investimentos → Capital de giro",
            "Aumente o capital de giro ou alongue o prazo das dívidas."),
        "VPL positivo": ("simulador", "Simulador e sensibilidade",
            "Veja quais variáveis mais afetam o resultado e teste ajustes de preço, volume e custos."),
    }
    base["destino"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[0])
    base["Onde corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[1])
    base["Como corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[2])
    extras = []
    mk = mercado(proj, res)
    if mk["preenchido"]:
        ok = mk["pct_som"] <= 1
        extras.append({"Teste": "Meta de vendas cabe no mercado (SOM)", "Valor apurado": f"Receita = {pct(mk['pct_som'])} do SOM",
                       "Situação": "OK" if ok else "VERIFICAR",
                       "Observação": "A receita do Ano 1 não deve superar o mercado que a empresa consegue conquistar.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Revise as quantidades vendidas ou a estimativa de mercado."})
    else:
        extras.append({"Teste": "Tamanho do mercado estimado", "Valor apurado": "não preenchido", "Situação": "VERIFICAR",
                       "Observação": "Estime TAM, SAM e SOM para verificar se a meta de vendas é realista.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Preencha clientes, percentual de participação e gasto anual por cliente."})
    cd = capacidade_demanda(d, res)
    if not cd.empty:
        exced = cd[cd["Uso Mês 12"] > 1]
        extras.append({"Teste": "Capacidade produtiva suficiente", "Valor apurado": f"{len(exced)} recurso(s) acima de 100%",
                       "Situação": "OK" if exced.empty else "VERIFICAR",
                       "Observação": ", ".join(exced["Recurso"]) or "A capacidade atende a demanda prevista até o Mês 12.",
                       "destino": "operacional", "Onde corrigir": "Marketing e operação → Plano operacional",
                       "Como corrigir": "Aumente a capacidade (equipamentos, pessoas, turnos) ou reveja a meta de vendas."})
    return pd.concat([base, pd.DataFrame(extras)], ignore_index=True)


# ------------------------------------------------------------------------------
# 15. RAIO-X DO PLANO (nota de 0 a 100 e principais correções)
# ------------------------------------------------------------------------------
def raio_x(res: dict, chk: pd.DataFrame, completude: float, sem: list) -> dict:
    pontos_chk = 30 * (chk["Situação"] == "OK").mean() if len(chk) else 0
    pontos_comp = 30 * completude
    verdes = sum(1 for s in sem if s["Status"] == "🟢")
    amarelos = sum(1 for s in sem if s["Status"] == "🟡")
    pontos_fin = 40 * (verdes + 0.5 * amarelos) / max(len(sem), 1)
    nota = round(pontos_chk + pontos_comp + pontos_fin)
    acoes = [f"{r['Teste']}: {r['Como corrigir'] or r['Observação']}" for _, r in chk[chk["Situação"] != "OK"].iterrows()]
    acoes += [f"{s['Indicador']} em {s['Valor']}: {s['Leitura']}" for s in sem if s["Status"] == "🔴"]
    return {"nota": nota, "acoes": acoes[:5], "pontos": (round(pontos_comp), round(pontos_chk), round(pontos_fin))}

# ##############################################################################
# PARTE: pnbox_pdf.py
# ##############################################################################



AZUL = (28, 49, 80)
def _fontes_matplotlib():
    try:
        import matplotlib
        base = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
        return [(str(base / "DejaVuSans.ttf"), str(base / "DejaVuSans-Bold.ttf"), str(base / "DejaVuSans-Oblique.ttf"))]
    except Exception:
        return []


FONTES_CANDIDATAS = [
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/ariali.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf", "/Library/Fonts/Arial Italic.ttf"),
] + _fontes_matplotlib()
TROCAS = {"–": "-", "—": "-", "−": "-", "×": "x", "÷": "/", "Σ": "Soma", "“": '"', "”": '"',
          "’": "'", "‘": "'", "…": "...", "•": "-", "≥": ">=", "≤": "<="}


class Relatorio(FPDF):
    def __init__(self, nome):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.nome = nome
        self.unicode = False
        for reg, neg, ita in FONTES_CANDIDATAS:
            if all(Path(p).exists() for p in (reg, neg, ita)):
                self.add_font("Base", "", reg)
                self.add_font("Base", "B", neg)
                self.add_font("Base", "I", ita)
                self.unicode = True
                break
        self.fam = "Base" if self.unicode else "Helvetica"
        self.set_auto_page_break(auto=True, margin=15)
        self.set_margins(15, 15, 15)

    def t(self, v) -> str:
        s = "" if v is None else str(v)
        if s == "nan":
            s = ""
        for a, b in TROCAS.items():
            s = s.replace(a, b)
        if not self.unicode:
            s = s.encode("latin-1", "replace").decode("latin-1")
        return s

    def fonte(self, estilo="", tam=9):
        self.set_font(self.fam, estilo, tam)

    def footer(self):
        self.set_y(-11)
        self.fonte("I", 7)
        self.set_text_color(120, 120, 120)
        self.cell(0, 5, self.t(f"Plano de Negócio - {self.nome}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.cell(0, 5, f"Pág. {self.page_no()}/{{nb}}", align="R")
        self.set_text_color(0, 0, 0)

    # --- blocos de conteúdo -------------------------------------------------
    def secao(self, texto, nova_pagina=False):
        self.n_secao = getattr(self, "n_secao", 0) + 1
        texto = f"{self.n_secao}. {texto}"
        if nova_pagina:
            self.add_page()
        elif self.get_y() > 245:
            self.add_page()
        self.ln(3)
        self.set_fill_color(*AZUL)
        self.set_text_color(255, 255, 255)
        self.fonte("B", 11)
        self.cell(0, 8, self.t(f"  {texto}"), fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def subtitulo(self, texto):
        if self.get_y() > 260:
            self.add_page()
        self.ln(1)
        self.fonte("B", 10)
        self.set_text_color(*AZUL)
        self.cell(0, 6, self.t(texto), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(0, 0, 0)

    def paragrafo(self, texto, estilo="", tam=9):
        if not str(texto).strip():
            return
        self.fonte(estilo, tam)
        self.multi_cell(0, 4.8, self.t(texto), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(1)

    def campo(self, rotulo, valor):
        if not str(valor).strip():
            return
        self.set_x(self.l_margin)
        self.fonte("B", 9)
        self.write(4.8, self.t(f"{rotulo}: "))
        self.fonte("", 9)
        self.write(4.8, self.t(valor))
        self.ln(5.5)

    def tabela(self, df: pd.DataFrame, larguras=None, alinhar=None, tam=8, destaque_total=False):
        if df is None or df.empty:
            self.paragrafo("(sem registros)", "I", 8)
            return
        self.fonte("", tam)
        self.set_fill_color(255, 255, 255)
        cols = list(df.columns)
        if alinhar is None:
            alinhar = ["LEFT"] + ["RIGHT"] * (len(cols) - 1)
        cab = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(70, 95, 130))
        with self.table(col_widths=larguras, text_align=alinhar, line_height=tam * 0.55,
                        headings_style=cab, first_row_as_headings=True,
                        cell_fill_color=(242, 245, 250), cell_fill_mode="ROWS") as tb:
            linha = tb.row()
            for c in cols:
                linha.cell(self.t(c))
            n = len(df)
            for i, (_, r) in enumerate(df.iterrows()):
                linha = tb.row()
                estilo = FontFace(emphasis="BOLD") if destaque_total and i == n - 1 else None
                for c in cols:
                    linha.cell(self.t(r[c]), style=estilo)
        self.ln(2)

    def matriz_swot(self, sw: pd.DataFrame):
        """Desenha a matriz SWOT 2x2 colorida (Forças, Fraquezas, Oportunidades, Ameaças)."""
        if self.get_y() > 150:
            self.add_page()
        x0, y0 = self.l_margin, self.get_y() + 2
        lat, cab, w, h, gap = 10, 9, 84, 78, 1.5
        cores = {"Força": ((38, 84, 140), (255, 255, 255)), "Fraqueza": ((214, 0, 110), (255, 255, 255)),
                 "Oportunidade": ((226, 230, 235), (40, 50, 60)), "Ameaça": ((92, 102, 114), (255, 255, 255))}
        titulos = {"Força": "FORÇAS", "Fraqueza": "FRAQUEZAS", "Oportunidade": "OPORTUNIDADES", "Ameaça": "AMEAÇAS"}
        # cabeçalhos
        self.fonte("B", 9)
        self.set_text_color(255, 255, 255)
        self.set_fill_color(38, 84, 140)
        self.set_xy(x0 + lat + gap, y0)
        self.cell(w, cab, self.t("FATORES POSITIVOS"), fill=True, align="C")
        self.set_fill_color(214, 0, 110)
        self.set_xy(x0 + lat + gap * 2 + w, y0)
        self.cell(w, cab, self.t("FATORES NEGATIVOS"), fill=True, align="C")
        for j, (rot, cor) in enumerate([("INTERNOS", (38, 84, 140)), ("EXTERNOS", (200, 204, 210))]):
            yy = y0 + cab + gap + j * (h + gap)
            self.set_fill_color(*cor)
            self.rect(x0, yy, lat, h, style="F")
            self.set_text_color(255, 255, 255) if j == 0 else self.set_text_color(40, 50, 60)
            self.fonte("B", 9)
            y_txt = yy + (h - len(rot) * 5) / 2
            for i, letra in enumerate(self.t(rot)):
                self.set_xy(x0, y_txt + i * 5)
                self.cell(lat, 5, letra, align="C")
        pos = {"Força": (0, 0), "Fraqueza": (1, 0), "Oportunidade": (0, 1), "Ameaça": (1, 1)}
        for nat, (c, l) in pos.items():
            xx = x0 + lat + gap + c * (w + gap)
            yy = y0 + cab + gap + l * (h + gap)
            fundo, texto = cores[nat]
            self.set_fill_color(*fundo)
            self.rect(xx, yy, w, h, style="F")
            self.set_text_color(*texto)
            self.set_xy(xx + 3, yy + 3)
            self.fonte("B", 11)
            self.cell(w - 6, 6, self.t(titulos[nat]))
            itens = sw[sw["natureza"] == nat]
            self.set_xy(xx + 3, yy + 10)
            self.fonte("", 7.5)
            linhas = []
            for _, r in itens.iterrows():
                marca = "*" if int(r.get("priorizado", 0) or 0) else "-"
                linhas.append(f"{marca} {r['fator']} ({r['impacto']})")
            txt = "\n".join(linhas) if linhas else "(nenhum fator cadastrado)"
            self.multi_cell(w - 6, 3.8, self.t(txt), new_x=XPos.LEFT, new_y=YPos.NEXT, max_line_height=3.8)
        self.set_text_color(0, 0, 0)
        self.set_xy(x0, y0 + cab + 2 * h + 3 * gap + 1)
        self.fonte("I", 7)
        self.cell(0, 4, self.t("* fator priorizado. Entre parênteses, o impacto (Alto/Baixo)."),
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)


# ==============================================================================
# GRÁFICOS (matplotlib) PARA O PDF
# ==============================================================================
PALETA = ["#26548c", "#d6006e", "#5c6672", "#3a8fd1", "#f39c12", "#27ae60", "#8e44ad", "#c0392b",
          "#16a085", "#7f8c8d", "#2c3e50", "#e67e22"]


def _fmt_rs(v):
    return ("R$ " + f"{v:,.0f}".replace(",", "."))


def _fmt_curto(v, _=None):
    a = abs(v)
    if a >= 1e6:
        return f"R$ {v/1e6:.1f} mi".replace(".", ",")
    if a >= 1e3:
        return f"R$ {v/1e3:.0f} mil"
    return f"R$ {v:.0f}"


def graficos_png(comp: dict) -> list:
    """Devolve [(título, bytes PNG)]. Se o matplotlib não estiver instalado, devolve lista vazia."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return []
    import io as _io
    plt.rcParams.update({"font.size": 9, "font.family": "DejaVu Sans"})
    saida = []

    def salvar(fig, titulo):
        b = _io.BytesIO()
        fig.savefig(b, format="png", dpi=170, bbox_inches="tight")
        plt.close(fig)
        saida.append((titulo, b.getvalue()))

    def rosca(df, titulo):
        if df.empty:
            return
        fig, ax = plt.subplots(figsize=(5.2, 4.4))
        cores = PALETA[: len(df)] if len(df) <= len(PALETA) else None
        cunhas, _, autot = ax.pie(df["Valor"], colors=cores, startangle=90, counterclock=False,
                                  autopct=lambda p: f"{p:.1f}%".replace(".", ",") if p >= 4 else "",
                                  pctdistance=0.78, wedgeprops=dict(width=0.42, edgecolor="white"))
        for t in autot:
            t.set_color("white"); t.set_fontsize(8); t.set_fontweight("bold")
        ax.text(0, 0, _fmt_rs(df["Valor"].sum()), ha="center", va="center", fontsize=9, fontweight="bold")
        ax.legend(cunhas, [f"{i} ({p*100:.1f}%)".replace(".", ",") for i, p in zip(df["Item"], df["%"])],
                  loc="upper center", bbox_to_anchor=(0.5, -0.02), ncol=1, frameon=False, fontsize=8)
        ax.set_title(titulo, fontsize=10, fontweight="bold")
        ax.axis("equal")
        salvar(fig, titulo)

    def barras(df, titulo):
        if df.empty:
            return
        df = df.iloc[::-1]
        fig, ax = plt.subplots(figsize=(7.5, 0.38 * len(df) + 1.0))
        ax.barh([str(i)[:55] for i in df["Item"]], df["Valor"], color="#26548c")
        for y, (v, p) in enumerate(zip(df["Valor"], df["%"])):
            ax.text(v, y, f"  {p*100:.1f}%".replace(".", ","), va="center", fontsize=8)
        ax.set_title(titulo, fontsize=10, fontweight="bold", loc="left")
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_curto)); ax.xaxis.set_major_locator(matplotlib.ticker.MaxNLocator(5))
        ax.spines[["top", "right"]].set_visible(False)
        ax.set_xlim(0, df["Valor"].max() * 1.15)
        salvar(fig, titulo)

    rosca(comp["investimento"], "Composição do investimento total")
    rosca(comp["fontes"], "Fontes de recursos")
    rosca(comp["estoque"], "Estoque inicial por categoria")
    rosca(comp["custo_direto"], "Custo direto mensal por categoria")
    rosca(comp["receita_categoria"], "Receita por categoria")
    barras(comp["inv_fixos"], "Maiores itens do ativo imobilizado")
    barras(comp["top_produtos"], "Produtos/serviços com maior receita mensal")
    barras(comp["estrutura_receita"], "Custos, despesas e lucro em % da receita (Ano 1)")

    eq = comp["equilibrio"]
    pe, rec = comp["pe_ponto"]
    fig, ax = plt.subplots(figsize=(7.5, 3.8))
    ax.plot(eq["Receita"], eq["Receita total"], color="#26548c", lw=2.5, label="Receita total")
    ax.plot(eq["Receita"], eq["Custo total"], color="#d6006e", lw=2.5, label="Custo total")
    ax.plot(eq["Receita"], eq["Custos fixos"], color="#5c6672", lw=1.5, ls="--", label="Custos fixos")
    ax.axvline(pe, color="black", ls=":", lw=1)
    ax.annotate(f"PE: {_fmt_rs(pe)}", (pe, pe), xytext=(10, -25), textcoords="offset points", fontsize=8)
    ax.axvline(rec, color="#27ae60", ls=":", lw=1)
    ax.annotate(f"Receita projetada: {_fmt_rs(rec)}", (rec, rec), xytext=(-150, 15), textcoords="offset points", fontsize=8)
    fmt = matplotlib.ticker.FuncFormatter(_fmt_curto)
    ax.xaxis.set_major_formatter(fmt); ax.yaxis.set_major_formatter(fmt); ax.xaxis.set_major_locator(matplotlib.ticker.MaxNLocator(6))
    ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Ponto de equilíbrio contábil (Ano 1)", fontsize=10, fontweight="bold", loc="left")
    salvar(fig, "Ponto de equilíbrio")

    fa = comp["fluxo_acumulado"]
    fig, ax = plt.subplots(figsize=(7.5, 3.5))
    x = range(len(fa))
    ax.bar([i - 0.2 for i in x], fa["Acumulado"], 0.4, color="#26548c", label="Acumulado")
    ax.bar([i + 0.2 for i in x], fa["Acumulado descontado (TMA)"], 0.4, color="#d6006e", label="Acumulado descontado (TMA)")
    ax.axhline(0, color="black", lw=0.8)
    ax.set_xticks(list(x), fa["Ano"]); ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_curto))
    ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Fluxo de caixa acumulado do projeto (payback)", fontsize=10, fontweight="bold", loc="left")
    salvar(fig, "Payback")
    return saida


# ==============================================================================
# MONTAGEM DO RELATÓRIO
# ==============================================================================
def gerar_pdf(proj, d, res, extras) -> bytes:
    """extras: dict com dre, indicadores, cenarios, checagens, swot_diag, quadrantes."""

    pdf = Relatorio(proj["nome_negocio"])
    pdf.alias_nb_pages()
    pdf.add_page()

    # ---------------- CAPA ----------------
    pdf.ln(50)
    pdf.fonte("B", 24)
    pdf.set_text_color(*AZUL)
    pdf.multi_cell(0, 12, pdf.t(proj["nome_negocio"]), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.fonte("", 14)
    pdf.cell(0, 9, pdf.t("Plano de Negócio e Análise de Viabilidade Econômico-Financeira"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(15)
    pdf.fonte("", 10)
    for rot, k in [("Objetivo", "objetivo"), ("Cidade", "cidade_uf"), ("Setor de atuação", "setor_atuacao"),
                   ("Enquadramento tributário", "enquadramento"), ("Porte", "porte"), ("Contatos", "contatos")]:
        pdf.cell(0, 6, pdf.t(f"{rot}: {proj[k]}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 6, pdf.t(f"Período de planejamento: {int(proj['periodo_anos'])} ano(s)"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    autores = [a.strip() for a in re.split(r"[;\n]", proj.get("autores") or "") if a.strip()]
    if autores or proj.get("turma") or proj.get("professor"):
        pdf.ln(12)
        pdf.fonte("B", 10)
        pdf.cell(0, 6, pdf.t("Elaborado por"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.fonte("", 10)
        for a in autores:
            pdf.cell(0, 5.5, pdf.t(a), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        for rot, k in [("Turma", "turma"), ("Professor(a)", "professor"), ("Entrega", "data_entrega")]:
            if proj.get(k):
                pdf.cell(0, 5.5, pdf.t(f"{rot}: {proj[k]}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if str(proj.get("sumario_executivo") or "").strip():
        pdf.secao("SUMÁRIO EXECUTIVO", nova_pagina=True)
        for par in str(proj["sumario_executivo"]).split("\n"):
            pdf.paragrafo(par, "", 10)

    # ---------------- 1. SÓCIOS ----------------
    pdf.secao("INFORMAÇÕES DO PLANO E QUADRO SOCIETÁRIO", nova_pagina=True)
    s = d["socios"]
    pdf.tabela(pd.DataFrame({
        "Sócio": s["nome"], "Função": s["funcao"], "Minicurrículo": s["perfil"],
        "Capital (R$)": s["capital_integralizado"].map(brl)}), larguras=(40, 34, 76, 30),
        alinhar=("LEFT", "LEFT", "LEFT", "RIGHT"))

    # ---------------- 2. SEGMENTOS ----------------
    pdf.secao("SEGMENTAÇÃO DE MERCADO")
    sg = d["segmentos"].copy()
    notas = ["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]
    sg["media"] = sg[notas].mean(axis=1)
    sg = sg.sort_values("media", ascending=False)
    pdf.tabela(pd.DataFrame({
        "Segmento": sg["nome"], "Descrição": sg["descricao"],
        "Tam.": sg["nota_tamanho"].map(lambda v: num_br(v, 0)), "Lucro": sg["nota_lucro"].map(lambda v: num_br(v, 0)),
        "Long.": sg["nota_longevidade"].map(lambda v: num_br(v, 0)),
        "Conc.": sg["nota_concorrencia"].map(lambda v: num_br(v, 0)),
        "Ader.": sg["nota_aderencia"].map(lambda v: num_br(v, 0)), "Média": sg["media"].map(num_br)}),
        larguras=(40, 70, 11, 11, 11, 11, 11, 15), alinhar=("LEFT", "LEFT") + ("CENTER",) * 6)
    pdf.paragrafo("Notas de 1 a 10: tamanho do segmento, potencial de lucro, longevidade, pouca concorrência "
                  "e aderência da oferta às demandas do segmento.", "I", 7.5)
    mk = mercado(proj, res)
    if mk["preenchido"]:
        pdf.subtitulo("Tamanho do mercado (TAM, SAM e SOM)")
        pdf.tabela(pd.DataFrame({
            "Nível": ["TAM - mercado total", "SAM - mercado alcançável", "SOM - mercado conquistável"],
            "Descrição": [proj.get("tam_desc", ""), proj.get("sam_desc", ""), f"{num_br(proj.get('som_pct', 0), 1)}% do SAM"],
            "Clientes": [num_br(mk["tam_c"], 0), num_br(mk["sam_c"], 0), num_br(mk["som_c"], 0)],
            "R$ por ano": [brl(mk["tam"]), brl(mk["sam"]), brl(mk["som"])]}),
            larguras=(42, 78, 25, 35), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT"))
        pdf.paragrafo(f"A receita prevista no Ano 1 ({brl(mk['receita_ano1'])}) corresponde a {pct(mk['pct_som'])} do SOM "
                      f"e a {pct(mk['pct_sam'])} do SAM. Fontes: {proj.get('mercado_fontes', '')}", "I", 8)

    # ---------------- 3. PROPOSTA DE VALOR ----------------
    pdf.secao("PROPOSTA DE VALOR")
    props = d.get("propostas")
    if props is not None and not props.empty:
        for _, pr in props.iterrows():
            pdf.subtitulo(f"Persona: {pr['persona_nome']}")
            if pr["valores"]:
                pdf.campo("Valores entregues", pr["valores"])
            pdf.paragrafo(f"{proj['nome_negocio']} ajuda {pr['persona_nome']}, que deseja {pr['que_desejam']}, por meio de "
                          f"{pr['por_meio']}. Diferentemente de {pr['diferentemente']}, oferecemos {pr['oferecem']}.", "", 10)
    pv = d["proposta_valor"]
    if (props is None or props.empty) and pv.get("segmento_alvo"):
        pdf.paragrafo(f"Nossa empresa ({proj['nome_negocio']}) ajuda {pv['segmento_alvo']} que desejam "
                      f"{pv['beneficios_desejados']} por meio de {pv['meios_entrega']}. Diferentemente de "
                      f"{pv['diferencial_competitivo']}, os nossos produtos ou serviços oferecem "
                      f"{pv['entrega_produtos_servicos']}", "", 10)

    # ---------------- 4. PERSONAS ----------------
    pdf.secao("PERSONAS", nova_pagina=True)
    for _, p in d["personas"].iterrows():
        pdf.subtitulo(f"{p['nome']}, {int(p['idade'])} anos - {p['profissao']}")
        pdf.campo("Estado civil", p["estado_civil"])
        pdf.campo("Renda mensal", brl(p["renda_mensal"]) if p["renda_mensal"] else "")
        pdf.campo("Escolaridade", p["escolaridade"])
        pdf.campo("Cidade", p["cidade"])
        pdf.paragrafo(p["descricao"])
        if str(p["citacao"]).strip():
            pdf.paragrafo(f'"{p["citacao"]}"', "I")
        for rot, k in [("Hábitos e comportamentos", "habitos"), ("Características", "caracteristicas"),
                       ("Fluência digital", "fluencia_digital"), ("Redes sociais", "redes_sociais"),
                       ("Quem influencia", "quem_influencia")]:
            pdf.campo(rot, p[k])
        rot_df = pd.DataFrame({"Manhã": [p["rotina_manha"]], "Tarde": [p["rotina_tarde"]], "Noite": [p["rotina_noite"]]})
        pdf.tabela(rot_df, alinhar=("LEFT",) * 3)
        gd = pd.DataFrame({"Ganhos": [p["ganhos"]], "Desafios": [p["desafios"]]})
        pdf.tabela(gd, alinhar=("LEFT",) * 2)
        rf = pd.DataFrame({"Dúvidas ou objeções": [p["duvidas_objecoes"]],
                           "Aborrecimentos a evitar": [p["aborrecimentos"]],
                           "Experiência ideal e como o negócio ajuda": [p["experiencia_ideal"]]})
        pdf.tabela(rf, alinhar=("LEFT",) * 3)

    # ---------------- 5. JORNADA ----------------
    pdf.secao("JORNADA DO CLIENTE")
    j = d["jornada_cliente"]
    for persona, grupo in j.groupby("persona_nome", sort=False):
        pdf.subtitulo(f"Persona: {persona}")
        pdf.tabela(pd.DataFrame({"Fase": grupo["macro_fase"], "Etapa": grupo["etapa"],
                                 "Ações do cliente": grupo["acoes_cliente"], "Pensamentos": grupo["pensamentos"],
                                 "Pontos de contato": grupo["pontos_contato"], "Emoção": grupo["estado_emocional"],
                                 "Oportunidades": grupo["oportunidades"]}),
                   larguras=(14, 22, 34, 26, 28, 20, 36), alinhar=("LEFT",) * 7, tam=7)

    # ---------------- 6. CONCORRÊNCIA ----------------
    pdf.secao("ANÁLISE DA CONCORRÊNCIA", nova_pagina=True)
    piv = extras["concorrencia_pivot"]
    if not piv.empty:
        tab = piv.copy()
        media = tab.drop(columns=["Atributo"]).mean()
        tab = pd.concat([tab, pd.DataFrame([{"Atributo": "Média geral", **media.to_dict()}])], ignore_index=True)
        for c in tab.columns[1:]:
            tab[c] = tab[c].map(lambda v: num_br(v, 1))
        n = len(tab.columns) - 1
        pdf.tabela(tab, larguras=(50,) + (130 / n,) * n, alinhar=("LEFT",) + ("CENTER",) * n,
                   destaque_total=True, tam=7.5)
        nossa = piv.set_index("Atributo")
        if "Nossa Empresa" in nossa.columns and len(nossa.columns) > 1:
            outros = nossa.drop(columns=["Nossa Empresa"]).max(axis=1)
            fortes = nossa.index[nossa["Nossa Empresa"] > outros].tolist()
            fracos = nossa.index[nossa["Nossa Empresa"] < outros].tolist()
            pdf.campo("Forças frente à concorrência", ", ".join(fortes) or "nenhuma")
            pdf.campo("Fraquezas frente à concorrência", ", ".join(fracos) or "nenhuma")
    else:
        pdf.paragrafo("(sem registros)", "I", 8)

    # ---------------- 7. AMBIENTE E SWOT ----------------
    pdf.secao("ANÁLISE DE AMBIENTE E MATRIZ SWOT", nova_pagina=True)
    sw = d["swot"]
    pdf.matriz_swot(sw)
    diag = extras["swot_diag"]
    q = diag["quadrantes"]
    pdf.subtitulo("Matriz de estratégias (pontuação: impacto alto = 2; baixo = 1)")
    pdf.tabela(pd.DataFrame({
        "Quadrante": [f"{QUADRANTES[k][0]} ({QUADRANTES[k][1]})" for k in q],
        "Pontuação": [str(v) for v in q.values()]}), larguras=(150, 30))
    if diag["predominante"]:
        nome, comb, txt = QUADRANTES[diag["predominante"]]
        pdf.paragrafo(f"Situação do negócio: {nome} ({comb}). {txt}", "B")
    pdf.subtitulo("Detalhamento dos fatores")
    pdf.tabela(pd.DataFrame({"Ambiente": sw["ambiente"], "Dimensão": sw["dimensao"], "Fator": sw["fator"],
                             "Classif.": sw["natureza"] + " / " + sw["impacto"], "Diagnóstico": sw["diagnostico"],
                             "Plano de ação": sw["plano_acao"]}),
               larguras=(16, 20, 30, 24, 45, 45), alinhar=("LEFT",) * 6, tam=7)

    # ---------------- 8. CANAIS E FUNIL ----------------
    pdf.secao("CANAIS DE AQUISIÇÃO E FUNIL DE VENDAS", nova_pagina=True)
    cn = d["canais"].copy()
    if not cn.empty:
        cn["media"] = cn[["atratividade", "alcance", "cliente", "relacionamento", "viabilidade"]].mean(axis=1)
        cn = cn.sort_values("media", ascending=False)
        pdf.tabela(pd.DataFrame({"Canal": cn["canal"], "Atratividade": cn["atratividade"].map(lambda v: num_br(v, 0)),
                                 "Alcance": cn["alcance"].map(lambda v: num_br(v, 0)),
                                 "Cliente": cn["cliente"].map(lambda v: num_br(v, 0)),
                                 "Relacionamento": cn["relacionamento"].map(lambda v: num_br(v, 0)),
                                 "Viabilidade": cn["viabilidade"].map(lambda v: num_br(v, 0)),
                                 "Média": cn["media"].map(lambda v: num_br(v, 1))}),
                   larguras=(52, 22, 18, 18, 26, 22, 22), alinhar=("LEFT",) + ("CENTER",) * 6)
        pdf.campo("Canais de tração", ", ".join(cn["canal"].head(3)))
    fu = extras["funil"]
    pdf.subtitulo("Funil de vendas")
    pdf.tabela(fu, alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"),
               larguras=(22, 50, 22, 20, 20, 22, 24), tam=7.5)

    # ---------------- 9. INVESTIMENTOS ----------------
    # ---------------- MARKETING E OPERAÇÃO ----------------
    if any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
        pdf.secao("PLANO DE MARKETING (4Ps)")
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            if proj.get(k):
                pdf.subtitulo(rot)
                pdf.paragrafo(proj[k])
    pdf.secao("PLANO OPERACIONAL", nova_pagina=True)
    for rot, k in [("Localização", "loc_endereco"), ("Justificativa do ponto", "loc_justificativa"),
                   ("Fluxo de pessoas", "loc_fluxo"), ("Horário de funcionamento", "op_horario")]:
        pdf.campo(rot, proj.get(k, ""))
    if proj.get("loc_area_m2"):
        pdf.campo("Área", f"{num_br(proj['loc_area_m2'], 0)} m²")
    if proj.get("op_layout"):
        pdf.subtitulo("Layout")
        pdf.paragrafo(proj["op_layout"])
    if proj.get("op_processos"):
        pdf.subtitulo("Processos operacionais")
        pdf.paragrafo(proj["op_processos"])
    cd = capacidade_demanda(d, res)
    if not cd.empty:
        pdf.subtitulo("Capacidade produtiva x demanda prevista")
        pdf.tabela(pd.DataFrame({"Recurso": cd["Recurso"], "Atende": cd["Atende"],
                                 "Capacidade/mês": cd["Capacidade/mês"].map(lambda v: num_br(v, 0)),
                                 "Demanda Mês 1": cd["Demanda Mês 1"].map(lambda v: num_br(v, 0)),
                                 "Demanda Mês 12": cd["Demanda Mês 12"].map(lambda v: num_br(v, 0)),
                                 "Uso Mês 12": cd["Uso Mês 12"].map(pct)}), larguras=(48, 42, 24, 22, 22, 22), tam=7.5)
    pes = d["pessoal"]
    if not pes.empty:
        pdf.subtitulo("Necessidade de pessoal por área")
        pdf.tabela(pd.DataFrame({"Área": pes["area"], "Cargo": pes["cargo"], "Nº": pes["num_empregados"].map(lambda v: num_br(v, 0)),
                                 "Atribuições": pes["atribuicoes"]}), larguras=(38, 52, 10, 80),
                   alinhar=("LEFT", "LEFT", "CENTER", "LEFT"), tam=7.5)
    fo = d.get("fornecedores")
    if fo is not None and not fo.empty:
        pdf.subtitulo("Fornecedores")
        pdf.tabela(pd.DataFrame({"Fornecedor": fo["nome"], "Itens": fo["itens"],
                                 "Pagamento": fo["prazo_pagamento_dias"].map(lambda v: f"{num_br(v, 0)} dias"),
                                 "Entrega": fo["prazo_entrega_dias"].map(lambda v: f"{num_br(v, 0)} dias"),
                                 "Condições": fo["condicoes"], "Crítico": fo["critico"].map(lambda v: "Sim" if v else "Não")}),
                   larguras=(36, 46, 20, 18, 46, 14), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "LEFT", "CENTER"), tam=7.5)

    pdf.secao("INVESTIMENTOS, CAPITAL DE GIRO E FONTES DE RECURSOS", nova_pagina=True)
    inv = res["inv"]
    pdf.subtitulo("Investimentos fixos e depreciação")
    tinv = pd.DataFrame({"Descrição": inv["descricao"], "Qtd.": inv["quantidade"].map(lambda v: num_br(v, 0)),
                         "Valor unit.": inv["valor_unitario"].map(brl), "Total": inv["total"].map(brl),
                         "Vida útil": inv["vida_util_anos"].map(lambda v: f"{num_br(v, 0)} anos"),
                         "Deprec. anual": inv["dep_anual"].map(brl)})
    tinv.loc[len(tinv)] = ["Total", "", "", brl(res["inv_fixo"]), "", brl(res["dep_anual"])]
    pdf.tabela(tinv, larguras=(62, 12, 26, 28, 20, 32), destaque_total=True, tam=7.5)
    pdf.subtitulo("Investimentos pré-operacionais")
    tp = pd.DataFrame({"Descrição": d["inv_pre"]["descricao"], "Valor": d["inv_pre"]["valor"].map(brl)})
    tp.loc[len(tp)] = ["Total", brl(res["inv_pre"])]
    pdf.tabela(tp, larguras=(140, 40), destaque_total=True)
    pdf.subtitulo("Estoque inicial")
    m = d["materiais"]
    m = m[m["qtd_estoque_inicial"] > 0]
    rv = res["cat"][(res["cat"]["categoria"] == "Revenda") & (res["cat"]["qtd_estoque_inicial"] > 0)]
    te = pd.DataFrame({"Item": list(m["nome"]) + list(rv["descricao"]),
                       "Tipo": ["Material/insumo"] * len(m) + ["Mercadoria p/ revenda"] * len(rv),
                       "Qtd.": [num_br(v, 1) for v in list(m["qtd_estoque_inicial"]) + list(rv["qtd_estoque_inicial"])],
                       "Valor unit.": [brl(v) for v in list(m["valor_unitario"]) + list(rv["custo_unit"])],
                       "Total": [brl(a * b) for a, b in zip(list(m["qtd_estoque_inicial"]) + list(rv["qtd_estoque_inicial"]),
                                                            list(m["valor_unitario"]) + list(rv["custo_unit"]))]})
    te.loc[len(te)] = ["Total", "", "", "", brl(res["inv_estoque"])]
    pdf.tabela(te, larguras=(70, 36, 18, 26, 30), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT"),
               destaque_total=True, tam=7.5)
    pdf.subtitulo("Capital de giro")
    pdf.tabela(pd.DataFrame({
        "Item": ["Prazo médio de recebimento (PMR)", "Prazo médio de pagamento (PMP)", "Prazo médio de estoques (PME)",
                 "Ciclo operacional (PME + PMR)", "Ciclo financeiro (PME + PMR - PMP)",
                 "Desembolso médio mensal (Mês 1)", "Necessidade de capital de giro (NCG)",
                 "Reserva financeira", "Capital de giro total"],
        "Valor": [f"{num_br(res['pmr'], 1)} dias", f"{num_br(res['pmp'], 1)} dias", f"{num_br(res['pme'], 1)} dias",
                  f"{num_br(res['ciclo_op'], 1)} dias", f"{num_br(res['ciclo_fin'], 1)} dias",
                  brl(res["desembolso_m1"]), brl(res["ncg"]), brl(res["reserva"]), brl(res["capital_giro"])]}),
        larguras=(130, 50), destaque_total=True)
    pdf.subtitulo("Investimento total e fontes de recursos")
    it = res["investimento_total"]
    pdf.tabela(pd.DataFrame({
        "Aplicação de recursos": ["Investimentos fixos", "Investimentos pré-operacionais", "Estoque inicial",
                                  "Capital de giro", "Investimento total"],
        "Valor": [brl(res["inv_fixo"]), brl(res["inv_pre"]), brl(res["inv_estoque"]), brl(res["capital_giro"]), brl(it)],
        "%": [pct(div(v, it)) for v in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)]}),
        larguras=(110, 45, 25), destaque_total=True)
    fontes = [(r["nome"], r["capital_integralizado"], "Próprio") for _, r in d["socios"].iterrows()] + \
             [(r["descricao"], r["valor"], f"Terceiros ({num_br(r['taxa_anual_pct'])}% a.a.; {int(r['prazo_meses'])} meses)")
              for _, r in d["emprestimos"].iterrows()]
    tf = pd.DataFrame({"Fonte": [f[0] for f in fontes], "Tipo": [f[2] for f in fontes],
                       "Valor": [brl(f[1]) for f in fontes], "%": [pct(div(f[1], res["fontes"])) for f in fontes]})
    tf.loc[len(tf)] = ["Total das fontes", "", brl(res["fontes"]), "100,00%"]
    pdf.tabela(tf, larguras=(65, 60, 35, 20), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT"), destaque_total=True)

    # ---------------- 10. PRODUTOS ----------------
    pdf.secao("PRODUTOS, SERVIÇOS E CUSTOS UNITÁRIOS", nova_pagina=True)
    c = res["cat"]
    for categoria in ("Fabricação Própria", "Revenda", "Serviço"):
        sub = c[c["categoria"] == categoria]
        if sub.empty:
            continue
        pdf.subtitulo(categoria)
        tc = pd.DataFrame({"Produto/serviço": sub["descricao"], "Preço": sub["preco_venda"].map(brl),
                           "Custo unit.": sub["custo_unit"].map(brl), "Margem": sub["margem_pct"].map(pct),
                           "Markup": sub["markup"].map(lambda v: f"{num_br(v)}x" if v == v else "n.a."),
                           "Qtd./mês": sub["qtd_mensal"].map(lambda v: num_br(v, 0)),
                           "Receita/mês": sub["receita_mensal"].map(brl)})
        pdf.tabela(tc, larguras=(58, 20, 20, 17, 15, 17, 33), tam=7.5)
    pdf.subtitulo("Fichas técnicas (composição do custo dos produtos)")
    ft = d["ficha_tecnica"]
    precos = dict(zip(d["materiais"]["nome"], d["materiais"]["valor_unitario"]))
    unids = dict(zip(d["materiais"]["nome"], d["materiais"]["unidade"]))
    for prod, g in ft.groupby("produto", sort=False):
        preco_v = c.loc[c["descricao"] == prod, "preco_venda"]
        rot = f"{prod}" + (f" - preço de venda {brl(preco_v.iloc[0])}" if len(preco_v) else " (produto não está no catálogo)")
        if pdf.get_y() > 240:
            pdf.add_page()
        pdf.fonte("B", 8.5)
        pdf.cell(0, 5, pdf.t(rot), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        vt = [q * precos.get(mm, 0) for mm, q in zip(g["material"], g["quantidade"])]
        tot = sum(vt)
        tf = pd.DataFrame({"Insumo": g["material"], "Quant.": g["quantidade"].map(lambda v: num_br(v, 3)),
                           "Unid.": [unids.get(mm, "?") for mm in g["material"]],
                           "Valor unit.": [brl(precos.get(mm, 0)) for mm in g["material"]],
                           "Total": [brl(v) for v in vt], "%": [pct(div(v, tot)) for v in vt]})
        tf.loc[len(tf)] = ["Custo unitário", "", "", "", brl(tot), "100,00%"]
        pdf.tabela(tf, larguras=(70, 20, 18, 26, 26, 20), destaque_total=True, tam=7.5)

    # ---------------- 11. CUSTOS ----------------
    pdf.secao("CUSTOS DE COMERCIALIZAÇÃO, PESSOAL E CUSTOS FIXOS", nova_pagina=True)
    a1 = res["anos"][1]
    pdf.subtitulo("Custos variáveis de comercialização (% sobre a receita)")
    cv = d["custos_variaveis"]
    tcv = pd.DataFrame({"Descrição": [f"Impostos sobre vendas ({proj['enquadramento']})"] + list(cv["descricao"]),
                        "%": [num_br(proj["aliquota_imposto_pct"])] + [num_br(v) for v in cv["percentual"]],
                        "Valor no Mês 1": [brl(res["receita_m1"] * res["aliq"])] +
                                          [brl(res["receita_m1"] * v / 100) for v in cv["percentual"]]})
    tcv.loc[len(tcv)] = ["Total", num_br((res["aliq"] + res["var_pct"]) * 100),
                         brl(res["receita_m1"] * (res["aliq"] + res["var_pct"]))]
    pdf.tabela(tcv, larguras=(110, 25, 45), destaque_total=True)
    pdf.subtitulo("Mão de obra")
    pe = d["pessoal"]
    sal = pe["num_empregados"] * pe["salario_mensal"]
    tpe = pd.DataFrame({"Cargo": pe["cargo"], "Nº": pe["num_empregados"].map(lambda v: num_br(v, 0)),
                        "Salário": pe["salario_mensal"].map(brl), "Encargos %": pe["encargos_pct"].map(num_br),
                        "Benefícios/func.": pe["beneficios_unit"].map(brl),
                        "Total mensal": (sal * (1 + pe["encargos_pct"] / 100) + pe["num_empregados"] * pe["beneficios_unit"]).map(brl)})
    tpe.loc[len(tpe)] = ["Total", num_br(pe["num_empregados"].sum(), 0), "", "", "", brl(res["pessoal_m"])]
    pdf.tabela(tpe, larguras=(62, 10, 25, 20, 28, 35), destaque_total=True, tam=7.5)
    pdf.subtitulo("Custos fixos mensais")
    tcf = pd.DataFrame({"Descrição": list(d["custos_fixos"]["descricao"]) +
                        ["Salários", "Encargos sociais", "Benefícios", "Depreciação"],
                        "Valor mensal": [brl(v) for v in d["custos_fixos"]["valor_mensal"]] +
                        [brl(res["salarios_m"]), brl(res["encargos_m"]), brl(res["beneficios_m"]), brl(res["dep_m"])]})
    tot_cf = res["fixos_m"] + res["pessoal_m"] + res["dep_m"]
    tcf.loc[len(tcf)] = ["Total", brl(tot_cf)]
    pdf.tabela(tcf, larguras=(130, 50), destaque_total=True)

    # ---------------- 12. PROJEÇÕES E DRE ----------------
    pdf.subtitulo("Classificação dos custos e despesas (mensal)")
    pdf.tabela(pd.DataFrame({
        "Classificação": ["Custos diretos (materiais e mercadorias) - Mês 1", "Mão de obra direta (produção)",
                          "Custos indiretos de produção", "Depreciação", "Pessoal administrativo e comercial",
                          "Despesas administrativas", "Despesas comerciais", "Despesas financeiras"],
        "Valor mensal": [brl(v) for v in (res["cdir_m1"], res["mod_m"], res["cif_m"], res["dep_m"],
                                          res["pessoal_adm_m"] + res["pessoal_com_m"], res["desp_adm_m"],
                                          res["desp_com_m"], res["desp_fin_m"])],
        "Vai para o produto (absorção)?": ["Sim (direto)", "Sim (rateio)", "Sim (rateio)", "Sim (rateio)",
                                           "Não (despesa do período)", "Não", "Não", "Não"]}),
        larguras=(90, 40, 50), alinhar=("LEFT", "RIGHT", "LEFT"))
    pdf.paragrafo(f"Critério de rateio dos custos indiretos: {res['criterio_rateio']}.", "I", 8)

    pdf.secao("FORMAÇÃO DE PREÇO E PONTO DE EQUILÍBRIO POR PRODUTO", nova_pagina=True)
    pdf.paragrafo(f"Markup divisor = 1 / [1 - (impostos {pct(res['aliq'])} + despesas variáveis {pct(res['var_pct'])} "
                  f"+ despesas fixas {pct(res['df_pct'])} + margem desejada {pct(res['margem_desejada'])})] = "
                  f"{num_br(res['markup_mult'], 3)}. Preço sugerido = custo direto x markup.", "", 9)
    cp = res["cat"]
    pdf.tabela(pd.DataFrame({"Produto/serviço": cp["descricao"], "Custo variável": cp["custo_unit"].map(brl),
                             "Custo absorção": cp["custo_absorcao_unit"].map(brl), "Preço": cp["preco_venda"].map(brl),
                             "Sugerido": cp["preco_sugerido"].map(brl), "Margem abs.": cp["margem_absorcao_pct"].map(pct),
                             "PE un./mês": cp["pe_unidades_mes"].map(lambda v: num_br(v, 0))}),
               larguras=(56, 21, 21, 20, 20, 21, 21), tam=7)

    pdf.secao("PROJEÇÃO DE RESULTADOS E DRE", nova_pagina=True)
    pdf.paragrafo(f"Premissas: crescimento de {num_br(proj['crescimento_mensal_ano1_pct'])}% ao mês no Ano 1 e "
                  f"{num_br(proj['crescimento_anual_pct'])}% ao ano a partir do Ano 2; custos fixos corrigidos em "
                  f"{num_br(proj['inflacao_custos_anual_pct'])}% ao ano; TMA de {num_br(proj['tma_anual_pct'])}% ao ano.", "I", 8)
    pdf.subtitulo("Projeção mensal do Ano 1")
    mm = res["meses"].copy()
    for col in mm.columns[1:]:
        mm[col] = mm[col].map(brl)
    pdf.tabela(mm, larguras=(18, 33, 33, 33, 30, 33), tam=7.5)
    pdf.subtitulo("Demonstração do resultado do exercício (Anos 1 a 3)")
    pdf.tabela(extras["dre"], larguras=(70, 30, 30, 30, 20), tam=7.5)

    # ---------------- 13. INDICADORES ----------------
    pdf.secao("FLUXO DE CAIXA MENSAL (ANO 1)", nova_pagina=True)
    fm = res["fluxo_mensal"]
    if not fm.empty:
        tf_ = fm.copy()
        for col in tf_.columns[1:]:
            tf_[col] = tf_[col].map(brl)
        tf_ = tf_.rename(columns={"Impostos e desp. variáveis": "Impostos/var.", "Pessoal e custos fixos": "Pessoal/fixos",
                                  "Parcelas de empréstimos": "Empréstimos"})
        pdf.tabela(tf_, larguras=(14, 25, 25, 23, 22, 24, 22, 25), tam=6.5)
        pdf.paragrafo(f"Menor saldo de caixa: {brl(fm['Saldo final'].min())} ({fm.loc[fm['Saldo final'].idxmin(), 'Mês']}).", "B", 9)

    pdf.secao("INDICADORES DE VIABILIDADE", nova_pagina=True)
    sem = semaforo(res)
    pdf.subtitulo("Semáforo da saúde do negócio")
    pdf.tabela(pd.DataFrame({"Situação": [{"🟢": "Bom", "🟡": "Atenção", "🔴": "Crítico"}.get(s_["Status"], "") for s_ in sem],
                             "Indicador": [s_["Indicador"] for s_ in sem], "Valor": [s_["Valor"] for s_ in sem],
                             "Leitura": [s_["Leitura"] for s_ in sem]}), larguras=(18, 45, 27, 90),
               alinhar=("CENTER", "LEFT", "RIGHT", "LEFT"), tam=7.5)
    rs = res["resumo"]
    pdf.subtitulo("Resumo financeiro (Ano 1, padrão PNBOX)")
    pdf.tabela(pd.DataFrame({
        "Indicador": ["Receita anual", "Custos variáveis", "Margem de contribuição", "Custos fixos (com depreciação)",
                      "Resultado operacional", "Lucratividade", "Rentabilidade", "Payback simples",
                      "PE contábil", "PE financeiro", "PE econômico"],
        "Valor": [brl(rs["receita_anual"]), brl(rs["custos_variaveis"]), brl(rs["margem_contrib"]),
                  brl(rs["custos_fixos"]), brl(rs["resultado_operacional"]), pct(rs["lucratividade"]),
                  pct(rs["rentabilidade"]),
                  f"{num_br(rs['payback_meses'], 1)} meses" if rs["payback_meses"] == rs["payback_meses"] else "Não recupera",
                  brl(a1["pe_contabil"]), brl(a1["pe_financeiro"]), brl(a1["pe_economico"])]}),
        larguras=(120, 60))
    pdf.subtitulo("Avaliação do investimento (horizonte de 3 anos)")
    fl = res["fluxos"]
    pdf.tabela(pd.DataFrame({
        "Indicador": ["Fluxos de caixa do projeto (Ano 0 a 3)", "VPL à TMA", "TIR", "Índice de lucratividade",
                      "Payback descontado"],
        "Valor": [" | ".join(brl(v) for v in fl), brl(res["vpl"]),
                  pct(res["tir"]) if res["tir"] is not None else "n.a.",
                  f"{num_br(res['il'])}x" if res["il"] == res["il"] else "n.a.",
                  f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "acima de 36 meses"]}),
        larguras=(55, 125), alinhar=("LEFT", "LEFT"))
    pdf.paragrafo("Fluxo do projeto: Ano 0 = investimento total; Anos 1 a 3 = EBITDA" +
                  (" (Ano 3 inclui a recuperação do capital de giro)." if proj.get("recuperar_giro") else "."), "I", 7.5)
    pdf.subtitulo("Matriz de indicadores em 5 blocos")
    ind = extras["indicadores"]
    for bloco, g in ind.groupby("Bloco", sort=False):
        pdf.fonte("B", 8.5)
        if pdf.get_y() > 250:
            pdf.add_page()
        pdf.cell(0, 5, pdf.t(bloco), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.tabela(g[["Indicador", "Fórmula", "Ano 1", "Ano 2", "Ano 3", "Var. Ano 2/Ano 1"]],
                   larguras=(42, 46, 25, 25, 25, 17), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"), tam=7)

    # ---------------- 14. GRÁFICOS ----------------
    imgs = graficos_png(extras["comp"]) if extras.get("comp") else []
    if imgs:
        pdf.secao("GRÁFICOS", nova_pagina=True)
        roscas = imgs[:5]
        outros = imgs[5:]
        largura = (pdf.epw - 6) / 2
        for i in range(0, len(roscas), 2):
            par = roscas[i:i + 2]
            if pdf.get_y() > 175:
                pdf.add_page()
            y = pdf.get_y()
            alturas = []
            for j, (tit, png) in enumerate(par):
                info = pdf.image(io.BytesIO(png), x=pdf.l_margin + j * (largura + 6), y=y, w=largura)
                alturas.append(info.rendered_height)
            pdf.set_y(y + max(alturas) + 4)
        for tit, png in outros:
            info_h = 95
            if pdf.get_y() > 297 - 15 - info_h:
                pdf.add_page()
            info = pdf.image(io.BytesIO(png), x=pdf.l_margin, w=pdf.epw * 0.92)
            pdf.ln(4)

    # ---------------- 15. CENÁRIOS E CHECAGENS ----------------
    # ---------------- FERRAMENTAS COMPLEMENTARES ----------------
    rk, cr, ex = d.get("riscos"), d.get("cronograma"), d.get("experimentos")
    if any(x is not None and not x.empty for x in (rk, cr, ex)):
        pdf.secao("RISCOS, CRONOGRAMA E VALIDAÇÃO", nova_pagina=True)
        if rk is not None and not rk.empty:
            pdf.subtitulo("Riscos e plano de contingência")
            pdf.tabela(pd.DataFrame({"Risco": rk["risco"], "Prob.": rk["probabilidade"].map(lambda v: num_br(v, 0)),
                                     "Impacto": rk["impacto"].map(lambda v: num_br(v, 0)),
                                     "Nível": (rk["probabilidade"] * rk["impacto"]).map(lambda v: num_br(v, 0)),
                                     "Mitigação": rk["mitigacao"], "Responsável": rk["responsavel"]}),
                       larguras=(42, 12, 14, 12, 72, 28), alinhar=("LEFT", "CENTER", "CENTER", "CENTER", "LEFT", "LEFT"), tam=7.5)
        if cr is not None and not cr.empty:
            pdf.subtitulo("Cronograma de implantação")
            pdf.tabela(pd.DataFrame({"Atividade": cr["atividade"], "Início (mês)": cr["mes_inicio"].map(lambda v: num_br(v, 0)),
                                     "Duração (meses)": cr["duracao_meses"].map(lambda v: num_br(v, 1)),
                                     "Responsável": cr["responsavel"], "Custo": cr["custo"].map(brl)}),
                       larguras=(70, 22, 25, 35, 28), alinhar=("LEFT", "CENTER", "CENTER", "LEFT", "RIGHT"), tam=7.5)
        if ex is not None and not ex.empty:
            pdf.subtitulo("Quadro de experimentação")
            pdf.tabela(pd.DataFrame({"Hipótese": ex["hipotese"], "Experimento": ex["experimento"], "Métrica": ex["metrica"],
                                     "Critério": ex["criterio_sucesso"], "Resultado": ex["resultado"], "Situação": ex["status"]}),
                       larguras=(38, 38, 24, 26, 34, 20), alinhar=("LEFT",) * 6, tam=7)

    pdf.secao("SIMULADOR DE CENÁRIOS E CHECAGENS", nova_pagina=True)
    pdf.tabela(extras["cenarios"], larguras=(45, 45, 45, 45), tam=7.5)
    pdf.subtitulo("Checagens de consistência")
    ch = extras["checagens"]
    pdf.tabela(ch[["Teste", "Valor apurado", "Situação", "Observação"]], larguras=(50, 40, 20, 70),
               alinhar=("LEFT", "LEFT", "CENTER", "LEFT"), tam=7.5)

    if str(proj.get("conclusao") or "").strip():
        pdf.secao("CONCLUSÃO E PARECER FINAL", nova_pagina=True)
        for par in str(proj["conclusao"]).split("\n"):
            pdf.paragrafo(par, "", 10)

    out = pdf.output()
    return bytes(out)

# ##############################################################################
# PARTE: pnbox_extras.py
# ##############################################################################







# ------------------------------------------------------------------------------
# 1. SITUAÇÃO DE PREENCHIMENTO DE CADA FERRAMENTA
# ------------------------------------------------------------------------------
def regras_situacao(proj: dict, d: dict) -> dict:
    def tem(t, n=1):
        return len(d[t]) >= n
    sw = d["swot"]
    return {
        "projeto": (bool(proj["nome_negocio"]) and bool(proj["cidade_uf"]), tem("socios")),
        "segmentos": (tem("segmentos", 2), tem("segmentos")),
        "personas": (tem("personas") and bool(d["personas"]["desafios"].astype(str).str.strip().ne("").any()), tem("personas")),
        "jornada": (d["jornada_cliente"]["etapa"].nunique() >= 6, tem("jornada_cliente")),
        "mercado": (proj["tam_clientes"] > 0 and proj["sam_clientes"] > 0 and proj["som_pct"] > 0 and proj["gasto_anual_cliente"] > 0,
                    proj["tam_clientes"] > 0 or proj["gasto_anual_cliente"] > 0),
        "proposta": (tem("propostas"), bool(d["proposta_valor"]["segmento_alvo"])),
        "concorrencia": (tem("concorrentes") and tem("concorrencia_notas", 3), tem("concorrencia_notas")),
        "forcas": ((sw["natureza"] == "Força").any() and (sw["natureza"] == "Fraqueza").any(),
                   sw["natureza"].isin(["Força", "Fraqueza"]).any()),
        "oportunidades": ((sw["natureza"] == "Oportunidade").any() and (sw["natureza"] == "Ameaça").any(),
                          sw["natureza"].isin(["Oportunidade", "Ameaça"]).any()),
        "swot": (sw["natureza"].nunique() >= 4, len(sw) > 0),
        "marketing": (all(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")),
                      any(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao"))),
        "canais": (tem("canais") and tem("funil_vendas"), tem("canais") or tem("funil_vendas")),
        "operacional": (bool(proj["loc_endereco"]) and tem("capacidade"), bool(proj["loc_endereco"]) or tem("capacidade")),
        "fornecedores": (tem("fornecedores", 2), tem("fornecedores")),
        "investimentos": (tem("inv_fixos") and (tem("socios") or tem("emprestimos")), tem("inv_fixos")),
        "produtos": (tem("catalogo") and tem("materiais"), tem("catalogo") or tem("materiais")),
        "custos": (tem("pessoal") and tem("custos_fixos"), tem("pessoal") or tem("custos_fixos")),
        "preco": (tem("catalogo"), tem("catalogo")),
        "dre": (tem("catalogo") and tem("custos_fixos"), tem("catalogo")),
        "fluxo": (tem("catalogo") and tem("capital_giro_prazos"), tem("catalogo")),
        "indicadores": (tem("catalogo") and tem("inv_fixos"), tem("catalogo")),
        "graficos": (tem("catalogo"), tem("catalogo")),
        "simulador": (tem("catalogo"), tem("catalogo")),
        "experimentos": (tem("experimentos"), tem("experimentos")),
        "riscos": (tem("riscos", 3), tem("riscos")),
        "cronograma": (tem("cronograma", 3), tem("cronograma")),
        "realizado": (tem("realizado"), tem("realizado")),
        "sumario": (len(str(proj["sumario_executivo"])) > 200, bool(proj["sumario_executivo"])),
    }


def completude_de(proj, d) -> float:
    r = regras_situacao(proj, d)
    # "Orçado x realizado" só faz sentido depois da abertura: não entra no cálculo
    vals = [1 if c else (0.5 if i else 0) for k, (c, i) in r.items() if k != "realizado"]
    return sum(vals) / max(len(vals), 1)


# ------------------------------------------------------------------------------
# 2. SUMÁRIO EXECUTIVO AUTOMÁTICO (rascunho para o aluno revisar)
# ------------------------------------------------------------------------------
def gerar_sumario(proj: dict, d: dict, res: dict) -> str:
    a1 = res["anos"][1]
    nome = proj["nome_negocio"] or "O negócio"
    seg = d["segmentos"]
    seg_top = ""
    if not seg.empty:
        s_ = seg.assign(m=seg.iloc[:, 2:7].mean(axis=1)).sort_values("m", ascending=False)
        seg_top = s_["nome"].iloc[0]
    props = d.get("propostas")
    valores = ""
    if props is not None and not props.empty:
        valores = "; ".join(dict.fromkeys(v.strip() for x in props["valores"] for v in str(x).split(";") if v.strip()))
    cat = res["cat"]
    top = cat.sort_values("receita_mensal", ascending=False)["descricao"].head(3).tolist() if not cat.empty else []
    socios = ", ".join(d["socios"]["nome"].tolist())
    diag = pontuacao_swot(d["swot"])
    estrategia = QUADRANTES[diag["predominante"]][0] if diag["predominante"] else "não definida"
    mk = mercado(proj, res)
    tir = res["tir"]
    pb = res["payback_desc"]
    viavel = res["vpl"] > 0 and a1["resultado"] > 0
    p = []
    p.append(f"{nome} é um empreendimento do setor de {proj['setor_atuacao'] or 'atuação a definir'}, localizado em "
             f"{proj['cidade_uf'] or 'cidade a definir'}, com o objetivo de {str(proj['objetivo']).lower()}. "
             f"O enquadramento tributário previsto é o {proj['enquadramento']} e o porte, {proj['porte']}.")
    if seg_top or valores:
        p.append(f"O segmento de clientes priorizado é {seg_top or 'a definir'}. A proposta de valor se apoia em "
                 f"{valores or 'benefícios a definir'}" + (f", com destaque para {', '.join(top)}." if top else "."))
    if mk["preenchido"]:
        p.append(f"O mercado alcançável foi estimado em {num_br(mk['sam_c'], 0)} clientes ({brl(mk['sam'])} por ano), e a meta "
                 f"de conquista (SOM) em {num_br(mk['som_c'], 0)} clientes. A receita prevista para o primeiro ano representa "
                 f"{pct(mk['pct_som'])} desse SOM.")
    p.append(f"A análise SWOT indica a {estrategia.lower()} como postura predominante.")
    p.append(f"O investimento total necessário é de {brl(res['investimento_total'])}, composto por {brl(res['inv_fixo'])} em "
             f"investimentos fixos, {brl(res['inv_pre'])} em gastos pré-operacionais, {brl(res['inv_estoque'])} em estoque inicial "
             f"e {brl(res['capital_giro'])} de capital de giro. As fontes são {brl(res['capital_proprio'])} de capital próprio"
             + (f" ({socios})" if socios else "") + f" e {brl(res['capital_terceiros'])} de capital de terceiros.")
    p.append(f"Para o primeiro ano, projeta-se receita bruta de {brl(a1['rec_bruta'])}, margem de contribuição de {pct(a1['imc'])} "
             f"e resultado líquido de {brl(a1['resultado'])} (lucratividade de {pct(div(a1['resultado'], a1['rec_bruta']))}). "
             f"O ponto de equilíbrio contábil é de {brl(a1['pe_contabil'])} por ano.")
    p.append(f"Na avaliação do investimento em três anos, à TMA de {pct(res['tma'])} ao ano, o VPL é de {brl(res['vpl'])}"
             + (f", a TIR é de {pct(tir)}" if tir is not None else "")
             + (f" e o payback descontado ocorre em {num_br(pb, 1)} meses." if pb == pb else " e o investimento não se paga em 36 meses."))
    p.append("Com base nesses resultados, o negócio se mostra " + ("economicamente viável" if viavel else
             "ainda não viável nas premissas atuais, sendo necessário rever preços, custos ou investimentos") + ".")
    return "\n".join(p)


# ------------------------------------------------------------------------------
# 3. TEXTO PARA ANÁLISE COM IA (copiar e colar no ChatGPT, Gemini, Claude...)
# ------------------------------------------------------------------------------
TAREFAS_IA = {
    "indicadores": "Diagnóstico de cada indicador: situação atual, o que significa para a saúde do negócio e sinais de alerta",
    "concorrencia": "Estratégias para se diferenciar e superar os concorrentes",
    "custos": "Plano de redução de custos e de melhoria de margens",
    "preco": "Revisão da política de preços por produto",
    "marketing": "Plano de marketing e de aquisição de clientes para os primeiros 6 meses",
    "riscos": "Principais riscos do plano e como mitigá-los",
    "graficos": "Gráficos interpretativos dos principais números",
    "banca": "Simulação de banca examinadora: perguntas difíceis sobre o plano",
}


def _md_tabela(df: pd.DataFrame, max_linhas=40) -> str:
    if df is None or df.empty:
        return "(sem dados)"
    df = df.head(max_linhas).copy()
    cols = list(df.columns)
    linhas = ["| " + " | ".join(map(str, cols)) + " |", "|" + "---|" * len(cols)]
    for _, r in df.iterrows():
        linhas.append("| " + " | ".join(str(r[c]).replace("|", "/").replace("\n", " ") for c in cols) + " |")
    return "\n".join(linhas)


def gerar_prompt_ia(proj: dict, d: dict, res: dict, tarefas: list, publico="estudante de Administração",
                    piv_conc: pd.DataFrame | None = None) -> str:
    a1 = res["anos"][1]
    sem = semaforo(res)
    cat = res["cat"]
    chk = tabela_checagens_v5(res, d, proj)
    mk = mercado(proj, res)
    diag = pontuacao_swot(d["swot"])
    b = []
    b.append(f"Você é um consultor de negócios e professor de finanças empresariais. Analise o plano de negócio abaixo, "
             f"elaborado por um {publico}, com linguagem clara, didática e em português do Brasil.")
    b.append("REGRAS: use SOMENTE os números fornecidos (não invente valores); quando fizer contas, mostre o cálculo; "
             "aponte dados incoerentes ou faltantes; separe fatos dos dados de suas sugestões; seja específico para este negócio.")
    b.append("\n## 1. Identificação")
    b.append(f"- Negócio: {proj['nome_negocio']} | Cidade: {proj['cidade_uf']} | Setor: {proj['setor_atuacao']}")
    b.append(f"- Enquadramento: {proj['enquadramento']} | Porte: {proj['porte']} | Objetivo: {proj['objetivo']}")
    seg = d["segmentos"]
    if not seg.empty:
        s_ = seg.assign(Média=seg.iloc[:, 2:7].mean(axis=1).round(1)).sort_values("Média", ascending=False)
        b.append("\n## 2. Segmentos de clientes (notas 1 a 10)")
        b.append(_md_tabela(s_[["nome", "descricao", "Média"]].rename(columns={"nome": "Segmento", "descricao": "Descrição"})))
    props = d.get("propostas")
    if props is not None and not props.empty:
        b.append("\n## 3. Propostas de valor")
        for _, p_ in props.iterrows():
            b.append(f"- Persona {p_['persona_nome']}: deseja {p_['que_desejam']}; por meio de {p_['por_meio']}; "
                     f"diferente de {p_['diferentemente']}; oferecemos {p_['oferecem']}. Valores: {p_['valores']}")
    if mk["preenchido"]:
        b.append("\n## 4. Mercado")
        b.append(f"- TAM {num_br(mk['tam_c'], 0)} clientes ({brl(mk['tam'])}/ano); SAM {num_br(mk['sam_c'], 0)} ({brl(mk['sam'])}); "
                 f"SOM {num_br(mk['som_c'], 0)} ({brl(mk['som'])}). Receita do Ano 1 = {pct(mk['pct_som'])} do SOM.")
    if piv_conc is not None and not piv_conc.empty:
        b.append("\n## 5. Matriz de concorrência (notas 0 a 10)")
        b.append(_md_tabela(piv_conc.round(1)))
    sw = d["swot"]
    if not sw.empty:
        b.append("\n## 6. SWOT")
        for nat in ("Força", "Fraqueza", "Oportunidade", "Ameaça"):
            itens = sw[sw["natureza"] == nat]
            if not itens.empty:
                b.append(f"- {nat}s: " + "; ".join(f"{r['fator']} ({r['impacto']})" for _, r in itens.iterrows()))
        if diag["predominante"]:
            b.append(f"- Quadrante predominante: {QUADRANTES[diag['predominante']][0]}")
    if any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
        b.append("\n## 7. Marketing (4Ps)")
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            b.append(f"- {rot}: {proj.get(k, '')}")
    b.append("\n## 8. Investimento e fontes")
    b.append(f"- Investimento total {brl(res['investimento_total'])}: fixos {brl(res['inv_fixo'])}; pré-operacionais {brl(res['inv_pre'])}; "
             f"estoque {brl(res['inv_estoque'])}; capital de giro {brl(res['capital_giro'])}.")
    b.append(f"- Fontes: próprio {brl(res['capital_proprio'])}; terceiros {brl(res['capital_terceiros'])}.")
    if not cat.empty:
        b.append("\n## 9. Produtos (Mês 1)")
        t = pd.DataFrame({"Produto": cat["descricao"], "Categoria": cat["categoria"], "Preço": cat["preco_venda"].map(brl),
                          "Custo variável": cat["custo_unit"].map(brl), "Custo absorção": cat["custo_absorcao_unit"].map(brl),
                          "Margem": cat["margem_pct"].map(pct), "Qtd/mês": cat["qtd_mensal"].map(lambda v: num_br(v, 0)),
                          "Receita/mês": cat["receita_mensal"].map(brl)})
        b.append(_md_tabela(t, 60))
    b.append("\n## 10. Estrutura de custos mensal")
    b.append(f"- Custos diretos {brl(res['cdir_m1'])}; impostos {pct(res['aliq'])} e despesas variáveis {pct(res['var_pct'])} da receita; "
             f"MOD {brl(res['mod_m'])}; custos indiretos de produção {brl(res['cif_m'])}; depreciação {brl(res['dep_m'])}; "
             f"pessoal adm./comercial {brl(res['pessoal_adm_m'] + res['pessoal_com_m'])}; despesas adm. {brl(res['desp_adm_m'])}, "
             f"comerciais {brl(res['desp_com_m'])}, financeiras {brl(res['desp_fin_m'])}.")
    cf = d["custos_fixos"]
    if not cf.empty:
        b.append(_md_tabela(pd.DataFrame({"Custo fixo": cf["descricao"], "Classificação": cf["classificacao"],
                                          "R$/mês": cf["valor_mensal"].map(brl)})))
    b.append("\n## 11. DRE (Anos 1 a 3)")
    b.append(_md_tabela(tabela_dre(res)))
    b.append("\n## 12. Indicadores e semáforo")
    b.append(_md_tabela(pd.DataFrame({"Indicador": [s_["Indicador"] for s_ in sem], "Valor": [s_["Valor"] for s_ in sem],
                                      "Situação": [{"🟢": "bom", "🟡": "atenção", "🔴": "crítico"}[s_["Status"]] for s_ in sem]})))
    b.append(f"- VPL {brl(res['vpl'])}; TIR {pct(res['tir']) if res['tir'] is not None else 'n.a.'}; TMA {pct(res['tma'])}; "
             f"payback descontado {num_br(res['payback_desc'], 1) + ' meses' if res['payback_desc'] == res['payback_desc'] else '> 36 meses'}; "
             f"PE contábil {brl(a1['pe_contabil'])}/ano; margem de segurança {pct(div(a1['rec_bruta'] - a1['pe_contabil'], a1['rec_bruta']))}.")
    sens = sensibilidade(res)
    b.append("- Sensibilidade (efeito de ±10% no resultado do Ano 1): " +
             "; ".join(f"{r['Variável']} {brl(r['Impacto absoluto'])}" for _, r in sens.iterrows()))
    pend = chk[chk["Situação"] != "OK"]
    if not pend.empty:
        b.append("\n## 13. Inconsistências apontadas pelo programa")
        for _, r in pend.iterrows():
            b.append(f"- {r['Teste']}: {r['Valor apurado']}")
    fm = res["fluxo_mensal"]
    if not fm.empty:
        b.append(f"\n- Fluxo de caixa: menor saldo {brl(fm['Saldo final'].min())} em {fm.loc[fm['Saldo final'].idxmin(), 'Mês']}; "
                 f"saldo no fim do Ano 1 {brl(fm['Saldo final'].iloc[-1])}.")

    b.append("\n# O QUE EU PRECISO QUE VOCÊ FAÇA")
    pedidos = {
        "indicadores": "Para CADA indicador da seção 12, explique em linguagem simples: (a) o que ele mede; (b) a situação atual deste "
                       "negócio com o valor exato; (c) o que isso significa para a saúde do negócio; (d) uma ação concreta para melhorar. "
                       "Termine com um parecer geral (viável, viável com ressalvas ou inviável) e justifique.",
        "concorrencia": "Com base na concorrência, na SWOT e nas propostas de valor, proponha 5 estratégias para driblar os concorrentes, "
                        "cada uma com: objetivo, ação prática, custo estimado (baixo/médio/alto), prazo e indicador de sucesso.",
        "custos": "Identifique os 5 itens de custo que mais pesam e proponha como reduzi-los sem prejudicar a qualidade. Para cada um, "
                  "estime o efeito no resultado anual em R$ usando os números fornecidos. Aponte também produtos com margem baixa.",
        "preco": "Analise preço x custo variável x custo por absorção de cada produto. Diga quais preços estão baixos ou altos demais, "
                 "sugira ajustes e explique o raciocínio (markup, concorrência, percepção de valor).",
        "marketing": "Monte um plano de marketing de 6 meses (mês a mês), com canais, ações, orçamento estimado e metas, coerente com a "
                     "persona e a proposta de valor.",
        "riscos": "Liste os 5 maiores riscos que você enxerga neste plano (financeiros, de mercado e operacionais), a probabilidade, o "
                  "impacto e uma ação de prevenção para cada um.",
        "graficos": "Gere gráficos interpretativos (se puder executar código, use Python/matplotlib; se não, descreva o gráfico e monte "
                    "uma tabela pronta para virar gráfico no Excel): (1) composição do investimento; (2) estrutura de custos; "
                    "(3) receita x custos x resultado nos 3 anos; (4) ponto de equilíbrio; (5) saldo de caixa mês a mês. "
                    "Abaixo de cada gráfico, escreva 2 ou 3 frases de interpretação.",
        "banca": "Faça o papel de uma banca examinadora: elabore 10 perguntas difíceis sobre este plano (mercado, números, premissas, "
                 "riscos) e, para cada uma, indique o que uma boa resposta deveria conter.",
    }
    for i, t in enumerate([t for t in tarefas if t in pedidos], start=1):
        b.append(f"{i}. {pedidos[t]}")
    b.append("\nOrganize a resposta com títulos, use tabelas quando ajudar e termine com uma lista de próximos passos.")
    return "\n".join(b)


# ------------------------------------------------------------------------------
# 4. PITCH DECK EM PDF (slides 16:9)
# ------------------------------------------------------------------------------
class Pitch(Relatorio):
    def __init__(self, nome):
        FPDF.__init__(self, orientation="L", unit="mm", format=(190.5, 338.7))
        self.nome = nome
        self.unicode = False
        for reg, neg, ita in FONTES_CANDIDATAS:
            if all(Path(p).exists() for p in (reg, neg, ita)):
                self.add_font("Base", "", reg)
                self.add_font("Base", "B", neg)
                self.add_font("Base", "I", ita)
                self.unicode = True
                break
        self.fam = "Base" if self.unicode else "Helvetica"
        self.set_auto_page_break(False)
        self.set_margins(18, 14, 18)

    def footer(self):
        self.set_y(-10)
        self.fonte("", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 5, self.t(f"{self.nome}  ·  {self.page_no()}"), align="R")
        self.set_text_color(0, 0, 0)

    def slide(self, titulo, sub=""):
        self.add_page()
        self.set_fill_color(*AZUL)
        self.rect(0, 0, self.w, 22, style="F")
        self.set_xy(18, 6)
        self.fonte("B", 20)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, self.t(titulo))
        self.set_text_color(0, 0, 0)
        self.set_xy(18, 28)
        if sub:
            self.fonte("I", 12)
            self.set_text_color(90, 90, 90)
            self.multi_cell(0, 6, self.t(sub), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.set_text_color(0, 0, 0)
            self.ln(2)

    def topicos(self, itens, tam=15):
        self.fonte("", tam)
        for it in itens:
            if not str(it).strip():
                continue
            self.set_x(self.l_margin)
            self.multi_cell(0, tam * 0.5, self.t(f"•  {it}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.ln(2)

    def numeros(self, pares):
        n = len(pares)
        w = (self.w - 36 - (n - 1) * 6) / n
        y = self.get_y() + 4
        for i, (rot, val) in enumerate(pares):
            x = 18 + i * (w + 6)
            self.set_fill_color(232, 240, 251)
            self.rect(x, y, w, 34, style="F")
            self.set_xy(x, y + 6)
            self.fonte("B", 20)
            self.set_text_color(*AZUL)
            self.cell(w, 10, self.t(val), align="C")
            self.set_xy(x, y + 20)
            self.fonte("", 11)
            self.set_text_color(60, 60, 60)
            self.cell(w, 6, self.t(rot), align="C")
        self.set_text_color(0, 0, 0)
        self.set_y(y + 40)


def gerar_pitch(proj: dict, d: dict, res: dict) -> bytes:
    a1 = res["anos"][1]
    pp = Pitch(proj["nome_negocio"])
    # 1. Capa
    pp.add_page()
    pp.set_fill_color(*AZUL)
    pp.rect(0, 0, pp.w, pp.h, style="F")
    pp.set_text_color(255, 255, 255)
    pp.set_xy(18, 60)
    pp.fonte("B", 40)
    pp.multi_cell(0, 18, pp.t(proj["nome_negocio"]), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.fonte("", 16)
    props = d.get("propostas")
    frase = props["oferecem"].iloc[0] if props is not None and not props.empty and props["oferecem"].iloc[0] else "Plano de negócio"
    pp.multi_cell(0, 9, pp.t(frase), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.ln(10)
    pp.fonte("", 12)
    pp.multi_cell(0, 7, pp.t(" · ".join(x for x in [proj["cidade_uf"], proj.get("turma", ""), proj.get("autores", "").replace("\n", ", ")] if x)),
                  align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.set_text_color(0, 0, 0)
    # 2. Problema
    pers = d["personas"]
    pp.slide("O problema", "Quem é o cliente e qual dor ele sente")
    itens = []
    for _, p_ in pers.head(2).iterrows():
        itens.append(f"{p_['nome']}, {int(p_['idade'])} anos, {p_['profissao']}: \"{p_['citacao']}\"")
        if p_["desafios"]:
            itens.append(f"Dores: {p_['desafios']}")
    pp.topicos(itens or ["Descreva as personas no programa para preencher este slide."])
    # 3. Solução
    pp.slide("A solução", "Proposta de valor")
    itens = [f"Para {p_['persona_nome']}: {p_['por_meio']} — oferecemos {p_['oferecem']}" for _, p_ in props.iterrows()] \
        if props is not None and not props.empty else []
    top = res["cat"].sort_values("receita_mensal", ascending=False)["descricao"].head(5).tolist()
    if top:
        itens.append("Principais produtos: " + ", ".join(top))
    pp.topicos(itens)
    # 4. Mercado
    mk = mercado(proj, res)
    pp.slide("Mercado", proj.get("sam_desc") or "Tamanho do mercado")
    if mk["preenchido"]:
        pp.numeros([("TAM - mercado total", brl(mk["tam"])), ("SAM - alcançável", brl(mk["sam"])), ("SOM - meta", brl(mk["som"]))])
        pp.topicos([f"{num_br(mk['som_c'], 0)} clientes conquistáveis; a receita do Ano 1 equivale a {pct(mk['pct_som'])} do SOM."], 13)
    else:
        pp.topicos(["Preencha a ferramenta Tamanho do mercado para completar este slide."])
    # 5. Concorrência
    pp.slide("Concorrência e diferencial", "Como nos posicionamos")
    notas = d["concorrencia_notas"]
    if not notas.empty:
        med = notas.groupby("empresa")["nota"].mean().sort_values(ascending=False)
        pp.topicos([f"{e}: média {num_br(v, 1)}" for e, v in med.items()] +
                   [f"Diferente de: {props['diferentemente'].iloc[0]}" if props is not None and not props.empty else ""], 14)
    # 6. Modelo de negócio / marketing
    pp.slide("Como vamos vender", "Plano de marketing (4Ps)")
    pp.topicos([f"Produto: {proj['mkt_produto']}", f"Preço: {proj['mkt_preco']}", f"Praça: {proj['mkt_praca']}",
                f"Promoção: {proj['mkt_promocao']}"], 13)
    # 7. Operação
    pp.slide("Operação", proj.get("loc_endereco") or "")
    cd = capacidade_demanda(d, res)
    itens = [proj.get("op_processos", "")] + [f"{r['Recurso']}: uso de {pct(r['Uso Mês 12'])} da capacidade no Mês 12"
                                             for _, r in cd.iterrows()]
    pp.topicos(itens, 13)
    # 8. Números
    pp.slide("Os números", "Projeção do primeiro ano")
    pp.numeros([("Investimento total", brl(res["investimento_total"])), ("Receita - Ano 1", brl(a1["rec_bruta"])),
                ("Lucro - Ano 1", brl(a1["resultado"])), ("Lucratividade", pct(div(a1["resultado"], a1["rec_bruta"])))])
    pp.numeros([("VPL (3 anos)", brl(res["vpl"])), ("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a."),
                ("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "> 36 meses"),
                ("Ponto de equilíbrio", brl(a1["pe_contabil"]))])
    # 9. Gráficos
    imgs = graficos_png(composicoes(res, d))
    escolhidos = [i for i in imgs if i[0] in ("Composição do investimento total", "Ponto de equilíbrio")]
    if escolhidos:
        pp.slide("Investimento e equilíbrio")
        x = 18
        for tit, png in escolhidos[:2]:
            pp.image(io.BytesIO(png), x=x, y=30, h=140)
            x += 160
    # 10. Investimento pedido
    pp.slide("O que precisamos", "Uso e fontes dos recursos")
    pp.numeros([("Capital próprio", brl(res["capital_proprio"])), ("Capital de terceiros", brl(res["capital_terceiros"])),
                ("Total", brl(res["fontes"]))])
    pp.topicos([f"Investimentos fixos: {brl(res['inv_fixo'])}", f"Pré-operacionais: {brl(res['inv_pre'])}",
                f"Estoque inicial: {brl(res['inv_estoque'])}", f"Capital de giro: {brl(res['capital_giro'])}"], 14)
    # 11. Equipe
    pp.slide("Equipe", "Quem faz acontecer")
    pp.topicos([f"{r['nome']} — {r['funcao']}: {r['perfil']}" for _, r in d["socios"].iterrows()], 13)
    # 12. Encerramento
    pp.add_page()
    pp.set_fill_color(*AZUL)
    pp.rect(0, 0, pp.w, pp.h, style="F")
    pp.set_text_color(255, 255, 255)
    pp.set_xy(18, 75)
    pp.fonte("B", 34)
    pp.cell(0, 16, pp.t("Obrigado!"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.fonte("", 14)
    pp.cell(0, 10, pp.t(proj.get("contatos") or ""), align="C")
    return bytes(pp.output())


# ------------------------------------------------------------------------------
# 5. EXPORTAÇÃO EM WORD (.docx)
# ------------------------------------------------------------------------------
def gerar_docx(proj: dict, d: dict, res: dict, piv_conc=None) -> bytes | None:
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        return None
    doc = Document()
    st_ = doc.styles["Normal"]
    st_.font.name = "Arial"
    st_.font.size = Pt(11)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.5)
        s.top_margin = s.bottom_margin = Cm(2.5)

    def tabela(df: pd.DataFrame):
        if df is None or df.empty:
            doc.add_paragraph("(sem registros)")
            return
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = "Light Grid Accent 1"
        for i, c in enumerate(df.columns):
            t.rows[0].cells[i].text = str(c)
        for _, r in df.iterrows():
            cells = t.add_row().cells
            for i, c in enumerate(df.columns):
                cells[i].text = str(r[c])
        for row in t.rows:
            for c in row.cells:
                for p_ in c.paragraphs:
                    for run in p_.runs:
                        run.font.size = Pt(8.5)
        doc.add_paragraph()

    p_ = doc.add_paragraph()
    p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = p_.add_run(proj["nome_negocio"])
    r_.bold, r_.font.size = True, Pt(26)
    r_.font.color.rgb = RGBColor(0x1C, 0x31, 0x50)
    p_ = doc.add_paragraph("Plano de Negócio e Análise de Viabilidade Econômico-Financeira")
    p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for linha in [proj["cidade_uf"], proj.get("turma", ""), proj.get("autores", ""), proj.get("professor", ""),
                  proj.get("data_entrega", "") or date.today().strftime("%d/%m/%Y")]:
        if linha:
            q = doc.add_paragraph(str(linha))
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    doc.add_heading("Sumário executivo", 1)
    for par in str(proj.get("sumario_executivo") or gerar_sumario(proj, d, res)).split("\n"):
        doc.add_paragraph(par)
    doc.add_heading("1. Análise de mercado", 1)
    doc.add_heading("1.1 Segmentos de clientes", 2)
    seg = d["segmentos"]
    if not seg.empty:
        tabela(seg.assign(Média=seg.iloc[:, 2:7].mean(axis=1).round(1))[["nome", "descricao", "Média"]]
               .rename(columns={"nome": "Segmento", "descricao": "Descrição"}))
    doc.add_heading("1.2 Personas", 2)
    for _, p in d["personas"].iterrows():
        doc.add_paragraph(f"{p['nome']}, {int(p['idade'])} anos, {p['profissao']}", style="List Bullet")
        doc.add_paragraph(f"{p['descricao']} Ganhos: {p['ganhos']} Dores: {p['desafios']}")
    mk = mercado(proj, res)
    if mk["preenchido"]:
        doc.add_heading("1.3 Tamanho do mercado", 2)
        doc.add_paragraph(f"TAM: {num_br(mk['tam_c'], 0)} clientes ({brl(mk['tam'])}/ano). SAM: {num_br(mk['sam_c'], 0)} clientes "
                          f"({brl(mk['sam'])}/ano). SOM: {num_br(mk['som_c'], 0)} clientes ({brl(mk['som'])}/ano). "
                          f"Fontes: {proj.get('mercado_fontes', '')}")
    doc.add_heading("1.4 Concorrência", 2)
    if piv_conc is not None and not piv_conc.empty:
        tabela(piv_conc.round(1))
    doc.add_heading("2. Proposta de valor", 1)
    props = d.get("propostas")
    if props is not None:
        for _, pr in props.iterrows():
            doc.add_paragraph(f"{proj['nome_negocio']} ajuda {pr['persona_nome']}, que deseja {pr['que_desejam']}, por meio de "
                              f"{pr['por_meio']}. Diferentemente de {pr['diferentemente']}, oferecemos {pr['oferecem']}.")
    doc.add_heading("3. Análise estratégica (SWOT)", 1)
    sw = d["swot"]
    tabela(sw[["ambiente", "natureza", "fator", "impacto", "plano_acao"]].rename(columns={
        "ambiente": "Ambiente", "natureza": "Classificação", "fator": "Fator", "impacto": "Impacto", "plano_acao": "Ação"}))
    diag = pontuacao_swot(sw)
    if diag["predominante"]:
        doc.add_paragraph(f"Estratégia predominante: {QUADRANTES[diag['predominante']][0]}. {QUADRANTES[diag['predominante']][2]}")
    doc.add_heading("4. Plano de marketing", 1)
    for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
        doc.add_paragraph(f"{rot}: {proj.get(k, '')}")
    doc.add_heading("5. Plano operacional", 1)
    for rot, k in [("Localização", "loc_endereco"), ("Justificativa", "loc_justificativa"), ("Layout", "op_layout"),
                   ("Processos", "op_processos"), ("Horário", "op_horario")]:
        if proj.get(k):
            doc.add_paragraph(f"{rot}: {proj[k]}")
    doc.add_heading("6. Plano financeiro", 1)
    doc.add_heading("6.1 Investimento total", 2)
    it = res["investimento_total"]
    tabela(pd.DataFrame({"Item": ["Investimentos fixos", "Pré-operacionais", "Estoque inicial", "Capital de giro", "Total"],
                         "Valor": [brl(x) for x in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)]}))
    doc.add_heading("6.2 Demonstração do resultado", 2)
    tabela(tabela_dre(res))
    doc.add_heading("6.3 Indicadores de viabilidade", 2)
    sem = semaforo(res)
    tabela(pd.DataFrame({"Indicador": [s["Indicador"] for s in sem], "Valor": [s["Valor"] for s in sem],
                         "Leitura": [s["Leitura"] for s in sem]}))
    imgs = graficos_png(composicoes(res, d))
    if imgs:
        doc.add_heading("6.4 Gráficos", 2)
        for tit, png in imgs:
            doc.add_picture(io.BytesIO(png), width=Cm(14))
            doc.add_paragraph(f"Figura - {tit}. Fonte: elaborado pelos autores com o PNBOX Acadêmico.").runs[0].italic = True
    doc.add_heading("7. Conclusão", 1)
    doc.add_paragraph(str(proj.get("conclusao") or "Escreva aqui a conclusão do grupo sobre a viabilidade do negócio."))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------------------
# 6. MODO PROFESSOR: AVALIAÇÃO DE VÁRIOS PROJETOS
# ------------------------------------------------------------------------------
RUBRICA_PADRAO = [
    ("Cliente e mercado (segmentos, personas, jornada, TAM/SAM/SOM)", 20),
    ("Proposta de valor, concorrência e estratégia (SWOT)", 20),
    ("Marketing e plano operacional", 15),
    ("Plano financeiro completo e consistente", 25),
    ("Viabilidade e interpretação dos indicadores", 10),
    ("Sumário executivo e conclusão", 10),
]
GRUPOS_RUBRICA = [
    ["segmentos", "personas", "jornada", "mercado"],
    ["proposta", "concorrencia", "forcas", "oportunidades", "swot"],
    ["marketing", "canais", "operacional", "fornecedores"],
    ["investimentos", "produtos", "custos"],
    None,
    ["sumario"],
]


def carregar_projeto_arquivo(caminho) -> tuple:
    conn = conectar(caminho)
    proj = ler_unica(conn, "projeto")
    d = {t: ler(conn, t) for t in SCHEMA if t not in ("projeto", "proposta_valor")}
    d["proposta_valor"] = ler_unica(conn, "proposta_valor")
    conn.close()
    return proj, d


def avaliar(proj: dict, d: dict, rubrica=RUBRICA_PADRAO) -> dict:
    res = executar_motor(proj, d)
    chk = tabela_checagens_v5(res, d, proj)
    sem = semaforo(res)
    comp = completude_de(proj, d)
    rx = raio_x(res, chk, comp, sem)
    regras = regras_situacao(proj, d)

    def frac(chaves):
        v = [1 if regras[k][0] else (0.5 if regras[k][1] else 0) for k in chaves]
        return sum(v) / len(v)

    ok_chk = float((chk["Situação"] == "OK").mean()) if len(chk) else 0
    verdes = sum(1 for s in sem if s["Status"] == "🟢") / max(len(sem), 1)
    notas = []
    for (crit, peso), grupo in zip(rubrica, GRUPOS_RUBRICA + [None] * max(0, len(rubrica) - len(GRUPOS_RUBRICA))):
        if grupo is None and "consistente" in crit.lower():
            f = 0.5 * frac(["investimentos", "produtos", "custos"]) + 0.5 * ok_chk
        elif grupo is None:
            f = 0.5 * verdes + 0.5 * (1 if len(str(proj.get("conclusao", ""))) > 150 else 0.3)
        else:
            f = frac(grupo)
            if "consistente" in crit.lower():
                f = 0.5 * f + 0.5 * ok_chk
            if grupo == ["sumario"]:
                f = 0.6 * f + 0.4 * (1 if len(str(proj.get("conclusao", ""))) > 150 else 0)
        notas.append(round(peso * f, 1))
    a1 = res["anos"][1]
    return {"Projeto": proj["nome_negocio"], "Integrantes": str(proj.get("autores", "")).replace("\n", "; "),
            "Turma": proj.get("turma", ""), "Preenchimento": comp, "Checagens OK": ok_chk, "Raio-X": rx["nota"],
            "Investimento": res["investimento_total"], "Receita Ano 1": a1["rec_bruta"], "Lucro Ano 1": a1["resultado"],
            "VPL": res["vpl"], "TIR": res["tir"], "Payback desc. (meses)": res["payback_desc"],
            "Pendências": "; ".join(chk.loc[chk["Situação"] != "OK", "Teste"]), "notas": notas,
            "Nota sugerida": round(sum(notas), 1), "res": res}


def ficha_avaliacao_pdf(avaliacoes: list, rubrica=RUBRICA_PADRAO, professor="", disciplina="") -> bytes:
    pdf = Relatorio("Fichas de avaliação")
    pdf.alias_nb_pages()
    for av in avaliacoes:
        pdf.add_page()
        pdf.n_secao = 0
        pdf.fonte("B", 15)
        pdf.set_text_color(*AZUL)
        pdf.cell(0, 9, pdf.t("FICHA DE AVALIAÇÃO DE PLANO DE NEGÓCIO"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.fonte("", 9)
        pdf.cell(0, 5, pdf.t(f"{disciplina}  ·  Professor(a): {professor}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)
        pdf.campo("Projeto", av["Projeto"])
        pdf.campo("Integrantes", av["Integrantes"])
        pdf.campo("Turma", av["Turma"])
        pdf.campo("Indicadores", f"Investimento {brl(av['Investimento'])} · Receita Ano 1 {brl(av['Receita Ano 1'])} · "
                                 f"Lucro Ano 1 {brl(av['Lucro Ano 1'])} · VPL {brl(av['VPL'])} · "
                                 f"TIR {pct(av['TIR']) if av['TIR'] is not None else 'n.a.'}")
        pdf.campo("Preenchimento / checagens", f"{pct(av['Preenchimento'], 0)} preenchido · {pct(av['Checagens OK'], 0)} das checagens OK · "
                                               f"Raio-X {av['Raio-X']}/100")
        if av["Pendências"]:
            pdf.campo("Pendências apontadas pelo programa", av["Pendências"])
        pdf.ln(2)
        notas_fin = av.get("notas_finais") or av["notas"]
        tab = pd.DataFrame({"Critério": [c for c, _ in rubrica], "Peso": [num_br(p, 0) for _, p in rubrica],
                            "Nota sugerida": [num_br(n, 1) for n in av["notas"]],
                            "Nota da banca": [num_br(n, 1) if av.get("notas_finais") else "" for n in notas_fin]})
        tab.loc[len(tab)] = ["TOTAL", num_br(sum(p for _, p in rubrica), 0), num_br(sum(av["notas"]), 1),
                             num_br(sum(notas_fin), 1) if av.get("notas_finais") else ""]
        pdf.tabela(tab, larguras=(110, 15, 27, 28), alinhar=("LEFT", "CENTER", "CENTER", "CENTER"), destaque_total=True)
        pdf.subtitulo("Comentários da banca")
        y = pdf.get_y() + 2
        for i in range(8):
            pdf.line(pdf.l_margin, y + i * 8, pdf.w - pdf.r_margin, y + i * 8)
        pdf.set_y(y + 70)
        pdf.cell(0, 6, pdf.t("Assinatura: ____________________________________    Data: ____/____/______"),
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    return bytes(pdf.output())

# ##############################################################################
# PARTE: pnbox_ui_base.py
# ##############################################################################
# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: BASE (configuração, projetos, navegação, salvamento)
# ==============================================================================





st.set_page_config(page_title="PNBOX Acadêmico - Plano de Negócio", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")
ss = st.session_state
VERSAO_PROGRAMA = "5.0"

# ------------------------------------------------------------------------------
# ENCERRAMENTO AUTOMÁTICO: quando todas as abas do navegador forem fechadas,
# o programa termina sozinho (só quando aberto pelo próprio arquivo / .bat / .exe).
# ------------------------------------------------------------------------------
TEMPO_ESPERA_FECHAR = 15


@st.cache_resource
def _vigia_de_encerramento():
    if os.environ.get("PNBOX_AUTOFECHAR") != "1":
        return False

    def vigiar():
        from streamlit.runtime import Runtime
        ja_conectou, desde = False, None
        while True:
            time.sleep(2)
            try:
                abertas = Runtime.instance()._session_mgr.num_active_sessions()
            except Exception:
                continue
            if abertas > 0:
                ja_conectou, desde = True, None
            elif ja_conectou:
                desde = desde or time.time()
                if time.time() - desde > TEMPO_ESPERA_FECHAR:
                    print("Nenhuma aba aberta. Encerrando o PNBOX...")
                    os._exit(0)

    threading.Thread(target=vigiar, daemon=True, name="pnbox-vigia").start()
    return True


_vigia_de_encerramento()

# ------------------------------------------------------------------------------
# ONDE O PROGRAMA ESTÁ RODANDO E QUAL VERSÃO ABRIR
# ------------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
NAVEGADOR = sys.platform == "emscripten"                      # GitHub Pages (stlite)
MODO_WEB = NAVEGADOR or Path("/mount/src").exists() or os.environ.get("PNBOX_WEB") == "1"


def _versao_inicial() -> str:
    """'modelo' abre com o exemplo FitBox; 'branco' abre com um projeto vazio."""
    try:
        q = st.query_params.get("versao", "")
    except Exception:
        q = ""
    if q in ("modelo", "branco"):
        return q
    if os.environ.get("PNBOX_VERSAO") in ("modelo", "branco"):
        return os.environ["PNBOX_VERSAO"]
    arq = APP_DIR / "versao.txt"
    if arq.exists():
        v = arq.read_text(encoding="utf-8").strip().lower()
        if v in ("modelo", "branco"):
            return v
    return "modelo"


VERSAO = _versao_inicial()

if NAVEGADOR:
    # Pasta guardada no próprio navegador (IndexedDB): sobrevive a atualizações do programa
    PASTA = Path("/home/pyodide/pnbox")
elif MODO_WEB:
    if "sessao_web" not in ss:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            ss.sessao_web = get_script_run_ctx().session_id
        except Exception:
            import uuid
            ss.sessao_web = uuid.uuid4().hex
    PASTA = Path(tempfile.gettempdir()) / "pnbox" / ss.sessao_web
elif (APP_DIR / "instalado.txt").exists():
    PASTA = Path.home() / "Documents" / "PNBOX"
else:
    PASTA = APP_DIR / "projetos"
PASTA.mkdir(parents=True, exist_ok=True)

for k, v in {"ver": {}, "buf": {}, "snap": {}, "hash": {}, "pagina": "inicio", "gen": 0,
             "alerta": None, "ultimo_salvo": None, "tema": "claro"}.items():
    ss.setdefault(k, v)

# ------------------------------------------------------------------------------
# LISTAS PADRÃO (o aluno pode acrescentar outras em Configurações)
# ------------------------------------------------------------------------------
LISTAS_PADRAO = {
    "unidade": ["kg", "g", "L", "ml", "unidade", "dúzia", "pacote", "maço", "caixa", "kit", "galão",
                "lata", "par", "m", "m²", "hora", "serviço"],
    "categoria_material": ["Matéria-prima", "Embalagem", "Limpeza e higiene", "Material de consumo",
                           "Mercadoria para revenda", "Outros insumos"],
    "dimensao_interna": ["Pessoas", "Marketing", "Finanças", "Operação", "Tecnologia", "Gestão"],
    "dimensao_externa": ["Político-Legal", "Econômico", "Social", "Tecnológico", "Ambiental",
                         "Demográfico", "Concorrencial"],
    "valor_proposta": ["Conforto e conveniência", "Economia de dinheiro", "Economia de tempo", "Esforço reduzido",
                       "Rapidez", "Segurança", "Variedade", "Personalização", "Qualidade", "Design e estética",
                       "Durabilidade", "Bem-estar", "Status", "Exclusividade", "Comunidade e pertencimento",
                       "Realização pessoal", "Sustentabilidade", "Confiança", "Desempenho", "Simplicidade",
                       "Aprendizado e autodesenvolvimento", "Integração de produtos"],
    "area_pessoal": ["Produção / operação", "Administrativo", "Comercial / vendas"],
    "classe_custo_fixo": ["Custo indireto de produção", "Despesa administrativa", "Despesa comercial",
                          "Despesa financeira"],
    "setor": ["Comércio", "Serviço", "Indústria", "Agronegócio", "Alimentos e bebidas", "Bares e restaurantes",
              "Varejo", "Moda", "Saúde e bem-estar", "Tecnologia", "Educação", "Turismo", "Beleza e estética"],
    "canal": ["Marketing de Conteúdo", "Publicidade em Redes Sociais", "Marketing Viral", "Comunidade",
              "Plataformas Existentes", "Acordos e Parcerias", "Eventos", "Indicação (boca a boca)",
              "Vendas Diretas", "E-mail e WhatsApp"],
}
NOMES_LISTAS = {
    "unidade": "Unidades de medida", "categoria_material": "Categorias de materiais",
    "dimensao_interna": "Dimensões internas (forças e fraquezas)",
    "dimensao_externa": "Dimensões externas (oportunidades e ameaças)",
    "valor_proposta": "Valores da proposta de valor", "area_pessoal": "Áreas da mão de obra",
    "classe_custo_fixo": "Classificação dos custos fixos", "setor": "Setores de atuação",
    "canal": "Canais de aquisição",
}
ETAPAS_FASE = {
    "Antes": ["Descoberta", "Consideração", "Decisão"],
    "Durante": ["Compra e Pagamento", "Entrega", "Uso"],
    "Depois": ["Avaliação", "Suporte", "Lealdade e advocacia"],
}
ETAPAS = [e for v in ETAPAS_FASE.values() for e in v]
EMOCOES = {"Entusiasmado": "🤩", "Satisfeito": "🙂", "Neutro": "😐", "Insatisfeito": "🙁", "Frustrado": "😠"}

# ------------------------------------------------------------------------------
# PÁGINAS (ferramentas) DO PLANO
# ------------------------------------------------------------------------------
PAGINAS = [
    ("projeto", "Identificação", "🏢", "Informações e sócios",
     "Dados do negócio, integrantes do grupo, premissas financeiras e quadro societário."),
    ("segmentos", "Cliente e mercado", "🎯", "Segmentação de mercado",
     "Liste os grupos de clientes possíveis e dê notas para descobrir qual atender primeiro."),
    ("personas", "Cliente e mercado", "🧑", "Gerador de personas",
     "Crie personagens que representam o cliente ideal: dados, características, rotina e insights."),
    ("jornada", "Cliente e mercado", "🧭", "Jornada do cliente",
     "Descreva o que a persona faz, pensa e sente antes, durante e depois da compra."),
    ("mercado", "Cliente e mercado", "📐", "Tamanho do mercado",
     "Estime o mercado total (TAM), o alcançável (SAM) e o conquistável (SOM) e confira se a meta de vendas é realista."),
    ("proposta", "Problema e solução", "❤️", "Proposta de valor",
     "Defina, para cada persona, os benefícios que o negócio entrega e o que o diferencia."),
    ("concorrencia", "Problema e solução", "🔍", "Análise da concorrência",
     "Compare o seu negócio com os concorrentes em notas de 0 a 10 por atributo."),
    ("forcas", "Estratégia", "💪", "Forças e fraquezas",
     "Fatores internos: o que o negócio faz bem e o que precisa melhorar."),
    ("oportunidades", "Estratégia", "💡", "Oportunidades e ameaças",
     "Fatores externos: tendências que ajudam e riscos que vêm do ambiente."),
    ("swot", "Estratégia", "🧩", "Análise SWOT",
     "Matriz SWOT gerada automaticamente e a estratégia indicada para o negócio."),
    ("marketing", "Marketing e operação", "📣", "Plano de marketing (4Ps)",
     "Produto, preço, praça e promoção: como o negócio vai chegar ao cliente."),
    ("canais", "Marketing e operação", "🛒", "Canais e funil de vendas",
     "Avalie os canais de aquisição e acompanhe a conversão do funil."),
    ("operacional", "Marketing e operação", "🏭", "Plano operacional",
     "Localização, layout, processos, capacidade produtiva e organograma."),
    ("fornecedores", "Marketing e operação", "🚚", "Fornecedores",
     "Quem fornece cada item, prazos de pagamento e de entrega."),
    ("investimentos", "Finanças", "💰", "Investimentos e fontes",
     "Investimentos fixos, pré-operacionais, capital de giro e de onde vem o dinheiro."),
    ("produtos", "Finanças", "📦", "Produtos e materiais",
     "Materiais, produtos e serviços, ficha técnica e custo de cada item."),
    ("custos", "Finanças", "🧾", "Custos e despesas",
     "Custos variáveis, mão de obra (produção e administrativa) e custos fixos classificados."),
    ("preco", "Finanças", "🏷️", "Formação de preço",
     "Custeio variável e por absorção, preço sugerido pelo markup e ponto de equilíbrio por produto."),
    ("dre", "Finanças", "📄", "DRE e projeções",
     "Demonstração do resultado dos Anos 1 a 3 e a projeção mês a mês."),
    ("fluxo", "Finanças", "💵", "Fluxo de caixa mensal",
     "Entradas e saídas de dinheiro em cada mês do Ano 1."),
    ("indicadores", "Finanças", "📊", "Indicadores financeiros",
     "Semáforo de saúde, VPL, TIR, payback, pontos de equilíbrio e checagens."),
    ("graficos", "Finanças", "📈", "Gráficos",
     "Composição do investimento, dos custos e da receita, ponto de equilíbrio e payback."),
    ("simulador", "Ferramentas complementares", "🎚️", "Simulador de resultados",
     "Cenários otimista e pessimista, simulador “e se?” e análise de sensibilidade."),
    ("experimentos", "Ferramentas complementares", "🧪", "Quadro de experimentação",
     "Hipóteses e testes para validar a ideia antes de investir."),
    ("riscos", "Ferramentas complementares", "⚠️", "Riscos e contingência",
     "Mapa de riscos (probabilidade x impacto) e o que fazer em cada caso."),
    ("cronograma", "Ferramentas complementares", "🗓️", "Cronograma de implantação",
     "O que precisa ser feito antes de abrir as portas, e quando."),
    ("realizado", "Ferramentas complementares", "📏", "Orçado x realizado",
     "Depois de abrir, lance os números reais e compare com o planejado."),
    ("ia", "Ferramentas complementares", "🤖", "Analisar com IA",
     "Monte o texto pronto para colar no ChatGPT, Gemini ou Claude e receber análises e estratégias."),
    ("sumario", "Conclusão e entrega", "📝", "Sumário executivo",
     "Resumo do plano gerado automaticamente, para revisar e completar."),
    ("apresentacao", "Conclusão e entrega", "🎤", "Modo apresentação",
     "Telas grandes com os principais números para apresentar o plano."),
    ("relatorio", "Conclusão e entrega", "📤", "Relatório e exportação",
     "PDF completo, pitch deck, Word, Excel e arquivo do projeto."),
    ("professor", "Conclusão e entrega", "🎓", "Modo professor",
     "Compare os projetos da turma, aplique a rubrica e gere fichas de avaliação."),
    ("config", "Conclusão e entrega", "⚙️", "Configurações e listas",
     "Listas personalizadas (unidades, categorias...), versões salvas e dados deste computador."),
]
PAG = {p[0]: p for p in PAGINAS}
GRUPOS = list(dict.fromkeys(p[1] for p in PAGINAS))


def ir(pagina: str, alerta: str | None = None):
    ss.pagina = pagina
    ss.alerta = (pagina, alerta) if alerta else None


# ------------------------------------------------------------------------------
# TEMA (CLARO / ESCURO) E ESTILO VISUAL
# ------------------------------------------------------------------------------
def aplicar_tema():
    base = "dark" if ss.tema == "escuro" else "light"
    try:
        if st.get_option("theme.base") != base:
            st._config.set_option("theme.base", base)
    except Exception:
        pass
    escuro = ss.tema == "escuro"
    fundo_card = "linear-gradient(135deg,#1f4f8f 0%,#2f6fd0 100%)" if not escuro else \
        "linear-gradient(135deg,#1b2d4f 0%,#24467d 100%)"
    css_escuro = """
      .stApp, [data-testid="stHeader"] {background-color:#0f1724 !important; color:#e6e9ef !important;}
      [data-testid="stSidebar"] {background-color:#151f30 !important;}
      .stApp p, .stApp span, .stApp label, .stApp li, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
      .stApp div[data-testid="stMarkdownContainer"] {color:#e6e9ef;}
      [data-testid="stMetricValue"] {color:#ffffff !important;}
      .pn-card-info {background:#1a2638 !important; border-color:#2b3a52 !important;}
    """ if escuro else ""
    st.markdown(f"""
    <style>
      [class*="st-key-card_"] button {{
        min-height: 128px; width: 100%; border-radius: 16px; border: 0;
        background: {fundo_card}; color: #fff; box-shadow: 0 4px 14px rgba(20,40,80,.18);
        transition: transform .12s ease, box-shadow .12s ease; white-space: pre-line;
      }}
      [class*="st-key-card_"] button:hover {{ transform: translateY(-3px); box-shadow: 0 8px 22px rgba(20,40,80,.28);
        color:#fff; border:0; }}
      [class*="st-key-card_"] button p {{ font-size: 1.02rem; font-weight: 600; line-height: 1.35; color:#fff; }}
      [class*="st-key-nav_"] button {{ justify-content: flex-start; text-align: left; padding: .25rem .6rem;
        min-height: 2rem; }}
      [class*="st-key-voltar"] button {{ font-weight: 600; }}
      .pn-grupo {{ font-size: 1.15rem; font-weight: 700; margin: 1.2rem 0 .4rem 0; }}
      .pn-card-info {{ border: 1px solid #dfe5ee; border-radius: 12px; padding: 14px 16px; background: #f7f9fc; }}
      .pn-face {{ font-size: 1.8rem; line-height: 1; }}
      .pn-etapa {{ border-left: 4px solid #2f6fd0; padding: 6px 12px; margin: 8px 0; }}
      {css_escuro}
    </style>""", unsafe_allow_html=True)


aplicar_tema()


# ------------------------------------------------------------------------------
# GERENCIAMENTO DE PROJETOS (cada projeto = um arquivo .pnbox)
# ------------------------------------------------------------------------------
def listar_projetos():
    return sorted(PASTA.glob("*.pnbox"), key=lambda p: p.stem.lower())


def caminho_livre(nome: str) -> Path:
    base = re.sub(r"[^\w\- ()]", "", nome, flags=re.UNICODE).strip() or "Projeto"
    p, i = PASTA / f"{base}.pnbox", 2
    while p.exists():
        p, i = PASTA / f"{base} ({i}).pnbox", i + 1
    return p


def trocar_projeto(caminho):
    ss.db = str(caminho)
    ss.gen += 1
    for k in ("buf", "snap", "hash"):
        ss[k] = {}
    ss.pagina = "inicio"
    ss.alerta = None


if not listar_projetos():
    antigo = APP_DIR / "plano_negocio_completo.pnbox"
    if antigo.exists():
        shutil.copy(antigo, PASTA / "Projeto da versão anterior.pnbox")
    if VERSAO == "branco":
        criar_projeto_em_branco(PASTA / "Meu Projeto.pnbox", "Meu Projeto")
    else:
        criar_projeto_modelo(PASTA / "FitBox (modelo).pnbox")

projetos = listar_projetos()
if "db" not in ss or not Path(ss.db).exists():
    trocar_projeto(projetos[0])

# ------------------------------------------------------------------------------
# CARREGAMENTO DOS DADOS DO PROJETO ATIVO
# ------------------------------------------------------------------------------
conn = conectar(Path(ss.db))
proj = ler_unica(conn, "projeto")
d = {t: ler(conn, t) for t in SCHEMA if t not in ("projeto", "proposta_valor")}
d["proposta_valor"] = ler_unica(conn, "proposta_valor")
PID = f"{Path(ss.db).stem}_{ss.gen}"


def opcoes(tipo: str, extras=()) -> list:
    """Lista padrão + itens personalizados do projeto + valores já usados nos dados."""
    pers = d["listas"].loc[d["listas"]["tipo"] == tipo, "valor"].tolist() if "listas" in d else []
    out = []
    for v in list(LISTAS_PADRAO.get(tipo, [])) + pers + [x for x in extras if x]:
        if v and v not in out:
            out.append(v)
    return out


def marcar_salvo():
    ss.ultimo_salvo = datetime.now().strftime("%H:%M:%S")


# ------------------------------------------------------------------------------
# MOTOR FINANCEIRO SOB DEMANDA
# ------------------------------------------------------------------------------
_CACHE = {}


def R() -> dict:
    """Resultado do motor com os dados atuais (recalcula quando algo muda)."""
    if "res" not in _CACHE:
        _CACHE["res"] = executar_motor(proj, d)
    return _CACHE["res"]


def invalidar():
    _CACHE.clear()


def CHK():
    if "chk" not in _CACHE:
        _CACHE["chk"] = tabela_checagens_v5(R(), d, proj)
    return _CACHE["chk"]


# ------------------------------------------------------------------------------
# TABELAS EDITÁVEIS COM SALVAMENTO AUTOMÁTICO
# ------------------------------------------------------------------------------
def _h(df: pd.DataFrame) -> str:
    try:
        return hashlib.md5(df.to_csv(index=False).encode("utf-8")).hexdigest()
    except Exception:
        return ""


def chave(tabela, extra=""):
    return f"{PID}_{tabela}_{extra}_{ss.ver.get((PID, tabela, extra), 0)}"


def _nova_versao(tabela, extra=""):
    ss.ver[(PID, tabela, extra)] = ss.ver.get((PID, tabela, extra), 0) + 1


def tabela_editavel(df_base, config, tabela, extra="", normalizar=None, salvar_fn=None, altura=None,
                    adicionar=True):
    """
    Tabela editável com botões "Adicionar linha" e "Apagar marcadas" e SALVAMENTO AUTOMÁTICO:
    toda alteração válida é gravada no projeto na hora, sem precisar clicar em Salvar.
    """
    kb = (PID, tabela, extra)
    normalizar = normalizar or (lambda x: x)
    wkey = chave(tabela, extra)
    if wkey not in ss.snap:
        ss.snap[wkey] = normalizar(ss.buf.get(kb, df_base).copy().reset_index(drop=True))
        ss.hash.setdefault(kb, _h(normalizar(df_base.copy().reset_index(drop=True))))
    base = ss.snap[wkey].copy()
    base.insert(0, "Excluir", False)
    cfg = {"Excluir": st.column_config.CheckboxColumn("🗑️", help="Marque as linhas que deseja apagar", width="small")}
    cfg.update(config)
    ed = st.data_editor(base, num_rows="dynamic" if adicionar else "fixed", width="stretch", hide_index=True,
                        column_config=cfg, key=wkey, **({"height": altura} if altura else {}))
    marcadas = ed["Excluir"].fillna(False).astype(bool)
    dados = normalizar(ed.drop(columns="Excluir"))

    def _gravar(df_novo):
        if salvar_fn:
            salvar_fn(df_novo)
        ss.hash[kb] = _h(df_novo)
        marcar_salvo()

    # Salvamento automático
    if _h(dados) != ss.hash.get(kb):
        _gravar(dados)
        invalidar()

    b1, b2, b3 = st.columns([1.2, 1.6, 3])
    if adicionar and b1.button("➕ Adicionar linha", key="add_" + wkey, width="stretch"):
        vazia = normalizar(pd.DataFrame([{c: None for c in dados.columns}]))
        ss.buf[kb] = pd.concat([normalizar(ed[~marcadas].drop(columns="Excluir")), vazia], ignore_index=True)
        _nova_versao(tabela, extra)
        st.rerun()
    if b2.button(f"🗑️ Apagar marcadas ({int(marcadas.sum())})", key="del_" + wkey,
                 disabled=not marcadas.any(), width="stretch"):
        novo = normalizar(ed[~marcadas].drop(columns="Excluir"))
        ss.buf[kb] = novo
        _gravar(novo)
        _nova_versao(tabela, extra)
        st.toast("Linhas apagadas.", icon="🗑️")
        st.rerun()
    b3.caption(f"💾 Salvo automaticamente{' às ' + ss.ultimo_salvo if ss.ultimo_salvo else ''}")
    return dados


def editor(tabela, config, campo_chave, altura=None, filtro=None):
    """
    Tabela editável ligada diretamente a uma tabela do projeto.
    filtro={coluna: valor}: edita só as linhas daquele grupo, mantendo as demais intactas.
    """
    filtro = filtro or {}
    if filtro:
        mask = pd.Series(True, index=d[tabela].index)
        for col, val in filtro.items():
            mask &= d[tabela][col] == val
        df_base = d[tabela][mask].reset_index(drop=True)
        extra = "|".join(f"{c}={v}" for c, v in filtro.items())
    else:
        df_base, extra = d[tabela], ""

    def norm(x):
        x = limpar(x, tabela)
        for col, val in filtro.items():
            x[col] = val
        return x

    def gravar(df_novo):
        if filtro:
            m = pd.Series(True, index=d[tabela].index)
            for col, val in filtro.items():
                m &= d[tabela][col] == val
            completo = pd.concat([d[tabela][~m], df_novo], ignore_index=True)
        else:
            completo = df_novo
        salvar(conn, tabela, completo, campo_chave)
        validos = completo[completo[campo_chave].astype(str).str.strip() != ""] if campo_chave else completo
        d[tabela] = limpar(validos.reset_index(drop=True), tabela)

    cfg = {k: v for k, v in config.items() if k not in filtro}
    cfg.update({k: None for k in filtro})          # esconde as colunas do filtro
    ed = tabela_editavel(df_base, cfg, tabela, extra, normalizar=norm, salvar_fn=gravar, altura=altura)
    if not filtro:
        d[tabela] = ed[ed[campo_chave].astype(str).str.strip() != ""].reset_index(drop=True) if campo_chave else ed
    invalidar()
    return ed


def auto_unica(tabela: str, novos: dict):
    """Salva automaticamente campos de tabelas de linha única (projeto, proposta) quando mudam."""
    atual = proj if tabela == "projeto" else d["proposta_valor"]
    mudou = {k: v for k, v in novos.items() if str(atual.get(k, "")) != str(v)}
    if mudou:
        salvar_unica(conn, tabela, mudou)
        atual.update(mudou)
        marcar_salvo()
        invalidar()


def wk(nome: str) -> str:
    """Chave de widget única por projeto (evita mostrar dados de outro projeto)."""
    return f"{PID}_w_{nome}"


T = st.column_config.TextColumn
N = st.column_config.NumberColumn
S = st.column_config.SelectboxColumn
CK = st.column_config.CheckboxColumn


def moeda(rot, **kw):
    return N(rot, min_value=0.0, format="R$ %.2f", **kw)


def nota(rot, maximo=10):
    return N(rot, min_value=0, max_value=maximo, step=1, format="%d")


# ------------------------------------------------------------------------------
# SITUAÇÃO DE CADA FERRAMENTA (selo nos cartões)
# ------------------------------------------------------------------------------
def _tem(tab, minimo=1):
    return len(d[tab]) >= minimo


def situacao(p: str) -> str:
    """✅ completo · ◐ em andamento · ○ não iniciado"""
    sw = d["swot"]
    regras = {
        "projeto": (bool(proj["nome_negocio"]) and bool(proj["cidade_uf"]), _tem("socios")),
        "segmentos": (_tem("segmentos", 2), _tem("segmentos")),
        "personas": (_tem("personas") and bool(d["personas"]["desafios"].astype(str).str.strip().ne("").any()), _tem("personas")),
        "jornada": (d["jornada_cliente"]["etapa"].nunique() >= 6, _tem("jornada_cliente")),
        "mercado": (proj["tam_clientes"] > 0 and proj["sam_clientes"] > 0 and proj["som_pct"] > 0 and proj["gasto_anual_cliente"] > 0,
                    proj["tam_clientes"] > 0 or proj["gasto_anual_cliente"] > 0),
        "proposta": (_tem("propostas"), bool(d["proposta_valor"]["segmento_alvo"])),
        "concorrencia": (_tem("concorrentes") and _tem("concorrencia_notas", 3), _tem("concorrencia_notas")),
        "forcas": ((sw["natureza"] == "Força").any() and (sw["natureza"] == "Fraqueza").any(),
                   sw["natureza"].isin(["Força", "Fraqueza"]).any()),
        "oportunidades": ((sw["natureza"] == "Oportunidade").any() and (sw["natureza"] == "Ameaça").any(),
                          sw["natureza"].isin(["Oportunidade", "Ameaça"]).any()),
        "swot": (sw["natureza"].nunique() >= 4, len(sw) > 0),
        "marketing": (all(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")),
                      any(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao"))),
        "canais": (_tem("canais") and _tem("funil_vendas"), _tem("canais") or _tem("funil_vendas")),
        "operacional": (bool(proj["loc_endereco"]) and _tem("capacidade"), bool(proj["loc_endereco"]) or _tem("capacidade")),
        "fornecedores": (_tem("fornecedores", 2), _tem("fornecedores")),
        "investimentos": (_tem("inv_fixos") and (_tem("socios") or _tem("emprestimos")), _tem("inv_fixos")),
        "produtos": (_tem("catalogo") and _tem("materiais"), _tem("catalogo") or _tem("materiais")),
        "custos": (_tem("pessoal") and _tem("custos_fixos"), _tem("pessoal") or _tem("custos_fixos")),
        "preco": (_tem("catalogo"), _tem("catalogo")),
        "dre": (_tem("catalogo") and _tem("custos_fixos"), _tem("catalogo")),
        "fluxo": (_tem("catalogo") and _tem("capital_giro_prazos"), _tem("catalogo")),
        "indicadores": (_tem("catalogo") and _tem("inv_fixos"), _tem("catalogo")),
        "graficos": (_tem("catalogo"), _tem("catalogo")),
        "simulador": (_tem("catalogo"), _tem("catalogo")),
        "experimentos": (_tem("experimentos"), _tem("experimentos")),
        "riscos": (_tem("riscos", 3), _tem("riscos")),
        "cronograma": (_tem("cronograma", 3), _tem("cronograma")),
        "realizado": (_tem("realizado", 1), _tem("realizado")),
        "sumario": (len(proj["sumario_executivo"]) > 200, bool(proj["sumario_executivo"])),
    }
    if p not in regras:
        return ""
    completo, iniciado = regras[p]
    return "✅" if completo else ("◐" if iniciado else "○")


def completude() -> float:
    sts = [situacao(p[0]) for p in PAGINAS if situacao(p[0])]
    return sum(1 if s == "✅" else (0.5 if s == "◐" else 0) for s in sts) / max(len(sts), 1)


# ------------------------------------------------------------------------------
# MENU LATERAL
# ------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📊 PNBOX Acadêmico")
    if st.button("🏠 Painel do plano", key="nav_inicio", width="stretch",
                 type="primary" if ss.pagina == "inicio" else "secondary"):
        ir("inicio")
        st.rerun()

    nomes = [p.stem for p in projetos]
    atual = Path(ss.db).stem
    escolha = st.selectbox("📁 Projeto ativo", nomes, index=nomes.index(atual) if atual in nomes else 0)
    if escolha != atual:
        trocar_projeto(PASTA / f"{escolha}.pnbox")
        st.rerun()

    with st.expander("➕ Novo projeto"):
        novo_nome = st.text_input("Nome do novo projeto", placeholder="Ex.: Padaria Pão Nosso")
        tipo = st.radio("Começar a partir de", ["Em branco (do zero)", "Modelo FitBox (alimentação saudável)",
                                                "Modelo Café Mucuri (cafeteria e padaria)", "Cópia do projeto atual"])
        if st.button("Criar projeto", type="primary", width="stretch"):
            if not novo_nome.strip():
                st.error("Informe um nome.")
            else:
                destino = caminho_livre(novo_nome)
                if tipo.startswith("Em branco"):
                    criar_projeto_em_branco(destino, novo_nome.strip())
                elif tipo.startswith("Modelo FitBox"):
                    criar_projeto_modelo(destino)
                elif tipo.startswith("Modelo Café"):
                    criar_projeto_cafe(destino)
                else:
                    conn.commit()
                    shutil.copy(ss.db, destino)
                c = conectar(destino)
                salvar_unica(c, "projeto", {"nome_negocio": novo_nome.strip()})
                c.close()
                trocar_projeto(destino)
                st.rerun()

    with st.expander("📂 Abrir arquivo .pnbox"):
        arq = st.file_uploader("Selecione um arquivo .pnbox", type=["pnbox"], key=f"up_{ss.gen}")
        if arq is not None and st.button("Abrir este arquivo", width="stretch"):
            dados_arq = arq.getvalue()
            if not dados_arq.startswith(b"SQLite format 3"):
                st.error("Arquivo inválido: não é um projeto .pnbox.")
            else:
                destino = caminho_livre(Path(arq.name).stem)
                destino.write_bytes(dados_arq)
                conectar(destino).close()      # atualiza a estrutura, se for de versão antiga
                trocar_projeto(destino)
                st.rerun()

    with st.expander("🗑️ Excluir projeto atual"):
        confirma = st.checkbox(f"Confirmo a exclusão de “{atual}”")
        if st.button("Excluir", disabled=not confirma, width="stretch"):
            conn.close()
            Path(ss.db).unlink(missing_ok=True)
            ss.pop("db")
            st.rerun()

    st.divider()
    tema_escuro = st.toggle("🌙 Tema escuro", value=ss.tema == "escuro", key="tg_tema")
    if tema_escuro != (ss.tema == "escuro"):
        ss.tema = "escuro" if tema_escuro else "claro"
        st.rerun()

    with st.expander("🧭 Ir direto para...", expanded=False):
        for g in GRUPOS:
            st.caption(g)
            for p in [x for x in PAGINAS if x[1] == g]:
                if st.button(f"{p[2]} {p[3]}", key=f"nav_{p[0]}", width="stretch",
                             type="primary" if ss.pagina == p[0] else "tertiary"):
                    ir(p[0])
                    st.rerun()

    st.divider()
    if NAVEGADOR:
        st.info("**Versão on-line:** seus projetos ficam guardados **neste navegador, neste computador**. "
                "Para levar a outro computador ou entregar ao professor, baixe o **.pnbox** em "
                "📤 Relatório e exportação.")
    elif MODO_WEB:
        st.warning("**Versão on-line:** o projeto fica só nesta aba. Ao terminar, baixe o **.pnbox** em "
                   "📤 Relatório e exportação.")
    else:
        st.caption(f"Projetos salvos em:\n`{PASTA}`")
    if not MODO_WEB and st.button("⏻ Encerrar o programa", width="stretch"):
        st.success("Programa encerrado. Pode fechar esta aba.")
        threading.Timer(1.0, lambda: os._exit(0)).start()
        st.stop()
    st.caption(f"Versão {VERSAO_PROGRAMA}")


# ------------------------------------------------------------------------------
# CABEÇALHO DAS FERRAMENTAS
# ------------------------------------------------------------------------------
def cabecalho(p: str):
    _, grupo, icone, titulo, descricao = PAG[p]
    c1, c2 = st.columns([1, 5])
    if c1.button("← Voltar ao plano", key="voltar_" + p, width="stretch"):
        ir("inicio")
        st.rerun()
    c2.caption(f"{grupo}  ›  {titulo}")
    st.title(f"{icone} {titulo}")
    st.markdown(f"<div class='pn-card-info'>ℹ️ {descricao}</div>", unsafe_allow_html=True)
    if ss.alerta and ss.alerta[0] == p:
        st.warning(f"**O que ajustar aqui:** {ss.alerta[1]}")
        if st.button("Entendi, ocultar aviso", key="ok_alerta"):
            ss.alerta = None
            st.rerun()
    st.write("")


def proxima(p: str):
    """Botões de navegação para a ferramenta anterior e a próxima."""
    chaves = [x[0] for x in PAGINAS if x[0] not in ("professor", "config")]
    if p not in chaves:
        return
    i = chaves.index(p)
    st.divider()
    c1, _, c3 = st.columns([1.3, 3, 1.3])
    if i > 0 and c1.button(f"← {PAG[chaves[i - 1]][3]}", key="ant_" + p, width="stretch"):
        ir(chaves[i - 1])
        st.rerun()
    if i < len(chaves) - 1 and c3.button(f"{PAG[chaves[i + 1]][3]} →", key="prox_" + p, width="stretch", type="primary"):
        ir(chaves[i + 1])
        st.rerun()

# ##############################################################################
# PARTE: pnbox_ui_paginas1.py
# ##############################################################################
# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: PAINEL E FERRAMENTAS DE ESTRATÉGIA E MERCADO
# ==============================================================================

# ------------------------------------------------------------------------------
# Registros editados um a um (personas, propostas): salvamento automático por id
# ------------------------------------------------------------------------------
def ler_com_id(tabela):
    cols = [c for c, _ in SCHEMA[tabela]]
    return pd.read_sql(f"SELECT id, {', '.join(cols)} FROM {tabela} ORDER BY id", conn)


def auto_registro(tabela, rid, atual: dict, novos: dict):
    mudou = {k: v for k, v in novos.items() if str(atual.get(k) if atual.get(k) is not None else "") != str(v)}
    if mudou:
        conn.execute(f"UPDATE {tabela} SET {', '.join(k + '=?' for k in mudou)} WHERE id=?",
                     list(mudou.values()) + [rid])
        conn.commit()
        d[tabela] = ler(conn, tabela)
        marcar_salvo()
        invalidar()


def novo_registro(tabela, valores: dict):
    cols = list(valores)
    cur = conn.execute(f"INSERT INTO {tabela} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                       [valores[c] for c in cols])
    conn.commit()
    return cur.lastrowid


def apagar_registro(tabela, rid):
    conn.execute(f"DELETE FROM {tabela} WHERE id=?", (rid,))
    conn.commit()


def salvo_txt():
    st.caption(f"💾 Salvo automaticamente{' às ' + ss.ultimo_salvo if ss.ultimo_salvo else ''}")


# ==============================================================================
# PAINEL DO PLANO (tela inicial)
# ==============================================================================
def pg_inicio():
    if not ss.get("bemvindo_ok"):
        with st.container(border=True):
            st.markdown(f"""
### 👋 Bem-vindo ao PNBOX Acadêmico
Este programa ajuda você a construir um **plano de negócio completo**, do cliente aos números, e a descobrir
se a ideia é **viável**. Como usar:

1. **Clique nos cartões** abaixo, de cima para baixo. Cada cartão é uma ferramenta, com explicação de como preencher.
2. **Tudo é salvo automaticamente** enquanto você digita. Não existe botão "Salvar".
3. O selo do cartão mostra o andamento: ✅ concluído · ◐ em andamento · ○ não iniciado.
4. Os **resultados financeiros** (DRE, indicadores, gráficos) são calculados sozinhos a partir do que você preencher.
5. No final, vá em **📤 Relatório e exportação** para baixar o PDF, o pitch, o Word e o arquivo **.pnbox** do projeto.

{"Você está na **versão on-line**: os dados ficam neste navegador. Para entregar ou levar a outro computador, baixe o arquivo .pnbox." if MODO_WEB else ""}
""")
            if st.button("Entendi, começar", type="primary"):
                ss.bemvindo_ok = True
                st.rerun()

    st.title(f"📊 {proj['nome_negocio'] or 'Meu plano de negócio'}")
    linha = " · ".join(x for x in [proj["cidade_uf"], proj["turma"], proj["autores"]] if x)
    if linha:
        st.caption(linha)

    comp = completude()
    try:
        res = R()
        chk = CHK()
        sem = semaforo(res)
        rx = raio_x(res, chk, comp, sem)
    except Exception as e:
        res, chk, sem, rx = None, None, [], None
        st.info(f"Os resultados financeiros aparecem quando houver produtos, custos e investimentos. ({e})")

    c1, c2, c3, c4 = st.columns([2.2, 1, 1, 1])
    with c1:
        st.markdown(f"**Plano {int(comp * 100)}% preenchido**")
        st.progress(comp)
    if rx:
        c2.metric("Raio-X do plano", f"{rx['nota']}/100", help="Preenchimento (30) + checagens OK (30) + saúde financeira (40)")
        c3.metric("Investimento total", brl(res["investimento_total"]).split(",")[0],
                  help=brl(res["investimento_total"]))
        c4.metric("Lucro no Ano 1", brl(res["anos"][1]["resultado"]).split(",")[0],
                  help=brl(res["anos"][1]["resultado"]))
        verdes = sum(1 for s in sem if s["Status"] == "🟢")
        if res["anos"][1]["rec_bruta"] <= 0:
            st.caption("Saúde financeira: aparece quando houver produtos com preço e quantidade cadastrados.")
        else:
            st.caption(f"Saúde financeira: {' '.join(s['Status'] for s in sem)}  ({verdes} de {len(sem)} indicadores no verde)")

    if chk is not None:
        pend = chk[chk["Situação"] != "OK"]
        if not pend.empty:
            with st.expander(f"⚠️ {len(pend)} ponto(s) para corrigir — clique para ver e ir direto ao ajuste", expanded=False):
                for i, r in pend.iterrows():
                    a, b = st.columns([4, 1.3])
                    a.markdown(f"**{r['Teste']}** — {r['Valor apurado']}  \n<small>{r['Como corrigir'] or r['Observação']}</small>",
                               unsafe_allow_html=True)
                    if r["destino"] and b.button("Ir para o ajuste →", key=f"ini_chk_{i}", width="stretch"):
                        ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir'] or r['Observação']}")
                        st.rerun()

    for g in GRUPOS:
        itens = [p for p in PAGINAS if p[1] == g]
        st.markdown(f"<div class='pn-grupo'>{g}</div>", unsafe_allow_html=True)
        for i in range(0, len(itens), 4):
            cols = st.columns(4)
            for col, p in zip(cols, itens[i:i + 4]):
                selo = situacao(p[0])
                rotulo = f"{p[2]}\n\n{p[3]}" + (f"\n\n{selo}" if selo else "")
                with col:
                    if st.button(rotulo, key=f"card_{p[0]}", width="stretch", help=p[4]):
                        ir(p[0])
                        st.rerun()


# ==============================================================================
# IDENTIFICAÇÃO
# ==============================================================================
def pg_projeto():
    cabecalho("projeto")
    t1, t2, t3 = st.tabs(["🏢 Negócio", "👥 Grupo e entrega", "📈 Premissas financeiras"])
    v = {}
    with t1:
        c1, c2, c3 = st.columns(3)
        v["nome_negocio"] = c1.text_input("Nome do negócio", proj["nome_negocio"], key=wk("nome"), max_chars=60)
        objs = ["Criar um novo negócio", "Expandir um negócio existente", "Captar recursos", "Outro"]
        v["objetivo"] = c2.selectbox("Objetivo do plano", objs,
                                     index=objs.index(proj["objetivo"]) if proj["objetivo"] in objs else 0, key=wk("obj"))
        v["periodo_anos"] = c3.number_input("Período de planejamento (anos)", 1, 3, int(proj["periodo_anos"] or 3), key=wk("per"))
        regs = ["Simples Nacional", "MEI", "Lucro Presumido", "Lucro Real"]
        v["enquadramento"] = c1.selectbox("Enquadramento tributário", regs, key=wk("enq"),
                                          index=regs.index(proj["enquadramento"]) if proj["enquadramento"] in regs else 0)
        portes = ["MEI", "Microempresa (ME)", "Pequena Empresa (EPP)", "Médio/Grande Porte"]
        v["porte"] = c2.selectbox("Porte", portes, key=wk("porte"),
                                  index=portes.index(proj["porte"]) if proj["porte"] in portes else 1)
        atuais = [x.strip() for x in proj["setor_atuacao"].split(";") if x.strip()]
        v["setor_atuacao"] = "; ".join(c3.multiselect("Setores de atuação", opcoes("setor", atuais), atuais, key=wk("setor"),
                                                      accept_new_options=True))
        v["cidade_uf"] = c1.text_input("Cidade / UF", proj["cidade_uf"], key=wk("cid"))
        v["contatos"] = c2.text_input("Contatos (site, redes, e-mail)", proj["contatos"], key=wk("cont"))
    with t2:
        st.caption("Estes dados aparecem na capa do relatório e no modo professor.")
        c1, c2 = st.columns(2)
        v["autores"] = c1.text_area("Integrantes do grupo (um por linha ou separados por ;)", proj["autores"], key=wk("aut"))
        v["turma"] = c2.text_input("Curso / disciplina / turma", proj["turma"], key=wk("turma"))
        v["professor"] = c2.text_input("Professor(a)", proj["professor"], key=wk("prof"))
        v["data_entrega"] = c2.text_input("Data de entrega", proj["data_entrega"], key=wk("data"), placeholder="dd/mm/aaaa")
    with t3:
        st.caption("Premissas usadas em todas as projeções. Passe o mouse no ícone ? para ver a explicação de cada uma.")
        p1, p2, p3 = st.columns(3)
        v["aliquota_imposto_pct"] = p1.number_input("Impostos sobre vendas (%)", 0.0, 40.0, proj["aliquota_imposto_pct"], 0.1, key=wk("aliq"),
                                                    help="Alíquota efetiva do Simples Nacional ou soma de PIS/COFINS/ISS/ICMS.")
        v["crescimento_mensal_ano1_pct"] = p2.number_input("Crescimento mensal das vendas no Ano 1 (%)", -10.0, 30.0,
                                                           proj["crescimento_mensal_ano1_pct"], 0.5, key=wk("gm"))
        v["crescimento_anual_pct"] = p3.number_input("Crescimento anual nos Anos 2 e 3 (%)", -50.0, 200.0,
                                                     proj["crescimento_anual_pct"], 1.0, key=wk("ga"))
        v["inflacao_custos_anual_pct"] = p1.number_input("Inflação dos custos fixos (% ao ano)", 0.0, 50.0,
                                                         proj["inflacao_custos_anual_pct"], 0.5, key=wk("inf"))
        v["tma_anual_pct"] = p2.number_input("TMA - taxa mínima de atratividade (% ao ano)", 0.0, 60.0, proj["tma_anual_pct"], 0.5,
                                             key=wk("tma"), help="Retorno mínimo que os sócios exigem (ex.: rendimento de uma aplicação segura + prêmio de risco).")
        v["pme_dias"] = p3.number_input("Prazo médio de estoques (dias)", 0.0, 365.0, proj["pme_dias"], 1.0, key=wk("pme"))
        v["reserva_caixa_extra"] = p1.number_input("Reserva financeira de segurança (R$)", 0.0, 1e8, proj["reserva_caixa_extra"], 1000.0, key=wk("res"))
        v["recuperar_giro"] = int(p2.checkbox("No VPL, recuperar o capital de giro no último ano", bool(proj["recuperar_giro"]), key=wk("rg")))
    auto_unica("projeto", v)
    salvo_txt()

    st.subheader("👥 Sócios")
    editor("socios", {"nome": T("Nome", required=True), "funcao": T("Função"), "contato": T("Contato"),
                      "perfil": T("Minicurrículo / atribuições", width="large"),
                      "capital_integralizado": moeda("Capital investido (R$)")}, "nome")
    st.caption(f"Capital próprio total: **{brl(d['socios']['capital_integralizado'].sum())}**")
    proxima("projeto")


# ==============================================================================
# CLIENTE E MERCADO
# ==============================================================================
def pg_segmentos():
    cabecalho("segmentos")
    t1, t2 = st.tabs(["🎯 Segmentos de clientes", "🏆 Priorização"])
    with t1:
        st.caption("Escreva até 4 a 6 grupos de clientes possíveis e dê notas de 1 a 10 para cada pergunta.")
        sg = editor("segmentos", {
            "nome": T("Segmento (máx. 40 caracteres)", required=True, max_chars=40),
            "descricao": T("Descrição do perfil", width="large"),
            "nota_tamanho": nota("É grande o suficiente?"), "nota_lucro": nota("Tem potencial de lucro?"),
            "nota_longevidade": nota("Não tende a desaparecer?"), "nota_concorrencia": nota("Tem poucos concorrentes?"),
            "nota_aderencia": nota("Seus produtos atendem esse segmento?")}, "nome")
    with t2:
        sg = d["segmentos"]
        if sg.empty:
            st.info("Cadastre os segmentos na primeira aba.")
        else:
            rk = sg.assign(Média=sg.iloc[:, 2:7].mean(axis=1)).sort_values("Média", ascending=False).reset_index(drop=True)
            for i, r in rk.iterrows():
                with st.container(border=True):
                    a, b = st.columns([5, 1])
                    a.markdown(f"### {i + 1}º  {r['nome']}\n{r['descricao']}")
                    b.metric("Média", num_br(r["Média"], 1))
            st.altair_chart(alt.Chart(rk).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("nome:N", sort="-x", title=None)),
                width="stretch")
    proxima("segmentos")


CAMPOS_PERSONA = {
    "Dados": [("nome", "Nome", "txt", 40), ("idade", "Idade (anos)", "int", 0), ("estado_civil", "Estado civil", "txt", 30),
              ("escolaridade", "Escolaridade", "txt", 40), ("profissao", "Profissão", "txt", 60),
              ("renda_mensal", "Renda mensal (R$)", "num", 0), ("cidade", "Cidade", "txt", 60),
              ("descricao", "Sobre a persona (máx. 380 caracteres)", "area", 380)],
    "Características": [("citacao", "Frase que representa a persona", "area", 250),
                        ("habitos", "Hábitos e comportamentos", "area", 500),
                        ("caracteristicas", "Características (#tags)", "txt", 200),
                        ("fluencia_digital", "Fluência digital (computador, smartphone, apps, compras on-line)", "txt", 200),
                        ("redes_sociais", "Redes sociais utilizadas", "txt", 150),
                        ("quem_influencia", "Quem influencia as decisões", "txt", 200)],
    "Rotina": [("rotina_manha", "Manhã", "area", 400), ("rotina_tarde", "Tarde", "area", 400), ("rotina_noite", "Noite", "area", 400)],
    "Insights": [("ganhos", "Ganhos esperados", "area", 500), ("desafios", "Dores e desafios", "area", 500),
                 ("duvidas_objecoes", "Dúvidas ou objeções", "area", 500),
                 ("aborrecimentos", "Aborrecimentos que devem ser evitados", "area", 500),
                 ("experiencia_ideal", "Experiência ideal e como o negócio ajuda", "area", 500)],
}


def pg_personas():
    cabecalho("personas")
    pers = ler_com_id("personas")
    c1, c2 = st.columns([4, 1.2])
    nomes = [f"{r['nome']}" for _, r in pers.iterrows()]
    if c2.button("➕ Nova persona", type="primary", width="stretch"):
        rid = novo_registro("personas", {"nome": f"Persona {len(pers) + 1}", "idade": 30})
        ss[wk("persona_sel")] = rid
        st.rerun()
    if pers.empty:
        st.info("Nenhuma persona cadastrada. Clique em **➕ Nova persona**.")
        proxima("personas")
        return
    ids = pers["id"].tolist()
    atual_id = ss.get(wk("persona_sel"), ids[0])
    atual_id = atual_id if atual_id in ids else ids[0]
    escolhido = c1.radio("Persona", ids, index=ids.index(atual_id), horizontal=True,
                         format_func=lambda i: "🧑 " + str(pers.loc[pers["id"] == i, "nome"].iloc[0]))
    ss[wk("persona_sel")] = escolhido
    row = pers[pers["id"] == escolhido].iloc[0].to_dict()

    st.markdown(f"<div style='text-align:center;font-size:4rem;line-height:1'>🧑</div>"
                f"<h3 style='text-align:center;margin-top:0'>{row['nome']}</h3>", unsafe_allow_html=True)
    novos = {}
    abas_p = st.tabs(list(CAMPOS_PERSONA))
    for aba, (grupo, campos) in zip(abas_p, CAMPOS_PERSONA.items()):
        with aba:
            cols = st.columns(3 if grupo == "Dados" else (3 if grupo == "Rotina" else 2))
            for i, (k, rot, tipo, mx) in enumerate(campos):
                alvo = cols[i % len(cols)] if tipo != "area" or grupo in ("Rotina", "Insights", "Características") else st
                val = row.get(k)
                chave_w = wk(f"pers_{escolhido}_{k}")
                if tipo == "int":
                    novos[k] = alvo.number_input(rot, 0, 120, int(val or 0), key=chave_w)
                elif tipo == "num":
                    novos[k] = alvo.number_input(rot, 0.0, 1e7, float(val or 0), 100.0, key=chave_w)
                elif tipo == "area":
                    novos[k] = alvo.text_area(rot, val or "", max_chars=mx, key=chave_w, height=110)
                else:
                    novos[k] = alvo.text_input(rot, val or "", max_chars=mx, key=chave_w)
    auto_registro("personas", escolhido, row, novos)
    a, b = st.columns([4, 1.2])
    with a:
        salvo_txt()
    if b.button("🗑️ Excluir esta persona", width="stretch"):
        apagar_registro("personas", escolhido)
        ss.pop(wk("persona_sel"), None)
        st.rerun()
    proxima("personas")


def pg_jornada():
    cabecalho("jornada")
    pers = d["personas"]["nome"].tolist()
    if not pers:
        st.info("Cadastre uma persona antes de montar a jornada.")
        if st.button("Ir para Personas"):
            ir("personas")
            st.rerun()
        return
    persona = st.radio("Persona", pers, horizontal=True, format_func=lambda n: "🧑 " + n, key=wk("jor_pers"))
    perguntas = {"Antes": "O que o cliente faz antes de entrar em contato com o seu produto ou serviço?",
                 "Durante": "Como é a compra, a entrega e o uso?",
                 "Depois": "O que acontece depois: avaliação, suporte e fidelização?"}
    fases = st.tabs(list(ETAPAS_FASE))
    for aba, fase in zip(fases, ETAPAS_FASE):
        with aba:
            st.markdown(f"**{perguntas[fase]}**")
            jc = d["jornada_cliente"]
            sub = jc[(jc["persona_nome"] == persona) & (jc["macro_fase"] == fase)]
            for etapa in ETAPAS_FASE[fase]:
                linhas = sub[sub["etapa"] == etapa]
                st.markdown(f"<div style='background:rgba(47,111,208,.08);padding:6px 12px;border-radius:6px;"
                            f"font-weight:600;margin-top:10px'>{etapa}</div>", unsafe_allow_html=True)
                if linhas.empty:
                    st.caption("— sem ações cadastradas nesta etapa —")
                for _, r in linhas.iterrows():
                    face = EMOCOES.get(r["estado_emocional"], "😐")
                    st.markdown(
                        f"<div class='pn-etapa'><span class='pn-face'>{face}</span> <b>{r['acoes_cliente']}</b><br>"
                        f"<small>💭 {r['pensamentos'] or '-'} · 📍 {r['pontos_contato'] or '-'}<br>"
                        f"💡 <i>Oportunidade:</i> {r['oportunidades'] or '-'}</small></div>", unsafe_allow_html=True)
            with st.expander(f"✏️ Editar as ações da fase “{fase}”", expanded=sub.empty):
                editor("jornada_cliente", {
                    "etapa": S("Etapa", options=ETAPAS_FASE[fase], required=True),
                    "acoes_cliente": T("Ação do cliente", width="medium", required=True),
                    "pensamentos": T("O que pensa"), "pontos_contato": T("Pontos de contato"),
                    "estado_emocional": S("Emoção", options=list(EMOCOES)),
                    "oportunidades": T("Oportunidade de melhoria", width="medium")},
                    "acoes_cliente", filtro={"persona_nome": persona, "macro_fase": fase})
    jc = d["jornada_cliente"]
    jp = jc[jc["persona_nome"] == persona]
    if not jp.empty:
        st.subheader("📉 Curva emocional da jornada")
        emo = {"Entusiasmado": 2, "Satisfeito": 1, "Neutro": 0, "Insatisfeito": -1, "Frustrado": -2}
        g = jp.assign(Emoção=jp["estado_emocional"].map(emo).fillna(0)).groupby("etapa", sort=False)["Emoção"].mean()
        g = g.reindex([e for e in ETAPAS if e in g.index]).reset_index()
        st.altair_chart(alt.Chart(g).mark_line(point=True, strokeWidth=3, color="#d6006e").encode(
            x=alt.X("etapa:N", sort=ETAPAS, title=None, axis=alt.Axis(labelAngle=0, labelLimit=120)),
            y=alt.Y("Emoção:Q", scale=alt.Scale(domain=[-2, 2]), title="😠  ←  emoção  →  🤩")).properties(height=260),
            width="stretch")
    proxima("jornada")


def pg_mercado():
    cabecalho("mercado")
    st.markdown("""
**Como estimar:** **TAM** é todo o mercado (todas as pessoas que poderiam comprar algo desse tipo na região).
**SAM** é a parte que o seu negócio consegue alcançar (pelo segmento escolhido, raio de entrega, canais).
**SOM** é a fatia que você realisticamente conquista (participação de mercado). Use dados do IBGE, Sebrae,
prefeitura, associações e pesquisa de campo, e anote as fontes.""")
    v = {}
    c1, c2 = st.columns([3, 1])
    v["tam_desc"] = c1.text_input("TAM - mercado total (quem é)", proj["tam_desc"], key=wk("tamd"))
    v["tam_clientes"] = c2.number_input("Nº de clientes potenciais (TAM)", 0.0, 1e9, float(proj["tam_clientes"]), 100.0, key=wk("tamc"))
    v["sam_desc"] = c1.text_input("SAM - mercado alcançável (quem é)", proj["sam_desc"], key=wk("samd"))
    v["sam_clientes"] = c2.number_input("Nº de clientes alcançáveis (SAM)", 0.0, 1e9, float(proj["sam_clientes"]), 100.0, key=wk("samc"))
    c3, c4 = st.columns(2)
    v["som_pct"] = c3.number_input("SOM - participação que pretende conquistar do SAM (%)", 0.0, 100.0, float(proj["som_pct"]), 0.5, key=wk("som"))
    v["gasto_anual_cliente"] = c4.number_input("Gasto médio anual de cada cliente com esse tipo de produto (R$)", 0.0, 1e7,
                                               float(proj["gasto_anual_cliente"]), 50.0, key=wk("gasto"))
    v["mercado_fontes"] = st.text_area("Fontes e premissas usadas", proj["mercado_fontes"], key=wk("mfont"))
    auto_unica("projeto", v)
    salvo_txt()
    mk = mercado(proj, R())
    if mk["preenchido"]:
        df = pd.DataFrame({"Nível": ["TAM", "SAM", "SOM"], "Clientes": [mk["tam_c"], mk["sam_c"], mk["som_c"]],
                           "Mercado (R$/ano)": [mk["tam"], mk["sam"], mk["som"]]})
        a, b = st.columns([1.3, 1])
        a.altair_chart(alt.Chart(df).mark_bar(color="#2f6fd0").encode(
            y=alt.Y("Nível:N", sort=["TAM", "SAM", "SOM"], title=None), x=alt.X("Mercado (R$/ano):Q", title="R$ por ano"),
            tooltip=["Nível", alt.Tooltip("Clientes:Q", format=",.0f"), alt.Tooltip("Mercado (R$/ano):Q", format=",.2f")]
        ).properties(height=200), width="stretch")
        b.metric("TAM", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
        b.metric("SAM", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
        b.metric("SOM", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
        st.subheader("A meta de vendas é realista?")
        p = mk["pct_som"]
        msg = (f"A receita prevista no Ano 1 ({brl(mk['receita_ano1'])}) corresponde a **{pct(p)} do SOM** "
               f"e a **{pct(mk['pct_sam'])} do SAM**.")
        if p > 1:
            st.error("🔴 " + msg + " A meta passa do que o negócio consegue conquistar: reveja as quantidades vendidas "
                     "(Produtos e materiais) ou a estimativa de mercado.")
        elif p > 0.8:
            st.warning("🟡 " + msg + " A meta é ambiciosa: depende de conquistar quase todo o mercado alcançável previsto.")
        else:
            st.success("🟢 " + msg + " A meta cabe no mercado estimado.")
    else:
        st.info("Preencha os quatro números para ver o tamanho do mercado e conferir a meta de vendas.")
    proxima("mercado")


# ==============================================================================
# PROBLEMA E SOLUÇÃO
# ==============================================================================
def pg_proposta():
    cabecalho("proposta")
    props = ler_com_id("propostas")
    pers = d["personas"]["nome"].tolist()
    c1, c2 = st.columns([4, 1.3])
    c1.caption("Uma proposta por persona: escolha os valores (etiquetas) e complete as frases.")
    if c2.button("➕ Nova proposta", type="primary", width="stretch"):
        novo_registro("propostas", {"persona_nome": pers[0] if pers else "", "valores": ""})
        st.rerun()
    if props.empty:
        st.info("Nenhuma proposta cadastrada. Clique em **➕ Nova proposta**.")
    for n, (_, r) in enumerate(props.iterrows(), start=1):
        row = r.to_dict()
        with st.expander(f"❤️ Proposta {n} — {row['persona_nome'] or 'sem persona'}:  {row['valores'][:90]}",
                         expanded=len(props) == 1):
            novos = {}
            opc_p = pers + ([row["persona_nome"]] if row["persona_nome"] and row["persona_nome"] not in pers else [])
            novos["persona_nome"] = st.selectbox("Persona", opc_p or [""], key=wk(f"pp_{row['id']}"),
                                                 index=opc_p.index(row["persona_nome"]) if row["persona_nome"] in opc_p else 0)
            atuais = [x.strip() for x in str(row["valores"]).split(";") if x.strip()]
            novos["valores"] = "; ".join(st.multiselect("Valores entregues (escolha ou digite novos)", opcoes("valor_proposta", atuais),
                                                        atuais, key=wk(f"pv_{row['id']}"), accept_new_options=True))
            a, b = st.columns(2)
            novos["que_desejam"] = a.text_area("Que desejam...", row["que_desejam"], key=wk(f"pq_{row['id']}"))
            novos["por_meio"] = b.text_area("Por meio de...", row["por_meio"], key=wk(f"pm_{row['id']}"))
            novos["diferentemente"] = a.text_area("Diferentemente de...", row["diferentemente"], key=wk(f"pd_{row['id']}"))
            novos["oferecem"] = b.text_area("Os nossos produtos ou serviços oferecem...", row["oferecem"], key=wk(f"po_{row['id']}"))
            auto_registro("propostas", row["id"], row, novos)
            st.info(f"**{proj['nome_negocio']}** ajuda **{novos['persona_nome']}**, que deseja **{novos['que_desejam']}**, "
                    f"por meio de **{novos['por_meio']}**. Diferentemente de **{novos['diferentemente']}**, "
                    f"oferecemos **{novos['oferecem']}**.")
            if st.button("🗑️ Excluir esta proposta", key=wk(f"pdel_{row['id']}")):
                apagar_registro("propostas", row["id"])
                st.rerun()
    salvo_txt()
    proxima("proposta")


def pivo_concorrencia():
    empresas = ["Nossa Empresa"] + [n for n in d["concorrentes"]["nome"] if str(n).strip()]
    notas = d["concorrencia_notas"]
    atributos = list(dict.fromkeys(notas["atributo"]))
    piv = pd.DataFrame({"Atributo": pd.Series(atributos, dtype=object)})
    for emp in empresas:
        mapa = notas[notas["empresa"] == emp].set_index("atributo")["nota"].to_dict()
        piv[emp] = pd.Series([mapa.get(a, None) for a in atributos], dtype=float)
    return piv, empresas


def pg_concorrencia():
    cabecalho("concorrencia")
    st.subheader("1. Concorrentes")
    editor("concorrentes", {"nome": T("Nome do concorrente", required=True),
                            "observacoes": T("Observações (onde fica, preço, pontos fortes)", width="large")}, "nome")
    piv, empresas = pivo_concorrencia()
    st.subheader("2. Notas por atributo (0 a 10)")

    def _norm(x):
        x = x.copy()
        x["Atributo"] = x["Atributo"].astype(object).where(x["Atributo"].notna(), "")
        for e in empresas:
            x[e] = pd.to_numeric(x[e], errors="coerce") if e in x else float("nan")
        return x[["Atributo"] + empresas]

    def _salvar(x):
        x = x[x["Atributo"].astype(str).str.strip() != ""]
        longo = x.melt(id_vars="Atributo", var_name="empresa", value_name="nota").rename(columns={"Atributo": "atributo"})
        salvar(conn, "concorrencia_notas", longo, "atributo")
        d["concorrencia_notas"] = ler(conn, "concorrencia_notas")

    cfg = {"Atributo": T("Atributo avaliado", required=True)}
    cfg.update({e: N(e, min_value=0, max_value=10, step=0.5) for e in empresas})
    pe = tabela_editavel(piv, cfg, "concorrencia_notas", "|".join(empresas), normalizar=_norm, salvar_fn=_salvar)
    pe = pe[pe["Atributo"].astype(str).str.strip() != ""]
    num = pe.copy()
    for e in empresas:
        num[e] = pd.to_numeric(num[e], errors="coerce").fillna(0)
    if not num.empty:
        st.subheader("3. Resultado")
        medias = num[empresas].mean()
        cols = st.columns(min(len(empresas), 5))
        for i, e in enumerate(empresas[:5]):
            cols[i].metric(e, num_br(medias[e], 1))
        longo = num.melt(id_vars="Atributo", var_name="Empresa", value_name="Nota")
        st.altair_chart(alt.Chart(longo).mark_line(point=True).encode(
            x=alt.X("Atributo:N", title=None, axis=alt.Axis(labelAngle=-35, labelLimit=160)),
            y=alt.Y("Nota:Q", scale=alt.Scale(domain=[0, 10])), color="Empresa:N").properties(height=340), width="stretch")
        if len(empresas) > 1:
            idx = num.set_index("Atributo")
            outros = idx[empresas[1:]].max(axis=1)
            st.success("**Forças frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] > outros]) or "nenhuma"))
            st.warning("**Fraquezas frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] < outros]) or "nenhuma"))
            dest = idx["Nossa Empresa"] - outros
            st.info(f"**Destaque do negócio:** {dest.idxmax()} (vantagem de {num_br(dest.max(), 1)} ponto(s) sobre o melhor concorrente).")
    proxima("concorrencia")


# ==============================================================================
# ESTRATÉGIA
# ==============================================================================
def _fatores(ambiente, naturezas, dims, rot_acao):
    editor("swot", {
        "dimensao": S("Dimensão", options=dims),
        "fator": T("Fator", required=True),
        "natureza": S("Classificação", options=naturezas, required=True),
        "impacto": S("Impacto", options=["Alto", "Baixo"], required=True),
        "diagnostico": T("Diagnóstico (por que é assim no seu negócio)", width="medium"),
        "plano_acao": T(rot_acao, width="medium"),
        "priorizado": CK("⭐ Priorizar")}, "fator", filtro={"ambiente": ambiente})


def pg_forcas():
    cabecalho("forcas")
    st.caption("Exemplos de fatores: qualificação da equipe, presença digital, controle de qualidade, controles financeiros, "
               "localização, relação com fornecedores, reputação da marca.")
    dims = opcoes("dimensao_interna", d["swot"].loc[d["swot"]["ambiente"] == "Interno", "dimensao"].tolist())
    _fatores("Interno", ["Força", "Fraqueza"], dims, "Como potencializar (força) ou corrigir (fraqueza)")
    proxima("forcas")


def pg_oportunidades():
    cabecalho("oportunidades")
    st.caption("Exemplos: preocupação com saúde, novas tecnologias de pagamento, inflação, renda dos consumidores, "
               "fiscalização, novos concorrentes, mudanças demográficas.")
    dims = opcoes("dimensao_externa", d["swot"].loc[d["swot"]["ambiente"] == "Externo", "dimensao"].tolist())
    _fatores("Externo", ["Oportunidade", "Ameaça"], dims, "Como aproveitar (oportunidade) ou neutralizar (ameaça)")
    proxima("oportunidades")


def html_swot(sw: pd.DataFrame) -> str:
    def lista(nat):
        itens = sw[sw["natureza"] == nat]
        if itens.empty:
            return "<i>nenhum fator</i>"
        return "<br>".join(("⭐ " if bool(r["priorizado"]) else "• ") + f"{r['fator']} <small>({r['impacto']})</small>"
                           for _, r in itens.iterrows())
    cel = "padding:14px 16px;border-radius:4px;min-height:200px;font-size:14px;line-height:1.55;"
    lat = ("writing-mode:vertical-rl;transform:rotate(180deg);font-weight:700;letter-spacing:2px;"
           "display:flex;align-items:center;justify-content:center;")
    return f"""
    <div style="display:grid;grid-template-columns:44px 1fr 1fr;gap:6px;font-family:sans-serif;">
      <div></div>
      <div style="background:#26548c;color:#fff;text-align:center;font-weight:700;padding:10px;">FATORES POSITIVOS</div>
      <div style="background:#d6006e;color:#fff;text-align:center;font-weight:700;padding:10px;">FATORES NEGATIVOS</div>
      <div style="background:#26548c;color:#fff;{lat}">INTERNOS</div>
      <div style="background:#26548c;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">💪 Forças</div>{lista("Força")}</div>
      <div style="background:#d6006e;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⛓️ Fraquezas</div>{lista("Fraqueza")}</div>
      <div style="background:#c8ccd2;color:#28323c;{lat}">EXTERNOS</div>
      <div style="background:#e2e6eb;color:#28323c;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⬆️ Oportunidades</div>{lista("Oportunidade")}</div>
      <div style="background:#5c6672;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⛈️ Ameaças</div>{lista("Ameaça")}</div>
    </div>"""


def pg_swot():
    cabecalho("swot")
    sw = d["swot"]
    c1, c2 = st.columns(2)
    if c1.button("✏️ Editar forças e fraquezas", width="stretch"):
        ir("forcas")
        st.rerun()
    if c2.button("✏️ Editar oportunidades e ameaças", width="stretch"):
        ir("oportunidades")
        st.rerun()
    st.markdown(html_swot(sw), unsafe_allow_html=True)
    st.caption("⭐ fator priorizado · entre parênteses, o impacto")
    diag = pontuacao_swot(sw)
    st.subheader("Matriz de estratégias")
    q = diag["quadrantes"]
    g1, g2 = st.columns(2)
    for col, k in zip([g1, g2, g1, g2], ["FO", "FrO", "FA", "FrA"]):
        nome_q, comb, _ = QUADRANTES[k]
        destaque = "border:3px solid #2f6fd0;" if k == diag["predominante"] else "border:1px solid #ccc;"
        col.markdown(f"<div style='{destaque}padding:12px;border-radius:8px;margin-bottom:8px;text-align:center;'>"
                     f"<div style='font-size:28px;font-weight:700;'>{q[k]}</div><b>{comb}</b><br>{nome_q}</div>",
                     unsafe_allow_html=True)
    if diag["predominante"]:
        nome_q, comb, txt = QUADRANTES[diag["predominante"]]
        st.success(f"**Situação do negócio: {nome_q}.** {txt}")
    pri = sw[sw["priorizado"].astype(bool)]
    if not pri.empty:
        st.subheader("⭐ Ações estratégicas dos fatores priorizados")
        verbo = {"Força": "Como potencializar", "Fraqueza": "Como corrigir", "Oportunidade": "Como aproveitar", "Ameaça": "Como neutralizar"}
        for _, r in pri.iterrows():
            st.markdown(f"**{r['natureza']} · {r['fator']}** ({r['impacto']} impacto)  \n{r['diagnostico']}  \n"
                        f"➡️ *{verbo.get(r['natureza'], 'Ação')}:* {r['plano_acao']}")
    st.caption("Pontuação: impacto alto = 2 pontos; impacto baixo = 1 ponto.")
    proxima("swot")


# ==============================================================================
# MARKETING E OPERAÇÃO
# ==============================================================================
def pg_marketing():
    cabecalho("marketing")
    v = {}
    c1, c2 = st.columns(2)
    v["mkt_produto"] = c1.text_area("📦 Produto: o que será vendido, linhas, diferenciais, embalagem, marca",
                                    proj["mkt_produto"], height=150, key=wk("m1"))
    v["mkt_preco"] = c2.text_area("🏷️ Preço: estratégia (popular, intermediário, premium), comparação com concorrentes, descontos",
                                  proj["mkt_preco"], height=150, key=wk("m2"))
    v["mkt_praca"] = c1.text_area("📍 Praça: onde e como o cliente compra (loja, delivery, site, apps, revendedores)",
                                  proj["mkt_praca"], height=150, key=wk("m3"))
    v["mkt_promocao"] = c2.text_area("📣 Promoção: como o cliente vai conhecer o negócio (redes, anúncios, parcerias, eventos)",
                                     proj["mkt_promocao"], height=150, key=wk("m4"))
    auto_unica("projeto", v)
    salvo_txt()
    try:
        cat = R()["cat"]
        if not cat.empty:
            st.subheader("Apoio para a decisão de preço")
            a, b, c = st.columns(3)
            a.metric("Ticket médio (Mês 1)", brl(div(cat["receita_mensal"].sum(), cat["qtd_mensal"].sum())))
            b.metric("Margem média sobre o preço", pct(div(cat["margem_unit"].mul(cat["qtd_mensal"]).sum(), cat["receita_mensal"].sum())))
            c.metric("Itens no catálogo", num_br(len(cat), 0))
            if st.button("🏷️ Ver formação de preço por produto"):
                ir("preco")
                st.rerun()
    except Exception:
        pass
    proxima("marketing")


def pg_canais():
    cabecalho("canais")
    t1, t2 = st.tabs(["🛒 Canais de aquisição", "🔻 Funil de vendas"])
    with t1:
        st.caption("Dê notas de 0 a 10. Os três canais com maior média são os **canais de tração**.")
        cn = editor("canais", {"canal": S("Canal", options=opcoes("canal", d["canais"]["canal"].tolist()), required=True),
                               "atratividade": nota("Atratividade"), "alcance": nota("Alcance esperado"),
                               "cliente": nota("Custo por cliente (10 = barato)"), "relacionamento": nota("Relacionamento"),
                               "viabilidade": nota("Viabilidade")}, "canal")
        if not cn.empty:
            cn2 = cn.assign(Média=cn.iloc[:, 1:6].mean(axis=1)).sort_values("Média", ascending=False)
            st.success("**Canais de tração:** " + ", ".join(cn2["canal"].head(3)))
            st.altair_chart(alt.Chart(cn2).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("canal:N", sort="-x", title=None)), width="stretch")
    with t2:
        fu = editor("funil_vendas", {
            "etapa": S("Etapa", options=["Topo do Funil", "Meio do Funil", "Fundo do Funil"], required=True),
            "acao": T("Ação de venda", width="large"), "orcamento": moeda("Orçamento (R$)"),
            "alcancados": N("Pessoas alcançadas", min_value=0, format="%d"),
            "convertidos": N("Pessoas que avançaram / compraram", min_value=0, format="%d")}, "etapa")
        if not fu.empty:
            st.dataframe(tabela_funil(fu), hide_index=True, width="stretch")
            fundo = fu[fu["etapa"] == "Fundo do Funil"]["convertidos"].sum()
            m1, m2, m3 = st.columns(3)
            m1.metric("Orçamento total", brl(fu["orcamento"].sum()))
            m2.metric("Clientes conquistados", num_br(fundo, 0))
            m3.metric("CAC - custo de aquisição por cliente", brl(div(fu["orcamento"].sum(), fundo)))
            agg = fu.groupby("etapa", sort=False)[["alcancados", "convertidos"]].sum().reindex(
                ["Topo do Funil", "Meio do Funil", "Fundo do Funil"]).dropna().reset_index()
            st.altair_chart(alt.Chart(agg).mark_bar(color="#d6006e").encode(
                y=alt.Y("etapa:N", sort=None, title=None), x=alt.X("alcancados:Q", title="Pessoas alcançadas")), width="stretch")
    proxima("canais")


def tabela_funil(fu: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "Etapa": fu["etapa"], "Ação": fu["acao"], "Orçamento": fu["orcamento"].map(brl),
        "Alcançados": fu["alcancados"].map(lambda v: num_br(v, 0)),
        "Convertidos": fu["convertidos"].map(lambda v: num_br(v, 0)),
        "Conversão": [pct(div(c, a)) for c, a in zip(fu["convertidos"], fu["alcancados"])],
        "Custo por conversão": [brl(div(o, c)) for o, c in zip(fu["orcamento"], fu["convertidos"])]})


def dot_organograma() -> str:
    def q(s):
        return '"' + str(s).replace('"', "'") + '"'
    linhas = ['digraph G {', 'rankdir=TB; node [shape=box, style="rounded,filled", fillcolor="#e8f0fb", fontname="Helvetica", fontsize=11];',
              'edge [color="#7a8aa0"];', f'raiz [label={q(proj["nome_negocio"] or "Empresa")}, fillcolor="#2f6fd0", fontcolor="white"];']
    for i, r in d["socios"].iterrows():
        linhas.append(f's{i} [label={q(r["nome"] + chr(10) + (r["funcao"] or "Sócio"))}, fillcolor="#cfe0f7"];')
        linhas.append(f'raiz -> s{i};')
    pes = d["pessoal"]
    for j, area in enumerate(dict.fromkeys(pes["area"].tolist())):
        if not area:
            continue
        linhas.append(f'a{j} [label={q(area)}, fillcolor="#fde3ef"];')
        linhas.append(f'raiz -> a{j};')
        for k, r in pes[pes["area"] == area].iterrows():
            linhas.append(f'c{k} [label={q(r["cargo"] + " (" + num_br(r["num_empregados"], 0) + ")")}];')
            linhas.append(f'a{j} -> c{k};')
    linhas.append("}")
    return "\n".join(linhas)


def pg_operacional():
    cabecalho("operacional")
    t1, t2, t3, t4 = st.tabs(["📍 Localização", "🏗️ Layout e processos", "⚙️ Capacidade x demanda", "🗂️ Organograma"])
    v = {}
    with t1:
        c1, c2 = st.columns([3, 1])
        v["loc_endereco"] = c1.text_input("Endereço / região", proj["loc_endereco"], key=wk("le"))
        v["loc_area_m2"] = c2.number_input("Área (m²)", 0.0, 1e6, float(proj["loc_area_m2"]), 5.0, key=wk("la"))
        v["loc_justificativa"] = st.text_area("Por que este ponto? (fluxo, acesso, concorrentes, custo, segurança)",
                                              proj["loc_justificativa"], key=wk("lj"))
        v["loc_fluxo"] = st.text_area("Fluxo de pessoas e horários de pico", proj["loc_fluxo"], key=wk("lf"))
        v["op_horario"] = st.text_input("Horário de funcionamento", proj["op_horario"], key=wk("oh"))
    with t2:
        v["op_layout"] = st.text_area("Layout (arranjo físico): como os espaços serão organizados", proj["op_layout"], height=130, key=wk("ol"))
        v["op_processos"] = st.text_area("Processos operacionais: passo a passo da compra à entrega (use →)",
                                         proj["op_processos"], height=130, key=wk("op"))
        etapas = [e.strip() for e in re.split(r"->|→|\n", proj["op_processos"] or "") if e.strip()]
        if len(etapas) >= 2:
            dot = "digraph P { rankdir=LR; node [shape=box, style=\"rounded,filled\", fillcolor=\"#e8f0fb\", fontname=Helvetica, fontsize=10];\n"
            dot += " -> ".join(f'"{e[:40]}"' for e in etapas) + ";}"
            st.graphviz_chart(dot, width="stretch")
    with t3:
        v["op_capacidade_txt"] = st.text_area("Como a capacidade foi estimada", proj["op_capacidade_txt"], key=wk("oc"))
        alvos = ["Todos os produtos", "Fabricação Própria", "Serviço", "Revenda"] + d["catalogo"]["descricao"].tolist()
        editor("capacidade", {"recurso": T("Recurso (equipe, máquina, espaço)", required=True, width="medium"),
                              "produto_ou_grupo": S("Atende", options=alvos, required=True),
                              "capacidade_mensal": N("Capacidade por mês (unidades)", min_value=0.0),
                              "observacao": T("Observação", width="medium")}, "recurso")
        cd = capacidade_demanda(d, R())
        if not cd.empty:
            show = cd.copy()
            for c in ("Uso Mês 1", "Uso Mês 12"):
                show[c] = show[c].map(lambda x: ("🔴 " if x > 1 else "🟡 " if x > 0.85 else "🟢 ") + pct(x))
            st.dataframe(show, hide_index=True, width="stretch")
            g = cd.melt(id_vars="Recurso", value_vars=["Uso Mês 1", "Uso Mês 12"], var_name="Momento", value_name="Uso")
            st.altair_chart((alt.Chart(g).mark_bar().encode(x=alt.X("Uso:Q", axis=alt.Axis(format="%")), y=alt.Y("Recurso:N", title=None),
                                                           color="Momento:N", yOffset="Momento:N")
                             + alt.Chart(pd.DataFrame({"x": [1]})).mark_rule(color="red", strokeDash=[5, 3]).encode(x="x:Q")
                             ).properties(height=60 * len(cd) + 40), width="stretch")
            st.caption("Linha vermelha = 100% da capacidade. Acima dela, a operação não dá conta da demanda prevista.")
    with t4:
        st.caption("Gerado a partir dos sócios (aba Informações) e da mão de obra (aba Custos), agrupada por área.")
        try:
            st.graphviz_chart(dot_organograma(), width="stretch")
        except Exception as e:
            st.info(f"Não foi possível desenhar o organograma: {e}")
    auto_unica("projeto", v)
    salvo_txt()
    proxima("operacional")


def pg_fornecedores():
    cabecalho("fornecedores")
    fo = editor("fornecedores", {"nome": T("Fornecedor", required=True), "itens": T("O que fornece", width="medium"),
                                 "prazo_pagamento_dias": N("Prazo de pagamento (dias)", min_value=0, step=1),
                                 "prazo_entrega_dias": N("Prazo de entrega (dias)", min_value=0, step=1),
                                 "condicoes": T("Condições (pedido mínimo, frete, desconto)", width="medium"),
                                 "critico": CK("Crítico?")}, "nome")
    if not fo.empty:
        a, b, c = st.columns(3)
        a.metric("Fornecedores cadastrados", len(fo))
        b.metric("Prazo médio de pagamento (fornecedores)", f"{num_br(fo['prazo_pagamento_dias'].mean(), 0)} dias")
        c.metric("PMP usado no capital de giro", f"{num_br(R()['pmp'], 1)} dias")
        crit = fo[fo["critico"].astype(bool)]
        if not crit.empty:
            st.warning("**Fornecedores críticos** (ter alternativa homologada): " + ", ".join(crit["nome"]))
    proxima("fornecedores")

# ##############################################################################
# PARTE: pnbox_ui_paginas2.py
# ##############################################################################
# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: FERRAMENTAS FINANCEIRAS
# ==============================================================================

def tabela_comp(df):
    return df.assign(Valor=df["Valor"].map(brl), **{"%": df["%"].map(pct)})


def rosca(df, titulo, chave_g, tabela=True):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(
        theta=alt.Theta("Valor:Q", stack=True),
        color=alt.Color("Item:N", sort=list(df["Item"]),
                        legend=alt.Legend(title=None, orient="bottom", columns=2, labelLimit=260)),
        tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    arco = base.mark_arc(innerRadius=55, outerRadius=105)
    rot = base.transform_filter(alt.datum["%"] >= 0.04).mark_text(radius=127, fontSize=12).encode(
        text=alt.Text("%:Q", format=".1%"))
    st.altair_chart((arco + rot).properties(height=360), width="stretch", key=chave_g)
    if tabela:
        st.dataframe(tabela_comp(df), hide_index=True, width="stretch")


def barras(df, titulo, chave_g, cor="#2f6fd0"):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(y=alt.Y("Item:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260)),
                                x=alt.X("Valor:Q", title="R$"),
                                tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    st.altair_chart((base.mark_bar(color=cor) + base.mark_text(align="left", dx=3).encode(text=alt.Text("%:Q", format=".1%")))
                    .properties(height=28 * len(df) + 40), width="stretch", key=chave_g)


# ==============================================================================
# INVESTIMENTOS E FONTES
# ==============================================================================
def pg_investimentos():
    cabecalho("investimentos")
    t1, t2, t3, t4 = st.tabs(["🏗️ Investimentos fixos", "📋 Pré-operacionais", "🔄 Capital de giro", "🏦 Fontes de recursos"])
    with t1:
        st.caption("Máquinas, equipamentos, móveis, veículos, instalações. Vida útil 0 = não deprecia (ex.: terreno).")
        invf = editor("inv_fixos", {
            "descricao": T("Descrição", required=True, width="large"), "quantidade": N("Qtd.", min_value=0.0),
            "valor_unitario": moeda("Valor unitário (R$)"),
            "vida_util_anos": N("Vida útil (anos)", min_value=0, max_value=50, step=1),
            "valor_residual_pct": N("Valor residual (%)", min_value=0.0, max_value=100.0)}, "descricao")
        r = R()
        a, b = st.columns(2)
        a.metric("Total dos investimentos fixos", brl(r["inv_fixo"]))
        b.metric("Depreciação mensal", brl(r["dep_m"]))
    with t2:
        st.caption("Gastos antes de abrir: reformas, legalização, marca, projeto, treinamento, marketing de inauguração.")
        editor("inv_pre", {"descricao": T("Descrição", required=True, width="large"), "valor": moeda("Valor (R$)")}, "descricao")
        st.metric("Total pré-operacional", brl(R()["inv_pre"]))
    with t3:
        st.caption("Recebimento: como os clientes pagam (0 dias = à vista). Pagamento: como a empresa paga os fornecedores. "
                   "Em cada tipo, os percentuais devem somar 100%.")
        editor("capital_giro_prazos", {
            "tipo": S("Tipo", options=["Recebimento", "Pagamento"], required=True),
            "prazo_dias": N("Prazo (dias)", min_value=0, max_value=360, step=1),
            "percentual": N("% das vendas / compras", min_value=0.0, max_value=100.0)}, "tipo")
        r = R()
        c = st.columns(4)
        c[0].metric("PMR (recebimento)", f"{num_br(r['pmr'], 1)} dias")
        c[1].metric("PMP (pagamento)", f"{num_br(r['pmp'], 1)} dias")
        c[2].metric("PME (estoques)", f"{num_br(r['pme'], 0)} dias")
        c[3].metric("Ciclo financeiro", f"{num_br(r['ciclo_fin'], 1)} dias")
        c = st.columns(3)
        c[0].metric("Necessidade de capital de giro", brl(r["ncg"]))
        c[1].metric("Reserva de segurança", brl(r["reserva"]), help="Altere em Informações e sócios → Premissas.")
        c[2].metric("Capital de giro total", brl(r["capital_giro"]))
        st.metric("Estoque inicial (calculado em Produtos e materiais)", brl(r["inv_estoque"]))
    with t4:
        st.caption("O capital próprio vem dos sócios (Informações e sócios). Aqui você cadastra empréstimos e financiamentos.")
        editor("emprestimos", {"descricao": T("Descrição (banco, linha de crédito)", required=True, width="large"),
                               "valor": moeda("Valor (R$)"), "taxa_anual_pct": N("Juros (% ao ano)", min_value=0.0),
                               "prazo_meses": N("Prazo (meses)", min_value=1, step=1)}, "descricao")
        r = R()
        it = r["investimento_total"]
        a, b = st.columns(2)
        a.dataframe(pd.DataFrame({
            "Aplicação": ["Investimentos fixos", "Pré-operacionais", "Estoque inicial", "Capital de giro", "Total"],
            "Valor": [brl(x) for x in (r["inv_fixo"], r["inv_pre"], r["inv_estoque"], r["capital_giro"], it)],
            "%": [pct(div(x, it)) for x in (r["inv_fixo"], r["inv_pre"], r["inv_estoque"], r["capital_giro"], it)]}),
            hide_index=True, width="stretch")
        b.dataframe(pd.DataFrame({
            "Fonte": ["Capital próprio (sócios)", "Capital de terceiros", "Total das fontes"],
            "Valor": [brl(r["capital_proprio"]), brl(r["capital_terceiros"]), brl(r["fontes"])]}), hide_index=True, width="stretch")
        dif = r["fontes"] - it
        if abs(dif) <= 1:
            st.success("🟢 As fontes de recursos cobrem exatamente o investimento total.")
        else:
            falta = "faltam" if dif < 0 else "sobram"
            st.error(f"🔴 Diferença de **{brl(dif)}**: {falta} {brl(abs(dif))} para fechar o investimento total de {brl(it)}.")
            st.markdown("**Como corrigir:** aumente/reduza o capital dos sócios ou o valor dos empréstimos.")
            if st.button("Ajustar o capital dos sócios →"):
                ir("projeto", f"Ajuste o capital dos sócios em {brl(-dif)} no total para igualar o investimento de {brl(it)}.")
                st.rerun()
    proxima("investimentos")


# ==============================================================================
# PRODUTOS, MATERIAIS E FICHA TÉCNICA
# ==============================================================================
def pg_produtos():
    cabecalho("produtos")
    t1, t2, t3, t4 = st.tabs(["① Materiais e insumos", "② Produtos e serviços", "③ Ficha técnica", "④ Resumo de custos"])
    with t1:
        st.caption("Matérias-primas, embalagens e insumos. A quantidade inicial forma o estoque inicial. "
                   "Falta uma unidade ou categoria? Acrescente em ⚙️ Configurações e listas, ou digite direto.")
        mat = editor("materiais", {
            "nome": T("Material / insumo", required=True, width="large"),
            "categoria": S("Categoria", options=opcoes("categoria_material", d["materiais"]["categoria"].tolist()), required=True),
            "unidade": S("Unidade", options=opcoes("unidade", d["materiais"]["unidade"].tolist()), required=True),
            "valor_unitario": moeda("Valor por unidade (R$)"),
            "qtd_estoque_inicial": N("Qtd. no estoque inicial", min_value=0.0)}, "nome")
        _nova_lista_rapida(["unidade", "categoria_material"], "mat")
    with t2:
        st.caption("**Revenda:** informe o valor de aquisição. **Fabricação própria e serviços:** monte a ficha técnica "
                   "(aba ③); se não houver ficha, vale o custo informado.")
        editor("catalogo", {
            "categoria": S("Categoria", options=CATEGORIAS, required=True),
            "descricao": T("Produto / serviço", required=True, width="large"),
            "preco_venda": moeda("Preço de venda (R$)"),
            "custo_aquisicao": moeda("Valor de aquisição / custo informado (R$)"),
            "qtd_mensal": N("Qtd. vendida no Mês 1", min_value=0, step=1),
            "qtd_estoque_inicial": N("Estoque inicial (só revenda)", min_value=0, step=1)}, "descricao")
    with t3:
        cat, mat = d["catalogo"], d["materiais"]
        fabricaveis = cat[cat["categoria"] != "Revenda"]["descricao"].tolist()
        if not fabricaveis:
            st.info("Cadastre produtos de fabricação própria ou serviços na aba ②.")
        elif mat.empty:
            st.info("Cadastre os materiais na aba ①.")
        else:
            prod = st.selectbox("Produto / serviço", fabricaveis, key=wk("ft_prod"))
            ft_all = d["ficha_tecnica"]
            ft = ft_all[ft_all["produto"] == prod][["material", "quantidade"]].reset_index(drop=True)

            def _norm(x):
                x = x[["material", "quantidade"]].copy()
                x["material"] = x["material"].astype(object).where(x["material"].notna(), "")
                x["quantidade"] = pd.to_numeric(x["quantidade"], errors="coerce").fillna(0.0)
                return x

            def _salvar(x):
                x = x[x["material"].astype(str).str.strip() != ""]
                novo = pd.concat([d["ficha_tecnica"][d["ficha_tecnica"]["produto"] != prod], x.assign(produto=prod)],
                                 ignore_index=True)
                salvar(conn, "ficha_tecnica", novo, "material")
                d["ficha_tecnica"] = ler(conn, "ficha_tecnica")

            ft_ed = tabela_editavel(ft, {"material": S("Material", options=mat["nome"].tolist(), required=True),
                                         "quantidade": N("Quantidade (na unidade do material)", min_value=0.0, format="%.3f")},
                                    "ficha_tecnica", prod, normalizar=_norm, salvar_fn=_salvar)
            ft_ed = ft_ed[ft_ed["material"].astype(str).str.strip() != ""]
            precos = dict(zip(mat["nome"], mat["valor_unitario"]))
            unid = dict(zip(mat["nome"], mat["unidade"]))
            catm = dict(zip(mat["nome"], mat["categoria"]))
            calc = ft_ed.assign(Unidade=ft_ed["material"].map(unid), Categoria=ft_ed["material"].map(catm),
                                **{"Valor unit.": ft_ed["material"].map(precos).fillna(0)})
            calc["Custo"] = calc["quantidade"] * calc["Valor unit."]
            custo_p = float(calc["Custo"].sum())
            preco_p = float(cat.loc[cat["descricao"] == prod, "preco_venda"].iloc[0])
            k = st.columns(4)
            k[0].metric("Custo unitário", brl(custo_p))
            k[1].metric("Preço de venda", brl(preco_p))
            k[2].metric("Margem bruta", pct(div(preco_p - custo_p, preco_p)))
            k[3].metric("Markup", f"{num_br(div(preco_p, custo_p))}x" if custo_p else "n.a.")
            if not calc.empty:
                st.dataframe(calc.assign(**{"Valor unit.": calc["Valor unit."].map(brl), "Custo": calc["Custo"].map(brl)}),
                             hide_index=True, width="stretch")
                por_cat = calc.groupby("Categoria")["Custo"].sum()
                st.caption("Composição do custo: " + " · ".join(f"{c}: {pct(div(v, custo_p))}" for c, v in por_cat.items()))
    with t4:
        cu = R()["cat"]
        if not cu.empty:
            st.dataframe(pd.DataFrame({
                "Categoria": cu["categoria"], "Produto / serviço": cu["descricao"], "Preço": cu["preco_venda"].map(brl),
                "Custo unit.": cu["custo_unit"].map(brl), "Origem do custo": cu["origem_custo"],
                "Margem": cu["margem_pct"].map(pct), "Markup": cu["markup"].map(lambda v: f"{num_br(v)}x" if v == v else "n.a."),
                "Qtd./mês": cu["qtd_mensal"].map(lambda v: num_br(v, 0)), "Receita/mês": cu["receita_mensal"].map(brl),
                "Custo/mês": cu["custo_mensal"].map(brl)}), hide_index=True, width="stretch")
            r = R()
            a, b, c = st.columns(3)
            a.metric("Receita do Mês 1", brl(r["receita_m1"]))
            b.metric("Custo direto do Mês 1", brl(r["cdir_m1"]))
            c.metric("Estoque inicial (materiais + revenda)", brl(r["inv_estoque"]))
    proxima("produtos")


def _nova_lista_rapida(tipos, sufixo):
    with st.expander("➕ Acrescentar uma opção às listas (unidade, categoria...)"):
        a, b, c = st.columns([1.5, 2, 1])
        tipo = a.selectbox("Lista", tipos, format_func=lambda t: NOMES_LISTAS[t], key=wk(f"lt_{sufixo}"))
        valor = b.text_input("Nova opção", key=wk(f"lv_{sufixo}"))
        if c.button("Acrescentar", key=wk(f"lb_{sufixo}"), width="stretch") and valor.strip():
            novo_registro("listas", {"tipo": tipo, "valor": valor.strip()})
            d["listas"] = ler(conn, "listas")
            for k in list(ss.snap):
                if k.startswith(f"{PID}_materiais") or k.startswith(f"{PID}_catalogo"):
                    ss.snap.pop(k)
            st.toast(f"“{valor}” acrescentado.")
            st.rerun()


# ==============================================================================
# CUSTOS E DESPESAS
# ==============================================================================
def pg_custos():
    cabecalho("custos")
    with st.expander("📚 Entenda a classificação dos custos", expanded=False):
        st.markdown("""
| Tipo | O que é | Exemplos | Onde entra |
|---|---|---|---|
| **Custo direto (variável)** | Vai direto no produto e varia com a quantidade | matéria-prima, embalagem, mercadoria de revenda | Ficha técnica / valor de aquisição |
| **Mão de obra direta (MOD)** | Pessoas que produzem ou prestam o serviço | cozinheiros, auxiliares de produção | Mão de obra com área "Produção / operação" |
| **Custo indireto de produção (CIF)** | Gasto da produção que não dá para medir por unidade | energia da cozinha, manutenção de máquinas, aluguel da área produtiva | Custos fixos classificados como "Custo indireto" |
| **Despesa** | Gasto para administrar e vender, fora da produção | pró-labore, contador, marketing, tarifas bancárias | Custos fixos "Despesa administrativa / comercial / financeira" |
| **Despesa variável de venda** | Proporcional à receita | impostos, comissões, taxas de cartão | Custos de comercialização (%) |
| **Investimento** | Compra de bens duráveis; entra no resultado só pela depreciação | máquinas, móveis, reforma | Investimentos e fontes |

No **custeio por absorção**, o produto recebe o custo direto **mais** uma parte (rateio) da MOD, dos CIF e da depreciação.
No **custeio variável**, o produto recebe só o custo direto; MOD, CIF e despesas ficam como custos do período.""")
    t1, t2, t3 = st.tabs(["% Custos de comercialização", "👷 Mão de obra", "🏢 Custos fixos"])
    with t1:
        st.caption(f"Impostos sobre vendas ({proj['enquadramento']}): **{num_br(proj['aliquota_imposto_pct'])}%** (altere em Informações e sócios).")
        cv = editor("custos_variaveis", {"descricao": T("Descrição", required=True, width="large"),
                                         "percentual": N("% sobre a receita", min_value=0.0, max_value=100.0)}, "descricao")
        st.metric("Total de custos variáveis de venda (impostos + comercialização)",
                  f"{num_br(proj['aliquota_imposto_pct'] + cv['percentual'].sum())}%")
    with t2:
        pe = editor("pessoal", {"cargo": T("Cargo", required=True, width="medium"),
                                "area": S("Área", options=opcoes("area_pessoal", d["pessoal"]["area"].tolist()), required=True,
                                          help="Produção / operação = mão de obra direta (entra no custo do produto por absorção)."),
                                "num_empregados": N("Nº", min_value=0, step=1), "salario_mensal": moeda("Salário (R$)"),
                                "encargos_pct": N("Encargos (%)", min_value=0.0), "beneficios_unit": moeda("Benefícios por pessoa (R$)"),
                                "atribuicoes": T("Atribuições", width="medium")}, "cargo")
        r = R()
        c = st.columns(4)
        c[0].metric("Total com pessoal / mês", brl(r["pessoal_m"]))
        c[1].metric("Mão de obra direta (produção)", brl(r["mod_m"]))
        c[2].metric("Administrativo", brl(r["pessoal_adm_m"]))
        c[3].metric("Comercial / vendas", brl(r["pessoal_com_m"]))
        st.caption("Referência: no Simples Nacional (anexos I a III), encargos de cerca de 30% a 40% (FGTS, férias, 13º, "
                   "INSS sobre 13º); fora do Simples, em geral acima de 60%. Confirme com o contador.")
    with t3:
        editor("custos_fixos", {"descricao": T("Descrição", required=True, width="large"),
                                "classificacao": S("Classificação", options=opcoes("classe_custo_fixo", d["custos_fixos"]["classificacao"].tolist()),
                                                   required=True),
                                "valor_mensal": moeda("Valor mensal (R$)")}, "descricao")
        r = R()
        c = st.columns(4)
        c[0].metric("Custos indiretos de produção", brl(r["cif_m"]))
        c[1].metric("Despesas administrativas", brl(r["desp_adm_m"]))
        c[2].metric("Despesas comerciais", brl(r["desp_com_m"]))
        c[3].metric("Despesas financeiras", brl(r["desp_fin_m"]))
    r = R()
    st.subheader("Estrutura de custos mensal (Mês 1)")
    est = pd.DataFrame({"Item": ["Custos diretos (materiais e mercadorias)", "Impostos e despesas variáveis de venda",
                                 "Mão de obra direta (produção)", "Custos indiretos de produção", "Depreciação",
                                 "Pessoal administrativo e comercial", "Despesas administrativas, comerciais e financeiras"],
                        "Valor": [r["cdir_m1"], r["receita_m1"] * (r["aliq"] + r["var_pct"]), r["mod_m"], r["cif_m"], r["dep_m"],
                                  r["pessoal_adm_m"] + r["pessoal_com_m"], r["desp_adm_m"] + r["desp_com_m"] + r["desp_fin_m"]]})
    est = est[est["Valor"] > 0]
    est["%"] = est["Valor"] / est["Valor"].sum()
    c1, c2 = st.columns(2)
    with c1:
        rosca(est, "Para onde vai o dinheiro no mês", "g_estcustos", tabela=False)
    with c2:
        st.dataframe(tabela_comp(est), hide_index=True, width="stretch")
        st.metric("Custos que vão no produto (absorção): MOD + CIF + depreciação", brl(r["overhead_m"]))
    proxima("custos")


# ==============================================================================
# FORMAÇÃO DE PREÇO
# ==============================================================================
def pg_preco():
    cabecalho("preco")
    v = {}
    c1, c2 = st.columns(2)
    v["margem_desejada_pct"] = c1.number_input("Margem de lucro desejada sobre o preço (%)", 0.0, 80.0,
                                               float(proj["margem_desejada_pct"]), 1.0, key=wk("mdes"))
    crits = ["Custo direto", "Receita", "Quantidade"]
    v["rateio_criterio"] = c2.selectbox("Critério de rateio dos custos indiretos (absorção)", crits, key=wk("rat"),
                                        index=crits.index(proj["rateio_criterio"]) if proj["rateio_criterio"] in crits else 0)
    auto_unica("projeto", v)
    r = R()
    st.markdown(f"""
**Markup divisor:** preço = custo ÷ [1 − (impostos + despesas variáveis + despesas fixas + margem desejada)]

= custo ÷ [1 − ({pct(r['aliq'])} + {pct(r['var_pct'])} + {pct(r['df_pct'])} + {pct(r['margem_desejada'])})]
= custo × **{num_br(r['markup_mult'], 3)}**

*Despesas fixas* = (custos fixos + pessoal + depreciação) ÷ receita do Mês 1 = {pct(r['df_pct'])}.""")
    if r["divisor_markup"] <= 0:
        st.error("A soma dos percentuais passou de 100%: não existe preço que cubra tudo. Reduza custos fixos ou a margem desejada.")
    cat = r["cat"]
    if cat.empty:
        st.info("Cadastre produtos em Produtos e materiais.")
        proxima("preco")
        return
    tab = pd.DataFrame({
        "Produto / serviço": cat["descricao"], "Categoria": cat["categoria"],
        "Custo direto (variável)": cat["custo_unit"].map(brl), "Custo por absorção": cat["custo_absorcao_unit"].map(brl),
        "Preço praticado": cat["preco_venda"].map(brl), "Preço sugerido (markup)": cat["preco_sugerido"].map(brl),
        "Diferença": cat["dif_preco_pct"].map(lambda v: ("🔻 " if v < -0.05 else "🔺 " if v > 0.05 else "✔️ ") + pct(v) if v == v else "n.a."),
        "Margem (variável)": cat["margem_pct"].map(pct), "Margem (absorção)": cat["margem_absorcao_pct"].map(pct),
        "PE (unid./mês)": cat["pe_unidades_mes"].map(lambda v: num_br(v, 0)), "Venda prevista (unid./mês)": cat["qtd_mensal"].map(lambda v: num_br(v, 0))})
    st.dataframe(tab, hide_index=True, width="stretch", height=480)
    st.caption("🔻 preço abaixo do sugerido (pode não cobrir todos os custos com a margem desejada) · 🔺 acima do sugerido "
               "(confira se o mercado aceita) · PE = quantidade mínima por mês para o negócio empatar, mantido o mix atual de vendas.")
    abaixo = cat[cat["dif_preco_pct"] < -0.05]
    if not abaixo.empty:
        st.warning(f"**{len(abaixo)} item(ns) com preço mais de 5% abaixo do sugerido:** " + ", ".join(abaixo["descricao"].head(8)))
    prej = cat[cat["margem_absorcao_pct"] < 0]
    if not prej.empty:
        st.error("**Itens que dão prejuízo no custeio por absorção:** " + ", ".join(prej["descricao"]))
    proxima("preco")


# ==============================================================================
# DRE, FLUXO DE CAIXA E INDICADORES
# ==============================================================================
def pg_dre():
    cabecalho("dre")
    r = R()
    st.subheader("Projeção mensal do Ano 1")
    mm = r["meses"]
    st.line_chart(mm.set_index("Mês")[["Receita", "Custos variáveis", "Custos fixos", "Resultado operacional"]])
    with st.expander("Ver tabela mensal"):
        st.dataframe(mm.assign(**{c: mm[c].map(brl) for c in mm.columns[1:]}), hide_index=True, width="stretch")
    st.subheader("Demonstração do resultado (Anos 1 a 3)")
    st.dataframe(tabela_dre(r), hide_index=True, width="stretch", height=560)
    st.bar_chart(pd.DataFrame({"Receita": [r["anos"][a]["rec_bruta"] for a in (1, 2, 3)],
                               "EBITDA": [r["anos"][a]["ebitda"] for a in (1, 2, 3)],
                               "Resultado líquido": [r["anos"][a]["resultado"] for a in (1, 2, 3)]},
                              index=["Ano 1", "Ano 2", "Ano 3"]), stack=False)
    proxima("dre")


def pg_fluxo():
    cabecalho("fluxo")
    r = R()
    fm = r["fluxo_mensal"]
    if fm.empty:
        st.info("Sem dados.")
        return
    minimo = fm["Saldo final"].min()
    mes = fm.loc[fm["Saldo final"].idxmin(), "Mês"]
    c = st.columns(3)
    c[0].metric("Saldo inicial (após investimentos)", brl(fm["Saldo inicial"].iloc[0]))
    c[1].metric("Menor saldo do ano", brl(minimo), mes, delta_color="off")
    c[2].metric("Saldo no fim do Ano 1", brl(fm["Saldo final"].iloc[-1]))
    if minimo < 0:
        st.error(f"🔴 O caixa fica negativo em {mes}. Aumente a reserva/capital de giro, alongue prazos ou reduza custos.")
    g = fm[["Mês", "Recebimentos"]].assign(Pagamentos=fm[["Fornecedores", "Impostos e desp. variáveis", "Pessoal e custos fixos",
                                                          "Parcelas de empréstimos"]].sum(axis=1))
    longo = g.melt(id_vars="Mês", var_name="Tipo", value_name="R$")
    barras_ = alt.Chart(longo).mark_bar(opacity=.8).encode(x=alt.X("Mês:N", sort=list(fm["Mês"]), title=None), y="R$:Q",
                                                            color=alt.Color("Tipo:N", scale=alt.Scale(range=["#2f6fd0", "#d6006e"])),
                                                            xOffset="Tipo:N")
    linha = alt.Chart(fm).mark_line(point=True, color="#27ae60", strokeWidth=3).encode(
        x=alt.X("Mês:N", sort=list(fm["Mês"])), y=alt.Y("Saldo final:Q", title="R$"))
    st.altair_chart((barras_ + linha).properties(height=380), width="stretch")
    st.caption("Barras: recebimentos (azul) e pagamentos (rosa) do mês. Linha verde: saldo acumulado em caixa.")
    st.dataframe(fm.assign(**{c: fm[c].map(brl) for c in fm.columns[1:]}), hide_index=True, width="stretch")
    proxima("fluxo")


def pg_indicadores():
    cabecalho("indicadores")
    r = R()
    sem = semaforo(r)
    st.subheader("🚦 Saúde do negócio")
    for i in range(0, len(sem), 3):
        cols = st.columns(3)
        for col, s_ in zip(cols, sem[i:i + 3]):
            with col.container(border=True):
                st.markdown(f"### {s_['Status']} {s_['Valor']}")
                st.markdown(f"**{s_['Indicador']}**")
                st.caption(s_["Leitura"])
                with st.popover("🔎 Como foi calculado"):
                    st.write(s_["Como foi calculado"])
    st.subheader("Resumo financeiro (Ano 1, padrão PNBOX)")
    rs, a1 = r["resumo"], r["anos"][1]
    k = st.columns(4)
    k[0].metric("Investimento total", brl(r["investimento_total"]))
    k[1].metric("Receita anual", brl(rs["receita_anual"]))
    k[2].metric("Resultado operacional", brl(rs["resultado_operacional"]))
    k[3].metric("Lucratividade", pct(rs["lucratividade"]))
    k = st.columns(4)
    k[0].metric("Rentabilidade", pct(rs["rentabilidade"]))
    k[1].metric("Payback simples", f"{num_br(rs['payback_meses'], 1)} meses" if rs["payback_meses"] == rs["payback_meses"] else "não recupera")
    k[2].metric(f"VPL (TMA {num_br(proj['tma_anual_pct'], 1)}%)", brl(r["vpl"]))
    k[3].metric("TIR", pct(r["tir"]) if r["tir"] is not None else "n.a.")
    k = st.columns(4)
    k[0].metric("PE contábil (anual)", brl(a1["pe_contabil"]))
    k[1].metric("PE financeiro (anual)", brl(a1["pe_financeiro"]))
    k[2].metric("PE econômico (anual)", brl(a1["pe_economico"]))
    k[3].metric("Payback descontado", f"{num_br(r['payback_desc'], 1)} meses" if r["payback_desc"] == r["payback_desc"] else "> 36 meses")

    st.subheader("Matriz de indicadores (5 blocos)")
    ind = tabela_indicadores(r)
    bsel = st.selectbox("Bloco", ["Todos"] + list(ind["Bloco"].unique()), key=wk("bloco"))
    st.dataframe(ind if bsel == "Todos" else ind[ind["Bloco"] == bsel], hide_index=True, width="stretch")

    st.subheader("✔️ Checagens de consistência")
    chk = CHK()
    for i, c_ in chk.iterrows():
        ok = c_["Situação"] == "OK"
        a, b, c = st.columns([3.2, 2.2, 1.4])
        a.markdown(f"{'🟢' if ok else '🔴'} **{c_['Teste']}**  \n<small>{c_['Observação']}</small>", unsafe_allow_html=True)
        b.markdown(f"{c_['Valor apurado']}" + ("" if ok else f"  \n<small>📍 {c_['Onde corrigir']}</small>"), unsafe_allow_html=True)
        if not ok and c_["destino"]:
            if c.button("VERIFICAR → ir ao ajuste", key=f"chk_{i}", type="primary", width="stretch"):
                ir(c_["destino"], f"{c_['Teste']}: {c_['Valor apurado']}. {c_['Como corrigir'] or c_['Observação']}")
                st.rerun()
        elif ok:
            c.markdown("✅ OK")
    with st.expander("Notas metodológicas"):
        st.markdown(
            "- **Custos variáveis**: custo direto dos itens (ficha técnica ou valor de aquisição) + impostos + custos de comercialização.\n"
            "- **Custos fixos totais**: pessoal + custos fixos + depreciação (linear, com valor residual).\n"
            "- **Fluxo do projeto** (VPL/TIR): Ano 0 = investimento total; Anos 1 a 3 = EBITDA (no Simples, os tributos já estão "
            "na alíquota sobre vendas). Opcionalmente recupera o capital de giro no Ano 3.\n"
            "- **Payback simples e rentabilidade** seguem o PNBOX: investimento ÷ resultado operacional mensal do Ano 1.\n"
            "- **Juros** pela Tabela Price, mês a mês. **Balanço simplificado**: caixa por diferença (Ativo = Passivo + PL), lucros retidos.")
    proxima("indicadores")


def pg_graficos():
    cabecalho("graficos")
    comp = composicoes(R(), d)
    st.subheader("Investimento inicial")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["investimento"], "Composição do investimento total", "g_inv")
    with g2:
        rosca(comp["fontes"], "Fontes de recursos", "g_fontes")
    barras(comp["inv_fixos"], "Maiores itens do ativo imobilizado", "g_fixos")
    st.subheader("Materiais e custos diretos")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["estoque"], "Estoque inicial por categoria", "g_est")
    with g2:
        rosca(comp["custo_direto"], "Custo direto mensal por categoria", "g_cd")
    st.subheader("Receitas")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["receita_categoria"], "Receita por categoria", "g_rc")
    with g2:
        barras(comp["top_produtos"], "10 itens com maior receita mensal", "g_top")
    st.subheader("Para onde vai cada R$ 1,00 de receita (Ano 1)")
    barras(comp["estrutura_receita"], "Custos, despesas e lucro em % da receita bruta", "g_estr")
    st.subheader("Ponto de equilíbrio contábil (Ano 1)")
    eq = comp["equilibrio"].melt(id_vars="Receita", var_name="Série", value_name="R$")
    pe, rec = comp["pe_ponto"]
    linhas = alt.Chart(eq).mark_line(strokeWidth=3).encode(x=alt.X("Receita:Q", title="Receita anual (R$)"), y="R$:Q",
                                                            color=alt.Color("Série:N", legend=alt.Legend(orient="bottom", title=None)))
    marcas = alt.Chart(pd.DataFrame({"x": [pe, rec], "rot": ["Ponto de equilíbrio", "Receita projetada"]})).mark_rule(
        strokeDash=[6, 4]).encode(x="x:Q", tooltip=["rot", alt.Tooltip("x:Q", format=",.2f")])
    st.altair_chart((linhas + marcas).properties(height=380), width="stretch", key="g_pe")
    st.caption(f"PE contábil: {brl(pe)} por ano · receita projetada: {brl(rec)} · margem de segurança: {pct(div(rec - pe, rec))}")
    st.subheader("Recuperação do investimento (payback)")
    fa = comp["fluxo_acumulado"]
    st.bar_chart(fa.set_index("Ano")[["Acumulado", "Acumulado descontado (TMA)"]], stack=False)
    proxima("graficos")


# ==============================================================================
# SIMULADOR, EXPERIMENTOS, RISCOS, CRONOGRAMA, ORÇADO X REALIZADO
# ==============================================================================
def pg_simulador():
    cabecalho("simulador")
    r = R()
    t1, t2, t3 = st.tabs(["🎚️ E se...?", "🌪️ Sensibilidade", "📊 Cenários"])
    with t1:
        st.caption("Mova as barras e veja o resultado do Ano 1 mudar na hora (nada é gravado no projeto).")
        c = st.columns(5)
        vp = c[0].slider("Preço de venda", -30, 30, 0, 1, format="%d%%", key=wk("s1"))
        vv = c[1].slider("Volume de vendas", -50, 50, 0, 1, format="%d%%", key=wk("s2"))
        vm = c[2].slider("Custo de materiais", -30, 30, 0, 1, format="%d%%", key=wk("s3"))
        vs = c[3].slider("Salários e encargos", -30, 30, 0, 1, format="%d%%", key=wk("s4"))
        vf = c[4].slider("Demais custos fixos", -30, 30, 0, 1, format="%d%%", key=wk("s5"))
        base, sim = simular(r), simular(r, vp / 100, vv / 100, vm / 100, vs / 100, vf / 100)
        k = st.columns(4)
        k[0].metric("Receita (Ano 1)", brl(sim["receita"]), brl(sim["receita"] - base["receita"]))
        k[1].metric("Margem de contribuição", brl(sim["mc"]), brl(sim["mc"] - base["mc"]))
        k[2].metric("Resultado líquido", brl(sim["resultado"]), brl(sim["resultado"] - base["resultado"]))
        k[3].metric("Lucratividade", pct(sim["lucratividade"]),
                    f"{num_br((sim['lucratividade'] - base['lucratividade']) * 100)} p.p." if sim["lucratividade"] == sim["lucratividade"] else None)
        if sim["resultado"] < 0:
            st.error("🔴 Nesse cenário o negócio tem prejuízo no Ano 1.")
    with t2:
        sens = sensibilidade(r)
        st.caption("Efeito de variar cada item em 10% (para mais e para menos) no resultado do Ano 1. "
                   "As barras mais longas mostram onde o negócio é mais sensível.")
        longo = sens.melt(id_vars="Variável", value_vars=["+10%", "-10%"], var_name="Variação", value_name="Efeito no resultado")
        st.altair_chart(alt.Chart(longo).mark_bar().encode(
            y=alt.Y("Variável:N", sort=list(sens["Variável"]), title=None), x=alt.X("Efeito no resultado:Q", title="R$"),
            color=alt.Color("Variação:N", scale=alt.Scale(range=["#27ae60", "#d6006e"]))).properties(height=260), width="stretch")
        st.info(f"O item que mais mexe no lucro é **{sens.iloc[0]['Variável']}**: 10% de variação muda o resultado em "
                f"{brl(sens.iloc[0]['Impacto absoluto'])}.")
    with t3:
        v = {}
        c = st.columns(4)
        v["cen_otim_receita_pct"] = c[0].number_input("Otimista: receita maior em (%)", 0.0, 100.0, proj["cen_otim_receita_pct"], key=wk("c1"))
        v["cen_otim_custo_pct"] = c[1].number_input("Otimista: custos menores em (%)", 0.0, 100.0, proj["cen_otim_custo_pct"], key=wk("c2"))
        v["cen_pess_receita_pct"] = c[2].number_input("Pessimista: receita menor em (%)", 0.0, 100.0, proj["cen_pess_receita_pct"], key=wk("c3"))
        v["cen_pess_custo_pct"] = c[3].number_input("Pessimista: custos maiores em (%)", 0.0, 100.0, proj["cen_pess_custo_pct"], key=wk("c4"))
        auto_unica("projeto", v)
        st.dataframe(cenarios(R(), proj), hide_index=True, width="stretch")
    proxima("simulador")


def pg_experimentos():
    cabecalho("experimentos")
    st.caption("Antes de investir, teste as hipóteses mais arriscadas. Ex.: “Pessoas pagariam R$ 29,90 por uma marmita com macros?” "
               "→ teste: pré-venda para 30 pessoas → métrica: % que compra → sucesso: acima de 40%.")
    ex = editor("experimentos", {"hipotese": T("Hipótese (o que acreditamos)", required=True, width="medium"),
                                 "experimento": T("Experimento (como testar)", width="medium"), "metrica": T("Métrica"),
                                 "criterio_sucesso": T("Critério de sucesso"), "resultado": T("Resultado obtido", width="medium"),
                                 "status": S("Situação", options=["A testar", "Em teste", "Validada", "Invalidada", "Pivotar"])}, "hipotese")
    if not ex.empty:
        cont = ex["status"].replace("", "A testar").value_counts()
        st.caption(" · ".join(f"{k}: {v}" for k, v in cont.items()))
    proxima("experimentos")


def pg_riscos():
    cabecalho("riscos")
    rk = editor("riscos", {"risco": T("Risco", required=True, width="medium"),
                           "probabilidade": N("Probabilidade (1 a 5)", min_value=1, max_value=5, step=1),
                           "impacto": N("Impacto (1 a 5)", min_value=1, max_value=5, step=1),
                           "mitigacao": T("Plano de contingência / mitigação", width="large"), "responsavel": T("Responsável")}, "risco")
    if not rk.empty:
        g = rk.assign(Nível=rk["probabilidade"] * rk["impacto"])
        fundo = pd.DataFrame([{"p": p, "i": i, "n": p * i} for p in range(1, 6) for i in range(1, 6)])
        mapa = alt.Chart(fundo).mark_rect(opacity=.55).encode(
            x=alt.X("p:O", title="Probabilidade"), y=alt.Y("i:O", sort="descending", title="Impacto"),
            color=alt.Color("n:Q", scale=alt.Scale(domain=[1, 8, 25], range=["#27ae60", "#f1c40f", "#e74c3c"]), legend=None))
        pts = alt.Chart(g).mark_text(fontSize=12, fontWeight="bold").encode(
            x="probabilidade:O", y=alt.Y("impacto:O", sort="descending"), text="risco:N")
        st.altair_chart((mapa + pts).properties(height=360), width="stretch")
        alto = g[g["Nível"] >= 12]
        if not alto.empty:
            st.warning("**Riscos críticos (nível ≥ 12):** " + ", ".join(alto["risco"]))
    proxima("riscos")


def pg_cronograma():
    cabecalho("cronograma")
    cr = editor("cronograma", {"atividade": T("Atividade", required=True, width="medium"),
                               "mes_inicio": N("Mês de início (1 = primeiro mês)", min_value=1, step=1),
                               "duracao_meses": N("Duração (meses)", min_value=0.25, step=0.25),
                               "responsavel": T("Responsável"), "custo": moeda("Custo previsto (R$)")}, "atividade")
    if not cr.empty:
        g = cr.assign(fim=cr["mes_inicio"] + cr["duracao_meses"])
        st.altair_chart(alt.Chart(g).mark_bar(cornerRadius=4, color="#2f6fd0").encode(
            x=alt.X("mes_inicio:Q", title="Mês"), x2="fim:Q", y=alt.Y("atividade:N", sort=None, title=None),
            tooltip=["atividade", "responsavel", "mes_inicio", "duracao_meses"]).properties(height=34 * len(g) + 40), width="stretch")
        st.metric("Custo total previsto no cronograma", brl(cr["custo"].sum()))
    proxima("cronograma")


def pg_realizado():
    cabecalho("realizado")
    st.caption("Depois que o negócio abrir, lance mês a mês o que aconteceu de verdade. O programa compara com o planejado.")
    rz = editor("realizado", {"mes": N("Mês (1 a 12)", min_value=1, max_value=12, step=1, required=True),
                              "receita_real": moeda("Receita real (R$)"), "custos_variaveis_real": moeda("Custos variáveis reais (R$)"),
                              "custos_fixos_real": moeda("Custos fixos reais (R$)")}, "mes")
    rz = rz[rz["mes"] > 0]
    if not rz.empty:
        mm = R()["meses"].reset_index(drop=True)
        linhas = []
        for _, x in rz.sort_values("mes").iterrows():
            m = int(x["mes"]) - 1
            if 0 <= m < len(mm):
                prev = mm.iloc[m]
                res_real = x["receita_real"] - x["custos_variaveis_real"] - x["custos_fixos_real"]
                linhas.append({"Mês": f"Mês {m + 1}", "Receita prevista": prev["Receita"], "Receita real": x["receita_real"],
                               "Desvio receita": div(x["receita_real"] - prev["Receita"], prev["Receita"]),
                               "Resultado previsto": prev["Resultado operacional"], "Resultado real": res_real})
        cmp_ = pd.DataFrame(linhas)
        if not cmp_.empty:
            st.dataframe(cmp_.assign(**{c: cmp_[c].map(brl) for c in ["Receita prevista", "Receita real", "Resultado previsto", "Resultado real"]},
                                     **{"Desvio receita": cmp_["Desvio receita"].map(lambda v: ("🔴 " if v < -0.1 else "🟡 " if v < 0 else "🟢 ") + pct(v))}),
                         hide_index=True, width="stretch")
            longo = cmp_.melt(id_vars="Mês", value_vars=["Receita prevista", "Receita real"], var_name="Série", value_name="R$")
            st.altair_chart(alt.Chart(longo).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("Mês:N", sort=None), y="R$:Q", color="Série:N").properties(height=320), width="stretch")
    proxima("realizado")

# ##############################################################################
# PARTE: pnbox_ui_paginas3.py
# ##############################################################################
# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: CONCLUSÃO, ENTREGA, PROFESSOR E CONFIGURAÇÕES
# ==============================================================================


# ------------------------------------------------------------------------------
# PONTOS DE RESTAURAÇÃO (versões salvas dentro do próprio projeto)
# ------------------------------------------------------------------------------
def foto_projeto() -> str:
    dados = {}
    for t in SCHEMA:
        if t == "versoes":
            continue
        dados[t] = ler(conn, t).to_dict("records")
    return json.dumps(dados, ensure_ascii=False, default=str)


def salvar_versao(nome: str, automatica=False):
    conn.execute("INSERT INTO versoes (nome, data, dados) VALUES (?,?,?)",
                 (("[auto] " if automatica else "") + nome, datetime.now().strftime("%d/%m/%Y %H:%M"), foto_projeto()))
    if automatica:   # guarda só as 5 cópias automáticas mais recentes
        ids = [r[0] for r in conn.execute("SELECT id FROM versoes WHERE nome LIKE '[auto]%' ORDER BY id DESC").fetchall()]
        for i in ids[5:]:
            conn.execute("DELETE FROM versoes WHERE id=?", (i,))
    conn.commit()


def restaurar_versao(vid: int):
    linha = conn.execute("SELECT dados FROM versoes WHERE id=?", (vid,)).fetchone()
    if not linha:
        return
    salvar_versao("Antes de restaurar", automatica=True)
    dados = json.loads(linha[0])
    for t, regs in dados.items():
        if t not in SCHEMA:
            continue
        cols = colunas_de(t)
        if t in ("projeto", "proposta_valor"):
            if regs:
                salvar_unica(conn, t, {c: regs[0][c] for c in cols if c in regs[0]})
            continue
        conn.execute(f"DELETE FROM {t}")
        for r in regs:
            cc = [c for c in cols if c in r]
            if cc:
                conn.execute(f"INSERT INTO {t} ({', '.join(cc)}) VALUES ({', '.join('?' * len(cc))})", [r[c] for c in cc])
    conn.commit()
    for k in ("buf", "snap", "hash"):
        ss[k] = {}
    ss.gen += 1
    invalidar()


def copia_automatica():
    """Cria um ponto de restauração a cada 10 minutos de uso."""
    agora = time.time()
    ultimo = ss.get("ultima_copia_auto", {}).get(ss.db)
    if ultimo is None:
        ss.setdefault("ultima_copia_auto", {})[ss.db] = agora
    elif agora - ultimo >= 600:
        salvar_versao("Cópia automática (10 min)", automatica=True)
        ss.ultima_copia_auto[ss.db] = agora


# ------------------------------------------------------------------------------
# ANÁLISE COM IA
# ------------------------------------------------------------------------------
def pg_ia():
    cabecalho("ia")
    res = R()
    st.markdown("""
<div class="pn-card-info">
<b>Como funciona:</b> o programa monta um texto com todos os números do seu plano e as instruções para a IA.
Você copia e cola em qualquer assistente gratuito (ChatGPT, Gemini, Copilot, Claude, DeepSeek...).
Nenhum dado sai do seu computador sem que você mesmo cole. <b>A IA comenta; quem decide é o grupo.</b>
Confira os números que ela citar com os do programa.
</div>""", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns([2, 1])
    with c1:
        escolha = st.multiselect("O que você quer que a IA faça?", list(TAREFAS_IA), format_func=lambda k: TAREFAS_IA[k],
                                 default=["indicadores", "concorrencia", "custos", "graficos"], key=wk("ia_tarefas"))
    with c2:
        publico = st.selectbox("Quem está pedindo", ["estudante de Administração", "estudante de Ciências Contábeis",
                                                      "estudante de Economia", "empreendedor iniciante", "professor"],
                               key=wk("ia_publico"))
    piv, _ = pivo_concorrencia()
    texto = gerar_prompt_ia(proj, d, res, escolha, publico, piv)
    st.caption(f"{len(texto):,} caracteres · cerca de {len(texto) // 4:,} palavras-token".replace(",", "."))
    t1, t2 = st.tabs(["📋 Copiar (botão no canto da caixa)", "✏️ Ver e editar antes de copiar"])
    with t1:
        st.code(texto, language="markdown", wrap_lines=True, height=420)
    with t2:
        editado = st.text_area("Texto", texto, height=420, label_visibility="collapsed", key=wk("ia_edit"))
        texto = editado
    st.download_button("⬇️ Baixar como arquivo .txt", texto.encode("utf-8"), f"prompt_IA_{PID}.txt", "text/plain",
                       width="stretch", type="primary")
    with st.expander("💡 Dicas para tirar mais proveito da IA"):
        st.markdown("""
- Se a IA disser que o texto é longo demais, cole em duas partes e escreva "aguarde a segunda parte".
- Peça depois: *"Agora explique como se eu nunca tivesse estudado finanças"* ou *"Resuma em 5 tópicos para o slide"*.
- Assistentes que executam código (ChatGPT com análise de dados, Gemini, Claude) conseguem desenhar os gráficos.
- Nunca copie a resposta da IA direto para o trabalho: use como ponto de partida e cite que a usou.
""")
    proxima("ia")


# ------------------------------------------------------------------------------
# SUMÁRIO EXECUTIVO E CONCLUSÃO
# ------------------------------------------------------------------------------
def pg_sumario():
    cabecalho("sumario")
    res = R()
    st.info("O sumário executivo é a primeira coisa que o leitor vê, mas a última que se escreve. "
            "Use o rascunho automático como ponto de partida e reescreva com as palavras do grupo.")
    if st.button("✨ Gerar rascunho automático a partir dos dados", type="primary"):
        novo = gerar_sumario(proj, d, res)
        salvar_unica(conn, "projeto", {"sumario_executivo": novo})
        ss.pop(wk("sumario_txt"), None)
        marcar_salvo()
        st.rerun()
    txt = st.text_area("Sumário executivo", proj["sumario_executivo"], height=360, key=wk("sumario_txt"),
                       placeholder="Clique no botão acima para gerar um rascunho, ou escreva aqui.")
    st.caption(f"{len(txt.split())} palavras. Recomenda-se entre 250 e 500.")
    st.divider()
    st.subheader("Conclusão e parecer do grupo")
    st.caption("O negócio é viável? Em que condições? O que precisa ser validado antes de investir? "
               "Use os indicadores e o semáforo para justificar.")
    conc = st.text_area("Conclusão", proj["conclusao"], height=240, key=wk("conclusao_txt"), label_visibility="collapsed")
    auto_unica("projeto", {"sumario_executivo": txt, "conclusao": conc})
    salvo_txt()
    proxima("sumario")


# ------------------------------------------------------------------------------
# MODO APRESENTAÇÃO
# ------------------------------------------------------------------------------
def pg_apresentacao():
    cabecalho("apresentacao")
    res = R()
    a1 = res["anos"][1]
    sem = semaforo(res)
    mk = mercado(proj, res)
    slides = ["capa", "mercado", "proposta", "numeros", "semaforo", "investimento", "dre", "caixa", "fim"]
    ss.setdefault("slide", 0)
    ss.slide = min(ss.slide, len(slides) - 1)
    n1, n2, n3 = st.columns([1, 3, 1])
    if n1.button("◀ Anterior", width="stretch", disabled=ss.slide == 0):
        ss.slide -= 1
        st.rerun()
    n2.progress((ss.slide + 1) / len(slides), text=f"Slide {ss.slide + 1} de {len(slides)}")
    if n3.button("Próximo ▶", width="stretch", disabled=ss.slide == len(slides) - 1, type="primary"):
        ss.slide += 1
        st.rerun()
    s = slides[ss.slide]
    st.markdown('<div class="pn-slide">', unsafe_allow_html=True)
    if s == "capa":
        st.markdown(f"<h1 style='text-align:center;font-size:3.4rem;margin-top:2rem'>{proj['nome_negocio']}</h1>", unsafe_allow_html=True)
        props = d["propostas"]
        if not props.empty:
            st.markdown(f"<h3 style='text-align:center;font-weight:400'>{props['oferecem'].iloc[0]}</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center'>{proj['cidade_uf']} · {str(proj['autores']).replace(chr(10), ', ')}</p>",
                    unsafe_allow_html=True)
    elif s == "mercado":
        st.header("🌍 O mercado")
        if mk["preenchido"]:
            c = st.columns(3)
            c[0].metric("TAM (mercado total)", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
            c[1].metric("SAM (alcançável)", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
            c[2].metric("SOM (nossa meta)", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
            st.markdown(f"### A receita do Ano 1 representa **{pct(mk['pct_som'])}** do SOM.")
        else:
            st.info("Preencha a ferramenta Tamanho do mercado.")
    elif s == "proposta":
        st.header("💎 Proposta de valor")
        for _, p_ in d["propostas"].iterrows():
            st.markdown(f"#### Para **{p_['persona_nome']}**, que deseja {p_['que_desejam']}")
            st.markdown(f"Oferecemos **{p_['oferecem']}**, por meio de {p_['por_meio']}, diferente de {p_['diferentemente']}.")
    elif s == "numeros":
        st.header("💰 Os números do primeiro ano")
        c = st.columns(4)
        c[0].metric("Investimento total", brl(res["investimento_total"]))
        c[1].metric("Receita bruta", brl(a1["rec_bruta"]))
        c[2].metric("Lucro líquido", brl(a1["resultado"]))
        c[3].metric("Lucratividade", pct(div(a1["resultado"], a1["rec_bruta"])))
        c = st.columns(4)
        c[0].metric("VPL (3 anos)", brl(res["vpl"]))
        c[1].metric("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a.")
        c[2].metric("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "> 36 meses")
        c[3].metric("Ponto de equilíbrio", brl(a1["pe_contabil"]))
    elif s == "semaforo":
        st.header("🚦 Saúde do negócio")
        for i in range(0, len(sem), 3):
            cols = st.columns(3)
            for c, it in zip(cols, sem[i:i + 3]):
                with c.container(border=True):
                    st.markdown(f"<div style='font-size:2.2rem'>{it['Status']}</div><b>{it['Indicador']}</b><br>"
                                f"<span style='font-size:1.6rem'>{it['Valor']}</span><br><small>{it['Leitura']}</small>",
                                unsafe_allow_html=True)
    elif s == "investimento":
        st.header("🏗️ Investimento e fontes")
        comp = composicoes(res, d)
        c1, c2 = st.columns(2)
        with c1:
            rosca(comp["investimento"], "Composição do investimento", "ap_inv", tabela=False)
        with c2:
            fontes = pd.DataFrame({"Item": ["Capital próprio", "Capital de terceiros"],
                                   "Valor": [res["capital_proprio"], res["capital_terceiros"]]})
            fontes["%"] = fontes["Valor"] / max(fontes["Valor"].sum(), 1)
            rosca(fontes, "Fontes de recursos", "ap_fontes", tabela=False)
    elif s == "dre":
        st.header("📈 Receita e resultado em 3 anos")
        anos = pd.DataFrame([{"Ano": f"Ano {k}", "Receita": v["rec_bruta"], "Resultado": v["resultado"]} for k, v in res["anos"].items()])
        longo = anos.melt("Ano", var_name="Série", value_name="R$")
        st.altair_chart(alt.Chart(longo).mark_bar().encode(x=alt.X("Ano:N", title=None), xOffset="Série:N", y="R$:Q",
                                                           color=alt.Color("Série:N", scale=alt.Scale(range=["#2f6fd0", "#18a058"])))
                        .properties(height=420), width="stretch")
    elif s == "caixa":
        st.header("🏦 Saldo de caixa mês a mês (Ano 1)")
        fm = res["fluxo_mensal"]
        if not fm.empty:
            st.altair_chart(alt.Chart(fm.assign(n=range(1, len(fm) + 1))).mark_area(
                line={"color": "#2f6fd0"}, color="#2f6fd033").encode(x=alt.X("n:Q", title="Mês"), y=alt.Y("Saldo final:Q", title="R$"))
                .properties(height=420), width="stretch")
    else:
        st.markdown("<h1 style='text-align:center;font-size:3.4rem;margin-top:3rem'>Obrigado! 🙌</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center;font-weight:400'>Perguntas?</h3>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.caption("Dica: aperte F11 para tela cheia e recolha o menu lateral (>> no topo) durante a apresentação.")


# ------------------------------------------------------------------------------
# RELATÓRIO E EXPORTAÇÃO
# ------------------------------------------------------------------------------
def planilha_excel(res) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([proj]).T.reset_index().set_axis(["Campo", "Valor"], axis=1).to_excel(w, sheet_name="Projeto", index=False)
        for t, df_t in d.items():
            if isinstance(df_t, pd.DataFrame) and t != "versoes":
                df_t.to_excel(w, sheet_name=t[:31], index=False)
        res["cat"].to_excel(w, sheet_name="Custos_e_precos", index=False)
        res["inv"].to_excel(w, sheet_name="Depreciacao", index=False)
        res["meses"].to_excel(w, sheet_name="Projecao_mensal", index=False)
        res["fluxo_mensal"].to_excel(w, sheet_name="Fluxo_caixa_mensal", index=False)
        tabela_dre(res).to_excel(w, sheet_name="DRE", index=False)
        tabela_indicadores(res).to_excel(w, sheet_name="Indicadores", index=False)
        pd.DataFrame(semaforo(res)).to_excel(w, sheet_name="Semaforo", index=False)
        cenarios(res, proj).to_excel(w, sheet_name="Cenarios", index=False)
        CHK().drop(columns=["destino"], errors="ignore").to_excel(w, sheet_name="Checagens", index=False)
    return buf.getvalue()


def pg_relatorio():
    cabecalho("relatorio")
    res = R()
    chk = CHK()
    pend = chk[chk["Situação"] != "OK"]
    comp = completude()
    c1, c2, c3 = st.columns(3)
    c1.metric("Plano preenchido", pct(comp, 0))
    c2.metric("Checagens com problema", len(pend))
    c3.metric("Páginas estimadas do PDF", "25 a 35")
    if not pend.empty:
        with st.expander(f"⚠️ {len(pend)} ponto(s) para revisar antes de entregar", expanded=True):
            for i, r in pend.iterrows():
                a, b = st.columns([4, 1])
                a.markdown(f"**{r['Teste']}**: {r['Valor apurado']}. {r['Como corrigir']}")
                if b.button("Ir corrigir →", key=f"rel_chk_{i}", width="stretch"):
                    ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir']}")
                    st.rerun()

    st.subheader("📦 Baixar")
    piv, _ = pivo_concorrencia()
    piv_i = piv
    k = st.columns(3)
    with k[0].container(border=True):
        st.markdown("#### 📄 Relatório completo (PDF)")
        st.caption("Todas as ferramentas, tabelas, gráficos e checagens, na ordem do PNBOX.")
        if st.button("Gerar PDF", type="primary", width="stretch", key="b_pdf"):
            with st.spinner("Montando o relatório... (pode levar alguns segundos)"):
                extras = {"comp": composicoes(res, d), "dre": tabela_dre(res), "indicadores": tabela_indicadores(res),
                          "cenarios": cenarios(res, proj), "checagens": tabela_checagens(res, d),
                          "swot_diag": pontuacao_swot(d["swot"]), "funil": tabela_funil(d["funil_vendas"]),
                          "concorrencia_pivot": piv_i}
                ss[f"pdf_{PID}"] = gerar_pdf(proj, d, res, extras)
        if ss.get(f"pdf_{PID}"):
            st.download_button("⬇️ Baixar PDF", ss[f"pdf_{PID}"], f"Plano_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🎤 Pitch deck (PDF 16:9)")
        st.caption("12 slides prontos para a apresentação oral.")
        if st.button("Gerar pitch", width="stretch", key="b_pitch"):
            with st.spinner("Montando os slides..."):
                ss[f"pitch_{PID}"] = gerar_pitch(proj, d, res)
        if ss.get(f"pitch_{PID}"):
            st.download_button("⬇️ Baixar pitch", ss[f"pitch_{PID}"], f"Pitch_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[2].container(border=True):
        st.markdown("#### 📝 Documento editável (Word)")
        st.caption("Para o grupo revisar o texto e formatar nas normas da ABNT.")
        if st.button("Gerar Word", width="stretch", key="b_docx"):
            with st.spinner("Montando o documento..."):
                ss[f"docx_{PID}"] = gerar_docx(proj, d, res, piv_i) or b""
        if ss.get(f"docx_{PID}") is not None and f"docx_{PID}" in ss:
            if ss[f"docx_{PID}"]:
                st.download_button("⬇️ Baixar .docx", ss[f"docx_{PID}"], f"Plano_{Path(ss.db).stem}.docx",
                                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
            else:
                st.warning("A biblioteca python-docx não está instalada neste computador.")
    k = st.columns(2)
    with k[0].container(border=True):
        st.markdown("#### 📊 Planilha (Excel)")
        st.caption("Todos os dados e cálculos, uma aba por tabela.")
        st.download_button("⬇️ Baixar planilha", planilha_excel(res), f"Planilha_{Path(ss.db).stem}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🗄️ Arquivo do projeto (.pnbox)")
        st.caption("Guarde no pen drive ou na nuvem. É o que o professor recebe para avaliar.")
        conn.commit()
        st.download_button("⬇️ Baixar .pnbox", Path(ss.db).read_bytes(), f"{Path(ss.db).stem}.pnbox",
                           "application/octet-stream", width="stretch", type="primary")
    st.divider()
    st.subheader("🕑 Pontos de restauração")
    st.caption("Salve uma versão antes de grandes mudanças. O programa também cria uma cópia automática a cada 10 minutos "
               "(as 5 mais recentes ficam guardadas).")
    a, b = st.columns([3, 1])
    nome_v = a.text_input("Nome da versão", placeholder="Ex.: Antes de mudar os preços", label_visibility="collapsed")
    if b.button("💾 Salvar versão", width="stretch"):
        salvar_versao(nome_v.strip() or "Versão manual")
        st.toast("Versão salva.")
        st.rerun()
    vs = pd.read_sql("SELECT id, nome, data FROM versoes ORDER BY id DESC", conn)
    if vs.empty:
        st.caption("Nenhuma versão salva ainda.")
    for _, v in vs.iterrows():
        a, b, c = st.columns([4, 1, 1])
        a.markdown(f"**{v['nome']}** · {v['data']}")
        if b.button("Restaurar", key=f"rest_{v['id']}", width="stretch"):
            restaurar_versao(int(v["id"]))
            st.toast("Versão restaurada. Uma cópia do estado anterior foi guardada.")
            st.rerun()
        if c.button("Apagar", key=f"apv_{v['id']}", width="stretch"):
            conn.execute("DELETE FROM versoes WHERE id=?", (int(v["id"]),))
            conn.commit()
            st.rerun()


# ------------------------------------------------------------------------------
# MODO PROFESSOR
# ------------------------------------------------------------------------------
def pg_professor():
    cabecalho("professor")
    st.markdown("""
<div class="pn-card-info">
Envie os arquivos <b>.pnbox</b> dos grupos. O programa recalcula tudo, confere as checagens, mede o preenchimento e
sugere uma nota pela rubrica. A nota sugerida é só um apoio: a banca registra a nota final.
Os arquivos não são guardados; ficam só na memória desta sessão.
</div>""", unsafe_allow_html=True)
    st.write("")
    with st.expander("📐 Rubrica de avaliação (critérios e pesos)", expanded=False):
        ss.setdefault("rubrica", pd.DataFrame(RUBRICA_PADRAO, columns=["Critério", "Peso"]))
        rub = st.data_editor(ss.rubrica, num_rows="fixed", hide_index=True, width="stretch", key="ed_rubrica",
                             column_config={"Critério": T(disabled=True, width="large"), "Peso": N(min_value=0, max_value=100, step=1)})
        st.caption(f"Total: {num_br(rub['Peso'].sum(), 0)} pontos. Os critérios são calculados automaticamente; ajuste apenas os pesos.")
        rubrica = list(zip(rub["Critério"], rub["Peso"]))
    disc, profn = st.columns(2)
    disciplina = disc.text_input("Disciplina", key="prof_disc")
    professor = profn.text_input("Professor(a)", key="prof_nome")
    arqs = st.file_uploader("Arquivos .pnbox dos grupos", type=["pnbox"], accept_multiple_files=True, key="prof_up")
    if not arqs:
        st.info("Nenhum arquivo enviado ainda.")
        return
    avs = []
    with st.spinner("Avaliando os projetos..."):
        tmp = Path(tempfile.mkdtemp())
        for a in arqs:
            dados = a.getvalue()
            if not dados.startswith(b"SQLite format 3"):
                st.warning(f"{a.name}: arquivo inválido, ignorado.")
                continue
            cam = tmp / a.name
            cam.write_bytes(dados)
            try:
                p_, d_ = carregar_projeto_arquivo(cam)
                av = avaliar(p_, d_, rubrica)
                av["Arquivo"] = a.name
                avs.append(av)
            except Exception as e:      # arquivo corrompido ou muito incompleto
                st.warning(f"{a.name}: não foi possível avaliar ({e}).")
    if not avs:
        return
    tab = pd.DataFrame([{k: v for k, v in av.items() if k not in ("res", "notas")} for av in avs])
    st.subheader("📋 Painel da turma")
    mostrar = tab[["Arquivo", "Projeto", "Integrantes", "Preenchimento", "Checagens OK", "Raio-X", "Investimento",
                   "Lucro Ano 1", "VPL", "TIR", "Nota sugerida"]]
    st.dataframe(mostrar, hide_index=True, width="stretch", column_config={
        "Preenchimento": st.column_config.ProgressColumn("Preenchido", format="percent", min_value=0, max_value=1),
        "Checagens OK": st.column_config.ProgressColumn("Checagens OK", format="percent", min_value=0, max_value=1),
        "Investimento": N(format="R$ %.0f"), "Lucro Ano 1": N(format="R$ %.0f"), "VPL": N(format="R$ %.0f"),
        "TIR": N(format="percent"), "Nota sugerida": N(format="%.1f")})
    c = st.columns(3)
    c[0].metric("Projetos", len(tab))
    c[1].metric("Nota média sugerida", num_br(tab["Nota sugerida"].mean(), 1))
    c[2].metric("Projetos com VPL positivo", f"{int((tab['VPL'] > 0).sum())} de {len(tab)}")

    st.subheader("✍️ Notas da banca")
    st.caption("A tabela vem com as notas sugeridas. Altere o que quiser; as fichas em PDF usam estas notas.")
    base = pd.DataFrame([[av["Projeto"]] + av["notas"] for av in avs], columns=["Projeto"] + [c_ for c_, _ in rubrica])
    ed = st.data_editor(base, hide_index=True, width="stretch", key="ed_notas_banca",
                        column_config={"Projeto": T(disabled=True)} |
                        {c_: N(c_[:28] + "…" if len(c_) > 28 else c_, min_value=0, max_value=float(p), step=0.5)
                         for c_, p in rubrica})
    ed["Total"] = ed[[c_ for c_, _ in rubrica]].sum(axis=1)
    st.dataframe(ed[["Projeto", "Total"]], hide_index=True)
    for av, (_, linha) in zip(avs, ed.iterrows()):
        av["notas_finais"] = [float(linha[c_]) for c_, _ in rubrica]

    st.subheader("🔍 Detalhes por grupo")
    for av in avs:
        with st.expander(f"{av['Projeto']} · nota sugerida {num_br(av['Nota sugerida'], 1)}"):
            st.write(f"**Integrantes:** {av['Integrantes'] or '-'}")
            st.write(f"**Pendências:** {av['Pendências'] or 'nenhuma'}")
            st.dataframe(pd.DataFrame({"Critério": [c_ for c_, _ in rubrica], "Peso": [p for _, p in rubrica],
                                       "Sugerida": av["notas"]}), hide_index=True, width="stretch")

    st.subheader("📦 Exportar")
    a, b = st.columns(2)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name="Turma", index=False)
        ed.to_excel(w, sheet_name="Notas_banca", index=False)
    a.download_button("📊 Planilha da turma (Excel)", buf.getvalue(), "avaliacao_turma.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    if b.button("📄 Gerar fichas de avaliação (PDF)", width="stretch", type="primary"):
        ss.fichas = ficha_avaliacao_pdf(avs, rubrica, professor, disciplina)
    if ss.get("fichas"):
        b.download_button("⬇️ Baixar fichas", ss.fichas, "fichas_avaliacao.pdf", "application/pdf", width="stretch")


# ------------------------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------------------------
def pg_config():
    cabecalho("config")
    t1, t2, t3 = st.tabs(["📚 Listas personalizadas", "💾 Onde ficam meus projetos", "🧹 Limpar este computador"])
    with t1:
        st.caption("As opções padrão continuam disponíveis. Aqui você acrescenta categorias, unidades, áreas e "
                   "classificações próprias do seu negócio; elas aparecem nas listas de escolha das ferramentas.")
        tipo = st.selectbox("Lista", list(NOMES_LISTAS), format_func=lambda k: NOMES_LISTAS[k], key="cfg_lista")
        st.markdown("**Padrão:** " + ", ".join(LISTAS_PADRAO.get(tipo, [])))
        editor("listas", {"valor": T("Novo item", required=True)}, "valor", filtro={"tipo": tipo})
    with t2:
        if NAVEGADOR:
            st.markdown("""
Você está usando a **versão on-line**. Os projetos ficam guardados **no armazenamento deste navegador, neste computador**.

- Atualizações do programa **não apagam** seus projetos: o arquivo do programa é trocado, mas a pasta de projetos do navegador continua.
- Eles **somem** se você limpar os dados de navegação do site, usar janela anônima ou trocar de navegador/computador.
- Por segurança, **baixe o .pnbox** (página Relatório e exportação) ao fim de cada sessão de trabalho.
""")
        elif MODO_WEB:
            st.warning("Nesta versão on-line temporária os projetos duram só enquanto a página estiver aberta. Baixe o .pnbox.")
        else:
            st.markdown(f"Os projetos ficam nesta pasta do computador:\n\n`{PASTA}`\n\n"
                        "Atualizar ou reinstalar o programa não apaga essa pasta.")
        st.write(f"Versão do programa: **{VERSAO_PROGRAMA}** · projetos encontrados: **{len(listar_projetos())}**")
    with t3:
        st.warning("Use em computadores de laboratório compartilhados, ao terminar. Todos os projetos guardados "
                   "**neste computador/navegador** serão apagados. Baixe antes o .pnbox do que quiser guardar.")
        ok = st.checkbox("Entendi e já baixei o que precisava", key="cfg_limpar_ok")
        if st.button("🧹 Apagar todos os projetos deste computador", disabled=not ok, type="primary"):
            conn.close()
            for p_ in listar_projetos():
                try:
                    p_.unlink()
                except OSError:
                    pass
            for k in ("db", "buf", "snap", "hash", "ultima_copia_auto"):
                ss.pop(k, None)
            ss.bemvindo_ok = False
            st.success("Pronto. Um projeto novo será criado ao recarregar.")
            st.rerun()


# ==============================================================================
# ROTEADOR
# ==============================================================================
ROTAS = {"inicio": pg_inicio, "projeto": pg_projeto, "segmentos": pg_segmentos, "personas": pg_personas,
         "jornada": pg_jornada, "mercado": pg_mercado, "proposta": pg_proposta, "concorrencia": pg_concorrencia,
         "forcas": pg_forcas, "oportunidades": pg_oportunidades, "swot": pg_swot, "marketing": pg_marketing,
         "canais": pg_canais, "operacional": pg_operacional, "fornecedores": pg_fornecedores,
         "investimentos": pg_investimentos, "produtos": pg_produtos, "custos": pg_custos, "preco": pg_preco,
         "dre": pg_dre, "fluxo": pg_fluxo, "indicadores": pg_indicadores, "graficos": pg_graficos,
         "simulador": pg_simulador, "experimentos": pg_experimentos, "riscos": pg_riscos, "cronograma": pg_cronograma,
         "realizado": pg_realizado, "ia": pg_ia, "sumario": pg_sumario, "apresentacao": pg_apresentacao,
         "relatorio": pg_relatorio, "professor": pg_professor, "config": pg_config}
try:
    ROTAS.get(ss.pagina, pg_inicio)()
    copia_automatica()
finally:
    try:
        conn.commit()
        conn.close()
    except Exception:
        pass
