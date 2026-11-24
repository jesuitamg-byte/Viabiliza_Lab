"""
================================================================================
PNBOX ACADÊMICO v4 - PLANO DE NEGÓCIO E VIABILIDADE ECONÔMICO-FINANCEIRA
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
(4) relatório PDF, (5) interface.
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
               "matplotlib": "matplotlib", "altair": "altair"}
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
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace
from pathlib import Path
import altair as alt
import io
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
    "pessoal": [("cargo", "TEXT"), ("num_empregados", "REAL"), ("salario_mensal", "REAL"),
                ("encargos_pct", "REAL"), ("beneficios_unit", "REAL")],
    "custos_fixos": [("descricao", "TEXT"), ("valor_mensal", "REAL")],
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
    """Converte dados de arquivos .pnbox gerados pelas versões 1 e 2 do programa."""
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
            if c == "priorizado":
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
         "Acompanhamento de indicadores, expansão e consolidação da empresa no mercado.", 42355.0),
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
    ins("""INSERT INTO pessoal (cargo, num_empregados, salario_mensal, encargos_pct, beneficios_unit)
           VALUES (?,?,?,?,?)""", [
        ("Gerente operacional", 1, 3800, 20, 200), ("Chef de cozinha / líder", 1, 3200, 20, 200),
        ("Cozinheiros", 2, 2400, 20, 200), ("Auxiliares de cozinha", 4, 1850, 20, 200),
        ("Atendentes / barista / caixa", 6, 1870, 20, 200),
        ("Vendedores / estoquista / limpeza / entregadores", 8, 1880, 20, 200),
    ])
    ins("INSERT INTO custos_fixos (descricao, valor_mensal) VALUES (?,?)", [
        ("Aluguel, condomínio e IPTU", 7000), ("Pró-labore", 6000), ("Energia elétrica", 4500),
        ("Água", 2000), ("Honorários contábeis", 1800), ("Manutenção", 1300), ("Seguros", 1200),
        ("Serviços de terceiros", 750), ("Marketing e propaganda (fixo)", 600),
        ("Materiais de limpeza", 600), ("Telefone e internet", 500), ("Material de escritório", 450),
        ("Tarifas bancárias", 250), ("Mensalidade de máquinas de cartão", 150),
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
            if m % 12 == 11:
                saldo_fim[ano] += max(saldo, 0.0)
    return {"juros": juros, "amort": amort, "saldo_fim": saldo_fim}


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

    return dict(
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

    # ---------------- 1. SÓCIOS ----------------
    pdf.secao("1. INFORMAÇÕES DO PLANO E QUADRO SOCIETÁRIO", nova_pagina=True)
    s = d["socios"]
    pdf.tabela(pd.DataFrame({
        "Sócio": s["nome"], "Função": s["funcao"], "Minicurrículo": s["perfil"],
        "Capital (R$)": s["capital_integralizado"].map(brl)}), larguras=(40, 34, 76, 30),
        alinhar=("LEFT", "LEFT", "LEFT", "RIGHT"))

    # ---------------- 2. SEGMENTOS ----------------
    pdf.secao("2. SEGMENTAÇÃO DE MERCADO")
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

    # ---------------- 3. PROPOSTA DE VALOR ----------------
    pdf.secao("3. PROPOSTA DE VALOR")
    pv = d["proposta_valor"]
    pdf.paragrafo(f"Nossa empresa ({proj['nome_negocio']}) ajuda {pv['segmento_alvo']} que desejam "
                  f"{pv['beneficios_desejados']} por meio de {pv['meios_entrega']}. Diferentemente de "
                  f"{pv['diferencial_competitivo']}, os nossos produtos ou serviços oferecem "
                  f"{pv['entrega_produtos_servicos']}", "", 10)

    # ---------------- 4. PERSONAS ----------------
    pdf.secao("4. PERSONAS", nova_pagina=True)
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
    pdf.secao("5. JORNADA DO CLIENTE")
    j = d["jornada_cliente"]
    for persona, grupo in j.groupby("persona_nome", sort=False):
        pdf.subtitulo(f"Persona: {persona}")
        pdf.tabela(pd.DataFrame({"Fase": grupo["macro_fase"], "Etapa": grupo["etapa"],
                                 "Ações do cliente": grupo["acoes_cliente"], "Pensamentos": grupo["pensamentos"],
                                 "Pontos de contato": grupo["pontos_contato"], "Emoção": grupo["estado_emocional"],
                                 "Oportunidades": grupo["oportunidades"]}),
                   larguras=(14, 22, 34, 26, 28, 20, 36), alinhar=("LEFT",) * 7, tam=7)

    # ---------------- 6. CONCORRÊNCIA ----------------
    pdf.secao("6. ANÁLISE DA CONCORRÊNCIA", nova_pagina=True)
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
    pdf.secao("7. ANÁLISE DE AMBIENTE E MATRIZ SWOT", nova_pagina=True)
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
    pdf.secao("8. CANAIS DE AQUISIÇÃO E FUNIL DE VENDAS", nova_pagina=True)
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
    pdf.secao("9. INVESTIMENTOS, CAPITAL DE GIRO E FONTES DE RECURSOS", nova_pagina=True)
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
    pdf.secao("10. PRODUTOS, SERVIÇOS E CUSTOS UNITÁRIOS", nova_pagina=True)
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
    pdf.secao("11. CUSTOS DE COMERCIALIZAÇÃO, PESSOAL E CUSTOS FIXOS", nova_pagina=True)
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
    pdf.secao("12. PROJEÇÃO DE RESULTADOS E DRE", nova_pagina=True)
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
    pdf.secao("13. INDICADORES DE VIABILIDADE", nova_pagina=True)
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
        pdf.secao("14. GRÁFICOS", nova_pagina=True)
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
    pdf.secao("15. SIMULADOR DE CENÁRIOS E CHECAGENS", nova_pagina=True)
    pdf.tabela(extras["cenarios"], larguras=(45, 45, 45, 45), tam=7.5)
    pdf.subtitulo("Checagens de consistência")
    ch = extras["checagens"]
    pdf.tabela(ch[["Teste", "Valor apurado", "Situação", "Observação"]], larguras=(50, 40, 20, 70),
               alinhar=("LEFT", "LEFT", "CENTER", "LEFT"), tam=7.5)

    out = pdf.output()
    return bytes(out)

# ##############################################################################
# PARTE: pnbox_app.py
# ##############################################################################
# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE (Streamlit)
# ==============================================================================






st.set_page_config(page_title="PNBOX Acadêmico - Plano de Negócio", page_icon="📊", layout="wide")


# ------------------------------------------------------------------------------
# ENCERRAMENTO AUTOMÁTICO: quando todas as abas do navegador forem fechadas,
# o programa termina sozinho (só quando aberto pelo próprio arquivo / .bat).
# ------------------------------------------------------------------------------
TEMPO_ESPERA_FECHAR = 15  # segundos sem nenhuma aba aberta (dá tempo de recarregar a página)


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

ss = st.session_state
APP_DIR = Path(__file__).resolve().parent
# Modo web: no navegador (stlite/GitHub Pages) ou no Streamlit Community Cloud
MODO_WEB = sys.platform == "emscripten" or Path("/mount/src").exists() or os.environ.get("PNBOX_WEB") == "1"


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

if sys.platform == "emscripten":
    # Navegador (GitHub Pages): pasta guardada no próprio navegador (IndexedDB) deste computador
    PASTA = Path("/home/pyodide/pnbox")
elif MODO_WEB:
    # Cada visitante tem sua própria pasta temporária (um aluno não vê o projeto do outro)
    if "sessao_web" not in ss:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            ss.sessao_web = get_script_run_ctx().session_id
        except Exception:
            import uuid
            ss.sessao_web = uuid.uuid4().hex
    PASTA = Path(tempfile.gettempdir()) / "pnbox" / ss.sessao_web
elif (APP_DIR / "instalado.txt").exists():
    # Versão instalada pelo .exe: projetos em Documentos\PNBOX
    PASTA = Path.home() / "Documents" / "PNBOX"
else:
    PASTA = APP_DIR / "projetos"
PASTA.mkdir(parents=True, exist_ok=True)
ss.setdefault("ver", {})

UNIDADES = ["kg", "g", "L", "ml", "unidade", "dúzia", "pacote", "maço", "caixa", "kit", "galão", "lata", "par", "m", "hora"]
DIMENSOES = ["Pessoas", "Marketing", "Finanças", "Operação", "Político-Legal", "Econômico", "Social",
             "Tecnológico", "Ambiental", "Demográfico"]
CATEG_MATERIAIS = ["Matéria-prima", "Embalagem", "Limpeza e higiene", "Material de consumo", "Outros insumos"]
ETAPAS = ["Descoberta", "Consideração", "Decisão", "Compra e Pagamento", "Entrega", "Uso",
          "Avaliação", "Suporte", "Lealdade e advocacia"]


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


if not listar_projetos():
    antigo = PASTA.parent / "plano_negocio_completo.pnbox"
    if antigo.exists():
        shutil.copy(antigo, PASTA / "Projeto da versão anterior.pnbox")
    if VERSAO == "branco":
        criar_projeto_em_branco(PASTA / "Meu Projeto.pnbox", "Meu Projeto")
    else:
        criar_projeto_modelo(PASTA / "FitBox (modelo).pnbox")

projetos = listar_projetos()
if "db" not in ss or not Path(ss.db).exists():
    ss.db = str(projetos[0])

with st.sidebar:
    st.header("📁 Projetos")
    nomes = [p.stem for p in projetos]
    atual = Path(ss.db).stem
    escolha = st.selectbox("Projeto ativo", nomes, index=nomes.index(atual) if atual in nomes else 0)
    if escolha != atual:
        ss.db = str(PASTA / f"{escolha}.pnbox")
        st.rerun()

    with st.expander("➕ Novo projeto", expanded=False):
        novo_nome = st.text_input("Nome do novo projeto", placeholder="Ex.: Padaria Pão Nosso")
        tipo = st.radio("Começar a partir de", ["Em branco (do zero)", "Modelo FitBox (exemplo preenchido)",
                                                "Cópia do projeto atual"])
        if st.button("Criar projeto", type="primary", width="stretch"):
            if not novo_nome.strip():
                st.error("Informe um nome.")
            else:
                destino = caminho_livre(novo_nome)
                if tipo.startswith("Em branco"):
                    criar_projeto_em_branco(destino, novo_nome.strip())
                elif tipo.startswith("Modelo"):
                    criar_projeto_modelo(destino)
                else:
                    shutil.copy(ss.db, destino)
                c = conectar(destino)
                salvar_unica(c, "projeto", {"nome_negocio": novo_nome.strip()})
                c.close()
                ss.db = str(destino)
                st.rerun()

    with st.expander("📂 Abrir arquivo .pnbox"):
        arq = st.file_uploader("Selecione um arquivo .pnbox", type=["pnbox"])
        if arq is not None and st.button("Importar", width="stretch"):
            dados = arq.getvalue()
            if not dados.startswith(b"SQLite format 3"):
                st.error("Arquivo inválido: não é um projeto .pnbox.")
            else:
                destino = caminho_livre(Path(arq.name).stem)
                destino.write_bytes(dados)
                conectar(destino).close()   # atualiza a estrutura, se for de versão antiga
                ss.db = str(destino)
                st.rerun()

    with st.expander("🗑️ Excluir projeto atual"):
        confirma = st.checkbox(f"Confirmo a exclusão de “{atual}”")
        if st.button("Excluir", disabled=not confirma, width="stretch"):
            Path(ss.db).unlink(missing_ok=True)
            ss.pop("db")
            st.rerun()

    if MODO_WEB:
        st.warning("**Versão on-line:** o projeto fica só nesta aba do navegador. Ao terminar, use "
                   "**🗄️ Baixar arquivo do projeto (.pnbox)** na aba Relatório. Para continuar depois, "
                   "abra o arquivo em **📂 Abrir arquivo .pnbox**.")
    else:
        st.caption(f"Arquivos salvos em:\n`{PASTA}`")
    st.divider()
    if not MODO_WEB and st.button("⏻ Encerrar o programa", width="stretch"):
        st.success("Programa encerrado. Pode fechar esta aba.")
        threading.Timer(1.0, lambda: os._exit(0)).start()
        st.stop()


# ------------------------------------------------------------------------------
# CARREGAMENTO DOS DADOS
# ------------------------------------------------------------------------------
conn = conectar(Path(ss.db))
proj = ler_unica(conn, "projeto")
d = {t: ler(conn, t) for t in SCHEMA if t not in ("projeto", "proposta_valor")}
d["proposta_valor"] = ler_unica(conn, "proposta_valor")
PID = Path(ss.db).stem


ss.setdefault("buf", {})


def chave(tabela, extra=""):
    return f"{PID}_{tabela}_{extra}_{ss.ver.get((PID, tabela, extra), 0)}"


def _nova_versao(tabela, extra=""):
    ss.ver[(PID, tabela, extra)] = ss.ver.get((PID, tabela, extra), 0) + 1


def gravar(tabela, df, campo_chave, extra=""):
    salvar(conn, tabela, df, campo_chave)
    ss.buf.pop((PID, tabela, extra), None)
    _nova_versao(tabela, extra)
    st.toast("Dados salvos!", icon="✅")
    st.rerun()


def tabela_editavel(df_base, config, tabela, extra="", normalizar=None, altura=None):
    """
    Mostra uma tabela editável com os botões "Adicionar linha" e "Apagar linhas marcadas".
    As alterações ficam guardadas na sessão até o usuário clicar em Salvar.
    Devolve (dados_editados, coluna_onde_vai_o_botao_salvar).
    """
    kb = (PID, tabela, extra)
    normalizar = normalizar or (lambda x: x)
    base = ss.buf.get(kb, df_base).copy().reset_index(drop=True)
    base.insert(0, "Excluir", False)
    cfg = {"Excluir": st.column_config.CheckboxColumn("🗑️", help="Marque as linhas que deseja apagar", width="small")}
    cfg.update(config)
    ed = st.data_editor(base, num_rows="dynamic", width="stretch", hide_index=True, column_config=cfg,
                        key=chave(tabela, extra), **({"height": altura} if altura else {}))
    marcadas = ed["Excluir"].fillna(False).astype(bool)
    dados = normalizar(ed.drop(columns="Excluir"))
    b1, b2, b3, _ = st.columns([1.1, 1.4, 1.4, 2])
    if b1.button("➕ Adicionar linha", key="add_" + chave(tabela, extra), width="stretch"):
        vazia = normalizar(pd.DataFrame([{c: None for c in dados.columns}]))
        ss.buf[kb] = pd.concat([normalizar(ed[~marcadas].drop(columns="Excluir")), vazia], ignore_index=True)
        _nova_versao(tabela, extra)
        st.rerun()
    if b2.button(f"🗑️ Apagar marcadas ({int(marcadas.sum())})", key="del_" + chave(tabela, extra),
                 disabled=not marcadas.any(), width="stretch"):
        ss.buf[kb] = normalizar(ed[~marcadas].drop(columns="Excluir"))
        _nova_versao(tabela, extra)
        st.toast("Linhas removidas. Clique em Salvar para gravar.", icon="🗑️")
        st.rerun()
    if kb in ss.buf:
        st.caption("⚠️ Há alterações ainda não salvas nesta tabela.")
    return dados, b3


def editor(tabela, config, campo_chave, rotulo="💾 Salvar", altura=None):
    """Tabela editável com botões de adicionar, apagar e salvar. Devolve os dados editados."""
    ed, col_salvar = tabela_editavel(d[tabela], config, tabela, normalizar=lambda x: limpar(x, tabela), altura=altura)
    if col_salvar.button(rotulo, key="bt_" + chave(tabela), type="primary", width="stretch"):
        gravar(tabela, ed, campo_chave)
    d[tabela] = ed
    return ed


T = st.column_config.TextColumn
N = st.column_config.NumberColumn
S = st.column_config.SelectboxColumn
CK = st.column_config.CheckboxColumn


def moeda(rot, **kw):
    return N(rot, min_value=0.0, format="R$ %.2f", **kw)


def nota(rot):
    return N(rot, min_value=0, max_value=10, step=1, format="%d")


st.title(f"📊 {proj['nome_negocio']}")
st.caption("Plano de Negócio e Viabilidade Econômico-Financeira · estrutura inspirada no PNBOX/Sebrae")

abas = st.tabs([
    "1. Projeto", "2. Segmentos e Valor", "3. Personas", "4. Jornada", "5. Concorrência",
    "6. SWOT", "7. Canais e Funil", "8. Investimentos", "9. Produtos e Materiais",
    "10. Custos", "11. Resultados", "12. Gráficos", "13. Relatório",
])

# ==============================================================================
# ABA 1 - PROJETO E SÓCIOS
# ==============================================================================
with abas[0]:
    st.subheader("Informações do plano")
    with st.form("f_proj"):
        c1, c2, c3 = st.columns(3)
        v = {}
        v["nome_negocio"] = c1.text_input("Nome do negócio", proj["nome_negocio"])
        objs = ["Criar um novo negócio", "Expandir um negócio existente", "Captar recursos", "Outro"]
        v["objetivo"] = c2.selectbox("Objetivo", objs, index=objs.index(proj["objetivo"]) if proj["objetivo"] in objs else 0)
        v["periodo_anos"] = c3.number_input("Período de planejamento (anos)", 1, 3, int(proj["periodo_anos"] or 3))
        regs = ["Simples Nacional", "MEI", "Lucro Presumido", "Lucro Real"]
        v["enquadramento"] = c1.selectbox("Enquadramento tributário", regs,
                                          index=regs.index(proj["enquadramento"]) if proj["enquadramento"] in regs else 0)
        portes = ["MEI", "Microempresa (ME)", "Pequena Empresa (EPP)", "Médio/Grande Porte"]
        v["porte"] = c2.selectbox("Porte", portes, index=portes.index(proj["porte"]) if proj["porte"] in portes else 1)
        setores = ["Comércio", "Serviço", "Indústria", "Agronegócio", "Alimentos e bebidas", "Bares e restaurantes",
                   "Varejo", "Moda", "Saúde e bem-estar", "Tecnologia", "Educação", "Turismo"]
        atuais = [x.strip() for x in proj["setor_atuacao"].split(";") if x.strip()]
        v["setor_atuacao"] = "; ".join(c3.multiselect("Setores de atuação", sorted(set(setores + atuais)), atuais))
        v["cidade_uf"] = c1.text_input("Cidade / UF", proj["cidade_uf"])
        v["contatos"] = c2.text_input("Contatos (site, redes, e-mail)", proj["contatos"])
        v["aliquota_imposto_pct"] = c3.number_input("Impostos sobre vendas (%)", 0.0, 40.0,
                                                    proj["aliquota_imposto_pct"], 0.1,
                                                    help="Alíquota efetiva do Simples ou soma de PIS/COFINS/ISS/ICMS etc.")
        st.markdown("**Premissas de projeção**")
        p1, p2, p3, p4, p5, p6 = st.columns(6)
        v["crescimento_mensal_ano1_pct"] = p1.number_input("Cresc. mensal Ano 1 (%)", -10.0, 30.0, proj["crescimento_mensal_ano1_pct"], 0.5)
        v["crescimento_anual_pct"] = p2.number_input("Cresc. anual Anos 2-3 (%)", -50.0, 200.0, proj["crescimento_anual_pct"], 1.0)
        v["inflacao_custos_anual_pct"] = p3.number_input("Inflação custos fixos (% a.a.)", 0.0, 50.0, proj["inflacao_custos_anual_pct"], 0.5)
        v["tma_anual_pct"] = p4.number_input("TMA (% a.a.)", 0.0, 60.0, proj["tma_anual_pct"], 0.5)
        v["pme_dias"] = p5.number_input("Prazo médio de estoques (dias)", 0.0, 365.0, proj["pme_dias"], 1.0)
        v["reserva_caixa_extra"] = p6.number_input("Reserva financeira (R$)", 0.0, 1e8, proj["reserva_caixa_extra"], 1000.0)
        v["recuperar_giro"] = int(st.checkbox("No cálculo do VPL, recuperar o capital de giro no último ano",
                                              bool(proj["recuperar_giro"])))
        if st.form_submit_button("💾 Salvar informações", type="primary"):
            salvar_unica(conn, "projeto", v)
            st.toast("Informações salvas!", icon="✅")
            st.rerun()

    st.subheader("Sócios")
    editor("socios", {
        "nome": T("Nome", required=True), "funcao": T("Função"), "contato": T("Contato"),
        "perfil": T("Minicurrículo / atribuições", width="large"),
        "capital_integralizado": moeda("Capital investido (R$)")}, "nome")
    st.caption(f"Capital próprio total: **{brl(d['socios']['capital_integralizado'].sum())}**")

# ==============================================================================
# ABA 2 - SEGMENTOS E PROPOSTA DE VALOR
# ==============================================================================
with abas[1]:
    st.subheader("Segmentos de clientes (notas de 1 a 10)")
    sg = editor("segmentos", {
        "nome": T("Segmento", required=True), "descricao": T("Descrição", width="large"),
        "nota_tamanho": nota("É grande o suficiente?"), "nota_lucro": nota("Potencial de lucro"),
        "nota_longevidade": nota("Não tende a desaparecer?"), "nota_concorrencia": nota("Poucos concorrentes?"),
        "nota_aderencia": nota("Oferta atende o segmento?")}, "nome")
    if not sg.empty:
        rk = sg.assign(Média=sg.iloc[:, 2:].mean(axis=1)).sort_values("Média", ascending=False)
        rk.insert(0, "Prioridade", [f"{i}º" for i in range(1, len(rk) + 1)])
        st.dataframe(rk[["Prioridade", "nome", "Média"]].rename(columns={"nome": "Segmento"}),
                     hide_index=True, width="stretch")

    st.divider()
    st.subheader("Proposta de valor")
    pv = d["proposta_valor"]
    with st.form("f_pv"):
        a = st.text_input("Nossa empresa ajuda (público / persona)", pv["segmento_alvo"])
        b = st.text_area("que desejam (benefícios funcionais, emocionais e sociais)", pv["beneficios_desejados"])
        c = st.text_area("por meio de (produtos, serviços, canais)", pv["meios_entrega"])
        e = st.text_area("diferentemente de (concorrentes / alternativas)", pv["diferencial_competitivo"])
        f = st.text_area("os nossos produtos ou serviços oferecem", pv["entrega_produtos_servicos"])
        if st.form_submit_button("💾 Salvar proposta de valor", type="primary"):
            salvar_unica(conn, "proposta_valor", {"segmento_alvo": a, "beneficios_desejados": b, "meios_entrega": c,
                                                  "diferencial_competitivo": e, "entrega_produtos_servicos": f})
            st.rerun()
    st.info(f"Nossa empresa (**{proj['nome_negocio']}**) ajuda **{pv['segmento_alvo']}** que desejam "
            f"**{pv['beneficios_desejados']}** por meio de **{pv['meios_entrega']}**. Diferentemente de "
            f"**{pv['diferencial_competitivo']}**, os nossos produtos ou serviços oferecem **{pv['entrega_produtos_servicos']}**")

# ==============================================================================
# ABA 3 - PERSONAS (formulário completo, uma persona por vez)
# ==============================================================================
CAMPOS_PERSONA = [
    ("nome", "Nome", "txt"), ("idade", "Idade", "int"), ("profissao", "Profissão", "txt"),
    ("estado_civil", "Estado civil", "txt"), ("renda_mensal", "Renda mensal (R$)", "num"),
    ("escolaridade", "Escolaridade", "txt"), ("cidade", "Cidade", "txt"),
    ("descricao", "Descrição da persona", "area"), ("citacao", "Frase que a representa", "area"),
    ("habitos", "Hábitos e comportamentos", "area"), ("caracteristicas", "Características (#tags)", "txt"),
    ("fluencia_digital", "Fluência digital (computador, smartphone, apps, compras on-line)", "txt"),
    ("redes_sociais", "Redes sociais utilizadas", "txt"), ("quem_influencia", "Quem a influencia", "txt"),
    ("rotina_manha", "Rotina - manhã", "area"), ("rotina_tarde", "Rotina - tarde", "area"),
    ("rotina_noite", "Rotina - noite", "area"), ("ganhos", "Ganhos esperados", "area"),
    ("desafios", "Desafios / dores", "area"), ("duvidas_objecoes", "Dúvidas ou objeções", "area"),
    ("aborrecimentos", "Aborrecimentos que devem ser evitados", "area"),
    ("experiencia_ideal", "Experiência ideal e como o negócio ajuda", "area"),
]
with abas[2]:
    st.subheader("Personas")
    pers = pd.read_sql("SELECT * FROM personas ORDER BY id", conn)
    opcoes = ["➕ Nova persona"] + [f"{r['id']} - {r['nome']}" for _, r in pers.iterrows()]
    sel = st.selectbox("Selecione a persona", opcoes, key=f"{PID}_sel_persona")
    atual_p = {} if sel.startswith("➕") else pers[pers["id"] == int(sel.split(" - ")[0])].iloc[0].to_dict()
    with st.form(f"f_persona_{sel}"):
        valores = {}
        cols = st.columns(4)
        for i, (k, rot, tipo) in enumerate(CAMPOS_PERSONA[:7]):
            val = atual_p.get(k)
            alvo = cols[i % 4]
            if tipo == "int":
                valores[k] = alvo.number_input(rot, 0, 120, int(val or 30))
            elif tipo == "num":
                valores[k] = alvo.number_input(rot, 0.0, 1e7, float(val or 0), 100.0)
            else:
                valores[k] = alvo.text_input(rot, val or "")
        c1, c2 = st.columns(2)
        for i, (k, rot, tipo) in enumerate(CAMPOS_PERSONA[7:]):
            alvo = c1 if i % 2 == 0 else c2
            val = atual_p.get(k) or ""
            valores[k] = alvo.text_area(rot, val, height=90) if tipo == "area" else alvo.text_input(rot, val)
        b1, b2 = st.columns([1, 5])
        salvar_p = b1.form_submit_button("💾 Salvar persona", type="primary")
        excluir_p = b2.form_submit_button("🗑️ Excluir persona", disabled=not atual_p)
    if salvar_p:
        if not valores["nome"].strip():
            st.error("Informe o nome da persona.")
        else:
            cols_p = list(valores)
            if atual_p:
                conn.execute(f"UPDATE personas SET {', '.join(c + '=?' for c in cols_p)} WHERE id=?",
                             [valores[c] for c in cols_p] + [atual_p["id"]])
            else:
                conn.execute(f"INSERT INTO personas ({', '.join(cols_p)}) VALUES ({', '.join('?' * len(cols_p))})",
                             [valores[c] for c in cols_p])
            conn.commit()
            st.rerun()
    if excluir_p and atual_p:
        conn.execute("DELETE FROM personas WHERE id=?", (atual_p["id"],))
        conn.commit()
        st.rerun()

# ==============================================================================
# ABA 4 - JORNADA DO CLIENTE
# ==============================================================================
with abas[3]:
    st.subheader("Jornada do cliente (antes, durante e depois da compra)")
    nomes_pers = d["personas"]["nome"].tolist() or ["(cadastre uma persona)"]
    editor("jornada_cliente", {
        "persona_nome": S("Persona", options=nomes_pers, required=True),
        "macro_fase": S("Fase", options=["Antes", "Durante", "Depois"], required=True),
        "etapa": S("Etapa", options=ETAPAS, required=True),
        "acoes_cliente": T("Ações do cliente", width="medium"), "pensamentos": T("Pensamentos"),
        "pontos_contato": T("Pontos de contato"),
        "estado_emocional": S("Emoção", options=["Entusiasmado", "Satisfeito", "Neutro", "Insatisfeito", "Frustrado"]),
        "oportunidades": T("Oportunidades de melhoria", width="medium")}, "persona_nome")
    emo = {"Entusiasmado": 2, "Satisfeito": 1, "Neutro": 0, "Insatisfeito": -1, "Frustrado": -2}
    jc = d["jornada_cliente"]
    if not jc.empty:
        st.markdown("**Curva emocional da jornada**")
        graf = jc.assign(Emoção=jc["estado_emocional"].map(emo)).pivot_table(
            index="etapa", columns="persona_nome", values="Emoção", aggfunc="mean")
        graf = graf.reindex([e for e in ETAPAS if e in graf.index])
        st.line_chart(graf)

# ==============================================================================
# ABA 5 - CONCORRÊNCIA
# ==============================================================================
with abas[4]:
    st.subheader("Concorrentes")
    conc = editor("concorrentes", {"nome": T("Nome do concorrente", required=True),
                                   "observacoes": T("Observações", width="large")}, "nome",
                  rotulo="💾 Salvar concorrentes")
    empresas = ["Nossa Empresa"] + [n for n in conc["nome"] if str(n).strip()]
    notas = d["concorrencia_notas"]
    atributos = list(dict.fromkeys(notas["atributo"]))
    piv = pd.DataFrame({"Atributo": pd.Series(atributos, dtype=object)})
    for emp in empresas:
        mapa = notas[notas["empresa"] == emp].set_index("atributo")["nota"].to_dict()
        piv[emp] = pd.Series([mapa.get(a, None) for a in atributos], dtype=float)
    st.subheader("Avaliação dos atributos (notas de 0 a 10)")
    cfg = {"Atributo": T("Atributo avaliado", required=True)}
    cfg.update({e: N(e, min_value=0, max_value=10, step=0.5) for e in empresas})
    def _norm_piv(x):
        x = x.copy()
        x["Atributo"] = x["Atributo"].astype(object).where(x["Atributo"].notna(), "")
        for e in empresas:
            x[e] = pd.to_numeric(x[e], errors="coerce") if e in x else float("nan")
        return x[["Atributo"] + empresas]
    piv_ed, col_s = tabela_editavel(piv, cfg, "concorrencia_notas", "|".join(empresas), normalizar=_norm_piv)
    piv_ed = piv_ed[piv_ed["Atributo"].astype(str).str.strip().ne("")]
    if col_s.button("💾 Salvar notas", key="bt_notas", type="primary", width="stretch"):
        longo = piv_ed.melt(id_vars="Atributo", var_name="empresa", value_name="nota").rename(columns={"Atributo": "atributo"})
        gravar("concorrencia_notas", longo, "atributo", "|".join(empresas))
    piv_num = piv_ed.copy()
    for e in empresas:
        piv_num[e] = pd.to_numeric(piv_num[e], errors="coerce").fillna(0)
    ss_piv = piv_num
    if not piv_num.empty:
        medias = piv_num[empresas].mean()
        mc = st.columns(min(len(empresas), 6))
        for i, e in enumerate(empresas[:6]):
            mc[i].metric(f"Média - {e}", num_br(medias[e]))
        st.bar_chart(piv_num.set_index("Atributo")[empresas])
        if len(empresas) > 1:
            outros = piv_num.set_index("Atributo")[empresas[1:]].max(axis=1)
            nossa = piv_num.set_index("Atributo")["Nossa Empresa"]
            st.success("**Forças frente à concorrência:** " + (", ".join(nossa.index[nossa > outros]) or "nenhuma"))
            st.warning("**Fraquezas frente à concorrência:** " + (", ".join(nossa.index[nossa < outros]) or "nenhuma"))

# ==============================================================================
# ABA 6 - AMBIENTE E MATRIZ SWOT
# ==============================================================================
def html_swot(sw: pd.DataFrame) -> str:
    def lista(nat):
        itens = sw[sw["natureza"] == nat]
        if itens.empty:
            return "<i>nenhum fator</i>"
        return "<br>".join(("⭐ " if int(r["priorizado"]) else "• ") + f"{r['fator']} <small>({r['impacto']})</small>"
                           for _, r in itens.iterrows())

    cel = "padding:14px 16px;border-radius:2px;min-height:190px;font-size:14px;line-height:1.5;"
    lat = "writing-mode:vertical-rl;transform:rotate(180deg);font-weight:700;letter-spacing:2px;" \
          "display:flex;align-items:center;justify-content:center;"
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


with abas[5]:
    st.subheader("Fatores internos e externos")
    sw = editor("swot", {
        "ambiente": S("Ambiente", options=["Interno", "Externo"], required=True),
        "dimensao": S("Dimensão", options=DIMENSOES),
        "fator": T("Fator", required=True),
        "natureza": S("Classificação", options=["Força", "Fraqueza", "Oportunidade", "Ameaça"], required=True),
        "impacto": S("Impacto", options=["Alto", "Baixo"], required=True),
        "diagnostico": T("Diagnóstico", width="medium"),
        "plano_acao": T("Como potencializar / corrigir / aproveitar / neutralizar", width="medium"),
        "priorizado": CK("Priorizado")}, "fator", rotulo="💾 Salvar fatores")
    st.subheader("Matriz SWOT")
    st.markdown(html_swot(sw), unsafe_allow_html=True)
    st.caption("⭐ fator priorizado · entre parênteses, o impacto")

    diag = pontuacao_swot(sw)
    st.subheader("Matriz de estratégias")
    q = diag["quadrantes"]
    g1, g2 = st.columns(2)
    for col, k in zip([g1, g2, g1, g2], ["FO", "FrO", "FA", "FrA"]):
        nome_q, comb, _ = QUADRANTES[k]
        destaque = "border:3px solid #26548c;" if k == diag["predominante"] else "border:1px solid #ccc;"
        col.markdown(f"<div style='{destaque}padding:12px;border-radius:6px;margin-bottom:8px;text-align:center;'>"
                     f"<div style='font-size:28px;font-weight:700;'>{q[k]}</div><b>{comb}</b><br>{nome_q}</div>",
                     unsafe_allow_html=True)
    if diag["predominante"]:
        nome_q, comb, txt = QUADRANTES[diag["predominante"]]
        st.success(f"**Situação do negócio: {nome_q}.** {txt}")
    st.caption("Pontuação: impacto alto = 2 pontos; impacto baixo = 1 ponto.")

# ==============================================================================
# ABA 7 - CANAIS E FUNIL
# ==============================================================================
def tabela_funil(fu: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "Etapa": fu["etapa"], "Ação": fu["acao"], "Orçamento": fu["orcamento"].map(brl),
        "Alcançados": fu["alcancados"].map(lambda v: num_br(v, 0)),
        "Convertidos": fu["convertidos"].map(lambda v: num_br(v, 0)),
        "Conversão": [pct(div(c, a)) for c, a in zip(fu["convertidos"], fu["alcancados"])],
        "Custo por conversão": [brl(div(o, c)) for o, c in zip(fu["orcamento"], fu["convertidos"])]})


with abas[6]:
    st.subheader("Avaliação dos canais de aquisição (notas de 0 a 10)")
    cn = editor("canais", {"canal": T("Canal", required=True), "atratividade": nota("Atratividade"),
                           "alcance": nota("Alcance esperado"), "cliente": nota("Custo por cliente"),
                           "relacionamento": nota("Relacionamento"), "viabilidade": nota("Viabilidade")}, "canal")
    if not cn.empty:
        cn2 = cn.assign(Média=cn.iloc[:, 1:].mean(axis=1)).sort_values("Média", ascending=False)
        st.success("**Canais de tração (3 maiores médias):** " + ", ".join(cn2["canal"].head(3)))
    st.subheader("Funil de vendas")
    fu = editor("funil_vendas", {
        "etapa": S("Etapa", options=["Topo do Funil", "Meio do Funil", "Fundo do Funil"], required=True),
        "acao": T("Ação de venda", width="large"), "orcamento": moeda("Orçamento (R$)"),
        "alcancados": N("Pessoas alcançadas", min_value=0, format="%d"),
        "convertidos": N("Pessoas que avançaram/compraram", min_value=0, format="%d")}, "etapa")
    if not fu.empty:
        st.dataframe(tabela_funil(fu), hide_index=True, width="stretch")
        fundo = fu[fu["etapa"] == "Fundo do Funil"]["convertidos"].sum()
        cac = div(fu["orcamento"].sum(), fundo)
        m1, m2, m3 = st.columns(3)
        m1.metric("Orçamento total", brl(fu["orcamento"].sum()))
        m2.metric("Clientes conquistados (fundo)", num_br(fundo, 0))
        m3.metric("CAC - custo de aquisição de cliente", brl(cac))

# ==============================================================================
# ABA 8 - INVESTIMENTOS, CAPITAL DE GIRO E FONTES
# ==============================================================================
with abas[7]:
    st.subheader("Investimentos fixos (máquinas, equipamentos, móveis, veículos, instalações)")
    invf = editor("inv_fixos", {
        "descricao": T("Descrição", required=True, width="large"), "quantidade": N("Qtd.", min_value=0.0),
        "valor_unitario": moeda("Valor unitário (R$)"),
        "vida_util_anos": N("Vida útil (anos)", min_value=0, max_value=50, step=1,
                            help="0 = não deprecia (ex.: terrenos)"),
        "valor_residual_pct": N("Valor residual (%)", min_value=0.0, max_value=100.0)}, "descricao")
    tot_fixo = float((invf["quantidade"] * invf["valor_unitario"]).sum())
    dep = sum((q * vu * (1 - r / 100) / vu_) if vu_ > 0 else 0 for q, vu, vu_, r in
              zip(invf["quantidade"], invf["valor_unitario"], invf["vida_util_anos"], invf["valor_residual_pct"]))
    a, b = st.columns(2)
    a.metric("Total dos investimentos fixos", brl(tot_fixo))
    b.metric("Depreciação mensal", brl(dep / 12))

    st.subheader("Investimentos pré-operacionais")
    ip = editor("inv_pre", {"descricao": T("Descrição", required=True, width="large"),
                            "valor": moeda("Valor (R$)")}, "descricao")
    st.metric("Total pré-operacional", brl(ip["valor"].sum()))
    st.info("O **estoque inicial** é calculado automaticamente na aba **9. Produtos e Materiais**, a partir da "
            "quantidade inicial de cada material e de cada mercadoria de revenda.")

    st.subheader("Capital de giro: prazos médios")
    editor("capital_giro_prazos", {
        "tipo": S("Tipo", options=["Recebimento", "Pagamento"], required=True),
        "prazo_dias": N("Prazo (dias)", min_value=0, max_value=360, step=1),
        "percentual": N("% das vendas/compras", min_value=0.0, max_value=100.0)}, "tipo")
    st.caption("Recebimento: como os clientes pagam (0 dias = à vista). Pagamento: como a empresa paga os fornecedores. "
               "Em cada tipo, os percentuais devem somar 100%.")

    st.subheader("Fontes de recursos de terceiros (empréstimos e financiamentos)")
    editor("emprestimos", {"descricao": T("Descrição", required=True, width="large"),
                           "valor": moeda("Valor (R$)"), "taxa_anual_pct": N("Juros (% a.a.)", min_value=0.0),
                           "prazo_meses": N("Prazo (meses)", min_value=1, step=1)}, "descricao")
    st.caption("Parcelas calculadas pela Tabela Price. O capital próprio vem da aba **1. Projeto** (sócios).")

# ==============================================================================
# ABA 9 - PRODUTOS, MATERIAIS E FICHA TÉCNICA
# ==============================================================================
with abas[8]:
    st.subheader("① Cadastro de materiais e insumos")
    st.caption("Matérias-primas, embalagens e insumos usados na fabricação e nos serviços. "
               "A quantidade inicial forma o estoque inicial.")
    mat = editor("materiais", {"nome": T("Material / insumo", required=True, width="large"),
                               "categoria": S("Categoria", options=CATEG_MATERIAIS, required=True),
                               "unidade": S("Unidade", options=UNIDADES, required=True),
                               "valor_unitario": moeda("Valor por unidade (R$)"),
                               "qtd_estoque_inicial": N("Qtd. no estoque inicial", min_value=0.0)}, "nome",
                 rotulo="💾 Salvar materiais")

    st.subheader("② Produtos e serviços")
    st.caption("**Revenda:** informe o valor de aquisição. **Fabricação própria e serviços:** monte a ficha "
               "técnica abaixo (se não houver ficha, vale o custo informado).")
    cat = editor("catalogo", {
        "categoria": S("Categoria", options=CATEGORIAS, required=True),
        "descricao": T("Produto / serviço", required=True, width="large"),
        "preco_venda": moeda("Preço de venda (R$)"),
        "custo_aquisicao": moeda("Valor de aquisição / custo informado (R$)"),
        "qtd_mensal": N("Qtd. vendida no Mês 1", min_value=0, step=1),
        "qtd_estoque_inicial": N("Estoque inicial (só revenda)", min_value=0, step=1)}, "descricao",
        rotulo="💾 Salvar produtos e serviços")

    st.subheader("③ Ficha técnica")
    fabricaveis = cat[cat["categoria"] != "Revenda"]["descricao"].tolist()
    if not fabricaveis:
        st.info("Cadastre produtos de fabricação própria ou serviços para montar a ficha técnica.")
    elif mat.empty:
        st.info("Cadastre os materiais (item ①) para montar a ficha técnica.")
    else:
        prod = st.selectbox("Produto / serviço", fabricaveis, key=f"{PID}_ft_prod")
        ft_all = d["ficha_tecnica"]
        ft = ft_all[ft_all["produto"] == prod][["material", "quantidade"]].reset_index(drop=True)
        def _norm_ft(x):
            x = x[["material", "quantidade"]].copy()
            x["quantidade"] = pd.to_numeric(x["quantidade"], errors="coerce").fillna(0.0)
            return x
        ft_ed, col_ft = tabela_editavel(ft, {"material": S("Material", options=mat["nome"].tolist(), required=True),
                                             "quantidade": N("Quantidade (na unidade do material)", min_value=0.0, format="%.3f")},
                                        "ficha_tecnica", prod, normalizar=_norm_ft)
        ft_ed = ft_ed.dropna(subset=["material"])
        ft_ed = ft_ed[ft_ed["material"].astype(str).str.strip() != ""]
        ft_ed["quantidade"] = pd.to_numeric(ft_ed["quantidade"], errors="coerce").fillna(0)
        precos = dict(zip(mat["nome"], mat["valor_unitario"]))
        unid = dict(zip(mat["nome"], mat["unidade"]))
        calc = ft_ed.assign(Unidade=ft_ed["material"].map(unid),
                            **{"Valor unit.": ft_ed["material"].map(precos).fillna(0)})
        calc["Custo"] = calc["quantidade"] * calc["Valor unit."]
        custo_p = float(calc["Custo"].sum())
        preco_p = float(cat.loc[cat["descricao"] == prod, "preco_venda"].iloc[0])
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Custo unitário", brl(custo_p))
        k2.metric("Preço de venda", brl(preco_p))
        k3.metric("Margem bruta", pct(div(preco_p - custo_p, preco_p)))
        k4.metric("Markup", f"{num_br(div(preco_p, custo_p))}x" if custo_p else "n.a.")
        if not calc.empty:
            st.dataframe(calc.assign(**{"Valor unit.": calc["Valor unit."].map(brl), "Custo": calc["Custo"].map(brl)}),
                         hide_index=True, width="stretch")
        if col_ft.button("💾 Salvar ficha técnica", key="bt_ft", type="primary", width="stretch"):
            nova = pd.concat([ft_all[ft_all["produto"] != prod], ft_ed.assign(produto=prod)], ignore_index=True)
            gravar("ficha_tecnica", nova, "material", prod)
        d["ficha_tecnica"] = pd.concat([ft_all[ft_all["produto"] != prod], ft_ed.assign(produto=prod)], ignore_index=True)
        d["ficha_tecnica"] = limpar(d["ficha_tecnica"], "ficha_tecnica")

    st.subheader("④ Resumo de custos e margens")
    cu = custos_unitarios(cat, mat, d["ficha_tecnica"])
    if not cu.empty:
        st.dataframe(pd.DataFrame({
            "Categoria": cu["categoria"], "Produto / serviço": cu["descricao"], "Preço": cu["preco_venda"].map(brl),
            "Custo unit.": cu["custo_unit"].map(brl), "Origem do custo": cu["origem_custo"],
            "Margem": cu["margem_pct"].map(pct), "Markup": cu["markup"].map(lambda v: f"{num_br(v)}x" if v == v else "n.a."),
            "Qtd./mês": cu["qtd_mensal"].map(lambda v: num_br(v, 0)), "Receita/mês": cu["receita_mensal"].map(brl),
            "Custo/mês": cu["custo_mensal"].map(brl)}), hide_index=True, width="stretch")
        r1, r2, r3 = st.columns(3)
        r1.metric("Receita do Mês 1", brl(cu["receita_mensal"].sum()))
        r2.metric("Custo direto do Mês 1", brl(cu["custo_mensal"].sum()))
        est = float((mat["qtd_estoque_inicial"] * mat["valor_unitario"]).sum() +
                    (cu.loc[cu["categoria"] == "Revenda", "qtd_estoque_inicial"] * cu.loc[cu["categoria"] == "Revenda", "custo_unit"]).sum())
        r3.metric("Estoque inicial (materiais + revenda)", brl(est))

# ==============================================================================
# ABA 10 - CUSTOS
# ==============================================================================
with abas[9]:
    st.subheader("Custos de comercialização (% sobre a receita)")
    st.caption(f"Impostos sobre vendas ({proj['enquadramento']}): **{num_br(proj['aliquota_imposto_pct'])}%** "
               "(altere na aba 1).")
    cv = editor("custos_variaveis", {"descricao": T("Descrição", required=True, width="large"),
                                     "percentual": N("% sobre a receita", min_value=0.0, max_value=100.0)}, "descricao")
    st.metric("Total de custos variáveis (impostos + comercialização)",
              f"{num_br(proj['aliquota_imposto_pct'] + cv['percentual'].sum())}%")

    st.subheader("Mão de obra")
    pe = editor("pessoal", {"cargo": T("Cargo", required=True, width="large"),
                            "num_empregados": N("Nº de empregados", min_value=0, step=1),
                            "salario_mensal": moeda("Salário (R$)"), "encargos_pct": N("Encargos (%)", min_value=0.0),
                            "beneficios_unit": moeda("Benefícios por empregado (R$)")}, "cargo")
    tot_pe = float((pe["num_empregados"] * pe["salario_mensal"] * (1 + pe["encargos_pct"] / 100) +
                    pe["num_empregados"] * pe["beneficios_unit"]).sum())
    st.metric("Custo mensal com pessoal", brl(tot_pe))
    st.caption("Referência: no Simples Nacional (anexos I a III), encargos de cerca de 30% a 40% (FGTS, férias, 13º e INSS do 13º); "
               "fora do Simples, em geral acima de 60%. Confirme com o contador.")

    st.subheader("Custos fixos operacionais mensais")
    cf = editor("custos_fixos", {"descricao": T("Descrição", required=True, width="large"),
                                 "valor_mensal": moeda("Valor mensal (R$)")}, "descricao")
    st.metric("Custos fixos (sem pessoal e depreciação)", brl(cf["valor_mensal"].sum()))

# ==============================================================================
# MOTOR FINANCEIRO (usa os dados atuais, inclusive edições ainda não salvas)
# ==============================================================================
res = executar_motor(proj, d)
df_dre = tabela_dre(res)
df_ind = tabela_indicadores(res)
df_cen = cenarios(res, proj)
df_chk = tabela_checagens(res, d)
A1 = res["anos"][1]

# ==============================================================================
# ABA 11 - RESULTADOS E INDICADORES
# ==============================================================================
with abas[10]:
    n_ver = int((df_chk["Situação"] == "VERIFICAR").sum())
    if n_ver:
        st.warning(f"⚠️ {n_ver} checagem(ns) pedem atenção. Veja a seção **Checagens** no fim desta aba.")
    st.subheader("Resumo financeiro (Ano 1)")
    rs = res["resumo"]
    k = st.columns(4)
    k[0].metric("Investimento total", brl(res["investimento_total"]))
    k[1].metric("Receita anual", brl(rs["receita_anual"]))
    k[2].metric("Resultado operacional", brl(rs["resultado_operacional"]))
    k[3].metric("Lucratividade", pct(rs["lucratividade"]))
    k = st.columns(4)
    k[0].metric("Rentabilidade", pct(rs["rentabilidade"]))
    k[1].metric("Payback simples", f"{num_br(rs['payback_meses'], 1)} meses" if rs["payback_meses"] == rs["payback_meses"] else "não recupera")
    k[2].metric(f"VPL (TMA {num_br(proj['tma_anual_pct'], 1)}%)", brl(res["vpl"]))
    k[3].metric("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a.")
    k = st.columns(4)
    k[0].metric("PE contábil (anual)", brl(A1["pe_contabil"]))
    k[1].metric("PE financeiro (anual)", brl(A1["pe_financeiro"]))
    k[2].metric("PE econômico (anual)", brl(A1["pe_economico"]))
    k[3].metric("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "> 36 meses")

    st.subheader("Investimento total e fontes")
    it = res["investimento_total"]
    c1, c2 = st.columns(2)
    c1.dataframe(pd.DataFrame({
        "Aplicação": ["Investimentos fixos", "Pré-operacionais", "Estoque inicial", "Capital de giro", "Total"],
        "Valor": [brl(x) for x in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)],
        "%": [pct(div(x, it)) for x in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)]}),
        hide_index=True, width="stretch")
    c2.dataframe(pd.DataFrame({
        "Fonte": ["Capital próprio (sócios)", "Capital de terceiros (empréstimos)", "Total", "Diferença (fontes − aplicações)"],
        "Valor": [brl(res["capital_proprio"]), brl(res["capital_terceiros"]), brl(res["fontes"]), brl(res["fontes"] - it)]}),
        hide_index=True, width="stretch")
    st.caption(f"Capital de giro: ciclo financeiro de {num_br(res['ciclo_fin'], 1)} dias (PME {num_br(res['pme'], 0)} + "
               f"PMR {num_br(res['pmr'], 1)} − PMP {num_br(res['pmp'], 1)}); NCG {brl(res['ncg'])} + reserva {brl(res['reserva'])}.")

    st.subheader("Projeção mensal do Ano 1")
    mm = res["meses"]
    st.line_chart(mm.set_index("Mês")[["Receita", "Custos variáveis", "Custos fixos", "Resultado operacional"]])
    with st.expander("Ver tabela mensal"):
        st.dataframe(mm.assign(**{c: mm[c].map(brl) for c in mm.columns[1:]}), hide_index=True, width="stretch")

    st.subheader("DRE projetada (Anos 1 a 3)")
    st.dataframe(df_dre, hide_index=True, width="stretch", height=560)
    st.bar_chart(pd.DataFrame({"Receita": [res["anos"][a]["rec_bruta"] for a in (1, 2, 3)],
                               "EBITDA": [res["anos"][a]["ebitda"] for a in (1, 2, 3)],
                               "Resultado líquido": [res["anos"][a]["resultado"] for a in (1, 2, 3)]},
                              index=["Ano 1", "Ano 2", "Ano 3"]), stack=False)

    st.subheader("Matriz de indicadores (5 blocos)")
    blocos = ["Todos"] + list(df_ind["Bloco"].unique())
    bsel = st.selectbox("Bloco", blocos)
    st.dataframe(df_ind if bsel == "Todos" else df_ind[df_ind["Bloco"] == bsel], hide_index=True, width="stretch")

    st.subheader("Simulador de cenários")
    with st.form("f_cen"):
        c = st.columns(4)
        vc = {"cen_otim_receita_pct": c[0].number_input("Otimista: receita maior em (%)", 0.0, 100.0, proj["cen_otim_receita_pct"]),
              "cen_otim_custo_pct": c[1].number_input("Otimista: custos menores em (%)", 0.0, 100.0, proj["cen_otim_custo_pct"]),
              "cen_pess_receita_pct": c[2].number_input("Pessimista: receita menor em (%)", 0.0, 100.0, proj["cen_pess_receita_pct"]),
              "cen_pess_custo_pct": c[3].number_input("Pessimista: custos maiores em (%)", 0.0, 100.0, proj["cen_pess_custo_pct"])}
        if st.form_submit_button("Atualizar cenários"):
            salvar_unica(conn, "projeto", vc)
            st.rerun()
    st.dataframe(df_cen, hide_index=True, width="stretch")

    st.subheader("Checagens de consistência")
    st.dataframe(df_chk.style.map(lambda v: "color:#b00020;font-weight:700" if v == "VERIFICAR" else
                                  ("color:#137333;font-weight:700" if v == "OK" else ""), subset=["Situação"]),
                 hide_index=True, width="stretch")
    with st.expander("Notas metodológicas"):
        st.markdown(
            "- **Custos variáveis**: custo direto dos itens (ficha técnica ou valor de aquisição) + impostos sobre vendas "
            "+ custos de comercialização (%).\n"
            "- **Custos fixos totais**: pessoal + custos fixos operacionais + depreciação (linear, com valor residual).\n"
            "- **Fluxo do projeto** para VPL/TIR: Ano 0 = investimento total; Anos 1 a 3 = EBITDA "
            "(no Simples, os tributos já estão na alíquota sobre vendas). Opcionalmente recupera o capital de giro no Ano 3.\n"
            "- **Payback simples e rentabilidade** seguem o PNBOX: investimento ÷ resultado operacional mensal do Ano 1.\n"
            "- **Juros** calculados pela Tabela Price, mês a mês.\n"
            "- **Balanço simplificado**: o caixa é obtido por diferença (Ativo = Passivo + PL); "
            "os lucros ficam retidos (sem distribuição).")

# ==============================================================================
# ABA 12 - GRÁFICOS
# ==============================================================================
comp = composicoes(res, d)


def tabela_comp(df):
    return df.assign(Valor=df["Valor"].map(brl), **{"%": df["%"].map(pct)})


def rosca(df, titulo, chave_g):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(
        theta=alt.Theta("Valor:Q", stack=True),
        color=alt.Color("Item:N", sort=list(df["Item"]), legend=alt.Legend(title=None, orient="bottom", columns=2, labelLimit=260)),
        tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    arco = base.mark_arc(innerRadius=55, outerRadius=105)
    rot = base.transform_filter(alt.datum["%"] >= 0.04).mark_text(radius=125, fontSize=12).encode(
        text=alt.Text("%:Q", format=".1%"))
    st.altair_chart((arco + rot).properties(height=360), width="stretch", key=chave_g)
    st.dataframe(tabela_comp(df), hide_index=True, width="stretch")


def barras(df, titulo, chave_g, rotulo_pct=True):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(y=alt.Y("Item:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260)),
                                x=alt.X("Valor:Q", title="R$"),
                                tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    txt = base.mark_text(align="left", dx=3).encode(text=alt.Text("%:Q", format=".1%")) if rotulo_pct else None
    graf = base.mark_bar(color="#26548c")
    st.altair_chart((graf + txt if txt is not None else graf).properties(height=28 * len(df) + 40),
                    width="stretch", key=chave_g)


with abas[11]:
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
        rosca(comp["custo_direto"], "Custo direto mensal por categoria (matéria-prima, embalagem...)", "g_cd")
    st.caption("O custo direto por categoria usa as fichas técnicas. Produtos sem ficha aparecem como "
               "“Custo informado”. Classifique cada material na aba 9 (coluna Categoria).")

    st.subheader("Receitas")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["receita_categoria"], "Receita por categoria", "g_rc")
    with g2:
        barras(comp["top_produtos"], "10 produtos/serviços com maior receita mensal", "g_top")

    st.subheader("Para onde vai cada R$ 1,00 de receita (Ano 1)")
    barras(comp["estrutura_receita"], "Custos, despesas e lucro em % da receita bruta", "g_estr")

    st.subheader("Ponto de equilíbrio contábil (Ano 1)")
    eq = comp["equilibrio"].melt(id_vars="Receita", var_name="Série", value_name="R$")
    pe, rec = comp["pe_ponto"]
    linhas = alt.Chart(eq).mark_line(strokeWidth=3).encode(
        x=alt.X("Receita:Q", title="Receita anual (R$)"), y=alt.Y("R$:Q", title="R$"),
        color=alt.Color("Série:N", legend=alt.Legend(orient="bottom", title=None)))
    marcas = alt.Chart(pd.DataFrame({"x": [pe, rec], "rot": ["Ponto de equilíbrio", "Receita projetada"]})).mark_rule(
        strokeDash=[6, 4]).encode(x="x:Q", tooltip=["rot", alt.Tooltip("x:Q", format=",.2f")])
    st.altair_chart((linhas + marcas).properties(height=380), width="stretch", key="g_pe")
    st.caption(f"PE contábil: {brl(pe)} por ano · receita projetada: {brl(rec)} · "
               f"margem de segurança: {pct(div(rec - pe, rec))}")

    st.subheader("Recuperação do investimento (payback)")
    fa = comp["fluxo_acumulado"]
    st.bar_chart(fa.set_index("Ano")[["Acumulado", "Acumulado descontado (TMA)"]], stack=False)
    st.dataframe(fa.assign(**{c: fa[c].map(brl) for c in fa.columns[1:]}), hide_index=True, width="stretch")

    st.subheader("Cenários")
    def _num(v):
        try:
            return float(str(v).replace("R$", "").replace(".", "").replace(",", ".").strip())
        except ValueError:
            return 0.0
    cen_res = pd.DataFrame({"Cenário": list(df_cen.columns[1:]),
                            "Resultado operacional": [_num(v) for v in df_cen.iloc[4, 1:]]})
    st.altair_chart(alt.Chart(cen_res).mark_bar().encode(
        x=alt.X("Cenário:N", sort=None, title=None, axis=alt.Axis(labelAngle=0, labelLimit=300)),
        y=alt.Y("Resultado operacional:Q"),
        color=alt.condition(alt.datum["Resultado operacional"] > 0, alt.value("#26548c"), alt.value("#d6006e"))
    ).properties(height=320), width="stretch", key="g_cen")

# ==============================================================================
# ABA 13 - RELATÓRIO E EXPORTAÇÃO
# ==============================================================================
with abas[12]:
    st.subheader("Exportar o plano")
    st.write("O **relatório em PDF** segue a ordem do PNBOX: informações do plano, segmentos, proposta de valor, "
             "personas, jornada, concorrência, matriz SWOT, canais e funil, investimentos, produtos e fichas técnicas, "
             "custos, DRE, indicadores, cenários e checagens.")
    extras = {"comp": comp, "dre": df_dre, "indicadores": df_ind, "cenarios": df_cen, "checagens": df_chk,
              "swot_diag": pontuacao_swot(d["swot"]), "funil": tabela_funil(d["funil_vendas"]),
              "concorrencia_pivot": ss_piv if "ss_piv" in globals() else pd.DataFrame()}
    c1, c2, c3 = st.columns(3)
    if c1.button("📄 Gerar relatório em PDF", type="primary", width="stretch"):
        with st.spinner("Gerando o PDF..."):
            ss[f"pdf_{PID}"] = gerar_pdf(proj, d, res, extras)
    if ss.get(f"pdf_{PID}"):
        c1.download_button("⬇️ Baixar PDF", ss[f"pdf_{PID}"], f"Plano_de_Negocio_{PID}.pdf",
                           "application/pdf", width="stretch")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([proj]).T.reset_index().set_axis(["Campo", "Valor"], axis=1).to_excel(w, sheet_name="Projeto", index=False)
        for t, df_t in d.items():
            if isinstance(df_t, pd.DataFrame):
                df_t.to_excel(w, sheet_name=t[:31], index=False)
        res["cat"].to_excel(w, sheet_name="Custos_produtos", index=False)
        res["inv"].to_excel(w, sheet_name="Depreciacao", index=False)
        res["meses"].to_excel(w, sheet_name="Projecao_mensal", index=False)
        df_dre.to_excel(w, sheet_name="DRE", index=False)
        df_ind.to_excel(w, sheet_name="Indicadores", index=False)
        df_cen.to_excel(w, sheet_name="Cenarios", index=False)
        df_chk.to_excel(w, sheet_name="Checagens", index=False)
    c2.download_button("📊 Baixar planilha (Excel)", buf.getvalue(), f"Planilha_{PID}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    conn.commit()
    c3.download_button("🗄️ Baixar arquivo do projeto (.pnbox)", Path(ss.db).read_bytes(), f"{PID}.pnbox",
                       "application/octet-stream", width="stretch")
    st.caption("O arquivo .pnbox guarda todo o projeto e pode ser reaberto no menu lateral (📂 Abrir arquivo .pnbox).")

conn.close()
