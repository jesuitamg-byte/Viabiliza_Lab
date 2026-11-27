# ==============================================================================
# VIABILIZALAB - INTERFACE: BASE (configuração, projetos, navegação, salvamento)
# ==============================================================================
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from pnbox_db import (SCHEMA, conectar, ler, limpar, salvar, ler_unica, salvar_unica,
                      criar_projeto_em_branco, criar_projeto_modelo)
from pnbox_motor import (CATEGORIAS, QUADRANTES, brl, num_br, pct, div, custos_unitarios,
                         executar_motor, cenarios, tabela_dre, tabela_indicadores, pontuacao_swot,
                         composicoes, simular, sensibilidade, mercado, capacidade_demanda, semaforo,
                         tabela_checagens_v5, raio_x)

st.set_page_config(page_title="ViabilizaLab - Plano de Negócio", page_icon="📊", layout="wide",
                   initial_sidebar_state="auto")      # aberto no computador, recolhido no celular
ss = st.session_state
VERSAO_PROGRAMA = "9.2"

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
                    print("Nenhuma aba aberta. Encerrando o ViabilizaLab...")
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
    """'modelo' abre com o FitBox; 'branco', com um projeto vazio; 'assinatura', com o Conecta Vale."""
    try:
        q = st.query_params.get("versao", "")
    except Exception:
        q = ""
    if q in ("modelo", "branco", "assinatura", "servico"):
        return q
    if os.environ.get("PNBOX_VERSAO") in ("modelo", "branco", "assinatura", "servico"):
        return os.environ["PNBOX_VERSAO"]
    arq = APP_DIR / "versao.txt"
    if arq.exists():
        v = arq.read_text(encoding="utf-8").strip().lower()
        if v in ("modelo", "branco", "assinatura", "servico"):
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
    PASTA = Path.home() / "Documents" / "ViabilizaLab"
    _antiga = Path.home() / "Documents" / "PNBOX"
    if _antiga.is_dir() and not PASTA.exists():
        # migração da versão anterior (PNBOX Acadêmico): copia os projetos para a nova pasta
        try:
            shutil.copytree(_antiga, PASTA)
        except Exception:
            PASTA = _antiga
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
    "classe_custo_fixo": ["Custo indireto de produção", "Custo de tecnologia / infraestrutura",
                          "Despesa administrativa", "Despesa comercial", "Despesa financeira"],
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
    ("trilha", "Identificação", "🚦", "Trilha guiada e o que falta",
     "A ordem recomendada de preenchimento, o próximo passo e tudo o que ainda falta, com atalho para cada ajuste."),
    ("projeto", "Identificação", "🏢", "Informações e sócios",
     "Dados do negócio, integrantes do grupo, premissas financeiras e quadro societário."),
    ("identidade", "Identificação", "🧭", "Missão, visão, valores e objetivos",
     "Por que a empresa existe, aonde quer chegar, em que acredita e objetivos SMART que transformam a visão em metas."),
    ("segmentos", "Cliente e mercado", "🎯", "Segmentação de mercado",
     "Liste os grupos de clientes possíveis e dê notas para descobrir qual atender primeiro."),
    ("personas", "Cliente e mercado", "🧑", "Gerador de personas",
     "Crie personagens que representam o cliente ideal: dados, características, rotina e insights."),
    ("jornada", "Cliente e mercado", "🧭", "Jornada do cliente",
     "Descreva o que a persona faz, pensa e sente antes, durante e depois da compra."),
    ("entrevistas", "Cliente e mercado", "🎙️", "Entrevistas com clientes",
     "Roteiro de entrevista de descoberta e registro das conversas: quais dores se repetem e quanto pagariam."),
    ("mercado", "Cliente e mercado", "📐", "Tamanho do mercado",
     "Estime o mercado total (TAM), o alcançável (SAM) e o conquistável (SOM) e confira se a meta de vendas é realista."),
    ("dados_publicos", "Cliente e mercado", "🏙️", "Dados do município e das famílias (IBGE)",
     "População, domicílios, PIB e empresas do município, e quanto as famílias gastam em cada categoria (POF)."),
    ("proposta", "Problema e solução", "❤️", "Proposta de valor",
     "Defina, para cada persona, os benefícios que o negócio entrega e o que o diferencia."),
    ("vpc", "Problema e solução", "🎁", "Canvas da proposta de valor",
     "Tarefas, dores e ganhos do cliente x produtos, aliviadores e criadores de ganho, e o encaixe entre eles."),
    ("concorrencia", "Problema e solução", "🔍", "Análise da concorrência",
     "Compare o seu negócio com os concorrentes em notas de 0 a 10 por atributo."),
    ("oceano", "Problema e solução", "🌊", "Curva de valor e ERRC",
     "Compare a sua curva com a dos concorrentes e decida o que eliminar, reduzir, elevar e criar (oceano azul)."),
    ("mapa", "Problema e solução", "🗺️", "Mapa de concorrentes",
     "Concorrentes e lugares de movimento no mapa do município (OpenStreetMap), com raio de atendimento."),
    ("posicionamento", "Problema e solução", "📍", "Posicionamento e mapa perceptual",
     "A frase que diz para quem, o que e por que é diferente, e onde a marca fica frente aos concorrentes."),
    ("forcas", "Estratégia", "💪", "Forças e fraquezas",
     "Fatores internos: o que o negócio faz bem e o que precisa melhorar."),
    ("oportunidades", "Estratégia", "💡", "Oportunidades e ameaças",
     "Fatores externos: tendências que ajudam e riscos que vêm do ambiente."),
    ("swot", "Estratégia", "🧩", "Análise SWOT",
     "Matriz SWOT gerada automaticamente e a estratégia indicada para o negócio."),
    ("swot_cruzada", "Estratégia", "🔀", "SWOT cruzada",
     "Transforme o diagnóstico em estratégias ofensivas, de confronto, de reforço e defensivas."),
    ("canvas", "Análise estratégica", "🗂️", "Business Model Canvas",
     "O modelo de negócio em uma página, montado com o que você já preencheu."),
    ("lean", "Análise estratégica", "🪧", "Lean Canvas",
     "A versão do Canvas para negócios novos: problema, solução, métricas-chave e vantagem injusta."),
    ("porter", "Análise estratégica", "⚔️", "Cinco Forças de Porter",
     "Pressão competitiva do setor: rivalidade, novos entrantes, substitutos, clientes e fornecedores."),
    ("ge", "Análise estratégica", "🧮", "Matriz GE / McKinsey",
     "Atratividade de cada segmento x força competitiva da empresa: onde investir, ser seletivo ou colher."),
    ("bcg", "Análise estratégica", "🫧", "Matriz BCG",
     "Linhas de produto por crescimento do mercado e participação relativa: estrelas, vacas leiteiras, interrogações e abacaxis."),
    ("bsc", "Análise estratégica", "🗺️", "Balanced Scorecard",
     "Objetivos e indicadores nas quatro perspectivas e o mapa estratégico que liga tudo."),
    ("plano_mkt", "Marketing e operação", "🧭", "Plano de marketing",
     "Situação, objetivos, estratégia, táticas, ações e controle, amarrando as ferramentas de marketing."),
    ("marketing", "Marketing e operação", "📣", "Mix de marketing (4Ps e 7Ps)",
     "Produto, preço, praça e promoção (e, em serviços, pessoas, processos e evidência física) ligados aos números do plano."),
    ("canais", "Marketing e operação", "🛒", "Canais e funil de vendas",
     "Avalie os canais de aquisição e acompanhe a conversão do funil."),
    ("orcamento_mkt", "Marketing e operação", "💸", "Orçamento de marketing e funil AARRR",
     "Gasto e CAC por canal, retorno de cada canal, funil de crescimento, calendário de ações e meta de NPS."),
    ("operacional", "Marketing e operação", "🏭", "Plano operacional",
     "Localização, layout, processos, capacidade produtiva e organograma."),
    ("fornecedores", "Marketing e operação", "🚚", "Fornecedores",
     "Quem fornece cada item, prazos de pagamento e de entrega."),
    ("investimentos", "Finanças", "💰", "Investimentos e fontes",
     "Investimentos fixos, pré-operacionais, capital de giro e de onde vem o dinheiro."),
    ("financiamento", "Finanças", "🏦", "Financiamento (SAC, Price e carência)",
     "Parcelas, juros e custo efetivo de cada empréstimo, e comparação entre propostas de crédito."),
    ("taxas", "Finanças", "🏛️", "Selic, inflação e TMA",
     "Taxas atuais do Banco Central e a TMA montada em degraus: livre de risco, risco do negócio e pequeno porte."),
    ("produtos", "Finanças", "📦", "Produtos e materiais",
     "Materiais, produtos e serviços, ficha técnica e custo de cada item."),
    ("sazonalidade", "Finanças", "📅", "Sazonalidade das vendas",
     "Peso de cada mês nas vendas: meses fortes e fracos, e o efeito no fluxo de caixa."),
    ("servicos", "Finanças", "🧑‍💻", "Serviços e custo da hora",
     "Para quem vende tempo e conhecimento: custo da hora produtiva, horas da equipe e preço de cada serviço."),
    ("assinaturas", "Finanças", "💳", "Assinaturas e receita recorrente",
     "Planos mensais, trimestrais ou anuais: assinantes, cancelamentos (churn), MRR, CAC, LTV e custos de tecnologia."),
    ("custos", "Finanças", "🧾", "Custos e despesas",
     "Custos variáveis, mão de obra (produção e administrativa) e custos fixos classificados."),
    ("reforma", "Finanças", "🏛️", "Reforma tributária (simulação)",
     "Simples puro ou IBS/CBS por fora a partir de 2027: imposto da empresa e crédito para os clientes empresas."),
    ("preco", "Finanças", "🏷️", "Formação de preço",
     "Custeio variável e por absorção, preço sugerido pelo markup e ponto de equilíbrio por produto."),
    ("teste_preco", "Finanças", "🙋", "Teste de preço e demanda",
     "Faixa de preço aceitável pelo método Van Westendorp e demanda ajustada pela intenção de compra."),
    ("dre", "Finanças", "📄", "DRE e projeções",
     "Demonstração do resultado dos Anos 1 a 3 e a projeção mês a mês."),
    ("fluxo", "Finanças", "💵", "Fluxo de caixa mensal",
     "Entradas e saídas de dinheiro em cada mês do Ano 1."),
    ("indicadores", "Finanças", "📊", "Indicadores financeiros",
     "Semáforo de saúde, VPL, TIR, payback, pontos de equilíbrio e checagens."),
    ("referencias", "Finanças", "📏", "Referências do setor (IBGE e Sebrae)",
     "Compare margem, peso da folha e produtividade com a média do setor e veja a taxa de sobrevivência das empresas."),
    ("graficos", "Finanças", "📈", "Gráficos",
     "Composição do investimento, dos custos e da receita, ponto de equilíbrio e payback."),
    ("simulador", "Ferramentas complementares", "🎚️", "Simulador de resultados",
     "Cenários otimista e pessimista, simulador “e se?” e análise de sensibilidade."),
    ("montecarlo", "Ferramentas complementares", "🎲", "Monte Carlo e perda aceitável",
     "Milhares de cenários sorteados: a chance de o projeto dar certo e quanto os sócios podem perder."),
    ("tornado", "Ferramentas complementares", "🌪️", "Sensibilidade do VPL (tornado)",
     "Quais premissas mais mexem no VPL, quanto cada uma pode mudar até o VPL zerar e o mapa preço × volume."),
    ("arvore", "Ferramentas complementares", "🌳", "Árvore de decisão",
     "Abrir ou não abrir em três cenários, opção de abandono e quanto vale pesquisar antes de decidir."),
    ("adocao", "Ferramentas complementares", "🌱", "Curva de adoção (Bass)",
     "Teste se o crescimento de clientes do plano é plausível pela difusão de produtos novos."),
    ("valuation", "Ferramentas complementares", "💎", "Valuation e rodada de investimento",
     "Quanto vale o negócio por três métodos e quanto um investidor levaria em troca do aporte."),
    ("impacto", "Ferramentas complementares", "🌍", "Impacto social e ODS",
     "Teoria da mudança, indicadores de impacto e Objetivos de Desenvolvimento Sustentável."),
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
    ("evidencias", "Validação e defesa", "🔬", "Placar de evidências",
     "Quais números do plano são palpite, pesquisa ou teste com clientes, e qual premissa validar primeiro."),
    ("prontidao", "Validação e defesa", "🪜", "Régua de prontidão",
     "Em que degrau, de 1 a 9, o negócio está: da ideia descrita ao pronto para lançar ou buscar investimento."),
    ("banca", "Validação e defesa", "🧑‍⚖️", "Banca simulada",
     "Perguntas difíceis geradas a partir dos pontos fracos do plano, para treinar a defesa."),
    ("formalizacao", "Validação e defesa", "📋", "Checklist de formalização",
     "Enquadramento (MEI, ME, EPP), registros, licenças, LGPD e marca."),
    ("editais", "Validação e defesa", "📝", "Editais de fomento (modelo Centelha)",
     "Rascunho da proposta para editais de inovação, com limite de caracteres e orçamento por rubrica."),
    ("sumario", "Conclusão e entrega", "📝", "Sumário executivo",
     "Resumo do plano gerado automaticamente, para revisar e completar."),
    ("apresentacao", "Conclusão e entrega", "🎤", "Modo apresentação",
     "Telas grandes com os principais números para apresentar o plano."),
    ("relatorio", "Conclusão e entrega", "📤", "Relatório e exportação",
     "PDF completo, pitch deck, Word, Excel e arquivo do projeto."),
    ("professor", "Conclusão e entrega", "🎓", "Modo professor",
     "Compare os projetos da turma, aplique a rubrica e gere fichas de avaliação."),
    ("glossario", "Conclusão e entrega", "📖", "Glossário e leitura dos números",
     "O que significa cada termo e o que os números deste projeto querem dizer, em linguagem simples."),
    ("planilha", "Conclusão e entrega", "📑", "Importar e exportar planilha",
     "Traga produtos, custos, entrevistas e respostas de pesquisa de uma planilha e exporte tudo para o Excel."),
    ("historico", "Conclusão e entrega", "🕑", "Histórico de versões do projeto",
     "Salve versões, compare os indicadores entre elas e restaure uma versão anterior."),
    ("novidades", "Conclusão e entrega", "🆕", "Novidades do programa",
     "O que mudou em cada versão do ViabilizaLab."),
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
      [class*="st-key-card_"] {{
        position: relative; min-height: 186px; height: 100%; border-radius: 16px; padding: 16px 18px 14px 18px;
        background: {fundo_card}; box-shadow: 0 4px 14px rgba(20,40,80,.18); gap: .25rem;
        transition: transform .12s ease, box-shadow .12s ease; cursor: pointer;
      }}
      [class*="st-key-card_"]:hover {{ transform: translateY(-3px); box-shadow: 0 10px 24px rgba(20,40,80,.30); }}
      [class*="st-key-card_"] p, [class*="st-key-card_"] div {{ color: #fff !important; }}
      [class*="st-key-card_"] .pn-ic {{ font-size: 2.1rem; line-height: 1.1; }}
      [class*="st-key-card_"] .pn-tit {{ font-size: 1.08rem; font-weight: 700; margin-top: .35rem; line-height: 1.25; }}
      [class*="st-key-card_"] .pn-desc {{ font-size: .84rem; opacity: .88; margin-top: .3rem; line-height: 1.3; }}
      [class*="st-key-card_"] .pn-selo {{ font-size: .8rem; margin-top: .45rem; opacity: .95; }}
      [class*="st-key-cardbtn_"] {{ position: absolute !important; inset: 0; margin: 0 !important; z-index: 3; }}
      [class*="st-key-cardbtn_"] > div, [class*="st-key-cardbtn_"] .stButton {{ height: 100%; width: 100%; }}
      [class*="st-key-cardbtn_"] button {{ width: 100%; height: 100% !important; opacity: 0; cursor: pointer; }}
      [class*="st-key-cardbtn_"] button:focus-visible {{ opacity: .25; }}
      [class*="st-key-nav_"] button {{ justify-content: flex-start; text-align: left; padding: .25rem .6rem;
        min-height: 2rem; }}
      [class*="st-key-voltar"] button {{ font-weight: 600; }}
      .pn-grupo {{ font-size: 1.15rem; font-weight: 700; margin: 1.2rem 0 .4rem 0; }}
      .pn-card-info {{ border: 1px solid #dfe5ee; border-radius: 12px; padding: 14px 16px; background: #f7f9fc; }}
      .pn-face {{ font-size: 1.8rem; line-height: 1; }}
      .pn-etapa {{ border-left: 4px solid #2f6fd0; padding: 6px 12px; margin: 8px 0; }}
      [data-testid="stAppDeployButton"] {{ display: none; }}
      /* ---------- Celular e telas estreitas ---------- */
      @media (max-width: 700px) {{
        [data-testid="stMainBlockContainer"], .block-container {{ padding: 2.6rem .7rem 5rem .7rem !important; }}
        h1 {{ font-size: 1.55rem !important; }}
        h2 {{ font-size: 1.3rem !important; }}
        h3 {{ font-size: 1.12rem !important; }}
        /* grade de cartões: 2 por linha, compactos */
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) {{ flex-wrap: wrap !important; gap: .55rem !important; }}
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) > [data-testid="stColumn"] {{
          flex: 1 1 calc(50% - .55rem) !important; min-width: calc(50% - .55rem) !important; width: calc(50% - .55rem) !important; }}
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) > [data-testid="stColumn"]:not(:has([class*="st-key-card_"])) {{
          display: none !important; }}
        [class*="st-key-card_"] {{ min-height: 118px; padding: 11px 12px 10px 12px; border-radius: 14px; }}
        [class*="st-key-card_"] .pn-ic {{ font-size: 1.6rem; }}
        [class*="st-key-card_"] .pn-tit {{ font-size: .92rem; margin-top: .2rem; }}
        [class*="st-key-card_"] .pn-desc {{ display: none; }}
        [class*="st-key-card_"] .pn-selo {{ font-size: .72rem; margin-top: .3rem; }}
        .pn-grupo {{ font-size: 1.02rem; margin: .9rem 0 .3rem 0; }}
        /* números e botões maiores para o dedo */
        [data-testid="stMetricValue"] {{ font-size: 1.25rem !important; }}
        [data-testid="stMetricLabel"] p {{ font-size: .8rem !important; }}
        .stButton button, .stDownloadButton button {{ min-height: 44px; }}
        [data-baseweb="tab-list"] {{ overflow-x: auto; scrollbar-width: none; }}
        [data-baseweb="tab"] {{ padding-left: .55rem !important; padding-right: .55rem !important; white-space: nowrap; }}
        .pn-card-info {{ padding: 11px 12px; font-size: .92rem; }}
      }}
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
    elif VERSAO == "assinatura":
        criar_projeto_assinatura(PASTA / "Conecta Vale (modelo assinatura).pnbox")
    elif VERSAO == "servico":
        criar_projeto_servico(PASTA / "TecVale Soluções (modelo serviços).pnbox")
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
    if not filtro and tabela in LEITURAS:
        try:
            frases = LEITURAS[tabela](limpar(d[tabela], tabela), proj, d, R)
        except Exception:
            frases = []
        mostrar_leitura(tabela, frases)
    return ed


def mostrar_leitura(chave_: str, frases: list):
    """Caixa "Leitura dos resultados" ao final de uma tabela. No modo exercício, o aluno escreve antes de ver."""
    if not frases or not ss.get("tg_leit", True):
        return
    with st.container(border=True):
        st.markdown("**📌 Leitura dos resultados**")
        if ss.get("tg_exerc", False):
            la = d["leituras_aluno"]
            atual = la.loc[la["tabela"] == chave_, "texto"]
            atual = str(atual.iloc[0]) if not atual.empty else ""
            txt = st.text_area("Antes de ver a leitura do programa, escreva a sua: o que esta tabela mostra?", atual,
                               key=wk("la_" + chave_), height=90)
            if txt != atual:
                novo = pd.concat([la[la["tabela"] != chave_], pd.DataFrame([{"tabela": chave_, "texto": txt}])], ignore_index=True)
                salvar(conn, "leituras_aluno", novo, "tabela")
                d["leituras_aluno"] = limpar(novo, "leituras_aluno")
                marcar_salvo()
            if len(txt.strip()) < 20:
                st.caption("A leitura do programa aparece depois que você escrever ao menos uma frase.")
                return
            st.caption("Compare com a leitura do programa:")
        st.markdown("\n".join(f"- {f}" for f in frases))
        if ss.get("tg_voz", False):
            import streamlit.components.v1 as _cmp
            _cmp.html(html_leitor(" ".join(frases)), height=48)


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


def _receita() -> bool:
    return _tem("catalogo") or _tem("assinaturas")


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
        "identidade": (bool(proj["missao"]) and bool(proj["visao"]) and bool(proj["valores"]) and _tem("objetivos", 3) and
                       all(all(ok for ok, _ in checar_smart(o).values()) for _, o in d["objetivos"].iterrows()),
                       bool(proj["missao"]) or bool(proj["visao"]) or _tem("objetivos")),
        "canvas": (bool(proj["canvas_relacionamento"]) and bool(proj["canvas_parcerias"]), _tem("segmentos") or _tem("propostas")),
        "porter": (len(d["porter"]) >= 5 and bool((d["porter"]["intensidade"] > 0).all()) and
                   not d["porter"]["justificativa"].astype(str).str.startswith("(sugestão)").any(), _tem("porter")),
        "ge": (_tem("segmentos", 2) and _tem("concorrencia_notas"), _tem("segmentos")),
        "bcg": (_tem("bcg", 2), _tem("bcg")),
        "bsc": (set(PERSPECTIVAS) <= set(d["objetivos"]["perspectiva"]), _tem("objetivos")),
        "servicos": (bool((d["catalogo"]["horas_unit"] > 0).any()) and float(proj["utilizacao_pct"]) > 0,
                     bool((d["catalogo"]["horas_unit"] > 0).any())),
        "assinaturas": (_tem("assinaturas") and bool((d["assinaturas"]["churn_mensal_pct"] > 0).all()), _tem("assinaturas")),
        "preco": (_tem("catalogo"), _tem("catalogo")),
        "dre": (_receita() and _tem("custos_fixos"), _receita()),
        "fluxo": (_receita() and _tem("capital_giro_prazos"), _receita()),
        "indicadores": (_receita() and _tem("inv_fixos"), _receita()),
        "graficos": (_receita(), _receita()),
        "simulador": (_receita(), _receita()),
        "experimentos": (_tem("experimentos"), _tem("experimentos")),
        "riscos": (_tem("riscos", 3), _tem("riscos")),
        "cronograma": (_tem("cronograma", 3), _tem("cronograma")),
        "realizado": (_tem("realizado", 1), _tem("realizado")),
        "sumario": (len(proj["sumario_executivo"]) > 200, bool(proj["sumario_executivo"])),
    }
    regras.update(regras_v71(proj, d))
    regras.update(regras_v8(proj, d))
    regras.update(regras_mix(proj, d))
    if p not in regras:
        return ""
    completo, iniciado = regras[p]
    tipo = str(proj.get("tipo_negocio") or "Produtos")
    opcional = (p == "servicos" and not tipo.startswith(("Serviço", "Misto"))) or \
               (p == "assinaturas" and not tipo.startswith(("Assinatura", "Misto"))) or p in ("reforma", "adocao") or p in OPCIONAIS_V8 or p in ("plano_mkt", "sazonalidade", "tornado", "taxas", "mapa", "arvore")
    if opcional and not iniciado:
        return ""            # ferramenta não se aplica ao tipo de negócio escolhido
    return "✅" if completo else ("◐" if iniciado else "○")


def completude() -> float:
    return completude_de(proj, d)


# ------------------------------------------------------------------------------
# MENU LATERAL
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# GOOGLE DRIVE (v9.1): login, pasta do aluno no Drive e salvamento automático
# ------------------------------------------------------------------------------
_MAPA_DRIVE = PASTA / ".drive.json"


def _drive_mapa() -> dict:
    try:
        return json.loads(_MAPA_DRIVE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _drive_registro(caminho=None) -> dict | None:
    """Arquivo do Drive ligado ao projeto local (para a conta conectada)."""
    if not ss.get("g_email"):
        return None
    return _drive_mapa().get(ss.g_email, {}).get(Path(caminho or ss.db).name)


def _drive_ligar(caminho, id_arq: str, mod: str, md5: str = ""):
    m = _drive_mapa()
    m.setdefault(ss.g_email, {})[Path(caminho).name] = {"id": id_arq, "mod": mod, "md5": md5}
    try:
        _MAPA_DRIVE.write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    ss.g_enviado = Path(caminho).stat().st_mtime if Path(caminho).exists() else 0
    ss.g_hora = datetime.now().strftime("%H:%M")


def drive_conectado() -> bool:
    return bool(ss.get("g_token")) and bool(ss.get("g_email")) and bool(ss.get("g_pasta")) and time.time() < ss.get("g_exp", 0)


def _drive_processar_canal():
    """Mensagens que o código da página manda pelo campo oculto (login, saída, arquivo escolhido)."""
    bruto = ss.get("gcanal") or ""
    if not bruto or bruto == ss.get("_gcanal_lido"):
        return
    ss._gcanal_lido = bruto
    try:
        msg = json.loads(bruto)
    except Exception:
        return
    tipo = msg.get("tipo")
    if tipo == "token":
        ss.g_token, ss.g_exp = msg["token"], time.time() + float(msg.get("expira", 3600)) - 120
        try:
            ss.g_email = drive_usuario(ss.g_token)
            ss.g_pasta = drive_pasta(ss.g_token)
            ss.g_erro = None
        except Exception as e:
            ss.g_erro = drive_mensagem_erro(e)
            for k in ("g_token", "g_exp", "g_email", "g_pasta"):
                ss.pop(k, None)
    elif tipo == "sair":
        for k in ("g_token", "g_exp", "g_email", "g_pasta", "g_erro", "g_conflito", "g_lista", "g_comp", "g_travado",
                  "g_leitura", "g_dono", "g_trava_ate", "g_perms"):
            ss.pop(k, None)
    elif tipo == "picker" and drive_conectado():
        _drive_abrir(msg["id"], msg.get("nome") or "Projeto do Drive.pnbox")
    elif tipo == "erro":
        ss.g_erro = {"popup_closed": "A janela do Google foi fechada antes de terminar o login.",
                     "popup_failed_to_open": "O navegador bloqueou a janela do Google. Libere as janelas para este site e tente de novo.",
                     "access_denied": "O acesso ao Drive não foi autorizado.",
                     "sem_login": "Entre com o Google antes de escolher um arquivo.",
                     "picker": "Não foi possível abrir a janela de escolha de arquivos do Google."}.get(
            msg.get("erro"), f"Não foi possível concluir o login ({msg.get('erro')}).")


def _drive_abrir(id_arq: str, nome: str):
    """Baixa um projeto do Drive e abre. Se o projeto já existir aqui, a cópia local é substituída."""
    try:
        dados = drive_baixar(ss.g_token, id_arq)
        meta = drive_metadados(ss.g_token, id_arq)
    except Exception as e:
        ss.g_erro = drive_mensagem_erro(e)
        return
    if not dados.startswith(b"SQLite format 3"):
        ss.g_erro = "O arquivo escolhido não é um projeto do ViabilizaLab (.pnbox)."
        return
    destino = None
    for nome_local, reg in _drive_mapa().get(ss.g_email, {}).items():
        if reg.get("id") == id_arq and (PASTA / nome_local).exists():
            destino = PASTA / nome_local
    destino = destino or caminho_livre(Path(nome).stem)
    try:
        conn.close()
    except Exception:
        pass
    destino.write_bytes(dados)
    conectar(destino).close()
    _drive_ligar(destino, id_arq, meta.get("modifiedTime", ""), meta.get("md5Checksum", ""))
    trocar_projeto(destino)
    ss.g_conflito = None
    ss.g_trava_ate = 0
    _drive_aplicar_meta(meta)


def _drive_aplicar_meta(meta: dict):
    """Guarda quem é o dono, se posso editar e se outra pessoa está editando agora."""
    ss.g_leitura = not (meta.get("capabilities") or {}).get("canEdit", True)
    ss.g_pode_compartilhar = (meta.get("capabilities") or {}).get("canShare", False)
    donos = meta.get("owners") or []
    ss.g_dono = donos[0].get("emailAddress", "") if donos else ""
    ss.g_travado = drive_estado_trava(meta, ss.g_email, time.time())


def _drive_sinal() -> dict | None:
    """A cada ciclo: lê o estado do arquivo no Drive e, se ninguém mais está editando, renova a minha trava."""
    reg = _drive_registro()
    if reg is None or not drive_conectado():
        return None
    try:
        meta = drive_metadados(ss.g_token, reg["id"])
        _drive_aplicar_meta(meta)
        if not ss.g_travado and not ss.g_leitura and ss.get("g_trava_ate", 0) - time.time() < TRAVA_SEGUNDOS / 2:
            ss.g_trava_ate = time.time() + TRAVA_SEGUNDOS
            drive_travar(ss.g_token, reg["id"], ss.g_email, ss.g_trava_ate)
        return meta
    except Exception as e:
        ss.g_erro = drive_mensagem_erro(e)
        return None


def _drive_liberar():
    """Retira a minha trava ao sair, para o colega não precisar esperar ela expirar."""
    reg = _drive_registro()
    if reg and drive_conectado() and ss.get("g_trava_ate", 0) > time.time():
        try:
            drive_travar(ss.g_token, reg["id"], ss.g_email, 0)
        except Exception:
            pass
    ss.g_trava_ate = 0


def drive_salvar(forcar: bool = False, sobrescrever: bool = False, meta: dict | None = None) -> bool:
    """Envia o projeto atual para o Drive se ele mudou desde o último envio. Devolve True se enviou."""
    if not drive_conectado() or ss.get("g_conflito") and not sobrescrever:
        return False
    if ss.get("g_travado") or ss.get("g_leitura"):
        return False
    caminho = Path(ss.db)
    reg = _drive_registro()
    if reg is None and not forcar:
        return False
    mtime = caminho.stat().st_mtime
    if not forcar and mtime <= ss.get("g_enviado", 0):
        return False
    if not forcar and time.time() - mtime < 4:          # espera o aluno parar de digitar
        return False
    try:
        conn.commit()
        dados = caminho.read_bytes()
        if reg is None:
            r = drive_criar(ss.g_token, caminho.name, dados, ss.g_pasta)
        else:
            if not sobrescrever:
                meta = meta or drive_metadados(ss.g_token, reg["id"])
                if meta.get("trashed"):
                    raise ErroDrive(404, "na lixeira")
                mudou = (meta.get("md5Checksum") != reg["md5"]) if reg.get("md5") and meta.get("md5Checksum") else \
                        bool(reg.get("mod") and meta.get("modifiedTime") != reg["mod"])
                if mudou:
                    ss.g_conflito = meta.get("modifiedTime") or "?"
                    return False
            r = drive_atualizar(ss.g_token, reg["id"], dados)
        _drive_ligar(caminho, r["id"], r.get("modifiedTime", ""), r.get("md5Checksum", ""))
        ss.g_enviado = mtime
        ss.g_erro = None
        return True
    except Exception as e:
        ss.g_erro = drive_mensagem_erro(e)
        return False


def _drive_status():
    reg = _drive_registro()
    if reg is not None and ss.get("g_leitura"):
        st.info(f"👁️ Você pode só visualizar este projeto{' de ' + ss.g_dono if ss.get('g_dono') else ''}. "
                "Alterações feitas aqui não são enviadas ao Drive.")
    elif reg is not None and ss.get("g_travado"):
        st.warning(f"🔒 **{ss.g_travado}** está editando este projeto agora. Suas alterações neste aparelho não serão "
                   "enviadas. Aguarde e atualize com a versão do Drive (a vez fica livre 2 minutos depois que a pessoa para).", icon="⏳")
        if st.button("🔄 Atualizar com a versão do Drive", key="g_trava_baixar", width="stretch"):
            _drive_abrir(reg["id"], Path(ss.db).name)
            st.rerun()
    if ss.get("g_conflito"):
        st.warning("Este projeto foi alterado em outro aparelho depois da última vez que você salvou daqui.", icon="⚠️")
        a, b = st.columns(2)
        if a.button("Usar a versão do Drive", key="g_conf_baixar", width="stretch"):
            _drive_abrir(reg["id"], Path(ss.db).name)
            st.rerun()
        if b.button("Manter a minha", key="g_conf_subir", width="stretch"):
            ss.g_conflito = None
            drive_salvar(forcar=True, sobrescrever=True)
            st.rerun()
    elif reg is None:
        st.caption("Este projeto ainda não está no seu Drive.")
    elif Path(ss.db).stat().st_mtime > ss.get("g_enviado", 0) + 0.5:
        st.caption("⏳ Alterações ainda não enviadas ao Drive...")
    else:
        st.caption(f"☁️ Salvo no Drive{' às ' + ss.g_hora if ss.get('g_hora') else ''}.")
    if ss.get("g_erro"):
        st.caption(f"⚠️ {ss.g_erro}")


if hasattr(st, "fragment"):
    @st.fragment(run_every=15)
    def _drive_vigia():
        """Roda sozinho a cada 15 segundos: envia ao Drive o que mudou e atualiza o aviso."""
        if drive_conectado():
            drive_salvar(meta=_drive_sinal())
            _drive_status()
else:
    def _drive_vigia():
        if drive_conectado():
            drive_salvar(meta=_drive_sinal())
            _drive_status()


def painel_drive():
    import streamlit.components.v1 as components
    _drive_processar_canal()
    st.markdown("<style>.st-key-gcanal{position:absolute!important;left:-10000px!important;top:0!important;width:1px!important;"
                "height:1px!important;overflow:hidden!important}</style>", unsafe_allow_html=True)
    st.text_input("canal", key="gcanal", label_visibility="collapsed")
    if ss.pop("_g_sair_js", False):        # revoga a autorização no Google e esquece o token na página
        components.html("<script>try{window.parent.__vlGoogle.sair();}catch(e){}</script>", height=0)
    with st.expander("☁️ Google Drive", expanded=bool(ss.get("g_erro") or ss.get("g_conflito"))):
        if not drive_conectado():
            if ss.get("g_email"):
                st.caption(f"A autorização de {ss.g_email} expirou. Entre de novo para continuar salvando no Drive.")
            else:
                st.caption("Entre com a sua conta Google para guardar os projetos numa pasta \"ViabilizaLab\" do seu Drive "
                           "e abri-los de qualquer computador ou celular.")
            components.html(drive_botao("entrar", "Entrar com Google"), height=46)
            if ss.get("g_erro"):
                st.caption(f"⚠️ {ss.g_erro}")
            st.caption("[Como seus dados são tratados](https://jesuitamg-byte.github.io/Viabiliza_Lab/privacidade.html)")
            return
        st.caption(f"Conectado como **{ss.g_email}**")
        _drive_vigia()
        if _drive_registro() is None:
            if st.button("☁️ Guardar este projeto no Drive", key="g_guardar", type="primary", width="stretch"):
                drive_salvar(forcar=True)
                st.rerun()
            st.caption("Depois disso, cada alteração é enviada ao Drive automaticamente.")
        elif st.button("Salvar no Drive agora", key="g_salvar", width="stretch"):
            drive_salvar(forcar=True)
            st.rerun()
        st.markdown("**Abrir do Drive**")
        if st.button("🔄 Ver meus projetos no Drive", key="g_listar", width="stretch"):
            try:
                ss.g_lista = drive_listar(ss.g_token, ss.g_pasta)
            except Exception as e:
                ss.g_erro = drive_mensagem_erro(e)
        lista = ss.get("g_lista") or []
        if lista:
            nomes = {f["id"]: f"{f['name']} ({f['modifiedTime'][8:10]}/{f['modifiedTime'][5:7]} {f['modifiedTime'][11:16]} UTC)" for f in lista}
            esc = st.selectbox("Projeto", list(nomes), format_func=nomes.get, key="g_esc", label_visibility="collapsed")
            if st.button("📂 Abrir este projeto", key="g_abrir", width="stretch"):
                _drive_abrir(esc, next(f["name"] for f in lista if f["id"] == esc))
                st.rerun()
        elif "g_lista" in ss:
            st.caption("Nenhum projeto na pasta ViabilizaLab do seu Drive ainda.")
        if st.button("👥 Ver projetos compartilhados comigo", key="g_comp_ver", width="stretch"):
            try:
                ss.g_comp = drive_compartilhados(ss.g_token)
            except Exception as e:
                ss.g_erro = drive_mensagem_erro(e)
        comp = ss.get("g_comp") or []
        if comp:
            nomes_c = {f["id"]: f"{f['name']} (de {(f.get('owners') or [{}])[0].get('emailAddress', '?')})" for f in comp}
            esc_c = st.selectbox("Compartilhado", list(nomes_c), format_func=nomes_c.get, key="g_esc_c", label_visibility="collapsed")
            if st.button("📂 Abrir o compartilhado", key="g_abrir_c", width="stretch"):
                _drive_abrir(esc_c, next(f["name"] for f in comp if f["id"] == esc_c))
                st.rerun()
        elif "g_comp" in ss:
            st.caption("Nenhum projeto compartilhado com você foi aberto aqui ainda.")
        st.caption("Recebeu um convite? Na primeira vez, abra o projeto pelo botão abaixo, na aba de compartilhados.")
        components.html(drive_botao("escolher", "Outro arquivo do Drive", "#ffffff", "#1a73e8"), height=46)
        if st.button("Sair da conta Google", key="g_sair", width="stretch"):
            _drive_liberar()
            for k in ("g_token", "g_exp", "g_email", "g_pasta", "g_erro", "g_conflito", "g_lista", "g_comp", "g_travado",
                      "g_leitura", "g_dono", "g_trava_ate", "g_perms"):
                ss.pop(k, None)
            ss._g_sair_js = True
            st.rerun()
    painel_colaboradores()


def painel_colaboradores():
    """Convites para colaborar no projeto (só para quem pode compartilhar o arquivo no Drive)."""
    reg = _drive_registro()
    if not drive_conectado() or reg is None:
        return
    with st.expander("👥 Colaboradores do projeto"):
        if ss.get("g_dono"):
            st.caption(f"O arquivo fica no Drive de **{ss.g_dono}**. As alterações de todos são salvas nele, uma pessoa de cada vez.")
        if not ss.get("g_pode_compartilhar"):
            st.caption("Só quem é dono ou pode editar e compartilhar o arquivo consegue convidar pessoas.")
            return
        if st.button("🔄 Ver quem tem acesso", key="g_perm_ver", width="stretch") or "g_perms" not in ss:
            try:
                ss.g_perms = drive_permissoes(ss.g_token, reg["id"])
            except Exception as e:
                ss.g_perms = []
                ss.g_erro = drive_mensagem_erro(e)
        nomes_papel = {"owner": "dono", "writer": "edita", "commenter": "comenta", "reader": "só vê"}
        for pm in ss.get("g_perms") or []:
            if pm.get("type") != "user":
                continue
            a, b = st.columns([3, 1])
            a.caption(f"{pm.get('emailAddress', '?')} · {nomes_papel.get(pm.get('role'), pm.get('role'))}")
            if pm.get("role") != "owner" and b.button("✕", key=f"g_rem_{pm['id']}", help="Retirar o acesso"):
                try:
                    drive_remover_permissao(ss.g_token, reg["id"], pm["id"])
                    ss.pop("g_perms", None)
                except Exception as e:
                    ss.g_erro = drive_mensagem_erro(e)
                st.rerun()
        st.markdown("**Convidar**")
        email_c = st.text_input("E-mail da pessoa (de preferência uma conta Google)", key="g_conv_email")
        papel = st.radio("Permissão", list(PAPEIS), horizontal=True, key="g_conv_papel")
        if st.button("✉️ Enviar convite", key="g_conv_env", type="primary", width="stretch"):
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email_c.strip()):
                st.error("Informe um e-mail válido.")
            else:
                msg = (f"{ss.g_email} convidou você para colaborar no plano de negócio \"{Path(ss.db).stem}\" no ViabilizaLab. "
                       "Para abrir: acesse https://jesuitamg-byte.github.io/Viabiliza_Lab/ , entre com a sua conta Google em "
                       "☁️ Google Drive e, na primeira vez, use \"Outro arquivo do Drive\" e escolha o projeto na aba de compartilhados.")
                try:
                    drive_convidar(ss.g_token, reg["id"], email_c, PAPEIS[papel], msg)
                    ss.pop("g_perms", None)
                    st.success(f"Convite enviado para {email_c.strip()}.")
                except Exception as e:
                    st.error(drive_mensagem_erro(e))
        st.caption("O Google envia o convite por e-mail. Você pode retirar o acesso a qualquer momento aqui ou no Drive.")



with st.sidebar:
    st.markdown("### 📊 ViabilizaLab")
    st.caption("Laboratório de plano de negócio e viabilidade (antigo PNBOX Acadêmico)")
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

    painel_drive()

    with st.expander("➕ Novo projeto"):
        novo_nome = st.text_input("Nome do novo projeto", placeholder="Ex.: Padaria Pão Nosso")
        tipo = st.radio("Começar a partir de", ["Em branco (do zero)", "Modelo FitBox (alimentação saudável)",
                                                "Modelo Café Mucuri (cafeteria e padaria)",
                                                "Modelo TecVale Soluções (serviços de TI)",
                                                "Modelo Conecta Vale (aplicativo por assinatura)", "Ponto de partida por setor (sem valores)",
                                                "Cópia do projeto atual"])
        setor_esc = st.selectbox("Setor", list(SETORES), key="novo_setor") if tipo.startswith("Ponto de partida") else ""
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
                elif tipo.startswith("Modelo Conecta"):
                    criar_projeto_assinatura(destino)
                elif tipo.startswith("Modelo TecVale"):
                    criar_projeto_servico(destino)
                elif tipo.startswith("Ponto de partida"):
                    criar_projeto_setor(destino, novo_nome.strip(), setor_esc)
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

    with st.expander("♿ Acessibilidade"):
        st.toggle("Letra maior", value=False, key="tg_fonte")
        st.toggle("Alto contraste", value=False, key="tg_contraste")
        st.toggle("Botão para ouvir as leituras", value=False, key="tg_voz",
                  help="Usa a voz do próprio navegador para ler em voz alta a leitura dos resultados.")
    _css_ac = css_acessibilidade(ss.get("tg_fonte", False), ss.get("tg_contraste", False))
    if _css_ac:
        st.markdown(_css_ac, unsafe_allow_html=True)
    st.toggle("📌 Leitura automática das tabelas", value=True, key="tg_leit",
              help="Ao final de cada tabela, frases que dizem o que os números mostram.")
    st.toggle("✍️ Modo exercício", value=False, key="tg_exerc",
              help="O aluno escreve a própria leitura de cada tabela antes de ver a do programa. As leituras ficam salvas no projeto.")

    with st.expander("🧭 Ir direto para...", expanded=False):
        for g in GRUPOS:
            st.caption(g)
            for p in [x for x in PAGINAS if x[1] == g]:
                if st.button(f"{p[2]} {p[3]}", key=f"nav_{p[0]}", width="stretch",
                             type="primary" if ss.pagina == p[0] else "tertiary"):
                    ir(p[0])
                    st.rerun()

    st.divider()
    if NAVEGADOR and drive_conectado() and _drive_registro() is not None:
        st.success("**Projeto no seu Google Drive:** cada alteração é enviada para a pasta **ViabilizaLab** do seu Drive. "
                   "Para continuar em outro computador ou no celular, entre com a mesma conta e use **Abrir do Drive**. "
                   "Para entregar ao professor, compartilhe o arquivo no Drive ou baixe o **.pnbox** em 📤 Relatório e exportação.")
    elif NAVEGADOR and drive_conectado():
        st.info("**Versão on-line:** este projeto ainda está só **neste navegador**. Use **☁️ Guardar este projeto no Drive** "
                "(acima) para salvá-lo automaticamente no seu Google Drive e abri-lo de qualquer lugar.")
    elif NAVEGADOR:
        st.info("**Versão on-line:** seus projetos ficam guardados **neste navegador, neste computador**. "
                "Para abrir de qualquer lugar, entre com a sua conta Google em **☁️ Google Drive** (acima). "
                "Para entregar ao professor, baixe o **.pnbox** em 📤 Relatório e exportação.")
    elif MODO_WEB:
        st.warning("**Versão on-line:** o projeto fica só nesta aba. Ao terminar, baixe o **.pnbox** em "
                   "📤 Relatório e exportação.")
    else:
        st.caption(f"Projetos salvos em:\n`{PASTA}`")
    if not MODO_WEB and st.button("⏻ Encerrar o programa", width="stretch"):
        tela_de_encerramento()
        threading.Timer(2.5, lambda: os._exit(0)).start()
        st.stop()
    st.caption(f"Versão {VERSAO_PROGRAMA} · [Política de privacidade](https://jesuitamg-byte.github.io/Viabiliza_Lab/privacidade.html)")


def tela_de_encerramento():
    """
    Troca a página por um aviso em português antes de o programa desligar. Sem isso, a aba aberta
    tentaria se reconectar e mostraria a mensagem técnica "Connection error" do Streamlit.
    O código é criado no contexto da página principal, porque este quadro some quando o servidor cai.
    """
    import streamlit.components.v1 as components
    components.html("""<script>
    try {
      const w = window.parent;
      new w.Function(`
        const html = '<div style="font-family:system-ui,Segoe UI,sans-serif;max-width:560px;margin:18vh auto;text-align:center;color:#26548c">' +
          '<div style="font-size:3rem">📊</div><h2>O ViabilizaLab foi encerrado</h2>' +
          '<p style="color:#333;font-size:1.05rem">Seus projetos estão salvos. Pode fechar esta aba.</p>' +
          '<p style="color:#666;font-size:.9rem">Para usar de novo, abra o atalho do ViabilizaLab na área de trabalho.</p></div>';
        const trocar = () => { document.title = "ViabilizaLab encerrado"; document.body.innerHTML = html;
                               document.body.style.background = "#fff"; };
        setTimeout(trocar, 600);
        setTimeout(() => { trocar(); try { window.close(); } catch (e) {} }, 1800);
        new MutationObserver(() => { if (!document.body.innerText.includes("foi encerrado")) trocar(); })
          .observe(document.documentElement, {childList: true, subtree: true});
      `)();
    } catch (e) {}
    </script>""", height=0)


# ------------------------------------------------------------------------------
# CABEÇALHO DAS FERRAMENTAS
# ------------------------------------------------------------------------------
def rodape_navegacao(p: str):
    """
    Fim de cada página: botão de voltar ao painel e, quando a página acabou de ser aberta,
    rola a tela para o início e registra a página no histórico do navegador, para que o
    botão Voltar do navegador leve de volta ao painel do plano.
    """
    if p != "inicio" and p in PAG:
        st.write("")
        if st.button("← Voltar ao painel do plano", key="voltar_fim_" + p, width="stretch"):
            ir("inicio")
            st.rerun()
    if ss.get("_pag_vista") == p:
        return
    ss["_pag_vista"] = p
    ss["_nav_n"] = ss.get("_nav_n", 0) + 1
    try:
        import streamlit.components.v1 as components
    except Exception:          # recurso opcional
        return
    import logging
    logging.getLogger("streamlit").setLevel(logging.ERROR)
    components.html("""<script>
    /* pagina: %s, abertura %d */
    try {
      const w = window.parent, d = w.document, pag = "%s";
      const topo = () => { for (const s of ['[data-testid="stMain"]', '[data-testid="stAppViewContainer"]', 'section.main', '.main']) {
          const el = d.querySelector(s); if (el) el.scrollTo({top: 0}); } w.scrollTo({top: 0}); };
      topo(); [80, 250, 600, 1200].forEach(t => setTimeout(topo, t));
      /* o ouvinte é criado no contexto da página principal: este quadro é removido na próxima atualização
         e uma função criada aqui deixaria de funcionar */
      if (!w.__vlHistorico2) {
        w.__vlHistorico2 = true;
        new w.Function('window.addEventListener("popstate", function () {' +
          ' if (!window.location.hash) { var b = document.querySelector(".st-key-nav_inicio button"); if (b) b.click(); } });')();
      }
      const base = w.location.pathname + w.location.search, h = w.location.hash.slice(1);
      if (pag === "inicio") { if (h) w.history.replaceState(null, "", base); }
      else if (!h) { w.history.pushState({vl: pag}, "", base + "#" + pag); }
      else if (h !== pag) { w.history.replaceState({vl: pag}, "", base + "#" + pag); }
    } catch (e) {}
    </script>""" % (p, ss["_nav_n"], p), height=0)


def cabecalho(p: str):
    _, grupo, icone, titulo, descricao = PAG[p]
    st.caption(f"{grupo}  ›  {titulo}")
    st.title(f"{icone} {titulo}")
    if ss.get("g_travado") and drive_conectado():
        st.warning(f"🔒 {ss.g_travado} está editando este projeto agora. O que você mudar aqui não será enviado ao Drive.", icon="⏳")
    st.markdown(f"<div class='pn-card-info'>ℹ️ {descricao}</div>", unsafe_allow_html=True)
    if ss.alerta and ss.alerta[0] == p:
        st.warning(f"**O que ajustar aqui:** {ss.alerta[1]}")
        if st.button("Entendi, ocultar aviso", key="ok_alerta"):
            ss.alerta = None
            st.rerun()
    comentarios_da_pagina(p)
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
