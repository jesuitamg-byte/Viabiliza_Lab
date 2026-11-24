# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE (Streamlit)
# ==============================================================================
import io
import os
import re
import threading
import time
import shutil
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from pnbox_db import (SCHEMA, conectar, ler, limpar, salvar, ler_unica, salvar_unica,
                      criar_projeto_em_branco, criar_projeto_modelo)
from pnbox_motor import (CATEGORIAS, QUADRANTES, brl, num_br, pct, div, custos_unitarios,
                         executar_motor, cenarios, tabela_dre, tabela_indicadores,
                         tabela_checagens, pontuacao_swot, composicoes)
from pnbox_pdf import gerar_pdf

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
import sys
import tempfile

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
