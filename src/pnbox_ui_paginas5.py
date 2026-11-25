# ==============================================================================
# VIABILIZALAB - INTERFACE: IDENTIDADE E OBJETIVOS SMART, ANÁLISE ESTRATÉGICA E REFORMA TRIBUTÁRIA
# ==============================================================================
import matplotlib.pyplot as plt

VALORES_SUGERIDOS = ["Ética", "Transparência", "Qualidade", "Inovação", "Respeito ao cliente", "Responsabilidade social",
                     "Sustentabilidade", "Valorização das pessoas", "Compromisso com a região", "Simplicidade",
                     "Excelência no atendimento", "Colaboração", "Agilidade", "Confiança"]


def recarregar_tabela(tabela: str, df: pd.DataFrame, chave_: str):
    """Grava uma tabela alterada por um botão e força o editor a reabrir com os dados novos."""
    salvar(conn, tabela, df, chave_)
    d[tabela] = limpar(df[df[chave_].astype(str).str.strip() != ""].reset_index(drop=True), tabela)
    ss.buf.pop((PID, tabela, ""), None)
    ss.hash.pop((PID, tabela, ""), None)
    _nova_versao(tabela)
    invalidar()


def _selos_smart(chk: dict) -> str:
    return " ".join(f"{'✅' if ok else '❌'} {letra}" for letra, (ok, _) in chk.items())


# ------------------------------------------------------------------------------
# MISSÃO, VISÃO, VALORES E OBJETIVOS SMART
# ------------------------------------------------------------------------------
def pg_identidade():
    cabecalho("identidade")
    t1, t2, t3 = st.tabs(["🧭 Missão, visão e valores", "🎯 Objetivos SMART", "📚 Como escrever"])
    with t1:
        v = {}
        v["missao"] = st.text_area("Missão: por que a empresa existe?", proj["missao"], height=110, key=wk("missao"),
                                   placeholder="Oferecer [o quê] para [quem], por meio de [como], gerando [qual benefício].",
                                   help="Curta, no presente, centrada no cliente e no benefício que a empresa entrega.")
        v["visao"] = st.text_area("Visão: aonde a empresa quer chegar?", proj["visao"], height=90, key=wk("visao"),
                                  placeholder="Ser reconhecida como [posição desejada] em [mercado/região] até [ano].",
                                  help="Futuro desejado, inspirador e com horizonte de tempo.")
        atuais = [x.strip() for x in str(proj["valores"]).split(";") if x.strip()]
        escolhidos = st.multiselect("Valores: em que a empresa acredita e como se comporta?",
                                    list(dict.fromkeys(VALORES_SUGERIDOS + atuais)), atuais, key=wk("valores"),
                                    accept_new_options=True, help="Escolha de 3 a 6 e acrescente os seus.")
        v["valores"] = "; ".join(escolhidos)
        auto_unica("projeto", v)
        salvo_txt()
        dicas = []
        if v["missao"] and len(v["missao"].split()) > 45:
            dicas.append("A missão está longa: tente caber em até 2 linhas.")
        if v["visao"] and not re.search(r"\b(19|20)\d{2}\b|\banos?\b", v["visao"]):
            dicas.append("A visão não tem prazo: inclua um ano ou horizonte (ex.: \"até 2030\").")
        if escolhidos and not 3 <= len(escolhidos) <= 6:
            dicas.append("Recomenda-se de 3 a 6 valores, para que todos sejam lembrados e praticados.")
        for dc in dicas:
            st.warning(dc)
    with t2:
        st.caption("Cada objetivo precisa ser **S**pecífico, **M**ensurável, **A**tingível, **R**elevante e **T**emporal. "
                   "A **perspectiva** liga o objetivo ao Balanced Scorecard, e **contribui para** desenha o mapa estratégico.")
        nomes_obj = [x for x in d["objetivos"]["objetivo"].astype(str) if x.strip()]
        editor("objetivos", {
            "objetivo": T("Objetivo (verbo + o que muda)", required=True, width="large"),
            "perspectiva": S("Perspectiva (BSC)", options=PERSPECTIVAS),
            "indicador": T("Indicador"), "meta": N("Meta", step=1.0), "unidade": T("Unidade"),
            "prazo_meses": N("Prazo (mês do plano)", min_value=0, max_value=120, step=1),
            "responsavel": T("Responsável"),
            "por_que_atingivel": T("Por que é atingível?", width="medium"),
            "relevancia": T("Por que é relevante?", width="medium"),
            "iniciativa": T("Iniciativa / plano de ação", width="medium"),
            "contribui_para": S("Contribui para (objetivo)", options=nomes_obj)}, "objetivo")
        c1, c2 = st.columns([1, 2])
        if c1.button("✨ Sugerir objetivos a partir do plano", width="stretch"):
            sug = pd.DataFrame(objetivos_sugeridos(proj, d, R()))
            existentes = set(d["objetivos"]["objetivo"].astype(str))
            sug = sug[~sug["objetivo"].isin(existentes)]
            recarregar_tabela("objetivos", pd.concat([d["objetivos"], sug], ignore_index=True), "objetivo")
            st.rerun()
        c2.caption("Os rascunhos usam os números do plano (receita, ponto de equilíbrio, churn, horas da equipe). "
                   "Revise e reescreva com as palavras do grupo.")
        ob = d["objetivos"]
        if not ob.empty:
            ok_total = 0
            st.markdown("#### Verificação SMART")
            for _, o in ob.iterrows():
                chk = checar_smart(o)
                n_ok = sum(1 for x in chk.values() if x[0])
                ok_total += n_ok == 5
                with st.container(border=True):
                    st.markdown(f"**{o['objetivo']}** · {o['perspectiva'] or 'sem perspectiva'} · {_selos_smart(chk)}")
                    if n_ok == 5:
                        st.caption("✍️ " + frase_smart(o))
                    for letra, (ok, dica) in chk.items():
                        if not ok:
                            st.caption(f"❌ **{SMART[letra][0]}:** {dica}")
            st.metric("Objetivos que cumprem os 5 critérios", f"{ok_total} de {len(ob)}")
    with t3:
        st.markdown("""
**Missão** responde *por que existimos*: o que a empresa faz, para quem e qual benefício entrega. Evite frases genéricas
("ser a melhor empresa"). Exemplo: *"Oferecer refeições saudáveis e saborosas a quem tem pouco tempo, com entrega rápida e
preço justo."*

**Visão** responde *aonde queremos chegar*: um futuro desafiador, com prazo. Exemplo: *"Ser a principal referência em
alimentação saudável de Teófilo Otoni até 2030."*

**Valores** são os princípios que orientam as decisões e o comportamento de todos na empresa.

**Objetivos SMART** transformam a visão em metas acompanháveis:

| Letra | Critério | Pergunta de checagem |
|---|---|---|
| S | Específico | Está claro *o que* muda? Tem verbo de ação? |
| M | Mensurável | Qual indicador e qual número provam que foi alcançado? |
| A | Atingível | Os recursos do plano (equipe, verba, capacidade) permitem chegar lá? |
| R | Relevante | Ajuda a cumprir a estratégia e a viabilidade do negócio? |
| T | Temporal | Até quando? |

Exemplo ruim: *"Melhorar as vendas."* Exemplo SMART: *"Aumentar a receita mensal para R$ 45.000 até o mês 12, com a
abertura do canal de delivery, sob responsabilidade da sócia comercial."*
""")
    proxima("identidade")


# ------------------------------------------------------------------------------
# BUSINESS MODEL CANVAS
# ------------------------------------------------------------------------------
ORIGEM_CANVAS = {
    "Parcerias-chave": "campo abaixo ou fornecedores críticos", "Atividades-chave": "processos do Plano operacional",
    "Recursos-chave": "investimentos fixos e equipe", "Proposta de valor": "Propostas de valor",
    "Relacionamento com clientes": "campo abaixo ou oportunidades da Jornada (pós-venda)", "Canais": "Canais e funil",
    "Segmentos de clientes": "Segmentação (os mais bem avaliados)", "Estrutura de custos": "DRE do Ano 1",
    "Fontes de receita": "Produtos, serviços e assinaturas",
}


def html_canvas(bl: dict) -> str:
    def caixa(nome, area, cor):
        itens = "".join(f"<li>{x}</li>" for x in bl.get(nome, [])) or "<li><i>preencha a ferramenta de origem</i></li>"
        return (f"<div class='cv-b' style='grid-area:{area};border-top:4px solid {cor}'><div class='cv-t'>{nome}</div>"
                f"<ul>{itens}</ul><div class='cv-o'>vem de: {ORIGEM_CANVAS[nome]}</div></div>")
    css = """<style>
    .cv{display:grid;gap:8px;grid-template-columns:repeat(10,1fr);grid-template-areas:
      "pa pa ak ak pv pv rc rc sc sc" "pa pa kr kr pv pv ch ch sc sc" "cs cs cs cs cs rs rs rs rs rs";}
    .cv-b{background:var(--secondary-background-color,#f4f6fa);border-radius:10px;padding:10px 12px;font-size:.86rem;min-height:130px}
    .cv-t{font-weight:700;margin-bottom:4px} .cv-b ul{margin:0;padding-left:18px} .cv-o{font-size:.7rem;opacity:.6;margin-top:6px}
    @media (max-width:700px){.cv{grid-template-columns:1fr;grid-template-areas:"sc" "pv" "ch" "rc" "rs" "kr" "ak" "pa" "cs";}}
    </style>"""
    return css + "<div class='cv'>" + "".join([
        caixa("Parcerias-chave", "pa", "#8e44ad"), caixa("Atividades-chave", "ak", "#2980b9"),
        caixa("Recursos-chave", "kr", "#2980b9"), caixa("Proposta de valor", "pv", "#c0392b"),
        caixa("Relacionamento com clientes", "rc", "#e67e22"), caixa("Canais", "ch", "#e67e22"),
        caixa("Segmentos de clientes", "sc", "#e67e22"), caixa("Estrutura de custos", "cs", "#16a085"),
        caixa("Fontes de receita", "rs", "#16a085")]) + "</div>"


def pg_canvas():
    cabecalho("canvas")
    st.caption("O quadro é montado automaticamente com o que você já preencheu. Complete os dois blocos que não têm "
               "ferramenta própria (um item por linha).")
    c1, c2 = st.columns(2)
    v = {"canvas_relacionamento": c1.text_area("Relacionamento com clientes", proj["canvas_relacionamento"], height=110,
                                               key=wk("cv_rel"), placeholder="Atendimento pelo WhatsApp\nCartão fidelidade\nComunidade no Instagram"),
         "canvas_parcerias": c2.text_area("Parcerias-chave", proj["canvas_parcerias"], height=110, key=wk("cv_par"),
                                          placeholder="Produtores locais\nAplicativo de entregas\nSecretaria de Turismo")}
    auto_unica("projeto", v)
    st.markdown(html_canvas(canvas_blocos(proj, d, R())), unsafe_allow_html=True)
    proxima("canvas")


# ------------------------------------------------------------------------------
# MATRIZ BCG
# ------------------------------------------------------------------------------
def pg_bcg():
    cabecalho("bcg")
    res = R()
    st.info("A matriz BCG foi criada para analisar um **portfólio**. Em um negócio novo, use-a para as linhas de produto: "
            "informe quanto o mercado de cada uma cresce por ano e sua participação **relativa** (sua fatia ÷ fatia do maior "
            "concorrente; 1,0 = empatado com o líder, 0,5 = metade do líder).")
    produtos = [p for p in res["cat"]["descricao"].astype(str) if p.strip()]
    if res.get("recorrencia", {}).get("tem"):
        produtos.append("Assinaturas (todos os planos)")
    lim = st.slider("A partir de quanto o crescimento do mercado é considerado alto? (% ao ano)", 0.0, 30.0,
                    float(proj["bcg_limite_crescimento_pct"] or 10), 0.5, key=wk("bcg_lim"))
    auto_unica("projeto", {"bcg_limite_crescimento_pct": lim})
    editor("bcg", {"produto": S("Produto / linha", options=produtos, required=True),
                   "crescimento_mercado_pct": N("Crescimento do mercado (% ao ano)", step=0.5),
                   "participacao_relativa": N("Participação relativa (x líder)", min_value=0.0, step=0.1)}, "produto")
    if st.button("Incluir todos os produtos com valores iniciais"):
        faltam = [p for p in produtos if p not in set(d["bcg"]["produto"])]
        novos = pd.DataFrame({"produto": faltam, "crescimento_mercado_pct": proj["crescimento_anual_pct"],
                              "participacao_relativa": 0.5})
        recarregar_tabela("bcg", pd.concat([d["bcg"], novos], ignore_index=True), "produto")
        st.rerun()
    tab = bcg_tabela(d, res, lim)
    if tab.empty:
        return proxima("bcg")
    g = tab.assign(Rótulo=tab["Produto"].str.slice(0, 22), Tam=tab["Receita anual"].clip(lower=1))
    base = alt.Chart(g).encode(
        x=alt.X("Participação relativa:Q", scale=alt.Scale(type="log", domain=[0.1, 10], reverse=True),
                title="Participação relativa de mercado (escala log, líder à esquerda)"),
        y=alt.Y("Crescimento do mercado (%):Q", title="Crescimento do mercado (% a.a.)"))
    pts = base.mark_circle(opacity=0.8, stroke="white").encode(
        size=alt.Size("Tam:Q", legend=None, scale=alt.Scale(range=[150, 2500])),
        color=alt.Color("Quadrante:N", scale=alt.Scale(domain=list(QUADRANTES_BCG), range=["#f1c40f", "#27ae60", "#2980b9", "#c0392b"]),
                        legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["Produto", "Quadrante", alt.Tooltip("Receita anual:Q", format=",.0f"), "Crescimento do mercado (%)", "Participação relativa"])
    txt = base.mark_text(dy=-16, fontSize=11).encode(text="Rótulo:N")
    regras = alt.Chart(pd.DataFrame({"x": [1.0]})).mark_rule(strokeDash=[4, 4]).encode(x="x:Q") + \
        alt.Chart(pd.DataFrame({"y": [lim]})).mark_rule(strokeDash=[4, 4]).encode(y="y:Q")
    st.altair_chart((pts + txt + regras).properties(height=420), width="stretch")
    st.caption("⭐ Estrela (alto, forte) · 🐄 Vaca leiteira (baixo, forte) · ❓ Interrogação (alto, fraca) · 🍍 Abacaxi (baixo, fraca). "
               "O tamanho da bolha é a receita do Ano 1.")
    vis = tab.copy()
    vis["Receita anual"] = vis["Receita anual"].map(brl)
    vis["Margem"] = vis["Margem"].map(pct)
    vis["Quadrante"] = vis["Quadrante"].map(lambda q: f"{QUADRANTES_BCG[q][0]} {q}")
    st.dataframe(vis, hide_index=True, width="stretch")
    for q, (ic, txt_) in QUADRANTES_BCG.items():
        itens = tab.loc[tab["Quadrante"] == q, "Produto"].tolist()
        if itens:
            st.markdown(f"{ic} **{q}** ({', '.join(itens)}): {txt_}")
    proxima("bcg")


# ------------------------------------------------------------------------------
# MATRIZ GE / McKINSEY
# ------------------------------------------------------------------------------
def pg_ge():
    cabecalho("ge")
    st.caption("Cada segmento de clientes é posicionado pela **atratividade do mercado** (média das notas de tamanho, lucro, "
               "longevidade e concorrência da Segmentação) e pela **força competitiva** da empresa nele (automática: média entre "
               "a aderência do produto ao segmento e a nota da empresa frente aos concorrentes; você pode informar a sua).")
    tab = ge_tabela(d)
    if tab.empty:
        st.warning("Cadastre os segmentos em Segmentação de mercado.")
        if st.button("Ir para Segmentação →"):
            ir("segmentos")
            st.rerun()
        return proxima("ge")
    with st.expander("✏️ Informar a força competitiva por segmento (opcional)"):
        editor("ge", {"unidade": S("Segmento", options=d["segmentos"]["nome"].tolist(), required=True),
                      "forca": N("Força competitiva (0 a 10)", min_value=0.0, max_value=10.0, step=0.5),
                      "justificativa": T("Justificativa", width="large")}, "unidade")
        tab = ge_tabela(d)
    cores = {(2, 2): "#27ae60", (2, 1): "#27ae60", (1, 2): "#27ae60", (2, 0): "#f1c40f", (1, 1): "#f1c40f",
             (0, 2): "#f1c40f", (1, 0): "#e74c3c", (0, 1): "#e74c3c", (0, 0): "#e74c3c"}
    celulas = pd.DataFrame([{"x0": j * 10 / 3, "x1": (j + 1) * 10 / 3, "y0": i * 10 / 3, "y1": (i + 1) * 10 / 3, "cor": cores[(i, j)]}
                            for i in range(3) for j in range(3)])
    fundo = alt.Chart(celulas).mark_rect(opacity=0.18).encode(x=alt.X("x0:Q", scale=alt.Scale(domain=[0, 10]), title="Força competitiva"),
                                                              x2="x1", y=alt.Y("y0:Q", scale=alt.Scale(domain=[0, 10]), title="Atratividade do mercado"),
                                                              y2="y1", color=alt.Color("cor:N", scale=None))
    pts = alt.Chart(tab).mark_circle(size=420, color="#1f4f8f", opacity=0.85).encode(
        x="Força competitiva:Q", y="Atratividade:Q", tooltip=["Segmento", "Atratividade", "Força competitiva", "Recomendação"])
    txt = alt.Chart(tab.assign(R=tab["Segmento"].str.slice(0, 24))).mark_text(dy=-18, fontSize=11).encode(
        x="Força competitiva:Q", y="Atratividade:Q", text="R:N")
    st.altair_chart((fundo + pts + txt).properties(height=430), width="stretch")
    st.caption("🟩 Investir e crescer · 🟨 Seletividade / manter · 🟥 Colher ou desinvestir")
    st.dataframe(tab, hide_index=True, width="stretch")
    proxima("ge")


# ------------------------------------------------------------------------------
# CINCO FORÇAS DE PORTER
# ------------------------------------------------------------------------------
def radar_porter(tab: pd.DataFrame, tamanho=(5.2, 5.2)):
    import numpy as np
    vals = tab["Intensidade"].tolist()
    rot = ["Rivalidade", "Novos\nentrantes", "Substitutos", "Clientes", "Fornecedores"]
    ang = np.linspace(0, 2 * np.pi, len(vals), endpoint=False).tolist()
    fig, ax = plt.subplots(figsize=tamanho, subplot_kw={"polar": True})
    ax.plot(ang + ang[:1], vals + vals[:1], color="#2f6fd0", linewidth=2)
    ax.fill(ang + ang[:1], vals + vals[:1], color="#2f6fd0", alpha=0.25)
    ax.set_xticks(ang)
    ax.set_xticklabels(rot, fontsize=10)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=8, color="#777")
    fig.tight_layout()
    return fig


def pg_porter():
    cabecalho("porter")
    sug = porter_sugestoes(proj, d, R())
    if d["porter"].empty:
        ini = pd.DataFrame({"forca": [f for f, _ in FORCAS_PORTER], "intensidade": [sug[f][0] for f, _ in FORCAS_PORTER],
                            "justificativa": [f"(sugestão) {sug[f][1]}" for f, _ in FORCAS_PORTER], "estrategia": ""})
        recarregar_tabela("porter", ini, "forca")
    st.caption("Nota de **1 (fraca)** a **5 (muito forte)**: quanto mais forte a força, mais ela pressiona preços e margens. "
               "O programa sugeriu notas iniciais a partir do plano; revise e justifique com as suas pesquisas.")
    with st.expander("❓ Perguntas que ajudam a avaliar cada força"):
        for f, q in FORCAS_PORTER:
            st.markdown(f"- **{f}:** {q}")
    editor("porter", {"forca": T("Força", disabled=True, width="medium"),
                      "intensidade": N("Intensidade (1 a 5)", min_value=0, max_value=5, step=1),
                      "justificativa": T("Justificativa", width="large"),
                      "estrategia": T("Como o negócio vai se proteger", width="large")}, "forca")
    tab = porter_tabela(d)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.pyplot(radar_porter(tab), width="stretch")
    with c2:
        st.markdown("#### Leitura")
        st.write(porter_leitura(tab))
        st.markdown("#### Sugestões do programa")
        for f, _ in FORCAS_PORTER:
            st.caption(f"**{f}:** nota {sug[f][0]}. {sug[f][1]}")
    proxima("porter")


# ------------------------------------------------------------------------------
# BALANCED SCORECARD E MAPA ESTRATÉGICO
# ------------------------------------------------------------------------------
def dot_mapa(ob: pd.DataFrame) -> str:
    cores = {"Financeira": "#d4efdf", "Clientes": "#fdebd0", "Processos internos": "#d6eaf8", "Aprendizado e crescimento": "#e8daef"}
    def q(s):
        return '"' + str(s).replace('"', "'") + '"'
    linhas = ['digraph G {', 'rankdir=BT; nodesep=0.3; ranksep=0.45;',
              'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11, width=2.2];']
    nomes = {}
    for i, (_, o) in enumerate(ob.iterrows()):
        nomes[o["objetivo"]] = f"o{i}"
        rot = "\\n".join(re.findall(r".{1,28}(?:\s|$)", str(o["objetivo"]))[:4]) or str(o["objetivo"])
        linhas.append(f'o{i} [label={q(rot)}, fillcolor="{cores.get(o["perspectiva"], "#eeeeee")}"];')
    for p in PERSPECTIVAS:
        ids = [nomes[o] for o in ob.loc[ob["perspectiva"] == p, "objetivo"]]
        if ids:
            linhas.append("{rank=same; " + "; ".join(ids) + "}")
    tem_lig = False
    for _, o in ob.iterrows():
        alvo = str(o["contribui_para"] or "").strip()
        if alvo in nomes and alvo != o["objetivo"]:
            linhas.append(f'{nomes[o["objetivo"]]} -> {nomes[alvo]};')
            tem_lig = True
    if not tem_lig:  # ligações automáticas: cada perspectiva sustenta a de cima
        for a, b in zip(PERSPECTIVAS[::-1][:-1], PERSPECTIVAS[::-1][1:]):
            de = [nomes[o] for o in ob.loc[ob["perspectiva"] == a, "objetivo"]]
            para = [nomes[o] for o in ob.loc[ob["perspectiva"] == b, "objetivo"]]
            for x in de:
                for y in para[:1]:
                    linhas.append(f"{x} -> {y} [style=dashed];")
    linhas.append("}")
    return "\n".join(linhas)


def pg_bsc():
    cabecalho("bsc")
    res = R()
    ob = d["objetivos"]
    ind = bsc_indicadores_plano(proj, d, res)
    st.caption("O BSC organiza os objetivos SMART em quatro perspectivas e mostra como o aprendizado sustenta os processos, "
               "que sustentam os clientes, que geram o resultado financeiro. Os objetivos vêm de 🧭 Missão, visão, valores e objetivos.")
    if ob.empty:
        st.warning("Ainda não há objetivos cadastrados.")
        if st.button("Cadastrar objetivos SMART →"):
            ir("identidade", "Na aba Objetivos SMART, use o botão de sugestões ou cadastre os seus, escolhendo a perspectiva de cada um.")
            st.rerun()
    cols = st.columns(2)
    for i, p in enumerate(PERSPECTIVAS[::-1][::-1]):
        with cols[i % 2].container(border=True):
            st.markdown(f"#### {['💰', '🤝', '⚙️', '🌱'][i]} {p}")
            sub = ob[ob["perspectiva"] == p]
            if sub.empty:
                st.caption("Nenhum objetivo nesta perspectiva.")
            for _, o in sub.iterrows():
                chk = checar_smart(o)
                st.markdown(f"**{o['objetivo']}**  \n{o['indicador']}: meta {num_br(o['meta'], 1)} {o['unidade']} até o mês "
                            f"{int(o['prazo_meses'] or 0)} · {_selos_smart(chk)}  \n*Iniciativa:* {o['iniciativa'] or '-'}")
            if ind.get(p):
                st.caption("Indicadores do plano: " + " · ".join(f"{n} = {v}" for n, v in ind[p]))
    if not ob.empty:
        st.markdown("#### 🗺️ Mapa estratégico")
        st.graphviz_chart(dot_mapa(ob), width="stretch")
        if not ob["contribui_para"].astype(str).str.strip().ne("").any():
            st.caption("Ligações tracejadas são automáticas. Para desenhar as suas, preencha \"Contribui para\" nos objetivos.")
    faltam = [p for p in PERSPECTIVAS if p not in set(ob["perspectiva"])]
    if faltam and not ob.empty:
        st.info("Perspectivas sem objetivo: " + ", ".join(faltam) + ".")
    proxima("bsc")


# ------------------------------------------------------------------------------
# REFORMA TRIBUTÁRIA
# ------------------------------------------------------------------------------
def aliquotas_padrao_reforma(ano: int) -> tuple:
    """CBS: referência de 11,3% calculada pelo TCU para 2027 (sujeita ao Senado). IBS: 0,1% em 2027-2028; de 2029 a 2032
    cresce 10% ao ano da alíquota cheia estimada (17,7%) e fica integral em 2033. Valores editáveis."""
    if ano <= 2028:
        return 11.3, 0.1
    if ano >= 2033:
        return 11.3, 17.7
    return 11.3, round(17.7 * 0.1 * (ano - 2028), 2)


def pg_reforma():
    cabecalho("reforma")
    res = R()
    a1 = res["anos"][1]
    st.markdown("""
<div class="pn-card-info">A partir de 2027, a empresa do Simples pode escolher entre continuar pagando tudo no <b>DAS</b>
("Simples puro") ou recolher <b>IBS e CBS por fora</b>, no regime regular ("Simples híbrido"). No híbrido ela paga IVA com
crédito sobre as compras, e os <b>clientes empresas</b> aproveitam crédito integral, o que torna o preço mais competitivo
em vendas B2B. Esta simulação é didática e simplificada: confirme a decisão com um contador.</div>""", unsafe_allow_html=True)
    st.write("")
    c1, c2, c3 = st.columns(3)
    ano = int(c1.select_slider("Ano simulado", list(range(2027, 2034)), int(proj["reforma_ano"] or 2027), key=wk("rf_ano")))
    cbs0, ibs0 = aliquotas_padrao_reforma(ano)
    anexos = list(TABELAS_SIMPLES)
    an0 = proj["anexo_simples"] if proj["anexo_simples"] in anexos else (
        "Anexo III - Serviços" if str(proj["tipo_negocio"]).startswith(("Serviço", "Assinatura")) else "Anexo I - Comércio")
    anexo = c2.selectbox("Anexo do Simples", anexos, index=anexos.index(an0), key=wk("rf_anexo"))
    b2b = c3.slider("Vendas para empresas que aproveitam crédito (%)", 0, 100, int(proj["reforma_b2b_pct"] or 0), 5, key=wk("rf_b2b"))
    c1, c2, c3 = st.columns(3)
    cbs = c1.number_input("CBS (%)", 0.0, 30.0, cbs0, 0.1, key=wk(f"rf_cbs_{ano}"))
    ibs = c2.number_input("IBS (%)", 0.0, 30.0, ibs0, 0.1, key=wk(f"rf_ibs_{ano}"))
    compras_pad = a1["custo_dir"] + a1["fixos"]
    cc = float(proj["reforma_compras_credito"])
    compras = c3.number_input("Compras com direito a crédito no ano (R$)", 0.0, 1e10, compras_pad if cc < 0 else cc, 1000.0,
                              key=wk("rf_compras"), help="Sugestão: custos diretos + custos fixos não salariais do Ano 1. "
                                                         "Salários e pró-labore não geram crédito.")
    auto_unica("projeto", {"reforma_ano": ano, "reforma_b2b_pct": b2b, "reforma_cbs_pct": cbs, "reforma_ibs_pct": ibs,
                           "reforma_compras_credito": compras})
    r = simular_reforma(a1["rec_bruta"], res["aliq"], anexo, ano, cbs / 100, ibs / 100, b2b / 100, compras)
    tab = pd.DataFrame({
        "Item": ["DAS", "IBS + CBS sobre as vendas", "(−) Crédito de IBS + CBS das compras", "Total de tributos (DAS + IBS/CBS)",
                 "Crédito que os clientes empresas aproveitam"],
        "Simples puro": [brl(r["das_puro"]), "incluso no DAS", "não aproveita", brl(r["puro_total"]), brl(r["puro_credito_cliente"])],
        "Simples híbrido": [brl(r["das_hibrido"]), brl(r["iva_debito"]), brl(-r["iva_credito"]), brl(r["hib_total"]),
                            brl(r["hib_credito_cliente"])]})
    st.dataframe(tab, hide_index=True, width="stretch")
    k = st.columns(3)
    k[0].metric("Diferença de imposto (híbrido − puro)", brl(r["diferenca"]))
    k[1].metric("Crédito a mais para clientes B2B no híbrido", brl(r["hib_credito_cliente"] - r["puro_credito_cliente"]))
    k[2].metric("Parcela do DAS substituída pela reforma", pct(r["parcela_substituida"]))
    if b2b == 0:
        st.info("Vendendo só para consumidor final, o crédito não beneficia o cliente: o Simples puro tende a ser mais vantajoso.")
    elif r["diferenca"] > 0 and r["hib_credito_cliente"] - r["puro_credito_cliente"] > r["diferenca"]:
        st.success("O híbrido custa mais imposto para a empresa, mas gera aos clientes empresas um crédito maior que essa diferença: "
                   "pode valer a pena se isso permitir vender mais ou cobrar um preço melhor no B2B.")
    elif r["diferenca"] <= 0:
        st.success("Neste cenário o híbrido sai mais barato e ainda gera mais crédito aos clientes.")
    else:
        st.warning("Neste cenário o híbrido aumenta o imposto mais do que o crédito que gera aos clientes.")
    st.caption("Fontes: LC 214/2025 (art. 41, opção pelo regime regular do IBS e da CBS); partilha do DAS da 1ª faixa da LC 123/2006 "
               "com redação da LC 155/2016; CBS de referência de 11,3% para 2027 calculada pelo TCU, sujeita à definição do Senado; "
               "estimativas de 26,5% a 28,6% para IBS + CBS em 2033. Todos os percentuais são editáveis.")
    proxima("reforma")
