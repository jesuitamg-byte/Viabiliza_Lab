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


# ------------------------------------------------------------------------------
# DESENHOS DAS MATRIZES (HTML/SVG responsivos)
# ------------------------------------------------------------------------------
import html as _html

CORES_GE = {"Seguro e crescimento": ("#3fd0c0", "🛡️"), "Seletivo ou cauteloso": ("#ffd23f", "➤"),
            "Expansão limitada ou colheita": ("#ffa630", "↗"), "Perigo ou colheita": ("#ec4561", "⚠️")}
# linhas de cima para baixo (atratividade alta -> baixa); colunas da esquerda para a direita (força alta -> baixa)
GRADE_GE = [["Seguro e crescimento", "Seguro e crescimento", "Seletivo ou cauteloso"],
            ["Seguro e crescimento", "Seletivo ou cauteloso", "Expansão limitada ou colheita"],
            ["Seletivo ou cauteloso", "Expansão limitada ou colheita", "Perigo ou colheita"]]


def _esc(s) -> str:
    return _html.escape(str(s))


def html_ge(tab: pd.DataFrame) -> str:
    celulas = "".join(
        f"<div class='ge-c' style='background:{CORES_GE[n][0]}'><div class='ge-i'>{CORES_GE[n][1]}</div><div>{n.replace(' ou ', ' ou<br>')}</div></div>"
        for linha in GRADE_GE for n in linha)
    pts = ""
    for k, (_, r) in enumerate(tab.iterrows(), start=1):
        left = min(max((10 - float(r["Força competitiva"])) / 10 * 100, 3), 97)
        top = min(max((10 - float(r["Atratividade"])) / 10 * 100, 3), 97)
        pts += (f"<div class='ge-p' style='left:{left:.1f}%;top:{top:.1f}%' title='{_esc(r['Segmento'])}: atratividade "
                f"{num_br(r['Atratividade'], 1)}, força {num_br(r['Força competitiva'], 1)}'>{k}</div>")
    leg = "".join(f"<li><b>{k}.</b> {_esc(r['Segmento'])} <span>→ {_esc(r['Recomendação'])}</span></li>"
                  for k, (_, r) in enumerate(tab.iterrows(), start=1))
    css = """<style>
    .ge-w{display:flex;gap:10px;align-items:stretch;max-width:760px;margin:auto}
    .ge-y{writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;color:#7a8595;font-size:.85rem;border-left:2px solid #7a8595;padding-left:6px}
    .ge-m{flex:1;position:relative}
    .ge-g{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;aspect-ratio:1.15/1}
    .ge-c{border-radius:14px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;
          font-weight:600;font-size:clamp(.62rem,1.6vw,.98rem);color:#253041;padding:4px;line-height:1.2}
    .ge-i{font-size:1.5em;margin-bottom:4px;opacity:.75}
    .ge-p{position:absolute;transform:translate(-50%,-50%);width:34px;height:34px;border-radius:50%;background:#1f3f75;color:#fff;
          border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;font-weight:700}
    .ge-x{text-align:center;color:#7a8595;font-size:.85rem;border-top:2px solid #7a8595;margin-top:8px;padding-top:4px}
    .ge-l{max-width:760px;margin:10px auto 0;padding-left:20px} .ge-l li{margin:2px 0} .ge-l span{opacity:.7}
    </style>"""
    return (css + "<div class='ge-w'><div class='ge-y'>Baixa ← Atratividade do mercado → Alta</div><div class='ge-m'>"
            f"<div class='ge-g'>{celulas}</div>{pts}<div class='ge-x'>Alta ← Força competitiva da empresa no segmento → Baixa</div>"
            f"</div></div><ol class='ge-l' style='list-style:none'>{leg}</ol>")


QUAD_BCG_VIS = {"Estrela": ("#ede3fb", "#7b3fe4", "⭐", ["Mercado atrativo", "Gera bom faturamento", "Exige investimento"]),
                "Interrogação": ("#fff4c7", "#c99700", "❓", ["Ainda gera pouco retorno", "Expectativa de crescimento", "Decisão de risco"]),
                "Vaca leiteira": ("#eefbd0", "#5e8f00", "🐄", ["Gera muito caixa", "Exige pouco investimento", "Recursos realocáveis"]),
                "Abacaxi": ("#fde2e5", "#c8102e", "🍍", ["Gera pouco caixa", "Mercado sem crescimento", "Provável saída"])}


def html_bcg(tab: pd.DataFrame, lim: float) -> str:
    import math
    ymax = max(2 * lim, 1.0)
    quad = ""
    for q in ("Estrela", "Interrogação", "Vaca leiteira", "Abacaxi"):
        bg, cor, ic, itens = QUAD_BCG_VIS[q]
        al = "flex-end" if q in ("Vaca leiteira", "Abacaxi") else "flex-start"
        quad += (f"<div class='bc-q' style='background:{bg};justify-content:{al}'><div class='bc-t' style='color:{cor}'>{ic} {q}</div>"
                 + "".join(f"<div class='bc-i'>✓ {i}</div>" for i in itens) + "</div>")
    mx = max(float(tab["Receita anual"].max()), 1.0)
    pts = ""
    for _, r in tab.iterrows():
        p = min(max(float(r["Participação relativa"]), 0.1), 10)
        left = (1 - math.log10(p)) / 2 * 100
        top = (1 - min(max(float(r["Crescimento do mercado (%)"]), 0), ymax) / ymax) * 100
        tam = 26 + 50 * math.sqrt(max(float(r["Receita anual"]), 0) / mx)
        cor = QUAD_BCG_VIS[r["Quadrante"]][1]
        pts += (f"<div class='bc-p' style='left:{min(max(left, 4), 96):.1f}%;top:{min(max(top, 5), 95):.1f}%;width:{tam:.0f}px;"
                f"height:{tam:.0f}px;background:{cor}' title='{_esc(r['Produto'])}: crescimento {num_br(r['Crescimento do mercado (%)'], 1)}%, "
                f"participação {num_br(r['Participação relativa'], 2)}x, receita {brl(r['Receita anual'])}'></div>"
                f"<div class='bc-r' style='left:{min(max(left, 8), 92):.1f}%;top:calc({min(max(top, 5), 95):.1f}% + {tam / 2 + 2:.0f}px)'>"
                f"{_esc(str(r['Produto'])[:26])}</div>")
    yt = "".join(f"<span style='top:{(1 - v / ymax) * 100:.0f}%'>{num_br(v, 0)}%</span>" for v in (0, lim / 2, lim, 1.5 * lim, ymax))
    xt = "".join(f"<span style='left:{(1 - math.log10(v)) / 2 * 100:.0f}%'>{num_br(v, 1 if v < 1 else 0)}x</span>" for v in (10, 5, 1, 0.5, 0.1))
    css = """<style>
    .bc-w{display:flex;gap:6px;max-width:780px;margin:auto}
    .bc-ya{writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;font-weight:700;color:#334;font-size:.85rem}
    .bc-yt{position:relative;width:34px;font-size:.72rem;color:#556} .bc-yt span{position:absolute;right:2px;transform:translateY(-50%)}
    .bc-m{flex:1;position:relative}
    .bc-g{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;aspect-ratio:1.2/1;border:2px solid #2d4f8a;border-radius:6px;overflow:hidden}
    .bc-q{display:flex;flex-direction:column;padding:10px 14px;border:1px solid rgba(45,79,138,.35);font-size:clamp(.62rem,1.4vw,.85rem);color:#253041}
    .bc-t{font-weight:800;font-size:1.25em;margin-bottom:4px} .bc-i{opacity:.8;line-height:1.35}
    .bc-p{position:absolute;transform:translate(-50%,-50%);border-radius:50%;opacity:.82;border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.3)}
    .bc-r{position:absolute;transform:translateX(-50%);font-size:.72rem;font-weight:600;background:rgba(255,255,255,.85);
          padding:0 4px;border-radius:4px;white-space:nowrap;color:#223}
    .bc-xt{position:relative;height:18px;font-size:.72rem;color:#556} .bc-xt span{position:absolute;transform:translateX(-50%)}
    .bc-xa{text-align:center;font-weight:700;color:#334;font-size:.85rem}
    @media (max-width:640px){.bc-i{display:none}.bc-r{font-size:.62rem}.bc-ya,.bc-xa{font-size:.72rem}}
    </style>"""
    return (css + "<div class='bc-w'><div class='bc-ya'>Baixo ← Crescimento do mercado (% a.a.) → Alto</div>"
            f"<div class='bc-yt'>{yt}</div><div class='bc-m'><div class='bc-g'>{quad}</div>{pts}<div class='bc-xt'>{xt}</div>"
            "<div class='bc-xa'>Alta ← Participação relativa de mercado (x líder) → Baixa</div></div></div>")


CICLO_BSC = [("Aprendizado e crescimento", "#0d6fb8", "Conhecimento e habilidades da equipe sustentam inovações e melhorias."),
             ("Processos internos", "#123a5e", "Equipe capacitada melhora os processos de trabalho."),
             ("Clientes", "#d7262e", "Processos melhores conduzem a clientes mais satisfeitos."),
             ("Financeira", "#b5191f", "Clientes satisfeitos levam a melhores resultados financeiros.")]


def svg_bsc(ob: pd.DataFrame) -> str:
    """Ciclo do BSC: anel com quatro segmentos em forma de seta, no sentido horário."""
    import math
    cx, cy, R1, R0 = 300, 260, 245, 120
    Rm = (R1 + R0) / 2

    def pt(r, a):
        return cx + r * math.cos(math.radians(a)), cy + r * math.sin(math.radians(a))

    partes = []
    inicio = [-90, 0, 90, 180]
    for (nome, cor, desc), a0 in zip(CICLO_BSC, inicio):
        a1 = a0 + 88
        ponta = a1 + 7
        o0, o1 = pt(R1, a0), pt(R1, a1)
        i0, i1 = pt(R0, a0), pt(R0, a1)
        tip, notch = pt(Rm, ponta), pt(Rm, a0 + 7)
        d = (f"M{o0[0]:.1f},{o0[1]:.1f} A{R1},{R1} 0 0 1 {o1[0]:.1f},{o1[1]:.1f} L{tip[0]:.1f},{tip[1]:.1f} "
             f"L{i1[0]:.1f},{i1[1]:.1f} A{R0},{R0} 0 0 0 {i0[0]:.1f},{i0[1]:.1f} L{notch[0]:.1f},{notch[1]:.1f} Z")
        partes.append(f"<path d='{d}' fill='{cor}'/>")
        tx, ty = pt(Rm - 2, a0 + 51)
        n = int((ob["perspectiva"] == nome).sum()) if not ob.empty else 0
        palavras = desc.split()
        linhas, atual = [], ""
        for w in palavras:
            if len(atual) + len(w) > 19:
                linhas.append(atual)
                atual = w
            else:
                atual = (atual + " " + w).strip()
        linhas.append(atual)
        titulo = {"Aprendizado e crescimento": ["APRENDIZADO", "E CRESCIMENTO"], "Processos internos": ["PROCESSOS", "INTERNOS"]}.get(
            nome, [nome.upper()])
        y0 = ty - (len(titulo) * 15 + len(linhas) * 14 + 16) / 2 + 12
        txt = "".join(f"<tspan x='{tx:.0f}' y='{y0 + i * 15:.0f}' font-weight='800' font-size='11.5'>{_esc(t)}</tspan>" for i, t in enumerate(titulo))
        y1 = y0 + len(titulo) * 15 + 2
        txt += "".join(f"<tspan x='{tx:.0f}' y='{y1 + i * 14:.0f}' font-size='11'>{_esc(t)}</tspan>" for i, t in enumerate(linhas))
        txt += (f"<tspan x='{tx:.0f}' y='{y1 + len(linhas) * 14 + 4:.0f}' font-size='11.5' font-weight='700'>"
                f"{n} objetivo{'s' if n != 1 else ''}</tspan>")
        partes.append(f"<text fill='white' text-anchor='middle' font-family='Helvetica,Arial,sans-serif'>{txt}</text>")
    centro = (f"<circle cx='{cx}' cy='{cy}' r='{R0 - 8}' fill='white'/>"
              f"<text x='{cx}' y='{cy + 26}' text-anchor='middle' font-family='Helvetica,Arial,sans-serif' font-weight='900' "
              f"font-size='74' fill='#333'>BSC</text>")
    return (f"<svg viewBox='0 0 600 520' style='width:100%;max-width:600px;display:block;margin:auto' "
            f"xmlns='http://www.w3.org/2000/svg'>{''.join(partes)}{centro}</svg>")


def html_canvas(bl: dict) -> str:
    def caixa(nome, area, cor):
        itens = "".join(f"<li>{_esc(x)}</li>" for x in bl.get(nome, [])) or "<li><i>preencha a ferramenta de origem</i></li>"
        return (f"<div class='cv-b' style='grid-area:{area};background:{cor}'><div class='cv-t'>{nome.upper()}</div>"
                f"<ul>{itens}</ul><div class='cv-o'>vem de: {ORIGEM_CANVAS[nome]}</div></div>")
    css = """<style>
    .cv{display:grid;gap:8px;grid-template-columns:repeat(10,1fr);grid-template-areas:
      "pa pa ak ak pv pv rc rc sc sc" "pa pa kr kr pv pv ch ch sc sc" "cs cs cs cs cs rs rs rs rs rs";}
    .cv-b{border-radius:12px;padding:14px 14px;font-size:.84rem;min-height:150px;color:#fff}
    .cv-t{font-weight:800;margin-bottom:6px;letter-spacing:.02em;font-size:.95rem;line-height:1.15}
    .cv-b ul{margin:0;padding-left:18px} .cv-b li{margin:2px 0} .cv-o{font-size:.7rem;opacity:.75;margin-top:8px}
    @media (max-width:700px){.cv{grid-template-columns:1fr;grid-template-areas:"sc" "pv" "ch" "rc" "rs" "kr" "ak" "pa" "cs";}}
    </style>"""
    return css + "<div class='cv'>" + "".join([
        caixa("Parcerias-chave", "pa", "#6a3184"), caixa("Atividades-chave", "ak", "#f0901f"),
        caixa("Recursos-chave", "kr", "#e5653e"), caixa("Proposta de valor", "pv", "#44923a"),
        caixa("Relacionamento com clientes", "rc", "#8db833"), caixa("Canais", "ch", "#14a49a"),
        caixa("Segmentos de clientes", "sc", "#1592c4"), caixa("Estrutura de custos", "cs", "#4456a0"),
        caixa("Fontes de receita", "rs", "#22617a")]) + "</div>"


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
    st.markdown(html_bcg(tab, lim), unsafe_allow_html=True)
    st.caption(f"O tamanho da bolha é a receita do Ano 1. A linha horizontal fica em {num_br(lim, 1)}% a.a. (limite escolhido acima) "
               "e a vertical em 1,0x (empatado com o líder). Passe o mouse sobre as bolhas para ver os valores.")
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
    st.markdown(html_ge(tab), unsafe_allow_html=True)
    st.write("")
    for n, (cor, ic) in CORES_GE.items():
        if (tab["Recomendação"] == n).any():
            st.markdown(f"{ic} **{n}:** {ACOES_GE[n]}")
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
    st.markdown(svg_bsc(ob), unsafe_allow_html=True)
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
