# ==============================================================================
# VIABILIZALAB - INTERFACE: VALIDAÇÃO, MARKETING, INCERTEZA E DEFESA (v7.1)
# ==============================================================================


def MC() -> dict:
    """Simulação de Monte Carlo com os dados atuais (recalcula quando algo muda)."""
    if "mc" not in _CACHE:
        res = R()
        _CACHE["mc"] = monte_carlo(res, faixas_mc(d)) if res["anos"][1]["rec_bruta"] > 0 else None
    return _CACHE["mc"]


def _semear(tabela: str, linhas: list, chave_: str):
    recarregar_tabela(tabela, pd.concat([d[tabela], pd.DataFrame(linhas)], ignore_index=True), chave_)
    st.rerun()


def _ir_botao(rotulo: str, destino: str, chave_: str):
    if st.button(rotulo, key=chave_):
        ir(destino)
        st.rerun()


# ------------------------------------------------------------------------------
# PLACAR DE EVIDÊNCIAS
# ------------------------------------------------------------------------------
def pg_evidencias():
    cabecalho("evidencias")
    res = R()
    st.markdown("""<div class="pn-card-info">Um plano de negócio é um conjunto de <b>hipóteses</b> até que os números sejam testados
com clientes reais (Ries, 2012; Blank, 2013). Marque a origem de cada premissa. O <b>peso</b> vem da análise de sensibilidade:
quanto mais a premissa mexe no resultado, mais ela conta no placar.</div>""", unsafe_allow_html=True)
    st.write("")
    pr = premissas_plano(proj, d, res)
    faltam = [p for p in pr["premissa"] if p not in set(d["evidencias"]["premissa"])]
    if faltam:
        recarregar_tabela("evidencias", pd.concat([d["evidencias"], pd.DataFrame(
            {"premissa": faltam, "status": "Estimada (sem fonte)", "fonte": "", "observacao": ""})], ignore_index=True), "premissa")
        st.rerun()
    pl = placar_evidencias(proj, d, res)
    k = st.columns(3)
    k[0].metric("Placar de evidências", pct(pl["score"], 0), help="Média das notas (estimada = 0; pesquisada = 0,5; testada = 1) "
                                                                  "ponderada pelo impacto de cada premissa no resultado.")
    k[1].metric("Premissas testadas com clientes", f"{pl['testadas']} de {pl['n']}")
    k[2].metric("Valide primeiro", pl["proxima"] or "tudo testado")
    st.progress(min(max(pl["score"], 0.0), 1.0))
    vis = pl["tabela"][["premissa", "valor_plano", "peso", "status"]].rename(columns={
        "premissa": "Premissa", "valor_plano": "Valor no plano", "peso": "Peso", "status": "Origem"})
    vis["Peso"] = vis["Peso"].map(lambda v: pct(v, 0))
    vis["Origem"] = vis["Origem"].map(STATUS_CURTO)
    st.dataframe(vis, hide_index=True, width="stretch")
    st.subheader("Informe a origem de cada premissa")
    ev = d["evidencias"][d["evidencias"]["premissa"].isin(set(pr["premissa"]))].reset_index(drop=True)

    def _salvar(x):
        outras = d["evidencias"][~d["evidencias"]["premissa"].isin(set(x["premissa"]))]
        completo = pd.concat([outras, x], ignore_index=True)
        salvar(conn, "evidencias", completo, "premissa")
        d["evidencias"] = limpar(completo.reset_index(drop=True), "evidencias")

    tabela_editavel(ev, {"premissa": T("Premissa", disabled=True, width="medium"),
                         "status": S("Origem do número", options=list(STATUS_EVIDENCIA), required=True, width="large"),
                         "fonte": T("Fonte (quem, onde, quando)", width="large"),
                         "observacao": T("O que o teste mostrou", width="large")},
                    "evidencias", "placar", normalizar=lambda x: limpar(x, "evidencias"), salvar_fn=_salvar, adicionar=False)
    st.caption("Exemplos de evidência: entrevistas registradas, respostas do teste de preço, pré-vendas, lista de espera, "
               "orçamentos de fornecedores, dados do IBGE ou do Sebrae, preços dos concorrentes.")
    c = st.columns(3)
    with c[0]:
        _ir_botao("🎙️ Registrar entrevistas", "entrevistas", "ev_ent")
    with c[1]:
        _ir_botao("🙋 Testar o preço", "teste_preco", "ev_vw")
    with c[2]:
        _ir_botao("🧪 Planejar experimentos", "experimentos", "ev_exp")
    proxima("evidencias")


# ------------------------------------------------------------------------------
# ENTREVISTAS COM CLIENTES
# ------------------------------------------------------------------------------
def pg_entrevistas():
    cabecalho("entrevistas")
    res = R()
    with st.expander("📋 Roteiro de entrevista de descoberta", expanded=d["entrevistas"].empty):
        st.markdown(ROTEIRO_ENTREVISTA)
    st.caption("Registre uma linha por pessoa. Em **dores**, separe as dores por ponto e vírgula e use sempre as mesmas palavras "
               "para a mesma dor: assim o programa consegue contá-las.")
    editor("entrevistas", {"data": T("Data"), "entrevistado": T("Entrevistado (nome ou código)", required=True),
                           "perfil": T("Perfil / segmento"), "dores": T("Dores citadas (separe com ;)", width="large"),
                           "solucao_atual": T("O que usa hoje", width="medium"),
                           "pagaria": S("Pagaria?", options=PAGARIA), "valor_pagaria": moeda("Quanto pagaria"),
                           "aprendizado": T("Principal aprendizado", width="large")}, "entrevistado")
    preco = preco_referencia(proj, res)
    r = resumo_entrevistas(d, preco)
    if r["n"] == 0:
        return proxima("entrevistas")
    k = st.columns(4)
    k[0].metric("Entrevistas", r["n"], help="Recomenda-se ao menos 10 a 15 por segmento antes de confiar nos padrões.")
    k[1].metric("Pagariam", pct(r["pct_sim"], 0), f"+ {pct(r['pct_talvez'], 0)} talvez", delta_color="off")
    k[2].metric("Valor mediano citado", brl(r["mediana_valor"]) if r["mediana_valor"] == r["mediana_valor"] else "n.a.")
    k[3].metric("Preço do plano", brl(preco) if preco == preco else "n.a.",
                f"{pct(r['valor_vs_preco'] - 1, 0)} vs. citado" if r["valor_vs_preco"] == r["valor_vs_preco"] else None,
                delta_color="off")
    if not r["dores"].empty:
        g = r["dores"].assign(Pct=r["dores"]["% das entrevistas"] * 100)
        st.altair_chart(alt.Chart(g).mark_bar(color="#2f6fd0").encode(
            x=alt.X("Pct:Q", title="% das entrevistas", scale=alt.Scale(domain=[0, 100])),
            y=alt.Y("Dor citada:N", sort="-x", title=None), tooltip=["Dor citada", "Entrevistas"]).properties(height=34 * len(g) + 40),
            width="stretch")
        top, qt = r["dor_top"]
        if qt >= r["n"] / 2:
            st.success(f"A dor **\"{top}\"** apareceu em {qt} de {r['n']} entrevistas: é um problema recorrente, bom sinal para a proposta de valor.")
        else:
            st.warning("Nenhuma dor apareceu em metade das entrevistas. Talvez o segmento esteja amplo demais ou o problema não seja tão forte.")
    if r["n"] < 10:
        st.info(f"Faltam {10 - r['n']} entrevista(s) para chegar ao mínimo de 10 usado na régua de prontidão.")
    proxima("entrevistas")


# ------------------------------------------------------------------------------
# TESTE DE PREÇO (VAN WESTENDORP) E INTENÇÃO DE COMPRA
# ------------------------------------------------------------------------------
def _grafico_vw(vw: dict, preco: float):
    cv = vw["curvas"].melt(id_vars="Preço", value_vars=["Muito barato", "Barato", "Caro", "Muito caro"],
                           var_name="Curva", value_name="Proporção")
    linhas = alt.Chart(cv).mark_line(strokeWidth=2.5).encode(
        x=alt.X("Preço:Q", title="Preço (R$)"), y=alt.Y("Proporção:Q", axis=alt.Axis(format="%"), title="% dos respondentes"),
        color=alt.Color("Curva:N", scale=alt.Scale(domain=["Muito barato", "Barato", "Caro", "Muito caro"],
                                                   range=["#7fb3d5", "#2471a3", "#e59866", "#c0392b"]),
                        legend=alt.Legend(orient="bottom", title=None)))
    faixa = alt.Chart(pd.DataFrame({"a": [vw["pmc"]], "b": [vw["pme"]]})).mark_rect(opacity=0.12, color="#27ae60").encode(x="a:Q", x2="b:Q")
    camadas = faixa + linhas
    if preco == preco:
        camadas += alt.Chart(pd.DataFrame({"p": [preco], "t": ["preço do plano"]})).mark_rule(strokeDash=[5, 4], color="black").encode(
            x="p:Q", tooltip=["t"])
    st.altair_chart(camadas.properties(height=380), width="stretch")


def pg_teste_preco():
    cabecalho("teste_preco")
    res = R()
    t1, t2 = st.tabs(["🙋 Teste de preço (Van Westendorp)", "📊 Intenção de compra ajustada"])
    with t1:
        st.caption("Pergunte a pessoas do público-alvo as quatro questões abaixo, sobre o mesmo produto, e lance as respostas. "
                   "O método (Van Westendorp, 1976) mostra a faixa de preço que os clientes consideram aceitável.")
        produtos = [p for p in res["cat"]["descricao"].astype(str) if p.strip()]
        if res.get("recorrencia", {}).get("tem"):
            produtos.append("Assinatura (ticket médio)")
        opcoes = ["(preço médio do plano)"] + produtos
        atual = proj["vw_produto"] if proj["vw_produto"] in opcoes else "(preço médio do plano)"
        esc = st.selectbox("Produto testado", opcoes, index=opcoes.index(atual), key=wk("vw_prod"))
        auto_unica("projeto", {"vw_produto": "" if esc.startswith("(") else esc})
        nome = esc if not esc.startswith("(") else "o produto"
        with st.expander("📝 Perguntas prontas para o Google Forms"):
            st.code(f"1. A que preço você acharia {nome} tão barato que duvidaria da qualidade?\n"
                    f"2. A que preço você acharia {nome} barato, uma boa compra?\n"
                    f"3. A que preço {nome} começaria a parecer caro, mas ainda pensaria em comprar?\n"
                    f"4. A que preço {nome} ficaria caro demais, e você não compraria de jeito nenhum?", language=None)
            st.caption("Use respostas numéricas (R$). Exporte as respostas para planilha e cole nas colunas abaixo.")
        editor("precos_vw", {"respondente": T("Respondente", required=True), "muito_barato": moeda("1. Barato demais"),
                             "barato": moeda("2. Barato"), "caro": moeda("3. Caro"), "muito_caro": moeda("4. Caro demais")},
               "respondente")
        preco = preco_referencia(proj, res)
        vw = van_westendorp(d["precos_vw"])
        if vw is None:
            st.info("Lance pelo menos 5 respostas para ver as curvas.")
        elif not vw["ok"]:
            st.info(f"{vw['n']} resposta(s) válida(s). São necessárias pelo menos 5 (de preferência 30 ou mais). "
                    "Respostas fora de ordem (barato maior que caro, por exemplo) são descartadas.")
        else:
            _grafico_vw(vw, preco)
            k = st.columns(4)
            k[0].metric("Limite inferior (PMC)", brl(vw["pmc"]), help="Abaixo disso o cliente desconfia da qualidade.")
            k[1].metric("Preço ótimo (OPP)", brl(vw["opp"]), help="Menos pessoas acham barato demais ou caro demais.")
            k[2].metric("Preço de indiferença (IPP)", brl(vw["ipp"]), help="Tantos acham caro quanto barato.")
            k[3].metric("Limite superior (PME)", brl(vw["pme"]), help="Acima disso a maioria considera caro demais.")
            if preco == preco:
                if vw["pmc"] <= preco <= vw["pme"]:
                    st.success(f"O preço do plano ({brl(preco)}) está dentro da faixa aceitável. Com {vw['n']} respostas, "
                               "esta premissa pode ser marcada como **testada com clientes** no placar de evidências.")
                elif preco > vw["pme"]:
                    st.error(f"O preço do plano ({brl(preco)}) está **acima** do que os clientes aceitam ({brl(vw['pme'])}). "
                             "Reveja custos, a proposta de valor ou o público.")
                else:
                    st.warning(f"O preço do plano ({brl(preco)}) está **abaixo** da faixa ({brl(vw['pmc'])}): há espaço para "
                               "cobrar mais, e preço baixo demais pode passar impressão de baixa qualidade.")
            if vw["n"] < 30:
                st.caption(f"Com {vw['n']} respostas o resultado é exploratório. Amostras de 30 ou mais tornam as curvas mais estáveis.")
    with t2:
        st.caption("Pesquisas de intenção superestimam a demanda: quem diz que \"certamente compraria\" nem sempre compra. "
                   "Aplique fatores de conversão à intenção declarada e compare com o número de clientes do plano.")
        c = st.columns(4)
        ic = c[0].number_input("Certamente compraria (%)", 0.0, 100.0, float(proj["intencao_certamente_pct"] or 0), 1.0, key=wk("int_c"))
        ip = c[1].number_input("Provavelmente compraria (%)", 0.0, 100.0, float(proj["intencao_provavel_pct"] or 0), 1.0, key=wk("int_p"))
        fc = c[2].number_input("Dos que \"certamente\", quantos compram (%)", 0.0, 100.0, float(proj["fator_certamente_pct"] or 70),
                               5.0, key=wk("fat_c"))
        fp = c[3].number_input("Dos que \"provavelmente\", quantos compram (%)", 0.0, 100.0, float(proj["fator_provavel_pct"] or 30),
                               5.0, key=wk("fat_p"))
        auto_unica("projeto", {"intencao_certamente_pct": ic, "intencao_provavel_pct": ip, "fator_certamente_pct": fc,
                               "fator_provavel_pct": fp})
        ia = intencao_ajustada(proj, res)
        if not ia["ok"]:
            st.info("Preencha o SAM (clientes alcançáveis) em Tamanho do mercado e os percentuais da pesquisa.")
        else:
            k = st.columns(3)
            k[0].metric("Clientes pela intenção declarada", num_br(ia["declarados"], 0))
            k[1].metric("Clientes após o ajuste", num_br(ia["ajustados"], 0))
            k[2].metric("Clientes que o plano supõe no Ano 1", num_br(ia["plano"], 0) if ia["plano"] == ia["plano"] else "n.a.",
                        help="Receita do Ano 1 ÷ gasto anual por cliente (Tamanho do mercado).")
            if ia["cobertura"] == ia["cobertura"]:
                if ia["cobertura"] >= 1:
                    st.success(f"A demanda ajustada cobre {pct(ia['cobertura'], 0)} dos clientes do plano.")
                else:
                    st.warning(f"A demanda ajustada cobre só {pct(ia['cobertura'], 0)} dos clientes do plano: a meta de vendas parece otimista.")
        st.caption("Fatores sugeridos (70% e 30%) são uma regra prática de pesquisa de mercado; ajuste conforme o setor e o preço.")
    proxima("teste_preco")


# ------------------------------------------------------------------------------
# CANVAS DA PROPOSTA DE VALOR
# ------------------------------------------------------------------------------
def pg_vpc():
    cabecalho("vpc")
    st.caption("Lado direito: o **perfil do cliente**. Lado esquerdo: o **mapa de valor** da empresa (Osterwalder et al., 2014). "
               "Escreva um item por linha. O encaixe acontece quando cada dor e ganho importante tem uma resposta da oferta.")
    job = st.text_input("Que trabalho (job) o cliente contrata o seu produto para fazer?", proj["vpc_job"], key=wk("vpc_job"),
                        placeholder="Quando estou sem tempo no almoço, quero comer algo saudável, para manter a dieta sem perder o horário.",
                        help="Formato de Christensen et al. (2016): quando [situação], quero [motivação], para [resultado].")
    v = {"vpc_job": job}
    ce, cd = st.columns(2)
    with cd:
        st.markdown("##### 🙂 Perfil do cliente")
        for k_ in ("vpc_tarefas", "vpc_dores", "vpc_ganhos"):
            _, rot, ajuda = next(b for b in BLOCOS_VPC if b[0] == k_)
            v[k_] = st.text_area(rot, proj[k_], height=130, key=wk(k_), help=ajuda)
    with ce:
        st.markdown("##### 🎁 Mapa de valor")
        for k_ in ("vpc_produtos", "vpc_aliviadores", "vpc_criadores"):
            _, rot, ajuda = next(b for b in BLOCOS_VPC if b[0] == k_)
            v[k_] = st.text_area(rot, proj[k_], height=130, key=wk(k_), help=ajuda)
    auto_unica("projeto", v)
    salvo_txt()
    enc = encaixe_vpc(proj, d)
    if enc["total"] == 0:
        return proxima("vpc")
    st.subheader("Encaixe: como a oferta responde a cada dor e ganho")
    opcoes = [""] + linhas_txt(proj["vpc_aliviadores"]) + linhas_txt(proj["vpc_criadores"])

    def _salvar(x):
        salvar(conn, "vpc_encaixe", x, "item")
        d["vpc_encaixe"] = limpar(x.reset_index(drop=True), "vpc_encaixe")

    tabela_editavel(enc["tabela"][["item", "tipo", "resposta"]], {
        "item": T("Dor ou ganho do cliente", disabled=True, width="large"), "tipo": T("Tipo", disabled=True, width="small"),
        "resposta": S("Respondido por", options=opcoes, width="large")}, "vpc_encaixe",
        "|".join(enc["tabela"]["item"])[:200], normalizar=lambda x: limpar(x, "vpc_encaixe"), salvar_fn=_salvar, adicionar=False)
    enc = encaixe_vpc(proj, d)
    st.metric("Encaixe produto-cliente", pct(enc["encaixe"], 0), f"{enc['cobertos']} de {enc['total']} itens respondidos",
              delta_color="off")
    if enc["encaixe"] < 0.6:
        st.warning("Muitas dores e ganhos ainda não têm resposta na oferta. Ou a oferta precisa mudar, ou esses itens não são prioridade.")
    proxima("vpc")


# ------------------------------------------------------------------------------
# CURVA DE VALOR E ERRC (OCEANO AZUL)
# ------------------------------------------------------------------------------
def pg_oceano():
    cabecalho("oceano")
    st.caption("A **curva de valor** mostra como cada empresa se sai nos atributos que o setor oferece (notas da Análise da "
               "concorrência). A matriz **ERRC** propõe uma curva nova: o que **E**liminar, **R**eduzir, **E**levar e **C**riar "
               "para se diferenciar sem aumentar custos (Kim; Mauborgne, 2005).")
    atributos = list(dict.fromkeys(d["concorrencia_notas"]["atributo"].astype(str)))
    if not atributos:
        st.warning("Primeiro dê notas aos atributos na Análise da concorrência.")
        _ir_botao("Ir para Análise da concorrência →", "concorrencia", "oc_conc")
        return proxima("oceano")
    editor("errc", {"atributo": S("Atributo (escolha ou digite um novo para Criar)", options=atributos, required=True, width="medium"),
                    "acao": S("Ação", options=ACOES_ERRC, required=True),
                    "nova_nota": N("Nova nota (0 a 10)", min_value=0.0, max_value=10.0, step=0.5),
                    "justificativa": T("Por quê?", width="large")}, "atributo")
    st.caption("Para **criar** um atributo novo, que nenhum concorrente oferece, escreva-o primeiro na Análise da concorrência "
               "com nota 0 para todos e depois marque aqui a ação Criar com a nota pretendida.")
    cv = curva_valor(d)
    ordem = list(dict.fromkeys(cv["Atributo"]))
    st.altair_chart(alt.Chart(cv).mark_line(point=True).encode(
        x=alt.X("Atributo:N", sort=ordem, title=None, axis=alt.Axis(labelAngle=-30)), y=alt.Y("Nota:Q", scale=alt.Scale(domain=[0, 10])),
        color=alt.Color("Empresa:N", legend=alt.Legend(orient="bottom", title=None)),
        strokeDash=alt.condition(alt.datum.Empresa == "Nova curva (ERRC)", alt.value([1, 0]), alt.value([4, 3])),
        strokeWidth=alt.condition(alt.datum.Empresa == "Nova curva (ERRC)", alt.value(4), alt.value(1.6)),
        tooltip=["Empresa", "Atributo", "Nota"]).properties(height=380), width="stretch")
    errc = d["errc"]
    if not errc.empty:
        cols = st.columns(4)
        for c_, (ac, ic) in zip(cols, [("Eliminar", "✂️"), ("Reduzir", "⬇️"), ("Elevar", "⬆️"), ("Criar", "✨")]):
            with c_.container(border=True):
                st.markdown(f"**{ic} {ac}**")
                for a in errc.loc[errc["acao"] == ac, "atributo"]:
                    st.markdown(f"- {a}")
        st.info(leitura_errc(d))
    proxima("oceano")


# ------------------------------------------------------------------------------
# POSICIONAMENTO E MAPA PERCEPTUAL
# ------------------------------------------------------------------------------
def pg_posicionamento():
    cabecalho("posicionamento")
    st.caption("A declaração de posicionamento diz, em uma frase, **para quem** é o negócio, **o que** entrega e **por que** é "
               "diferente (modelo de Moore, 1991). O mapa perceptual mostra onde a marca fica na cabeça do cliente.")
    c = st.columns(2)
    v = {}
    v["posic_segmento"] = c[0].text_input("Para (público-alvo)", proj["posic_segmento"], key=wk("ps_seg"))
    v["posic_necessidade"] = c[1].text_input("que (necessidade ou problema)", proj["posic_necessidade"], key=wk("ps_nec"))
    v["posic_marca"] = c[0].text_input("a marca", proj["posic_marca"] or proj["nome_negocio"], key=wk("ps_marca"))
    v["posic_categoria"] = c[1].text_input("é (categoria)", proj["posic_categoria"], key=wk("ps_cat"))
    v["posic_beneficio"] = c[0].text_input("que (benefício principal)", proj["posic_beneficio"], key=wk("ps_ben"))
    v["posic_concorrente"] = c[1].text_input("Diferentemente de (principal alternativa)", proj["posic_concorrente"], key=wk("ps_conc"))
    v["posic_diferencial"] = st.text_input("o nosso produto (diferencial que prova o benefício)", proj["posic_diferencial"], key=wk("ps_dif"))
    auto_unica("projeto", v)
    frase = frase_posicionamento(proj)
    if frase:
        st.success("**Declaração de posicionamento:** " + frase)
    st.subheader("Mapa perceptual")
    c = st.columns(2)
    ex = c[0].text_input("Eixo horizontal", proj["mapa_eixo_x"], key=wk("mp_x"))
    ey = c[1].text_input("Eixo vertical", proj["mapa_eixo_y"], key=wk("mp_y"))
    auto_unica("projeto", {"mapa_eixo_x": ex, "mapa_eixo_y": ey})
    if d["mapa_perceptual"].empty:
        nomes = [proj["nome_negocio"] or "Nossa Empresa"] + [n for n in d["concorrentes"]["nome"] if str(n).strip()]
        if st.button("Incluir a empresa e os concorrentes cadastrados"):
            _semear("mapa_perceptual", [{"marca": n, "eixo_x": 5.0, "eixo_y": 5.0} for n in nomes], "marca")
    editor("mapa_perceptual", {"marca": T("Marca", required=True), "eixo_x": N("Nota no eixo horizontal (0 a 10)", min_value=0.0,
                                                                                max_value=10.0, step=0.5),
                               "eixo_y": N("Nota no eixo vertical (0 a 10)", min_value=0.0, max_value=10.0, step=0.5)}, "marca")
    mp = d["mapa_perceptual"]
    if not mp.empty:
        nossa = proj["nome_negocio"] or "Nossa Empresa"
        g = mp.assign(Tipo=["Nossa empresa" if m in (nossa, "Nossa Empresa") else "Concorrente" for m in mp["marca"]])
        base = alt.Chart(g).encode(x=alt.X("eixo_x:Q", scale=alt.Scale(domain=[0, 10]), title=ex),
                                   y=alt.Y("eixo_y:Q", scale=alt.Scale(domain=[0, 10]), title=ey))
        pts = base.mark_circle(size=380, opacity=0.85).encode(
            color=alt.Color("Tipo:N", scale=alt.Scale(domain=["Nossa empresa", "Concorrente"], range=["#e67e22", "#2f6fd0"]),
                            legend=alt.Legend(orient="bottom", title=None)), tooltip=["marca", "eixo_x", "eixo_y"])
        txt = base.mark_text(dy=-18, fontSize=11).encode(text="marca:N")
        meio = alt.Chart(pd.DataFrame({"v": [5]})).mark_rule(color="#999").encode(x="v:Q") + \
            alt.Chart(pd.DataFrame({"v": [5]})).mark_rule(color="#999").encode(y="v:Q")
        st.altair_chart((meio + pts + txt).properties(height=420), width="stretch")
        st.caption("Procure um espaço vazio no mapa que seja valorizado pelo cliente. Se a sua marca está colada em um concorrente, "
                   "o cliente não tem motivo para trocar.")
    proxima("posicionamento")


# ------------------------------------------------------------------------------
# ORÇAMENTO DE MARKETING, FUNIL AARRR, CALENDÁRIO E NPS
# ------------------------------------------------------------------------------
def pg_orcamento_mkt():
    cabecalho("orcamento_mkt")
    res = R()
    t1, t2, t3, t4 = st.tabs(["💸 Orçamento e CAC por canal", "🪜 Funil AARRR", "🗓️ Calendário de ações", "⭐ Meta de NPS"])
    with t1:
        st.caption("Quanto será gasto por mês em cada canal e quantos clientes novos cada um deve trazer. O programa calcula o custo "
                   "de aquisição (CAC) e o retorno de cada canal.")
        canais = [c for c in d["canais"]["canal"].astype(str) if c.strip()]
        editor("orcamento_canais", {"canal": S("Canal", options=canais, required=True) if canais else T("Canal", required=True),
                                    "gasto_mensal": moeda("Gasto mensal"), "clientes_mes": N("Clientes novos por mês", min_value=0.0, step=1.0),
                                    "observacao": T("Como será medido", width="large")}, "canal")
        oc = orcamento_canais(proj, d, res)
        if not oc.empty:
            vis = oc.copy()
            for c_ in ("Gasto mensal", "CAC", "Valor do cliente"):
                vis[c_] = vis[c_].map(lambda x: brl(x) if x == x else "n.a.")
            vis["Retorno por R$ 1"] = vis["Retorno por R$ 1"].map(lambda x: f"R$ {num_br(x, 2)}" if x == x else "n.a.")
            vis = vis.rename(columns={"Valor do cliente": f"Valor do cliente ({oc.attrs.get('rotulo_valor', '')})"})
            st.dataframe(vis, hide_index=True, width="stretch")
            tot_g, tot_c = float(oc["Gasto mensal"].sum()), float(oc["Clientes novos/mês"].sum())
            k = st.columns(3)
            k[0].metric("Orçamento mensal de aquisição", brl(tot_g))
            k[1].metric("CAC médio", brl(div(tot_g, tot_c)) if tot_c else "n.a.")
            verba = float(proj["verba_aquisicao_mensal"] or 0)
            com = float(res["desp_com_m"] + res["pessoal_com_m"])
            k[2].metric("Gasto comercial no plano (mês)", brl(verba if verba > 0 else com),
                        help="Verba de aquisição informada em Assinaturas ou despesas comerciais fixas.")
            if tot_g > (verba if verba > 0 else com) * 1.05:
                st.warning("O orçamento por canal é maior que o gasto comercial previsto nos custos: inclua a diferença em Custos e "
                           "despesas, senão a DRE fica otimista.")
            pior = oc.sort_values("Retorno por R$ 1").iloc[0]
            if pior["Retorno por R$ 1"] == pior["Retorno por R$ 1"] and pior["Retorno por R$ 1"] < 1:
                st.error(f"O canal **{pior['Canal']}** devolve menos do que custa: cada R$ 1 investido traz R$ {num_br(pior['Retorno por R$ 1'], 2)} "
                         "em valor de cliente.")
    with t2:
        st.caption("Funil de crescimento em cinco etapas (Ellis; Brown, 2017): para cada uma, a métrica que mostra se está "
                   "funcionando, a meta e a ação.")
        if d["aarrr"].empty:
            if st.button("Criar as cinco etapas"):
                _semear("aarrr", [{"etapa": e, "metrica": "", "meta": "", "acao": ""} for e, _ in ETAPAS_AARRR], "etapa")
            for e, perg in ETAPAS_AARRR:
                st.markdown(f"- **{e}:** {perg}")
        editor("aarrr", {"etapa": S("Etapa", options=[e for e, _ in ETAPAS_AARRR], required=True),
                         "metrica": T("Métrica", width="medium"), "meta": T("Meta"), "acao": T("Ação principal", width="large")}, "etapa")
    with t3:
        editor("calendario_mkt", {"mes": N("Mês do plano", min_value=1, max_value=120, step=1), "acao": T("Ação", required=True, width="large"),
                                  "canal": T("Canal"), "custo": moeda("Custo")}, "acao")
        cal = d["calendario_mkt"]
        if not cal.empty:
            g = cal.groupby("mes", as_index=False)["custo"].sum()
            st.altair_chart(alt.Chart(g).mark_bar(color="#e67e22").encode(
                x=alt.X("mes:O", title="Mês"), y=alt.Y("custo:Q", title="Custo (R$)")).properties(height=240), width="stretch")
            st.caption(f"Total do calendário: {brl(cal['custo'].sum())}")
    with t4:
        st.markdown("O **Net Promoter Score** pergunta: *de 0 a 10, quanto você recomendaria a empresa a um amigo?* NPS = % de "
                    "promotores (9 e 10) menos % de detratores (0 a 6), variando de -100 a +100 (Reichheld, 2003).")
        nps = st.slider("Meta de NPS para o primeiro ano", -100, 100, int(proj["nps_meta"] or 50), 5, key=wk("nps"))
        auto_unica("projeto", {"nps_meta": nps})
        st.caption("Referência comum: acima de 50 é considerado muito bom e acima de 75, excelente. Ligue esta meta a um objetivo da "
                   "perspectiva de clientes no Balanced Scorecard.")
    proxima("orcamento_mkt")


# ------------------------------------------------------------------------------
# SIMULAÇÃO DE MONTE CARLO E PERDA ACEITÁVEL
# ------------------------------------------------------------------------------
def pg_montecarlo():
    cabecalho("montecarlo")
    res = R()
    if res["anos"][1]["rec_bruta"] <= 0:
        st.warning("Cadastre produtos ou assinaturas para simular.")
        return proxima("montecarlo")
    st.caption("Em vez de um único cenário, o programa sorteia 3.000 combinações de preço, volume, custos e crescimento dentro das "
               "faixas abaixo (distribuição triangular: mínimo, valor do plano, máximo) e mostra a **chance** de o projeto dar certo.")
    fx = faixas_mc(d)

    def _salvar(x):
        salvar(conn, "mc_faixas", x, "variavel")
        d["mc_faixas"] = limpar(x.reset_index(drop=True), "mc_faixas")

    tabela_editavel(fx, {"variavel": T("Variável", disabled=True, width="large"),
                         "minimo_pct": N("Pior caso (%)", step=1.0, max_value=0.0),
                         "maximo_pct": N("Melhor caso (%)", step=1.0, min_value=0.0)},
                    "mc_faixas", "faixas", normalizar=lambda x: limpar(x, "mc_faixas"), salvar_fn=_salvar, adicionar=False)
    st.caption("Para custos, o \"melhor caso\" positivo significa custo maior (é o limite de alta). Use pesquisas, orçamentos e o "
               "histórico do setor para definir as faixas.")
    mc = MC()
    k = st.columns(4)
    k[0].metric("Chance de VPL positivo", pct(mc["p_vpl_pos"], 0))
    k[1].metric("VPL mediano", brl(mc["p50"]))
    k[2].metric("VPL nos 5% piores casos", brl(mc["p5"]))
    k[3].metric("Chance de caixa negativo no Ano 1", pct(mc["p_caixa_neg"], 0))
    hist = pd.DataFrame({"VPL": mc["vpl"]})
    st.altair_chart(alt.Chart(hist).mark_bar(opacity=0.85).encode(
        x=alt.X("VPL:Q", bin=alt.Bin(maxbins=45), title="VPL (R$)"), y=alt.Y("count():Q", title="Cenários"),
        color=alt.condition(alt.datum.VPL > 0, alt.value("#27ae60"), alt.value("#c0392b"))).properties(height=300) +
        alt.Chart(pd.DataFrame({"x": [0]})).mark_rule(color="black").encode(x="x:Q"), width="stretch")
    st.info(leitura_mc(mc))
    st.subheader("Perda aceitável dos sócios")
    st.caption("Em vez de perguntar só \"quanto posso ganhar?\", o empreendedor experiente pergunta \"quanto aceito perder?\" e "
               "desenha o negócio para que o pior caso caiba nesse limite (Sarasvathy, 2001). Informe a perda aceitável de cada "
               "sócio em Informações e sócios.")
    rv = st.slider("Quanto do investimento fixo seria recuperado vendendo os bens, se o negócio fechar (%)", 0, 100,
                   int(proj["revenda_ativos_pct"] if proj["revenda_ativos_pct"] is not None else 40), 5, key=wk("revenda"))
    auto_unica("projeto", {"revenda_ativos_pct": rv})
    pa = perda_aceitavel(proj, d, res)
    k = st.columns(3)
    k[0].metric("Perda potencial se fechar", brl(pa["potencial"]),
                help="Capital dos sócios + empréstimos (normalmente garantidos pelos sócios) − venda dos bens.")
    k[1].metric("Perda aceitável somada", brl(pa["aceitavel"]))
    k[2].metric("Situação", ("cabe no limite" if pa["cabe"] else "acima do limite") if pa["ok"] else "não informada")
    if pa["ok"] and not pa["cabe"]:
        st.warning("O pior caso passa do que os sócios aceitam perder. Opções: começar menor (MVP), alugar em vez de comprar "
                   "equipamentos, buscar parceiros ou pré-vendas antes de investir.")
    _ir_botao("👥 Informar a perda aceitável dos sócios", "projeto", "mc_socios")
    proxima("montecarlo")


# ------------------------------------------------------------------------------
# CURVA DE ADOÇÃO (BASS)
# ------------------------------------------------------------------------------
def pg_adocao():
    cabecalho("adocao")
    res = R()
    st.caption("O modelo de Bass (1969) descreve como um produto novo se espalha: alguns clientes compram por iniciativa própria "
               "(**inovação**, p) e a maioria compra porque viu outros usando (**imitação**, q). A curva serve para testar se o "
               "crescimento previsto no plano é plausível.")
    c = st.columns(3)
    m0 = float(proj["bass_mercado"] or 0) or float(proj["sam_clientes"] or 0)
    m = c[0].number_input("Mercado potencial (clientes)", 0.0, 1e9, m0, 100.0, key=wk("bass_m"),
                          help="Padrão: SAM (clientes alcançáveis) de Tamanho do mercado.")
    p = c[1].number_input("Coeficiente de inovação (p)", 0.001, 0.5, float(proj["bass_p"] or 0.03), 0.005, format="%.3f", key=wk("bass_p"))
    q = c[2].number_input("Coeficiente de imitação (q)", 0.0, 1.5, float(proj["bass_q"] or 0.38), 0.02, format="%.2f", key=wk("bass_q"))
    auto_unica("projeto", {"bass_mercado": m if m != float(proj["sam_clientes"] or 0) else 0.0, "bass_p": p, "bass_q": q})
    st.caption("Valores médios de p = 0,03 e q = 0,38 vêm de uma meta-análise de estudos de difusão (Sultan; Farley; Lehmann, 1990). "
               "Produtos que dependem muito de indicação (aplicativos, serviços locais) costumam ter q maior.")
    bc = bass_comparacao(proj, res)
    if not bc["ok"]:
        st.info("Informe o mercado potencial ou preencha o SAM em Tamanho do mercado.")
        return proxima("adocao")
    tb = bc["tabela"]
    series = ["Adotantes acumulados"] + (["Clientes no plano"] if bc["tem_plano"] else [])
    g = tb.melt(id_vars="Ano", value_vars=series, var_name="Série", value_name="Clientes")
    st.altair_chart(alt.Chart(g).mark_line(point=True, strokeWidth=3).encode(
        x=alt.X("Ano:O"), y=alt.Y("Clientes:Q"), color=alt.Color("Série:N", legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["Ano", "Série", alt.Tooltip("Clientes:Q", format=",.0f")]).properties(height=320), width="stretch")
    vis = tb.copy()
    for c_ in ("Adotantes acumulados", "Novos adotantes no ano", "Clientes no plano"):
        if c_ in vis:
            vis[c_] = vis[c_].map(lambda x: num_br(x, 0))
    vis["% do mercado potencial"] = vis["% do mercado potencial"].map(lambda x: pct(x, 1))
    if "Plano ÷ Bass" in vis:
        vis["Plano ÷ Bass"] = vis["Plano ÷ Bass"].map(lambda x: f"{num_br(x, 2)}x")
    st.dataframe(vis, hide_index=True, width="stretch")
    if bc["tem_plano"]:
        acima = tb[tb["Plano ÷ Bass"] > 1.2]
        if acima.empty:
            st.success("O número de clientes do plano está dentro do que a difusão do produto permite.")
        else:
            st.warning(f"Nos anos {', '.join(str(a) for a in acima['Ano'])} o plano supõe mais clientes do que a curva de adoção "
                       "sugere. Ou o crescimento está otimista, ou o grupo precisa explicar o que acelera a adoção (parcerias, "
                       "indicação, mídia).")
    else:
        st.caption("Preencha o gasto anual por cliente em Tamanho do mercado para comparar com o plano.")
    proxima("adocao")


# ------------------------------------------------------------------------------
# RÉGUA DE PRONTIDÃO
# ------------------------------------------------------------------------------
def pg_prontidao():
    cabecalho("prontidao")
    res = R()
    pr = avaliar_prontidao(proj, d, res, MC(), preco_referencia(proj, res))
    st.metric("Nível de prontidão do negócio", f"{pr['nivel']} de 9")
    st.progress(pr["nivel"] / 9)
    st.caption("Inspirada nos níveis de prontidão usados por aceleradoras e programas de inovação: cada degrau só conta se os "
               "anteriores estiverem cumpridos.")
    for n_, nome, crit in NIVEIS_PRONTIDAO:
        ok = pr["cond"][n_]
        atingido = n_ <= pr["nivel"]
        ic = "✅" if atingido else ("🟡" if ok else "⬜")
        with st.container(border=True):
            st.markdown(f"{ic} **{n_}. {nome}**: {crit}")
            if n_ == pr["nivel"] + 1:
                st.caption("👉 Próximo degrau: é aqui que o grupo deve concentrar o esforço agora.")
            elif ok and not atingido:
                st.caption("Critério cumprido, mas um degrau anterior ainda está pendente.")
    proxima("prontidao")


# ------------------------------------------------------------------------------
# BANCA SIMULADA
# ------------------------------------------------------------------------------
def pg_banca():
    cabecalho("banca")
    res = R()
    perg = perguntas_banca(proj, d, res, MC(), preco_referencia(proj, res))
    st.caption("Perguntas que uma banca faria, geradas a partir dos pontos fracos do próprio plano. Escreva a resposta do grupo: "
               "o treino vale mais do que decorar o pitch. As respostas vão para o relatório e para o prompt de IA.")
    resp = {str(r["pergunta"]): str(r["resposta"]) for _, r in d["banca_respostas"].iterrows()}
    novas = {}
    for i, q in enumerate(perg):
        with st.container(border=True):
            st.markdown(f"**{i + 1}. [{q['tema']}]** {q['pergunta']}")
            st.caption(f"Por que a banca perguntaria: {q['motivo']}")
            novas[q["pergunta"]] = st.text_area("Resposta do grupo", resp.get(q["pergunta"], ""), key=wk(f"banca_{i}_{abs(hash(q['pergunta'])) % 10**6}"),
                                                label_visibility="collapsed", placeholder="Resposta com dados do plano...")
    mudou = any(novas[k] != resp.get(k, "") for k in novas)
    if mudou:
        todas = {**resp, **novas}
        df = pd.DataFrame({"pergunta": list(todas), "resposta": list(todas.values())})
        df = df[df["resposta"].astype(str).str.strip() != ""]
        salvar(conn, "banca_respostas", df, "pergunta")
        d["banca_respostas"] = limpar(df.reset_index(drop=True), "banca_respostas")
        marcar_salvo()
        invalidar()
    n_ok = sum(1 for k in novas if len(novas[k].strip()) > 20)
    st.metric("Perguntas respondidas", f"{n_ok} de {len(perg)}")
    proxima("banca")


# ------------------------------------------------------------------------------
# FORMALIZAÇÃO
# ------------------------------------------------------------------------------
def pg_formalizacao():
    cabecalho("formalizacao")
    res = R()
    ps = porte_sugerido(res, d)
    k = st.columns(2)
    k[0].metric("Receita prevista no Ano 1", brl(ps["receita"]))
    k[1].metric("Enquadramento sugerido", ps["porte"])
    for o in ps["obs"]:
        st.warning(o)
    st.caption(f"Limites usados: MEI até {brl(LIMITE_MEI)} por ano (com um sócio, no máximo um empregado e atividade permitida), "
               f"ME até {brl(LIMITE_ME)} e EPP até {brl(LIMITE_EPP)}. O teto do MEI está em discussão no Congresso (proposta aprovada "
               "na Câmara em 2026 prevê R$ 110 mil em 2027 e R$ 140 mil em 2028): confira a regra vigente no Portal do Empreendedor.")
    itens = itens_formalizacao(proj, d)
    faltam = [i for i in itens if i not in set(d["formalizacao"]["item"])]
    if faltam and st.button(f"Incluir {len(faltam)} item(ns) sugerido(s) para este negócio"):
        _semear("formalizacao", [{"item": i, "status": "Pendente", "observacao": ""} for i in faltam], "item")
    editor("formalizacao", {"item": T("Item", required=True, width="large"), "status": S("Situação", options=STATUS_FORMAL),
                            "observacao": T("Observação (custo, prazo, órgão)", width="large")}, "item")
    fo = d["formalizacao"]
    if not fo.empty:
        feito = float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).mean())
        st.progress(feito, text=f"{pct(feito, 0)} encaminhado")
    proxima("formalizacao")
