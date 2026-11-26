"""
Interface do mix de marketing integrado (v8.1 em diante).
Substitui a antiga página de quatro caixas de texto por um painel em que cada P
tem escolhas estruturadas ligadas aos números do plano. Os textos antigos
continuam valendo como justificativa de cada P.
"""
from pnbox_ui_base import *
from pnbox_mix import *
from pnbox_v9 import *

ICONE_NIVEL = {"ok": "✅", "atencao": "⚠️", "critico": "⛔"}


def mostrar_alertas(alertas):
    for nivel, msg in alertas:
        {"ok": st.success, "atencao": st.warning, "critico": st.error}[nivel](msg, icon=ICONE_NIVEL[nivel])


def _select(rotulo, campo, opcoes_, ajuda=None):
    atual = proj.get(campo) or ""
    lista = [""] + list(opcoes_)
    return st.selectbox(rotulo, lista, index=lista.index(atual) if atual in lista else 0,
                        format_func=lambda x: x or "(escolha)", key=wk(campo), help=ajuda)


def _aba_produto():
    st.caption("Kotler e Keller (2012) descrevem o produto em níveis: o benefício que o cliente compra, "
               "o produto em si e o que vem junto com ele.")
    v = {}
    c1, c2, c3 = st.columns(3)
    v["mix_nucleo"] = c1.text_area("🎯 Benefício central: o que o cliente realmente compra",
                                   proj.get("mix_nucleo") or "", height=120, key=wk("mx_n"),
                                   placeholder="Ex.: não é um bolo, é a comemoração resolvida sem trabalho.")
    v["mix_real"] = c2.text_area("📦 Produto em si: marca, qualidade, embalagem, linhas",
                                 proj.get("mix_real") or "", height=120, key=wk("mx_r"))
    v["mix_ampliado"] = c3.text_area("🎁 Produto ampliado: garantia, entrega, suporte, pós-venda",
                                     proj.get("mix_ampliado") or "", height=120, key=wk("mx_a"))
    v["mix_ciclo"] = _select("Fase do ciclo de vida do produto no mercado", "mix_ciclo", CICLO_VIDA)
    if v["mix_ciclo"]:
        st.info(CICLO_VIDA[v["mix_ciclo"]], icon="📈")
    v["mkt_produto"] = st.text_area("📝 Justificativa e detalhes do produto", proj["mkt_produto"], height=110,
                                    key=wk("m1"))
    auto_unica("projeto", v)
    cat = R()["cat"]
    if not cat.empty:
        st.markdown("**Portfólio (do catálogo de produtos)**")
        pt = pareto_produtos(cat)
        a, b, c = st.columns(3)
        a.metric("Itens no catálogo", num_br(pt["total"], 0))
        b.metric("Itens que fazem 80% da receita", num_br(pt["n"], 0))
        c.metric("Peso do item principal", pct(pt.get("share_top", float("nan"))))
        if pt["n"]:
            st.caption("Itens que sustentam a receita: " + ", ".join(pt["itens"]))
    dif = [proj.get("vpc_criadores"), proj.get("posic_diferencial")]
    if not any(str(x or "").strip() for x in dif):
        st.warning("Os diferenciais do produto ainda não aparecem no canvas da proposta de valor nem no "
                   "posicionamento. Um diferencial que só está no texto do 4Ps é difícil de defender na banca.",
                   icon="⚠️")


def _aba_preco():
    res = R()
    cat = res["cat"]
    c1, c2 = st.columns(2)
    with c1:
        est = _select("Estratégia de preço", "mix_estrategia_preco", ESTRATEGIAS_PRECO)
        if est:
            st.caption(ESTRATEGIAS_PRECO[est][0])
    with c2:
        pos = _select("Posição de preço em relação ao mercado", "mix_posicao_preco", POSICOES_PRECO)
    auto_unica("projeto", {"mix_estrategia_preco": est, "mix_posicao_preco": pos})

    st.markdown("#### Preços dos concorrentes")
    st.caption("Anote o preço que cada concorrente cobra pelo produto equivalente ao seu. "
               "Vale preço de vitrine, cardápio, site ou orçamento pedido.")
    prods = [x for x in cat["descricao"].astype(str) if x.strip()] if not cat.empty else []
    concs = [n for n in d["concorrentes"]["nome"].astype(str) if n.strip()]
    editor("precos_concorrentes", {
        "concorrente": S("Concorrente", options=concs, required=True) if concs else T("Concorrente", required=True),
        "produto": S("Produto equivalente ao seu", options=prods) if prods else T("Produto equivalente ao seu"),
        "preco": moeda("Preço cobrado"), "fonte": T("Onde viu o preço", width="medium")}, "concorrente")
    comp = comparar_concorrentes(cat, d["precos_concorrentes"])
    if not comp.empty:
        mostra = comp.copy()
        for col in ["Seu preço", "Média dos concorrentes", "Menor", "Maior"]:
            mostra[col] = mostra[col].map(brl)
        mostra["Índice"] = mostra["Índice"].map(lambda x: f"{x:.2f}".replace(".", ","))
        st.dataframe(mostra, hide_index=True, width="stretch")
        st.caption("Índice = seu preço ÷ média dos concorrentes. Acima de 1, você é mais caro.")

    st.markdown("#### Margem de cada item")
    mc = mc_produtos(res)
    if not mc.empty:
        t = mc.copy()
        t["Preço"], t["Custo direto"] = t["Preço"].map(brl), t["Custo direto"].map(brl)
        t["MC %"], t["Peso na receita"] = t["MC %"].map(pct), t["Peso na receita"].map(pct)
        st.dataframe(t, hide_index=True, width="stretch")
        st.caption("MC % = preço menos custo direto, impostos sobre a venda e outros custos variáveis, sobre o preço.")

    st.markdown("#### Desconto ou aumento: quanto o volume precisa mudar")
    mm = mc_media(res)
    if mm == mm and mm > 0:
        a, b = st.columns(2)
        dsc = a.slider("Desconto em estudo (%)", 0, 50, 10, key=wk("mx_desc")) / 100
        aum = b.slider("Aumento em estudo (%)", 0, 50, 10, key=wk("mx_aum")) / 100
        vol = volume_para_desconto(mm, dsc)
        queda = queda_tolerada_aumento(mm, aum)
        a.metric("Volume a mais para empatar", "impossível" if math.isinf(vol) else pct(vol))
        b.metric("Queda de volume tolerada", pct(queda))
        st.caption(f"Com margem de contribuição média de {pct(mm)}, um desconto de {dsc * 100:.0f}% só mantém "
                   f"o lucro se as vendas subirem {'mais do que é possível' if math.isinf(vol) else pct(vol)}. "
                   f"Um aumento de {aum * 100:.0f}% mantém o lucro mesmo que as vendas caiam até {pct(queda)}. "
                   "É uma conta de equilíbrio, não uma previsão de como os clientes vão reagir.")
    else:
        st.caption("Cadastre produtos com preço e custo para ver esta conta.")

    st.markdown("#### Condições de pagamento")
    st.caption("Como os clientes vão pagar. As taxas e prazos de exemplo são apenas referência: "
               "confira os da sua maquininha ou banco.")
    if d["meios_pagamento"].empty and st.button("Preencher com um exemplo", key=wk("mx_pag_ex")):
        recarregar_tabela("meios_pagamento", PAGAMENTO_EXEMPLO.copy(), "forma")
        st.rerun()
    editor("meios_pagamento", {"forma": T("Forma de pagamento", required=True),
                               "pct_vendas": N("% das vendas", min_value=0.0, max_value=100.0, format="%.1f"),
                               "taxa_pct": N("Taxa cobrada (%)", min_value=0.0, max_value=30.0, format="%.2f"),
                               "prazo_dias": N("Dias até o dinheiro entrar", min_value=0.0, format="%.0f")}, "forma")
    pg = d["meios_pagamento"]
    if not pg.empty:
        r = resumo_pagamento(pg)
        a, b, c = st.columns(3)
        a.metric("Soma das formas", f"{r['soma']:.0f}%".replace(".", ","))
        b.metric("Taxa média sobre a receita", f"{r['taxa']:.2f}%".replace(".", ","))
        c.metric("Prazo médio de recebimento", f"{r['prazo']:.0f} dias")
        if abs(r["soma"] - 100) > 0.5:
            st.warning("As formas de pagamento precisam somar 100% das vendas.", icon="⚠️")
        else:
            st.caption("Ao aplicar, a taxa média substitui a linha \"Taxas de cartões\" nos custos variáveis "
                       "e os prazos substituem os prazos de recebimento do capital de giro.")
            if st.button("✅ Aplicar ao plano financeiro", key=wk("mx_pag_ap")):
                cv, pz = aplicar_pagamento(pg, d["custos_variaveis"], d["capital_giro_prazos"])
                recarregar_tabela("custos_variaveis", cv, "descricao")
                recarregar_tabela("capital_giro_prazos", pz, "tipo")
                st.toast("Taxas e prazos aplicados ao plano.", icon="✅")
                st.rerun()

    v = {"mkt_preco": st.text_area("📝 Justificativa e detalhes do preço", proj["mkt_preco"], height=110, key=wk("m2"))}
    auto_unica("projeto", v)


def _aba_praca():
    dist = _select("Intensidade da distribuição", "mix_distribuicao", DISTRIBUICAO)
    if dist:
        st.caption(DISTRIBUICAO[dist])
    auto_unica("projeto", {"mix_distribuicao": dist})
    st.markdown("#### Canais de venda")
    st.caption("Onde o cliente compra, que parte das vendas passa por cada canal e quanto o canal cobra "
               "(comissão do aplicativo ou do marketplace, margem do revendedor). Confira as taxas atuais de cada canal.")
    editor("canais_venda", {"canal": S("Canal", options=opcoes("canal_venda", CANAIS_VENDA), required=True),
                            "pct_vendas": N("% das vendas", min_value=0.0, max_value=100.0, format="%.1f"),
                            "comissao_pct": N("Comissão ou taxa do canal (%)", min_value=0.0, max_value=60.0, format="%.1f"),
                            "observacao": T("Observação", width="medium")}, "canal")
    cv = d["canais_venda"]
    if not cv.empty:
        r = resumo_canais(cv)
        a, b, c = st.columns(3)
        a.metric("Soma dos canais", f"{num_br(r['soma'], 0)}%")
        b.metric("Comissão média sobre a receita", f"{num_br(r['comissao'], 2)}%")
        c.metric("Canal principal", r["maior"], f"{num_br(r['maior_pct'], 0)}% das vendas", delta_color="off")
        if abs(r["soma"] - 100) <= 0.5:
            st.caption(f"Ao aplicar, a comissão média entra nos custos variáveis como \"{_LINHA_COMISSAO}\" e a linha "
                       "\"Comissões de aplicativos de delivery\" é zerada, para a comissão não ser contada duas vezes.")
            if st.button("✅ Aplicar comissões ao plano financeiro", key=wk("mx_can_ap")):
                recarregar_tabela("custos_variaveis", aplicar_canais(cv, d["custos_variaveis"]), "descricao")
                st.toast("Comissões aplicadas.", icon="✅")
                st.rerun()
        else:
            st.warning("Os canais precisam somar 100% das vendas.", icon="⚠️")
    st.markdown("#### Localização")
    st.caption((f"Endereço informado: {proj['loc_endereco']}. " if str(proj.get("loc_endereco") or "").strip() else
                "A localização ainda não foi informada. ") + "Endereço, área e fluxo de pessoas ficam no plano operacional; "
               "o aluguel entra nos custos fixos.")
    if st.button("🏭 Abrir o plano operacional", key=wk("mx_op")):
        ir("operacional")
        st.rerun()
    auto_unica("projeto", {"mkt_praca": st.text_area("📝 Justificativa e detalhes da praça", proj["mkt_praca"], height=110,
                                                     key=wk("m3"))})


def _aba_promocao():
    res = R()
    obj = _select("Objetivo principal da comunicação nesta fase", "mix_objetivo_com", OBJETIVOS_COMUNICACAO)
    if obj:
        st.caption("Ações que costumam servir a esse objetivo: " + OBJETIVOS_COMUNICACAO[obj])
    st.markdown("#### O volume do plano e a promoção")
    st.caption("O volume de vendas do plano precisa de clientes. Aqui o programa estima quantos clientes ativos o "
               "primeiro mês exige e compara com os clientes novos que o orçamento de promoção promete trazer.")
    a, b = st.columns(2)
    tk = a.number_input("Gasto médio por compra (R$); deixe 0 para usar o preço médio do catálogo", min_value=0.0,
                        value=float(proj.get("mix_ticket") or 0), step=1.0, key=wk("mx_tk"))
    fr = b.number_input("Quantas vezes por mês o cliente típico compra", min_value=0.0,
                        value=float(proj.get("mix_freq") or 0), step=0.5, key=wk("mx_fr"))
    auto_unica("projeto", {"mix_objetivo_com": obj, "mix_ticket": tk, "mix_freq": fr})
    pv = promocao_vs_volume(proj, d, res)
    k = st.columns(4)
    k[0].metric("Clientes ativos que o plano exige", num_br(pv["cn"]["clientes"], 0) if pv["cn"].get("ok") else "n.a.")
    k[1].metric("Clientes novos por mês (promoção)", num_br(pv["novos"], 0) if pv["novos"] else "n.a.")
    k[2].metric("Custo por cliente novo (CAC)", brl(pv["cac"]) if pv["cac"] == pv["cac"] else "n.a.")
    k[3].metric("Verba de promoção ÷ receita", pct(pv["verba_receita"]) if pv["verba_receita"] == pv["verba_receita"] else "n.a.")
    if pv["novos"] <= 0:
        st.info("Os clientes novos por mês vêm do orçamento de marketing por canal.", icon="💡")
    if st.button("💸 Abrir orçamento de marketing e calendário", key=wk("mx_orc")):
        ir("orcamento_mkt")
        st.rerun()
    auto_unica("projeto", {"mkt_promocao": st.text_area("📝 Justificativa e detalhes da promoção", proj["mkt_promocao"],
                                                        height=110, key=wk("m4"))})


def _aba_servico():
    res = R()
    st.caption("Em serviços o cliente compra algo que não pode ver antes. Booms e Bitner (1981) acrescentam três Ps "
               "ao mix: as pessoas que atendem, o processo pelo qual o serviço acontece e a evidência física que o torna visível.")
    if not e_servico(proj, res):
        st.info("O projeto não está marcado como serviço, mas estes três Ps também ajudam no varejo e na alimentação, "
                "onde o atendimento faz parte do que se vende.", icon="💡")
    st.markdown("#### 👥 Pessoas")
    pes = d["pessoal"]
    n = float(pes["num_empregados"].sum()) if not pes.empty else 0.0
    st.caption(f"Equipe no plano: {num_br(n, 0)} pessoa(s), cadastradas em Custos e despesas.")
    v = {"mix_pessoas": st.text_area("Quem atende o cliente, que preparo precisa ter e como o atendimento será padronizado",
                                     proj.get("mix_pessoas") or "", height=100, key=wk("mx_pe"))}
    st.markdown("#### 🔁 Processos: mapa do serviço")
    st.caption("Liste as etapas do serviço na ordem em que acontecem, separando o que o cliente vê do que acontece nos "
               "bastidores (Shostack, 1984). Marque onde algo costuma dar errado.")
    editor("blueprint", {"etapa": T("Etapa", required=True, width="medium"),
                         "linha": S("Onde acontece", options=LINHAS_BLUEPRINT),
                         "tempo_min": N("Tempo (minutos)", min_value=0.0, step=1.0),
                         "ponto_falha": T("O que pode dar errado", width="medium")}, "etapa")
    bp = resumo_blueprint(d["blueprint"])
    if bp["etapas"]:
        a, b, c = st.columns(3)
        a.metric("Etapas", num_br(bp["etapas"], 0))
        b.metric("Tempo total", f"{num_br(bp['tempo'], 0)} min")
        c.metric("Pontos de falha", num_br(len(bp["falhas"]), 0))
        if bp["falhas"]:
            st.caption("Etapas com risco de falha: " + ", ".join(bp["falhas"]) +
                       ". Cada uma merece um padrão de atendimento ou uma checagem.")
    v["mix_processos"] = st.text_area("Observações sobre o processo", proj.get("mix_processos") or "", height=80,
                                      key=wk("mx_pr"))
    st.markdown("#### 🏠 Evidência física")
    v["mix_evidencia"] = st.text_area("O que o cliente vê e toca: ambiente, uniforme, embalagem, site, contrato, "
                                      "certificados, avaliações", proj.get("mix_evidencia") or "", height=100, key=wk("mx_ev"))
    auto_unica("projeto", v)


def _aba_diagnostico():
    vw = None
    try:
        vw = van_westendorp(d["precos_vw"])
    except Exception:
        pass
    res = R()
    st.markdown("#### Coerência entre os Ps")
    st.caption("Os Ps precisam contar a mesma história. As combinações abaixo aparecem conforme as escolhas feitas nas outras abas.")
    co = coerencia_mix(proj, d, res)
    if co.empty:
        st.info("Escolha a posição e a estratégia de preço, a distribuição e o objetivo da comunicação para ver a coerência do mix.",
                icon="💡")
    else:
        st.dataframe(co, hide_index=True, width="stretch")
        ok = int(co["Situação"].str.startswith("✅").sum())
        st.caption(f"{ok} de {len(co)} combinações coerentes.")
    st.markdown("#### Produto e preço")
    mostrar_alertas(checar_produto_preco(proj, d, res, vw))
    st.markdown("#### Praça e promoção")
    mostrar_alertas(checar_praca_promocao(proj, d, res))
    st.markdown("#### Autoavaliação pelos olhos do cliente (4Cs)")
    st.caption("Lauterborn (1990) propõe olhar cada P do ponto de vista de quem compra. Dê uma nota de 1 a 5 e escreva a evidência.")
    v, notas = {}, []
    for campo_n, campo_t, titulo, pergunta in QUATRO_CS:
        c1, c2 = st.columns([1, 3])
        v[campo_n] = c1.slider(titulo, 0, 5, int(proj.get(campo_n) or 0), key=wk(campo_n), help="0 = ainda não avaliado")
        v[campo_t] = c2.text_area(pergunta, proj.get(campo_t) or "", height=80, key=wk(campo_t))
        notas.append((titulo.split(" (")[0], v[campo_n]))
    auto_unica("projeto", v)
    avaliadas = [(t, n) for t, n in notas if n > 0]
    if len(avaliadas) == 4:
        pior = min(avaliadas, key=lambda x: x[1])
        st.altair_chart(alt.Chart(pd.DataFrame(avaliadas, columns=["C", "Nota"])).mark_bar(color="#2f6fd0").encode(
            x=alt.X("Nota:Q", scale=alt.Scale(domain=[0, 5])), y=alt.Y("C:N", sort=None, title=None)), width="stretch")
        st.caption(f"O ponto mais fraco pelo olhar do cliente é **{pior[0].lower()}** (nota {pior[1]}). "
                   "É por ele que vale começar a ajustar o mix.")


def pg_marketing():
    cabecalho("marketing")
    st.caption("Cada P tem escolhas que alimentam o restante do plano. Os campos de texto continuam "
               "valendo como justificativa.")
    abas = st.tabs(["📦 Produto", "🏷️ Preço", "📍 Praça", "📣 Promoção", "👥 Pessoas, processos e evidência física",
                    "🔎 Diagnóstico do mix"])
    with abas[0]:
        _aba_produto()
    with abas[1]:
        _aba_preco()
    with abas[2]:
        _aba_praca()
    with abas[3]:
        _aba_promocao()
    with abas[4]:
        _aba_servico()
    with abas[5]:
        _aba_diagnostico()
    proxima("marketing")


# ------------------------------------------------------------------------------
# PLANO DE MARKETING (v8.4)
# ------------------------------------------------------------------------------
SELO_ETAPA = {2: "✅", 1: "◐", 0: "○"}


def _botao_ir(rotulo, pagina, chave_):
    if st.button(rotulo, key=wk(chave_)):
        ir(pagina)
        st.rerun()


def pg_plano_mkt():
    cabecalho("plano_mkt")
    res = R()
    st.caption("O plano segue a sequência proposta na literatura: análise da situação, objetivos, estratégia, táticas, "
               "ações e controle (Smith, 2011; McDonald, 2007; Las Casas, 2011). Quase tudo vem das outras ferramentas; "
               "aqui o grupo amarra as partes e completa o que falta.")
    et = etapas_plano_mkt(proj, d)
    cols = st.columns(6)
    for c, (nome, s, falta) in zip(cols, et):
        c.metric(nome, SELO_ETAPA[s], help=f"Completa quando houver {falta}.")
    cont = conteudo_plano_mkt(proj, d, res)
    abas = st.tabs([f"{i + 1}. {n}" for i, (n, _) in enumerate(ETAPAS_PLANO)])
    with abas[0]:
        st.caption(ETAPAS_PLANO[0][1])
        st.markdown(cont[0]["texto"] if not _tem_txt(proj.get("pm_situacao")) else
                    cont[0]["texto"].replace(str(proj["pm_situacao"]), "").strip() or "")
        c = st.columns(4)
        for col, (rot, pg) in zip(c, [("🎯 Mercado", "mercado"), ("🥊 Concorrência", "concorrencia"), ("🧩 SWOT", "swot"),
                                      ("🔀 SWOT cruzada", "swot_cruzada")]):
            with col:
                _botao_ir(rot, pg, "pm_ir_" + pg)
        auto_unica("projeto", {"pm_situacao": st.text_area(
            "Síntese do diagnóstico: as três conclusões mais importantes para o marketing", proj.get("pm_situacao") or "",
            height=120, key=wk("pm_sit"))})
    with abas[1]:
        st.caption(ETAPAS_PLANO[1][1] + " Entram aqui os objetivos da perspectiva Clientes cadastrados em Missão, visão, valores e objetivos.")
        if cont[1]["tabela"] is not None:
            st.dataframe(cont[1]["tabela"], hide_index=True, width="stretch")
        else:
            st.info("Nenhum objetivo da perspectiva Clientes ainda. Exemplos: conquistar 300 clientes ativos até o mês 6; "
                    "chegar a 40% de recompra em 30 dias até o mês 9.", icon="💡")
        _botao_ir("🧭 Abrir os objetivos SMART", "identidade", "pm_ir_obj")
    with abas[2]:
        st.caption(ETAPAS_PLANO[2][1])
        fp = frase_posicionamento(proj)
        st.markdown("**Declaração de posicionamento:** " + (fp or "_ainda não preenchida_"))
        _botao_ir("📌 Abrir o posicionamento", "posicionamento", "pm_ir_pos")
        a, b = st.columns(2)
        with a:
            ans = _select("Direção de crescimento (matriz de Ansoff)", "pm_ansoff", ANSOFF,
                          "Para um negócio novo, pense no principal movimento dos próximos anos depois da abertura.")
            if ans:
                st.caption(f"**{ANSOFF[ans][0]}.** {ANSOFF[ans][1]}")
        with b:
            gen = _select("Estratégia competitiva (Porter)", "pm_generica", GENERICAS)
            if gen:
                st.caption(GENERICAS[gen])
        auto_unica("projeto", {"pm_ansoff": ans, "pm_generica": gen})
        if gen.startswith("Liderança") and proj.get("mix_posicao_preco") == "Premium":
            st.warning("Liderança em custo com preço premium: as duas escolhas puxam para lados opostos.", icon="⚠️")
        if gen and "diferenciação" in gen.lower() and proj.get("mix_estrategia_preco") == "Penetração":
            st.warning("Diferenciação com preço de penetração: se o diferencial existe, ele pode ser cobrado.", icon="⚠️")
    with abas[3]:
        st.caption(ETAPAS_PLANO[3][1])
        if cont[3]["tabela"] is not None:
            st.dataframe(cont[3]["tabela"], hide_index=True, width="stretch")
        co = coerencia_mix(proj, d, res)
        if not co.empty:
            ok = int(co["Situação"].str.startswith("✅").sum())
            (st.success if ok == len(co) else st.warning)(f"Coerência do mix: {ok} de {len(co)} combinações coerentes.")
        _botao_ir("📣 Abrir o mix de marketing", "marketing", "pm_ir_mix")
    with abas[4]:
        st.caption(ETAPAS_PLANO[4][1])
        if cont[4]["tabela"] is not None:
            st.markdown(cont[4]["texto"])
            st.dataframe(cont[4]["tabela"], hide_index=True, width="stretch")
        else:
            st.info("As ações ficam no calendário do orçamento de marketing.", icon="💡")
        _botao_ir("🗓️ Abrir o calendário de marketing", "orcamento_mkt", "pm_ir_cal")
    with abas[5]:
        st.caption(ETAPAS_PLANO[5][1] + " McDonald (2007) lembra que o plano só fecha o ciclo quando os resultados "
                   "voltam para a análise da situação seguinte.")
        if d["kpis_mkt"].empty and st.button("✨ Sugerir indicadores", key=wk("pm_kpi_sug")):
            recarregar_tabela("kpis_mkt", pd.DataFrame([{"indicador": i, "meta": "", "frequencia": f, "fonte": fo, "responsavel": ""}
                                                        for i, f, fo in KPIS_SUGERIDOS]), "indicador")
            st.rerun()
        editor("kpis_mkt", {"indicador": T("Indicador", required=True, width="medium"), "meta": T("Meta"),
                            "frequencia": S("Frequência", options=FREQUENCIAS), "fonte": T("De onde vem o dado", width="medium"),
                            "responsavel": T("Responsável")}, "indicador")
        sem_meta = [i for i, m in zip(d["kpis_mkt"]["indicador"], d["kpis_mkt"]["meta"]) if not str(m).strip()]
        if sem_meta:
            st.warning("Indicadores sem meta não servem para controle: " + ", ".join(sem_meta) + ".", icon="⚠️")
        _botao_ir("📊 Abrir o acompanhamento do realizado", "realizado", "pm_ir_real")
    with st.expander("📄 Ver o plano de marketing montado"):
        for i, sec in enumerate(cont, 1):
            st.markdown(f"**{i}. {sec['titulo']}**")
            st.markdown(sec["texto"])
            if sec["tabela"] is not None:
                st.dataframe(sec["tabela"], hide_index=True, width="stretch")
    proxima("plano_mkt")


# ------------------------------------------------------------------------------
# SAZONALIDADE (v8.6)
# ------------------------------------------------------------------------------
def pg_sazonalidade():
    cabecalho("sazonalidade")
    st.caption("Poucos negócios vendem igual o ano todo. Aqui o grupo informa o peso de cada mês (100 = mês médio). "
               "O total do ano continua o mesmo; muda a distribuição dos meses e, com ela, o fluxo de caixa.")
    a, b = st.columns(2)
    ini = a.selectbox("Mês do calendário em que o negócio abre", list(range(1, 13)),
                      index=int(proj.get("saz_mes_inicio") or 1) - 1, format_func=lambda m: MESES_NOME[m - 1], key=wk("saz_ini"))
    perfil = b.selectbox("Perfil de partida (ilustrativo)", list(PERFIS_SAZONAIS), key=wk("saz_perfil"),
                         index=list(PERFIS_SAZONAIS).index(proj["saz_perfil"]) if proj.get("saz_perfil") in PERFIS_SAZONAIS else 0)
    auto_unica("projeto", {"saz_mes_inicio": float(ini)})
    if st.button("Usar este perfil", key=wk("saz_usar")):
        auto_unica("projeto", {"saz_perfil": perfil})
        recarregar_tabela("sazonalidade", tabela_perfil(perfil), "mes")
        st.rerun()
    st.caption("Os perfis são ilustrativos. Ajuste os índices com o que o grupo sabe do negócio: vendas de concorrentes, "
               "conversa com fornecedores, datas comemorativas e o clima da cidade.")
    editor("sazonalidade", {"mes": N("Mês do calendário", min_value=1, max_value=12, step=1, required=True),
                            "indice": N("Índice (100 = mês médio)", min_value=0.0, step=5.0)}, "mes")
    res = R()
    mm = res["meses"]
    graf = pd.DataFrame({"Mês do plano": [f"{i + 1:02d}" for i in range(12)], "Receita": mm["Receita"],
                         "Resultado operacional": mm["Resultado operacional"]})
    st.altair_chart(alt.Chart(graf.melt("Mês do plano", var_name="Série", value_name="R$")).mark_line(point=True).encode(
        x="Mês do plano:N", y="R$:Q", color="Série:N"), width="stretch")
    mostrar_leitura("sazonalidade_resultado", leitura_sazonalidade(proj, res))
    if st.button("💵 Ver o fluxo de caixa", key=wk("saz_fluxo")):
        ir("fluxo")
        st.rerun()
    proxima("sazonalidade")


# ------------------------------------------------------------------------------
# SENSIBILIDADE (v8.7)
# ------------------------------------------------------------------------------
def pg_tornado():
    cabecalho("tornado")
    res = R()
    if res["anos"][1]["rec_bruta"] <= 0:
        st.info("Cadastre produtos ou assinaturas para analisar a sensibilidade.", icon="💡")
        return
    st.caption("O gráfico de tornado mostra quanto o VPL muda quando cada variável vai do pior ao melhor valor da faixa "
               "informada no Monte Carlo, uma de cada vez. As barras mais compridas são as premissas que mais merecem pesquisa.")
    fx = faixas_mc(d)
    tor = tornado(res, fx)
    base = float(tor["VPL base"].iloc[0])
    g = []
    for _, r in tor.iterrows():
        g.append({"Variável": r["Variável"], "Caso": "Pior valor da faixa" if r["VPL no mínimo"] <= r["VPL no máximo"] else "Melhor valor da faixa",
                  "ini": base, "fim": min(r["VPL no mínimo"], r["VPL no máximo"])})
        g.append({"Variável": r["Variável"], "Caso": "Melhor valor da faixa" if r["VPL no mínimo"] <= r["VPL no máximo"] else "Pior valor da faixa",
                  "ini": base, "fim": max(r["VPL no mínimo"], r["VPL no máximo"])})
    gd = pd.DataFrame(g)
    ordem = tor["Variável"].tolist()
    barras = alt.Chart(gd).mark_bar().encode(
        y=alt.Y("Variável:N", sort=ordem, title=None), x=alt.X("ini:Q", title="VPL (R$)"), x2="fim:Q",
        color=alt.Color("Caso:N", scale=alt.Scale(domain=["Pior valor da faixa", "Melhor valor da faixa"], range=["#d62828", "#2f6fd0"])))
    linha = alt.Chart(pd.DataFrame({"x": [base]})).mark_rule(strokeDash=[4, 3], color="#333").encode(x="x:Q")
    st.altair_chart(barras + linha, width="stretch")
    st.caption(f"Linha tracejada: VPL do plano ({brl(base)}). As faixas de variação são editadas em Monte Carlo.")
    vis = tor[["Variável", "Variação mínima", "Variação máxima", "VPL no mínimo", "VPL no máximo"]].copy()
    vis["Variação mínima"], vis["Variação máxima"] = vis["Variação mínima"].map(pct), vis["Variação máxima"].map(pct)
    vis["VPL no mínimo"], vis["VPL no máximo"] = vis["VPL no mínimo"].map(brl), vis["VPL no máximo"].map(brl)
    st.dataframe(vis, hide_index=True, width="stretch")
    st.markdown("#### Ponto de ruptura")
    st.caption("Quanto cada variável pode mudar, sozinha, até o VPL chegar a zero.")
    rup = tabela_rupturas(res)
    st.dataframe(rup.assign(**{"Variação que zera o VPL": rup["Variação que zera o VPL"].map(
        lambda x: pct(x) if x == x else "não zera entre -100% e +500%")}), hide_index=True, width="stretch")
    mostrar_leitura("tornado", leitura_tornado(tor, rup))
    st.markdown("#### Valide isto primeiro")
    try:
        mostrar_alertas(valide_primeiro(tor, placar_evidencias(proj, d, res)["tabela"]))
    except Exception:
        mostrar_alertas(valide_primeiro(tor, None))
    if st.button("🔬 Abrir o placar de evidências", key=wk("tor_ev")):
        ir("evidencias")
        st.rerun()
    st.markdown("#### Preço × volume")
    st.caption("VPL para combinações de variação de preço e de volume. Em vermelho, as combinações em que o projeto não se paga.")
    mp = mapa_preco_volume(res)
    mp["Preço"], mp["Volume"] = mp["Preço"].map(lambda x: f"{x * 100:+.0f}%"), mp["Volume"].map(lambda x: f"{x * 100:+.0f}%")
    ordem_ = [f"{x:+.0f}%" for x in range(-20, 25, 5)]
    st.altair_chart(alt.Chart(mp).mark_rect().encode(
        x=alt.X("Preço:N", sort=ordem_, title="Variação do preço"), y=alt.Y("Volume:N", sort=ordem_[::-1], title="Variação do volume"),
        color=alt.Color("VPL:Q", scale=alt.Scale(scheme="redblue", domainMid=0)), tooltip=["Preço", "Volume", alt.Tooltip("VPL:Q", format=",.0f")]),
        width="stretch")
    proxima("tornado")
