# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: SERVIÇOS (CUSTO DA HORA) E ASSINATURAS (RECEITA RECORRENTE)
# ==============================================================================


def _info(html: str):
    st.markdown(f'<div class="pn-card-info">{html}</div>', unsafe_allow_html=True)
    st.write("")


# ------------------------------------------------------------------------------
# SERVIÇOS E CUSTO DA HORA
# ------------------------------------------------------------------------------
def pg_servicos():
    cabecalho("servicos")
    _info("Em serviços, o principal custo é o <b>tempo da equipe</b>. O programa calcula quanto custa cada hora "
          "efetivamente cobrável e, a partir dela, o custo pleno e o preço sugerido de cada serviço. Cadastre a equipe da "
          "operação em <b>🧾 Custos e despesas</b> (área \"Produção / operação\").")
    t1, t2, t3 = st.tabs(["⏱️ Horas da equipe", "🧾 Custo e preço dos serviços", "📚 Como funciona"])
    v = {}
    with t1:
        c1, c2 = st.columns(2)
        v["horas_mes_pessoa"] = c1.number_input("Horas contratadas por pessoa no mês", 0.0, 400.0, float(proj["horas_mes_pessoa"] or 168),
                                                4.0, key=wk("hmes"),
                                                help="44 h semanais ≈ 220 h/mês com o descanso remunerado; 40 h ≈ 176 h; 6 h/dia em 20 dias = 120 h.")
        v["utilizacao_pct"] = c2.slider("Taxa de utilização (% das horas que vira trabalho cobrável)", 10, 100,
                                        int(proj["utilizacao_pct"] or 75), 5, key=wk("util"),
                                        help="Desconte reuniões, deslocamentos, orçamentos, treinamento, retrabalho e férias. "
                                             "Empresas de serviços costumam ficar entre 60% e 80%.")
        auto_unica("projeto", v)
        res = R()
        ch = custo_hora(res, d, proj)
        k = st.columns(4)
        k[0].metric("Pessoas na operação", num_br(ch["pessoas"], 0))
        k[1].metric("Horas contratadas/mês", num_br(ch["horas_contratadas"], 0))
        k[2].metric("Horas cobráveis/mês", num_br(ch["horas_cobraveis"], 0))
        k[3].metric("Custo da hora produtiva", brl(ch["custo_hora"]) if ch["custo_hora"] == ch["custo_hora"] else "n.a.",
                    help=f"(Mão de obra da operação {brl(res['mod_m'])} + custos indiretos e de tecnologia {brl(res['cif_m'])} "
                         f"+ depreciação {brl(res['dep_m'])}) ÷ {num_br(ch['horas_cobraveis'], 0)} horas cobráveis")
        if ch["pessoas"] == 0:
            st.warning("Não há ninguém na área \"Produção / operação\". Sem equipe não há horas para vender.")
            if st.button("Ir para Custos e despesas →"):
                ir("custos", "Cadastre na aba Mão de obra quem executa os serviços, com a área \"Produção / operação\".")
                st.rerun()
        if ch["horas_cobraveis"] > 0:
            st.markdown("**Ocupação da equipe (horas vendidas ÷ horas cobráveis)**")
            o1, o12 = ch["ocupacao_m1"], ch["ocupacao_m12"]
            st.progress(min(1.0, o1 if o1 == o1 else 0.0), text=f"Mês 1: {num_br(ch['horas_demandadas_m1'], 0)} h vendidas = {pct(o1)}")
            st.progress(min(1.0, o12 if o12 == o12 else 0.0), text=f"Mês 12: {num_br(ch['horas_demandadas_m12'], 0)} h vendidas = {pct(o12)}")
            if o12 == o12 and o12 > 1:
                st.error("No Mês 12 a equipe não dá conta das horas vendidas. Planeje contratações ou reduza a meta.")
            elif o12 == o12 and o12 < 0.5:
                st.info("Há muita hora ociosa: dá para vender mais serviços sem aumentar a equipe.")
    with t2:
        st.caption("Serviços cadastrados (os mesmos de Produtos e materiais, categoria Serviço). **Insumos** são materiais "
                   "gastos em cada execução; o tempo da equipe entra pelas **horas por serviço**.")
        editor("catalogo", {
            "descricao": T("Serviço", required=True, width="large"),
            "preco_venda": moeda("Preço cobrado (R$)"),
            "custo_aquisicao": moeda("Insumos por serviço (R$)"),
            "qtd_mensal": N("Qtd. no Mês 1", min_value=0, step=1),
            "horas_unit": N("Horas por serviço", min_value=0.0, step=0.5),
            "qtd_estoque_inicial": None}, "descricao", filtro={"categoria": "Serviço"})
        res = R()
        ch = custo_hora(res, d, proj)
        tab = ch["tabela"]
        if tab.empty:
            st.info("Informe as horas por serviço para ver o custo pleno e o preço sugerido.")
        else:
            st.markdown(f"**Custo da hora:** {brl(ch['custo_hora'])} · **Despesas de estrutura:** {pct(ch['desp_pct'])} da receita "
                        f"(média do Ano 1) · **Margem desejada:** {pct(res['margem_desejada'])}")
            vis = tab.copy()
            for c in ("Insumos", "Mão de obra (horas x custo hora)", "Custo pleno", "Preço atual", "Preço sugerido",
                      "Valor da hora cobrada"):
                vis[c] = vis[c].map(lambda x: brl(x) if x == x else "não calculável")
            vis["Margem após custos diretos"] = tab["Margem após custos diretos"].map(pct)
            for c in ("Horas por serviço", "Qtd/mês", "Horas/mês"):
                vis[c] = vis[c].map(lambda x: num_br(x, 1))
            st.dataframe(vis, hide_index=True, width="stretch")
            if ch["divisor"] <= 0.05:
                st.warning("As despesas de estrutura, os impostos e a margem desejada somam quase 100% da receita: nenhum preço "
                           "cobre tudo isso. Reduza a estrutura, revise a margem ou aumente o volume de vendas.")
            baixos = tab[(tab["Preço sugerido"] == tab["Preço sugerido"]) & (tab["Preço atual"] < tab["Preço sugerido"])]
            if not baixos.empty:
                st.warning("Abaixo do preço sugerido: " + ", ".join(baixos["Serviço"]) + ".")
            g = tab.melt("Serviço", ["Custo pleno", "Preço atual", "Preço sugerido"], "Valor", "R$").dropna()
            st.altair_chart(alt.Chart(g).mark_bar().encode(
                x=alt.X("Serviço:N", title=None, axis=alt.Axis(labelAngle=0, labelLimit=180)), xOffset="Valor:N",
                y=alt.Y("R$:Q", title="R$"),
                color=alt.Color("Valor:N", scale=alt.Scale(domain=["Custo pleno", "Preço atual", "Preço sugerido"],
                                                           range=["#c0392b", "#2f6fd0", "#18a058"]),
                                legend=alt.Legend(orient="bottom", title=None)),
                tooltip=["Serviço", "Valor", alt.Tooltip("R$:Q", format=",.2f")]).properties(height=320), width="stretch")
    with t3:
        st.markdown(r"""
#### Passo a passo do preço de um serviço
1. **Horas cobráveis** = pessoas da operação × horas contratadas no mês × taxa de utilização.
   Ninguém trabalha 100% do tempo em atividades que o cliente paga: há deslocamento, orçamento, reunião interna, estudo.
2. **Custo da hora produtiva** = (salários, encargos e benefícios da operação + custos indiretos e de tecnologia +
   depreciação dos equipamentos) ÷ horas cobráveis.
3. **Custo pleno do serviço** = insumos + horas do serviço × custo da hora.
4. **Preço sugerido** = custo pleno ÷ [1 − (impostos + despesas variáveis + despesas de estrutura + margem desejada)],
   o mesmo *markup* divisor usado para produtos.

#### Leituras úteis
- **Valor da hora cobrada** (preço ÷ horas) mostra se o serviço remunera o tempo da equipe.
- **Ocupação** acima de 100% indica que faltam pessoas; abaixo de 50% indica capacidade ociosa.
- Para profissional autônomo, o mesmo raciocínio vale: some os custos mensais e o pró-labore desejado, divida pelas horas
  produtivas e acrescente os impostos.
- Serviços vendidos por pacote (mensalidade de manutenção, plano de suporte) também podem ser tratados como
  **assinatura**, em 💳 Assinaturas e receita recorrente.
""")
    proxima("servicos")


# ------------------------------------------------------------------------------
# ASSINATURAS E RECEITA RECORRENTE
# ------------------------------------------------------------------------------
CICLOS = {1: "Mensal", 3: "Trimestral", 6: "Semestral", 12: "Anual"}


def pg_assinaturas():
    cabecalho("assinaturas")
    _info("Em negócios por assinatura o cliente não compra uma vez: ele <b>fica</b> e paga de novo a cada ciclo, até "
          "cancelar. Por isso o programa projeta <b>safras de assinantes</b> mês a mês, separando o dinheiro que entra "
          "(<b>cobranças</b>) da receita que pertence a cada mês (<b>receita reconhecida</b>). Um plano anual pago adiantado "
          "entra todo no caixa, mas vira receita 1/12 por mês.")
    t1, t2, t3, t4, t5 = st.tabs(["📋 Planos", "📈 Projeção 36 meses", "📊 MRR, churn, CAC e LTV", "☁️ Custos de tecnologia",
                                  "📚 Como funciona"])
    with t1:
        st.caption("Um plano por linha. **Preço por ciclo** é o valor cobrado de uma vez (ex.: R$ 3.000 no plano anual). "
                   "**Churn** é a % de assinantes que cancelam por mês. **Custo de entrega** é o gasto mensal por assinante "
                   "ativo (nuvem, APIs, SMS/e-mail, suporte). **Taxa de cobrança** é o % do meio de pagamento ou da loja de "
                   "aplicativos.")
        editor("assinaturas", {
            "plano": T("Plano", required=True, width="medium"),
            "publico": S("Público", options=["Empresas (B2B)", "Consumidores (B2C)"]),
            "preco_ciclo": moeda("Preço por ciclo (R$)"),
            "ciclo_meses": N("Ciclo (meses)", min_value=1, max_value=36, step=1,
                             help="1 = mensal, 3 = trimestral, 6 = semestral, 12 = anual"),
            "novos_mes1": N("Novos assinantes no Mês 1", min_value=0.0, step=1.0),
            "crescimento_novos_pct": N("Crescimento dos novos (% ao mês)", step=0.5,
                                       help="Quanto a captação de novos assinantes cresce por mês no Ano 1."),
            "churn_mensal_pct": N("Churn (% ao mês)", min_value=0.0, max_value=100.0, step=0.5),
            "custo_entrega_mensal": moeda("Custo de entrega por assinante/mês (R$)"),
            "taxa_cobranca_pct": N("Taxa de cobrança (%)", min_value=0.0, max_value=50.0, step=0.5)}, "plano")
        ass = d["assinaturas"]
        if not ass.empty:
            resumo = pd.DataFrame({
                "Plano": ass["plano"], "Ciclo": ass["ciclo_meses"].map(lambda c: CICLOS.get(int(c), f"{int(c)} meses")),
                "Equivale por mês": [brl(div(p, c)) for p, c in zip(ass["preco_ciclo"], ass["ciclo_meses"].clip(lower=1))],
                "Vida média do assinante": [f"{num_br(div(1, c / 100), 0)} meses" if c > 0 else "infinita (churn zero!)"
                                            for c in ass["churn_mensal_pct"]],
                "Lucro bruto por assinante/mês": [brl(div(p, c) * (1 - t / 100) - e) for p, c, t, e in
                                                  zip(ass["preco_ciclo"], ass["ciclo_meses"].clip(lower=1),
                                                      ass["taxa_cobranca_pct"], ass["custo_entrega_mensal"])]})
            st.dataframe(resumo, hide_index=True, width="stretch")
            if (ass["churn_mensal_pct"] <= 0).any():
                st.warning("Plano com churn zero: na prática, todo negócio por assinatura perde clientes. Informe uma estimativa.")
    res = R()
    s, planos, rc = res["ass"], res["ass_planos"], res["recorrencia"]
    with t2:
        if s.empty:
            st.info("Cadastre ao menos um plano na aba 📋 Planos.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Assinantes ativos por plano**")
                st.altair_chart(alt.Chart(planos).mark_area(opacity=0.85).encode(
                    x=alt.X("Mês:Q", scale=alt.Scale(domain=[1, 36])), y=alt.Y("Ativos:Q", stack=True, title="Assinantes"),
                    color=alt.Color("Plano:N", legend=alt.Legend(orient="bottom", columns=1, title=None)),
                    tooltip=["Mês", "Plano", alt.Tooltip("Ativos:Q", format=",.0f")]).properties(height=330), width="stretch")
            with c2:
                st.markdown("**Receita reconhecida (MRR) x dinheiro que entra (cobranças)**")
                lg = s.melt("Mês", ["Receita reconhecida", "Cobranças (caixa)"], "Série", "R$")
                st.altair_chart(alt.Chart(lg).mark_line(point=False, strokeWidth=2.5).encode(
                    x=alt.X("Mês:Q", scale=alt.Scale(domain=[1, 36])), y="R$:Q",
                    color=alt.Color("Série:N", scale=alt.Scale(range=["#2f6fd0", "#e67e22"]),
                                    legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["Mês", "Série", alt.Tooltip("R$:Q", format=",.2f")]).properties(height=330), width="stretch")
            anos = []
            for k in (1, 2, 3):
                bloco = s.iloc[(k - 1) * 12:k * 12]
                r_, c_ = float(bloco["Receita reconhecida"].sum()), float(bloco["Cobranças (caixa)"].sum())
                ce, tx = float(bloco["Custo de entrega"].sum()), float(bloco["Taxas de cobrança"].sum())
                anos.append({"Ano": f"Ano {k}", "Novos": num_br(bloco["Novos"].sum(), 0),
                             "Cancelamentos": num_br(bloco["Cancelamentos"].sum(), 0),
                             "Ativos no fim": num_br(bloco["Ativos"].iloc[-1], 0), "Receita reconhecida": brl(r_),
                             "Cobranças (caixa)": brl(c_), "Receita antecipada acumulada": brl(res["anos"][k]["receita_diferida"]),
                             "Custo de entrega": brl(ce), "Taxas": brl(tx), "Margem bruta": pct(div(r_ - ce - tx, r_))})
            st.dataframe(pd.DataFrame(anos), hide_index=True, width="stretch")
            with st.expander("Ver a projeção mês a mês"):
                vis = s.copy()
                for c in ("Novos", "Cancelamentos", "Ativos"):
                    vis[c] = vis[c].map(lambda x: num_br(x, 1))
                for c in ("Receita reconhecida", "Cobranças (caixa)", "Custo de entrega", "Taxas de cobrança"):
                    vis[c] = vis[c].map(brl)
                st.dataframe(vis, hide_index=True, width="stretch", height=420)
            st.caption("Assinantes aparecem com casas decimais porque o churn é uma média: 2,5 assinantes perdidos significa "
                       "\"entre 2 e 3\". A receita das assinaturas entra automaticamente na DRE, no fluxo de caixa e nos indicadores.")
    with t3:
        if not rc.get("tem"):
            st.info("Cadastre planos com preço e novos assinantes para ver os indicadores.")
        else:
            k = st.columns(4)
            k[0].metric("MRR no Mês 12", brl(rc["mrr12"]), f"Mês 1: {brl(rc['mrr1'])}", delta_color="off",
                        help="Receita recorrente mensal: soma do valor mensal equivalente de todos os assinantes ativos.")
            k[1].metric("ARR (MRR x 12)", brl(rc["arr12"]), help="Receita recorrente anualizada no ritmo do Mês 12.")
            k[2].metric("Assinantes no Mês 12", num_br(rc["ativos12"], 0))
            k[3].metric("ARPU (receita média por assinante)", brl(rc["arpu"]), help="MRR ÷ assinantes ativos no Mês 12.")
            k = st.columns(4)
            k[0].metric("Churn mensal médio", pct(rc["churn"]), f"vida média ≈ {num_br(rc['vida_meses'], 0)} meses",
                        delta_color="off", help="Churn de cada plano ponderado pelos assinantes do Mês 12. Vida média = 1 ÷ churn.")
            k[1].metric("Margem bruta recorrente (Ano 1)", pct(rc["margem_bruta"]),
                        help="(Receita reconhecida − custo de entrega − taxas de cobrança) ÷ receita reconhecida.")
            k[2].metric("LTV (lucro bruto por assinante na vida)", brl(rc["ltv"]) if rc["ltv"] == rc["ltv"] else "n.a.",
                        help="ARPU × margem bruta ÷ churn.")
            k[3].metric("CAC (custo para conquistar 1 assinante)", brl(rc["cac"]) if rc["cac"] == rc["cac"] else "n.a.",
                        help=f"Gasto comercial fixo do Mês 1 ({brl(rc['gasto_aquisicao_m1'])}: pessoal comercial + despesas "
                             f"comerciais) ÷ {num_br(rc['novos1'], 0)} novos assinantes no Mês 1.")
            lc, pb = rc["ltv_cac"], rc["payback_cac"]
            k = st.columns(3)
            k[0].metric("LTV/CAC", f"{num_br(lc, 1)}x" if lc == lc else "n.a.",
                        help="Referência: 3x ou mais. Abaixo de 1x, cada cliente novo dá prejuízo.")
            k[1].metric("Payback do CAC", f"{num_br(pb, 1)} meses" if pb == pb else "n.a.",
                        help="Meses de lucro bruto de um assinante para pagar o custo de conquistá-lo. Referência: até 12 meses.")
            k[2].metric("Receita antecipada no fim do Ano 1", brl(rc["receita_diferida_a1"]),
                        help="Dinheiro de planos já cobrados e ainda não prestados. É caixa, mas também é dívida com o cliente.")
            st.markdown("#### Leitura rápida")
            leit = []
            leit.append(("🟢" if lc == lc and lc >= 3 else ("🟡" if lc == lc and lc >= 1 else "🔴"),
                         f"LTV/CAC de {num_br(lc, 1)}x: " + ("a aquisição de clientes se paga com folga." if lc == lc and lc >= 3 else
                                                           "a conquista de clientes está cara para o que eles deixam de lucro.")))
            leit.append(("🟢" if rc["churn"] <= 0.03 else ("🟡" if rc["churn"] <= 0.07 else "🔴"),
                         f"Churn de {pct(rc['churn'])} ao mês: em um ano, cerca de {pct(1 - (1 - rc['churn']) ** 12, 0)} da base "
                         "seria substituída só para manter o tamanho."))
            leit.append(("🟢" if rc["margem_bruta"] >= 0.7 else ("🟡" if rc["margem_bruta"] >= 0.5 else "🔴"),
                         f"Margem bruta de {pct(rc['margem_bruta'])}: empresas de software costumam operar entre 70% e 85%."))
            for ic, tx in leit:
                st.markdown(f"{ic} {tx}")
    with t4:
        st.markdown("""
Classifique cada gasto de tecnologia pelo jeito como ele cresce:

| Tipo de gasto | Exemplos | Onde lançar no programa |
|---|---|---|
| **Cresce com cada assinante** | servidor/nuvem por uso, banco de dados, armazenamento, envio de SMS e e-mail, APIs pagas por chamada, uso de inteligência artificial, suporte ao cliente | Coluna **Custo de entrega por assinante** (aba Planos) |
| **% de cada cobrança** | meio de pagamento (cartão, Pix, boleto), comissão das lojas de aplicativos sobre assinaturas vendidas dentro do app | Coluna **Taxa de cobrança** (aba Planos) |
| **Fixo por mês** | infraestrutura mínima de nuvem, licenças de ferramentas (repositório de código, design, monitoramento), domínio, contas de desenvolvedor | 🧾 Custos e despesas → Custos fixos, classificação **Custo de tecnologia / infraestrutura** |
| **Equipe técnica** | desenvolvedores, suporte, DevOps | 🧾 Custos e despesas → Mão de obra, área **Produção / operação** |
| **Antes de abrir** | desenvolvimento da primeira versão (MVP), identidade visual, registro de marca | 🏗️ Investimentos → Pré-operacionais |
| **Equipamentos** | computadores e celulares para teste | 🏗️ Investimentos → Investimentos fixos (com depreciação) |

As lojas de aplicativos cobram comissão sobre compras e assinaturas feitas dentro do app, e essas regras vêm mudando
por decisões judiciais e regulatórias em vários países. Consulte as condições vigentes antes de fechar o plano. Vender para
empresas por boleto ou Pix, fora da loja, costuma reduzir a taxa de cobrança.
""")
        cf = d["custos_fixos"]
        tec = cf[cf["classificacao"].astype(str).str.startswith("Custo de tecnologia")]
        st.markdown(f"**Custos fixos de tecnologia cadastrados:** {brl(tec['valor_mensal'].sum())}/mês")
        if not tec.empty:
            st.dataframe(tec[["descricao", "valor_mensal"]].rename(columns={"descricao": "Item", "valor_mensal": "R$/mês"})
                         .assign(**{"R$/mês": lambda x: x["R$/mês"].map(brl)}), hide_index=True, width="stretch")
        if st.button("Lançar custos fixos de tecnologia →"):
            ir("custos", "Na aba Custos fixos, use a classificação \"Custo de tecnologia / infraestrutura\".")
            st.rerun()
    with t5:
        st.markdown(r"""
#### Faturamento não é receita
Um cliente que paga **R$ 3.000 por um plano anual** em janeiro:
- **Caixa:** entra R$ 3.000 em janeiro.
- **Receita (DRE):** R$ 250 por mês, de janeiro a dezembro, porque o serviço é prestado ao longo do ano.
- **Balanço:** em janeiro sobra R$ 2.750 de **receita antecipada**, uma obrigação com o cliente que diminui a cada mês.

#### Erro comum em planos de negócio
Lançar os planos trimestrais, semestrais e anuais como se fossem **vendidos de novo todo mês** (por exemplo, "8 planos
anuais de R$ 4.200 por mês" somando R$ 33.600 de receita mensal). Isso multiplica a receita: o plano anual é cobrado
uma vez por ano, e quem já assinou não paga de novo no mês seguinte. O certo é projetar **novos assinantes**,
**cancelamentos** e **renovações**, como esta ferramenta faz.

#### Os indicadores
- **MRR** (receita recorrente mensal) = soma dos valores mensais equivalentes dos assinantes ativos. **ARR** = MRR × 12.
- **Churn** = % de assinantes que cancelam por mês. Vida média do cliente ≈ 1 ÷ churn.
- **ARPU** = MRR ÷ assinantes ativos.
- **Margem bruta** = (receita − custo de entrega − taxas) ÷ receita.
- **LTV** (valor do cliente na vida) = ARPU × margem bruta ÷ churn.
- **CAC** (custo de aquisição) = gastos de marketing e vendas ÷ novos assinantes.
- **LTV/CAC**: referência de 3x ou mais. **Payback do CAC** = CAC ÷ (ARPU × margem bruta), idealmente até 12 meses.

#### Como o programa calcula
Cada mês de entrada forma uma **safra**. A cada renovação, a safra perde a parte que cancela. Quem paga por ciclo
(trimestral, anual) só sai na renovação: o churn mensal informado é convertido para o ciclo,
1 − (1 − churn)^meses do ciclo. Nos Anos 2 e 3, a entrada de novos assinantes parte do Mês 12 e cresce à taxa anual de
crescimento das premissas.
""")
    proxima("assinaturas")
