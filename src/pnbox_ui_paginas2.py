# ==============================================================================
# VIABILIZALAB - INTERFACE: FERRAMENTAS FINANCEIRAS
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
                               "prazo_meses": N("Prazo (meses)", min_value=1, step=1),
                               "carencia_meses": N("Carência (meses)", min_value=0, step=1),
                               "sistema": S("Sistema", options=["Price", "SAC"])}, "descricao")
        if st.button("🏦 Ver parcelas e comparar propostas de crédito", key="inv_fin"):
            ir("financiamento")
            st.rerun()
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
            "qtd_estoque_inicial": N("Estoque inicial (só revenda)", min_value=0, step=1),
            "horas_unit": N("Horas de trabalho por unidade (serviços)", min_value=0.0, step=0.5,
                            help="Quanto tempo da equipe cada serviço consome. Usado em 🧑‍💻 Serviços e custo da hora.")},
            "descricao")
        st.caption("Vende planos com pagamento recorrente (mensal, trimestral, anual)? Cadastre-os em "
                   "**💳 Assinaturas e receita recorrente**, não aqui: lá o programa considera renovações e cancelamentos.")
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
| **Custo de tecnologia / infraestrutura** | Base fixa para entregar um serviço digital | nuvem mínima, licenças de software, domínio, contas de desenvolvedor | Custos fixos classificados como "Custo de tecnologia" (gastos por assinante ficam em 💳 Assinaturas) |
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
                                "atribuicoes": T("Atribuições", width="medium"),
                                "mes_inicio": N("Começa no mês", min_value=1, max_value=120, step=1,
                                                help="1 = desde a abertura. Use 13 para contratar no início do Ano 2, 25 no Ano 3...")},
                    "cargo")
        st.caption("📅 **Contratações futuras:** para planejar quem entra depois, crie uma linha com o cargo e informe o mês de "
                   "início. O custo entra na DRE e no fluxo a partir desse mês.")
        r = R()
        if (d["pessoal"]["mes_inicio"] > 1).any():
            ps = r["pes_serie"]
            st.altair_chart(alt.Chart(pd.DataFrame({"Mês": range(1, len(ps) + 1), "Custo com pessoal": ps})).mark_area(
                line={"color": "#2f6fd0"}, color="#2f6fd033").encode(x="Mês:Q", y=alt.Y("Custo com pessoal:Q", title="R$/mês (sem inflação)"))
                .properties(height=200, title="Custo com pessoal mês a mês, com as contratações planejadas"), width="stretch")
        c = st.columns(4)
        c[0].metric("Total com pessoal / mês (Mês 1)", brl(r["pessoal_m"]))
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
    st.bar_chart(pd.DataFrame({"Receita": [r["anos"][a]["rec_bruta"] for a in r["ANOS"]],
                               "EBITDA": [r["anos"][a]["ebitda"] for a in r["ANOS"]],
                               "Resultado líquido": [r["anos"][a]["resultado"] for a in r["ANOS"]]},
                              index=[f"Ano {a:02d}" for a in r["ANOS"]]), stack=False)
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
    with st.expander("📖 O que significam estes números? (leitura em linguagem simples)"):
        for nome, valor, leitura in explicar_indicadores(r):
            st.markdown(f"**{nome}: {valor}**  \n{leitura}")
    st.subheader("Resumo financeiro (Ano 1, padrão Sebrae)")
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
    k[3].metric("Payback descontado", f"{num_br(r['payback_desc'], 1)} meses" if r["payback_desc"] == r["payback_desc"] else f"> {r['MESES']} meses")

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
            "na alíquota sobre vendas). Opcionalmente recupera o capital de giro no último ano do horizonte.\n"
            "- **Payback simples e rentabilidade** seguem o modelo do Sebrae: investimento ÷ resultado operacional mensal do Ano 1.\n"
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
