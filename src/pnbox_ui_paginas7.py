# ==============================================================================
# VIABILIZALAB 8 - INTERFACE DAS NOVAS FERRAMENTAS
# ==============================================================================


# ------------------------------------------------------------------------------
# TRILHA GUIADA E "O QUE FALTA"
# ------------------------------------------------------------------------------
def proximo_passo() -> str | None:
    """Primeira ferramenta essencial que ainda não está concluída."""
    for p in ESSENCIAIS:
        if p in PAG and situacao(p) != "✅":
            return p
    for p in [x[0] for x in PAGINAS]:
        if situacao(p) == "◐":
            return p
    return None


def pg_trilha():
    cabecalho("trilha")
    st.caption("A trilha mostra a ordem recomendada de preenchimento. As ferramentas essenciais formam o plano mínimo; "
               "as demais aprofundam a análise.")
    prox = proximo_passo()
    if prox:
        with st.container(border=True):
            st.markdown(f"### 👉 Próximo passo: {PAG[prox][2]} {PAG[prox][3]}")
            st.caption(PAG[prox][4])
            if st.button("Ir para esta ferramenta", type="primary", key="trilha_ir"):
                ir(prox)
                st.rerun()
    else:
        st.success("Todas as ferramentas essenciais estão concluídas. Agora vale aprofundar a validação e preparar a defesa.")
    ess = [p for p in ESSENCIAIS if p in PAG]
    feitos = sum(1 for p in ess if situacao(p) == "✅")
    st.progress(feitos / len(ess), text=f"Plano mínimo: {feitos} de {len(ess)} ferramentas essenciais concluídas")
    st.subheader("O que falta, por ordem de importância")
    chk = CHK()
    pend = chk[chk["Situação"] != "OK"]
    n = 0
    for i, r in pend.iterrows():
        n += 1
        a, b = st.columns([4, 1.2])
        a.markdown(f"⚠️ **{r['Teste']}**: {r['Valor apurado']}  \n<small>{r['Como corrigir'] or r['Observação']}</small>",
                   unsafe_allow_html=True)
        if r["destino"] and b.button("Corrigir →", key=f"tri_chk_{i}", width="stretch"):
            ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir'] or r['Observação']}")
            st.rerun()
    for p in ess:
        s_ = situacao(p)
        if s_ != "✅":
            n += 1
            a, b = st.columns([4, 1.2])
            a.markdown(f"{'◐' if s_ == '◐' else '○'} **{PAG[p][3]}** ({'em andamento' if s_ == '◐' else 'não iniciada'})")
            if b.button("Abrir →", key=f"tri_pg_{p}", width="stretch"):
                ir(p)
                st.rerun()
    if n == 0:
        st.success("Nada pendente nas ferramentas essenciais e nas checagens.")
    st.subheader("Mapa completo")
    for g in GRUPOS:
        itens = [p for p in PAGINAS if p[1] == g]
        st.markdown(f"**{g}**: " + " · ".join(f"{situacao(p[0]) or '·'} {p[3]}" for p in itens))


# ------------------------------------------------------------------------------
# GLOSSÁRIO
# ------------------------------------------------------------------------------
def pg_glossario():
    cabecalho("glossario")
    try:
        st.subheader("📖 O que significam os números deste projeto")
        for nome, valor, leitura in explicar_indicadores(R()):
            with st.container(border=True):
                st.markdown(f"**{nome}: {valor}**  \n{leitura}")
    except Exception:
        st.info("A leitura dos números aparece quando houver produtos, custos e investimentos.")
    st.subheader("🔤 Termos")
    busca = st.text_input("Procurar termo", key="glos_busca", placeholder="Ex.: VPL, markup, churn")
    for termo, defi in sorted(GLOSSARIO.items(), key=lambda x: x[0].lower()):
        if busca and busca.lower() not in (termo + defi).lower():
            continue
        st.markdown(f"**{termo}**: {defi}")
    proxima("glossario")


# ------------------------------------------------------------------------------
# IMPORTAR E EXPORTAR PLANILHAS
# ------------------------------------------------------------------------------
TABELAS_IMPORTAVEIS = {
    "catalogo": ("Produtos e serviços", {"categoria": "Categoria", "descricao": "Descrição", "preco_venda": "Preço de venda",
                                         "custo_aquisicao": "Custo unitário", "qtd_mensal": "Quantidade por mês",
                                         "qtd_estoque_inicial": "Estoque inicial", "horas_unit": "Horas por unidade"}),
    "custos_fixos": ("Custos fixos", {"descricao": "Descrição", "classificacao": "Classificação", "valor_mensal": "Valor mensal"}),
    "inv_fixos": ("Investimentos fixos", {"descricao": "Descrição", "quantidade": "Quantidade", "valor_unitario": "Valor unitário",
                                          "vida_util_anos": "Vida útil (anos)", "valor_residual_pct": "Valor residual (%)"}),
    "materiais": ("Materiais", {"nome": "Material", "categoria": "Categoria", "unidade": "Unidade", "valor_unitario": "Valor unitário"}),
    "pessoal": ("Equipe", {"cargo": "Cargo", "area": "Área", "num_empregados": "Pessoas", "salario_mensal": "Salário",
                           "encargos_pct": "Encargos (%)", "beneficios_unit": "Benefícios", "mes_inicio": "Mês de início"}),
    "entrevistas": ("Entrevistas", {"data": "Data", "entrevistado": "Entrevistado", "perfil": "Perfil", "dores": "Dores",
                                    "solucao_atual": "Solução atual", "pagaria": "Pagaria", "valor_pagaria": "Valor que pagaria",
                                    "aprendizado": "Aprendizado"}),
    "precos_vw": ("Teste de preço (Van Westendorp)", {"respondente": "Respondente", "muito_barato": "Barato demais",
                                                      "barato": "Barato", "caro": "Caro", "muito_caro": "Caro demais"}),
    "concorrentes": ("Concorrentes", {"nome": "Nome"}),
    "riscos": ("Riscos", {"risco": "Risco", "probabilidade": "Probabilidade", "impacto": "Impacto", "mitigacao": "Mitigação",
                          "responsavel": "Responsável"}),
}


def pg_planilha():
    cabecalho("planilha")
    t1, t2 = st.tabs(["⬆️ Importar de planilha", "⬇️ Exportar para Excel"])
    with t1:
        st.caption("Traga listas grandes de uma planilha em vez de digitar linha a linha. A primeira linha da planilha precisa ter "
                   "os nomes das colunas; o programa reconhece os nomes do modelo, com ou sem acentos.")
        tab = st.selectbox("Para qual ferramenta?", list(TABELAS_IMPORTAVEIS), format_func=lambda k: TABELAS_IMPORTAVEIS[k][0],
                           key="imp_tab")
        nome_t, rot = TABELAS_IMPORTAVEIS[tab]
        modelo = pd.DataFrame(columns=list(rot.values()))
        b = io.BytesIO()
        with pd.ExcelWriter(b, engine="openpyxl") as w:
            modelo.to_excel(w, index=False, sheet_name=nome_t[:31])
        st.download_button(f"📄 Baixar planilha-modelo de {nome_t}", b.getvalue(), f"modelo_{tab}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        arq = st.file_uploader("Planilha preenchida (.xlsx ou .csv)", type=["xlsx", "csv"], key=wk(f"imp_up_{tab}"))
        if arq:
            try:
                bruto = ler_planilha(arq.name, arq.getvalue())
            except Exception as e:
                st.error(f"Não foi possível ler a planilha: {e}")
                return proxima("planilha")
            mapa = mapear_colunas(bruto, colunas_de(tab), rot)
            if not mapa:
                st.error("Nenhuma coluna reconhecida. Use os nomes da planilha-modelo.")
                return proxima("planilha")
            novo = bruto[list(mapa)].rename(columns=mapa)
            for c in colunas_de(tab):
                if c not in novo:
                    novo[c] = None
            tipos = dict(SCHEMA[tab])
            for c in novo.columns:
                if tipos.get(c) in ("REAL", "INTEGER"):
                    novo[c] = novo[c].map(converter_numero)
            novo = limpar(novo[colunas_de(tab)].dropna(how="all"), tab)
            st.write(f"**{len(novo)} linha(s)** reconhecida(s). Colunas usadas: " + ", ".join(f"{a} → {rot.get(b_, b_)}"
                                                                                              for a, b_ in mapa.items()))
            st.dataframe(novo.head(20), hide_index=True, width="stretch")
            modo = st.radio("Como incluir", ["Acrescentar às linhas atuais", "Substituir as linhas atuais"], horizontal=True,
                            key="imp_modo")
            if st.button("✅ Importar", type="primary", key="imp_ok"):
                base = novo if modo.startswith("Substituir") else pd.concat([d[tab], novo], ignore_index=True)
                recarregar_tabela(tab, base, colunas_de(tab)[0])
                st.success(f"{len(novo)} linha(s) importada(s) em {nome_t}.")
    with t2:
        st.caption("Uma pasta de trabalho com todas as tabelas do projeto, a DRE, o fluxo de caixa e os indicadores.")
        if st.button("Gerar planilha", key="exp_xlsx"):
            ss.xlsx_v8 = planilha_excel(R())
        if ss.get("xlsx_v8"):
            st.download_button("⬇️ Baixar planilha (Excel)", ss.xlsx_v8, f"{proj['nome_negocio'] or 'projeto'}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    proxima("planilha")


# ------------------------------------------------------------------------------
# HISTÓRICO DE VERSÕES DO PROJETO (com comparação de indicadores)
# ------------------------------------------------------------------------------
def indicadores_de_versao(vid: int) -> dict | None:
    linha = conn.execute("SELECT dados FROM versoes WHERE id=?", (vid,)).fetchone()
    if not linha:
        return None
    dados = json.loads(linha[0])
    p_ = dados.get("projeto", [{}])[0] if dados.get("projeto") else {}
    p_ = {**PROJETO_PADRAO, **{k: v for k, v in p_.items() if v is not None}}
    d_ = {}
    for t in SCHEMA:
        if t in ("projeto", "versoes"):
            continue
        cols = colunas_de(t)
        df = pd.DataFrame(dados.get(t, []))
        for c in cols:
            if c not in df:
                df[c] = None
        d_[t] = limpar(df[cols] if len(df) else pd.DataFrame(columns=cols), t)
    pv = dados.get("proposta_valor")
    d_["proposta_valor"] = pd.DataFrame(pv) if pv else d_.get("proposta_valor", pd.DataFrame())
    r_ = executar_motor(p_, d_)
    return {"Receita Ano 1": r_["anos"][1]["rec_bruta"], "Lucro Ano 1": r_["anos"][1]["resultado"],
            "Investimento": r_["investimento_total"], "VPL": r_["vpl"], "TIR": r_["tir"]}


def pg_historico():
    cabecalho("historico")
    st.caption("Cada versão guarda uma foto completa do projeto. Compare os indicadores para ver o efeito de cada mudança e "
               "restaure uma versão anterior se precisar. O programa também cria cópias automáticas a cada 10 minutos.")
    a, b = st.columns([3, 1])
    nome_v = a.text_input("Nome da versão", placeholder="Ex.: Depois de testar o preço", label_visibility="collapsed", key="hist_nome")
    if b.button("💾 Salvar versão agora", width="stretch", key="hist_salvar"):
        salvar_versao(nome_v.strip() or "Versão manual")
        st.rerun()
    vs = pd.read_sql("SELECT id, nome, data FROM versoes ORDER BY id DESC", conn)
    if vs.empty:
        st.info("Nenhuma versão salva ainda.")
        return proxima("historico")
    if st.button("📊 Comparar os indicadores de todas as versões", key="hist_comp"):
        linhas_ = []
        for _, v in vs.iterrows():
            try:
                ind = indicadores_de_versao(int(v["id"]))
            except Exception:
                ind = None
            if ind:
                linhas_.append({"Versão": v["nome"], "Data": v["data"], **ind})
        r_ = R()
        linhas_.insert(0, {"Versão": "▶ Projeto atual", "Data": "agora", "Receita Ano 1": r_["anos"][1]["rec_bruta"],
                           "Lucro Ano 1": r_["anos"][1]["resultado"], "Investimento": r_["investimento_total"],
                           "VPL": r_["vpl"], "TIR": r_["tir"]})
        ss.hist_comp = pd.DataFrame(linhas_)
    if ss.get("hist_comp") is not None:
        st.dataframe(ss.hist_comp, hide_index=True, width="stretch", column_config={
            c: N(format="R$ %.0f") for c in ("Receita Ano 1", "Lucro Ano 1", "Investimento", "VPL")} | {"TIR": N(format="percent")})
    for _, v in vs.iterrows():
        a, b, c = st.columns([4, 1, 1])
        a.markdown(f"**{v['nome']}** · {v['data']}")
        if b.button("Restaurar", key=f"hrest_{v['id']}", width="stretch"):
            restaurar_versao(int(v["id"]))
            ss.hist_comp = None
            st.toast("Versão restaurada. Uma cópia do estado anterior foi guardada.")
            st.rerun()
        if c.button("Apagar", key=f"hapv_{v['id']}", width="stretch"):
            conn.execute("DELETE FROM versoes WHERE id=?", (int(v["id"]),))
            conn.commit()
            ss.hist_comp = None
            st.rerun()
    proxima("historico")


# ------------------------------------------------------------------------------
# FINANCIAMENTO: SAC, PRICE, CARÊNCIA E COMPARAÇÃO DE PROPOSTAS
# ------------------------------------------------------------------------------
def pg_financiamento():
    cabecalho("financiamento")
    t1, t2 = st.tabs(["🏦 Empréstimos do plano", "⚖️ Comparar propostas de crédito"])
    with t1:
        st.caption("Estes empréstimos entram no investimento, no fluxo de caixa e na DRE. Escolha o sistema de amortização e a "
                   "carência (meses em que se paga só os juros).")
        editor("emprestimos", {"descricao": T("Linha de crédito", required=True, width="large"), "valor": moeda("Valor (R$)"),
                               "taxa_anual_pct": N("Juros (% ao ano)", min_value=0.0), "prazo_meses": N("Prazo total (meses)", min_value=1, step=1),
                               "carencia_meses": N("Carência (meses)", min_value=0, step=1),
                               "sistema": S("Sistema", options=["Price", "SAC"])}, "descricao")
        emp = d["emprestimos"]
        for _, r in emp.iterrows():
            if float(r["valor"] or 0) <= 0:
                continue
            rs = resumo_emprestimo(r["valor"], r["taxa_anual_pct"], r["prazo_meses"], r["carencia_meses"], r["sistema"])
            with st.expander(f"📄 {r['descricao']}: parcela de {brl(rs['primeira'])} a {brl(rs['menor'])} · juros totais {brl(rs['juros'])}"):
                k = st.columns(4)
                k[0].metric("Total pago", brl(rs["total_pago"]))
                k[1].metric("Juros totais", brl(rs["juros"]))
                k[2].metric("Maior parcela", brl(rs["maior"]))
                k[3].metric("CET (a.a.)", pct(rs["cet"]))
                g = rs["tabela"].melt(id_vars="Mês", value_vars=["Juros", "Amortização"], var_name="Parte", value_name="R$")
                st.altair_chart(alt.Chart(g).mark_bar().encode(x="Mês:O", y="R$:Q", color=alt.Color("Parte:N", legend=alt.Legend(orient="bottom")))
                                .properties(height=220), width="stretch")
                st.dataframe(rs["tabela"], hide_index=True, width="stretch", height=240,
                             column_config={c: N(format="R$ %.2f") for c in ("Juros", "Amortização", "Parcela", "Saldo devedor")})
    with t2:
        st.caption("Compare propostas antes de decidir. As taxas mudam com frequência: use as condições informadas pelo banco "
                   "ou pela cooperativa na data da simulação.")
        base = ss.setdefault("propostas_credito", pd.DataFrame([
            {"Proposta": "Proposta A (exemplo)", "Valor": 30000.0, "Juros (% a.a.)": 18.0, "Prazo (meses)": 36, "Carência (meses)": 6, "Sistema": "Price"},
            {"Proposta": "Proposta B (exemplo)", "Valor": 30000.0, "Juros (% a.a.)": 22.0, "Prazo (meses)": 24, "Carência (meses)": 0, "Sistema": "SAC"}]))
        ed = st.data_editor(base, num_rows="dynamic", hide_index=True, width="stretch", key="ed_propostas",
                            column_config={"Sistema": S(options=["Price", "SAC"]), "Valor": N(format="R$ %.2f")})
        ss.propostas_credito = ed
        linhas_ = []
        for _, r in ed.dropna(subset=["Valor"]).iterrows():
            try:
                rs = resumo_emprestimo(r["Valor"], r["Juros (% a.a.)"], r["Prazo (meses)"], r["Carência (meses)"], r["Sistema"])
            except Exception:
                continue
            linhas_.append({"Proposta": r["Proposta"], "Primeira parcela": rs["primeira"], "Maior parcela": rs["maior"],
                            "Total pago": rs["total_pago"], "Juros totais": rs["juros"], "CET (a.a.)": rs["cet"]})
        if linhas_:
            cmp_ = pd.DataFrame(linhas_)
            st.dataframe(cmp_, hide_index=True, width="stretch", column_config={
                c: N(format="R$ %.2f") for c in ("Primeira parcela", "Maior parcela", "Total pago", "Juros totais")} | {"CET (a.a.)": N(format="percent")})
            mel = cmp_.sort_values("CET (a.a.)").iloc[0]
            st.info(f"Menor custo efetivo: **{mel['Proposta']}** ({pct(mel['CET (a.a.)'])} ao ano). Confira também se a parcela cabe "
                    "no fluxo de caixa dos primeiros meses.")
    proxima("financiamento")


# ------------------------------------------------------------------------------
# DADOS PÚBLICOS DO MUNICÍPIO (IBGE) E POF
# ------------------------------------------------------------------------------
def _gravar_dado(chave, rotulo, valor, periodo, fonte):
    dp = d["dados_publicos"]
    dp = dp[dp["chave"] != chave]
    dp = pd.concat([dp, pd.DataFrame([{"chave": chave, "rotulo": rotulo, "valor": valor, "periodo": periodo, "fonte": fonte}])],
                   ignore_index=True)
    salvar(conn, "dados_publicos", dp, "chave")
    d["dados_publicos"] = limpar(dp.reset_index(drop=True), "dados_publicos")


def pg_dados_publicos():
    cabecalho("dados_publicos")
    st.caption("Busca dados oficiais do IBGE pela internet (gratuito) para dimensionar o mercado com números verificáveis. "
               "Os valores ficam guardados no projeto e entram no relatório com a fonte.")
    t1, t2 = st.tabs(["🏙️ Município (Censo 2022, PIB, empresas)", "🛒 Gastos das famílias (POF 2017-2018)"])
    with t1:
        c = st.columns([1, 3])
        ufs = list(UFS)
        uf = c[0].selectbox("UF", ufs, index=ufs.index(proj["uf_sigla"]) if proj["uf_sigla"] in ufs else ufs.index("MG"), key=wk("dp_uf"))
        chave_m = f"mun_{uf}"
        if chave_m not in ss:
            try:
                ss[chave_m] = municipios_da_uf(uf)
            except Exception as e:
                ss[chave_m] = []
                st.warning(f"Não foi possível consultar a lista de municípios do IBGE agora ({e}). Verifique a internet.")
        lista = ss[chave_m]
        nomes = [n for _, n in lista]
        idx = nomes.index(proj["municipio_nome"]) if proj["municipio_nome"] in nomes else 0
        nome = c[1].selectbox("Município", nomes, index=idx, key=wk("dp_mun")) if nomes else None
        if nome and st.button("🔎 Buscar dados do IBGE", type="primary", key="dp_buscar"):
            cod = dict((n, i) for i, n in lista)[nome]
            with st.spinner("Consultando a API do IBGE..."):
                dados = dados_municipio(cod)
            if not dados:
                st.error("O IBGE não respondeu. Tente de novo em alguns minutos.")
            else:
                for ch, rot, v, per, fonte in dados:
                    _gravar_dado(ch, rot, v, per, fonte)
                auto_unica("projeto", {"uf_sigla": uf, "municipio_nome": nome, "municipio_ibge": cod})
                st.rerun()
        dp = d["dados_publicos"]
        mun = dp[~dp["chave"].astype(str).str.startswith("pof")]
        if not mun.empty:
            st.markdown(f"**{proj['municipio_nome'] or ''} ({proj['uf_sigla'] or ''})**")
            cols = st.columns(3)
            for i, (_, r) in enumerate(mun.iterrows()):
                v = r["valor"]
                txt = brl(v) if "R$" in r["rotulo"] and "mil" not in r["rotulo"] else num_br(v, 1 if v < 100 else 0)
                cols[i % 3].metric(f"{r['rotulo']} ({r['periodo']})", txt, help=r["fonte"])
            m = {r["chave"]: r["valor"] for _, r in mun.iterrows()}
            st.markdown("**Usar no Tamanho do mercado (TAM)**")
            b = st.columns(3)
            if "populacao" in m and b[0].button(f"TAM = população ({num_br(m['populacao'], 0)})", key="dp_tam_pop"):
                auto_unica("projeto", {"tam_clientes": m["populacao"], "tam_desc": f"População de {proj['municipio_nome']} (Censo 2022, IBGE)"})
                st.toast("TAM atualizado.")
            if "domicilios" in m and b[1].button(f"TAM = domicílios ({num_br(m['domicilios'], 0)})", key="dp_tam_dom"):
                auto_unica("projeto", {"tam_clientes": m["domicilios"], "tam_desc": f"Domicílios de {proj['municipio_nome']} (Censo 2022, IBGE)"})
                st.toast("TAM atualizado.")
            if "unidades_locais" in m and b[2].button(f"TAM = empresas ({num_br(m['unidades_locais'], 0)})", key="dp_tam_emp"):
                auto_unica("projeto", {"tam_clientes": m["unidades_locais"],
                                       "tam_desc": f"Unidades locais de empresas em {proj['municipio_nome']} (Cempre, IBGE)"})
                st.toast("TAM atualizado.")
    with t2:
        uf2 = proj["uf_sigla"] or "MG"
        st.caption(f"Despesa média mensal das famílias por tipo de gasto em {UFS.get(uf2, (0, 'Brasil'))[1]} (POF 2017-2018). "
                   "Útil para estimar quanto um domicílio gasta por ano na categoria do seu negócio.")
        if st.button("Carregar gastos das famílias", key="dp_pof"):
            with st.spinner("Consultando a POF no IBGE..."):
                ss.pof_df, ss.pof_origem = pof_uf(uf2)
        if ss.get("pof_df") is not None:
            pof = ss.pof_df
            st.caption(f"Origem: {ss.pof_origem}. Valores de 2018, sem correção pela inflação.")
            esc = st.selectbox("Categoria de gasto do seu negócio", list(pof["Tipo de despesa"]), key="dp_pof_cat")
            v = float(pof.loc[pof["Tipo de despesa"] == esc, "R$ por mês por família"].iloc[0])
            k = st.columns(2)
            k[0].metric("Por mês, por família", brl(v))
            k[1].metric("Por ano, por família", brl(v * 12))
            if st.button("Guardar e usar como gasto anual por cliente", key="dp_pof_usar"):
                _gravar_dado("pof_" + _norm(esc)[:30], esc.split(" ", 1)[-1], v, "2017-2018",
                             f"IBGE. Pesquisa de Orçamentos Familiares 2017-2018, tabela 6715 ({UFS.get(uf2, (0, 'Brasil'))[1]}).")
                auto_unica("projeto", {"gasto_anual_cliente": round(v * 12, 2)})
                st.toast("Gasto anual por cliente atualizado.")
            st.dataframe(pof, hide_index=True, width="stretch", height=300, column_config={"R$ por mês por família": N(format="R$ %.2f")})
    proxima("dados_publicos")


# ------------------------------------------------------------------------------
# REFERÊNCIAS DO SETOR E SOBREVIVÊNCIA
# ------------------------------------------------------------------------------
def pg_referencias():
    cabecalho("referencias")
    res = R()
    st.caption("Compare os números do plano com a média das empresas brasileiras da mesma atividade, segundo o IBGE. Diferenças "
               "grandes não são erro por si só, mas precisam de explicação.")
    ops = atividades_referencia()
    atual = proj["setor_referencia"] if proj["setor_referencia"] in ops else ops[0]
    esc = st.selectbox("Atividade mais parecida com o seu negócio", ops, index=ops.index(atual), key=wk("ref_set"))
    auto_unica("projeto", {"setor_referencia": esc})
    if res["anos"][1]["rec_bruta"] > 0:
        cs = comparar_setor(esc, proj, d, res)
        vis = cs.copy()
        for c in ("Plano", "Setor (Brasil)"):
            vis[c] = [pct(v) if ("÷" in i or "margem" in i.lower()) else brl(v) for i, v in zip(cs["Indicador"], cs[c])]
        vis["Plano ÷ setor"] = cs["Plano ÷ setor"].map(lambda v: f"{num_br(v, 2)}x" if v == v else "n.a.")
        st.dataframe(vis, hide_index=True, width="stretch")
        for a in leitura_setor(cs):
            st.warning(a)
        st.caption(FONTE_PAC if esc.startswith("Comércio") else FONTE_PAS)
    st.subheader("Quantas empresas sobrevivem")
    k = st.columns(3)
    for col, (rot, v) in zip(k, SOBREVIVENCIA_5A.items()):
        col.metric(rot, pct(v, 1), help="Percentual de empresas ainda ativas cinco anos depois de abertas.")
    st.caption("Fonte: Sebrae, seminário DataSebrae \"Aplicações de análise de sobrevivência\" (jul. 2026), com dados da Receita Federal.")
    st.markdown("**Causas mais citadas por quem fechou** (pesquisa Sobrevivência de Empresas, Sebrae): " +
                "; ".join(f"{c} ({pct(v, 0)})" for c, v in CAUSAS_FECHAMENTO) + ".")
    a1 = res["anos"][1]
    minimo = float(res["fluxo_mensal"]["Saldo final"].min())
    if minimo < res["reserva"] * 0.5:
        st.warning(f"Atenção ao capital de giro: o menor saldo de caixa do Ano 1 é {brl(minimo)}. É a principal causa de fechamento.")
    if a1["rec_bruta"] and a1["rec_bruta"] < a1["pe_contabil"] * 1.2:
        st.warning("A receita prevista está a menos de 20% do ponto de equilíbrio: pouca folga se as vendas vierem abaixo do esperado.")
    proxima("referencias")


# ------------------------------------------------------------------------------
# SWOT CRUZADA
# ------------------------------------------------------------------------------
def pg_swot_cruzada():
    cabecalho("swot_cruzada")
    st.caption("A SWOT só vale se virar estratégia (Wright; Kroll; Parnell, 2000). Cruze os fatores priorizados e escreva o que a "
               "empresa vai fazer em cada combinação.")
    sug = sugestoes_cruzada(d)
    cols = st.columns(2)
    for i, (q, s_) in enumerate(sug.items()):
        with cols[i % 2].container(border=True):
            st.markdown(f"**{q}**  \n<small>{s_['dica']}</small>", unsafe_allow_html=True)
            st.caption(("Fatores: " + "; ".join(s_["a"][:3]) + "  ×  " + "; ".join(s_["b"][:3])) if s_["a"] and s_["b"]
                       else "Cadastre os fatores na análise SWOT.")
            n = int((d["swot_estrategias"]["quadrante"] == q).sum())
            st.caption(f"{n} estratégia(s) registrada(s)")
    if d["swot_estrategias"].empty and any(s_["pares"] for s_ in sug.values()):
        if st.button("Criar uma linha para cada combinação de fatores priorizados", key="swc_seed"):
            linhas_ = [{"quadrante": q, "estrategia": f"[{p}] ", "acao": "", "prazo_meses": 6.0}
                       for q, s_ in sug.items() for p in s_["pares"][:2]]
            recarregar_tabela("swot_estrategias", pd.DataFrame(linhas_), "estrategia")
            st.rerun()
    editor("swot_estrategias", {"quadrante": S("Tipo de estratégia", options=list(QUADRANTES_CRUZADA), required=True, width="medium"),
                                "estrategia": T("Estratégia", required=True, width="large"), "acao": T("Ação concreta", width="large"),
                                "prazo_meses": N("Prazo (meses)", min_value=1, step=1)}, "estrategia")
    proxima("swot_cruzada")


# ------------------------------------------------------------------------------
# LEAN CANVAS
# ------------------------------------------------------------------------------
def html_lean(bl: dict) -> str:
    cores = {"Problema": "#c0392b", "Alternativas existentes": "#c0392b", "Solução": "#2c7fb8", "Métricas-chave": "#2c7fb8",
             "Proposta única de valor": "#44923a", "Vantagem injusta": "#7b3fe4", "Canais": "#14a49a",
             "Segmentos de clientes": "#1592c4", "Primeiros clientes (early adopters)": "#1592c4",
             "Estrutura de custos": "#4456a0", "Fontes de receita": "#22617a"}
    areas = {"Problema": "pr", "Alternativas existentes": "al", "Solução": "so", "Métricas-chave": "me",
             "Proposta única de valor": "pv", "Vantagem injusta": "va", "Canais": "ca", "Segmentos de clientes": "se",
             "Primeiros clientes (early adopters)": "ea", "Estrutura de custos": "cu", "Fontes de receita": "re"}

    def caixa(nome):
        itens = "".join(f"<li>{_esc(x)}</li>" for x in bl.get(nome, [])) or "<li><i>a preencher</i></li>"
        return (f"<div class='lc-b' style='grid-area:{areas[nome]};background:{cores[nome]}'><div class='lc-t'>{nome.upper()}</div>"
                f"<ul>{itens}</ul></div>")
    css = """<style>
    .lc{display:grid;gap:8px;grid-template-columns:repeat(10,1fr);grid-template-areas:
      "pr pr so so pv pv va va se se" "pr pr so so pv pv va va se se" "al al me me pv pv ca ca ea ea" "cu cu cu cu cu re re re re re";}
    .lc-b{border-radius:12px;padding:12px;color:#fff;font-size:.83rem;min-height:110px}
    .lc-t{font-weight:800;font-size:.9rem;margin-bottom:6px} .lc-b ul{margin:0;padding-left:18px}
    @media (max-width:700px){.lc{grid-template-columns:1fr;grid-template-areas:"pr" "al" "se" "ea" "pv" "so" "ca" "re" "cu" "me" "va";}}
    </style>"""
    return css + "<div class='lc'>" + "".join(caixa(n) for n in areas) + "</div>"


def pg_lean():
    cabecalho("lean")
    st.caption("O Lean Canvas (Maurya, 2012) adapta o Business Model Canvas para negócios novos e incertos: começa pelo problema e "
               "pelas métricas que mostram se a solução funciona. Ele é montado com o que você já preencheu; complete os três campos abaixo.")
    c = st.columns(3)
    v = {"lean_alternativas": c[0].text_area("Alternativas existentes (uma por linha)", proj["lean_alternativas"], height=110,
                                             key=wk("lean_alt"), help="Como os clientes resolvem o problema hoje."),
         "lean_metricas": c[1].text_area("Métricas-chave (uma por linha)", proj["lean_metricas"], height=110, key=wk("lean_met"),
                                         help="Os poucos números que mostram se o negócio está funcionando."),
         "lean_vantagem": c[2].text_area("Vantagem injusta (uma por linha)", proj["lean_vantagem"], height=110, key=wk("lean_van"),
                                         help="Algo que não pode ser copiado ou comprado com facilidade.")}
    auto_unica("projeto", v)
    st.markdown(html_lean(lean_canvas(proj, d, R())), unsafe_allow_html=True)
    proxima("lean")


# ------------------------------------------------------------------------------
# VALUATION E RODADA DE INVESTIMENTO
# ------------------------------------------------------------------------------
def pg_valuation():
    cabecalho("valuation")
    res = R()
    st.caption("Quanto vale o negócio? Três métodos dão uma faixa de valor, não um número exato. A rodada mostra quanto da empresa "
               "um investidor levaria em troca do dinheiro.")
    c = st.columns(3)
    g = c[0].number_input("Crescimento na perpetuidade (% a.a.)", -5.0, 15.0, float(proj["val_crescimento_perp_pct"] or 3), 0.5, key=wk("val_g"),
                          help="Quanto o fluxo cresce por ano depois do horizonte. Deve ser menor que a TMA.")
    meb = c[1].number_input("Múltiplo de EBITDA", 0.5, 20.0, float(proj["val_multiplo_ebitda"] or 4), 0.5, key=wk("val_meb"))
    mrec = c[2].number_input("Múltiplo de receita", 0.1, 10.0, float(proj["val_multiplo_receita"] or 1), 0.1, key=wk("val_mrec"))
    auto_unica("projeto", {"val_crescimento_perp_pct": g, "val_multiplo_ebitda": meb, "val_multiplo_receita": mrec})
    vl = valuation(res, proj, d)
    st.dataframe(vl["metodos"], hide_index=True, width="stretch", column_config={"Valor do patrimônio dos sócios": N(format="R$ %.0f")})
    k = st.columns(3)
    k[0].metric("Valor presente dos fluxos do horizonte", brl(vl["vp_fluxos"]))
    k[1].metric("Valor presente da perpetuidade", brl(vl["vp_terminal"]))
    k[2].metric("Dívida (empréstimos)", brl(vl["divida"]))
    if vl["vp_terminal"] > vl["vp_fluxos"]:
        st.info("Mais da metade do valor vem da perpetuidade, ou seja, de depois do horizonte do plano. Isso é comum, mas torna o "
                "valor muito sensível à taxa de crescimento escolhida.")
    st.caption("Múltiplos usados em pequenos negócios costumam ser baixos e variam muito por setor; trate-os como referência didática.")
    st.subheader("Rodada de investimento")
    c = st.columns(2)
    pre = c[0].number_input("Valor da empresa antes do aporte (pre-money, R$)", 0.0, 1e10,
                            float(proj["rodada_pre_money"] or max(vl["equity_dcf"], 0.0)), 1000.0, key=wk("rod_pre"))
    inv = c[1].number_input("Valor do aporte do investidor (R$)", 0.0, 1e10, float(proj["rodada_investimento"] or 0), 1000.0, key=wk("rod_inv"))
    auto_unica("projeto", {"rodada_pre_money": pre, "rodada_investimento": inv})
    ct = cap_table(d, pre, inv)
    st.dataframe(ct, hide_index=True, width="stretch", column_config={
        "Participação antes": N(format="percent"), "Participação depois": N(format="percent"), "Valor da participação": N(format="R$ %.0f")})
    if inv > 0 and pre > 0:
        st.markdown(f"Com post-money de **{brl(pre + inv)}**, o investidor fica com **{pct(inv / (pre + inv))}** da empresa e cada 1% "
                    f"vale **{brl((pre + inv) / 100)}**.")
    proxima("valuation")


# ------------------------------------------------------------------------------
# IMPACTO SOCIAL E ODS
# ------------------------------------------------------------------------------
def pg_impacto():
    cabecalho("impacto")
    st.caption("Para negócios de impacto, editais e projetos de extensão: a teoria da mudança liga o que a empresa faz às "
               "transformações que pretende causar, e os indicadores mostram se isso está acontecendo.")
    v = {}
    cols = st.columns(3)
    for i, (k_, rot, ajuda) in enumerate(ETAPAS_MUDANCA):
        v[k_] = cols[i % 3].text_area(f"{i + 1}. {rot}", proj[k_], height=110, key=wk(k_), help=ajuda)
    ods_atual = [o for o in str(proj["ods"] or "").split(";") if o in ODS]
    sel = st.multiselect("Objetivos de Desenvolvimento Sustentável (ODS) relacionados", ODS, default=ods_atual, key=wk("ods"))
    v["ods"] = ";".join(sel)
    auto_unica("projeto", v)
    if any(proj[k_] for k_, _, _ in ETAPAS_MUDANCA):
        st.markdown(" ➜ ".join(f"**{rot}**: {proj[k_]}" for k_, rot, _ in ETAPAS_MUDANCA if proj[k_]))
    st.subheader("Indicadores de impacto")
    editor("impacto_indicadores", {"indicador": T("Indicador", required=True, width="large"),
                                   "tipo": S("Tipo", options=TIPOS_INDICADOR_IMPACTO), "linha_base": T("Situação atual"),
                                   "meta": T("Meta"), "ods": S("ODS", options=sel or ODS)}, "indicador")
    proxima("impacto")


# ------------------------------------------------------------------------------
# EDITAIS DE FOMENTO
# ------------------------------------------------------------------------------
def pg_editais():
    cabecalho("editais")
    st.warning("Estrutura baseada nas seções usuais de editais de inovação, como o Programa Centelha. Cada edital tem campos e "
               "limites próprios: ajuste os limites de caracteres conforme o edital vigente antes de enviar.")
    res = R()
    trl = st.select_slider("Nível de maturidade tecnológica (TRL)", options=list(TRL), value=int(proj["edital_trl"] or 3),
                           format_func=lambda n: f"{n}: {TRL[n]}", key=wk("trl"))
    auto_unica("projeto", {"edital_trl": trl})
    ed = d["edital_textos"]
    if ed.empty or st.button("🔄 Refazer o rascunho a partir do plano (substitui os textos)", key="edt_refazer"):
        ras = rascunho_edital(proj, d, res)
        base = pd.DataFrame([{"secao": s_, "texto": ras.get(s_, ""), "limite": float(lim)} for s_, lim in SECOES_EDITAL])
        recarregar_tabela("edital_textos", base, "secao")
        st.rerun()
    novos = []
    mudou = False
    for i, r in ed.iterrows():
        lim = int(r["limite"] or 2000)
        with st.container(border=True):
            c1, c2 = st.columns([4, 1])
            c1.markdown(f"**{r['secao']}**")
            nl = c2.number_input("Limite", 50, 20000, lim, 50, key=wk(f"edl_{i}"), label_visibility="collapsed")
            txt = st.text_area(r["secao"], r["texto"], height=120, key=wk(f"edt_{i}"), label_visibility="collapsed")
            n = len(txt)
            st.caption(("🔴 " if n > nl else "🟢 ") + f"{n} de {nl} caracteres")
            novos.append({"secao": r["secao"], "texto": txt, "limite": float(nl)})
            mudou = mudou or txt != r["texto"] or nl != lim
    if mudou:
        df = pd.DataFrame(novos)
        salvar(conn, "edital_textos", df, "secao")
        d["edital_textos"] = limpar(df, "edital_textos")
        marcar_salvo()
    st.subheader("Orçamento do projeto")
    editor("edital_orcamento", {"rubrica": S("Rubrica", options=RUBRICAS_EDITAL, required=True), "item": T("Item", width="large"),
                                "valor": moeda("Valor (R$)")}, "item")
    orc = d["edital_orcamento"]
    if not orc.empty:
        st.dataframe(orc.groupby("rubrica", as_index=False)["valor"].sum().rename(columns={"rubrica": "Rubrica", "valor": "Total"}),
                     hide_index=True, column_config={"Total": N(format="R$ %.2f")})
        st.metric("Total solicitado", brl(orc["valor"].sum()))
    texto = "\n\n".join(f"{r['secao'].upper()}\n{r['texto']}" for _, r in d["edital_textos"].iterrows())
    if not orc.empty:
        texto += "\n\nORÇAMENTO\n" + "\n".join(f"{r['rubrica']} - {r['item']}: {brl(r['valor'])}" for _, r in orc.iterrows())
    st.download_button("⬇️ Baixar texto da proposta (.txt)", texto.encode("utf-8"), "proposta_edital.txt", "text/plain")
    proxima("editais")


# ------------------------------------------------------------------------------
# NOVIDADES DO PROGRAMA (histórico de versões do ViabilizaLab)
# ------------------------------------------------------------------------------
def pg_novidades():
    cabecalho("novidades")
    st.caption(f"Você está usando a versão **{VERSAO_PROGRAMA}**. O histórico completo também está no arquivo CHANGELOG.md do "
               "repositório no GitHub.")
    st.markdown(NOVIDADES_MD)


# ------------------------------------------------------------------------------
# COMENTÁRIOS DO PROFESSOR (aparecem no topo de cada ferramenta)
# ------------------------------------------------------------------------------
def comentarios_da_pagina(p: str):
    cm = d.get("comentarios")
    if cm is None:
        return
    da_pag = cm[(cm["pagina"] == p) & (cm["resolvido"].fillna(0).astype(int) == 0)]
    for i, r in da_pag.iterrows():
        with st.container(border=True):
            a, b = st.columns([5, 1])
            a.markdown(f"💬 **Comentário de {r['autor'] or 'professor(a)'}** ({r['data']}): {r['comentario']}")
            if b.button("Resolvido ✓", key=f"cm_ok_{p}_{i}"):
                cm.loc[i, "resolvido"] = 1
                salvar(conn, "comentarios", cm, "comentario")
                d["comentarios"] = limpar(cm.reset_index(drop=True), "comentarios")
                st.rerun()
    if ss.get("modo_correcao"):
        with st.expander("✍️ Escrever comentário nesta ferramenta (modo de correção)"):
            txt = st.text_area("Comentário", key=f"cm_txt_{p}")
            if st.button("Adicionar comentário", key=f"cm_add_{p}") and txt.strip():
                novo = pd.DataFrame([{"pagina": p, "comentario": txt.strip(), "autor": ss.get("prof_nome") or "Professor(a)",
                                      "data": datetime.now().strftime("%d/%m/%Y"), "resolvido": 0}])
                cm2 = pd.concat([cm, novo], ignore_index=True)
                salvar(conn, "comentarios", cm2, "comentario")
                d["comentarios"] = limpar(cm2, "comentarios")
                st.rerun()


# ------------------------------------------------------------------------------
# MODO PROFESSOR: SEMELHANÇA ENTRE PLANOS E RODADA DE INVESTIMENTO DA TURMA
# ------------------------------------------------------------------------------
def professor_extras(projetos: list):
    """projetos: lista de (rotulo, proj, d) já carregados no modo professor."""
    st.subheader("🧭 Prontidão e evidências da turma")
    linhas_ = []
    for rot, p_, d_ in projetos:
        try:
            r_ = executar_motor(p_, d_)
            pl = placar_evidencias(p_, d_, r_)
            mc = monte_carlo(r_, faixas_mc(d_), n=800) if r_["anos"][1]["rec_bruta"] > 0 else None
            pr = avaliar_prontidao(p_, d_, r_, mc, preco_referencia(p_, r_))
            linhas_.append({"Projeto": rot, "Prontidão (1-9)": pr["nivel"], "Placar de evidências": pl["score"],
                            "Entrevistas": len(d_["entrevistas"]), "Chance de VPL > 0": mc["p_vpl_pos"] if mc else float("nan"),
                            "Comentários abertos": int((d_["comentarios"]["resolvido"].fillna(0).astype(int) == 0).sum())})
        except Exception:
            continue
    if linhas_:
        st.dataframe(pd.DataFrame(linhas_), hide_index=True, width="stretch", column_config={
            "Placar de evidências": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "Chance de VPL > 0": N(format="percent")})
    if len(projetos) >= 2:
        st.subheader("🔁 Planos parecidos entre si")
        st.caption("Compara trechos de texto (sequências de três palavras) e valores numéricos. Semelhança alta pede uma conversa "
                   "com os grupos; não é prova de cópia.")
        sm = semelhanca_planos(projetos)
        st.dataframe(sm, hide_index=True, width="stretch", column_config={
            "Textos em comum": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "Números em comum": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1)})
        alto = sm[(sm["Textos em comum"] > 0.35) | (sm["Números em comum"] > 0.5)]
        if not alto.empty:
            st.warning(f"{len(alto)} par(es) com semelhança alta.")
        st.subheader("💰 Rodada de investimento da turma")
        st.caption("Cada grupo recebe um capital fictício e distribui entre os planos dos colegas (o próprio plano não conta). "
                   "Lance os valores combinados em sala.")
        cap = st.number_input("Capital fictício por grupo (R$)", 1000.0, 1e9, 100000.0, 10000.0, key="rod_cap")
        nomes = [r for r, _, _ in projetos]
        if ss.get("rodada_nomes") != nomes:
            ss.rodada_df = pd.DataFrame(0.0, index=nomes, columns=nomes)
            ss.rodada_nomes = nomes
        ed = st.data_editor(ss.rodada_df, width="stretch", key="ed_rodada",
                            column_config={c: N(format="R$ %.0f", min_value=0.0) for c in nomes})
        ss.rodada_df = ed
        excedeu = [i for i in ed.index if float(ed.loc[i].drop(i, errors="ignore").sum()) > cap]
        if excedeu:
            st.error("Passaram do capital disponível: " + ", ".join(excedeu))
        st.dataframe(resultado_rodada(ed), hide_index=True, width="stretch", column_config={"Capital recebido": N(format="R$ %.0f")})
