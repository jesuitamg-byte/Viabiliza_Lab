# ==============================================================================
# VIABILIZALAB - INTERFACE: PAINEL E FERRAMENTAS DE ESTRATÉGIA E MERCADO
# ==============================================================================

# ------------------------------------------------------------------------------
# Registros editados um a um (personas, propostas): salvamento automático por id
# ------------------------------------------------------------------------------
def ler_com_id(tabela):
    cols = [c for c, _ in SCHEMA[tabela]]
    return pd.read_sql(f"SELECT id, {', '.join(cols)} FROM {tabela} ORDER BY id", conn)


def auto_registro(tabela, rid, atual: dict, novos: dict):
    mudou = {k: v for k, v in novos.items() if str(atual.get(k) if atual.get(k) is not None else "") != str(v)}
    if mudou:
        conn.execute(f"UPDATE {tabela} SET {', '.join(k + '=?' for k in mudou)} WHERE id=?",
                     list(mudou.values()) + [rid])
        conn.commit()
        d[tabela] = ler(conn, tabela)
        marcar_salvo()
        invalidar()


def novo_registro(tabela, valores: dict):
    cols = list(valores)
    cur = conn.execute(f"INSERT INTO {tabela} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                       [valores[c] for c in cols])
    conn.commit()
    return cur.lastrowid


def apagar_registro(tabela, rid):
    conn.execute(f"DELETE FROM {tabela} WHERE id=?", (rid,))
    conn.commit()


def salvo_txt():
    st.caption(f"💾 Salvo automaticamente{' às ' + ss.ultimo_salvo if ss.ultimo_salvo else ''}")


# ==============================================================================
# PAINEL DO PLANO (tela inicial)
# ==============================================================================
def pg_inicio():
    if ss.get("_pag_vista") != "inicio":
        ss["_pag_vista"] = "inicio"
        rolar_para_topo()
    if not ss.get("bemvindo_ok"):
        with st.container(border=True):
            st.markdown(f"""
### 👋 Bem-vindo ao ViabilizaLab
Este programa ajuda você a construir um **plano de negócio completo**, do cliente aos números, e a descobrir
se a ideia é **viável**. Como usar:

1. **Clique nos cartões** abaixo, de cima para baixo. Cada cartão é uma ferramenta, com explicação de como preencher.
2. **Tudo é salvo automaticamente** enquanto você digita. Não existe botão "Salvar".
3. O selo do cartão mostra o andamento: ✅ concluído · ◐ em andamento · ○ não iniciado.
4. Os **resultados financeiros** (DRE, indicadores, gráficos) são calculados sozinhos a partir do que você preencher.
5. No final, vá em **📤 Relatório e exportação** para baixar o PDF, o pitch, o Word e o arquivo **.pnbox** do projeto.

{"Você está na **versão on-line**: os dados ficam neste navegador. Para entregar ou levar a outro computador, baixe o arquivo .pnbox." if MODO_WEB else ""}
""")
            if st.button("Entendi, começar", type="primary"):
                ss.bemvindo_ok = True
                st.rerun()

    st.title(f"📊 {proj['nome_negocio'] or 'Meu plano de negócio'}")
    linha = " · ".join(x for x in [proj["cidade_uf"], proj["turma"], proj["autores"]] if x)
    if linha:
        st.caption(linha)

    comp = completude()
    try:
        res = R()
        chk = CHK()
        sem = semaforo(res)
        rx = raio_x(res, chk, comp, sem)
    except Exception as e:
        res, chk, sem, rx = None, None, [], None
        st.info(f"Os resultados financeiros aparecem quando houver produtos, custos e investimentos. ({e})")

    c1, c2, c3, c4 = st.columns([2.2, 1, 1, 1])
    with c1:
        st.markdown(f"**Plano {int(comp * 100)}% preenchido**")
        st.progress(comp)
    if rx:
        c2.metric("Raio-X do plano", f"{rx['nota']}/100", help="Preenchimento (30) + checagens OK (30) + saúde financeira (40)")
        c3.metric("Investimento total", brl(res["investimento_total"]).split(",")[0],
                  help=brl(res["investimento_total"]))
        c4.metric("Lucro no Ano 1", brl(res["anos"][1]["resultado"]).split(",")[0],
                  help=brl(res["anos"][1]["resultado"]))
        verdes = sum(1 for s in sem if s["Status"] == "🟢")
        if res["anos"][1]["rec_bruta"] <= 0:
            st.caption("Saúde financeira: aparece quando houver produtos com preço e quantidade cadastrados.")
        else:
            st.caption(f"Saúde financeira: {' '.join(s['Status'] for s in sem)}  ({verdes} de {len(sem)} indicadores no verde)")

    if chk is not None:
        pend = chk[chk["Situação"] != "OK"]
        if not pend.empty:
            with st.expander(f"⚠️ {len(pend)} ponto(s) para corrigir — clique para ver e ir direto ao ajuste", expanded=False):
                for i, r in pend.iterrows():
                    a, b = st.columns([4, 1.3])
                    a.markdown(f"**{r['Teste']}** — {r['Valor apurado']}  \n<small>{r['Como corrigir'] or r['Observação']}</small>",
                               unsafe_allow_html=True)
                    if r["destino"] and b.button("Ir para o ajuste →", key=f"ini_chk_{i}", width="stretch"):
                        ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir'] or r['Observação']}")
                        st.rerun()

    for g in GRUPOS:
        itens = [p for p in PAGINAS if p[1] == g]
        st.markdown(f"<div class='pn-grupo'>{g}</div>", unsafe_allow_html=True)
        for i in range(0, len(itens), 4):
            cols = st.columns(4)
            for col, p in zip(cols, itens[i:i + 4]):
                selo = situacao(p[0])
                txt_selo = {"✅": "✅ concluído", "◐": "◐ em andamento", "○": "○ não iniciado"}.get(selo, "")
                with col.container(key=f"card_{p[0]}"):
                    st.markdown(f"<div class='pn-ic'>{p[2]}</div><div class='pn-tit'>{p[3]}</div>"
                                f"<div class='pn-desc'>{p[4]}</div><div class='pn-selo'>{txt_selo}</div>",
                                unsafe_allow_html=True)
                    if st.button(f"Abrir {p[3]}", key=f"cardbtn_{p[0]}", width="stretch"):
                        ir(p[0])
                        st.rerun()


# ==============================================================================
# IDENTIFICAÇÃO
# ==============================================================================
TIPOS_NEGOCIO = {
    "Produtos": "Fabrica ou revende mercadorias. Use Produtos e materiais, ficha técnica e estoques.",
    "Serviços (hora ou projeto)": "Vende tempo e conhecimento (consultoria, manutenção, design, desenvolvimento sob encomenda, "
                                  "estética...). Use também 🧑‍💻 Serviços e custo da hora.",
    "Assinatura (receita recorrente)": "Cobra planos que se renovam (aplicativos, software, clubes, academias, mensalidades). "
                                       "Use 💳 Assinaturas e receita recorrente.",
    "Misto": "Combina produtos, serviços e/ou assinaturas.",
}


def calculadora_simples():
    res = R()
    fr = fator_r(res, d)
    st.markdown("""
No Simples Nacional a alíquota depende do **anexo** da atividade e da **receita bruta dos últimos 12 meses (RBT12)**:
alíquota efetiva = (RBT12 × alíquota nominal − parcela a deduzir) ÷ RBT12. Serviços de tecnologia, engenharia e outros de
natureza intelectual ficam no **Anexo V**, mas passam ao **Anexo III**, mais barato, quando o **fator R** (folha de pagamento,
incluindo pró-labore, ÷ receita bruta) é de 28% ou mais. Confirme sempre o CNAE e o anexo com um contador.
""")
    c1, c2, c3 = st.columns(3)
    anexos = list(TABELAS_SIMPLES)
    sugerido = proj["anexo_simples"] if proj["anexo_simples"] in anexos else (
        "Anexo III - Serviços" if str(proj["tipo_negocio"]).startswith(("Serviço", "Assinatura")) else "Anexo I - Comércio")
    anexo = c1.selectbox("Anexo da atividade", anexos, index=anexos.index(sugerido), key=wk("anexo"))
    rbt = c2.number_input("Receita bruta de 12 meses (RBT12)", 0.0, 1e9, float(round(res["anos"][1]["rec_bruta"], 2)), 1000.0,
                          key=wk("rbt12"), help="Sugestão: a receita prevista para o Ano 1.")
    usar_r = anexo in ("Anexo III - Serviços", "Anexo V - Serviços")
    anexo_final = fr["anexo_por_fator_r"] if usar_r else anexo
    r = aliquota_simples(anexo_final, rbt)
    c3.metric("Fator R previsto", pct(fr["fator_r"]) if fr["fator_r"] == fr["fator_r"] else "n.a.",
              help=f"Folha de 12 meses {brl(fr['folha12'])} (salários, encargos e pró-labore de {brl(fr['pro_labore_m'])}/mês) "
                   f"÷ receita {brl(fr['receita12'])}")
    if usar_r:
        st.info(f"Com fator R de {pct(fr['fator_r'])}, a atividade é tributada pelo **{anexo_final}**.")
    if r["efetiva"] == r["efetiva"]:
        a, b, c = st.columns(3)
        a.metric("Faixa", "acima do limite" if r.get("acima_limite") else f"{r['faixa']}ª")
        b.metric("Alíquota nominal", pct(r["nominal"]))
        c.metric("Alíquota efetiva", pct(r["efetiva"]))
        if st.button(f"Usar {pct(r['efetiva'])} como imposto sobre vendas", key=wk("usar_aliq")):
            salvar_unica(conn, "projeto", {"aliquota_imposto_pct": round(r["efetiva"] * 100, 2), "anexo_simples": anexo})
            ss.pop(wk("aliq"), None)
            invalidar()
            marcar_salvo()
            st.rerun()
    st.caption("Tabelas da LC 123/2006 com a redação da LC 155/2016. A reforma tributária (EC 132/2023 e LC 214/2025) mantém o "
               "Simples, mas a partir de 2027 a empresa pode recolher IBS e CBS dentro da guia única ou por fora, no regime "
               "regular. Para o plano, use a regra vigente, cite a fonte e confirme com um contador.")


def pg_projeto():
    cabecalho("projeto")
    t1, t2, t3 = st.tabs(["🏢 Negócio", "👥 Grupo e entrega", "📈 Premissas financeiras"])
    v = {}
    with t1:
        tipos = list(TIPOS_NEGOCIO)
        atual_t = proj["tipo_negocio"] if proj["tipo_negocio"] in tipos else tipos[0]
        v["tipo_negocio"] = st.radio("Como o negócio ganha dinheiro?", tipos, index=tipos.index(atual_t), horizontal=True,
                                     key=wk("tipo_neg"), help="Define quais ferramentas o programa recomenda e quais checagens aplica.")
        st.caption(TIPOS_NEGOCIO[v["tipo_negocio"]])
        c1, c2, c3 = st.columns(3)
        v["nome_negocio"] = c1.text_input("Nome do negócio", proj["nome_negocio"], key=wk("nome"), max_chars=60)
        objs = ["Criar um novo negócio", "Expandir um negócio existente", "Captar recursos", "Outro"]
        v["objetivo"] = c2.selectbox("Objetivo do plano", objs,
                                     index=objs.index(proj["objetivo"]) if proj["objetivo"] in objs else 0, key=wk("obj"))
        v["periodo_anos"] = c3.number_input("Horizonte de planejamento (anos)", 1, 10, int(proj["periodo_anos"] or 3), key=wk("per"),
                                            help="Quantos anos o plano projeta: DRE, fluxos, VPL, TIR e payback usam este período. "
                                                 "Planos de bancos e do Sebrae costumam usar de 3 a 5 anos.")
        regs = ["Simples Nacional", "MEI", "Lucro Presumido", "Lucro Real"]
        v["enquadramento"] = c1.selectbox("Enquadramento tributário", regs, key=wk("enq"),
                                          index=regs.index(proj["enquadramento"]) if proj["enquadramento"] in regs else 0)
        portes = ["MEI", "Microempresa (ME)", "Pequena Empresa (EPP)", "Médio/Grande Porte"]
        v["porte"] = c2.selectbox("Porte", portes, key=wk("porte"),
                                  index=portes.index(proj["porte"]) if proj["porte"] in portes else 1)
        atuais = [x.strip() for x in proj["setor_atuacao"].split(";") if x.strip()]
        v["setor_atuacao"] = "; ".join(c3.multiselect("Setores de atuação", opcoes("setor", atuais), atuais, key=wk("setor"),
                                                      accept_new_options=True))
        v["cidade_uf"] = c1.text_input("Cidade / UF", proj["cidade_uf"], key=wk("cid"))
        v["contatos"] = c2.text_input("Contatos (site, redes, e-mail)", proj["contatos"], key=wk("cont"))
    with t2:
        st.caption("Estes dados aparecem na capa do relatório e no modo professor.")
        c1, c2 = st.columns(2)
        v["autores"] = c1.text_area("Integrantes do grupo (um por linha ou separados por ;)", proj["autores"], key=wk("aut"))
        v["turma"] = c2.text_input("Curso / disciplina / turma", proj["turma"], key=wk("turma"))
        v["professor"] = c2.text_input("Professor(a)", proj["professor"], key=wk("prof"))
        v["data_entrega"] = c2.text_input("Data de entrega", proj["data_entrega"], key=wk("data"), placeholder="dd/mm/aaaa")
    with t3:
        st.caption("Premissas usadas em todas as projeções. Passe o mouse no ícone ? para ver a explicação de cada uma.")
        p1, p2, p3 = st.columns(3)
        v["aliquota_imposto_pct"] = p1.number_input("Impostos sobre vendas (%)", 0.0, 40.0, proj["aliquota_imposto_pct"], 0.1, key=wk("aliq"),
                                                    help="Alíquota efetiva do Simples Nacional ou soma de PIS/COFINS/ISS/ICMS.")
        v["crescimento_mensal_ano1_pct"] = p2.number_input("Crescimento mensal das vendas no Ano 1 (%)", -10.0, 30.0,
                                                           proj["crescimento_mensal_ano1_pct"], 0.5, key=wk("gm"))
        v["crescimento_anual_pct"] = p3.number_input("Crescimento anual nos Anos 2 e 3 (%)", -50.0, 200.0,
                                                     proj["crescimento_anual_pct"], 1.0, key=wk("ga"))
        v["inflacao_custos_anual_pct"] = p1.number_input("Inflação dos custos fixos (% ao ano)", 0.0, 50.0,
                                                         proj["inflacao_custos_anual_pct"], 0.5, key=wk("inf"))
        v["tma_anual_pct"] = p2.number_input("TMA - taxa mínima de atratividade (% ao ano)", 0.0, 60.0, proj["tma_anual_pct"], 0.5,
                                             key=wk("tma"), help="Retorno mínimo que os sócios exigem (ex.: rendimento de uma aplicação segura + prêmio de risco).")
        v["pme_dias"] = p3.number_input("Prazo médio de estoques (dias)", 0.0, 365.0, proj["pme_dias"], 1.0, key=wk("pme"))
        v["reserva_caixa_extra"] = p1.number_input("Reserva financeira de segurança (R$)", 0.0, 1e8, proj["reserva_caixa_extra"], 1000.0, key=wk("res"))
        v["recuperar_giro"] = int(p2.checkbox("No VPL, recuperar o capital de giro no último ano", bool(proj["recuperar_giro"]), key=wk("rg")))
        with st.expander("🧮 Calcular a alíquota do Simples Nacional (anexos e fator R)"):
            calculadora_simples()
    auto_unica("projeto", v)
    salvo_txt()

    st.subheader("👥 Sócios")
    editor("socios", {"nome": T("Nome", required=True), "funcao": T("Função"), "contato": T("Contato"),
                      "perfil": T("Minicurrículo / atribuições", width="large"),
                      "capital_integralizado": moeda("Capital investido (R$)"),
                      "perda_aceitavel": moeda("Perda aceitável (R$)", help="Quanto este sócio aceita perder se o negócio não "
                                               "der certo, sem comprometer a vida pessoal (Sarasvathy, 2001).")}, "nome")
    st.caption(f"Capital próprio total: **{brl(d['socios']['capital_integralizado'].sum())}** · perda aceitável somada: "
               f"**{brl(d['socios']['perda_aceitavel'].sum())}** (usada em Monte Carlo e perda aceitável)")
    proxima("projeto")


# ==============================================================================
# CLIENTE E MERCADO
# ==============================================================================
def pg_segmentos():
    cabecalho("segmentos")
    t1, t2 = st.tabs(["🎯 Segmentos de clientes", "🏆 Priorização"])
    with t1:
        st.caption("Escreva até 4 a 6 grupos de clientes possíveis e dê notas de 1 a 10 para cada pergunta.")
        sg = editor("segmentos", {
            "nome": T("Segmento (máx. 40 caracteres)", required=True, max_chars=40),
            "descricao": T("Descrição do perfil", width="large"),
            "nota_tamanho": nota("É grande o suficiente?"), "nota_lucro": nota("Tem potencial de lucro?"),
            "nota_longevidade": nota("Não tende a desaparecer?"), "nota_concorrencia": nota("Tem poucos concorrentes?"),
            "nota_aderencia": nota("Seus produtos atendem esse segmento?")}, "nome")
    with t2:
        sg = d["segmentos"]
        if sg.empty:
            st.info("Cadastre os segmentos na primeira aba.")
        else:
            rk = sg.assign(Média=sg.iloc[:, 2:7].mean(axis=1)).sort_values("Média", ascending=False).reset_index(drop=True)
            for i, r in rk.iterrows():
                with st.container(border=True):
                    a, b = st.columns([5, 1])
                    a.markdown(f"### {i + 1}º  {r['nome']}\n{r['descricao']}")
                    b.metric("Média", num_br(r["Média"], 1))
            st.altair_chart(alt.Chart(rk).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("nome:N", sort="-x", title=None)),
                width="stretch")
    proxima("segmentos")


CAMPOS_PERSONA = {
    "Dados": [("nome", "Nome", "txt", 40), ("idade", "Idade (anos)", "int", 0), ("estado_civil", "Estado civil", "txt", 30),
              ("escolaridade", "Escolaridade", "txt", 40), ("profissao", "Profissão", "txt", 60),
              ("renda_mensal", "Renda mensal (R$)", "num", 0), ("cidade", "Cidade", "txt", 60),
              ("descricao", "Sobre a persona (máx. 380 caracteres)", "area", 380)],
    "Características": [("citacao", "Frase que representa a persona", "area", 250),
                        ("habitos", "Hábitos e comportamentos", "area", 500),
                        ("caracteristicas", "Características (#tags)", "txt", 200),
                        ("fluencia_digital", "Fluência digital (computador, smartphone, apps, compras on-line)", "txt", 200),
                        ("redes_sociais", "Redes sociais utilizadas", "txt", 150),
                        ("quem_influencia", "Quem influencia as decisões", "txt", 200)],
    "Rotina": [("rotina_manha", "Manhã", "area", 400), ("rotina_tarde", "Tarde", "area", 400), ("rotina_noite", "Noite", "area", 400)],
    "Insights": [("ganhos", "Ganhos esperados", "area", 500), ("desafios", "Dores e desafios", "area", 500),
                 ("duvidas_objecoes", "Dúvidas ou objeções", "area", 500),
                 ("aborrecimentos", "Aborrecimentos que devem ser evitados", "area", 500),
                 ("experiencia_ideal", "Experiência ideal e como o negócio ajuda", "area", 500)],
}


def pg_personas():
    cabecalho("personas")
    pers = ler_com_id("personas")
    c1, c2 = st.columns([4, 1.2])
    nomes = [f"{r['nome']}" for _, r in pers.iterrows()]
    if c2.button("➕ Nova persona", type="primary", width="stretch"):
        rid = novo_registro("personas", {"nome": f"Persona {len(pers) + 1}", "idade": 30})
        ss[wk("persona_sel")] = rid
        st.rerun()
    if pers.empty:
        st.info("Nenhuma persona cadastrada. Clique em **➕ Nova persona**.")
        proxima("personas")
        return
    ids = pers["id"].tolist()
    atual_id = ss.get(wk("persona_sel"), ids[0])
    atual_id = atual_id if atual_id in ids else ids[0]
    escolhido = c1.radio("Persona", ids, index=ids.index(atual_id), horizontal=True,
                         format_func=lambda i: "🧑 " + str(pers.loc[pers["id"] == i, "nome"].iloc[0]))
    ss[wk("persona_sel")] = escolhido
    row = pers[pers["id"] == escolhido].iloc[0].to_dict()

    st.markdown(f"<div style='text-align:center;font-size:4rem;line-height:1'>🧑</div>"
                f"<h3 style='text-align:center;margin-top:0'>{row['nome']}</h3>", unsafe_allow_html=True)
    novos = {}
    abas_p = st.tabs(list(CAMPOS_PERSONA))
    for aba, (grupo, campos) in zip(abas_p, CAMPOS_PERSONA.items()):
        with aba:
            cols = st.columns(3 if grupo == "Dados" else (3 if grupo == "Rotina" else 2))
            for i, (k, rot, tipo, mx) in enumerate(campos):
                alvo = cols[i % len(cols)] if tipo != "area" or grupo in ("Rotina", "Insights", "Características") else st
                val = row.get(k)
                chave_w = wk(f"pers_{escolhido}_{k}")
                if tipo == "int":
                    novos[k] = alvo.number_input(rot, 0, 120, int(val or 0), key=chave_w)
                elif tipo == "num":
                    novos[k] = alvo.number_input(rot, 0.0, 1e7, float(val or 0), 100.0, key=chave_w)
                elif tipo == "area":
                    novos[k] = alvo.text_area(rot, val or "", max_chars=mx, key=chave_w, height=110)
                else:
                    novos[k] = alvo.text_input(rot, val or "", max_chars=mx, key=chave_w)
    auto_registro("personas", escolhido, row, novos)
    a, b = st.columns([4, 1.2])
    with a:
        salvo_txt()
    if b.button("🗑️ Excluir esta persona", width="stretch"):
        apagar_registro("personas", escolhido)
        ss.pop(wk("persona_sel"), None)
        st.rerun()
    proxima("personas")


def pg_jornada():
    cabecalho("jornada")
    pers = d["personas"]["nome"].tolist()
    if not pers:
        st.info("Cadastre uma persona antes de montar a jornada.")
        if st.button("Ir para Personas"):
            ir("personas")
            st.rerun()
        return
    persona = st.radio("Persona", pers, horizontal=True, format_func=lambda n: "🧑 " + n, key=wk("jor_pers"))
    perguntas = {"Antes": "O que o cliente faz antes de entrar em contato com o seu produto ou serviço?",
                 "Durante": "Como é a compra, a entrega e o uso?",
                 "Depois": "O que acontece depois: avaliação, suporte e fidelização?"}
    fases = st.tabs(list(ETAPAS_FASE))
    for aba, fase in zip(fases, ETAPAS_FASE):
        with aba:
            st.markdown(f"**{perguntas[fase]}**")
            jc = d["jornada_cliente"]
            sub = jc[(jc["persona_nome"] == persona) & (jc["macro_fase"] == fase)]
            for etapa in ETAPAS_FASE[fase]:
                linhas = sub[sub["etapa"] == etapa]
                st.markdown(f"<div style='background:rgba(47,111,208,.08);padding:6px 12px;border-radius:6px;"
                            f"font-weight:600;margin-top:10px'>{etapa}</div>", unsafe_allow_html=True)
                if linhas.empty:
                    st.caption("— sem ações cadastradas nesta etapa —")
                for _, r in linhas.iterrows():
                    face = EMOCOES.get(r["estado_emocional"], "😐")
                    st.markdown(
                        f"<div class='pn-etapa'><span class='pn-face'>{face}</span> <b>{r['acoes_cliente']}</b><br>"
                        f"<small>💭 {r['pensamentos'] or '-'} · 📍 {r['pontos_contato'] or '-'}<br>"
                        f"💡 <i>Oportunidade:</i> {r['oportunidades'] or '-'}</small></div>", unsafe_allow_html=True)
            with st.expander(f"✏️ Editar as ações da fase “{fase}”", expanded=sub.empty):
                editor("jornada_cliente", {
                    "etapa": S("Etapa", options=ETAPAS_FASE[fase], required=True),
                    "acoes_cliente": T("Ação do cliente", width="medium", required=True),
                    "pensamentos": T("O que pensa"), "pontos_contato": T("Pontos de contato"),
                    "estado_emocional": S("Emoção", options=list(EMOCOES)),
                    "oportunidades": T("Oportunidade de melhoria", width="medium")},
                    "acoes_cliente", filtro={"persona_nome": persona, "macro_fase": fase})
    jc = d["jornada_cliente"]
    jp = jc[jc["persona_nome"] == persona]
    if not jp.empty:
        st.subheader("📉 Curva emocional da jornada")
        emo = {"Entusiasmado": 2, "Satisfeito": 1, "Neutro": 0, "Insatisfeito": -1, "Frustrado": -2}
        g = jp.assign(Emoção=jp["estado_emocional"].map(emo).fillna(0)).groupby("etapa", sort=False)["Emoção"].mean()
        g = g.reindex([e for e in ETAPAS if e in g.index]).reset_index()
        st.altair_chart(alt.Chart(g).mark_line(point=True, strokeWidth=3, color="#d6006e").encode(
            x=alt.X("etapa:N", sort=ETAPAS, title=None, axis=alt.Axis(labelAngle=0, labelLimit=120)),
            y=alt.Y("Emoção:Q", scale=alt.Scale(domain=[-2, 2]), title="😠  ←  emoção  →  🤩")).properties(height=260),
            width="stretch")
    proxima("jornada")


def pg_mercado():
    cabecalho("mercado")
    st.markdown("""
**Como estimar:** **TAM** é todo o mercado (todas as pessoas que poderiam comprar algo desse tipo na região).
**SAM** é a parte que o seu negócio consegue alcançar (pelo segmento escolhido, raio de entrega, canais).
**SOM** é a fatia que você realisticamente conquista (participação de mercado). Use dados do IBGE, Sebrae,
prefeitura, associações e pesquisa de campo, e anote as fontes.""")
    v = {}
    c1, c2 = st.columns([3, 1])
    v["tam_desc"] = c1.text_input("TAM - mercado total (quem é)", proj["tam_desc"], key=wk("tamd"))
    v["tam_clientes"] = c2.number_input("Nº de clientes potenciais (TAM)", 0.0, 1e9, float(proj["tam_clientes"]), 100.0, key=wk("tamc"))
    v["sam_desc"] = c1.text_input("SAM - mercado alcançável (quem é)", proj["sam_desc"], key=wk("samd"))
    v["sam_clientes"] = c2.number_input("Nº de clientes alcançáveis (SAM)", 0.0, 1e9, float(proj["sam_clientes"]), 100.0, key=wk("samc"))
    c3, c4 = st.columns(2)
    v["som_pct"] = c3.number_input("SOM - participação que pretende conquistar do SAM (%)", 0.0, 100.0, float(proj["som_pct"]), 0.5, key=wk("som"))
    v["gasto_anual_cliente"] = c4.number_input("Gasto médio anual de cada cliente com esse tipo de produto (R$)", 0.0, 1e7,
                                               float(proj["gasto_anual_cliente"]), 50.0, key=wk("gasto"))
    v["mercado_fontes"] = st.text_area("Fontes e premissas usadas", proj["mercado_fontes"], key=wk("mfont"))
    auto_unica("projeto", v)
    salvo_txt()
    mk = mercado(proj, R())
    if mk["preenchido"]:
        df = pd.DataFrame({"Nível": ["TAM", "SAM", "SOM"], "Clientes": [mk["tam_c"], mk["sam_c"], mk["som_c"]],
                           "Mercado (R$/ano)": [mk["tam"], mk["sam"], mk["som"]]})
        a, b = st.columns([1.3, 1])
        a.altair_chart(alt.Chart(df).mark_bar(color="#2f6fd0").encode(
            y=alt.Y("Nível:N", sort=["TAM", "SAM", "SOM"], title=None), x=alt.X("Mercado (R$/ano):Q", title="R$ por ano"),
            tooltip=["Nível", alt.Tooltip("Clientes:Q", format=",.0f"), alt.Tooltip("Mercado (R$/ano):Q", format=",.2f")]
        ).properties(height=200), width="stretch")
        b.metric("TAM", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
        b.metric("SAM", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
        b.metric("SOM", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
        st.subheader("A meta de vendas é realista?")
        p = mk["pct_som"]
        msg = (f"A receita prevista no Ano 1 ({brl(mk['receita_ano1'])}) corresponde a **{pct(p)} do SOM** "
               f"e a **{pct(mk['pct_sam'])} do SAM**.")
        if p > 1:
            st.error("🔴 " + msg + " A meta passa do que o negócio consegue conquistar: reveja as quantidades vendidas "
                     "(Produtos e materiais) ou a estimativa de mercado.")
        elif p > 0.8:
            st.warning("🟡 " + msg + " A meta é ambiciosa: depende de conquistar quase todo o mercado alcançável previsto.")
        else:
            st.success("🟢 " + msg + " A meta cabe no mercado estimado.")
    else:
        st.info("Preencha os quatro números para ver o tamanho do mercado e conferir a meta de vendas.")
    proxima("mercado")


# ==============================================================================
# PROBLEMA E SOLUÇÃO
# ==============================================================================
def pg_proposta():
    cabecalho("proposta")
    props = ler_com_id("propostas")
    pers = d["personas"]["nome"].tolist()
    c1, c2 = st.columns([4, 1.3])
    c1.caption("Uma proposta por persona: escolha os valores (etiquetas) e complete as frases.")
    if c2.button("➕ Nova proposta", type="primary", width="stretch"):
        novo_registro("propostas", {"persona_nome": pers[0] if pers else "", "valores": ""})
        st.rerun()
    if props.empty:
        st.info("Nenhuma proposta cadastrada. Clique em **➕ Nova proposta**.")
    for n, (_, r) in enumerate(props.iterrows(), start=1):
        row = r.to_dict()
        with st.expander(f"❤️ Proposta {n} — {row['persona_nome'] or 'sem persona'}:  {row['valores'][:90]}",
                         expanded=len(props) == 1):
            novos = {}
            opc_p = pers + ([row["persona_nome"]] if row["persona_nome"] and row["persona_nome"] not in pers else [])
            novos["persona_nome"] = st.selectbox("Persona", opc_p or [""], key=wk(f"pp_{row['id']}"),
                                                 index=opc_p.index(row["persona_nome"]) if row["persona_nome"] in opc_p else 0)
            atuais = [x.strip() for x in str(row["valores"]).split(";") if x.strip()]
            novos["valores"] = "; ".join(st.multiselect("Valores entregues (escolha ou digite novos)", opcoes("valor_proposta", atuais),
                                                        atuais, key=wk(f"pv_{row['id']}"), accept_new_options=True))
            a, b = st.columns(2)
            novos["que_desejam"] = a.text_area("Que desejam...", row["que_desejam"], key=wk(f"pq_{row['id']}"))
            novos["por_meio"] = b.text_area("Por meio de...", row["por_meio"], key=wk(f"pm_{row['id']}"))
            novos["diferentemente"] = a.text_area("Diferentemente de...", row["diferentemente"], key=wk(f"pd_{row['id']}"))
            novos["oferecem"] = b.text_area("Os nossos produtos ou serviços oferecem...", row["oferecem"], key=wk(f"po_{row['id']}"))
            auto_registro("propostas", row["id"], row, novos)
            st.info(f"**{proj['nome_negocio']}** ajuda **{novos['persona_nome']}**, que deseja **{novos['que_desejam']}**, "
                    f"por meio de **{novos['por_meio']}**. Diferentemente de **{novos['diferentemente']}**, "
                    f"oferecemos **{novos['oferecem']}**.")
            if st.button("🗑️ Excluir esta proposta", key=wk(f"pdel_{row['id']}")):
                apagar_registro("propostas", row["id"])
                st.rerun()
    salvo_txt()
    proxima("proposta")


def pivo_concorrencia():
    empresas = ["Nossa Empresa"] + [n for n in d["concorrentes"]["nome"] if str(n).strip()]
    notas = d["concorrencia_notas"]
    atributos = list(dict.fromkeys(notas["atributo"]))
    piv = pd.DataFrame({"Atributo": pd.Series(atributos, dtype=object)})
    for emp in empresas:
        mapa = notas[notas["empresa"] == emp].set_index("atributo")["nota"].to_dict()
        piv[emp] = pd.Series([mapa.get(a, None) for a in atributos], dtype=float)
    return piv, empresas


def pg_concorrencia():
    cabecalho("concorrencia")
    st.subheader("1. Concorrentes")
    editor("concorrentes", {"nome": T("Nome do concorrente", required=True),
                            "observacoes": T("Observações (onde fica, preço, pontos fortes)", width="large")}, "nome")
    piv, empresas = pivo_concorrencia()
    st.subheader("2. Notas por atributo (0 a 10)")

    def _norm(x):
        x = x.copy()
        x["Atributo"] = x["Atributo"].astype(object).where(x["Atributo"].notna(), "")
        for e in empresas:
            x[e] = pd.to_numeric(x[e], errors="coerce") if e in x else float("nan")
        return x[["Atributo"] + empresas]

    def _salvar(x):
        x = x[x["Atributo"].astype(str).str.strip() != ""]
        longo = x.melt(id_vars="Atributo", var_name="empresa", value_name="nota").rename(columns={"Atributo": "atributo"})
        salvar(conn, "concorrencia_notas", longo, "atributo")
        d["concorrencia_notas"] = ler(conn, "concorrencia_notas")

    cfg = {"Atributo": T("Atributo avaliado", required=True)}
    cfg.update({e: N(e, min_value=0, max_value=10, step=0.5) for e in empresas})
    pe = tabela_editavel(piv, cfg, "concorrencia_notas", "|".join(empresas), normalizar=_norm, salvar_fn=_salvar)
    pe = pe[pe["Atributo"].astype(str).str.strip() != ""]
    num = pe.copy()
    for e in empresas:
        num[e] = pd.to_numeric(num[e], errors="coerce").fillna(0)
    if not num.empty:
        st.subheader("3. Resultado")
        medias = num[empresas].mean()
        cols = st.columns(min(len(empresas), 5))
        for i, e in enumerate(empresas[:5]):
            cols[i].metric(e, num_br(medias[e], 1))
        longo = num.melt(id_vars="Atributo", var_name="Empresa", value_name="Nota")
        st.altair_chart(alt.Chart(longo).mark_line(point=True).encode(
            x=alt.X("Atributo:N", title=None, axis=alt.Axis(labelAngle=-35, labelLimit=160)),
            y=alt.Y("Nota:Q", scale=alt.Scale(domain=[0, 10])), color="Empresa:N").properties(height=340), width="stretch")
        if len(empresas) > 1:
            idx = num.set_index("Atributo")
            outros = idx[empresas[1:]].max(axis=1)
            st.success("**Forças frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] > outros]) or "nenhuma"))
            st.warning("**Fraquezas frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] < outros]) or "nenhuma"))
            dest = idx["Nossa Empresa"] - outros
            st.info(f"**Destaque do negócio:** {dest.idxmax()} (vantagem de {num_br(dest.max(), 1)} ponto(s) sobre o melhor concorrente).")
    proxima("concorrencia")


# ==============================================================================
# ESTRATÉGIA
# ==============================================================================
def _fatores(ambiente, naturezas, dims, rot_acao):
    editor("swot", {
        "dimensao": S("Dimensão", options=dims),
        "fator": T("Fator", required=True),
        "natureza": S("Classificação", options=naturezas, required=True),
        "impacto": S("Impacto", options=["Alto", "Baixo"], required=True),
        "diagnostico": T("Diagnóstico (por que é assim no seu negócio)", width="medium"),
        "plano_acao": T(rot_acao, width="medium"),
        "priorizado": CK("⭐ Priorizar")}, "fator", filtro={"ambiente": ambiente})


def pg_forcas():
    cabecalho("forcas")
    st.caption("Exemplos de fatores: qualificação da equipe, presença digital, controle de qualidade, controles financeiros, "
               "localização, relação com fornecedores, reputação da marca.")
    dims = opcoes("dimensao_interna", d["swot"].loc[d["swot"]["ambiente"] == "Interno", "dimensao"].tolist())
    _fatores("Interno", ["Força", "Fraqueza"], dims, "Como potencializar (força) ou corrigir (fraqueza)")
    proxima("forcas")


def pg_oportunidades():
    cabecalho("oportunidades")
    st.caption("Exemplos: preocupação com saúde, novas tecnologias de pagamento, inflação, renda dos consumidores, "
               "fiscalização, novos concorrentes, mudanças demográficas.")
    dims = opcoes("dimensao_externa", d["swot"].loc[d["swot"]["ambiente"] == "Externo", "dimensao"].tolist())
    _fatores("Externo", ["Oportunidade", "Ameaça"], dims, "Como aproveitar (oportunidade) ou neutralizar (ameaça)")
    proxima("oportunidades")


def html_swot(sw: pd.DataFrame) -> str:
    def lista(nat):
        itens = sw[sw["natureza"] == nat]
        if itens.empty:
            return "<i>nenhum fator</i>"
        return "<br>".join(("⭐ " if bool(r["priorizado"]) else "• ") + f"{r['fator']} <small>({r['impacto']})</small>"
                           for _, r in itens.iterrows())
    cel = "padding:14px 16px;border-radius:4px;min-height:200px;font-size:14px;line-height:1.55;"
    lat = ("writing-mode:vertical-rl;transform:rotate(180deg);font-weight:700;letter-spacing:2px;"
           "display:flex;align-items:center;justify-content:center;")
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


def pg_swot():
    cabecalho("swot")
    sw = d["swot"]
    c1, c2 = st.columns(2)
    if c1.button("✏️ Editar forças e fraquezas", width="stretch"):
        ir("forcas")
        st.rerun()
    if c2.button("✏️ Editar oportunidades e ameaças", width="stretch"):
        ir("oportunidades")
        st.rerun()
    st.markdown(html_swot(sw), unsafe_allow_html=True)
    st.caption("⭐ fator priorizado · entre parênteses, o impacto")
    diag = pontuacao_swot(sw)
    st.subheader("Matriz de estratégias")
    q = diag["quadrantes"]
    g1, g2 = st.columns(2)
    for col, k in zip([g1, g2, g1, g2], ["FO", "FrO", "FA", "FrA"]):
        nome_q, comb, _ = QUADRANTES[k]
        destaque = "border:3px solid #2f6fd0;" if k == diag["predominante"] else "border:1px solid #ccc;"
        col.markdown(f"<div style='{destaque}padding:12px;border-radius:8px;margin-bottom:8px;text-align:center;'>"
                     f"<div style='font-size:28px;font-weight:700;'>{q[k]}</div><b>{comb}</b><br>{nome_q}</div>",
                     unsafe_allow_html=True)
    if diag["predominante"]:
        nome_q, comb, txt = QUADRANTES[diag["predominante"]]
        st.success(f"**Situação do negócio: {nome_q}.** {txt}")
    pri = sw[sw["priorizado"].astype(bool)]
    if not pri.empty:
        st.subheader("⭐ Ações estratégicas dos fatores priorizados")
        verbo = {"Força": "Como potencializar", "Fraqueza": "Como corrigir", "Oportunidade": "Como aproveitar", "Ameaça": "Como neutralizar"}
        for _, r in pri.iterrows():
            st.markdown(f"**{r['natureza']} · {r['fator']}** ({r['impacto']} impacto)  \n{r['diagnostico']}  \n"
                        f"➡️ *{verbo.get(r['natureza'], 'Ação')}:* {r['plano_acao']}")
    st.caption("Pontuação: impacto alto = 2 pontos; impacto baixo = 1 ponto.")
    proxima("swot")


# ==============================================================================
# MARKETING E OPERAÇÃO
# ==============================================================================
def pg_marketing():
    cabecalho("marketing")
    v = {}
    c1, c2 = st.columns(2)
    v["mkt_produto"] = c1.text_area("📦 Produto: o que será vendido, linhas, diferenciais, embalagem, marca",
                                    proj["mkt_produto"], height=150, key=wk("m1"))
    v["mkt_preco"] = c2.text_area("🏷️ Preço: estratégia (popular, intermediário, premium), comparação com concorrentes, descontos",
                                  proj["mkt_preco"], height=150, key=wk("m2"))
    v["mkt_praca"] = c1.text_area("📍 Praça: onde e como o cliente compra (loja, delivery, site, apps, revendedores)",
                                  proj["mkt_praca"], height=150, key=wk("m3"))
    v["mkt_promocao"] = c2.text_area("📣 Promoção: como o cliente vai conhecer o negócio (redes, anúncios, parcerias, eventos)",
                                     proj["mkt_promocao"], height=150, key=wk("m4"))
    auto_unica("projeto", v)
    salvo_txt()
    try:
        cat = R()["cat"]
        if not cat.empty:
            st.subheader("Apoio para a decisão de preço")
            a, b, c = st.columns(3)
            a.metric("Ticket médio (Mês 1)", brl(div(cat["receita_mensal"].sum(), cat["qtd_mensal"].sum())))
            b.metric("Margem média sobre o preço", pct(div(cat["margem_unit"].mul(cat["qtd_mensal"]).sum(), cat["receita_mensal"].sum())))
            c.metric("Itens no catálogo", num_br(len(cat), 0))
            if st.button("🏷️ Ver formação de preço por produto"):
                ir("preco")
                st.rerun()
    except Exception:
        pass
    proxima("marketing")


def pg_canais():
    cabecalho("canais")
    t1, t2 = st.tabs(["🛒 Canais de aquisição", "🔻 Funil de vendas"])
    with t1:
        st.caption("Dê notas de 0 a 10. Os três canais com maior média são os **canais de tração**.")
        cn = editor("canais", {"canal": S("Canal", options=opcoes("canal", d["canais"]["canal"].tolist()), required=True),
                               "atratividade": nota("Atratividade"), "alcance": nota("Alcance esperado"),
                               "cliente": nota("Custo por cliente (10 = barato)"), "relacionamento": nota("Relacionamento"),
                               "viabilidade": nota("Viabilidade")}, "canal")
        if not cn.empty:
            cn2 = cn.assign(Média=cn.iloc[:, 1:6].mean(axis=1)).sort_values("Média", ascending=False)
            st.success("**Canais de tração:** " + ", ".join(cn2["canal"].head(3)))
            st.altair_chart(alt.Chart(cn2).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("canal:N", sort="-x", title=None)), width="stretch")
    with t2:
        fu = editor("funil_vendas", {
            "etapa": S("Etapa", options=["Topo do Funil", "Meio do Funil", "Fundo do Funil"], required=True),
            "acao": T("Ação de venda", width="large"), "orcamento": moeda("Orçamento (R$)"),
            "alcancados": N("Pessoas alcançadas", min_value=0, format="%d"),
            "convertidos": N("Pessoas que avançaram / compraram", min_value=0, format="%d")}, "etapa")
        if not fu.empty:
            st.dataframe(tabela_funil(fu), hide_index=True, width="stretch")
            fundo = fu[fu["etapa"] == "Fundo do Funil"]["convertidos"].sum()
            m1, m2, m3 = st.columns(3)
            m1.metric("Orçamento total", brl(fu["orcamento"].sum()))
            m2.metric("Clientes conquistados", num_br(fundo, 0))
            m3.metric("CAC - custo de aquisição por cliente", brl(div(fu["orcamento"].sum(), fundo)))
            agg = fu.groupby("etapa", sort=False)[["alcancados", "convertidos"]].sum().reindex(
                ["Topo do Funil", "Meio do Funil", "Fundo do Funil"]).dropna().reset_index()
            st.altair_chart(alt.Chart(agg).mark_bar(color="#d6006e").encode(
                y=alt.Y("etapa:N", sort=None, title=None), x=alt.X("alcancados:Q", title="Pessoas alcançadas")), width="stretch")
    proxima("canais")


def tabela_funil(fu: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "Etapa": fu["etapa"], "Ação": fu["acao"], "Orçamento": fu["orcamento"].map(brl),
        "Alcançados": fu["alcancados"].map(lambda v: num_br(v, 0)),
        "Convertidos": fu["convertidos"].map(lambda v: num_br(v, 0)),
        "Conversão": [pct(div(c, a)) for c, a in zip(fu["convertidos"], fu["alcancados"])],
        "Custo por conversão": [brl(div(o, c)) for o, c in zip(fu["orcamento"], fu["convertidos"])]})


def dot_organograma() -> str:
    def q(s):
        return '"' + str(s).replace('"', "'").replace("\n", "\\n") + '"'
    linhas = ['digraph G {', 'rankdir=TB; nodesep=0.35; ranksep=0.5;',
              'node [shape=box, style="rounded,filled", fillcolor="#e8f0fb", fontname="Helvetica", fontsize=13, margin="0.25,0.12"];',
              'edge [color="#7a8aa0"];', f'raiz [label={q(proj["nome_negocio"] or "Empresa")}, fillcolor="#2f6fd0", fontcolor="white", fontsize=16];']
    soc = d["socios"]
    if not soc.empty:
        txt = "Sócios\n" + "\n".join(f"{r['nome']} ({r['funcao'] or 'sócio'})" for _, r in soc.iterrows())
        linhas.append(f'socios [label={q(txt)}, fillcolor="#cfe0f7"];')
        linhas.append("socios -> raiz [dir=none, style=dashed];")
        linhas.append("{rank=min; socios}")
    pes = d["pessoal"]
    for j, area in enumerate(dict.fromkeys(pes["area"].tolist())):
        if not area:
            continue
        sub = pes[pes["area"] == area]
        cargos = "\n".join(f"{r['cargo']} ({num_br(r['num_empregados'], 0)})" for _, r in sub.iterrows())
        total = num_br(sub["num_empregados"].sum(), 0)
        linhas.append(f'a{j} [label={q(area + " - " + total + " pessoa(s)")}, fillcolor="#fde3ef", fontsize=14];')
        linhas.append(f'c{j} [label={q(cargos)}, fillcolor="#ffffff"];')
        linhas.append(f'raiz -> a{j}; a{j} -> c{j};')
    linhas.append("}")
    return "\n".join(linhas)


def pg_operacional():
    cabecalho("operacional")
    t1, t2, t3, t4 = st.tabs(["📍 Localização", "🏗️ Layout e processos", "⚙️ Capacidade x demanda", "🗂️ Organograma"])
    v = {}
    with t1:
        c1, c2 = st.columns([3, 1])
        v["loc_endereco"] = c1.text_input("Endereço / região", proj["loc_endereco"], key=wk("le"))
        v["loc_area_m2"] = c2.number_input("Área (m²)", 0.0, 1e6, float(proj["loc_area_m2"]), 5.0, key=wk("la"))
        v["loc_justificativa"] = st.text_area("Por que este ponto? (fluxo, acesso, concorrentes, custo, segurança)",
                                              proj["loc_justificativa"], key=wk("lj"))
        v["loc_fluxo"] = st.text_area("Fluxo de pessoas e horários de pico", proj["loc_fluxo"], key=wk("lf"))
        v["op_horario"] = st.text_input("Horário de funcionamento", proj["op_horario"], key=wk("oh"))
    with t2:
        v["op_layout"] = st.text_area("Layout (arranjo físico): como os espaços serão organizados", proj["op_layout"], height=130, key=wk("ol"))
        v["op_processos"] = st.text_area("Processos operacionais: passo a passo da compra à entrega (use →)",
                                         proj["op_processos"], height=130, key=wk("op"))
        etapas = [e.strip() for e in re.split(r"->|→|\n", proj["op_processos"] or "") if e.strip()]
        if len(etapas) >= 2:
            dot = "digraph P { rankdir=LR; node [shape=box, style=\"rounded,filled\", fillcolor=\"#e8f0fb\", fontname=Helvetica, fontsize=10];\n"
            dot += " -> ".join(f'"{e[:40]}"' for e in etapas) + ";}"
            st.graphviz_chart(dot, width="stretch")
    with t3:
        v["op_capacidade_txt"] = st.text_area("Como a capacidade foi estimada", proj["op_capacidade_txt"], key=wk("oc"))
        alvos = ["Todos os produtos", "Fabricação Própria", "Serviço", "Revenda"] + d["catalogo"]["descricao"].tolist()
        editor("capacidade", {"recurso": T("Recurso (equipe, máquina, espaço)", required=True, width="medium"),
                              "produto_ou_grupo": S("Atende", options=alvos, required=True),
                              "capacidade_mensal": N("Capacidade por mês (unidades)", min_value=0.0),
                              "observacao": T("Observação", width="medium")}, "recurso")
        cd = capacidade_demanda(d, R())
        if not cd.empty:
            show = cd.copy()
            for c in ("Uso Mês 1", "Uso Mês 12"):
                show[c] = show[c].map(lambda x: ("🔴 " if x > 1 else "🟡 " if x > 0.85 else "🟢 ") + pct(x))
            st.dataframe(show, hide_index=True, width="stretch")
            g = cd.melt(id_vars="Recurso", value_vars=["Uso Mês 1", "Uso Mês 12"], var_name="Momento", value_name="Uso")
            st.altair_chart((alt.Chart(g).mark_bar().encode(x=alt.X("Uso:Q", axis=alt.Axis(format="%")), y=alt.Y("Recurso:N", title=None),
                                                           color="Momento:N", yOffset="Momento:N")
                             + alt.Chart(pd.DataFrame({"x": [1]})).mark_rule(color="red", strokeDash=[5, 3]).encode(x="x:Q")
                             ).properties(height=60 * len(cd) + 40), width="stretch")
            st.caption("Linha vermelha = 100% da capacidade. Acima dela, a operação não dá conta da demanda prevista.")
    with t4:
        st.caption("Gerado a partir dos sócios (aba Informações) e da mão de obra (aba Custos), agrupada por área.")
        try:
            st.graphviz_chart(dot_organograma(), width="stretch")
        except Exception as e:
            st.info(f"Não foi possível desenhar o organograma: {e}")
    auto_unica("projeto", v)
    salvo_txt()
    proxima("operacional")


def pg_fornecedores():
    cabecalho("fornecedores")
    fo = editor("fornecedores", {"nome": T("Fornecedor", required=True), "itens": T("O que fornece", width="medium"),
                                 "prazo_pagamento_dias": N("Prazo de pagamento (dias)", min_value=0, step=1),
                                 "prazo_entrega_dias": N("Prazo de entrega (dias)", min_value=0, step=1),
                                 "condicoes": T("Condições (pedido mínimo, frete, desconto)", width="medium"),
                                 "critico": CK("Crítico?")}, "nome")
    if not fo.empty:
        a, b, c = st.columns(3)
        a.metric("Fornecedores cadastrados", len(fo))
        b.metric("Prazo médio de pagamento (fornecedores)", f"{num_br(fo['prazo_pagamento_dias'].mean(), 0)} dias")
        c.metric("PMP usado no capital de giro", f"{num_br(R()['pmp'], 1)} dias")
        crit = fo[fo["critico"].astype(bool)]
        if not crit.empty:
            st.warning("**Fornecedores críticos** (ter alternativa homologada): " + ", ".join(crit["nome"]))
    proxima("fornecedores")
