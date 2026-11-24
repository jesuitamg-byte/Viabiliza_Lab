# ==============================================================================
# PNBOX ACADÊMICO - INTERFACE: CONCLUSÃO, ENTREGA, PROFESSOR E CONFIGURAÇÕES
# ==============================================================================
import json
import tempfile


# ------------------------------------------------------------------------------
# PONTOS DE RESTAURAÇÃO (versões salvas dentro do próprio projeto)
# ------------------------------------------------------------------------------
def foto_projeto() -> str:
    dados = {}
    for t in SCHEMA:
        if t == "versoes":
            continue
        dados[t] = ler(conn, t).to_dict("records")
    return json.dumps(dados, ensure_ascii=False, default=str)


def salvar_versao(nome: str, automatica=False):
    conn.execute("INSERT INTO versoes (nome, data, dados) VALUES (?,?,?)",
                 (("[auto] " if automatica else "") + nome, datetime.now().strftime("%d/%m/%Y %H:%M"), foto_projeto()))
    if automatica:   # guarda só as 5 cópias automáticas mais recentes
        ids = [r[0] for r in conn.execute("SELECT id FROM versoes WHERE nome LIKE '[auto]%' ORDER BY id DESC").fetchall()]
        for i in ids[5:]:
            conn.execute("DELETE FROM versoes WHERE id=?", (i,))
    conn.commit()


def restaurar_versao(vid: int):
    linha = conn.execute("SELECT dados FROM versoes WHERE id=?", (vid,)).fetchone()
    if not linha:
        return
    salvar_versao("Antes de restaurar", automatica=True)
    dados = json.loads(linha[0])
    for t, regs in dados.items():
        if t not in SCHEMA:
            continue
        cols = colunas_de(t)
        if t in ("projeto", "proposta_valor"):
            if regs:
                salvar_unica(conn, t, {c: regs[0][c] for c in cols if c in regs[0]})
            continue
        conn.execute(f"DELETE FROM {t}")
        for r in regs:
            cc = [c for c in cols if c in r]
            if cc:
                conn.execute(f"INSERT INTO {t} ({', '.join(cc)}) VALUES ({', '.join('?' * len(cc))})", [r[c] for c in cc])
    conn.commit()
    for k in ("buf", "snap", "hash"):
        ss[k] = {}
    ss.gen += 1
    invalidar()


def copia_automatica():
    """Cria um ponto de restauração a cada 10 minutos de uso."""
    agora = time.time()
    ultimo = ss.get("ultima_copia_auto", {}).get(ss.db)
    if ultimo is None:
        ss.setdefault("ultima_copia_auto", {})[ss.db] = agora
    elif agora - ultimo >= 600:
        salvar_versao("Cópia automática (10 min)", automatica=True)
        ss.ultima_copia_auto[ss.db] = agora


# ------------------------------------------------------------------------------
# ANÁLISE COM IA
# ------------------------------------------------------------------------------
def pg_ia():
    cabecalho("ia")
    res = R()
    st.markdown("""
<div class="pn-card-info">
<b>Como funciona:</b> o programa monta um texto com todos os números do seu plano e as instruções para a IA.
Você copia e cola em qualquer assistente gratuito (ChatGPT, Gemini, Copilot, Claude, DeepSeek...).
Nenhum dado sai do seu computador sem que você mesmo cole. <b>A IA comenta; quem decide é o grupo.</b>
Confira os números que ela citar com os do programa.
</div>""", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns([2, 1])
    with c1:
        escolha = st.multiselect("O que você quer que a IA faça?", list(TAREFAS_IA), format_func=lambda k: TAREFAS_IA[k],
                                 default=["indicadores", "concorrencia", "custos", "graficos"], key=wk("ia_tarefas"))
    with c2:
        publico = st.selectbox("Quem está pedindo", ["estudante de Administração", "estudante de Ciências Contábeis",
                                                      "estudante de Economia", "empreendedor iniciante", "professor"],
                               key=wk("ia_publico"))
    piv, _ = pivo_concorrencia()
    texto = gerar_prompt_ia(proj, d, res, escolha, publico, piv)
    st.caption(f"{len(texto):,} caracteres · cerca de {len(texto) // 4:,} palavras-token".replace(",", "."))
    t1, t2 = st.tabs(["📋 Copiar (botão no canto da caixa)", "✏️ Ver e editar antes de copiar"])
    with t1:
        st.code(texto, language="markdown", wrap_lines=True, height=420)
    with t2:
        editado = st.text_area("Texto", texto, height=420, label_visibility="collapsed", key=wk("ia_edit"))
        texto = editado
    st.download_button("⬇️ Baixar como arquivo .txt", texto.encode("utf-8"), f"prompt_IA_{PID}.txt", "text/plain",
                       width="stretch", type="primary")
    with st.expander("💡 Dicas para tirar mais proveito da IA"):
        st.markdown("""
- Se a IA disser que o texto é longo demais, cole em duas partes e escreva "aguarde a segunda parte".
- Peça depois: *"Agora explique como se eu nunca tivesse estudado finanças"* ou *"Resuma em 5 tópicos para o slide"*.
- Assistentes que executam código (ChatGPT com análise de dados, Gemini, Claude) conseguem desenhar os gráficos.
- Nunca copie a resposta da IA direto para o trabalho: use como ponto de partida e cite que a usou.
""")
    proxima("ia")


# ------------------------------------------------------------------------------
# SUMÁRIO EXECUTIVO E CONCLUSÃO
# ------------------------------------------------------------------------------
def pg_sumario():
    cabecalho("sumario")
    res = R()
    st.info("O sumário executivo é a primeira coisa que o leitor vê, mas a última que se escreve. "
            "Use o rascunho automático como ponto de partida e reescreva com as palavras do grupo.")
    if st.button("✨ Gerar rascunho automático a partir dos dados", type="primary"):
        novo = gerar_sumario(proj, d, res)
        salvar_unica(conn, "projeto", {"sumario_executivo": novo})
        ss.pop(wk("sumario_txt"), None)
        marcar_salvo()
        st.rerun()
    txt = st.text_area("Sumário executivo", proj["sumario_executivo"], height=360, key=wk("sumario_txt"),
                       placeholder="Clique no botão acima para gerar um rascunho, ou escreva aqui.")
    st.caption(f"{len(txt.split())} palavras. Recomenda-se entre 250 e 500.")
    st.divider()
    st.subheader("Conclusão e parecer do grupo")
    st.caption("O negócio é viável? Em que condições? O que precisa ser validado antes de investir? "
               "Use os indicadores e o semáforo para justificar.")
    conc = st.text_area("Conclusão", proj["conclusao"], height=240, key=wk("conclusao_txt"), label_visibility="collapsed")
    auto_unica("projeto", {"sumario_executivo": txt, "conclusao": conc})
    salvo_txt()
    proxima("sumario")


# ------------------------------------------------------------------------------
# MODO APRESENTAÇÃO
# ------------------------------------------------------------------------------
def pg_apresentacao():
    cabecalho("apresentacao")
    res = R()
    a1 = res["anos"][1]
    sem = semaforo(res)
    mk = mercado(proj, res)
    slides = ["capa", "mercado", "proposta", "numeros", "semaforo", "investimento", "dre", "caixa", "fim"]
    ss.setdefault("slide", 0)
    ss.slide = min(ss.slide, len(slides) - 1)
    n1, n2, n3 = st.columns([1, 3, 1])
    if n1.button("◀ Anterior", width="stretch", disabled=ss.slide == 0):
        ss.slide -= 1
        st.rerun()
    n2.progress((ss.slide + 1) / len(slides), text=f"Slide {ss.slide + 1} de {len(slides)}")
    if n3.button("Próximo ▶", width="stretch", disabled=ss.slide == len(slides) - 1, type="primary"):
        ss.slide += 1
        st.rerun()
    s = slides[ss.slide]
    st.markdown('<div class="pn-slide">', unsafe_allow_html=True)
    if s == "capa":
        st.markdown(f"<h1 style='text-align:center;font-size:3.4rem;margin-top:2rem'>{proj['nome_negocio']}</h1>", unsafe_allow_html=True)
        props = d["propostas"]
        if not props.empty:
            st.markdown(f"<h3 style='text-align:center;font-weight:400'>{props['oferecem'].iloc[0]}</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center'>{proj['cidade_uf']} · {str(proj['autores']).replace(chr(10), ', ')}</p>",
                    unsafe_allow_html=True)
    elif s == "mercado":
        st.header("🌍 O mercado")
        if mk["preenchido"]:
            c = st.columns(3)
            c[0].metric("TAM (mercado total)", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
            c[1].metric("SAM (alcançável)", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
            c[2].metric("SOM (nossa meta)", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
            st.markdown(f"### A receita do Ano 1 representa **{pct(mk['pct_som'])}** do SOM.")
        else:
            st.info("Preencha a ferramenta Tamanho do mercado.")
    elif s == "proposta":
        st.header("💎 Proposta de valor")
        for _, p_ in d["propostas"].iterrows():
            st.markdown(f"#### Para **{p_['persona_nome']}**, que deseja {p_['que_desejam']}")
            st.markdown(f"Oferecemos **{p_['oferecem']}**, por meio de {p_['por_meio']}, diferente de {p_['diferentemente']}.")
    elif s == "numeros":
        st.header("💰 Os números do primeiro ano")
        c = st.columns(4)
        c[0].metric("Investimento total", brl(res["investimento_total"]))
        c[1].metric("Receita bruta", brl(a1["rec_bruta"]))
        c[2].metric("Lucro líquido", brl(a1["resultado"]))
        c[3].metric("Lucratividade", pct(div(a1["resultado"], a1["rec_bruta"])))
        c = st.columns(4)
        c[0].metric("VPL (3 anos)", brl(res["vpl"]))
        c[1].metric("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a.")
        c[2].metric("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "> 36 meses")
        c[3].metric("Ponto de equilíbrio", brl(a1["pe_contabil"]))
    elif s == "semaforo":
        st.header("🚦 Saúde do negócio")
        for i in range(0, len(sem), 3):
            cols = st.columns(3)
            for c, it in zip(cols, sem[i:i + 3]):
                with c.container(border=True):
                    st.markdown(f"<div style='font-size:2.2rem'>{it['Status']}</div><b>{it['Indicador']}</b><br>"
                                f"<span style='font-size:1.6rem'>{it['Valor']}</span><br><small>{it['Leitura']}</small>",
                                unsafe_allow_html=True)
    elif s == "investimento":
        st.header("🏗️ Investimento e fontes")
        comp = composicoes(res, d)
        c1, c2 = st.columns(2)
        with c1:
            rosca(comp["investimento"], "Composição do investimento", "ap_inv", tabela=False)
        with c2:
            fontes = pd.DataFrame({"Item": ["Capital próprio", "Capital de terceiros"],
                                   "Valor": [res["capital_proprio"], res["capital_terceiros"]]})
            fontes["%"] = fontes["Valor"] / max(fontes["Valor"].sum(), 1)
            rosca(fontes, "Fontes de recursos", "ap_fontes", tabela=False)
    elif s == "dre":
        st.header("📈 Receita e resultado em 3 anos")
        anos = pd.DataFrame([{"Ano": f"Ano {k}", "Receita": v["rec_bruta"], "Resultado": v["resultado"]} for k, v in res["anos"].items()])
        longo = anos.melt("Ano", var_name="Série", value_name="R$")
        st.altair_chart(alt.Chart(longo).mark_bar().encode(x=alt.X("Ano:N", title=None), xOffset="Série:N", y="R$:Q",
                                                           color=alt.Color("Série:N", scale=alt.Scale(range=["#2f6fd0", "#18a058"])))
                        .properties(height=420), width="stretch")
    elif s == "caixa":
        st.header("🏦 Saldo de caixa mês a mês (Ano 1)")
        fm = res["fluxo_mensal"]
        if not fm.empty:
            st.altair_chart(alt.Chart(fm.assign(n=range(1, len(fm) + 1))).mark_area(
                line={"color": "#2f6fd0"}, color="#2f6fd033").encode(x=alt.X("n:Q", title="Mês"), y=alt.Y("Saldo final:Q", title="R$"))
                .properties(height=420), width="stretch")
    else:
        st.markdown("<h1 style='text-align:center;font-size:3.4rem;margin-top:3rem'>Obrigado! 🙌</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center;font-weight:400'>Perguntas?</h3>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.caption("Dica: aperte F11 para tela cheia e recolha o menu lateral (>> no topo) durante a apresentação.")


# ------------------------------------------------------------------------------
# RELATÓRIO E EXPORTAÇÃO
# ------------------------------------------------------------------------------
def planilha_excel(res) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([proj]).T.reset_index().set_axis(["Campo", "Valor"], axis=1).to_excel(w, sheet_name="Projeto", index=False)
        for t, df_t in d.items():
            if isinstance(df_t, pd.DataFrame) and t != "versoes":
                df_t.to_excel(w, sheet_name=t[:31], index=False)
        res["cat"].to_excel(w, sheet_name="Custos_e_precos", index=False)
        res["inv"].to_excel(w, sheet_name="Depreciacao", index=False)
        res["meses"].to_excel(w, sheet_name="Projecao_mensal", index=False)
        res["fluxo_mensal"].to_excel(w, sheet_name="Fluxo_caixa_mensal", index=False)
        tabela_dre(res).to_excel(w, sheet_name="DRE", index=False)
        tabela_indicadores(res).to_excel(w, sheet_name="Indicadores", index=False)
        pd.DataFrame(semaforo(res)).to_excel(w, sheet_name="Semaforo", index=False)
        cenarios(res, proj).to_excel(w, sheet_name="Cenarios", index=False)
        CHK().drop(columns=["destino"], errors="ignore").to_excel(w, sheet_name="Checagens", index=False)
    return buf.getvalue()


def pg_relatorio():
    cabecalho("relatorio")
    res = R()
    chk = CHK()
    pend = chk[chk["Situação"] != "OK"]
    comp = completude()
    c1, c2, c3 = st.columns(3)
    c1.metric("Plano preenchido", pct(comp, 0))
    c2.metric("Checagens com problema", len(pend))
    c3.metric("Páginas estimadas do PDF", "25 a 35")
    if not pend.empty:
        with st.expander(f"⚠️ {len(pend)} ponto(s) para revisar antes de entregar", expanded=True):
            for i, r in pend.iterrows():
                a, b = st.columns([4, 1])
                a.markdown(f"**{r['Teste']}**: {r['Valor apurado']}. {r['Como corrigir']}")
                if b.button("Ir corrigir →", key=f"rel_chk_{i}", width="stretch"):
                    ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir']}")
                    st.rerun()

    st.subheader("📦 Baixar")
    piv, _ = pivo_concorrencia()
    piv_i = piv
    k = st.columns(3)
    with k[0].container(border=True):
        st.markdown("#### 📄 Relatório completo (PDF)")
        st.caption("Todas as ferramentas, tabelas, gráficos e checagens, na ordem do PNBOX.")
        if st.button("Gerar PDF", type="primary", width="stretch", key="b_pdf"):
            with st.spinner("Montando o relatório... (pode levar alguns segundos)"):
                extras = {"comp": composicoes(res, d), "dre": tabela_dre(res), "indicadores": tabela_indicadores(res),
                          "cenarios": cenarios(res, proj), "checagens": tabela_checagens(res, d),
                          "swot_diag": pontuacao_swot(d["swot"]), "funil": tabela_funil(d["funil_vendas"]),
                          "concorrencia_pivot": piv_i}
                ss[f"pdf_{PID}"] = gerar_pdf(proj, d, res, extras)
        if ss.get(f"pdf_{PID}"):
            st.download_button("⬇️ Baixar PDF", ss[f"pdf_{PID}"], f"Plano_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🎤 Pitch deck (PDF 16:9)")
        st.caption("12 slides prontos para a apresentação oral.")
        if st.button("Gerar pitch", width="stretch", key="b_pitch"):
            with st.spinner("Montando os slides..."):
                ss[f"pitch_{PID}"] = gerar_pitch(proj, d, res)
        if ss.get(f"pitch_{PID}"):
            st.download_button("⬇️ Baixar pitch", ss[f"pitch_{PID}"], f"Pitch_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[2].container(border=True):
        st.markdown("#### 📝 Documento editável (Word)")
        st.caption("Para o grupo revisar o texto e formatar nas normas da ABNT.")
        if st.button("Gerar Word", width="stretch", key="b_docx"):
            with st.spinner("Montando o documento..."):
                ss[f"docx_{PID}"] = gerar_docx(proj, d, res, piv_i) or b""
        if ss.get(f"docx_{PID}") is not None and f"docx_{PID}" in ss:
            if ss[f"docx_{PID}"]:
                st.download_button("⬇️ Baixar .docx", ss[f"docx_{PID}"], f"Plano_{Path(ss.db).stem}.docx",
                                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
            else:
                st.warning("A biblioteca python-docx não está instalada neste computador.")
    k = st.columns(2)
    with k[0].container(border=True):
        st.markdown("#### 📊 Planilha (Excel)")
        st.caption("Todos os dados e cálculos, uma aba por tabela.")
        st.download_button("⬇️ Baixar planilha", planilha_excel(res), f"Planilha_{Path(ss.db).stem}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🗄️ Arquivo do projeto (.pnbox)")
        st.caption("Guarde no pen drive ou na nuvem. É o que o professor recebe para avaliar.")
        conn.commit()
        st.download_button("⬇️ Baixar .pnbox", Path(ss.db).read_bytes(), f"{Path(ss.db).stem}.pnbox",
                           "application/octet-stream", width="stretch", type="primary")
    st.divider()
    st.subheader("🕑 Pontos de restauração")
    st.caption("Salve uma versão antes de grandes mudanças. O programa também cria uma cópia automática a cada 10 minutos "
               "(as 5 mais recentes ficam guardadas).")
    a, b = st.columns([3, 1])
    nome_v = a.text_input("Nome da versão", placeholder="Ex.: Antes de mudar os preços", label_visibility="collapsed")
    if b.button("💾 Salvar versão", width="stretch"):
        salvar_versao(nome_v.strip() or "Versão manual")
        st.toast("Versão salva.")
        st.rerun()
    vs = pd.read_sql("SELECT id, nome, data FROM versoes ORDER BY id DESC", conn)
    if vs.empty:
        st.caption("Nenhuma versão salva ainda.")
    for _, v in vs.iterrows():
        a, b, c = st.columns([4, 1, 1])
        a.markdown(f"**{v['nome']}** · {v['data']}")
        if b.button("Restaurar", key=f"rest_{v['id']}", width="stretch"):
            restaurar_versao(int(v["id"]))
            st.toast("Versão restaurada. Uma cópia do estado anterior foi guardada.")
            st.rerun()
        if c.button("Apagar", key=f"apv_{v['id']}", width="stretch"):
            conn.execute("DELETE FROM versoes WHERE id=?", (int(v["id"]),))
            conn.commit()
            st.rerun()


# ------------------------------------------------------------------------------
# MODO PROFESSOR
# ------------------------------------------------------------------------------
def pg_professor():
    cabecalho("professor")
    st.markdown("""
<div class="pn-card-info">
Envie os arquivos <b>.pnbox</b> dos grupos. O programa recalcula tudo, confere as checagens, mede o preenchimento e
sugere uma nota pela rubrica. A nota sugerida é só um apoio: a banca registra a nota final.
Os arquivos não são guardados; ficam só na memória desta sessão.
</div>""", unsafe_allow_html=True)
    st.write("")
    with st.expander("📐 Rubrica de avaliação (critérios e pesos)", expanded=False):
        ss.setdefault("rubrica", pd.DataFrame(RUBRICA_PADRAO, columns=["Critério", "Peso"]))
        rub = st.data_editor(ss.rubrica, num_rows="fixed", hide_index=True, width="stretch", key="ed_rubrica",
                             column_config={"Critério": T(disabled=True, width="large"), "Peso": N(min_value=0, max_value=100, step=1)})
        st.caption(f"Total: {num_br(rub['Peso'].sum(), 0)} pontos. Os critérios são calculados automaticamente; ajuste apenas os pesos.")
        rubrica = list(zip(rub["Critério"], rub["Peso"]))
    disc, profn = st.columns(2)
    disciplina = disc.text_input("Disciplina", key="prof_disc")
    professor = profn.text_input("Professor(a)", key="prof_nome")
    arqs = st.file_uploader("Arquivos .pnbox dos grupos", type=["pnbox"], accept_multiple_files=True, key="prof_up")
    if not arqs:
        st.info("Nenhum arquivo enviado ainda.")
        return
    avs = []
    with st.spinner("Avaliando os projetos..."):
        tmp = Path(tempfile.mkdtemp())
        for a in arqs:
            dados = a.getvalue()
            if not dados.startswith(b"SQLite format 3"):
                st.warning(f"{a.name}: arquivo inválido, ignorado.")
                continue
            cam = tmp / a.name
            cam.write_bytes(dados)
            try:
                p_, d_ = carregar_projeto_arquivo(cam)
                av = avaliar(p_, d_, rubrica)
                av["Arquivo"] = a.name
                avs.append(av)
            except Exception as e:      # arquivo corrompido ou muito incompleto
                st.warning(f"{a.name}: não foi possível avaliar ({e}).")
    if not avs:
        return
    tab = pd.DataFrame([{k: v for k, v in av.items() if k not in ("res", "notas")} for av in avs])
    st.subheader("📋 Painel da turma")
    mostrar = tab[["Arquivo", "Projeto", "Integrantes", "Preenchimento", "Checagens OK", "Raio-X", "Investimento",
                   "Lucro Ano 1", "VPL", "TIR", "Nota sugerida"]]
    st.dataframe(mostrar, hide_index=True, width="stretch", column_config={
        "Preenchimento": st.column_config.ProgressColumn("Preenchido", format="percent", min_value=0, max_value=1),
        "Checagens OK": st.column_config.ProgressColumn("Checagens OK", format="percent", min_value=0, max_value=1),
        "Investimento": N(format="R$ %.0f"), "Lucro Ano 1": N(format="R$ %.0f"), "VPL": N(format="R$ %.0f"),
        "TIR": N(format="percent"), "Nota sugerida": N(format="%.1f")})
    c = st.columns(3)
    c[0].metric("Projetos", len(tab))
    c[1].metric("Nota média sugerida", num_br(tab["Nota sugerida"].mean(), 1))
    c[2].metric("Projetos com VPL positivo", f"{int((tab['VPL'] > 0).sum())} de {len(tab)}")

    st.subheader("✍️ Notas da banca")
    st.caption("A tabela vem com as notas sugeridas. Altere o que quiser; as fichas em PDF usam estas notas.")
    base = pd.DataFrame([[av["Projeto"]] + av["notas"] for av in avs], columns=["Projeto"] + [c_ for c_, _ in rubrica])
    ed = st.data_editor(base, hide_index=True, width="stretch", key="ed_notas_banca",
                        column_config={"Projeto": T(disabled=True)} |
                        {c_: N(c_[:28] + "…" if len(c_) > 28 else c_, min_value=0, max_value=float(p), step=0.5)
                         for c_, p in rubrica})
    ed["Total"] = ed[[c_ for c_, _ in rubrica]].sum(axis=1)
    st.dataframe(ed[["Projeto", "Total"]], hide_index=True)
    for av, (_, linha) in zip(avs, ed.iterrows()):
        av["notas_finais"] = [float(linha[c_]) for c_, _ in rubrica]

    st.subheader("🔍 Detalhes por grupo")
    for av in avs:
        with st.expander(f"{av['Projeto']} · nota sugerida {num_br(av['Nota sugerida'], 1)}"):
            st.write(f"**Integrantes:** {av['Integrantes'] or '-'}")
            st.write(f"**Pendências:** {av['Pendências'] or 'nenhuma'}")
            st.dataframe(pd.DataFrame({"Critério": [c_ for c_, _ in rubrica], "Peso": [p for _, p in rubrica],
                                       "Sugerida": av["notas"]}), hide_index=True, width="stretch")

    st.subheader("📦 Exportar")
    a, b = st.columns(2)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name="Turma", index=False)
        ed.to_excel(w, sheet_name="Notas_banca", index=False)
    a.download_button("📊 Planilha da turma (Excel)", buf.getvalue(), "avaliacao_turma.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    if b.button("📄 Gerar fichas de avaliação (PDF)", width="stretch", type="primary"):
        ss.fichas = ficha_avaliacao_pdf(avs, rubrica, professor, disciplina)
    if ss.get("fichas"):
        b.download_button("⬇️ Baixar fichas", ss.fichas, "fichas_avaliacao.pdf", "application/pdf", width="stretch")


# ------------------------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------------------------
def pg_config():
    cabecalho("config")
    t1, t2, t3 = st.tabs(["📚 Listas personalizadas", "💾 Onde ficam meus projetos", "🧹 Limpar este computador"])
    with t1:
        st.caption("As opções padrão continuam disponíveis. Aqui você acrescenta categorias, unidades, áreas e "
                   "classificações próprias do seu negócio; elas aparecem nas listas de escolha das ferramentas.")
        tipo = st.selectbox("Lista", list(NOMES_LISTAS), format_func=lambda k: NOMES_LISTAS[k], key="cfg_lista")
        st.markdown("**Padrão:** " + ", ".join(LISTAS_PADRAO.get(tipo, [])))
        editor("listas", {"valor": T("Novo item", required=True)}, "valor", filtro={"tipo": tipo})
    with t2:
        if NAVEGADOR:
            st.markdown("""
Você está usando a **versão on-line**. Os projetos ficam guardados **no armazenamento deste navegador, neste computador**.

- Atualizações do programa **não apagam** seus projetos: o arquivo do programa é trocado, mas a pasta de projetos do navegador continua.
- Eles **somem** se você limpar os dados de navegação do site, usar janela anônima ou trocar de navegador/computador.
- Por segurança, **baixe o .pnbox** (página Relatório e exportação) ao fim de cada sessão de trabalho.
""")
        elif MODO_WEB:
            st.warning("Nesta versão on-line temporária os projetos duram só enquanto a página estiver aberta. Baixe o .pnbox.")
        else:
            st.markdown(f"Os projetos ficam nesta pasta do computador:\n\n`{PASTA}`\n\n"
                        "Atualizar ou reinstalar o programa não apaga essa pasta.")
        st.write(f"Versão do programa: **{VERSAO_PROGRAMA}** · projetos encontrados: **{len(listar_projetos())}**")
    with t3:
        st.warning("Use em computadores de laboratório compartilhados, ao terminar. Todos os projetos guardados "
                   "**neste computador/navegador** serão apagados. Baixe antes o .pnbox do que quiser guardar.")
        ok = st.checkbox("Entendi e já baixei o que precisava", key="cfg_limpar_ok")
        if st.button("🧹 Apagar todos os projetos deste computador", disabled=not ok, type="primary"):
            conn.close()
            for p_ in listar_projetos():
                try:
                    p_.unlink()
                except OSError:
                    pass
            for k in ("db", "buf", "snap", "hash", "ultima_copia_auto"):
                ss.pop(k, None)
            ss.bemvindo_ok = False
            st.success("Pronto. Um projeto novo será criado ao recarregar.")
            st.rerun()


# ==============================================================================
# ROTEADOR
# ==============================================================================
ROTAS = {"inicio": pg_inicio, "projeto": pg_projeto, "segmentos": pg_segmentos, "personas": pg_personas,
         "jornada": pg_jornada, "mercado": pg_mercado, "proposta": pg_proposta, "concorrencia": pg_concorrencia,
         "forcas": pg_forcas, "oportunidades": pg_oportunidades, "swot": pg_swot, "marketing": pg_marketing,
         "canais": pg_canais, "operacional": pg_operacional, "fornecedores": pg_fornecedores,
         "investimentos": pg_investimentos, "produtos": pg_produtos, "custos": pg_custos, "preco": pg_preco,
         "dre": pg_dre, "fluxo": pg_fluxo, "indicadores": pg_indicadores, "graficos": pg_graficos,
         "simulador": pg_simulador, "experimentos": pg_experimentos, "riscos": pg_riscos, "cronograma": pg_cronograma,
         "realizado": pg_realizado, "ia": pg_ia, "sumario": pg_sumario, "apresentacao": pg_apresentacao,
         "relatorio": pg_relatorio, "professor": pg_professor, "config": pg_config}
try:
    ROTAS.get(ss.pagina, pg_inicio)()
    copia_automatica()
finally:
    try:
        conn.commit()
        conn.close()
    except Exception:
        pass
