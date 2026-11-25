"""
================================================================================
VIABILIZALAB - ANÁLISE ESTRATÉGICA (pnbox_estrategia.py)
================================================================================
Objetivos SMART, Business Model Canvas, matriz BCG, matriz GE/McKinsey,
Cinco Forças de Porter, Balanced Scorecard e simulação da reforma tributária.
Todos os cálculos aproveitam o que o aluno já preencheu nas outras ferramentas.
================================================================================
"""
import re

import pandas as pd

from pnbox_motor import brl, num_br, pct, div, custo_hora, mercado


# ------------------------------------------------------------------------------
# 1. OBJETIVOS SMART
# ------------------------------------------------------------------------------
PERSPECTIVAS = ["Financeira", "Clientes", "Processos internos", "Aprendizado e crescimento"]
VERBOS_VAGOS = {"melhorar", "otimizar", "fortalecer", "buscar", "tentar", "procurar", "desenvolver", "trabalhar"}

SMART = {
    "S": ("Específico", "Diz exatamente o que será alcançado, com um verbo de ação e um foco claro."),
    "M": ("Mensurável", "Tem um indicador e uma meta numérica."),
    "A": ("Atingível", "Explica por que é possível com os recursos do plano."),
    "R": ("Relevante", "Está ligado à estratégia (SWOT, proposta de valor ou viabilidade)."),
    "T": ("Temporal", "Tem prazo definido (mês do plano)."),
}


def checar_smart(o: dict) -> dict:
    """Verifica cada critério SMART de um objetivo e devolve {letra: (ok, dica)}."""
    txt = str(o.get("objetivo") or "").strip()
    palavras = re.findall(r"\w+", txt.lower())
    vago = bool(palavras) and palavras[0] in VERBOS_VAGOS and len(palavras) < 8
    s_ok = len(palavras) >= 5 and not vago
    m_ok = bool(str(o.get("indicador") or "").strip()) and float(o.get("meta") or 0) != 0
    a_ok = len(str(o.get("por_que_atingivel") or "").strip()) >= 15
    r_ok = len(str(o.get("relevancia") or "").strip()) >= 10 and bool(str(o.get("perspectiva") or "").strip())
    t_ok = float(o.get("prazo_meses") or 0) > 0
    return {
        "S": (s_ok, "Escreva com verbo de ação e o que muda (ex.: \"Conquistar 300 assinantes pagantes\")." if not s_ok else ""),
        "M": (m_ok, "Informe o indicador e a meta numérica." if not m_ok else ""),
        "A": (a_ok, "Explique em uma frase por que a meta é possível (equipe, verba, capacidade, histórico)." if not a_ok else ""),
        "R": (r_ok, "Escolha a perspectiva e diga a que estratégia ou problema ele responde." if not r_ok else ""),
        "T": (t_ok, "Defina o prazo em meses." if not t_ok else ""),
    }


def frase_smart(o: dict) -> str:
    meta = num_br(o.get("meta") or 0, 2 if abs(float(o.get("meta") or 0)) < 100 else 0)
    un = str(o.get("unidade") or "").strip()
    prazo = int(float(o.get("prazo_meses") or 0))
    resp = str(o.get("responsavel") or "").strip()
    return (f"{str(o.get('objetivo') or '').strip()}: atingir {meta} {un} em {o.get('indicador') or 'indicador'} "
            f"até o mês {prazo}" + (f", sob responsabilidade de {resp}" if resp else "") + ".")


def objetivos_sugeridos(proj: dict, d: dict, res: dict) -> list:
    """Rascunhos de objetivos SMART a partir dos números do plano."""
    a1 = res["anos"][1]
    sug = [
        {"objetivo": "Alcançar o faturamento previsto no primeiro ano", "perspectiva": "Financeira",
         "indicador": "Receita bruta acumulada", "meta": round(a1["rec_bruta"], 0), "unidade": "R$", "prazo_meses": 12,
         "por_que_atingivel": "Meta calculada no plano a partir do volume e dos preços previstos.",
         "relevancia": "Garante a viabilidade financeira projetada.", "iniciativa": "Acompanhar vendas semanalmente."},
        {"objetivo": "Operar acima do ponto de equilíbrio contábil todos os meses", "perspectiva": "Financeira",
         "indicador": "Receita mensal acima do PE", "meta": round(a1["pe_contabil"] / 12, 0), "unidade": "R$/mês", "prazo_meses": 12,
         "por_que_atingivel": "O ponto de equilíbrio foi calculado com os custos do plano.",
         "relevancia": "Evita prejuízo e consumo do capital de giro.", "iniciativa": "Controle mensal de custos fixos."},
    ]
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        sug.append({"objetivo": "Manter o cancelamento de assinantes sob controle", "perspectiva": "Clientes",
                    "indicador": "Churn mensal", "meta": round(rc["churn"] * 100, 1), "unidade": "% ao mês (máximo)",
                    "prazo_meses": 12, "por_que_atingivel": "Churn usado na projeção dos planos.",
                    "relevancia": "Cada ponto de churn reduz o LTV e o MRR.", "iniciativa": "Programa de sucesso do cliente."})
        sug.append({"objetivo": "Chegar ao número de assinantes previsto no fim do primeiro ano", "perspectiva": "Clientes",
                    "indicador": "Assinantes ativos", "meta": round(rc["ativos12"], 0), "unidade": "assinantes", "prazo_meses": 12,
                    "por_que_atingivel": "Resultado da captação e do churn projetados.",
                    "relevancia": "Base de assinantes sustenta a receita recorrente.", "iniciativa": "Campanhas mensais de captação."})
    else:
        sug.append({"objetivo": "Conquistar e manter clientes satisfeitos", "perspectiva": "Clientes",
                    "indicador": "Nota de satisfação (NPS)", "meta": 70, "unidade": "pontos", "prazo_meses": 12,
                    "por_que_atingivel": "Pesquisa simples após a compra, com ajustes mensais.",
                    "relevancia": "Clientes satisfeitos voltam e indicam, reduzindo o custo de aquisição.",
                    "iniciativa": "Pesquisa de satisfação pelo WhatsApp."})
    ch = custo_hora(res, d, proj)
    if ch["horas_cobraveis"] > 0 and ch["horas_demandadas_m1"] > 0:
        sug.append({"objetivo": "Usar bem as horas da equipe de serviços", "perspectiva": "Processos internos",
                    "indicador": "Taxa de utilização", "meta": round(ch["utilizacao"] * 100, 0), "unidade": "%", "prazo_meses": 6,
                    "por_que_atingivel": "Taxa prevista na ferramenta de serviços.",
                    "relevancia": "Hora ociosa é custo sem receita.", "iniciativa": "Agenda semanal e controle de horas."})
    else:
        sug.append({"objetivo": "Padronizar os processos de produção e atendimento", "perspectiva": "Processos internos",
                    "indicador": "Processos documentados", "meta": 5, "unidade": "processos", "prazo_meses": 6,
                    "por_que_atingivel": "Equipe pequena e processos já descritos no plano operacional.",
                    "relevancia": "Reduz erros e desperdício, protegendo a margem.", "iniciativa": "Fichas e checklists."})
    sug.append({"objetivo": "Capacitar a equipe nas competências críticas do negócio", "perspectiva": "Aprendizado e crescimento",
                "indicador": "Horas de capacitação por pessoa", "meta": 20, "unidade": "horas/ano", "prazo_meses": 12,
                "por_que_atingivel": "Cursos gratuitos do Sebrae e treinamentos internos.",
                "relevancia": "Equipe preparada sustenta a qualidade prometida na proposta de valor.",
                "iniciativa": "Plano anual de treinamento."})
    return sug


# ------------------------------------------------------------------------------
# 2. BUSINESS MODEL CANVAS (montado a partir das outras ferramentas)
# ------------------------------------------------------------------------------
def _lista(x, n=6):
    return [str(v).strip() for v in x if str(v).strip()][:n]


def canvas_blocos(proj: dict, d: dict, res: dict) -> dict:
    seg = d["segmentos"]
    seg_l = []
    if not seg.empty:
        s_ = seg.assign(m=seg[["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]].mean(axis=1))
        seg_l = _lista(s_.sort_values("m", ascending=False)["nome"], 5)
    props = d.get("propostas")
    prop_l = _lista(props["oferecem"], 4) if props is not None and not props.empty else []
    if props is not None and not props.empty:
        vals = [v.strip() for x in props["valores"] for v in str(x).split(";") if v.strip()]
        if vals:
            prop_l.append("Valores: " + ", ".join(dict.fromkeys(vals)))
    canais = _lista(d["canais"]["canal"], 6) if "canal" in d["canais"] else []
    rel = [x.strip() for x in str(proj.get("canvas_relacionamento") or "").split("\n") if x.strip()]
    if not rel:
        etapas = d["jornada_cliente"]
        rel = _lista(etapas.loc[etapas["macro_fase"] == "Depois", "oportunidades"], 4)
    rec = []
    cat = res["cat"]
    if not cat.empty:
        for c_, v_ in cat.groupby("categoria")["receita_mensal"].sum().sort_values(ascending=False).items():
            rec.append(f"{c_}: {pct(div(v_ * 12, res['anos'][1]['rec_bruta']), 0)} da receita")
    if res.get("recorrencia", {}).get("tem"):
        a_ = d["assinaturas"]
        rec.append("Assinaturas: " + ", ".join(f"{p} ({brl(v)})" for p, v in zip(a_["plano"], a_["preco_ciclo"])))
    recursos = _lista(d["inv_fixos"].sort_values("valor_unitario", ascending=False)["descricao"], 3) + \
        [f"Equipe: {num_br(d['pessoal']['num_empregados'].sum(), 0)} pessoa(s)"] if not d["pessoal"].empty else \
        _lista(d["inv_fixos"]["descricao"], 3)
    atividades = [x.strip() for x in re.split(r"->|→|;|\n", str(proj.get("op_processos") or "")) if x.strip()][:6]
    parc = [x.strip() for x in str(proj.get("canvas_parcerias") or "").split("\n") if x.strip()]
    if not parc and not d["fornecedores"].empty:
        fo = d["fornecedores"]
        parc = _lista(fo.sort_values("critico", ascending=False)["nome"], 5)
    a1 = res["anos"][1]
    custos = [f"Pessoal: {pct(div(a1['pessoal'], a1['rec_bruta']), 0)} da receita",
              f"Custos diretos: {pct(div(a1['custo_dir'], a1['rec_bruta']), 0)}",
              f"Demais custos fixos: {pct(div(a1['fixos'], a1['rec_bruta']), 0)}",
              f"Investimento inicial: {brl(res['investimento_total'])}"]
    return {"Parcerias-chave": parc, "Atividades-chave": atividades, "Recursos-chave": recursos,
            "Proposta de valor": prop_l, "Relacionamento com clientes": rel, "Canais": canais,
            "Segmentos de clientes": seg_l, "Estrutura de custos": custos, "Fontes de receita": rec}


# ------------------------------------------------------------------------------
# 3. MATRIZ BCG (crescimento do mercado x participação relativa)
# ------------------------------------------------------------------------------
QUADRANTES_BCG = {
    "Estrela": ("⭐", "Mercado em alta e posição forte. Investir para manter a liderança; consome caixa, mas tende a virar vaca leiteira."),
    "Vaca leiteira": ("🐄", "Mercado maduro e posição forte. Gera caixa: manter eficiência e usar o dinheiro para financiar estrelas e interrogações."),
    "Interrogação": ("❓", "Mercado em alta, posição fraca. Decidir: investir para ganhar participação ou abandonar."),
    "Abacaxi": ("🍍", "Mercado parado e posição fraca. Rever preço e custo, reposicionar ou descontinuar."),
}


def bcg_tabela(d: dict, res: dict, limite_cresc: float = 10.0) -> pd.DataFrame:
    cat = res["cat"]
    b = d["bcg"]
    linhas = []
    for _, r in b.iterrows():
        prod = str(r["produto"]).strip()
        if not prod:
            continue
        linha = cat[cat["descricao"] == prod]
        receita = float(linha["receita_mensal"].sum()) * 12 if not linha.empty else 0.0
        margem = float(linha["margem_pct"].iloc[0]) if not linha.empty else float("nan")
        if prod.startswith("Assinatura") and res.get("recorrencia", {}).get("tem"):
            receita = res["anos"][1]["rec_cat"].get("Assinatura", 0.0)
        cres, part = float(r["crescimento_mercado_pct"]), float(r["participacao_relativa"])
        alto, forte = cres >= limite_cresc, part >= 1.0
        quad = "Estrela" if alto and forte else "Vaca leiteira" if forte else "Interrogação" if alto else "Abacaxi"
        linhas.append({"Produto": prod, "Crescimento do mercado (%)": cres, "Participação relativa": part,
                       "Receita anual": receita, "Margem": margem, "Quadrante": quad})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 4. MATRIZ GE / McKINSEY (atratividade do mercado x força competitiva)
# ------------------------------------------------------------------------------
def forca_competitiva_padrao(d: dict) -> float:
    """Nota da 'Nossa Empresa' na matriz de concorrência comparada à média dos concorrentes, em escala 0-10."""
    n = d["concorrencia_notas"]
    if n.empty:
        return 5.0
    nossa = n.loc[n["empresa"] == "Nossa Empresa", "nota"].mean()
    outras = n.loc[n["empresa"] != "Nossa Empresa", "nota"].mean()
    if nossa != nossa:
        return 5.0
    if outras != outras or outras == 0:
        return float(nossa)
    return float(min(10, max(0, 5 * nossa / outras)))


def ge_tabela(d: dict) -> pd.DataFrame:
    seg = d["segmentos"]
    ge = d["ge"]
    padrao = forca_competitiva_padrao(d)
    forcas = dict(zip(ge["unidade"], ge["forca"]))
    linhas = []
    for _, s_ in seg.iterrows():
        atr = float(pd.Series([s_["nota_tamanho"], s_["nota_lucro"], s_["nota_longevidade"], s_["nota_concorrencia"]]).mean())
        f_inf = forcas.get(s_["nome"], 0) or 0
        forca = float(f_inf) if f_inf > 0 else round((padrao + float(s_["nota_aderencia"])) / 2, 1)
        def nivel(v):
            return "Alta" if v >= 6.67 else ("Média" if v >= 3.33 else "Baixa")
        na, nf = nivel(atr), nivel(forca)
        pontos = {"Alta": 2, "Média": 1, "Baixa": 0}
        soma = pontos[na] + pontos[nf]
        rec = "Investir e crescer" if soma >= 3 else ("Seletividade / manter" if soma == 2 else "Colher ou desinvestir")
        linhas.append({"Segmento": s_["nome"], "Atratividade": round(atr, 1), "Força competitiva": round(forca, 1),
                       "Nível atratividade": na, "Nível força": nf, "Recomendação": rec,
                       "Origem da força": "informada" if f_inf > 0 else "automática (aderência + concorrência)"})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 5. CINCO FORÇAS DE PORTER
# ------------------------------------------------------------------------------
FORCAS_PORTER = [
    ("Rivalidade entre concorrentes", "Quantos concorrentes, quão parecidos e quão agressivos em preço?"),
    ("Ameaça de novos entrantes", "É fácil abrir um negócio igual? Exige muito capital, licença, marca ou tecnologia?"),
    ("Ameaça de produtos substitutos", "O cliente resolve o mesmo problema de outro jeito?"),
    ("Poder de barganha dos clientes", "O cliente troca de fornecedor com facilidade? Compra em grande volume? Compara preço?"),
    ("Poder de barganha dos fornecedores", "Poucos fornecedores? Insumo crítico ou difícil de substituir?"),
]


def porter_sugestoes(proj: dict, d: dict, res: dict) -> dict:
    """Nota sugerida (1 a 5) e motivo para cada força, a partir dos dados do plano."""
    n_conc = int(d["concorrentes"]["nome"].astype(str).str.strip().ne("").sum())
    nota_conc = forca_competitiva_padrao(d)
    riv = 2 if n_conc <= 1 else (3 if n_conc <= 3 else 4)
    if nota_conc < 5:
        riv = min(5, riv + 1)
    inv = res["investimento_total"]
    ent = 4 if inv < 50000 else (3 if inv < 200000 else 2)
    imc = res["anos"][1]["imc"]
    subs = 3
    cli = 4 if str(proj.get("tipo_negocio", "")).startswith("Produtos") and imc < 0.4 else 3
    fo = d["fornecedores"]
    criticos = int(fo["critico"].sum()) if not fo.empty else 0
    forn = 2 if fo.empty or criticos == 0 else (3 if criticos <= 2 else 4)
    return {
        FORCAS_PORTER[0][0]: (riv, f"{n_conc} concorrente(s) cadastrado(s); nota da empresa frente à média deles: {num_br(nota_conc, 1)}/10."),
        FORCAS_PORTER[1][0]: (ent, f"Investimento inicial de {brl(inv)}: quanto menor, mais fácil alguém copiar o negócio."),
        FORCAS_PORTER[2][0]: (subs, "Avalie os substitutos citados na proposta de valor (o que o cliente usa hoje)."),
        FORCAS_PORTER[3][0]: (cli, f"Margem de contribuição de {pct(imc, 0)}: margens baixas indicam clientes sensíveis a preço."),
        FORCAS_PORTER[4][0]: (forn, f"{criticos} fornecedor(es) marcado(s) como crítico(s)."),
    }


def porter_tabela(d: dict) -> pd.DataFrame:
    p = d["porter"]
    base = pd.DataFrame({"Força": [f for f, _ in FORCAS_PORTER]})
    if p.empty:
        base["Intensidade"] = 0.0
        base["Justificativa"] = ""
        base["Estratégia"] = ""
        return base
    m = p.set_index("forca")
    base["Intensidade"] = base["Força"].map(lambda f: float(m["intensidade"].get(f, 0)) if f in m.index else 0.0)
    base["Justificativa"] = base["Força"].map(lambda f: str(m["justificativa"].get(f, "")) if f in m.index else "")
    base["Estratégia"] = base["Força"].map(lambda f: str(m["estrategia"].get(f, "")) if f in m.index else "")
    return base


def porter_leitura(tab: pd.DataFrame) -> str:
    v = tab["Intensidade"]
    v = v[v > 0]
    if v.empty:
        return "Atribua uma intensidade de 1 a 5 a cada força."
    media = float(v.mean())
    forte = tab.loc[tab["Intensidade"].idxmax(), "Força"]
    nivel = "alta" if media >= 3.5 else ("moderada" if media >= 2.5 else "baixa")
    return (f"Intensidade média de {num_br(media, 1)} em 5: a pressão competitiva do setor é {nivel}. "
            f"A força mais intensa é \"{forte}\"; ela deve orientar as primeiras ações estratégicas. "
            + ("Setores de pressão alta exigem diferenciação clara ou custos muito baixos para sustentar margens."
               if nivel == "alta" else "Há espaço para margens saudáveis se o negócio entregar a proposta de valor."))


# ------------------------------------------------------------------------------
# 6. BALANCED SCORECARD
# ------------------------------------------------------------------------------
def bsc_indicadores_plano(proj: dict, d: dict, res: dict) -> dict:
    """Indicadores que o próprio plano já calcula, organizados pelas quatro perspectivas."""
    a1 = res["anos"][1]
    fin = [("Lucratividade líquida (Ano 1)", pct(div(a1["resultado"], a1["rec_bruta"]))),
           (f"VPL ({res['N']} anos)", brl(res["vpl"])),
           ("Margem de contribuição", pct(a1["imc"]))]
    cli = []
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        fin.append(("MRR no Mês 12", brl(rc["mrr12"])))
        cli += [("Churn mensal", pct(rc["churn"])), ("LTV/CAC", f"{num_br(rc['ltv_cac'], 1)}x" if rc["ltv_cac"] == rc["ltv_cac"] else "n.a.")]
    mk = mercado(proj, res)
    if mk["preenchido"]:
        cli.append(("Participação no SOM (Ano 1)", pct(mk["pct_som"])))
    proc = []
    ch = custo_hora(res, d, proj)
    if ch["horas_demandadas_m1"] > 0:
        proc.append(("Ocupação da equipe (Mês 12)", pct(ch["ocupacao_m12"])))
    proc.append(("Ciclo financeiro", f"{num_br(res['ciclo_fin'], 0)} dias"))
    apr = [("Peso da folha na receita líquida", pct(div(a1["pessoal"], a1["rec_liquida"])))]
    return {"Financeira": fin, "Clientes": cli, "Processos internos": proc, "Aprendizado e crescimento": apr}


# ------------------------------------------------------------------------------
# 7. REFORMA TRIBUTÁRIA: SIMPLES "PURO" x IBS/CBS POR FORA ("HÍBRIDO")
# ------------------------------------------------------------------------------
# Parcela do DAS (1ª faixa) que corresponde aos tributos substituídos pela reforma, pelas tabelas da LC 123/2006.
# PIS + Cofins são substituídos pela CBS (a partir de 2027); ICMS/ISS pelo IBS (2029 a 2033); IPI zerado.
PARTILHA_DAS = {
    "Anexo I - Comércio": {"federais_consumo": 0.1274 + 0.0276, "icms_iss": 0.34},
    "Anexo II - Indústria": {"federais_consumo": 0.1151 + 0.0249 + 0.075, "icms_iss": 0.32},
    "Anexo III - Serviços": {"federais_consumo": 0.1282 + 0.0278, "icms_iss": 0.335},
    "Anexo IV - Serviços": {"federais_consumo": 0.1767 + 0.0383, "icms_iss": 0.445},
    "Anexo V - Serviços": {"federais_consumo": 0.1410 + 0.0305, "icms_iss": 0.14},
}


def simular_reforma(receita: float, aliq_simples: float, anexo: str, ano: int, cbs: float, ibs: float,
                    b2b: float, compras_credito: float) -> dict:
    """
    Comparação didática e simplificada (alíquotas por dentro, sobre a receita):
    - Simples "puro": paga o DAS cheio; o cliente empresa só aproveita como crédito a parcela de IBS/CBS do DAS.
    - Simples "híbrido": tira do DAS a parcela dos tributos substituídos e recolhe IBS e CBS pelo regime regular,
      com crédito sobre as compras; o cliente empresa aproveita o crédito integral.
    """
    part = PARTILHA_DAS.get(anexo, PARTILHA_DAS["Anexo III - Serviços"])
    # Em 2027-2028 só PIS/Cofins já foram substituídos (CBS); ICMS/ISS seguem no DAS até a transição do IBS
    # ICMS e ISS são reduzidos em 10% ao ano de 2029 a 2032 e extintos em 2033 (EC 132/2023)
    frac_icms = 0.0 if ano <= 2028 else (1.0 if ano >= 2033 else 0.1 * (ano - 2028))
    substituida = part["federais_consumo"] + part["icms_iss"] * frac_icms
    das = receita * aliq_simples
    iva = cbs + ibs
    puro_total = das
    puro_credito_cliente = das * substituida * b2b
    das_hibrido = das * (1 - substituida)
    iva_debito = receita * iva
    iva_credito = compras_credito * iva
    hib_total = das_hibrido + max(0.0, iva_debito - iva_credito)
    hib_credito_cliente = iva_debito * b2b
    return {"das_puro": das, "puro_total": puro_total, "puro_credito_cliente": puro_credito_cliente,
            "das_hibrido": das_hibrido, "iva_debito": iva_debito, "iva_credito": iva_credito, "hib_total": hib_total,
            "hib_credito_cliente": hib_credito_cliente, "parcela_substituida": substituida, "iva": iva,
            "diferenca": hib_total - puro_total,
            "custo_liquido_cliente_puro": -puro_credito_cliente, "custo_liquido_cliente_hib": -hib_credito_cliente}
