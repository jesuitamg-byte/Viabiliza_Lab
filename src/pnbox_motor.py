"""
================================================================================
VIABILIZALAB - MOTOR ECONÔMICO-FINANCEIRO (pnbox_motor.py)
================================================================================
Recebe as tabelas do projeto e calcula: custos unitários (ficha técnica),
investimento total, capital de giro, depreciação, cronograma dos empréstimos,
projeção mensal do Ano 1, DRE dos Anos 1 a 3, balanço simplificado, fluxos de
caixa, VPL, TIR, payback, pontos de equilíbrio, cenários e checagens.
================================================================================
"""

import math

import pandas as pd

CATEGORIAS = ["Fabricação Própria", "Revenda", "Serviço"]


# ------------------------------------------------------------------------------
# Formatação no padrão brasileiro
# ------------------------------------------------------------------------------
def brl(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def num_br(v, casas=2) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, casas=2) -> str:
    return "n.a." if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))) \
        else f"{num_br(v * 100, casas)}%"


def div(a, b):
    return a / b if b not in (0, 0.0) and b is not None else float("nan")


# ------------------------------------------------------------------------------
# 1. CUSTO UNITÁRIO DOS PRODUTOS (FICHA TÉCNICA x CUSTO INFORMADO)
# ------------------------------------------------------------------------------
def custos_unitarios(cat: pd.DataFrame, mat: pd.DataFrame, ficha: pd.DataFrame) -> pd.DataFrame:
    precos = dict(zip(mat["nome"], mat["valor_unitario"]))
    f = ficha.copy()
    f["valor_unitario"] = f["material"].map(precos).fillna(0.0)
    f["custo"] = f["quantidade"] * f["valor_unitario"]
    custo_ficha = f.groupby("produto")["custo"].sum().to_dict()

    df = cat.copy()
    custos, origens = [], []
    for _, r in df.iterrows():
        if r["categoria"] != "Revenda" and r["descricao"] in custo_ficha:
            custos.append(custo_ficha[r["descricao"]])
            origens.append("Ficha técnica")
        else:
            custos.append(float(r["custo_aquisicao"]))
            origens.append("Valor de aquisição" if r["categoria"] == "Revenda" else "Custo informado")
    df["custo_unit"] = custos
    df["origem_custo"] = origens
    df["margem_unit"] = df["preco_venda"] - df["custo_unit"]
    df["margem_pct"] = [div(m, p) for m, p in zip(df["margem_unit"], df["preco_venda"])]
    df["markup"] = [div(p, c) for p, c in zip(df["preco_venda"], df["custo_unit"])]
    df["receita_mensal"] = df["preco_venda"] * df["qtd_mensal"]
    df["custo_mensal"] = df["custo_unit"] * df["qtd_mensal"]
    return df


# ------------------------------------------------------------------------------
# 2. CRONOGRAMA DOS EMPRÉSTIMOS (TABELA PRICE, PARCELAS MENSAIS)
# ------------------------------------------------------------------------------
def cronograma_emprestimos(emp: pd.DataFrame, anos: int = 4) -> dict:
    juros = [0.0] * anos
    amort = [0.0] * anos
    saldo_fim = [0.0] * anos
    juros_m = [0.0] * (anos * 12)
    amort_m = [0.0] * (anos * 12)
    for _, r in emp.iterrows():
        pv, n = float(r["valor"]), int(round(r["prazo_meses"]))
        if pv <= 0:
            continue
        n = max(n, 1)
        i = (1 + r["taxa_anual_pct"] / 100.0) ** (1 / 12) - 1
        pmt = pv / n if i == 0 else pv * i / (1 - (1 + i) ** (-n))
        saldo = pv
        for m in range(anos * 12):
            ano = m // 12
            if m < n and saldo > 1e-9:
                j = saldo * i
                a = min(pmt - j, saldo)
                saldo -= a
                juros[ano] += j
                amort[ano] += a
                juros_m[m] += j
                amort_m[m] += a
            if m % 12 == 11:
                saldo_fim[ano] += max(saldo, 0.0)
    return {"juros": juros, "amort": amort, "saldo_fim": saldo_fim, "juros_m": juros_m, "amort_m": amort_m}


# ------------------------------------------------------------------------------
# 3. TIR (bisseção)
# ------------------------------------------------------------------------------
def calcular_tir(fluxos):
    if not fluxos or all(f >= 0 for f in fluxos) or all(f <= 0 for f in fluxos):
        return None
    vpl = lambda r: sum(f / (1 + r) ** t for t, f in enumerate(fluxos))
    lo, hi = -0.99, 10.0
    if vpl(lo) * vpl(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if vpl(lo) * vpl(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


# ------------------------------------------------------------------------------
# 3b. RECEITA RECORRENTE (ASSINATURAS) POR SAFRAS DE CLIENTES
# ------------------------------------------------------------------------------
COLS_ASS = ["Mês", "Novos", "Cancelamentos", "Ativos", "Receita reconhecida", "Cobranças (caixa)",
            "Custo de entrega", "Taxas de cobrança"]


def simular_assinaturas(ass: pd.DataFrame | None, g_a: float, inf: float, meses: int = 36):
    """
    Cada mês de entrada forma uma safra (coorte) de assinantes. A safra encolhe a cada renovação
    conforme o churn; quem paga por ciclo (trimestral, anual...) só pode sair na renovação.
    - Cobrança (caixa): no início de cada ciclo, preço do ciclo x assinantes ativos.
    - Receita reconhecida: preço do ciclo / meses do ciclo, todos os meses em que o assinante está ativo.
    - Custo de entrega: custo mensal por assinante ativo (nuvem, APIs, suporte), corrigido pela inflação.
    - Taxas de cobrança: % sobre o valor cobrado (meio de pagamento, loja de aplicativos).
    - Novos assinantes: crescem à taxa mensal do plano no Ano 1; nos Anos 2 e 3 partem do Mês 12
      e crescem à taxa anual de crescimento do projeto.
    """
    vazio = pd.DataFrame(columns=COLS_ASS), pd.DataFrame(columns=["Mês", "Plano", "Ativos", "Receita reconhecida"])
    if ass is None or ass.empty:
        return vazio
    tot = {c: [0.0] * meses for c in COLS_ASS[1:]}
    por_plano = []
    for _, p in ass.iterrows():
        preco, c = float(p["preco_ciclo"] or 0), max(1, int(round(float(p["ciclo_meses"] or 1))))
        n1, g = float(p["novos_mes1"] or 0), float(p["crescimento_novos_pct"] or 0) / 100
        churn_m = min(max(float(p["churn_mensal_pct"] or 0) / 100, 0.0), 1.0)
        sobra_ciclo = (1 - churn_m) ** c              # fração que renova a cada ciclo
        custo, taxa = float(p["custo_entrega_mensal"] or 0), float(p["taxa_cobranca_pct"] or 0) / 100
        if preco <= 0 and n1 <= 0:
            continue
        novos = [n1 * (1 + g) ** t if t < 12 else n1 * (1 + g) ** 11 * (1 + g_a) ** (t // 12) for t in range(meses)]
        ativos, rec, cob, canc = [0.0] * meses, [0.0] * meses, [0.0] * meses, [0.0] * meses
        for t0 in range(meses):
            for t in range(t0, meses):
                a = t - t0
                ren = a // c
                vivos = novos[t0] * sobra_ciclo ** ren
                ativos[t] += vivos
                rec[t] += vivos * preco / c
                if a % c == 0:
                    cob[t] += vivos * preco
                    if a > 0:
                        canc[t] += novos[t0] * sobra_ciclo ** (ren - 1) - vivos
        for t in range(meses):
            tot["Novos"][t] += novos[t]
            tot["Cancelamentos"][t] += canc[t]
            tot["Ativos"][t] += ativos[t]
            tot["Receita reconhecida"][t] += rec[t]
            tot["Cobranças (caixa)"][t] += cob[t]
            tot["Custo de entrega"][t] += ativos[t] * custo * (1 + inf) ** (t // 12)
            tot["Taxas de cobrança"][t] += cob[t] * taxa
            por_plano.append({"Mês": t + 1, "Plano": p["plano"], "Ativos": ativos[t], "Receita reconhecida": rec[t]})
    if not por_plano:
        return vazio
    df = pd.DataFrame({"Mês": list(range(1, meses + 1)), **tot})
    return df, pd.DataFrame(por_plano)


def metricas_recorrencia(s: pd.DataFrame, planos: pd.DataFrame, ass: pd.DataFrame, gasto_aquisicao_m1: float) -> dict:
    """Indicadores de negócios por assinatura (Ano 1)."""
    if s is None or s.empty or float(s["Receita reconhecida"].sum()) <= 0:
        return {"tem": False}
    a1 = s.iloc[:12]
    mrr1, mrr12 = float(s["Receita reconhecida"].iloc[0]), float(s["Receita reconhecida"].iloc[11])
    ativos12 = float(s["Ativos"].iloc[11])
    arpu = div(mrr12, ativos12)
    # churn mensal médio: churn informado em cada plano, ponderado pelos assinantes ativos no Mês 12
    p12 = planos[planos["Mês"] == 12].groupby("Plano")["Ativos"].sum()
    taxas = dict(zip(ass["plano"], ass["churn_mensal_pct"] / 100))
    churn = div(sum(p12.get(k, 0.0) * taxas.get(k, 0.0) for k in p12.index), float(p12.sum())) if p12.sum() > 0 else 0.0
    rec_a1 = float(a1["Receita reconhecida"].sum())
    margem_bruta = div(rec_a1 - float(a1["Custo de entrega"].sum()) - float(a1["Taxas de cobrança"].sum()), rec_a1)
    lucro_bruto_cliente = arpu * margem_bruta
    ltv = div(lucro_bruto_cliente, churn) if churn and churn > 0 else float("nan")
    vida_meses = div(1, churn) if churn and churn > 0 else float("nan")
    novos1 = float(s["Novos"].iloc[0])
    cac = div(gasto_aquisicao_m1, novos1) if novos1 > 0 else float("nan")
    ltv_cac = div(ltv, cac) if cac == cac and cac and cac > 0 and ltv == ltv else float("nan")
    payback_cac = div(cac, lucro_bruto_cliente) if lucro_bruto_cliente > 0 and cac == cac else float("nan")
    diferida = float(a1["Cobranças (caixa)"].sum()) - rec_a1
    return {"tem": True, "mrr1": mrr1, "mrr12": mrr12, "arr12": mrr12 * 12, "ativos12": ativos12, "arpu": arpu,
            "churn": churn, "vida_meses": vida_meses, "margem_bruta": margem_bruta, "ltv": ltv, "cac": cac,
            "gasto_aquisicao_m1": gasto_aquisicao_m1, "novos1": novos1, "ltv_cac": ltv_cac,
            "payback_cac": payback_cac, "receita_diferida_a1": diferida, "rec_a1": rec_a1,
            "cobrancas_a1": float(a1["Cobranças (caixa)"].sum()),
            "crescimento_mrr": div(mrr12, mrr1) - 1 if mrr1 > 0 else float("nan")}


# ------------------------------------------------------------------------------
# 3c. SIMPLES NACIONAL (Anexos I a V, LC 123/2006 com redação da LC 155/2016)
# ------------------------------------------------------------------------------
TABELAS_SIMPLES = {
    "Anexo I - Comércio": [(180000, .04, 0), (360000, .073, 5940), (720000, .095, 13860), (1800000, .107, 22500),
                           (3600000, .143, 87300), (4800000, .19, 378000)],
    "Anexo II - Indústria": [(180000, .045, 0), (360000, .078, 5940), (720000, .10, 13860), (1800000, .112, 22500),
                             (3600000, .147, 85500), (4800000, .30, 720000)],
    "Anexo III - Serviços": [(180000, .06, 0), (360000, .112, 9360), (720000, .135, 17640), (1800000, .16, 35640),
                             (3600000, .21, 125640), (4800000, .33, 648000)],
    "Anexo IV - Serviços": [(180000, .045, 0), (360000, .09, 8100), (720000, .102, 12420), (1800000, .14, 39780),
                            (3600000, .22, 183780), (4800000, .33, 828000)],
    "Anexo V - Serviços": [(180000, .155, 0), (360000, .18, 4500), (720000, .195, 9900), (1800000, .205, 17100),
                           (3600000, .23, 62100), (4800000, .305, 540000)],
}


def aliquota_simples(anexo: str, rbt12: float) -> dict:
    """Alíquota efetiva = (RBT12 x alíquota nominal - parcela a deduzir) / RBT12."""
    tab = TABELAS_SIMPLES.get(anexo)
    if not tab or rbt12 <= 0:
        return {"faixa": 0, "nominal": float("nan"), "deduzir": 0.0, "efetiva": float("nan")}
    for i, (lim, nom, ded) in enumerate(tab, start=1):
        if rbt12 <= lim:
            return {"faixa": i, "nominal": nom, "deduzir": ded, "efetiva": (rbt12 * nom - ded) / rbt12}
    lim, nom, ded = tab[-1]
    return {"faixa": 7, "nominal": nom, "deduzir": ded, "efetiva": (rbt12 * nom - ded) / rbt12, "acima_limite": True}


def fator_r(res: dict, d: dict) -> dict:
    """Fator R = folha de salários dos últimos 12 meses (com pró-labore e encargos) ÷ receita bruta."""
    cf = d["custos_fixos"]
    pro = float(cf.loc[cf["descricao"].astype(str).str.lower().str.replace("-", "").str.replace(" ", "")
                       .str.contains("prolabore|prólabore"), "valor_mensal"].sum())
    folha12 = res.get("folha_ano1", (res["salarios_m"] + res["encargos_m"]) * 12) + pro * 12
    rec = res["anos"][1]["rec_bruta"]
    fr = div(folha12, rec) if rec > 0 else float("nan")
    return {"folha12": folha12, "pro_labore_m": pro, "receita12": rec, "fator_r": fr,
            "anexo_por_fator_r": "Anexo III - Serviços" if fr == fr and fr >= 0.28 else "Anexo V - Serviços"}


# ------------------------------------------------------------------------------
# 4. MOTOR PRINCIPAL
# ------------------------------------------------------------------------------
def executar_motor(proj: dict, d: dict) -> dict:
    """d = dicionário com os DataFrames do projeto (nomes das tabelas)."""
    cat = custos_unitarios(d["catalogo"], d["materiais"], d["ficha_tecnica"])
    aliq = proj["aliquota_imposto_pct"] / 100.0
    var_pct = float(d["custos_variaveis"]["percentual"].sum()) / 100.0

    # --- Receita e custos diretos do Mês 1, por categoria
    receita_m1 = float(cat["receita_mensal"].sum())
    cdir_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "custo_mensal"].sum()) for c in CATEGORIAS}
    rec_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "receita_mensal"].sum()) for c in CATEGORIAS}
    cdir_m1 = sum(cdir_cat_m1.values())

    # --- Horizonte de planejamento escolhido pelo usuário (1 a 10 anos)
    N = int(min(max(round(float(proj.get("periodo_anos") or 3)), 1), 10))
    ANOS = list(range(1, N + 1))
    MESES = 12 * N

    # --- Assinaturas (receita recorrente), todos os meses do horizonte
    g_m = proj["crescimento_mensal_ano1_pct"] / 100
    g_a = proj["crescimento_anual_pct"] / 100
    inf = proj["inflacao_custos_anual_pct"] / 100
    ass_df, ass_planos = simular_assinaturas(d.get("assinaturas"), g_a, inf, meses=MESES)
    tem_ass = not ass_df.empty

    def serie(col):
        return [float(v) for v in ass_df[col]] if tem_ass else [0.0] * MESES
    s_rec, s_cob, s_custo, s_taxa = (serie("Receita reconhecida"), serie("Cobranças (caixa)"),
                                     serie("Custo de entrega"), serie("Taxas de cobrança"))

    # --- Pessoal, custos fixos e depreciação (mensais)
    pes_todos = d["pessoal"]
    inicio = (pes_todos["mes_inicio"] if "mes_inicio" in pes_todos else pd.Series([1.0] * len(pes_todos))).fillna(1).clip(lower=1)
    linha_total = (pes_todos["num_empregados"] * pes_todos["salario_mensal"] * (1 + pes_todos["encargos_pct"] / 100) +
                   pes_todos["num_empregados"] * pes_todos["beneficios_unit"])
    linha_sal_enc = pes_todos["num_empregados"] * pes_todos["salario_mensal"] * (1 + pes_todos["encargos_pct"] / 100)
    # custo de pessoal mês a mês (sem inflação): cada cargo entra no seu mês de início
    pes_serie = [float(linha_total[inicio <= t + 1].sum()) for t in range(MESES)]
    sal_enc_serie = [float(linha_sal_enc[inicio <= t + 1].sum()) for t in range(MESES)]
    pes = pes_todos[inicio <= 1]                                    # equipe do Mês 1 (custeio unitário e custo da hora)
    sal = pes["num_empregados"] * pes["salario_mensal"]
    salarios_m = float(sal.sum())
    encargos_m = float((sal * pes["encargos_pct"] / 100).sum())
    beneficios_m = float((pes["num_empregados"] * pes["beneficios_unit"]).sum())
    pessoal_m = salarios_m + encargos_m + beneficios_m
    fixos_m = float(d["custos_fixos"]["valor_mensal"].sum())
    pes_total_linha = sal * (1 + pes["encargos_pct"] / 100) + pes["num_empregados"] * pes["beneficios_unit"]
    area = pes["area"].fillna("") if "area" in pes else pd.Series([""] * len(pes))
    mod_m = float(pes_total_linha[area.str.startswith("Produ")].sum())          # mão de obra direta (produção)
    pessoal_adm_m = float(pes_total_linha[area.str.startswith("Admin")].sum())
    pessoal_com_m = pessoal_m - mod_m - pessoal_adm_m
    cfx = d["custos_fixos"]
    cls = cfx["classificacao"].fillna("") if "classificacao" in cfx else pd.Series([""] * len(cfx))
    cif_m = float(cfx.loc[cls.str.startswith("Custo"), "valor_mensal"].sum())   # indiretos de produção e de tecnologia
    desp_adm_m = float(cfx.loc[cls.str.contains("administrativa") | (cls == ""), "valor_mensal"].sum())
    desp_com_m = float(cfx.loc[cls.str.contains("comercial"), "valor_mensal"].sum())
    desp_fin_m = float(cfx.loc[cls.str.contains("financeira"), "valor_mensal"].sum())

    inv = d["inv_fixos"].copy()
    inv["total"] = inv["quantidade"] * inv["valor_unitario"]
    inv["dep_anual"] = [
        (t * (1 - res / 100) / v) if v > 0 else 0.0
        for t, res, v in zip(inv["total"], inv["valor_residual_pct"], inv["vida_util_anos"])
    ]
    inv["dep_mensal"] = inv["dep_anual"] / 12
    dep_anual = float(inv["dep_anual"].sum())
    dep_m = dep_anual / 12
    cf_desemb_m = pessoal_m + fixos_m

    # --- Investimentos
    inv_fixo = float(inv["total"].sum())
    inv_pre = float(d["inv_pre"]["valor"].sum())
    mat = d["materiais"]
    est_mat = float((mat["qtd_estoque_inicial"] * mat["valor_unitario"]).sum())
    rev = cat[cat["categoria"] == "Revenda"]
    est_rev = float((rev["qtd_estoque_inicial"] * rev["custo_unit"]).sum())
    inv_estoque = est_mat + est_rev

    # --- Prazos médios e capital de giro
    pz = d["capital_giro_prazos"]
    rec_pz, pag_pz = pz[pz["tipo"] == "Recebimento"], pz[pz["tipo"] == "Pagamento"]
    pmr = float((rec_pz["prazo_dias"] * rec_pz["percentual"] / 100).sum())
    pmp = float((pag_pz["prazo_dias"] * pag_pz["percentual"] / 100).sum())
    pme = float(proj["pme_dias"])
    ciclo_op = pme + pmr
    ciclo_fin = ciclo_op - pmp
    desembolso_m1 = (cdir_m1 + (receita_m1 + s_rec[0]) * (aliq + var_pct) + cf_desemb_m + s_custo[0] + s_taxa[0])
    ncg = max(0.0, desembolso_m1 / 30.0 * ciclo_fin)
    reserva = float(proj["reserva_caixa_extra"])
    capital_giro = ncg + reserva
    investimento_total = inv_fixo + inv_pre + inv_estoque + capital_giro

    # --- Fontes
    capital_proprio = float(d["socios"]["capital_integralizado"].sum())
    capital_terceiros = float(d["emprestimos"]["valor"].sum())
    fontes = capital_proprio + capital_terceiros
    crono = cronograma_emprestimos(d["emprestimos"], anos=N + 1)

    # --- Projeção mensal do Ano 1
    tma = proj["tma_anual_pct"] / 100
    meses = []
    for m in range(12):
        f = (1 + g_m) ** m
        rec = receita_m1 * f + s_rec[m]
        cv = cdir_m1 * f + rec * (aliq + var_pct) + s_custo[m] + s_taxa[m]
        cf = pes_serie[m] + fixos_m + dep_m
        meses.append({"Mês": f"Mês {m + 1}", "Receita": rec, "Custos variáveis": cv,
                      "Margem de contribuição": rec - cv, "Custos fixos": cf,
                      "Resultado operacional": rec - cv - cf})
    df_meses = pd.DataFrame(meses)
    fator_a1 = sum((1 + g_m) ** m for m in range(12))

    # --- DRE e balanço dos Anos 1 a 3
    anos = {}
    lucros_acum = 0.0
    for k in ANOS:
        fr = fator_a1 * (1 + g_a) ** (k - 1)
        fi = (1 + inf) ** (k - 1)
        meses_k = slice((k - 1) * 12, k * 12)
        ass_rec, ass_custo, ass_taxa = sum(s_rec[meses_k]), sum(s_custo[meses_k]), sum(s_taxa[meses_k])
        rec_prod = receita_m1 * fr
        rec_bruta = rec_prod + ass_rec
        rec_cat = {c: v * fr for c, v in rec_cat_m1.items()}
        rec_cat["Assinatura"] = ass_rec
        cmv, cpv, csv = (cdir_cat_m1["Revenda"] * fr, cdir_cat_m1["Fabricação Própria"] * fr,
                         cdir_cat_m1["Serviço"] * fr + ass_custo)
        custo_dir = cmv + cpv + csv
        impostos = rec_bruta * aliq
        outras_var = rec_bruta * var_pct + ass_taxa
        rec_liquida = rec_bruta - impostos
        cv_tot = custo_dir + impostos + outras_var
        mc = rec_bruta - cv_tot
        imc = div(mc, rec_bruta)
        pessoal = sum(pes_serie[meses_k]) * fi
        fixos = fixos_m * 12 * fi
        cf_desemb = pessoal + fixos
        ebitda = mc - cf_desemb
        ebit = ebitda - dep_anual
        juros = crono["juros"][k - 1]
        resultado = ebit - juros
        lucros_acum += resultado
        fco = resultado + dep_anual
        amort = crono["amort"][k - 1]
        fluxo_socio = fco - amort

        # Balanço simplificado (caixa calculado por diferença: Ativo = Passivo + PL)
        receber = rec_bruta / 360 * pmr
        estoques = inv_estoque * (rec_prod / (receita_m1 * fator_a1)) if receita_m1 > 0 else inv_estoque
        imobilizado = max(0.0, inv_fixo - dep_anual * k)
        fornecedores = custo_dir / 360 * pmp
        obrigacoes = (pessoal + impostos) / 12
        saldo_div = crono["saldo_fim"][k - 1]
        div_cp = crono["amort"][k] if k < len(crono["amort"]) else 0.0
        div_cp = min(div_cp, saldo_div)
        div_lp = saldo_div - div_cp
        # Receita antecipada (assinaturas cobradas e ainda não prestadas): passivo circulante
        receita_diferida = sum(s_cob[:k * 12]) - sum(s_rec[:k * 12])
        pc = fornecedores + obrigacoes + div_cp + receita_diferida
        pnc = div_lp
        pl = capital_proprio + lucros_acum
        caixa = pc + pnc + pl - (receber + estoques + imobilizado + inv_pre)
        ac = caixa + receber + estoques
        ativo_total = ac + imobilizado + inv_pre

        cf_total = cf_desemb + dep_anual
        pe_cont = div(cf_total, imc) if imc and imc > 0 else float("nan")
        pe_fin = div(cf_desemb, imc) if imc and imc > 0 else float("nan")
        pe_eco = div(cf_total + investimento_total * tma, imc) if imc and imc > 0 else float("nan")

        anos[k] = dict(
            rec_bruta=rec_bruta, rec_cat=rec_cat, impostos=impostos, rec_liquida=rec_liquida,
            cmv=cmv, cpv=cpv, csv=csv, custo_dir=custo_dir, outras_var=outras_var,
            cv_tot=cv_tot, mc=mc, imc=imc, pessoal=pessoal, fixos=fixos, cf_desemb=cf_desemb,
            dep=dep_anual, cf_total=cf_total, ebitda=ebitda, ebit=ebit, juros=juros,
            resultado=resultado, fco=fco, amort=amort, fluxo_socio=fluxo_socio,
            caixa=caixa, receber=receber, estoques=estoques, ac=ac, imobilizado=imobilizado,
            ativo_total=ativo_total, fornecedores=fornecedores, obrigacoes=obrigacoes, receita_diferida=receita_diferida,
            div_cp=div_cp, div_lp=div_lp, pc=pc, pnc=pnc, pl=pl, divida_bruta=saldo_div,
            divida_liquida=saldo_div - caixa, pe_contabil=pe_cont, pe_financeiro=pe_fin,
            pe_economico=pe_eco,
        )

    # --- Avaliação do projeto (fluxo do projeto = EBITDA, investimento no ano 0)
    fluxos = [-investimento_total] + [anos[k]["ebitda"] for k in ANOS]
    if proj.get("recuperar_giro"):
        fluxos[N] += capital_giro
    vpl = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos))
    tir = calcular_tir(fluxos)
    vp_entradas = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos) if t > 0)
    il = div(vp_entradas, investimento_total)

    lucro_mensal_a1 = anos[1]["resultado"] / 12
    payback_simples = div(investimento_total, lucro_mensal_a1) if lucro_mensal_a1 > 0 else float("nan")
    acum, payback_desc = -investimento_total, float("nan")
    for t in ANOS:
        fd = fluxos[t] / (1 + tma) ** t
        if fd > 0 and acum + fd >= 0:
            payback_desc = ((t - 1) + (-acum) / fd) * 12
            break
        acum += fd

    a1 = anos[1]
    resumo = {
        "receita_anual": a1["rec_bruta"], "custos_variaveis": a1["cv_tot"], "margem_contrib": a1["mc"],
        "custos_fixos": a1["cf_total"], "resultado_operacional": a1["mc"] - a1["cf_total"],
        "lucratividade": div(a1["mc"] - a1["cf_total"], a1["rec_bruta"]),
        "rentabilidade": div(a1["mc"] - a1["cf_total"], investimento_total),
        "payback_meses": div(investimento_total, (a1["mc"] - a1["cf_total"]) / 12)
        if a1["mc"] - a1["cf_total"] > 0 else float("nan"),
    }

    # --- Custeio por absorção: MOD + custos indiretos de produção + depreciação rateados
    overhead_m = mod_m + cif_m + dep_m
    criterio = proj.get("rateio_criterio") or "Custo direto"
    prod_mask = cat["categoria"] != "Revenda"
    base = {"Custo direto": cat["custo_mensal"], "Receita": cat["receita_mensal"],
            "Quantidade": cat["qtd_mensal"]}.get(criterio, cat["custo_mensal"]).where(prod_mask, 0.0)
    tot_base = float(base.sum())
    cat["rateio_mensal"] = base / tot_base * overhead_m if tot_base > 0 else 0.0
    cat["custo_absorcao_unit"] = [cu + (r / q if q > 0 else 0.0) for cu, r, q in
                                  zip(cat["custo_unit"], cat["rateio_mensal"], cat["qtd_mensal"])]
    cat["margem_absorcao_pct"] = [div(p - c, p) for p, c in zip(cat["preco_venda"], cat["custo_absorcao_unit"])]

    # --- Formação de preço (markup divisor)
    margem_desej = float(proj.get("margem_desejada_pct") or 0) / 100
    df_pct = div(cf_desemb_m + dep_m, receita_m1 + s_rec[0]) if receita_m1 + s_rec[0] > 0 else 0.0
    df_pct = 0.0 if df_pct != df_pct else df_pct
    divisor = 1 - (aliq + var_pct + df_pct + margem_desej)
    markup_mult = div(1, divisor) if divisor > 0 else float("nan")
    cat["preco_sugerido"] = [c * markup_mult if divisor > 0 else float("nan") for c in cat["custo_unit"]]
    cat["dif_preco_pct"] = [div(p - s_, s_) for p, s_ in zip(cat["preco_venda"], cat["preco_sugerido"])]

    # --- Ponto de equilíbrio em unidades por produto (mix atual de vendas)
    pe_m = anos[1]["pe_contabil"] / 12 if anos[1]["pe_contabil"] == anos[1]["pe_contabil"] else float("nan")
    cat["pe_unidades_mes"] = [div(pe_m * div(r, receita_m1), p) if receita_m1 > 0 else float("nan")
                              for r, p in zip(cat["receita_mensal"], cat["preco_venda"])]

    # --- Fluxo de caixa mensal do Ano 1
    def distribuir(valores, prazos_df):
        """Distribui cada valor mensal no tempo conforme os prazos (dias) e percentuais."""
        saida = [0.0] * 14
        regras = [(float(p), float(q) / 100) for p, q in zip(prazos_df["prazo_dias"], prazos_df["percentual"])] \
            or [(0.0, 1.0)]
        for m, v in enumerate(valores):
            for dias, frac in regras:
                lag, resto = int(dias // 30), (dias % 30) / 30
                if m + lag < 14:
                    saida[m + lag] += v * frac * (1 - resto)
                if m + lag + 1 < 14:
                    saida[m + lag + 1] += v * frac * resto
        return saida[:12]

    rec_mes = [receita_m1 * (1 + g_m) ** m for m in range(12)]
    cdir_mes = [cdir_m1 * (1 + g_m) ** m for m in range(12)]
    entradas = [e + c for e, c in zip(distribuir(rec_mes, rec_pz), s_cob[:12])]   # assinaturas: caixa na cobrança
    pag_forn = distribuir(cdir_mes, pag_pz)
    saldo = fontes - inv_fixo - inv_pre - inv_estoque
    fluxo_m = []
    for m in range(12):
        sai_var = (rec_mes[m] + s_rec[m]) * (aliq + var_pct) + s_custo[m] + s_taxa[m]
        parcela = crono["juros_m"][m] + crono["amort_m"][m]
        sai = pag_forn[m] + sai_var + pes_serie[m] + fixos_m + parcela
        ini = saldo
        saldo = saldo + entradas[m] - sai
        fluxo_m.append({"Mês": f"Mês {m + 1}", "Saldo inicial": ini, "Recebimentos": entradas[m],
                        "Fornecedores": pag_forn[m], "Impostos e desp. variáveis": sai_var,
                        "Pessoal e custos fixos": pes_serie[m] + fixos_m, "Parcelas de empréstimos": parcela,
                        "Saldo final": saldo})
    df_fluxo_m = pd.DataFrame(fluxo_m)

    # --- Indicadores de recorrência (CAC = gasto comercial fixo do Mês 1 ÷ novos assinantes do Mês 1)
    verba = float(proj.get("verba_aquisicao_mensal") or 0)
    rec_ind = metricas_recorrencia(ass_df, ass_planos, d.get("assinaturas"), verba if verba > 0 else desp_com_m + pessoal_com_m)
    if rec_ind.get("tem"):
        rec_ind["origem_cac"] = "verba de aquisição informada" if verba > 0 else "gasto comercial fixo do Mês 1"

    return dict(
        mod_m=mod_m, pessoal_adm_m=pessoal_adm_m, pessoal_com_m=pessoal_com_m, cif_m=cif_m,
        desp_adm_m=desp_adm_m, desp_com_m=desp_com_m, desp_fin_m=desp_fin_m, overhead_m=overhead_m,
        criterio_rateio=criterio, margem_desejada=margem_desej, df_pct=df_pct, markup_mult=markup_mult,
        divisor_markup=divisor, fluxo_mensal=df_fluxo_m, g_m=g_m,
        cat=cat, inv=inv, receita_m1=receita_m1, cdir_m1=cdir_m1, var_pct=var_pct, aliq=aliq,
        salarios_m=salarios_m, encargos_m=encargos_m, beneficios_m=beneficios_m, pessoal_m=pessoal_m,
        fixos_m=fixos_m, dep_m=dep_m, dep_anual=dep_anual, inv_fixo=inv_fixo, inv_pre=inv_pre,
        est_mat=est_mat, est_rev=est_rev, inv_estoque=inv_estoque, pmr=pmr, pmp=pmp, pme=pme,
        ciclo_op=ciclo_op, ciclo_fin=ciclo_fin, desembolso_m1=desembolso_m1, ncg=ncg,
        reserva=reserva, capital_giro=capital_giro, investimento_total=investimento_total,
        capital_proprio=capital_proprio, capital_terceiros=capital_terceiros, fontes=fontes,
        crono=crono, meses=df_meses, anos=anos, fluxos=fluxos, vpl=vpl, tir=tir, il=il,
        payback_simples=payback_simples, payback_desc=payback_desc, resumo=resumo, tma=tma,
        ass=ass_df, ass_planos=ass_planos, recorrencia=rec_ind, g_a=g_a, inf=inf,
        N=N, ANOS=ANOS, MESES=MESES, pes_serie=pes_serie, folha_ano1=sum(sal_enc_serie[:12]),
        recupera_giro=bool(proj.get("recuperar_giro")),
    )


# ------------------------------------------------------------------------------
# 5. CENÁRIOS (estilo "Simulador de Resultados" do PNBOX)
# ------------------------------------------------------------------------------
def cenarios(res: dict, proj: dict) -> pd.DataFrame:
    a1 = res["anos"][1]
    inv, tma = res["investimento_total"], res["tma"]

    def calc(fr, fc):
        rec, cv, cf = a1["rec_bruta"] * fr, a1["cv_tot"] * fc, a1["cf_total"] * fc
        mc = rec - cv
        r = mc - cf
        imc = div(mc, rec)
        ok = imc == imc and imc > 0
        return [brl(rec), brl(cv), brl(mc), brl(cf), brl(r), pct(div(r, rec)), pct(div(r, inv)),
                f"{num_br(div(inv, r / 12), 1)} meses" if r > 0 else "Não recupera",
                brl(cf / imc) if ok else "n.a.", brl((cf - a1["dep"] * fc) / imc) if ok else "n.a.",
                brl((cf + inv * tma) / imc) if ok else "n.a."]

    rotulos = ["Receita anual", "Custos variáveis", "Margem de contribuição", "Custos fixos",
               "Resultado operacional", "Lucratividade", "Rentabilidade", "Payback simples",
               "PE contábil", "PE financeiro", "PE econômico"]
    ro, co = proj["cen_otim_receita_pct"] / 100, proj["cen_otim_custo_pct"] / 100
    rp, cp = proj["cen_pess_receita_pct"] / 100, proj["cen_pess_custo_pct"] / 100
    return pd.DataFrame({
        "Indicador (Ano 1)": rotulos,
        f"Pessimista (-{num_br(rp*100,0)}% rec. / +{num_br(cp*100,0)}% custos)": calc(1 - rp, 1 + cp),
        "Provável": calc(1, 1),
        f"Otimista (+{num_br(ro*100,0)}% rec. / -{num_br(co*100,0)}% custos)": calc(1 + ro, 1 - co),
    })


# ------------------------------------------------------------------------------
# 6. DRE EM TABELA
# ------------------------------------------------------------------------------
def tabela_dre(res: dict) -> pd.DataFrame:
    A = res["anos"]
    linhas = [
        ("1. Receita operacional bruta", "rec_bruta", 1),
        ("   (-) Impostos sobre vendas", "impostos", -1),
        ("2. (=) Receita líquida", "rec_liquida", 1),
        ("   (-) CMV - mercadorias de revenda", "cmv", -1),
        ("   (-) CPV - produtos de fabricação própria", "cpv", -1),
        ("   (-) CSV - serviços vendidos", "csv", -1),
        ("   (-) Outras despesas variáveis", "outras_var", -1),
        ("3. (=) Margem de contribuição", "mc", 1),
        ("   (-) Pessoal (salários, encargos e benefícios)", "pessoal", -1),
        ("   (-) Demais custos fixos", "fixos", -1),
        ("4. (=) EBITDA", "ebitda", 1),
        ("   (-) Depreciação", "dep", -1),
        ("5. (=) Resultado operacional (EBIT)", "ebit", 1),
        ("   (-) Juros dos empréstimos", "juros", -1),
        ("6. (=) Resultado líquido do exercício", "resultado", 1),
    ]
    rows = []
    for nome, k, s in linhas:
        lin = {"Conta": nome}
        for a in res["ANOS"]:
            lin[f"Ano {a}"] = brl(A[a][k] * s)
        lin["% Receita (Ano 1)"] = pct(div(A[1][k], A[1]["rec_bruta"]))
        rows.append(lin)
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------------
# 7. INDICADORES EM 5 BLOCOS
# ------------------------------------------------------------------------------
def tabela_indicadores(res: dict) -> pd.DataFrame:
    A = res["anos"]
    inv = res["investimento_total"]
    linhas = []

    AN = res["ANOS"]

    def add(bloco, nome, formula, vals, fmt, leitura, var="pct"):
        f = {"brl": brl, "pct": pct, "x": lambda v: "n.a." if v != v else f"{num_br(v)}x"}[fmt]
        lin = {"Bloco": bloco, "Indicador": nome, "Fórmula": formula}
        for i, a in enumerate(AN):
            lin[f"Ano {a}"] = f(vals[i])
        if len(AN) >= 2:
            v1, v2 = vals[0], vals[1]
            if var == "pp":
                vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br((v2 - v1) * 100):s} p.p."
            elif var == "dif":
                vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br(v2 - v1)}x"
            else:
                vr = "n.a." if not v1 or v1 != v1 else f"{num_br((v2 - v1) / abs(v1) * 100, 1)}%"
            lin["Var. Ano 2/Ano 1"] = vr
        lin["Leitura"] = leitura
        linhas.append(lin)

    g = lambda k: [A[a][k] for a in AN]
    r = lambda n, d_: [div(A[a][n], A[a][d_]) for a in AN]

    b = "1. Rentabilidade e resultado"
    add(b, "Receita operacional bruta", "Σ preço × quantidade", g("rec_bruta"), "brl", "Faturamento antes das deduções.")
    add(b, "Margem de contribuição", "Receita − custos e despesas variáveis", g("mc"), "brl", "Valor que sobra para pagar os custos fixos.")
    add(b, "Índice de margem de contribuição", "MC ÷ Receita bruta", g("imc"), "pct", "Parcela de cada R$ 1,00 vendido que cobre a estrutura fixa.", "pp")
    add(b, "EBITDA", "MC − custos fixos desembolsáveis", g("ebitda"), "brl", "Geração operacional de caixa antes de juros e depreciação.")
    add(b, "Margem EBITDA", "EBITDA ÷ Receita líquida", r("ebitda", "rec_liquida"), "pct", "Eficiência operacional.", "pp")
    add(b, "Resultado líquido", "EBIT − juros", g("resultado"), "brl", "Lucro (ou prejuízo) do exercício.")
    add(b, "Lucratividade líquida", "Resultado líquido ÷ Receita bruta", r("resultado", "rec_bruta"), "pct", "Ganho para cada real faturado.", "pp")
    add(b, "Rentabilidade (ROI)", "Resultado líquido ÷ Investimento total", [div(A[a]["resultado"], inv) for a in AN], "pct", "Retorno contábil sobre o capital investido.", "pp")
    add(b, "ROE", "Resultado líquido ÷ Patrimônio líquido", r("resultado", "pl"), "pct", "Retorno sobre o capital dos sócios.", "pp")

    b = "2. Endividamento e estrutura de capital"
    add(b, "Dívida bruta", "Saldo dos empréstimos no fim do ano", g("divida_bruta"), "brl", "Obrigações financeiras onerosas.")
    add(b, "Dívida líquida", "Dívida bruta − caixa", g("divida_liquida"), "brl", "Negativa indica caixa maior que a dívida.")
    add(b, "Endividamento geral", "(PC + PNC) ÷ Ativo total", [div(A[a]["pc"] + A[a]["pnc"], A[a]["ativo_total"]) for a in AN], "pct", "Parcela do ativo financiada por terceiros.", "pp")
    add(b, "Dívida líquida ÷ EBITDA", "Dívida líquida ÷ EBITDA", [div(A[a]["divida_liquida"], A[a]["ebitda"]) if A[a]["ebitda"] > 0 else float("nan") for a in AN], "x", "Anos de geração operacional para quitar a dívida.", "dif")
    add(b, "Cobertura de juros", "EBITDA ÷ juros", [div(A[a]["ebitda"], A[a]["juros"]) if A[a]["juros"] > 0 else float("nan") for a in AN], "x", "Abaixo de 1,0x a operação não paga os juros.", "dif")

    b = "3. Liquidez e capital de giro"
    add(b, "Liquidez corrente", "AC ÷ PC", r("ac", "pc"), "x", "Acima de 1,0x indica folga de curto prazo.", "dif")
    add(b, "Liquidez seca", "(AC − estoques) ÷ PC", [div(A[a]["ac"] - A[a]["estoques"], A[a]["pc"]) for a in AN], "x", "Capacidade de pagamento sem vender estoques.", "dif")
    add(b, "Liquidez imediata", "Caixa ÷ PC", r("caixa", "pc"), "x", "Disponibilidade imediata.", "dif")
    add(b, "Capital circulante líquido", "AC − PC", [A[a]["ac"] - A[a]["pc"] for a in AN], "brl", "Folga (ou falta) de recursos de curto prazo.")
    add(b, "Caixa projetado", "Ativo − demais contas (por diferença)", g("caixa"), "brl", "Negativo indica necessidade de mais recursos.")

    b = "4. Viabilidade do investimento"
    add(b, "Custo com pessoal ÷ Receita líquida", "Folha total ÷ Receita líquida", r("pessoal", "rec_liquida"), "pct", "Peso da mão de obra.", "pp")
    add(b, "Fluxo de caixa operacional (FCO)", "Resultado líquido + depreciação", g("fco"), "brl", "Caixa gerado pelas operações.")
    add(b, "Fluxo de caixa do sócio", "FCO − amortização dos empréstimos", g("fluxo_socio"), "brl", "Caixa que sobra após pagar o principal das dívidas.")

    b = "5. Pontos de equilíbrio"
    add(b, "PE contábil (anual)", "Custos fixos totais ÷ IMC", g("pe_contabil"), "brl", "Receita para lucro zero.")
    add(b, "PE financeiro (anual)", "(Custos fixos − depreciação) ÷ IMC", g("pe_financeiro"), "brl", "Receita para cobrir os desembolsos.")
    add(b, "PE econômico (anual)", "(Custos fixos + TMA × investimento) ÷ IMC", g("pe_economico"), "brl", "Receita para remunerar o capital à TMA.")
    add(b, "Margem de segurança", "(Receita − PE contábil) ÷ Receita", [div(A[a]["rec_bruta"] - A[a]["pe_contabil"], A[a]["rec_bruta"]) for a in AN], "pct", "Quanto a receita pode cair antes do prejuízo.", "pp")
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 8. CHECAGENS DE CONSISTÊNCIA
# ------------------------------------------------------------------------------
def tabela_checagens(res: dict, d: dict) -> pd.DataFrame:
    cat = res["cat"]
    A = res["anos"]
    chk = []

    def add(t, valor, ok, obs):
        chk.append({"Teste": t, "Valor apurado": valor, "Situação": "OK" if ok else "VERIFICAR", "Observação": obs})

    dif = res["fontes"] - res["investimento_total"]
    add("Fontes de recursos x investimento total", f"Diferença: {brl(dif)}", abs(dif) <= 1,
        "Capital dos sócios + empréstimos deve igualar o investimento total.")
    neg = cat[cat["margem_unit"] <= 0]["descricao"].tolist()
    add("Margem unitária positiva", f"{len(neg)} item(ns) com margem <= 0", not neg,
        ("Itens: " + ", ".join(neg[:5])) if neg else "Todos os itens têm preço maior que o custo.")
    sem = cat[cat["custo_unit"] <= 0]["descricao"].tolist()
    add("Produtos com custo cadastrado", f"{len(sem)} item(ns) sem custo", not sem,
        ("Sem ficha técnica/custo: " + ", ".join(sem[:5])) if sem else "Todos os itens têm custo.")
    mats = set(d["materiais"]["nome"])
    orf = sorted(set(d["ficha_tecnica"]["material"]) - mats)
    add("Fichas técnicas com materiais válidos", f"{len(orf)} material(is) não cadastrado(s)", not orf,
        ", ".join(orf[:5]) if orf else "Todos os insumos existem no cadastro de materiais.")
    for tipo in ("Recebimento", "Pagamento"):
        s = d["capital_giro_prazos"].query("tipo == @tipo")["percentual"].sum()
        add(f"Prazos de {tipo.lower()} somam 100%", f"{num_br(s)}%", abs(s - 100) < 0.01,
            "A soma dos percentuais de cada tipo de prazo deve ser 100%.")
    sw = d["swot"]
    incoer = sw[((sw["ambiente"] == "Interno") & sw["natureza"].isin(["Oportunidade", "Ameaça"])) |
                ((sw["ambiente"] == "Externo") & sw["natureza"].isin(["Força", "Fraqueza"]))]
    add("SWOT: ambiente coerente com a classificação", f"{len(incoer)} fator(es) incoerente(s)", incoer.empty,
        "Interno = Força/Fraqueza; Externo = Oportunidade/Ameaça.")
    neg_cx = [k for k in res["ANOS"] if A[k]["caixa"] < 0]
    add("Caixa projetado não negativo", "Anos com caixa negativo: " + (", ".join(map(str, neg_cx)) or "nenhum"),
        not neg_cx, "Caixa negativo indica que as fontes não cobrem a operação.")
    lc = div(A[1]["ac"], A[1]["pc"])
    add("Liquidez corrente do Ano 1 >= 1,00", f"{num_br(lc)}x", lc == lc and lc >= 1, "Abaixo de 1,00 há aperto de curto prazo.")
    add("VPL positivo", brl(res["vpl"]), res["vpl"] > 0, "VPL positivo indica criação de valor acima da TMA.")
    return pd.DataFrame(chk)


# ------------------------------------------------------------------------------
# 9. DIAGNÓSTICO SWOT (matriz de estratégias)
# ------------------------------------------------------------------------------
QUADRANTES = {
    "FO": ("Estratégia de Desenvolvimento", "Forças + Oportunidades",
           "Cenário mais favorável: aproveitar oportunidades e expandir (novos mercados, produtos e parcerias)."),
    "FrO": ("Estratégia de Crescimento", "Fraquezas + Oportunidades",
            "Há oportunidades, mas fraquezas internas limitam o aproveitamento: corrigir pontos fracos para crescer."),
    "FA": ("Estratégia de Manutenção", "Forças + Ameaças",
           "Usar as forças para enfrentar as ameaças e manter a posição conquistada."),
    "FrA": ("Estratégia de Sobrevivência", "Fraquezas + Ameaças",
            "Situação mais delicada: reduzir custos, rever o modelo e proteger o caixa."),
}


def pontuacao_swot(sw: pd.DataFrame) -> dict:
    peso = lambda nat: int(sum(2 if i == "Alto" else 1 for i in sw.loc[sw["natureza"] == nat, "impacto"]))
    F, Fr, O, A = peso("Força"), peso("Fraqueza"), peso("Oportunidade"), peso("Ameaça")
    q = {"FO": F + O, "FrO": Fr + O, "FA": F + A, "FrA": Fr + A}
    pred = max(q, key=q.get) if any(q.values()) else None
    return {"F": F, "Fr": Fr, "O": O, "A": A, "quadrantes": q, "predominante": pred}


# ------------------------------------------------------------------------------
# 10. COMPOSIÇÕES PARA GRÁFICOS
# ------------------------------------------------------------------------------
def _serie(dic: dict) -> pd.DataFrame:
    df = pd.DataFrame({"Item": list(dic.keys()), "Valor": [float(v) for v in dic.values()]})
    df = df[df["Valor"] > 0.005].sort_values("Valor", ascending=False).reset_index(drop=True)
    tot = df["Valor"].sum()
    df["%"] = df["Valor"] / tot if tot else 0.0
    return df


def composicoes(res: dict, d: dict) -> dict:
    mat = d["materiais"]
    cat_mat = {n: (c or "Não classificado") for n, c in zip(mat["nome"], mat["categoria"])}
    preco_mat = dict(zip(mat["nome"], mat["valor_unitario"]))
    out = {}

    out["investimento"] = _serie({
        "Ativo imobilizado (investimentos fixos)": res["inv_fixo"], "Investimentos pré-operacionais": res["inv_pre"],
        "Estoque inicial": res["inv_estoque"], "Capital de giro": res["capital_giro"]})

    est = {}
    for _, r in mat.iterrows():
        k = r["categoria"] or "Não classificado"
        est[k] = est.get(k, 0) + r["qtd_estoque_inicial"] * r["valor_unitario"]
    est["Mercadorias para revenda"] = res["est_rev"]
    out["estoque"] = _serie(est)

    ficha = d["ficha_tecnica"]
    custo = {}
    for _, p in res["cat"].iterrows():
        if p["qtd_mensal"] <= 0:
            continue
        if p["origem_custo"] == "Ficha técnica":
            for _, f in ficha[ficha["produto"] == p["descricao"]].iterrows():
                k = cat_mat.get(f["material"], "Não classificado")
                custo[k] = custo.get(k, 0) + p["qtd_mensal"] * f["quantidade"] * preco_mat.get(f["material"], 0)
        elif p["categoria"] == "Revenda":
            custo["Mercadorias para revenda"] = custo.get("Mercadorias para revenda", 0) + p["custo_mensal"]
        else:
            custo["Custo informado (sem ficha técnica)"] = custo.get("Custo informado (sem ficha técnica)", 0) + p["custo_mensal"]
    out["custo_direto"] = _serie(custo)

    out["fontes"] = _serie({"Capital próprio (sócios)": res["capital_proprio"],
                            "Capital de terceiros (empréstimos)": res["capital_terceiros"]})
    out["receita_categoria"] = _serie(res["anos"][1]["rec_cat"])
    top = res["cat"].sort_values("receita_mensal", ascending=False).head(10)
    out["top_produtos"] = _serie(dict(zip(top["descricao"], top["receita_mensal"])))
    inv = res["inv"].sort_values("total", ascending=False)
    fixos = dict(zip(inv["descricao"].head(10), inv["total"].head(10)))
    resto = inv["total"].iloc[10:].sum()
    if resto > 0:
        fixos["Demais itens"] = resto
    out["inv_fixos"] = _serie(fixos)

    a1 = res["anos"][1]
    estrutura = {"CMV (revenda)": a1["cmv"], "CPV (fabricação)": a1["cpv"], "CSV (serviços)": a1["csv"],
                 "Impostos sobre vendas": a1["impostos"], "Outras despesas variáveis": a1["outras_var"],
                 "Pessoal": a1["pessoal"], "Demais custos fixos": a1["fixos"], "Depreciação": a1["dep"],
                 "Juros": a1["juros"], "Lucro líquido": max(a1["resultado"], 0)}
    df = pd.DataFrame({"Item": list(estrutura), "Valor": list(estrutura.values())})
    df["%"] = df["Valor"] / a1["rec_bruta"] if a1["rec_bruta"] else 0.0
    out["estrutura_receita"] = df[df["Valor"] > 0.005].reset_index(drop=True)

    # Gráfico do ponto de equilíbrio (receita x custo total, Ano 1)
    imc = a1["imc"] if a1["imc"] == a1["imc"] else 0.0
    pe = a1["pe_contabil"] if a1["pe_contabil"] == a1["pe_contabil"] else a1["rec_bruta"]
    xmax = max(a1["rec_bruta"], pe) * 1.3 or 1.0
    xs = [xmax * i / 20 for i in range(21)]
    out["equilibrio"] = pd.DataFrame({"Receita": xs, "Receita total": xs,
                                      "Custo total": [a1["cf_total"] + x * (1 - imc) for x in xs],
                                      "Custos fixos": [a1["cf_total"]] * len(xs)})
    out["pe_ponto"] = (pe, a1["rec_bruta"])

    fl = res["fluxos"]
    acum, acum_d, s, sd = [], [], 0.0, 0.0
    for t, f in enumerate(fl):
        s += f
        sd += f / (1 + res["tma"]) ** t
        acum.append(s)
        acum_d.append(sd)
    out["fluxo_acumulado"] = pd.DataFrame({"Ano": [f"Ano {t}" for t in range(len(fl))], "Fluxo do ano": fl,
                                           "Acumulado": acum, "Acumulado descontado (TMA)": acum_d})
    return out


# ------------------------------------------------------------------------------
# 11. SIMULADOR "E SE?" E ANÁLISE DE SENSIBILIDADE (Ano 1)
# ------------------------------------------------------------------------------
def simular(res: dict, preco=0.0, volume=0.0, materiais=0.0, salarios=0.0, fixos=0.0) -> dict:
    """Recalcula o resultado do Ano 1 aplicando variações percentuais (0.10 = +10%)."""
    a1 = res["anos"][1]
    fp, fv, fm, fs, ff = 1 + preco, 1 + volume, 1 + materiais, 1 + salarios, 1 + fixos
    rec = a1["rec_bruta"] * fp * fv
    cdir = a1["custo_dir"] * fv * fm
    var = rec * (res["aliq"] + res["var_pct"])
    mc = rec - cdir - var
    cf = a1["pessoal"] * fs + a1["fixos"] * ff + a1["dep"]
    resultado = mc - cf - a1["juros"]
    imc = div(mc, rec)
    return {"receita": rec, "custos_variaveis": cdir + var, "mc": mc, "custos_fixos": cf,
            "resultado": resultado, "lucratividade": div(resultado, rec),
            "pe": div(cf, imc) if imc == imc and imc > 0 else float("nan")}


def sensibilidade(res: dict, delta: float = 0.10) -> pd.DataFrame:
    base = simular(res)["resultado"]
    linhas = []
    for nome, chave_, sinal in [("Preço de venda", "preco", 1), ("Volume de vendas", "volume", 1),
                                ("Custo de materiais e mercadorias", "materiais", -1),
                                ("Salários e encargos", "salarios", -1), ("Demais custos fixos", "fixos", -1)]:
        up = simular(res, **{chave_: delta})["resultado"] - base
        dn = simular(res, **{chave_: -delta})["resultado"] - base
        linhas.append({"Variável": nome, f"+{int(delta*100)}%": up, f"-{int(delta*100)}%": dn,
                       "Impacto absoluto": max(abs(up), abs(dn))})
    return pd.DataFrame(linhas).sort_values("Impacto absoluto", ascending=False).reset_index(drop=True)


# ------------------------------------------------------------------------------
# 12. MERCADO (TAM / SAM / SOM) E CAPACIDADE x DEMANDA
# ------------------------------------------------------------------------------
def mercado(proj: dict, res: dict) -> dict:
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    tam_c, sam_c = float(proj.get("tam_clientes") or 0), float(proj.get("sam_clientes") or 0)
    som_c = sam_c * float(proj.get("som_pct") or 0) / 100
    tam, sam, som = tam_c * gasto, sam_c * gasto, som_c * gasto
    rec1 = res["anos"][1]["rec_bruta"]
    return {"tam": tam, "sam": sam, "som": som, "tam_c": tam_c, "sam_c": sam_c, "som_c": som_c,
            "receita_ano1": rec1, "pct_som": div(rec1, som), "pct_sam": div(rec1, sam),
            "preenchido": tam > 0 and sam > 0 and som > 0}


def capacidade_demanda(d: dict, res: dict) -> pd.DataFrame:
    cat = res["cat"]
    fator12 = (1 + res["g_m"]) ** 11
    linhas = []
    for _, r in d["capacidade"].iterrows():
        alvo = str(r["produto_ou_grupo"])
        if alvo in ("Fabricação Própria", "Serviço", "Revenda"):
            q = cat.loc[cat["categoria"] == alvo, "qtd_mensal"].sum()
        elif alvo == "Todos os produtos":
            q = cat["qtd_mensal"].sum()
        else:
            q = cat.loc[cat["descricao"] == alvo, "qtd_mensal"].sum()
        dem1, dem12 = float(q), float(q) * fator12
        cap = float(r["capacidade_mensal"])
        linhas.append({"Recurso": r["recurso"], "Atende": alvo, "Capacidade/mês": cap,
                       "Demanda Mês 1": dem1, "Demanda Mês 12": dem12,
                       "Uso Mês 1": div(dem1, cap), "Uso Mês 12": div(dem12, cap)})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 13. SEMÁFORO DOS INDICADORES COM "EXPLIQUE ESTE NÚMERO"
# ------------------------------------------------------------------------------
def semaforo(res: dict) -> list:
    a1 = res["anos"][1]
    it, tma = res["investimento_total"], res["tma"]
    out = []

    def add(nome, valor, txt, status, leitura, conta):
        out.append({"Indicador": nome, "Valor": txt, "Status": status, "Leitura": leitura,
                    "Como foi calculado": conta, "valor_num": valor})

    luc = div(a1["resultado"], a1["rec_bruta"])
    add("Lucratividade líquida (Ano 1)", luc, pct(luc),
        "🟢" if luc >= 0.10 else ("🟡" if luc > 0 else "🔴"),
        "Quanto sobra de lucro para cada R$ 1,00 vendido. Acima de 10% é bom para pequenos negócios de alimentação e varejo; "
        "negativo significa prejuízo.",
        f"{brl(a1['resultado'])} ÷ {brl(a1['rec_bruta'])} = {pct(luc)}")
    imc = a1["imc"]
    add("Margem de contribuição", imc, pct(imc),
        "🟢" if imc >= 0.40 else ("🟡" if imc >= 0.25 else "🔴"),
        "Parte de cada venda que sobra, depois dos custos variáveis, para pagar a estrutura fixa. Margens baixas "
        "exigem volume muito alto.",
        f"{brl(a1['mc'])} ÷ {brl(a1['rec_bruta'])} = {pct(imc)}")
    ms = div(a1["rec_bruta"] - a1["pe_contabil"], a1["rec_bruta"])
    add("Margem de segurança", ms, pct(ms),
        "🟢" if ms >= 0.20 else ("🟡" if ms > 0 else "🔴"),
        "Quanto as vendas podem cair antes de o negócio ter prejuízo. Abaixo de 0% significa que a receita prevista "
        "não cobre o ponto de equilíbrio.",
        f"({brl(a1['rec_bruta'])} − PE {brl(a1['pe_contabil'])}) ÷ {brl(a1['rec_bruta'])} = {pct(ms)}")
    add(f"VPL ({res['N']} anos)", res["vpl"], brl(res["vpl"]),
        "🟢" if res["vpl"] > 0 else "🔴",
        "Valor criado pelo projeto acima do retorno mínimo exigido (TMA). Positivo: o investimento compensa.",
        "Soma dos fluxos descontados: " + " + ".join(f"{brl(f)}/(1+{num_br(tma*100,1)}%)^{t}" for t, f in enumerate(res["fluxos"])))
    tir = res["tir"]
    add("TIR", tir if tir is not None else float("nan"), pct(tir) if tir is not None else "n.a.",
        "🟢" if tir is not None and tir > tma else "🔴",
        f"Rentabilidade anual do projeto. Deve superar a TMA de {pct(tma)}.",
        "Taxa que faz o VPL ser zero nos fluxos: " + "; ".join(brl(f) for f in res["fluxos"]))
    pb = res["payback_desc"]
    add("Payback descontado", pb, f"{num_br(pb,1)} meses" if pb == pb else f"> {res['MESES']} meses",
        "🟢" if pb == pb and pb <= 24 else ("🟡" if pb == pb else "🔴"),
        "Tempo para recuperar o investimento considerando o valor do dinheiro no tempo.",
        f"Investimento de {brl(it)} recuperado pelos fluxos descontados à TMA de {pct(tma)}")
    lc = div(a1["ac"], a1["pc"])
    add("Liquidez corrente (Ano 1)", lc, f"{num_br(lc)}x",
        "🟢" if lc >= 1.5 else ("🟡" if lc >= 1 else "🔴"),
        "Quantos reais de bens e direitos de curto prazo existem para cada R$ 1,00 de dívida de curto prazo.",
        f"{brl(a1['ac'])} ÷ {brl(a1['pc'])} = {num_br(lc)}x")
    pes = div(a1["pessoal"], a1["rec_liquida"])
    add("Peso da folha de pagamento", pes, pct(pes),
        "🟢" if pes <= 0.25 else ("🟡" if pes <= 0.35 else "🔴"),
        "Quanto da receita líquida vai para salários, encargos e benefícios. Acima de 35% costuma apertar o resultado.",
        f"{brl(a1['pessoal'])} ÷ {brl(a1['rec_liquida'])} = {pct(pes)}")
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        lc = rc["ltv_cac"]
        add("LTV/CAC (assinaturas)", lc, f"{num_br(lc, 1)}x" if lc == lc else "n.a.",
            "🟢" if lc == lc and lc >= 3 else ("🟡" if lc == lc and lc >= 1 else "🔴"),
            "Quantas vezes o lucro bruto que um assinante deixa ao longo da vida cobre o custo de conquistá-lo. "
            "Abaixo de 1 cada cliente novo dá prejuízo; 3 ou mais é a referência usada por investidores.",
            f"LTV {brl(rc['ltv'])} ÷ CAC {brl(rc['cac'])}" if lc == lc else "Sem gasto comercial ou sem novos assinantes no Mês 1")
        ch_ = rc["churn"]
        add("Churn mensal efetivo", ch_, pct(ch_),
            "🟢" if ch_ <= 0.03 else ("🟡" if ch_ <= 0.07 else "🔴"),
            "Percentual de assinantes que cancelam por mês. Cada ponto de churn encurta a vida do cliente e o LTV.",
            f"Churn de cada plano ponderado pelos assinantes ativos no Mês 12 = {pct(ch_)}; vida média ≈ 1 ÷ churn = "
            f"{num_br(rc['vida_meses'], 0)} meses")
    fm = res["fluxo_mensal"]
    minimo = float(fm["Saldo final"].min()) if not fm.empty else 0.0
    mes_min = fm.loc[fm["Saldo final"].idxmin(), "Mês"] if not fm.empty else "-"
    add("Menor saldo de caixa no Ano 1", minimo, brl(minimo),
        "🟢" if minimo > res["desembolso_m1"] * 0.5 else ("🟡" if minimo >= 0 else "🔴"),
        f"O caixa fica mais apertado em {mes_min}. Negativo indica que faltará dinheiro nesse mês.",
        f"Saldo inicial + recebimentos − pagamentos, mês a mês (menor valor em {mes_min})")
    return out


# ------------------------------------------------------------------------------
# 14. CHECAGENS COM INDICAÇÃO DE ONDE CORRIGIR
# ------------------------------------------------------------------------------
def tabela_checagens_v5(res: dict, d: dict, proj: dict) -> pd.DataFrame:
    base = tabela_checagens(res, d)
    destinos = {
        "Fontes de recursos x investimento total": ("investimentos", "Investimentos → Fontes de recursos",
            "Ajuste o capital dos sócios (aba Projeto) ou os empréstimos para que a soma seja igual ao investimento total."),
        "Margem unitária positiva": ("produtos", "Produtos e materiais → Produtos e serviços",
            "Revise o preço de venda ou o custo dos itens listados."),
        "Produtos com custo cadastrado": ("produtos", "Produtos e materiais → Ficha técnica",
            "Monte a ficha técnica ou informe o custo dos itens listados."),
        "Fichas técnicas com materiais válidos": ("produtos", "Produtos e materiais → Ficha técnica",
            "Cadastre os materiais que faltam ou troque o insumo na ficha técnica."),
        "Prazos de recebimento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de recebimento para somarem 100%."),
        "Prazos de pagamento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de pagamento para somarem 100%."),
        "SWOT: ambiente coerente com a classificação": ("forcas", "Estratégia → Fatores internos e externos",
            "Fatores internos são Forças/Fraquezas; externos são Oportunidades/Ameaças."),
        "Caixa projetado não negativo": ("fluxo", "Finanças → Fluxo de caixa",
            "Aumente a reserva/capital de giro, renegocie prazos ou reduza custos."),
        "Liquidez corrente do Ano 1 >= 1,00": ("investimentos", "Investimentos → Capital de giro",
            "Aumente o capital de giro ou alongue o prazo das dívidas."),
        "VPL positivo": ("simulador", "Simulador e sensibilidade",
            "Veja quais variáveis mais afetam o resultado e teste ajustes de preço, volume e custos."),
    }
    base["destino"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[0])
    base["Onde corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[1])
    base["Como corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[2])
    extras = []
    mk = mercado(proj, res)
    if mk["preenchido"]:
        ok = mk["pct_som"] <= 1
        extras.append({"Teste": "Meta de vendas cabe no mercado (SOM)", "Valor apurado": f"Receita = {pct(mk['pct_som'])} do SOM",
                       "Situação": "OK" if ok else "VERIFICAR",
                       "Observação": "A receita do Ano 1 não deve superar o mercado que a empresa consegue conquistar.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Revise as quantidades vendidas ou a estimativa de mercado."})
    else:
        extras.append({"Teste": "Tamanho do mercado estimado", "Valor apurado": "não preenchido", "Situação": "VERIFICAR",
                       "Observação": "Estime TAM, SAM e SOM para verificar se a meta de vendas é realista.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Preencha clientes, percentual de participação e gasto anual por cliente."})
    cd = capacidade_demanda(d, res)
    if not cd.empty:
        exced = cd[cd["Uso Mês 12"] > 1]
        extras.append({"Teste": "Capacidade produtiva suficiente", "Valor apurado": f"{len(exced)} recurso(s) acima de 100%",
                       "Situação": "OK" if exced.empty else "VERIFICAR",
                       "Observação": ", ".join(exced["Recurso"]) or "A capacidade atende a demanda prevista até o Mês 12.",
                       "destino": "operacional", "Onde corrigir": "Marketing e operação → Plano operacional",
                       "Como corrigir": "Aumente a capacidade (equipamentos, pessoas, turnos) ou reveja a meta de vendas."})
    # --- Serviços: horas demandadas x horas disponíveis
    ch = custo_hora(res, d, proj)
    if ch["horas_demandadas_m1"] > 0:
        ok = ch["ocupacao_m12"] == ch["ocupacao_m12"] and ch["ocupacao_m12"] <= 1
        extras.append({"Teste": "Horas da equipe suficientes para os serviços",
                       "Valor apurado": f"Ocupação no Mês 12: {pct(ch['ocupacao_m12'])}" if ch["horas_cobraveis"] > 0
                       else "nenhuma pessoa na área de produção/operação",
                       "Situação": "OK" if ok else "VERIFICAR",
                       "Observação": "As horas vendidas não podem passar das horas cobráveis da equipe.",
                       "destino": "servicos", "Onde corrigir": "Finanças → Serviços e custo da hora",
                       "Como corrigir": "Contrate mais pessoas, aumente a utilização ou reduza as horas vendidas."})
    # --- Assinaturas
    ass = d.get("assinaturas")
    rc = res.get("recorrencia", {"tem": False})
    if ass is not None and not ass.empty:
        sem_churn = ass[(ass["churn_mensal_pct"] <= 0) & (ass["preco_ciclo"] > 0)]
        extras.append({"Teste": "Assinaturas com taxa de cancelamento (churn)",
                       "Valor apurado": f"{len(sem_churn)} plano(s) com churn zero",
                       "Situação": "OK" if sem_churn.empty else "VERIFICAR",
                       "Observação": ", ".join(sem_churn["plano"].astype(str)) or
                       "Todos os planos preveem cancelamentos, como acontece na prática.",
                       "destino": "assinaturas", "Onde corrigir": "Finanças → Assinaturas e receita recorrente",
                       "Como corrigir": "Informe um churn mensal realista (pesquisas citam algo entre 3% e 8% ao mês para pequenas "
                                        "empresas de software e aplicativos)."})
        if rc.get("tem"):
            lc = rc["ltv_cac"]
            ok = lc == lc and lc >= 3
            extras.append({"Teste": "LTV/CAC de pelo menos 3", "Valor apurado": f"{num_br(lc, 1)}x" if lc == lc else "sem CAC",
                           "Situação": "OK" if ok else "VERIFICAR",
                           "Observação": "Cada cliente deve gerar, ao longo da vida, ao menos 3 vezes o que custou para conquistá-lo.",
                           "destino": "assinaturas", "Onde corrigir": "Finanças → Assinaturas → Indicadores",
                           "Como corrigir": "Reduza o churn, aumente o preço/ARPU ou diminua o gasto comercial por novo cliente."})
    if str(proj.get("tipo_negocio", "")).startswith(("Serviço", "Assinatura")) and res["inv_estoque"] > 0:
        extras.append({"Teste": "Estoque inicial em negócio de serviço", "Valor apurado": brl(res["inv_estoque"]),
                       "Situação": "VERIFICAR",
                       "Observação": "Serviços e assinaturas costumam não ter estoque; confira se os materiais são mesmo necessários.",
                       "destino": "produtos", "Onde corrigir": "Produtos e materiais → Materiais",
                       "Como corrigir": "Zere o estoque inicial dos itens que não serão comprados antes da abertura."})
    return pd.concat([base, pd.DataFrame(extras)], ignore_index=True)


# ------------------------------------------------------------------------------
# 15. RAIO-X DO PLANO (nota de 0 a 100 e principais correções)
# ------------------------------------------------------------------------------
def raio_x(res: dict, chk: pd.DataFrame, completude: float, sem: list) -> dict:
    pontos_chk = 30 * (chk["Situação"] == "OK").mean() if len(chk) else 0
    pontos_comp = 30 * completude
    verdes = sum(1 for s in sem if s["Status"] == "🟢")
    amarelos = sum(1 for s in sem if s["Status"] == "🟡")
    pontos_fin = 40 * (verdes + 0.5 * amarelos) / max(len(sem), 1)
    nota = round(pontos_chk + pontos_comp + pontos_fin)
    acoes = [f"{r['Teste']}: {r['Como corrigir'] or r['Observação']}" for _, r in chk[chk["Situação"] != "OK"].iterrows()]
    acoes += [f"{s['Indicador']} em {s['Valor']}: {s['Leitura']}" for s in sem if s["Status"] == "🔴"]
    return {"nota": nota, "acoes": acoes[:5], "pontos": (round(pontos_comp), round(pontos_chk), round(pontos_fin))}


# ------------------------------------------------------------------------------
# 15. SERVIÇOS: CUSTO DA HORA PRODUTIVA, CAPACIDADE EM HORAS E PREÇO DO SERVIÇO
# ------------------------------------------------------------------------------
def custo_hora(res: dict, d: dict, proj: dict) -> dict:
    """
    Custo da hora produtiva = (mão de obra da operação + custos indiretos + depreciação) ÷ horas cobráveis.
    Horas cobráveis = pessoas da operação x horas contratadas por mês x taxa de utilização.
    Custo pleno do serviço = insumos + horas x custo da hora.
    Preço sugerido = custo pleno ÷ [1 − (impostos + despesas variáveis + despesas de estrutura % + margem)].
    """
    pes = d["pessoal"]
    area = pes["area"].fillna("").astype(str)
    pessoas = float(pes.loc[area.str.startswith("Produ"), "num_empregados"].sum())
    h_mes = float(proj.get("horas_mes_pessoa") or 168)
    util = float(proj.get("utilizacao_pct") or 75) / 100
    horas_contr = pessoas * h_mes
    horas_fat = horas_contr * util
    custo_base = res["mod_m"] + res["cif_m"] + res["dep_m"]
    ch = div(custo_base, horas_fat) if horas_fat > 0 else float("nan")
    rec_a1 = res["anos"][1]["rec_bruta"]
    estrutura = res["pessoal_adm_m"] + res["pessoal_com_m"] + res["desp_adm_m"] + res["desp_com_m"] + res["desp_fin_m"]
    desp_pct = div(estrutura * 12, rec_a1) if rec_a1 > 0 else 0.0       # média do Ano 1
    margem = res["margem_desejada"]
    divisor = 1 - (res["aliq"] + res["var_pct"] + desp_pct + margem)
    cat = res["cat"]
    sv = cat[cat["horas_unit"] > 0].copy() if "horas_unit" in cat else cat.iloc[0:0].copy()
    linhas = []
    for _, r in sv.iterrows():
        h = float(r["horas_unit"])
        mao = h * ch if ch == ch else float("nan")
        pleno = float(r["custo_unit"]) + (mao if mao == mao else 0.0)
        preco = float(r["preco_venda"])
        sug = div(pleno, divisor) if divisor > 0.05 else float("nan")
        margem_real = div(preco - pleno - preco * (res["aliq"] + res["var_pct"]), preco) if preco > 0 else float("nan")
        linhas.append({"Serviço": r["descricao"], "Horas por serviço": h, "Qtd/mês": float(r["qtd_mensal"]),
                       "Horas/mês": h * float(r["qtd_mensal"]), "Insumos": float(r["custo_unit"]),
                       "Mão de obra (horas x custo hora)": mao, "Custo pleno": pleno, "Preço atual": preco,
                       "Preço sugerido": sug, "Margem após custos diretos": margem_real,
                       "Valor da hora cobrada": div(preco, h)})
    tab = pd.DataFrame(linhas)
    horas_dem = float(tab["Horas/mês"].sum()) if not tab.empty else 0.0
    return {"pessoas": pessoas, "horas_mes_pessoa": h_mes, "utilizacao": util, "horas_contratadas": horas_contr,
            "horas_cobraveis": horas_fat, "custo_base": custo_base, "custo_hora": ch, "desp_pct": desp_pct,
            "divisor": divisor, "tabela": tab, "horas_demandadas_m1": horas_dem,
            "horas_demandadas_m12": horas_dem * (1 + res["g_m"]) ** 11,
            "ocupacao_m1": div(horas_dem, horas_fat) if horas_fat > 0 else float("nan"),
            "ocupacao_m12": div(horas_dem * (1 + res["g_m"]) ** 11, horas_fat) if horas_fat > 0 else float("nan")}
