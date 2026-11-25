"""
================================================================================
VIABILIZALAB - VALIDAÇÃO, MARKETING, INCERTEZA E DEFESA (pnbox_validacao.py)
================================================================================
Cálculos das ferramentas da versão 7.1:
    - placar de evidências das premissas (pesos pela sensibilidade do resultado);
    - registro de entrevistas, teste de preço de Van Westendorp e intenção de compra ajustada;
    - Canvas da proposta de valor, curva de valor (ERRC), posicionamento, orçamento por canal e AARRR;
    - simulação de Monte Carlo, curva de adoção de Bass e perda aceitável;
    - régua de prontidão, banca simulada e checklist de formalização.
"""
import math
import re
from collections import Counter

import numpy as np
import pandas as pd

from pnbox_motor import brl, num_br, pct, div, sensibilidade, mercado, simular

# ------------------------------------------------------------------------------
# 1. PLACAR DE EVIDÊNCIAS
# ------------------------------------------------------------------------------
STATUS_EVIDENCIA = {
    "Estimada (sem fonte)": 0.0,
    "Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)": 0.5,
    "Testada com clientes (entrevista, pré-venda, piloto)": 1.0,
}
STATUS_CURTO = {"Estimada (sem fonte)": "estimada",
                "Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)": "pesquisada",
                "Testada com clientes (entrevista, pré-venda, piloto)": "testada"}


def premissas_plano(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    """Premissas que movem o resultado, com o valor usado no plano e o peso de cada uma (soma 100%)."""
    a1 = res["anos"][1]
    sens = sensibilidade(res) if a1["rec_bruta"] > 0 else pd.DataFrame(columns=["Variável", "Impacto absoluto"])
    imp = dict(zip(sens["Variável"], sens["Impacto absoluto"]))
    base = float(np.mean(list(imp.values()))) if imp else 1.0
    base = base if base > 0 else 1.0
    cat = res.get("cat", pd.DataFrame())
    qtd = float(cat["qtd_mensal"].sum()) if "qtd_mensal" in cat else 0.0
    ticket = div(float((cat["preco_venda"] * cat["qtd_mensal"]).sum()), qtd) if qtd else float("nan")
    linhas = [
        ("Preço de venda", f"preço médio {brl(ticket)}" if ticket == ticket else "assinaturas", imp.get("Preço de venda", base)),
        ("Volume de vendas do Ano 1", f"receita {brl(a1['rec_bruta'])}", imp.get("Volume de vendas", base)),
        ("Custo de materiais e mercadorias", brl(a1["custo_dir"]), imp.get("Custo de materiais e mercadorias", base)),
        ("Salários e encargos", brl(a1["pessoal"]), imp.get("Salários e encargos", base)),
        ("Demais custos fixos", brl(a1["fixos"]), imp.get("Demais custos fixos", base)),
        ("Crescimento das vendas", f"{num_br(proj.get('crescimento_mensal_ano1_pct') or 0, 1)}% a.m. no Ano 1; "
                                   f"{num_br(proj.get('crescimento_anual_pct') or 0, 1)}% a.a. depois",
         0.5 * imp.get("Volume de vendas", base)),
    ]
    m = mercado(proj, res)
    if m["preenchido"]:
        linhas.append(("Participação de mercado (SOM)", f"{num_br(proj.get('som_pct') or 0, 1)}% do SAM", 0.4 * base))
    if not d["funil_vendas"].empty:
        linhas.append(("Conversão do funil de vendas", "taxas do funil", 0.4 * base))
    rec = res.get("recorrencia", {})
    if rec.get("tem"):
        linhas.append(("Cancelamento mensal (churn)", pct(rec.get("churn", 0)), 0.8 * base))
        linhas.append(("Custo de aquisição (CAC)", brl(rec.get("cac", 0)), 0.4 * base))
    linhas.append(("Prazo de recebimento das vendas", "prazos do capital de giro", 0.2 * base))
    df = pd.DataFrame(linhas, columns=["premissa", "valor_plano", "peso"])
    df["peso"] = df["peso"].clip(lower=0) / max(df["peso"].clip(lower=0).sum(), 1e-9)
    return df


def placar_evidencias(proj: dict, d: dict, res: dict) -> dict:
    pr = premissas_plano(proj, d, res)
    ev = d.get("evidencias", pd.DataFrame(columns=["premissa", "status", "fonte", "observacao"]))
    mapa = {str(r["premissa"]): r for _, r in ev.iterrows()}
    st_, fo = [], []
    for p in pr["premissa"]:
        r = mapa.get(p)
        s = str(r["status"]) if r is not None and str(r["status"]) in STATUS_EVIDENCIA else "Estimada (sem fonte)"
        st_.append(s)
        fo.append(str(r["fonte"]) if r is not None else "")
    pr["status"], pr["fonte"] = st_, fo
    pr["nota"] = pr["status"].map(STATUS_EVIDENCIA)
    score = float((pr["peso"] * pr["nota"]).sum())
    pend = pr[pr["nota"] < 1].sort_values("peso", ascending=False)
    prox = pend.iloc[0]["premissa"] if not pend.empty else ""
    return {"tabela": pr, "score": score, "proxima": prox,
            "testadas": int((pr["nota"] == 1).sum()), "n": len(pr)}


# ------------------------------------------------------------------------------
# 2. ENTREVISTAS, VAN WESTENDORP E INTENÇÃO DE COMPRA
# ------------------------------------------------------------------------------
PAGARIA = ["Sim", "Talvez", "Não"]

ROTEIRO_ENTREVISTA = """**Antes:** converse com pessoas do segmento-alvo, uma de cada vez, sem apresentar a sua ideia no começo.
Pergunte sobre o **passado e o comportamento real**, não sobre opiniões e intenções futuras (Fitzpatrick, 2013).

1. Conte a última vez que você precisou resolver [o problema]. O que aconteceu?
2. O que foi mais difícil ou chato nisso? Por quê?
3. O que você já tentou para resolver? Quanto gastou com isso (dinheiro e tempo)?
4. O que você usa hoje? O que gosta e o que não gosta nessa solução?
5. Com que frequência isso acontece?
6. *(Só agora apresente a ideia em uma frase.)* O que achou? O que faltaria para você usar?
7. Quanto pagaria por isso? Pagaria hoje, com um pré-cadastro ou sinal?
8. Quem mais você conhece que tem esse problema? Pode me apresentar?

**Sinais fortes:** a pessoa já gasta dinheiro ou tempo com o problema, pede para ser avisada, aceita pagar um sinal ou indica
outras pessoas. **Sinais fracos:** elogios ("que ideia legal!") e promessas sobre o futuro ("eu compraria, com certeza")."""


def _itens(txt) -> list:
    return [re.sub(r"\s+", " ", x).strip().lower() for x in re.split(r"[;\n,]", str(txt or "")) if x.strip()]


def resumo_entrevistas(d: dict, preco_ref: float = float("nan")) -> dict:
    e = d.get("entrevistas", pd.DataFrame())
    e = e[e["entrevistado"].astype(str).str.strip() != ""] if not e.empty else e
    n = len(e)
    if n == 0:
        return {"n": 0}
    dores = Counter(x for t in e["dores"] for x in set(_itens(t)))
    tab_dores = pd.DataFrame([{"Dor citada": k, "Entrevistas": v, "% das entrevistas": v / n}
                              for k, v in dores.most_common(12)])
    pag = e["pagaria"].astype(str)
    vals = pd.to_numeric(e["valor_pagaria"], errors="coerce")
    vals = vals[vals > 0]
    med = float(vals.median()) if len(vals) else float("nan")
    return {"n": n, "dores": tab_dores, "dor_top": dores.most_common(1)[0] if dores else ("", 0),
            "pct_sim": float((pag == "Sim").mean()), "pct_talvez": float((pag == "Talvez").mean()),
            "pct_nao": float((pag == "Não").mean()), "mediana_valor": med,
            "valor_vs_preco": div(med, preco_ref) if preco_ref == preco_ref and preco_ref > 0 else float("nan")}


def van_westendorp(df: pd.DataFrame) -> dict | None:
    """Price Sensitivity Meter (Van Westendorp, 1976). Retorna as curvas e os quatro pontos de preço."""
    if df is None or df.empty:
        return None
    x = df[["muito_barato", "barato", "caro", "muito_caro"]].apply(pd.to_numeric, errors="coerce")
    x = x[(x > 0).all(axis=1)]
    # descarta respostas incoerentes (a ordem precisa ser muito barato <= barato <= caro <= muito caro)
    x = x[(x["muito_barato"] <= x["barato"]) & (x["barato"] <= x["caro"]) & (x["caro"] <= x["muito_caro"])]
    n = len(x)
    if n < 5:
        return {"n": n, "ok": False}
    lo, hi = float(x.values.min()) * 0.8, float(x.values.max()) * 1.1
    grade = np.linspace(lo, hi, 400)
    mb = np.array([(x["muito_barato"] >= p).mean() for p in grade])
    ba = np.array([(x["barato"] >= p).mean() for p in grade])
    ca = np.array([(x["caro"] <= p).mean() for p in grade])
    mc = np.array([(x["muito_caro"] <= p).mean() for p in grade])
    nao_barato, nao_caro = 1 - ba, 1 - ca

    def cruz(a, b):
        dif = a - b
        idx = np.where(np.diff(np.sign(dif)) != 0)[0]
        if len(idx) == 0:
            return float("nan")
        i = idx[0]
        t = dif[i] / (dif[i] - dif[i + 1]) if dif[i] != dif[i + 1] else 0.0
        return float(grade[i] + t * (grade[i + 1] - grade[i]))

    curvas = pd.DataFrame({"Preço": grade, "Muito barato": mb, "Barato": ba, "Caro": ca, "Muito caro": mc,
                           "Não é barato": nao_barato, "Não é caro": nao_caro})
    return {"n": n, "ok": True, "curvas": curvas,
            "pmc": cruz(mb, nao_barato),      # preço marginal barato (limite inferior)
            "pme": cruz(mc, nao_caro),        # preço marginal caro (limite superior)
            "opp": cruz(mb, mc),              # preço ótimo
            "ipp": cruz(nao_barato, nao_caro)}  # preço de indiferença


def intencao_ajustada(proj: dict, res: dict) -> dict:
    """Converte intenção de compra declarada em clientes prováveis (fatores editáveis)."""
    base = float(proj.get("sam_clientes") or 0)
    c, p = float(proj.get("intencao_certamente_pct") or 0) / 100, float(proj.get("intencao_provavel_pct") or 0) / 100
    fc, fp = float(proj.get("fator_certamente_pct") or 70) / 100, float(proj.get("fator_provavel_pct") or 30) / 100
    declarados = base * (c + p)
    ajustados = base * (c * fc + p * fp)
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    plano = div(res["anos"][1]["rec_bruta"], gasto) if gasto > 0 else float("nan")
    return {"base": base, "declarados": declarados, "ajustados": ajustados, "plano": plano,
            "cobertura": div(ajustados, plano) if plano == plano and plano > 0 else float("nan"),
            "ok": base > 0 and (c + p) > 0}


# ------------------------------------------------------------------------------
# 3. MARKETING: PROPOSTA DE VALOR, CURVA DE VALOR, POSICIONAMENTO, CANAIS
# ------------------------------------------------------------------------------
BLOCOS_VPC = [("vpc_tarefas", "Tarefas do cliente (jobs)", "O que o cliente tenta resolver ou realizar"),
              ("vpc_dores", "Dores", "O que incomoda, atrapalha ou gera risco antes, durante e depois da tarefa"),
              ("vpc_ganhos", "Ganhos", "Resultados e benefícios que o cliente quer ou que o surpreenderiam"),
              ("vpc_produtos", "Produtos e serviços", "O que a empresa oferece"),
              ("vpc_aliviadores", "Aliviadores de dores", "Como a oferta elimina ou reduz cada dor"),
              ("vpc_criadores", "Criadores de ganhos", "Como a oferta produz cada ganho")]


def linhas_txt(t) -> list:
    return [x.strip(" -•\t") for x in str(t or "").splitlines() if x.strip(" -•\t")]


def encaixe_vpc(proj: dict, d: dict) -> dict:
    dores, ganhos = linhas_txt(proj.get("vpc_dores")), linhas_txt(proj.get("vpc_ganhos"))
    enc = d.get("vpc_encaixe", pd.DataFrame(columns=["item", "tipo", "resposta"]))
    resp = {(str(r["tipo"]), str(r["item"])): str(r["resposta"]) for _, r in enc.iterrows()}
    linhas = [{"tipo": "Dor", "item": x, "resposta": resp.get(("Dor", x), "")} for x in dores] + \
             [{"tipo": "Ganho", "item": x, "resposta": resp.get(("Ganho", x), "")} for x in ganhos]
    tab = pd.DataFrame(linhas, columns=["tipo", "item", "resposta"])
    cobertos = int((tab["resposta"].astype(str).str.strip() != "").sum()) if not tab.empty else 0
    return {"tabela": tab, "cobertos": cobertos, "total": len(tab), "encaixe": div(cobertos, len(tab)) if len(tab) else float("nan")}


ACOES_ERRC = ["Manter", "Eliminar", "Reduzir", "Elevar", "Criar"]


def curva_valor(d: dict) -> pd.DataFrame:
    """Notas da concorrência (formato longo) + nova curva proposta pelo ERRC."""
    notas = d["concorrencia_notas"].copy()
    notas = notas[notas["atributo"].astype(str).str.strip() != ""]
    errc = d.get("errc", pd.DataFrame(columns=["atributo", "acao", "nova_nota", "justificativa"]))
    novos = errc[errc["atributo"].astype(str).str.strip() != ""]
    linhas = [{"Atributo": r["atributo"], "Empresa": r["empresa"], "Nota": float(r["nota"])} for _, r in notas.iterrows()]
    nossa = {r["atributo"]: float(r["nota"]) for _, r in notas[notas["empresa"] == "Nossa Empresa"].iterrows()}
    for _, r in novos.iterrows():
        a = r["atributo"]
        if str(r["acao"]) == "Eliminar":
            v = 0.0
        elif float(r["nova_nota"] or 0) > 0 or str(r["acao"]) in ("Criar", "Reduzir"):
            v = float(r["nova_nota"] or 0)
        else:
            v = nossa.get(a, 0.0)
        linhas.append({"Atributo": a, "Empresa": "Nova curva (ERRC)", "Nota": v})
    for a, v in nossa.items():
        if a not in set(novos["atributo"]):
            linhas.append({"Atributo": a, "Empresa": "Nova curva (ERRC)", "Nota": v})
    df = pd.DataFrame(linhas, columns=["Atributo", "Empresa", "Nota"])
    ordem = list(dict.fromkeys(list(notas["atributo"]) + list(novos["atributo"])))
    df["ordem"] = df["Atributo"].map({a: i for i, a in enumerate(ordem)})
    return df.sort_values(["ordem", "Empresa"]).drop(columns="ordem").reset_index(drop=True)


def leitura_errc(d: dict) -> str:
    errc = d.get("errc", pd.DataFrame())
    if errc.empty:
        return ""
    cont = Counter(str(a) for a in errc["acao"])
    if cont.get("Criar", 0) == 0 and cont.get("Eliminar", 0) + cont.get("Reduzir", 0) == 0:
        return ("A nova curva só eleva ou mantém atributos: isso tende a aumentar custos sem diferenciar. "
                "Procure o que eliminar ou reduzir (o que o cliente não valoriza) e algo novo para criar.")
    if cont.get("Criar", 0) == 0:
        return "Há cortes de atributos, mas nada novo foi criado: que fator, que o setor não oferece, poderia atrair não clientes?"
    return (f"A nova curva elimina ou reduz {cont.get('Eliminar', 0) + cont.get('Reduzir', 0)} atributo(s), eleva "
            f"{cont.get('Elevar', 0)} e cria {cont.get('Criar', 0)}: é o movimento de diferenciação com custo controlado "
            "proposto pela estratégia do oceano azul.")


def frase_posicionamento(proj: dict) -> str:
    c = {k: str(proj.get(k) or "").strip() for k in ("posic_segmento", "posic_necessidade", "posic_marca", "posic_categoria",
                                                       "posic_beneficio", "posic_concorrente", "posic_diferencial")}
    if not (c["posic_segmento"] and c["posic_beneficio"]):
        return ""
    return (f"Para {c['posic_segmento']} que {c['posic_necessidade'] or '...'}, {c['posic_marca'] or proj.get('nome_negocio', '')} "
            f"é {c['posic_categoria'] or '...'} que {c['posic_beneficio']}. Diferentemente de {c['posic_concorrente'] or '...'}, "
            f"{c['posic_diferencial'] or '...'}.")


def orcamento_canais(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    oc = d.get("orcamento_canais", pd.DataFrame())
    oc = oc[oc["canal"].astype(str).str.strip() != ""] if not oc.empty else oc
    if oc.empty:
        return pd.DataFrame()
    rec = res.get("recorrencia", {})
    valor = float(rec.get("ltv") or 0) if rec.get("tem") and rec.get("ltv", 0) == rec.get("ltv", 0) else 0.0
    rotulo = "LTV"
    if valor <= 0:
        valor, rotulo = float(proj.get("gasto_anual_cliente") or 0), "gasto anual"
    out = pd.DataFrame({"Canal": oc["canal"], "Gasto mensal": oc["gasto_mensal"], "Clientes novos/mês": oc["clientes_mes"]})
    out["CAC"] = [div(g, c) if c > 0 else float("nan") for g, c in zip(out["Gasto mensal"], out["Clientes novos/mês"])]
    out["Valor do cliente"] = valor
    out["Retorno por R$ 1"] = [div(c * valor, g) if g > 0 else float("nan") for g, c in zip(out["Gasto mensal"], out["Clientes novos/mês"])]
    out.attrs["rotulo_valor"] = rotulo
    return out


ETAPAS_AARRR = [("Aquisição", "Como as pessoas descobrem o negócio?"),
                ("Ativação", "Qual é a primeira boa experiência do cliente?"),
                ("Retenção", "O que faz o cliente voltar ou continuar assinando?"),
                ("Receita", "Como e quando o cliente paga?"),
                ("Indicação", "Por que o cliente indicaria o negócio a outras pessoas?")]


# ------------------------------------------------------------------------------
# 4. INCERTEZA: MONTE CARLO, BASS E PERDA ACEITÁVEL
# ------------------------------------------------------------------------------
VARIAVEIS_MC = [("Preço de venda", -10.0, 5.0), ("Volume de vendas", -30.0, 15.0),
                ("Custo de materiais e mercadorias", -5.0, 15.0), ("Salários e encargos", 0.0, 10.0),
                ("Demais custos fixos", -5.0, 15.0), ("Crescimento anual (pontos percentuais)", -5.0, 3.0)]


def faixas_mc(d: dict) -> pd.DataFrame:
    f = d.get("mc_faixas", pd.DataFrame(columns=["variavel", "minimo_pct", "maximo_pct"]))
    mapa = {str(r["variavel"]): (float(r["minimo_pct"]), float(r["maximo_pct"])) for _, r in f.iterrows()}
    return pd.DataFrame([{"variavel": v, "minimo_pct": mapa.get(v, (a, b))[0], "maximo_pct": mapa.get(v, (a, b))[1]}
                         for v, a, b in VARIAVEIS_MC])


def monte_carlo(res: dict, faixas: pd.DataFrame, n: int = 3000, seed: int = 42) -> dict:
    """
    Sorteia variações (distribuição triangular: mínimo, 0 = provável, máximo) e recalcula o VPL do projeto
    e o menor saldo de caixa do Ano 1 com a mesma lógica do motor (fluxo do projeto = EBITDA).
    """
    rng = np.random.default_rng(seed)
    fx = {r["variavel"]: (float(r["minimo_pct"]) / 100, float(r["maximo_pct"]) / 100) for _, r in faixas.iterrows()}

    def sorteio(nome):
        a, b = fx.get(nome, (0.0, 0.0))
        a, b = min(a, 0.0), max(b, 0.0)
        if a == b:
            return np.zeros(n)
        return rng.triangular(a, 0.0, b, n)

    fp, fv = 1 + sorteio("Preço de venda"), 1 + sorteio("Volume de vendas")
    fm, fs = 1 + sorteio("Custo de materiais e mercadorias"), 1 + sorteio("Salários e encargos")
    ff, dg = 1 + sorteio("Demais custos fixos"), sorteio("Crescimento anual (pontos percentuais)")
    g_a, tma, N = res["g_a"], res["tma"], res["N"]
    adj = (1 + g_a + dg) / (1 + g_a)
    vpl = -res["investimento_total"] * np.ones(n)
    fluxo_anos = []
    for k in res["ANOS"]:
        a = res["anos"][k]
        gk = adj ** (k - 1)
        rec = a["rec_bruta"] * fp * fv * gk
        var_ratio = div(a["impostos"] + a["outras_var"], a["rec_bruta"]) if a["rec_bruta"] else 0.0
        ebitda = rec - a["custo_dir"] * fv * fm * gk - rec * var_ratio - a["pessoal"] * fs - a["fixos"] * ff
        if k == N and res.get("recupera_giro"):
            ebitda = ebitda + res.get("capital_giro", 0.0)
        fluxo_anos.append(ebitda)
        vpl = vpl + ebitda / (1 + tma) ** k
    # caixa mês a mês do Ano 1
    fm_df = res["fluxo_mensal"]
    a1 = res["anos"][1]
    w_s = div(a1["pessoal"], a1["pessoal"] + a1["fixos"]) if (a1["pessoal"] + a1["fixos"]) else 0.5
    saldo = float(fm_df["Saldo inicial"].iloc[0]) * np.ones(n)
    minimo = np.full(n, np.inf)
    for _, r in fm_df.iterrows():
        saldo = saldo + r["Recebimentos"] * fp * fv - r["Fornecedores"] * fv * fm - r["Impostos e desp. variáveis"] * fp * fv \
            - r["Pessoal e custos fixos"] * (w_s * fs + (1 - w_s) * ff) - r["Parcelas de empréstimos"]
        minimo = np.minimum(minimo, saldo)
    return {"vpl": vpl, "caixa_min": minimo, "n": n,
            "p_vpl_pos": float((vpl > 0).mean()), "p_caixa_neg": float((minimo < 0).mean()),
            "p5": float(np.percentile(vpl, 5)), "p50": float(np.percentile(vpl, 50)), "p95": float(np.percentile(vpl, 95)),
            "caixa_p5": float(np.percentile(minimo, 5)), "vpl_medio": float(vpl.mean())}


def leitura_mc(mc: dict) -> str:
    p = mc["p_vpl_pos"]
    if p >= 0.8:
        txt = "O projeto é robusto: na grande maioria dos cenários sorteados ele cria valor."
    elif p >= 0.6:
        txt = "O projeto tende a criar valor, mas há uma chance relevante de fracasso: proteja-se das variáveis mais sensíveis."
    elif p >= 0.4:
        txt = "Resultado incerto, quase uma aposta: valide as premissas críticas antes de investir."
    else:
        txt = "Na maior parte dos cenários o projeto destrói valor: o modelo de negócio precisa ser revisto."
    if mc["p_caixa_neg"] > 0.2:
        txt += f" Em {pct(mc['p_caixa_neg'], 0)} dos cenários o caixa fica negativo no Ano 1: aumente a reserva ou o capital de giro."
    return txt


def bass_curva(m: float, p: float, q: float, anos: int) -> pd.DataFrame:
    """Modelo de difusão de Bass (1969): adotantes acumulados ao fim de cada ano."""
    linhas = []
    for t in range(1, anos + 1):
        F = (1 - math.exp(-(p + q) * t)) / (1 + (q / p) * math.exp(-(p + q) * t)) if p > 0 else 0.0
        linhas.append({"Ano": t, "Adotantes acumulados": m * F, "% do mercado potencial": F})
    df = pd.DataFrame(linhas)
    df["Novos adotantes no ano"] = df["Adotantes acumulados"].diff().fillna(df["Adotantes acumulados"])
    return df


def bass_comparacao(proj: dict, res: dict) -> dict:
    m = float(proj.get("bass_mercado") or 0) or float(proj.get("sam_clientes") or 0)
    p, q = float(proj.get("bass_p") or 0.03), float(proj.get("bass_q") or 0.38)
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    if m <= 0:
        return {"ok": False}
    cv = bass_curva(m, p, q, res["N"])
    if gasto > 0:
        cv["Clientes no plano"] = [div(res["anos"][k]["rec_bruta"], gasto) for k in res["ANOS"]]
        cv["Plano ÷ Bass"] = [div(a, b) for a, b in zip(cv["Clientes no plano"], cv["Adotantes acumulados"])]
    return {"ok": True, "m": m, "p": p, "q": q, "tabela": cv, "tem_plano": gasto > 0}


def perda_aceitavel(proj: dict, d: dict, res: dict) -> dict:
    soc = d["socios"]
    aceit = float(soc["perda_aceitavel"].sum()) if "perda_aceitavel" in soc else 0.0
    proprio = float(soc["capital_integralizado"].sum())
    divida = float(d["emprestimos"]["valor"].sum())
    revenda = float(proj.get("revenda_ativos_pct") if proj.get("revenda_ativos_pct") is not None else 40) / 100
    inv_fixo = float((d["inv_fixos"]["quantidade"] * d["inv_fixos"]["valor_unitario"]).sum())
    potencial = max(proprio + divida - inv_fixo * revenda, 0.0)
    return {"aceitavel": aceit, "proprio": proprio, "divida": divida, "revenda": inv_fixo * revenda,
            "potencial": potencial, "ok": aceit > 0, "cabe": aceit >= potencial}


# ------------------------------------------------------------------------------
# 5. DEFESA: FORMALIZAÇÃO, PRONTIDÃO E BANCA SIMULADA
# ------------------------------------------------------------------------------
LIMITE_MEI, LIMITE_ME, LIMITE_EPP = 81_000.0, 360_000.0, 4_800_000.0
STATUS_FORMAL = ["Pendente", "Em andamento", "Concluído", "Não se aplica"]


def itens_formalizacao(proj: dict, d: dict) -> list:
    tipo = str(proj.get("tipo_negocio") or "")
    setor = (str(proj.get("setor_atuacao") or "") + " " + str(proj.get("nome_negocio") or "")).lower()
    itens = ["Escolher a natureza jurídica (MEI, empresário individual, SLU ou LTDA)",
             "Definir as atividades (CNAE) e conferir se são permitidas no Simples/MEI",
             "Consultar a viabilidade do endereço na prefeitura (zoneamento)",
             "Registrar a empresa (Junta Comercial / Redesim) e obter o CNPJ",
             "Inscrição municipal e alvará de funcionamento",
             "Optar pelo Simples Nacional (prazo no início da atividade)",
             "Contratar contador ou escritório contábil",
             "Abrir conta bancária PJ e meios de pagamento",
             "Contrato social com acordo de sócios (entrada, saída, divisão de lucros)",
             "Registrar a marca no INPI"]
    if any(w in setor for w in ("aliment", "café", "cafe", "restaurante", "lanche", "refei", "fit", "padaria", "bebida")):
        itens.append("Licença sanitária (Vigilância Sanitária) e boas práticas de manipulação")
    if float(proj.get("loc_area_m2") or 0) > 0 or proj.get("loc_endereco"):
        itens.append("Auto de vistoria / licença do Corpo de Bombeiros")
    if not d["pessoal"].empty:
        itens.append("Registro de empregados (eSocial), exames admissionais e segurança do trabalho")
    if tipo.startswith(("Assinatura", "Serviço", "Misto")) or "app" in setor or "aplicativo" in setor:
        itens += ["LGPD: política de privacidade, bases legais e encarregado de dados",
                  "Termos de uso e contrato de assinatura (renovação, cancelamento, reembolso)"]
    return itens


def porte_sugerido(res: dict, d: dict) -> dict:
    rec = res["anos"][1]["rec_bruta"]
    n_soc = len(d["socios"])
    n_emp = float(d["pessoal"]["num_empregados"].sum()) if not d["pessoal"].empty else 0
    if rec <= LIMITE_MEI and n_soc <= 1 and n_emp <= 1:
        porte = "MEI (se a atividade for permitida)"
    elif rec <= LIMITE_ME:
        porte = "Microempresa (ME)"
    elif rec <= LIMITE_EPP:
        porte = "Empresa de Pequeno Porte (EPP)"
    else:
        porte = "Acima do limite do Simples Nacional (Lucro Presumido ou Real)"
    obs = []
    if rec <= LIMITE_MEI and n_soc > 1:
        obs.append("O MEI não admite sócios: com mais de um sócio, o caminho é ME (LTDA).")
    return {"porte": porte, "receita": rec, "obs": obs}


NIVEIS_PRONTIDAO = [
    (1, "Ideia descrita", "Missão, proposta de valor e segmento definidos."),
    (2, "Mercado dimensionado", "TAM, SAM e SOM estimados e meta de vendas coerente com o SOM."),
    (3, "Hipóteses mapeadas", "Canvas, premissas listadas e status de evidência informado."),
    (4, "Problema validado", "Pelo menos 10 entrevistas e uma dor citada por metade ou mais dos entrevistados."),
    (5, "Solução e preço validados", "Preço dentro da faixa aceitável (Van Westendorp) ou 40% ou mais dispostos a pagar."),
    (6, "Modelo econômico viável", "VPL positivo e caixa do Ano 1 sem ficar negativo."),
    (7, "Viabilidade robusta", "Chance de VPL positivo de 70% ou mais (Monte Carlo) e placar de evidências de 50% ou mais."),
    (8, "Plano de execução", "Objetivos SMART, cronograma, riscos e 60% ou mais do checklist de formalização encaminhado."),
    (9, "Pronto para lançar ou buscar investimento", "Placar de 75% ou mais, perda aceitável coberta e perguntas da banca respondidas."),
]


def avaliar_prontidao(proj: dict, d: dict, res: dict, mc: dict | None, preco_ref: float) -> dict:
    a1 = res["anos"][1]
    pl = placar_evidencias(proj, d, res)
    ent = resumo_entrevistas(d, preco_ref)
    vw = van_westendorp(d.get("precos_vw"))
    m = mercado(proj, res)
    fo = d.get("formalizacao", pd.DataFrame())
    fo_ok = div(float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).sum()), len(fo)) if len(fo) else 0.0
    perda = perda_aceitavel(proj, d, res)
    resp = d.get("banca_respostas", pd.DataFrame())
    n_resp = int((resp["resposta"].astype(str).str.len() > 20).sum()) if not resp.empty else 0
    preco_ok = False
    if vw and vw.get("ok") and preco_ref == preco_ref:
        preco_ok = vw["pmc"] <= preco_ref <= vw["pme"]
    cond = {
        1: bool(proj.get("missao")) and not d["propostas"].empty and not d["segmentos"].empty,
        2: m["preenchido"] and (m["pct_som"] != m["pct_som"] or m["pct_som"] <= 1.0),
        3: (d.get("evidencias") is not None and len(d["evidencias"]) >= 3),
        4: ent["n"] >= 10 and ent.get("dor_top", ("", 0))[1] >= ent["n"] / 2,
        5: preco_ok or (ent["n"] >= 10 and ent.get("pct_sim", 0) >= 0.4),
        6: res["vpl"] > 0 and float(res["fluxo_mensal"]["Saldo final"].min()) >= 0,
        7: mc is not None and mc["p_vpl_pos"] >= 0.7 and pl["score"] >= 0.5,
        8: len(d["objetivos"]) >= 3 and len(d["cronograma"]) >= 3 and len(d["riscos"]) >= 3 and fo_ok >= 0.6,
        9: pl["score"] >= 0.75 and perda["ok"] and perda["cabe"] and n_resp >= 5,
    }
    nivel = 0
    for k in range(1, 10):
        if cond[k]:
            nivel = k
        else:
            break
    return {"nivel": nivel, "cond": cond, "a_rec": a1["rec_bruta"]}


def perguntas_banca(proj: dict, d: dict, res: dict, mc: dict | None, preco_ref: float) -> list:
    """Perguntas que uma banca faria, geradas a partir das fragilidades do próprio plano."""
    P = []

    def add(tema, pergunta, motivo):
        P.append({"tema": tema, "pergunta": pergunta, "motivo": motivo})

    a1 = res["anos"][1]
    rec = a1["rec_bruta"]
    m = mercado(proj, res)
    pl = placar_evidencias(proj, d, res)
    ent = resumo_entrevistas(d, preco_ref)
    vw = van_westendorp(d.get("precos_vw"))
    rr = res.get("recorrencia", {})
    if pl["score"] < 0.5:
        add("Evidências", f"Quais números do plano foram testados com clientes reais? Por que devemos acreditar em "
                          f"\"{pl['proxima']}\"?", f"Placar de evidências em {pct(pl['score'], 0)}.")
    if ent["n"] < 10:
        add("Cliente", "Com quantos clientes do segmento-alvo vocês conversaram? O que aprenderam que mudou o plano?",
            f"{ent['n']} entrevista(s) registrada(s).")
    if vw and vw.get("ok") and preco_ref == preco_ref and not (vw["pmc"] <= preco_ref <= vw["pme"]):
        lado = "acima" if preco_ref > vw["pme"] else "abaixo"
        add("Preço", f"O preço médio ({brl(preco_ref)}) está {lado} da faixa que os clientes consideram aceitável "
                     f"({brl(vw['pmc'])} a {brl(vw['pme'])}). Como justificam?", "Teste de Van Westendorp.")
    if m["preenchido"] and m["pct_som"] == m["pct_som"] and m["pct_som"] > 1:
        add("Mercado", "A receita do Ano 1 é maior que o mercado que vocês dizem conseguir conquistar (SOM). De onde vêm esses clientes?",
            f"Receita = {pct(m['pct_som'], 0)} do SOM.")
    if res["vpl"] <= 0:
        add("Viabilidade", "O VPL é negativo. O que mudaria no modelo de negócio para torná-lo viável, e por que ainda vale a pena?",
            f"VPL de {brl(res['vpl'])}.")
    tir = res.get("tir")
    if tir == tir and tir is not None and res["vpl"] > 0 and tir < res["tma"] * 1.3:
        add("Retorno", "A TIR é pouco maior que a TMA. Por que não deixar o dinheiro aplicado, com menos risco?",
            f"TIR {pct(tir)} x TMA {pct(res['tma'])}.")
    minimo = float(res["fluxo_mensal"]["Saldo final"].min())
    if minimo < 0:
        add("Caixa", "O caixa fica negativo no primeiro ano. Quem cobre esse buraco e em que mês?", f"Menor saldo: {brl(minimo)}.")
    if a1["imc"] == a1["imc"] and a1["imc"] < 0.3:
        add("Margem", "A margem de contribuição é baixa. O que acontece se o fornecedor aumentar o preço em 10%?",
            f"Margem de contribuição de {pct(a1['imc'])}.")
    if rec > 0 and div(a1["pessoal"], rec) > 0.45:
        add("Equipe", "A folha consome quase metade da receita. A equipe não está grande demais para o início?",
            f"Pessoal = {pct(div(a1['pessoal'], rec), 0)} da receita.")
    if rr.get("tem"):
        if rr.get("churn", 0) > 0.05:
            add("Retenção", "O churn mensal passa de 5%. O que fará o assinante continuar depois do terceiro mês?",
                f"Churn de {pct(rr['churn'])} ao mês.")
        lc = rr.get("ltv_cac", float("nan"))
        if lc == lc and lc < 3:
            add("Aquisição", "O LTV é menor que três vezes o CAC. O crescimento não vai queimar caixa?", f"LTV/CAC = {num_br(lc, 1)}.")
    if mc is not None and mc["p_vpl_pos"] < 0.6:
        add("Risco", "Em quase metade das simulações o negócio não se paga. Qual é o plano B?",
            f"Chance de VPL positivo: {pct(mc['p_vpl_pos'], 0)}.")
    perda = perda_aceitavel(proj, d, res)
    if perda["ok"] and not perda["cabe"]:
        add("Sócios", "Se der errado, vocês podem perder mais do que disseram aceitar. Como reduzir o investimento inicial?",
            f"Perda potencial {brl(perda['potencial'])} x aceitável {brl(perda['aceitavel'])}.")
    errc = d.get("errc", pd.DataFrame())
    if errc.empty or not (errc["acao"] == "Criar").any():
        add("Diferenciação", "O que vocês oferecem que nenhum concorrente oferece? Por que um concorrente grande não copiaria amanhã?",
            "Nenhum atributo novo na curva de valor.")
    forn = d["fornecedores"]
    if not forn.empty and int(forn["critico"].sum()) >= 1 and len(forn) <= 2:
        add("Operação", "Há dependência de poucos fornecedores críticos. O que acontece se o principal falhar?", "Fornecedores críticos.")
    add("Timing", "Por que este é o momento certo para abrir este negócio?", "Pergunta clássica de banca.")
    add("Equipe", "Por que este grupo é o mais capaz de executar este plano?", "Pergunta clássica de banca.")
    return P


def preco_referencia(proj: dict, res: dict) -> float:
    """Preço do produto escolhido no teste de preço; sem escolha, preço médio ponderado (ou ARPU das assinaturas)."""
    cat = res.get("cat", pd.DataFrame())
    alvo = str(proj.get("vw_produto") or "")
    if alvo and not cat.empty and alvo in set(cat["descricao"]):
        return float(cat.loc[cat["descricao"] == alvo, "preco_venda"].iloc[0])
    if alvo.startswith("Assinatura") or (cat.empty and res.get("recorrencia", {}).get("tem")):
        return float(res["recorrencia"].get("arpu", float("nan")))
    if cat.empty or float(cat["qtd_mensal"].sum()) <= 0:
        return float("nan")
    return div(float((cat["preco_venda"] * cat["qtd_mensal"]).sum()), float(cat["qtd_mensal"].sum()))


def regras_v71(proj: dict, d: dict) -> dict:
    """Situação de preenchimento das ferramentas novas: (completo, iniciado)."""
    def n(t):
        x = d.get(t)
        return 0 if x is None else len(x)
    pos = [proj.get(k) for k in ("posic_segmento", "posic_beneficio", "posic_diferencial")]
    vpc = [proj.get(k) for k, _, _ in BLOCOS_VPC]
    ev = d.get("evidencias", pd.DataFrame(columns=["status"]))
    fo = d.get("formalizacao", pd.DataFrame(columns=["status"]))
    resp = d.get("banca_respostas", pd.DataFrame(columns=["resposta"]))
    n_resp = int((resp["resposta"].astype(str).str.len() > 20).sum()) if len(resp) else 0
    errc = d.get("errc", pd.DataFrame(columns=["acao"]))
    rec = n("catalogo") > 0 or n("assinaturas") > 0
    return {
        "entrevistas": (n("entrevistas") >= 10, n("entrevistas") >= 1),
        "vpc": (all(vpc), any(vpc)),
        "oceano": (n("errc") >= 3 and bool((errc["acao"] == "Criar").any()), n("errc") >= 1),
        "posicionamento": (all(pos) and n("mapa_perceptual") >= 2, any(pos) or n("mapa_perceptual") >= 1),
        "orcamento_mkt": (n("orcamento_canais") >= 1 and n("aarrr") >= 5, n("orcamento_canais") >= 1 or n("calendario_mkt") >= 1),
        "teste_preco": (n("precos_vw") >= 5 or float(proj.get("intencao_certamente_pct") or 0) > 0,
                        n("precos_vw") >= 1 or float(proj.get("intencao_certamente_pct") or 0) > 0),
        "montecarlo": (rec and float(d["socios"]["perda_aceitavel"].sum()) > 0, rec),
        "adocao": (float(proj.get("sam_clientes") or 0) > 0 and float(proj.get("gasto_anual_cliente") or 0) > 0,
                   float(proj.get("sam_clientes") or 0) > 0 or float(proj.get("bass_mercado") or 0) > 0),
        "evidencias": (len(ev) >= 3 and bool((ev["status"] != "Estimada (sem fonte)").any()), len(ev) >= 1),
        "banca": (n_resp >= 5, n_resp >= 1),
        "formalizacao": (len(fo) > 0 and float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).mean()) >= 0.6,
                         len(fo) >= 1),
    }
