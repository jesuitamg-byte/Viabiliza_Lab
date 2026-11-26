"""
Mix de marketing integrado (v8.1 em diante): regras e cálculos dos 4Ps/7Ps,
sem interface. As funções recebem o projeto (proj), as tabelas (d) e o resultado
do motor financeiro (res) e devolvem números, tabelas e alertas.

Referências usadas nas regras:
- Kotler; Keller (2012): níveis de produto, ciclo de vida, estratégias de preço.
- Booms; Bitner (1981): 7Ps de serviços.
- Lauterborn (1990): 4Cs.
- Van Westendorp (1976): faixa aceitável de preço.
"""

import math

import pandas as pd

from pnbox_motor import div


def _p(x: float, casas: int = 1) -> str:
    """Percentual no formato brasileiro (0,257 -> 25,7%)."""
    return f"{x * 100:.{casas}f}%".replace(".", ",")


# ------------------------------------------------------------------------------
# PRODUTO
# ------------------------------------------------------------------------------
CICLO_VIDA = {
    "Introdução": "Vendas baixas e custo alto de conquistar cliente. Foco em fazer o produto ser conhecido e "
                  "experimentado; o preço costuma ser de penetração ou de desnatação.",
    "Crescimento": "As vendas sobem e surgem concorrentes. Foco em ampliar a distribuição e reforçar a marca.",
    "Maturidade": "O mercado para de crescer e a disputa é por participação. Foco em diferenciar, "
                  "fidelizar e controlar custos; guerra de preço corrói a margem.",
    "Declínio": "As vendas caem. Foco em reduzir custos, enxugar a linha ou reposicionar o produto.",
}


def pareto_produtos(cat: pd.DataFrame) -> dict:
    """Quantos itens respondem por 80% da receita (concentração do portfólio)."""
    if cat is None or cat.empty or cat["receita_mensal"].sum() <= 0:
        return {"n": 0, "total": 0, "itens": []}
    c = cat.sort_values("receita_mensal", ascending=False)
    acum = c["receita_mensal"].cumsum() / c["receita_mensal"].sum()
    n = int((acum < 0.8).sum()) + 1
    return {"n": min(n, len(c)), "total": len(c), "itens": c["descricao"].head(n).tolist(),
            "share_top": float(c["receita_mensal"].iloc[0] / c["receita_mensal"].sum())}


# ------------------------------------------------------------------------------
# PREÇO
# ------------------------------------------------------------------------------
ESTRATEGIAS_PRECO = {
    "Penetração": ("Preço abaixo do mercado para ganhar clientes depressa. Exige volume alto e custo baixo.", "abaixo"),
    "Desnatação": ("Preço alto no lançamento, para quem valoriza a novidade; cai depois.", "acima"),
    "Valor percebido": ("Preço definido pelo benefício que o cliente enxerga, medido em pesquisa.", None),
    "Custo mais margem": ("Custo unitário mais uma margem (markup). Simples, mas ignora o cliente e a concorrência.", None),
    "Referência na concorrência": ("Preço alinhado ao praticado pelos concorrentes diretos.", "igual"),
}
POSICOES_PRECO = {
    "Econômico": (0.0, 0.95),
    "Intermediário": (0.90, 1.10),
    "Premium": (1.05, 99.0),
}


def mc_produtos(res: dict) -> pd.DataFrame:
    """Margem de contribuição percentual de cada item: preço − custo direto − impostos − outros variáveis."""
    cat = res.get("cat")
    if cat is None or cat.empty:
        return pd.DataFrame(columns=["Produto", "Preço", "Custo direto", "MC %", "Peso na receita"])
    aliq, var = float(res.get("aliq") or 0), float(res.get("var_pct") or 0)
    tot = float(cat["receita_mensal"].sum()) or 1.0
    return pd.DataFrame({
        "Produto": cat["descricao"], "Preço": cat["preco_venda"], "Custo direto": cat["custo_unit"],
        "MC %": [div(p - c, p) - aliq - var if p > 0 else float("nan")
                 for p, c in zip(cat["preco_venda"], cat["custo_unit"])],
        "Peso na receita": cat["receita_mensal"] / tot,
    })


def mc_media(res: dict) -> float:
    mc = mc_produtos(res)
    if mc.empty:
        return float("nan")
    ok = mc.dropna(subset=["MC %"])
    return float((ok["MC %"] * ok["Peso na receita"]).sum() / ok["Peso na receita"].sum()) if ok["Peso na receita"].sum() > 0 else float("nan")


def volume_para_desconto(mc: float, desconto: float) -> float:
    """Aumento de volume necessário para manter a margem total após um desconto (frações)."""
    if mc != mc or mc <= 0 or desconto >= mc:
        return float("inf")
    return desconto / (mc - desconto)


def queda_tolerada_aumento(mc: float, aumento: float) -> float:
    """Quanto o volume pode cair, após um aumento de preço, sem reduzir a margem total (frações)."""
    if mc != mc or mc + aumento <= 0:
        return float("nan")
    return aumento / (mc + aumento)


def comparar_concorrentes(cat: pd.DataFrame, precos: pd.DataFrame) -> pd.DataFrame:
    """Preço próprio x média dos concorrentes, produto a produto."""
    if cat is None or cat.empty or precos is None or precos.empty:
        return pd.DataFrame(columns=["Produto", "Seu preço", "Média dos concorrentes", "Menor", "Maior", "Índice"])
    p = precos[(precos["preco"] > 0) & (precos["produto"].astype(str).str.strip() != "")]
    linhas = []
    for prod, g in p.groupby("produto"):
        meu = cat.loc[cat["descricao"] == prod, "preco_venda"]
        if meu.empty:
            continue
        m = float(meu.iloc[0])
        media = float(g["preco"].mean())
        linhas.append({"Produto": prod, "Seu preço": m, "Média dos concorrentes": media,
                       "Menor": float(g["preco"].min()), "Maior": float(g["preco"].max()),
                       "Índice": div(m, media)})
    return pd.DataFrame(linhas)


def indice_preco_medio(cat: pd.DataFrame, comp: pd.DataFrame) -> float:
    """Índice de preço (seu preço ÷ média dos concorrentes), ponderado pela receita dos itens comparados."""
    if comp.empty:
        return float("nan")
    pesos = [float(cat.loc[cat["descricao"] == p, "receita_mensal"].sum()) for p in comp["Produto"]]
    if sum(pesos) <= 0:
        return float(comp["Índice"].mean())
    return float(sum(i * w for i, w in zip(comp["Índice"], pesos)) / sum(pesos))


def checar_produto_preco(proj: dict, d: dict, res: dict, vw: dict | None = None) -> list[tuple[str, str]]:
    """Alertas de Produto e Preço. Cada item é (nível, mensagem) com nível 'ok', 'atencao' ou 'critico'."""
    out = []
    cat = res.get("cat")
    if cat is None or cat.empty:
        return [("atencao", "Cadastre os produtos ou serviços para que o programa confira preço e margem.")]
    # Margens
    mc = mc_produtos(res)
    neg = mc[mc["MC %"] <= 0]["Produto"].tolist()
    if neg:
        out.append(("critico", "Itens vendidos abaixo do custo variável: " + ", ".join(neg) +
                    ". Cada venda deles aumenta o prejuízo."))
    mm = mc_media(res)
    if mm == mm:
        nivel = "ok" if mm >= 0.30 else ("atencao" if mm >= 0.15 else "critico")
        out.append((nivel, f"Margem de contribuição média de {_p(mm)}. "
                    + ("Há folga para custos fixos e descontos." if nivel == "ok" else
                       "Pouca folga: qualquer desconto ou aumento de custo pesa muito." if nivel == "atencao" else
                       "Margem muito estreita: o negócio depende de volume muito alto para pagar os custos fixos.")))
    # Concorrência e posição de preço
    comp = comparar_concorrentes(cat, d.get("precos_concorrentes"))
    idx = indice_preco_medio(cat, comp)
    pos = proj.get("mix_posicao_preco") or ""
    est = proj.get("mix_estrategia_preco") or ""
    if idx == idx:
        txt = f"Seus preços estão, em média, {_p(abs(idx - 1), 0)} {'acima' if idx >= 1 else 'abaixo'} dos concorrentes pesquisados."
        if pos in POSICOES_PRECO:
            lo, hi = POSICOES_PRECO[pos]
            if not (lo <= idx <= hi):
                out.append(("atencao", txt + f" Isso não combina com o posicionamento {pos.lower()} declarado."))
            else:
                out.append(("ok", txt + f" Coerente com o posicionamento {pos.lower()}."))
        else:
            out.append(("ok", txt))
        esperado = ESTRATEGIAS_PRECO.get(est, ("", None))[1]
        if esperado == "abaixo" and idx >= 1:
            out.append(("atencao", "A estratégia é de penetração, mas os preços não estão abaixo da concorrência."))
        if esperado == "acima" and idx <= 1:
            out.append(("atencao", "A estratégia é de desnatação, mas os preços não estão acima da concorrência."))
        if esperado == "igual" and not (0.9 <= idx <= 1.1):
            out.append(("atencao", "A estratégia é seguir a concorrência, mas a diferença passa de 10%."))
    else:
        out.append(("atencao", "Sem preços de concorrentes cadastrados: a comparação de mercado não pôde ser feita."))
    if est == "Penetração" and mm == mm and mm < 0.25:
        out.append(("atencao", "Penetração com margem estreita é arriscada: confira se o volume previsto é realista."))
    # Van Westendorp
    if vw and vw.get("ok"):
        alvo = proj.get("vw_produto") or ""
        linha = cat[cat["descricao"] == alvo]
        preco = float(linha["preco_venda"].iloc[0]) if not linha.empty else float("nan")
        if preco == preco and vw["pmc"] == vw["pmc"] and vw["pme"] == vw["pme"]:
            if preco < vw["pmc"]:
                out.append(("atencao", f"O preço de {alvo} está abaixo da faixa aceitável da pesquisa: pode passar imagem de baixa qualidade."))
            elif preco > vw["pme"]:
                out.append(("critico", f"O preço de {alvo} está acima da faixa aceitável da pesquisa: muitos clientes o acham caro demais."))
            else:
                out.append(("ok", f"O preço de {alvo} está dentro da faixa aceitável da pesquisa de preço."))
    elif est == "Valor percebido":
        out.append(("atencao", "A estratégia é de valor percebido, mas não há pesquisa de preço (Van Westendorp) com respostas suficientes."))
    # Ciclo de vida
    ciclo = proj.get("mix_ciclo") or ""
    if ciclo == "Maturidade" and est == "Desnatação":
        out.append(("atencao", "Desnatação funciona no lançamento; num mercado maduro o cliente já conhece as alternativas."))
    if ciclo == "Declínio" and est == "Penetração":
        out.append(("atencao", "Penetração num mercado em declínio dificilmente compensa: o volume total está caindo."))
    # Concentração
    pt = pareto_produtos(cat)
    if pt["total"] >= 3 and pt["n"] == 1:
        out.append(("atencao", f"Um único item ({pt['itens'][0]}) responde por {_p(pt['share_top'], 0)} da receita: "
                    "o plano fica muito dependente dele."))
    return out


# ------------------------------------------------------------------------------
# PREÇO: CONDIÇÕES DE PAGAMENTO
# ------------------------------------------------------------------------------
PAGAMENTO_EXEMPLO = pd.DataFrame([
    {"forma": "Pix", "pct_vendas": 35.0, "taxa_pct": 0.0, "prazo_dias": 0.0},
    {"forma": "Dinheiro", "pct_vendas": 10.0, "taxa_pct": 0.0, "prazo_dias": 0.0},
    {"forma": "Cartão de débito", "pct_vendas": 20.0, "taxa_pct": 1.5, "prazo_dias": 1.0},
    {"forma": "Cartão de crédito à vista", "pct_vendas": 25.0, "taxa_pct": 3.5, "prazo_dias": 30.0},
    {"forma": "Cartão de crédito parcelado", "pct_vendas": 10.0, "taxa_pct": 6.0, "prazo_dias": 60.0},
])


def resumo_pagamento(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"soma": 0.0, "taxa": 0.0, "prazo": 0.0}
    w = df["pct_vendas"].astype(float)
    soma = float(w.sum())
    taxa = float((w * df["taxa_pct"]).sum() / 100) if soma else 0.0          # % sobre a receita total
    prazo = float((w * df["prazo_dias"]).sum() / soma) if soma else 0.0
    return {"soma": soma, "taxa": taxa, "prazo": prazo}


def aplicar_pagamento(df_pag: pd.DataFrame, cv: pd.DataFrame, prazos: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Atualiza 'Taxas de cartões' nos custos variáveis e substitui os prazos de recebimento."""
    r = resumo_pagamento(df_pag)
    cv = cv.copy()
    mask = cv["descricao"].astype(str).str.lower().str.startswith("taxas de cart")
    if mask.any():
        cv.loc[mask, "percentual"] = 0.0
        cv.loc[cv.index[mask][0], "percentual"] = round(r["taxa"], 4)
    else:
        cv = pd.concat([cv, pd.DataFrame([{"descricao": "Taxas de cartões", "percentual": round(r["taxa"], 4)}])],
                       ignore_index=True)
    g = df_pag[df_pag["pct_vendas"] > 0].groupby("prazo_dias", as_index=False)["pct_vendas"].sum()
    novos = pd.DataFrame({"tipo": "Recebimento", "prazo_dias": g["prazo_dias"], "percentual": g["pct_vendas"]})
    prazos = pd.concat([prazos[prazos["tipo"] != "Recebimento"], novos], ignore_index=True)
    return cv, prazos
