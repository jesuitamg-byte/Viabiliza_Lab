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


def _n(x: float, casas: int = 0) -> str:
    """Número no formato brasileiro (1234,5 -> 1.234,5)."""
    t = f"{x:,.{casas}f}"
    return t.replace(",", "X").replace(".", ",").replace("X", ".")


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


# ------------------------------------------------------------------------------
# PRAÇA (v8.2)
# ------------------------------------------------------------------------------
CANAIS_VENDA = ["Loja física", "Entrega própria", "Aplicativo de delivery", "Marketplace", "Site próprio",
                "WhatsApp e redes sociais", "Revendedores", "Venda direta a empresas", "Feiras e eventos"]
DISTRIBUICAO = {
    "Intensiva": "Estar no maior número possível de pontos de venda. Combina com produtos de compra frequente e preço baixo.",
    "Seletiva": "Poucos pontos ou canais escolhidos. Combina com produtos em que o cliente compara antes de comprar.",
    "Exclusiva": "Um ou poucos pontos, com controle total da experiência. Combina com posicionamento premium.",
}
_LINHA_COMISSAO = "Comissões de canais de venda"
TERCEIROS = {"Aplicativo de delivery", "Marketplace", "Revendedores"}


def resumo_canais(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"soma": 0.0, "comissao": 0.0, "maior": "", "maior_pct": 0.0}
    w = df["pct_vendas"].astype(float)
    soma = float(w.sum())
    i = int(w.values.argmax())
    return {"soma": soma, "comissao": float((w * df["comissao_pct"]).sum() / 100) if soma else 0.0,
            "maior": str(df["canal"].iloc[i]), "maior_pct": float(w.iloc[i])}


def aplicar_canais(df: pd.DataFrame, cv: pd.DataFrame) -> pd.DataFrame:
    """Grava a comissão média ponderada nos custos variáveis, zerando a linha antiga de aplicativos para não contar duas vezes."""
    r = resumo_canais(df)
    cv = cv.copy()
    antiga = cv["descricao"].astype(str).str.lower().str.startswith("comissões de aplicativos")
    cv.loc[antiga, "percentual"] = 0.0
    m = cv["descricao"] == _LINHA_COMISSAO
    if m.any():
        cv.loc[m, "percentual"] = round(r["comissao"], 4)
    else:
        cv = pd.concat([cv, pd.DataFrame([{"descricao": _LINHA_COMISSAO, "percentual": round(r["comissao"], 4)}])],
                       ignore_index=True)
    return cv


# ------------------------------------------------------------------------------
# PROMOÇÃO (v8.2)
# ------------------------------------------------------------------------------
OBJETIVOS_COMUNICACAO = {
    "Tornar a marca conhecida": "Alcance e frequência: anúncios, parcerias, presença em eventos e conteúdo.",
    "Fazer o cliente experimentar": "Amostras, degustação, primeira compra com vantagem e indicação.",
    "Converter em venda": "Oferta clara, prova social, facilidade de pagamento e contato rápido.",
    "Fidelizar e fazer voltar": "Programa de fidelidade, pós-venda, relacionamento por WhatsApp ou e-mail.",
}


def clientes_necessarios(proj: dict, res: dict) -> dict:
    """Clientes ativos por mês que o volume do plano exige (receita do mês 1 ÷ gasto mensal por cliente)."""
    cat = res.get("cat")
    receita = float(res.get("receita_m1") or 0)
    ticket = float(proj.get("mix_ticket") or 0)
    if ticket <= 0 and cat is not None and not cat.empty and cat["qtd_mensal"].sum() > 0:
        ticket = receita / float(cat["qtd_mensal"].sum())
    freq = float(proj.get("mix_freq") or 0)
    if receita <= 0 or ticket <= 0 or freq <= 0:
        return {"ok": False, "ticket": ticket, "freq": freq}
    return {"ok": True, "ticket": ticket, "freq": freq, "clientes": receita / (ticket * freq),
            "gasto_cliente_mes": ticket * freq}


def promocao_vs_volume(proj: dict, d: dict, res: dict) -> dict:
    """Confronta os clientes que o plano exige com os clientes novos que a promoção promete trazer."""
    cn = clientes_necessarios(proj, res)
    oc = d.get("orcamento_canais")
    novos = float(oc["clientes_mes"].sum()) if oc is not None and not oc.empty else 0.0
    gasto = float(oc["gasto_mensal"].sum()) if oc is not None and not oc.empty else 0.0
    out = {"cn": cn, "novos": novos, "gasto": gasto, "cac": div(gasto, novos) if novos > 0 else float("nan")}
    if cn["ok"] and novos > 0:
        out["meses_base"] = cn["clientes"] / novos
    mm = mc_media(res)
    if cn.get("ok") and mm == mm and mm > 0 and out["cac"] == out["cac"]:
        margem_cliente = cn["gasto_cliente_mes"] * mm
        out["payback_cac"] = div(out["cac"], margem_cliente)
    receita = float(res.get("receita_m1") or 0)
    out["verba_receita"] = div(gasto, receita) if receita > 0 else float("nan")
    return out


def checar_praca_promocao(proj: dict, d: dict, res: dict) -> list[tuple[str, str]]:
    out = []
    cv = d.get("canais_venda")
    if cv is None or cv.empty:
        out.append(("atencao", "Os canais de venda (onde o cliente compra e quanto cada canal cobra) ainda não foram informados."))
    else:
        r = resumo_canais(cv)
        if abs(r["soma"] - 100) > 0.5:
            out.append(("atencao", f"Os canais de venda somam {_n(r['soma'])}% das vendas; precisam somar 100%."))
        mm = mc_media(res)
        if r["comissao"] > 0 and mm == mm:
            cvv = d["custos_variaveis"]
            desc = cvv["descricao"].astype(str)
            aplicado = float(cvv.loc[(desc == _LINHA_COMISSAO) | desc.str.lower().str.startswith("comissões de aplicativos"),
                                     "percentual"].sum())
            if abs(aplicado - r["comissao"]) > 0.01:
                out.append(("atencao", f"A comissão média dos canais ({_p(r['comissao'] / 100, 2)} da receita) ainda não foi "
                            "aplicada aos custos variáveis: a margem do plano está maior do que a real."))
        if r["maior_pct"] >= 70 and r["maior"] in TERCEIROS:
            out.append(("atencao", f"{r['maior']} concentra {_n(r['maior_pct'])}% das vendas. Se esse canal mudar as regras ou a taxa, "
                        "o negócio sente direto."))
        pos, dist = proj.get("mix_posicao_preco") or "", proj.get("mix_distribuicao") or ""
        mkt = float(cv.loc[cv["canal"].isin(["Marketplace", "Aplicativo de delivery"]), "pct_vendas"].sum())
        if pos == "Premium" and (dist == "Intensiva" or mkt >= 50):
            out.append(("atencao", "Posição premium com distribuição intensiva ou vendas concentradas em marketplace e aplicativos: "
                        "o cliente compara preço lado a lado e a diferenciação some."))
    pv = promocao_vs_volume(proj, d, res)
    if pv["novos"] <= 0:
        out.append(("atencao", "Não há orçamento de promoção por canal com clientes novos por mês: não dá para saber de onde vem o volume do plano."))
    elif not pv["cn"]["ok"]:
        out.append(("atencao", "Informe quantas vezes por mês o cliente compra para comparar o volume do plano com a promoção."))
    else:
        mb = pv.get("meses_base", float("nan"))
        nivel = "ok" if mb <= 3 else ("atencao" if mb <= 6 else "critico")
        cvn = d.get("canais_venda")
        loja = float(cvn.loc[cvn["canal"] == "Loja física", "pct_vendas"].sum()) if cvn is not None and not cvn.empty else 0.0
        if nivel == "critico" and loja >= 50:
            nivel = "atencao"
        out.append((nivel, f"O plano precisa de cerca de {_n(pv['cn']['clientes'])} clientes ativos por mês, e a promoção traz "
                    f"{_n(pv['novos'])} novos por mês: são uns {_n(mb, 1)} meses para formar essa base, sem contar quem deixa de comprar."
                    + (" O volume do primeiro mês parece otimista." if nivel != "ok" else "")
                    + (" Numa loja de rua parte dos clientes vem do fluxo de pessoas e não da promoção paga; "
                       "estime esse fluxo no plano operacional para sustentar o volume." if loja >= 50 and nivel != "ok" else "")))
        pb = pv.get("payback_cac", float("nan"))
        if pb == pb:
            nivel = "ok" if pb <= 3 else ("atencao" if pb <= 12 else "critico")
            out.append((nivel, f"Cada cliente novo custa R$ {_n(pv['cac'], 2)} e devolve isso em margem de contribuição em cerca de {_n(pb, 1)} meses."))
    return out


# ------------------------------------------------------------------------------
# 7Ps DE SERVIÇOS (v8.3) - Booms; Bitner (1981)
# ------------------------------------------------------------------------------
LINHAS_BLUEPRINT = ["Visível ao cliente", "Bastidores", "Apoio (sistemas e fornecedores)"]


def resumo_blueprint(df: pd.DataFrame) -> dict:
    """Mapa do serviço (Shostack, 1984): tempo total, tempo visível ao cliente e pontos de falha."""
    if df is None or df.empty:
        return {"etapas": 0, "tempo": 0.0, "visivel": 0.0, "falhas": []}
    t = df["tempo_min"].astype(float)
    vis = df["linha"] == "Visível ao cliente"
    falhas = df.loc[df["ponto_falha"].astype(str).str.strip() != "", "etapa"].tolist()
    return {"etapas": len(df), "tempo": float(t.sum()), "visivel": float(t[vis].sum()), "falhas": falhas}


def e_servico(proj: dict, res: dict) -> bool:
    tipo = str(proj.get("tipo_negocio") or "")
    if tipo in ("Serviços", "Assinatura", "Misto"):
        return True
    cat = res.get("cat")
    if cat is not None and not cat.empty and "categoria" in cat:
        return bool((cat["categoria"] == "Serviço").any())
    return False


# ------------------------------------------------------------------------------
# COERÊNCIA DO MIX (v8.3)
# ------------------------------------------------------------------------------
def coerencia_mix(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    """Combinações entre os Ps que se reforçam ou se contradizem. Cada linha: combinação, situação e comentário."""
    pos = proj.get("mix_posicao_preco") or ""
    est = proj.get("mix_estrategia_preco") or ""
    dist = proj.get("mix_distribuicao") or ""
    obj = proj.get("mix_objetivo_com") or ""
    ciclo = proj.get("mix_ciclo") or ""
    ampliado = str(proj.get("mix_ampliado") or "").strip()
    mm = mc_media(res)
    linhas = []

    def add(comb, ok, bom, ruim):
        linhas.append({"Combinação": comb, "Situação": "✅ coerente" if ok else "⚠️ contradição",
                       "Comentário": bom if ok else ruim})

    if pos == "Premium":
        add("Preço premium × produto ampliado", bool(ampliado),
            "O preço mais alto vem acompanhado de serviços e garantias que o justificam.",
            "Preço premium sem nada além do produto básico: o cliente não vê por que pagar mais.")
        if dist:
            add("Preço premium × distribuição", dist != "Intensiva",
                "Distribuição seletiva ou exclusiva protege a imagem e o controle da experiência.",
                "Distribuição intensiva espalha o produto em pontos onde ele disputa preço lado a lado.")
        if est:
            add("Preço premium × estratégia de preço", est != "Penetração",
                "A estratégia de preço sustenta a posição premium.",
                "Penetração é preço baixo para ganhar volume: contradiz a posição premium.")
    if pos == "Econômico":
        if dist:
            add("Preço econômico × distribuição", dist != "Exclusiva",
                "Preço baixo pede volume, e a distribuição ajuda a alcançá-lo.",
                "Preço baixo em poucos pontos exclusivos: falta volume para compensar a margem menor.")
        if mm == mm:
            add("Preço econômico × margem", mm >= 0.20,
                "Mesmo com preço baixo, a margem de contribuição deixa folga para os custos fixos.",
                "Preço econômico com margem estreita: o negócio depende de um volume muito alto.")
    if est == "Penetração" and obj:
        add("Penetração × objetivo de comunicação", obj in ("Tornar a marca conhecida", "Fazer o cliente experimentar"),
            "A comunicação busca alcance e experimentação, que é o que a penetração precisa.",
            "Penetração depende de muita gente conhecendo e experimentando; a comunicação está focada em outra coisa.")
    if est == "Desnatação" and ciclo:
        add("Desnatação × ciclo de vida", ciclo == "Introdução",
            "Preço alto no lançamento, quando o produto ainda é novidade.",
            "Desnatação perde força depois do lançamento, quando o cliente já conhece as alternativas.")
    if ciclo and obj:
        add("Ciclo de vida × objetivo de comunicação",
            not (ciclo == "Introdução" and obj == "Fidelizar e fazer voltar"),
            "O objetivo da comunicação combina com a fase do produto.",
            "Na introdução quase ninguém conhece o produto: antes de fidelizar é preciso ser conhecido e experimentado.")
    cv = d.get("canais_venda")
    if cv is not None and not cv.empty and dist:
        n = int((cv["pct_vendas"] > 0).sum())
        add("Distribuição declarada × canais usados", not (dist == "Intensiva" and n <= 1) and not (dist == "Exclusiva" and n >= 4),
            "Os canais cadastrados batem com a intensidade de distribuição escolhida.",
            "O número de canais cadastrados não bate com a intensidade de distribuição escolhida.")
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 4Cs (v8.3) - Lauterborn (1990)
# ------------------------------------------------------------------------------
QUATRO_CS = [
    ("mix_4c_n1", "mix_4c_cliente", "Cliente (em vez de produto)",
     "O produto resolve uma necessidade que o cliente reconhece? Que dor das personas ele alivia?"),
    ("mix_4c_n2", "mix_4c_custo", "Custo para o cliente (em vez de preço)",
     "Quanto custa ao cliente comprar, somando preço, deslocamento, tempo e o risco de não gostar?"),
    ("mix_4c_n3", "mix_4c_conveniencia", "Conveniência (em vez de praça)",
     "É fácil encontrar, comprar e pagar? Horários, entrega, formas de pagamento e atendimento ajudam?"),
    ("mix_4c_n4", "mix_4c_comunicacao", "Comunicação (em vez de promoção)",
     "Há conversa com o cliente, e não só anúncio? Ele consegue perguntar, reclamar e indicar?"),
]


# ------------------------------------------------------------------------------
# PLANO DE MARKETING (v8.4): situação, objetivos, estratégia, táticas, ações e controle
# Smith (2011), McDonald (2007) e Las Casas (2011)
# ------------------------------------------------------------------------------
ANSOFF = {
    "Penetração de mercado": ("Produto atual, mercado atual", "Vender mais do mesmo para o mesmo público: frequência, fidelidade e "
                              "conquista de clientes dos concorrentes. É o caminho de menor risco."),
    "Desenvolvimento de produto": ("Produto novo, mercado atual", "Oferecer novos produtos ou serviços a quem já é cliente. "
                                   "Risco moderado: exige criar e testar o novo produto."),
    "Desenvolvimento de mercado": ("Produto atual, mercado novo", "Levar o produto a outros bairros, cidades, públicos ou canais. "
                                   "Risco moderado: exige conhecer o novo público."),
    "Diversificação": ("Produto novo, mercado novo", "Entrar com produto novo em mercado novo. É o caminho de maior risco, "
                       "porque nada do que a empresa já sabe se aplica diretamente."),
}
GENERICAS = {
    "Liderança em custo": "Ser o mais barato de produzir e vender, com preço competitivo.",
    "Diferenciação": "Oferecer algo que o cliente valoriza e pelo qual aceita pagar mais.",
    "Foco em custo": "Atender um nicho específico com o menor custo.",
    "Foco em diferenciação": "Atender um nicho específico com uma oferta feita sob medida para ele.",
}
FREQUENCIAS = ["Diária", "Semanal", "Mensal", "Trimestral", "Semestral", "Anual"]
KPIS_SUGERIDOS = [
    ("Clientes novos por mês", "Mensal", "Cadastro de clientes ou vendas"),
    ("Custo por cliente novo (CAC)", "Mensal", "Gasto de marketing ÷ clientes novos"),
    ("Gasto médio por compra (ticket médio)", "Semanal", "Sistema de caixa"),
    ("Taxa de recompra em 30 dias", "Mensal", "Cadastro de clientes ou cartão fidelidade"),
    ("Conversão de contatos em vendas", "Semanal", "WhatsApp, redes sociais ou CRM"),
    ("Nota de satisfação (NPS)", "Trimestral", "Pesquisa com clientes"),
    ("Receita por canal de venda", "Mensal", "Relatórios dos canais e do caixa"),
]
ETAPAS_PLANO = [
    ("Situação", "Onde estamos? Mercado, clientes, concorrentes e fatores internos e externos."),
    ("Objetivos", "Aonde queremos chegar? Objetivos de marketing mensuráveis e com prazo."),
    ("Estratégia", "Como chegar lá? Público-alvo, posicionamento e direção de crescimento."),
    ("Táticas", "Com que ferramentas? O mix de marketing (4Ps ou 7Ps)."),
    ("Ações", "Quem faz o quê e quando? Calendário, responsáveis e custos."),
    ("Controle", "Chegamos lá? Indicadores, metas e frequência de acompanhamento."),
]


def _tem_txt(x) -> bool:
    return bool(str(x or "").strip())


def objetivos_marketing(d: dict) -> pd.DataFrame:
    ob = d.get("objetivos")
    if ob is None or ob.empty:
        return pd.DataFrame(columns=["objetivo", "indicador", "meta", "unidade", "prazo_meses"])
    return ob[ob["perspectiva"] == "Clientes"]


def etapas_plano_mkt(proj: dict, d: dict) -> list[tuple[str, int, str]]:
    """Situação de cada etapa: (etapa, 0 = vazia / 1 = parcial / 2 = completa, o que falta)."""
    def n(t):
        x = d.get(t)
        return 0 if x is None else len(x)
    s = [proj.get("tam_clientes") or 0, n("concorrentes"), n("swot"), _tem_txt(proj.get("pm_situacao"))]
    sit = sum(bool(x) for x in s)
    obm = objetivos_marketing(d)
    est = [_tem_txt(proj.get("posic_segmento")), _tem_txt(proj.get("pm_ansoff")), _tem_txt(proj.get("pm_generica"))]
    tat = [_tem_txt(proj.get("mix_nucleo")), _tem_txt(proj.get("mix_estrategia_preco")), n("canais_venda") > 0,
           _tem_txt(proj.get("mix_objetivo_com"))]
    acoes = n("calendario_mkt")
    kp = n("kpis_mkt")
    return [
        ("Situação", 2 if sit >= 4 else (1 if sit else 0), "mercado, concorrentes, SWOT e a síntese do diagnóstico"),
        ("Objetivos", 2 if len(obm) >= 2 else (1 if len(obm) else 0), "ao menos dois objetivos da perspectiva Clientes"),
        ("Estratégia", 2 if all(est) else (1 if any(est) else 0), "público-alvo, direção de crescimento e estratégia competitiva"),
        ("Táticas", 2 if all(tat) else (1 if any(tat) else 0), "produto, estratégia de preço, canais de venda e objetivo da comunicação"),
        ("Ações", 2 if acoes >= 4 else (1 if acoes else 0), "ao menos quatro ações no calendário de marketing"),
        ("Controle", 2 if kp >= 3 else (1 if kp else 0), "ao menos três indicadores com meta e frequência"),
    ]


def regras_mix(proj: dict, d: dict) -> dict:
    et = etapas_plano_mkt(proj, d)
    completas = sum(1 for _, s, _ in et if s == 2)
    saz = d.get("sazonalidade")
    tem_saz = saz is not None and len(saz) >= 12 and float(saz["indice"].max() - saz["indice"].min()) > 0
    return {"plano_mkt": (completas >= 5, any(s for _, s, _ in et)),
            "sazonalidade": (tem_saz, saz is not None and len(saz) > 0),
            "tornado": (len(d.get("mc_faixas", [])) > 0 and (len(d["catalogo"]) > 0 or len(d["assinaturas"]) > 0),
                        len(d["catalogo"]) > 0 or len(d["assinaturas"]) > 0),
            "taxas": ("bcb_selic" in set(d["dados_publicos"]["chave"]), "bcb_selic" in set(d["dados_publicos"]["chave"])),
            "mapa": (len(d.get("mapa_pontos", [])) > 0 and float(proj.get("loc_lat") or 0) != 0, len(d.get("mapa_pontos", [])) > 0),
            "arvore": (len(d["catalogo"]) > 0 or len(d["assinaturas"]) > 0, len(d["catalogo"]) > 0 or len(d["assinaturas"]) > 0)}


def frase_posicionamento(proj: dict) -> str:
    p = {k: str(proj.get(k) or "").strip() for k in ("posic_segmento", "posic_necessidade", "posic_marca", "posic_categoria",
                                                     "posic_beneficio", "posic_concorrente", "posic_diferencial")}
    if not (p["posic_segmento"] and p["posic_marca"]):
        return ""
    return (f"Para {p['posic_segmento']}, que {p['posic_necessidade'] or '...'}, {p['posic_marca']} é {p['posic_categoria'] or '...'} "
            f"que {p['posic_beneficio'] or '...'}. Diferentemente de {p['posic_concorrente'] or '...'}, "
            f"{p['posic_diferencial'] or '...'}.")


def resumo_mix_linhas(proj: dict, d: dict, res: dict) -> list[tuple[str, str]]:
    """Uma linha por P, montada a partir das escolhas estruturadas (para o plano e os relatórios)."""
    cv = d.get("canais_venda")
    canais = ", ".join(f"{c} ({_n(p)}%)" for c, p in zip(cv["canal"], cv["pct_vendas"])) if cv is not None and not cv.empty else ""
    mm = mc_media(res)
    linhas = [
        ("Produto", "; ".join(x for x in [proj.get("mix_nucleo"), proj.get("mix_ampliado")] if _tem_txt(x))
         + (f". Fase do ciclo de vida: {proj['mix_ciclo'].lower()}." if _tem_txt(proj.get("mix_ciclo")) else "")),
        ("Preço", ", ".join(x for x in [f"estratégia de {proj['mix_estrategia_preco'].lower()}" if _tem_txt(proj.get("mix_estrategia_preco")) else "",
                                         f"posição {proj['mix_posicao_preco'].lower()}" if _tem_txt(proj.get("mix_posicao_preco")) else "",
                                         f"margem de contribuição média de {_p(mm)}" if mm == mm else ""] if x)),
        ("Praça", ", ".join(x for x in [f"distribuição {proj['mix_distribuicao'].lower()}" if _tem_txt(proj.get("mix_distribuicao")) else "",
                                         f"canais: {canais}" if canais else ""] if x)),
        ("Promoção", proj.get("mix_objetivo_com") or ""),
    ]
    for rot, campo in (("Pessoas", "mix_pessoas"), ("Processos", "mix_processos"), ("Evidência física", "mix_evidencia")):
        if _tem_txt(proj.get(campo)):
            linhas.append((rot, proj[campo]))
    texto = {"Produto": "mkt_produto", "Preço": "mkt_preco", "Praça": "mkt_praca", "Promoção": "mkt_promocao"}
    def cap(x):
        x = str(x).strip()
        return x[:1].upper() + x[1:]
    return [(r, cap(t if t else str(proj.get(texto.get(r, ""), "") or ""))) for r, t in linhas
            if t or _tem_txt(proj.get(texto.get(r, "")))]


def conteudo_plano_mkt(proj: dict, d: dict, res: dict) -> list[dict]:
    """Seções do plano de marketing para a tela e os relatórios: titulo, texto, tabela (DataFrame ou None), refs."""
    out = []
    # 1 Situação
    t = []
    if proj.get("tam_clientes"):
        t.append(f"O mercado atendível foi estimado em {_n(float(proj.get('sam_clientes') or 0))} clientes, com meta de "
                 f"participação de {_n(float(proj.get('som_pct') or 0), 1)}%.")
    nc = len(d.get("concorrentes", []))
    if nc:
        comp = comparar_concorrentes(res.get("cat"), d.get("precos_concorrentes"))
        idx = indice_preco_medio(res.get("cat"), comp) if not comp.empty else float("nan")
        t.append(f"Foram analisados {nc} concorrentes" + (f", cujos preços ficam, em média, {_p(abs(idx - 1), 0)} "
                 f"{'abaixo' if idx >= 1 else 'acima'} dos do projeto." if idx == idx else "."))
    if _tem_txt(proj.get("pm_situacao")):
        t.append(str(proj["pm_situacao"]))
    out.append({"titulo": "Análise da situação", "texto": " ".join(t) or "Diagnóstico a elaborar.", "tabela": None, "refs": ["smith"]})
    # 2 Objetivos
    obm = objetivos_marketing(d)
    tab = None
    if not obm.empty:
        tab = pd.DataFrame({"Objetivo": obm["objetivo"], "Indicador": obm["indicador"],
                            "Meta": [f"{_n(m, 0)} {u}".strip() for m, u in zip(obm["meta"], obm["unidade"])],
                            "Prazo (mês)": obm["prazo_meses"].map(lambda v: _n(v))})
    out.append({"titulo": "Objetivos de marketing", "texto": "Os objetivos seguem os critérios SMART: específicos, mensuráveis, "
                "atingíveis, relevantes e com prazo." if tab is not None else "Objetivos a definir.", "tabela": tab, "refs": []})
    # 3 Estratégia
    t = []
    fp = frase_posicionamento(proj)
    if fp:
        t.append("Declaração de posicionamento: " + fp)
    if _tem_txt(proj.get("pm_ansoff")):
        t.append(f"Direção de crescimento (matriz de Ansoff): {proj['pm_ansoff'].lower()}, isto é, "
                 f"{ANSOFF.get(proj['pm_ansoff'], ('', ''))[0].lower()}.")
    if _tem_txt(proj.get("pm_generica")):
        t.append(f"Estratégia competitiva: {proj['pm_generica'].lower()}. {GENERICAS.get(proj['pm_generica'], '')}")
    out.append({"titulo": "Estratégia de marketing", "texto": " ".join(t) or "Estratégia a definir.", "tabela": None,
                "refs": ["ansoff"] if _tem_txt(proj.get("pm_ansoff")) else []})
    # 4 Táticas
    lm = resumo_mix_linhas(proj, d, res)
    refs = ["kotler"] + (["booms"] if any(r in ("Pessoas", "Processos", "Evidência física") for r, _ in lm) else [])
    out.append({"titulo": "Táticas: mix de marketing", "texto": "O mix de marketing traduz a estratégia em decisões sobre cada P.",
                "tabela": pd.DataFrame(lm, columns=["Elemento", "Decisão"]) if lm else None, "refs": refs})
    # 5 Ações
    cal = d.get("calendario_mkt")
    tab = None
    if cal is not None and not cal.empty:
        c = cal.sort_values("mes")
        tab = pd.DataFrame({"Mês": c["mes"].map(lambda v: _n(v)), "Ação": c["acao"], "Canal": c["canal"],
                            "Custo (R$)": c["custo"].map(lambda v: _n(v, 2))})
    out.append({"titulo": "Plano de ação", "texto": (f"O calendário prevê {len(cal)} ações, com custo total de R$ "
                f"{_n(float(cal['custo'].sum()), 2)}." if tab is not None else "Ações a programar."), "tabela": tab, "refs": []})
    # 6 Controle
    kp = d.get("kpis_mkt")
    tab = None
    if kp is not None and not kp.empty:
        tab = pd.DataFrame({"Indicador": kp["indicador"], "Meta": kp["meta"], "Frequência": kp["frequencia"],
                            "Fonte do dado": kp["fonte"], "Responsável": kp["responsavel"]})
    out.append({"titulo": "Controle", "texto": "Os indicadores abaixo serão acompanhados e comparados com as metas, e os desvios "
                "levam à revisão das táticas, fechando o ciclo de planejamento." if tab is not None else "Indicadores a definir.",
                "tabela": tab, "refs": ["mcdonald"]})
    return out


REFS_MKT = {
    "smith": "SMITH, P. R. **The SOSTAC guide**: to writing the perfect plan. [S. l.]: PR Smith, 2011.",
    "mcdonald": "McDONALD, Malcolm. **Marketing plans**: how to prepare them, how to use them. 6. ed. Oxford: Butterworth-Heinemann, 2007.",
    "kotler": "KOTLER, Philip; KELLER, Kevin Lane. **Administração de marketing**. 14. ed. São Paulo: Pearson Education do Brasil, 2012.",
    "booms": "BOOMS, Bernard H.; BITNER, Mary Jo. Marketing strategies and organization structures for service firms. In: DONNELLY, "
             "James H.; GEORGE, William R. (ed.). **Marketing of services**. Chicago: American Marketing Association, 1981. p. 47-51.",
    "ansoff": "ANSOFF, H. Igor. Strategies for diversification. **Harvard Business Review**, v. 35, n. 5, p. 113-124, 1957.",
    "lascasas": "LAS CASAS, Alexandre Luzzi. **Plano de marketing para micro e pequena empresa**. 6. ed. São Paulo: Atlas, 2011.",
}
CITA_MKT = {"smith": "Smith, 2011", "mcdonald": "McDonald, 2007", "kotler": "Kotler; Keller, 2012",
            "booms": "Booms; Bitner, 1981", "ansoff": "Ansoff, 1957", "lascasas": "Las Casas, 2011"}
