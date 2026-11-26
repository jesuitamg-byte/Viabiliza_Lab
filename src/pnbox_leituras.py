"""
Leitura automática dos resultados (v8.5): ao final de cada tabela, duas a quatro frases
que dizem o que os números mostram, em linguagem simples. As regras são fixas e
transparentes (não usam IA), para que o resultado seja o mesmo em qualquer computador
e o professor saiba de onde veio cada frase.

Cada função recebe a tabela (df), o projeto (proj), as tabelas (d) e uma função R()
que devolve o resultado do motor financeiro só quando a regra precisar dele.
"""
import pandas as pd

from pnbox_mix import (_n, _p, comparar_concorrentes, indice_preco_medio, mc_media,
                       resumo_blueprint, resumo_canais, resumo_pagamento)


def _q(n, singular: str, plural: str) -> str:
    """Quantidade com a palavra no singular ou no plural: _q(1, "item", "itens") -> "1 item"."""
    n = float(n)
    return f"{_n(n)} {singular if abs(n) == 1 else plural}"


PLURAL_SWOT = {"Força": "forças", "Fraqueza": "fraquezas", "Oportunidade": "oportunidades", "Ameaça": "ameaças"}


def _brl(x: float) -> str:
    return "R$ " + _n(float(x), 2)


def _validas(df: pd.DataFrame, col: str) -> pd.DataFrame:
    return df[df[col].astype(str).str.strip() != ""] if col in df else df


def _top(df, col_valor, col_nome):
    if df.empty or df[col_valor].sum() <= 0:
        return None, 0.0
    i = df[col_valor].astype(float).idxmax()
    return str(df.loc[i, col_nome]), float(df.loc[i, col_valor]) / float(df[col_valor].sum())


# ------------------------------------------------------------------------------
# FINANÇAS
# ------------------------------------------------------------------------------
def _l_catalogo(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    res = R()
    cat = res["cat"]
    out = [f"{_q(len(df), 'item soma', 'itens somam')} receita de {_brl(res['receita_m1'])} no primeiro mês."]
    nome, share = _top(cat, "receita_mensal", "descricao")
    if nome:
        out.append(f"O item que mais fatura é {nome}, com {_p(share, 0)} da receita."
                   + (" O plano depende muito dele." if share >= 0.5 and len(cat) > 2 else ""))
    neg = cat.loc[cat["margem_unit"] <= 0, "descricao"].tolist()
    if neg:
        out.append("Vendidos sem margem (preço menor ou igual ao custo): " + ", ".join(neg) + ".")
    mm = mc_media(res)
    if mm == mm:
        out.append(f"Depois de impostos e custos variáveis, sobra em média {_p(mm)} de cada venda para pagar os custos fixos.")
    return out


def _l_custos_fixos(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["valor_mensal"].sum())
    out = [f"Os custos fixos (sem pessoal) somam {_brl(tot)} por mês, {_brl(tot * 12)} por ano."]
    nome, share = _top(df, "valor_mensal", "descricao")
    if nome:
        out.append(f"O maior é {nome}, com {_p(share, 0)} do total.")
    res = R()
    if res["receita_m1"] > 0:
        r = (tot + res["pessoal_m"]) / res["receita_m1"]
        out.append(f"Somados à equipe, os custos fixos equivalem a {_p(r, 0)} da receita do primeiro mês"
                   + (": o negócio começa sem cobrir a estrutura e precisa crescer para empatar." if r > 1 else "."))
    return out


def _l_custos_variaveis(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["percentual"].sum())
    out = [f"Os custos variáveis sobre a venda somam {_n(tot, 2)}%: de cada R$ 100 vendidos, R$ {_n(tot, 2)} saem para eles,"
           " além dos impostos."]
    nome, share = _top(df, "percentual", "descricao")
    if nome:
        out.append(f"O mais pesado é {nome}.")
    return out


def _l_pessoal(df, proj, d, R):
    df = _validas(df, "cargo")
    if df.empty:
        return []
    res = R()
    n = float(df["num_empregados"].sum())
    out = [f"{_q(n, 'pessoa', 'pessoas')} na equipe, com custo de {_brl(res['pessoal_m'])} por mês no primeiro mês, contando encargos e benefícios."]
    fut = df[df["mes_inicio"] > 1]
    if not fut.empty:
        out.append(f"{_q(len(fut), 'contratação prevista', 'contratações previstas')} para depois da abertura, o que aumenta o custo ao longo do plano.")
    if res["receita_m1"] > 0:
        out.append(f"A equipe consome {_p(res['pessoal_m'] / res['receita_m1'], 0)} da receita do primeiro mês.")
    return out


def _l_inv_fixos(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float((df["quantidade"] * df["valor_unitario"]).sum())
    res = R()
    out = [f"Os investimentos fixos somam {_brl(tot)}, {_p(tot / res['investimento_total'], 0) if res['investimento_total'] else '-'} "
           "do investimento total."]
    tmp = df.assign(total=df["quantidade"] * df["valor_unitario"])
    nome, share = _top(tmp, "total", "descricao")
    if nome:
        out.append(f"O item mais caro é {nome} ({_p(share, 0)}).")
    out.append(f"A depreciação, que entra como custo sem sair do caixa, é de {_brl(res['dep_anual'])} por ano.")
    return out


def _l_inv_pre(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["valor"].sum())
    return [f"Os gastos antes de abrir somam {_brl(tot)}. São gastos que não voltam se o negócio não sair do papel, "
            "por isso convém conferir orçamentos reais de cada um."]


def _l_emprestimos(df, proj, d, R):
    df = _validas(df, "descricao")
    df = df[df["valor"] > 0]
    if df.empty:
        return []
    tot = float(df["valor"].sum())
    taxa = float((df["valor"] * df["taxa_anual_pct"]).sum() / tot)
    res = R()
    out = [f"{_q(len(df), 'empréstimo soma', 'empréstimos somam')} {_brl(tot)}, com taxa média de {_n(taxa, 1)}% ao ano."]
    if res["fontes"] > 0:
        out.append(f"O dinheiro de terceiros cobre {_p(tot / res['fontes'], 0)} das fontes de recursos.")
    if taxa / 100 > res["tma"] and res["tma"] > 0:
        out.append("A taxa do empréstimo é maior que a TMA do projeto: o crédito custa mais do que o retorno mínimo exigido.")
    return out


def _l_socios(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    tot = float(df["capital_integralizado"].sum())
    out = [f"{_q(len(df), 'sócio investe', 'sócios investem')} {_brl(tot)} de capital próprio."]
    if tot > 0 and len(df) > 1:
        nome, share = _top(df, "capital_integralizado", "nome")
        out.append(f"{nome} entra com {_p(share, 0)} do capital" + (", o que garante o controle da sociedade." if share > 0.5 else "."))
    return out


def _l_prazos(df, proj, d, R):
    res = R()
    out = [f"O cliente paga, em média, em {_n(res['pmr'])} dias e a empresa paga fornecedores em {_n(res['pmp'])} dias; "
           f"o estoque gira em {_n(res['pme'])} dias."]
    cf = res["ciclo_fin"]
    out.append((f"O ciclo financeiro é de {_n(cf)} dias" if cf >= 0 else f"O ciclo financeiro é negativo ({_n(cf)} dias)") + (": a empresa financia a operação nesse período, e é isso que pede capital de giro."
                                                         if cf > 0 else ": os fornecedores financiam a operação."))
    return out


def _l_materiais(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    est = float((df["qtd_estoque_inicial"] * df["valor_unitario"]).sum())
    return [f"{_q(len(df), 'material cadastrado', 'materiais cadastrados')}; o estoque inicial deles custa {_brl(est)}."]


def _l_assinaturas(df, proj, d, R):
    df = _validas(df, "plano")
    if df.empty:
        return []
    out = []
    for _, r in df.iterrows():
        ch = float(r["churn_mensal_pct"] or 0)
        if ch > 0:
            out.append(f"No plano {r['plano']}, com cancelamento de {_n(ch, 1)}% ao mês, o assinante fica em média {_n(100 / ch, 0)} meses.")
    return out[:3]


def _l_realizado(df, proj, d, R):
    df = df[df["mes"] > 0]
    if df.empty:
        return []
    res = R()
    plan = res["meses"]["Receita"].tolist()
    reais = [(int(m), float(r)) for m, r in zip(df["mes"], df["receita_real"]) if 1 <= m <= 12]
    if not reais:
        return []
    p = sum(plan[m - 1] for m, _ in reais)
    r = sum(v for _, v in reais)
    return [f"Em {_q(len(reais), 'mês lançado', 'meses lançados')}, a receita real foi {_brl(r)}, contra {_brl(p)} planejados "
            f"({_p(r / p - 1, 0) if p else '-'} de diferença).",
            "Diferença acima de 10% para baixo pede revisão do volume ou do preço no plano." if p and r / p < 0.9 else
            "O realizado está próximo ou acima do planejado."]


# ------------------------------------------------------------------------------
# MERCADO, ESTRATÉGIA E MARKETING
# ------------------------------------------------------------------------------
def _l_segmentos(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    notas = ["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]
    m = df.assign(media=df[notas].mean(axis=1)).sort_values("media", ascending=False)
    out = [f"O segmento mais atraente pelas notas é {m['nome'].iloc[0]} (média {_n(m['media'].iloc[0], 1)})."]
    if len(m) > 1 and m["media"].iloc[0] - m["media"].iloc[1] < 0.5:
        out.append(f"A diferença para {m['nome'].iloc[1]} é pequena: vale olhar qual dos dois o grupo consegue atender melhor.")
    return out


def _l_swot(df, proj, d, R):
    df = _validas(df, "fator")
    if df.empty:
        return []
    c = df["natureza"].value_counts()
    out = ["Fatores: " + ", ".join(_q(v, k.lower(), PLURAL_SWOT.get(k, k.lower())) for k, v in c.items()) + "."]
    if c.get("Ameaça", 0) > c.get("Oportunidade", 0):
        out.append("Há mais ameaças que oportunidades no ambiente: as estratégias defensivas da SWOT cruzada ganham peso.")
    if c.get("Fraqueza", 0) > c.get("Força", 0):
        out.append("As fraquezas superam as forças: o plano precisa mostrar como cada uma será tratada.")
    return out


def _l_porter(df, proj, d, R):
    df = _validas(df, "forca")
    if df.empty:
        return []
    m = float(df["intensidade"].mean())
    i = df["intensidade"].idxmax()
    return [f"Intensidade média das forças: {_n(m, 1)}. A mais forte é {df.loc[i, 'forca'].lower()}, "
            "que é onde a estratégia precisa responder primeiro."]


def _l_riscos(df, proj, d, R):
    df = _validas(df, "risco")
    if df.empty:
        return []
    sev = df["probabilidade"] * df["impacto"]
    alto = df[sev >= 15]
    i = sev.idxmax()
    out = [f"O risco mais grave é {df.loc[i, 'risco']} (probabilidade × impacto = {_n(sev[i])})."]
    if not alto.empty:
        out.append(f"{_q(len(alto), 'risco está', 'riscos estão')} na zona alta (15 ou mais) e pedem plano de contingência antes da abertura.")
    sem = df[df["mitigacao"].astype(str).str.strip() == ""]
    if not sem.empty:
        out.append(f"{_q(len(sem), 'risco', 'riscos')} ainda sem mitigação.")
    return out


def _l_precos_vw(df, proj, d, R):
    df = _validas(df, "respondente")
    n = len(df)
    if not n:
        return []
    return [f"{_q(n, 'resposta', 'respostas')}. " + ("Com menos de 5 respostas coerentes a faixa de preço não é calculada; o recomendado é passar de 30."
                                    if n < 5 else "A faixa aceitável e o preço ótimo aparecem no gráfico; quanto mais respostas, mais confiável."
                                    if n < 30 else "Amostra razoável para uma primeira faixa de preço.")]


def _l_entrevistas(df, proj, d, R):
    df = _validas(df, "entrevistado")
    if df.empty:
        return []
    sim = float((df["pagaria"] == "Sim").mean())
    out = [f"{_q(len(df), 'entrevista', 'entrevistas')}; {_p(sim, 0)} disseram que pagariam."]
    if len(df) < 10:
        out.append("Com menos de 10 conversas é cedo para conclusões: as respostas ainda podem mudar muito.")
    if sim > 0.8:
        out.append("Intenção muito alta costuma ser otimista: quem diz \"pagaria\" nem sempre compra.")
    return out


def _l_canais(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    m = df.assign(media=df.iloc[:, 1:6].mean(axis=1)).sort_values("media", ascending=False)
    return [f"Os canais com maior nota média são {', '.join(m['canal'].head(3))}. Concentrar a verba neles costuma render mais "
            "do que espalhar em muitos canais."]


def _l_funil(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    topo = float(df.loc[df["etapa"] == "Topo do Funil", "alcancados"].sum())
    fundo = float(df.loc[df["etapa"] == "Fundo do Funil", "convertidos"].sum())
    if topo <= 0:
        return []
    return [f"De {_n(topo)} pessoas alcançadas no topo, {_n(fundo)} compram: conversão total de {_p(fundo / topo, 1)}."]


def _l_orcamento_canais(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    g, c = float(df["gasto_mensal"].sum()), float(df["clientes_mes"].sum())
    out = [f"Orçamento de {_brl(g)} por mês para trazer {_n(c)} clientes novos: custo médio de {_brl(g / c) if c else '-'} por cliente."]
    x = df[df["clientes_mes"] > 0].assign(cac=lambda t: t["gasto_mensal"] / t["clientes_mes"]).sort_values("cac")
    if len(x) > 1:
        out.append(f"O canal mais barato por cliente é {x['canal'].iloc[0]}; o mais caro, {x['canal'].iloc[-1]}.")
    return out


def _l_fornecedores(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    crit = df[df["critico"].astype(bool)]
    out = [f"{_q(len(df), 'fornecedor', 'fornecedores')}, prazo médio de pagamento de {_n(float(df['prazo_pagamento_dias'].mean()))} dias."]
    if not crit.empty:
        out.append(f"{_q(len(crit), 'fornecedor está marcado como crítico', 'fornecedores estão marcados como críticos')}: vale ter um segundo fornecedor para cada um.")
    return out


def _l_cronograma(df, proj, d, R):
    df = _validas(df, "atividade")
    if df.empty:
        return []
    fim = float((df["mes_inicio"] + df["duracao_meses"] - 1).max())
    out = [f"{_q(len(df), 'atividade', 'atividades')}, terminando por volta do mês {_n(fim)}, com custo previsto de {_brl(df['custo'].sum())}."]
    sem = df[df["responsavel"].astype(str).str.strip() == ""]
    if not sem.empty:
        out.append(f"{_q(len(sem), 'atividade', 'atividades')} sem responsável.")
    return out


def _l_calendario(df, proj, d, R):
    df = _validas(df, "acao")
    if df.empty:
        return []
    meses = df["mes"].nunique()
    return [f"{_q(len(df), 'ação', 'ações')} em {_q(meses, 'mês', 'meses diferentes')}, custo total de {_brl(df['custo'].sum())}."
            + (" As ações estão concentradas em poucos meses; depois do lançamento a comunicação some." if meses <= 2 and len(df) >= 3 else "")]


def _l_bcg(df, proj, d, R):
    df = _validas(df, "produto")
    if df.empty:
        return []
    lim = float(proj.get("bcg_limite_crescimento_pct") or 10)
    q = {"Estrela": 0, "Vaca leiteira": 0, "Interrogação": 0, "Abacaxi": 0}
    for g, p in zip(df["crescimento_mercado_pct"], df["participacao_relativa"]):
        q["Estrela" if g >= lim and p >= 1 else "Interrogação" if g >= lim else "Vaca leiteira" if p >= 1 else "Abacaxi"] += 1
    out = ["; ".join(f"{k}: {v}" for k, v in q.items() if v) + "."]
    if q["Vaca leiteira"] == 0:
        out.append("Sem vaca leiteira, não há item que gere caixa folgado para financiar os outros.")
    return out


def _l_errc(df, proj, d, R):
    df = _validas(df, "atributo")
    if df.empty:
        return []
    c = df["acao"].value_counts()
    return [", ".join(f"{k}: {v}" for k, v in c.items()) + ". A nova curva de valor só se diferencia se houver algo em criar ou elevar."]


def _l_experimentos(df, proj, d, R):
    df = _validas(df, "hipotese")
    if df.empty:
        return []
    c = df["status"].value_counts()
    return [", ".join(f"{k or 'sem status'}: {v}" for k, v in c.items()) + "."]


def _l_formalizacao(df, proj, d, R):
    df = _validas(df, "item")
    if df.empty:
        return []
    c = df["status"].value_counts()
    return [", ".join(f"{k or 'sem status'}: {v}" for k, v in c.items()) + "."]


def _l_aarrr(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    sem = df[(df["metrica"].astype(str).str.strip() == "") | (df["meta"].astype(str).str.strip() == "")]
    return [f"{len(df) - len(sem)} de {len(df)} etapas têm métrica e meta." + (" Etapa sem meta não pode ser acompanhada." if len(sem) else "")]


def _l_objetivos(df, proj, d, R):
    df = _validas(df, "objetivo")
    if df.empty:
        return []
    c = df["perspectiva"].replace("", "sem perspectiva").value_counts()
    out = ["Objetivos por perspectiva: " + ", ".join(f"{k} {v}" for k, v in c.items()) + "."]
    falta = [p for p in ("Financeira", "Clientes", "Processos internos", "Aprendizado e crescimento") if p not in c]
    if falta:
        out.append("Sem objetivo em: " + ", ".join(falta).lower() + ". O mapa estratégico fica incompleto.")
    return out


def _l_precos_concorrentes(df, proj, d, R):
    cat = R()["cat"]
    comp = comparar_concorrentes(cat, df)
    if comp.empty:
        return []
    idx = indice_preco_medio(cat, comp)
    return [f"Considerando {_q(len(comp), 'produto comparado', 'produtos comparados')}, seus preços estão {_p(abs(idx - 1), 0)} "
            f"{'acima' if idx >= 1 else 'abaixo'} da média dos concorrentes."]


def _l_meios_pagamento(df, proj, d, R):
    df = _validas(df, "forma")
    if df.empty:
        return []
    r = resumo_pagamento(df)
    return [f"As taxas de pagamento consomem {_n(r['taxa'], 2)}% da receita e o dinheiro entra, em média, {_n(r['prazo'])} dias "
            "depois da venda. Quanto maior esse prazo, mais capital de giro o negócio precisa."]


def _l_canais_venda(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    r = resumo_canais(df)
    return [f"{r['maior']} responde por {_n(r['maior_pct'])}% das vendas; a comissão média dos canais é de {_n(r['comissao'], 2)}% da receita."]


def _l_blueprint(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    b = resumo_blueprint(df)
    return [f"O serviço tem {b['etapas']} etapas e leva {_n(b['tempo'])} minutos, dos quais {_n(b['visivel'])} na frente do cliente.",
            f"{_q(len(b['falhas']), 'ponto de falha identificado', 'pontos de falha identificados')}." if b["falhas"] else "Nenhum ponto de falha anotado; vale perguntar onde o cliente costuma reclamar."]


def _l_kpis(df, proj, d, R):
    df = _validas(df, "indicador")
    if df.empty:
        return []
    sem = int((df["meta"].astype(str).str.strip() == "").sum())
    return [f"{_q(len(df), 'indicador', 'indicadores')} de controle" + (f", {sem} sem meta." if sem else ", todos com meta.")]


def _l_swot_estrategias(df, proj, d, R):
    df = _validas(df, "estrategia")
    if df.empty:
        return []
    q = df["quadrante"].str.split(" ").str[0].value_counts()
    return ["Estratégias por tipo: " + ", ".join(f"{k} {v}" for k, v in q.items()) + "."]


def _l_capacidade(df, proj, d, R):
    df = _validas(df, "recurso")
    if df.empty:
        return []
    cat = R()["cat"]
    out = []
    for _, r in df.iterrows():
        q = cat.loc[cat["descricao"] == r["produto_ou_grupo"], "qtd_mensal"].sum() if not cat.empty else 0
        if q and r["capacidade_mensal"]:
            u = q / r["capacidade_mensal"]
            out.append(f"{r['recurso']} trabalha a {_p(u, 0)} da capacidade no primeiro mês"
                       + (": acima de 85% sobra pouca folga para crescer." if u > 0.85 else "."))
    return out[:3] or [f"{_q(len(df), 'recurso de capacidade cadastrado', 'recursos de capacidade cadastrados')}."]


LEITURAS = {
    "catalogo": _l_catalogo, "custos_fixos": _l_custos_fixos, "custos_variaveis": _l_custos_variaveis, "pessoal": _l_pessoal,
    "inv_fixos": _l_inv_fixos, "inv_pre": _l_inv_pre, "emprestimos": _l_emprestimos, "socios": _l_socios,
    "capital_giro_prazos": _l_prazos, "materiais": _l_materiais, "assinaturas": _l_assinaturas, "realizado": _l_realizado,
    "segmentos": _l_segmentos, "swot": _l_swot, "porter": _l_porter, "riscos": _l_riscos, "precos_vw": _l_precos_vw,
    "entrevistas": _l_entrevistas, "canais": _l_canais, "funil_vendas": _l_funil, "orcamento_canais": _l_orcamento_canais,
    "fornecedores": _l_fornecedores, "cronograma": _l_cronograma, "calendario_mkt": _l_calendario, "bcg": _l_bcg,
    "errc": _l_errc, "experimentos": _l_experimentos, "formalizacao": _l_formalizacao, "aarrr": _l_aarrr,
    "objetivos": _l_objetivos, "precos_concorrentes": _l_precos_concorrentes, "meios_pagamento": _l_meios_pagamento,
    "canais_venda": _l_canais_venda, "blueprint": _l_blueprint, "kpis_mkt": _l_kpis, "swot_estrategias": _l_swot_estrategias,
    "capacidade": _l_capacidade,
}


# ------------------------------------------------------------------------------
# TABELAS DE RESULTADO (DRE, fluxo de caixa)
# ------------------------------------------------------------------------------
def leitura_dre(res: dict) -> list[str]:
    a = res["anos"]
    N = res["N"]
    a1, an = a[1], a[N]
    out = [f"No Ano 1 a receita é de {_brl(a1['rec_bruta'])} e o resultado líquido de {_brl(a1['resultado'])} "
           f"({_p(a1['resultado'] / a1['rec_bruta'], 1) if a1['rec_bruta'] else '-'} da receita)."]
    if N > 1:
        out.append(f"No Ano {N} o resultado chega a {_brl(an['resultado'])}"
                   + (", saindo do prejuízo." if a1["resultado"] < 0 <= an["resultado"] else "."))
    prim = next((k for k in res["ANOS"] if a[k]["resultado"] > 0), None)
    if prim is None:
        out.append("O resultado não fica positivo em nenhum ano do horizonte: o plano, como está, não se paga.")
    m = res["meses"]
    pos = m.index[m["Resultado operacional"] > 0]
    if len(pos):
        out.append("O resultado operacional já é positivo no primeiro mês." if pos[0] == m.index[0] else
                   f"Mês a mês, o resultado operacional fica positivo a partir do {m.loc[pos[0], 'Mês'].lower()}.")
    return out


def leitura_fluxo(res: dict) -> list[str]:
    fm = res["fluxo_mensal"]
    if fm.empty:
        return []
    i = fm["Saldo final"].idxmin()
    mn, mes = float(fm.loc[i, "Saldo final"]), fm.loc[i, "Mês"]
    out = [f"O caixa chega ao ponto mais baixo no {str(mes).lower()}, com {_brl(mn)}."]
    out.append("Como fica negativo, falta dinheiro nesse período: é preciso mais capital de giro, prazo maior de fornecedores "
               "ou recebimento mais rápido." if mn < 0 else "Como não fica negativo, o capital de giro previsto cobre o período mais apertado.")
    return out
