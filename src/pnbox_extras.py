"""
================================================================================
VIABILIZALAB - RECURSOS EXTRAS (pnbox_extras.py)
================================================================================
Sumário executivo automático, texto para análise com IA, pitch deck em PDF,
exportação em Word, situação de preenchimento e avaliação (modo professor).
================================================================================
"""
import io
import re
from datetime import date
from pathlib import Path

import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos

from pnbox_motor import (brl, num_br, pct, div, QUADRANTES, pontuacao_swot, mercado, semaforo,
                         tabela_dre, tabela_indicadores, composicoes, sensibilidade, capacidade_demanda,
                         tabela_checagens_v5, raio_x, executar_motor, custo_hora, cenarios)
from pnbox_db import SCHEMA, conectar, ler, ler_unica
from pnbox_estrategia import (checar_smart, PERSPECTIVAS, canvas_blocos, bcg_tabela, ge_tabela, porter_tabela,
from pnbox_validacao import *  # noqa
from pnbox_v8 import *  # noqa
from pnbox_mix import *  # noqa
                              porter_leitura, bsc_indicadores_plano, simular_reforma, frase_smart)
from pnbox_pdf import Relatorio, graficos_png, AZUL, FONTES_CANDIDATAS


# ------------------------------------------------------------------------------
# 1. SITUAÇÃO DE PREENCHIMENTO DE CADA FERRAMENTA
# ------------------------------------------------------------------------------
def regras_situacao(proj: dict, d: dict) -> dict:
    def tem(t, n=1):
        return len(d[t]) >= n
    sw = d["swot"]
    rec = tem("catalogo") or tem("assinaturas")
    return {
        "projeto": (bool(proj["nome_negocio"]) and bool(proj["cidade_uf"]), tem("socios")),
        "segmentos": (tem("segmentos", 2), tem("segmentos")),
        "personas": (tem("personas") and bool(d["personas"]["desafios"].astype(str).str.strip().ne("").any()), tem("personas")),
        "jornada": (d["jornada_cliente"]["etapa"].nunique() >= 6, tem("jornada_cliente")),
        "mercado": (proj["tam_clientes"] > 0 and proj["sam_clientes"] > 0 and proj["som_pct"] > 0 and proj["gasto_anual_cliente"] > 0,
                    proj["tam_clientes"] > 0 or proj["gasto_anual_cliente"] > 0),
        "proposta": (tem("propostas"), bool(d["proposta_valor"]["segmento_alvo"])),
        "concorrencia": (tem("concorrentes") and tem("concorrencia_notas", 3), tem("concorrencia_notas")),
        "forcas": ((sw["natureza"] == "Força").any() and (sw["natureza"] == "Fraqueza").any(),
                   sw["natureza"].isin(["Força", "Fraqueza"]).any()),
        "oportunidades": ((sw["natureza"] == "Oportunidade").any() and (sw["natureza"] == "Ameaça").any(),
                          sw["natureza"].isin(["Oportunidade", "Ameaça"]).any()),
        "swot": (sw["natureza"].nunique() >= 4, len(sw) > 0),
        "marketing": (all(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")),
                      any(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao"))),
        "canais": (tem("canais") and tem("funil_vendas"), tem("canais") or tem("funil_vendas")),
        "operacional": (bool(proj["loc_endereco"]) and tem("capacidade"), bool(proj["loc_endereco"]) or tem("capacidade")),
        "fornecedores": (tem("fornecedores", 2), tem("fornecedores")),
        "investimentos": (tem("inv_fixos") and (tem("socios") or tem("emprestimos")), tem("inv_fixos")),
        "produtos": (tem("catalogo") and tem("materiais"), tem("catalogo") or tem("materiais")),
        "custos": (tem("pessoal") and tem("custos_fixos"), tem("pessoal") or tem("custos_fixos")),
        "identidade": (bool(proj["missao"]) and bool(proj["visao"]) and bool(proj["valores"]) and tem("objetivos", 3) and
                       all(all(ok for ok, _ in checar_smart(o).values()) for _, o in d["objetivos"].iterrows()),
                       bool(proj["missao"]) or bool(proj["visao"]) or tem("objetivos")),
        "canvas": (bool(proj["canvas_relacionamento"]) and bool(proj["canvas_parcerias"]), tem("segmentos") or tem("propostas")),
        "porter": (len(d["porter"]) >= 5 and bool((d["porter"]["intensidade"] > 0).all()) and
                   not d["porter"]["justificativa"].astype(str).str.startswith("(sugestão)").any(), tem("porter")),
        "ge": (tem("segmentos", 2) and tem("concorrencia_notas"), tem("segmentos")),
        "bcg": (tem("bcg", 2), tem("bcg")),
        "bsc": (set(PERSPECTIVAS) <= set(d["objetivos"]["perspectiva"]), tem("objetivos")),
        "servicos": (bool((d["catalogo"]["horas_unit"] > 0).any()), bool((d["catalogo"]["horas_unit"] > 0).any())),
        "assinaturas": (tem("assinaturas") and bool((d["assinaturas"]["churn_mensal_pct"] > 0).all()), tem("assinaturas")),
        "preco": (tem("catalogo"), tem("catalogo")),
        "dre": (rec and tem("custos_fixos"), rec),
        "fluxo": (rec and tem("capital_giro_prazos"), rec),
        "indicadores": (rec and tem("inv_fixos"), rec),
        "graficos": (rec, rec),
        "simulador": (rec, rec),
        "experimentos": (tem("experimentos"), tem("experimentos")),
        "riscos": (tem("riscos", 3), tem("riscos")),
        "cronograma": (tem("cronograma", 3), tem("cronograma")),
        "realizado": (tem("realizado"), tem("realizado")),
        "sumario": (len(str(proj["sumario_executivo"])) > 200, bool(proj["sumario_executivo"])),
        **regras_v71(proj, d),
        **regras_v8(proj, d),
        **regras_mix(proj, d),
    }


def completude_de(proj, d) -> float:
    r = regras_situacao(proj, d)
    # "Orçado x realizado" só faz sentido depois da abertura: não entra no cálculo
    tipo = str(proj.get("tipo_negocio") or "Produtos")
    fora = {"realizado", "adocao", "plano_mkt", "sazonalidade", "tornado", "taxas", "mapa", "arvore"} | OPCIONAIS_V8
    if not tipo.startswith(("Serviço", "Misto")):
        fora.add("servicos")
    if not tipo.startswith(("Assinatura", "Misto")):
        fora.add("assinaturas")
    if tipo.startswith("Assinatura") and not r["produtos"][1]:
        fora.update({"produtos", "preco"})
    vals = [1 if c else (0.5 if i else 0) for k, (c, i) in r.items() if k not in fora]
    return sum(vals) / max(len(vals), 1)


# ------------------------------------------------------------------------------
# 2. SUMÁRIO EXECUTIVO AUTOMÁTICO (rascunho para o aluno revisar)
# ------------------------------------------------------------------------------
def gerar_sumario(proj: dict, d: dict, res: dict) -> str:
    a1 = res["anos"][1]
    nome = proj["nome_negocio"] or "O negócio"
    seg = d["segmentos"]
    seg_top = ""
    if not seg.empty:
        s_ = seg.assign(m=seg.iloc[:, 2:7].mean(axis=1)).sort_values("m", ascending=False)
        seg_top = s_["nome"].iloc[0]
    props = d.get("propostas")
    valores = ""
    if props is not None and not props.empty:
        valores = "; ".join(dict.fromkeys(v.strip() for x in props["valores"] for v in str(x).split(";") if v.strip()))
    cat = res["cat"]
    top = cat.sort_values("receita_mensal", ascending=False)["descricao"].head(3).tolist() if not cat.empty else []
    socios = ", ".join(d["socios"]["nome"].tolist())
    diag = pontuacao_swot(d["swot"])
    estrategia = QUADRANTES[diag["predominante"]][0] if diag["predominante"] else "não definida"
    mk = mercado(proj, res)
    tir = res["tir"]
    pb = res["payback_desc"]
    viavel = res["vpl"] > 0 and a1["resultado"] > 0
    p = []
    setores = [s.strip().lower() for s in str(proj["setor_atuacao"] or "").split(";") if s.strip()]
    setores = [s for s in setores if s not in ("comércio", "serviço", "indústria")] or setores
    setor_txt = (", ".join(setores[:2]) if len(setores) <= 2 else f"{setores[0]} e {setores[1]}") or "atuação a definir"
    p.append(f"{nome} é um empreendimento do setor de {setor_txt}, localizado em "
             f"{proj['cidade_uf'] or 'cidade a definir'}, com o objetivo de {str(proj['objetivo']).lower()}. "
             f"O enquadramento tributário previsto é o {proj['enquadramento']} e o porte, {proj['porte']}.")
    if seg_top or valores:
        p.append(f"O segmento de clientes priorizado é {seg_top or 'a definir'}. A proposta de valor se apoia em "
                 f"{valores or 'benefícios a definir'}" + (f", com destaque para {', '.join(top)}." if top else "."))
    if mk["preenchido"]:
        p.append(f"O mercado alcançável foi estimado em {num_br(mk['sam_c'], 0)} clientes ({brl(mk['sam'])} por ano), e a meta "
                 f"de conquista (SOM) em {num_br(mk['som_c'], 0)} clientes. A receita prevista para o primeiro ano representa "
                 f"{pct(mk['pct_som'])} desse SOM.")
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        p.append(f"O modelo de receita é baseado em assinaturas: a projeção chega a {num_br(rc['ativos12'], 0)} assinantes e MRR de "
                 f"{brl(rc['mrr12'])} no Mês 12, com churn médio de {pct(rc['churn'])} ao mês e relação LTV/CAC de "
                 f"{num_br(rc['ltv_cac'], 1)}x.")
    p.append(f"A análise SWOT indica a {estrategia.lower()} como postura predominante.")
    p.append(f"O investimento total necessário é de {brl(res['investimento_total'])}, composto por {brl(res['inv_fixo'])} em "
             f"investimentos fixos, {brl(res['inv_pre'])} em gastos pré-operacionais, {brl(res['inv_estoque'])} em estoque inicial "
             f"e {brl(res['capital_giro'])} de capital de giro. As fontes são {brl(res['capital_proprio'])} de capital próprio"
             + (f" ({socios})" if socios else "") + f" e {brl(res['capital_terceiros'])} de capital de terceiros.")
    p.append(f"Para o primeiro ano, projeta-se receita bruta de {brl(a1['rec_bruta'])}, margem de contribuição de {pct(a1['imc'])} "
             f"e resultado líquido de {brl(a1['resultado'])} (lucratividade de {pct(div(a1['resultado'], a1['rec_bruta']))}). "
             f"O ponto de equilíbrio contábil é de {brl(a1['pe_contabil'])} por ano.")
    p.append(f"Na avaliação do investimento em {res['N']} ano{'s' if res['N'] > 1 else ''}, à TMA de {pct(res['tma'])} ao ano, o VPL é de {brl(res['vpl'])}"
             + (f", a TIR é de {pct(tir)}" if tir is not None else "")
             + (f" e o payback descontado ocorre em {num_br(pb, 1)} meses." if pb == pb else f" e o investimento não se paga em {res['MESES']} meses."))
    p.append("Com base nesses resultados, o negócio se mostra " + ("economicamente viável" if viavel else
             "ainda não viável nas premissas atuais, sendo necessário rever preços, custos ou investimentos") + ".")
    return "\n".join(p)


# ------------------------------------------------------------------------------
# 3. TEXTO PARA ANÁLISE COM IA (copiar e colar no ChatGPT, Gemini, Claude...)
# ------------------------------------------------------------------------------
TAREFAS_IA = {
    "viabilidade_real": "Teste de viabilidade real: premissas comparadas com o mercado, cenário realista, veredito e ajustes para viabilizar",
    "validacao": "Plano de validação de 90 dias para as premissas mais arriscadas",
    "indicadores": "Diagnóstico de cada indicador: situação atual, o que significa para a saúde do negócio e sinais de alerta",
    "concorrencia": "Estratégias para se diferenciar e superar os concorrentes",
    "custos": "Plano de redução de custos e de melhoria de margens",
    "preco": "Revisão da política de preços por produto",
    "marketing": "Plano de marketing e de aquisição de clientes para os primeiros 6 meses",
    "riscos": "Principais riscos do plano e como mitigá-los",
    "graficos": "Gráficos interpretativos dos principais números",
    "banca": "Simulação de banca examinadora: avaliação das respostas do grupo e novas perguntas difíceis",
    "recorrencia": "Assinaturas e serviços: churn, LTV/CAC, preço dos planos e custo da hora",
    "smart": "Revisão da missão, visão, valores e objetivos SMART",
    "estrategia": "Análise estratégica integrada: Canvas, Cinco Forças, GE/McKinsey, BCG e Balanced Scorecard",
    "coerencia": "Coerência entre as ferramentas (mercado x vendas x capacidade x equipe x finanças)",
    "reforma": "Efeito da reforma tributária e escolha entre Simples puro e híbrido",
    "posicionamento": "Proposta de valor, posicionamento, curva de valor (ERRC) e orçamento de marketing por canal",
    "setor": "Comparação com a média do setor (IBGE) e com os dados do município e das famílias",
    "investidor": "Valuation, rodada de investimento e financiamento do ponto de vista de um investidor",
    "impacto": "Impacto social, ODS e adequação a editais de fomento",
}


def _md_tabela(df: pd.DataFrame, max_linhas=40) -> str:
    if df is None or df.empty:
        return "(sem dados)"
    df = df.head(max_linhas).copy()
    cols = list(df.columns)
    linhas = ["| " + " | ".join(map(str, cols)) + " |", "|" + "---|" * len(cols)]
    for _, r in df.iterrows():
        linhas.append("| " + " | ".join(str(r[c]).replace("|", "/").replace("\n", " ") for c in cols) + " |")
    return "\n".join(linhas)


def _prompt_v71(proj: dict, d: dict, res: dict) -> list:
    """Seções do prompt com validação, marketing, incerteza e defesa (v7.1)."""
    b = []
    preco = preco_referencia(proj, res)
    pl = placar_evidencias(proj, d, res)
    b.append(f"\n## 12. Placar de evidências ({pct(pl['score'], 0)} do resultado apoiado em premissas validadas)")
    tb = pl["tabela"].assign(peso=pl["tabela"]["peso"].map(lambda v: pct(v, 0)), status=pl["tabela"]["status"].map(STATUS_CURTO))
    b.append(_md_tabela(tb[["premissa", "valor_plano", "peso", "status", "fonte"]].rename(columns={
        "premissa": "Premissa", "valor_plano": "Valor no plano", "peso": "Peso no resultado", "status": "Origem", "fonte": "Fonte"})))
    ent = resumo_entrevistas(d, preco)
    if ent["n"]:
        b.append(f"\n## 12. Entrevistas com clientes ({ent['n']})")
        b.append(f"- Pagariam: {pct(ent['pct_sim'], 0)} sim, {pct(ent['pct_talvez'], 0)} talvez, {pct(ent['pct_nao'], 0)} não; "
                 f"valor mediano citado {brl(ent['mediana_valor'])} x preço de referência do plano {brl(preco)}.")
        b.append(_md_tabela(ent["dores"].assign(**{"% das entrevistas": ent["dores"]["% das entrevistas"].map(lambda v: pct(v, 0))})))
        b.append(_md_tabela(d["entrevistas"][["perfil", "dores", "solucao_atual", "pagaria", "valor_pagaria", "aprendizado"]], 20))
    vw = van_westendorp(d.get("precos_vw"))
    if vw and vw.get("ok"):
        b.append(f"\n## 12. Teste de preço Van Westendorp ({vw['n']} respostas, produto: {proj.get('vw_produto') or 'preço médio'})")
        b.append(f"- Faixa aceitável {brl(vw['pmc'])} a {brl(vw['pme'])}; preço ótimo {brl(vw['opp'])}; indiferença {brl(vw['ipp'])}; "
                 f"preço do plano {brl(preco)}.")
    ia = intencao_ajustada(proj, res)
    if ia["ok"]:
        b.append(f"- Intenção de compra: {num_br(proj.get('intencao_certamente_pct'), 0)}% certamente e "
                 f"{num_br(proj.get('intencao_provavel_pct'), 0)}% provavelmente; após fatores de "
                 f"{num_br(proj.get('fator_certamente_pct'), 0)}% e {num_br(proj.get('fator_provavel_pct'), 0)}%: "
                 f"{num_br(ia['ajustados'], 0)} clientes x {num_br(ia['plano'], 0)} clientes supostos pelo plano no Ano 1.")
    if proj.get("vpc_job") or proj.get("vpc_dores"):
        b.append("\n## 12. Canvas da proposta de valor")
        b.append(f"- Job do cliente: {proj.get('vpc_job') or ''}")
        for k_, rot, _ in BLOCOS_VPC:
            if proj.get(k_):
                b.append(f"- {rot}: " + "; ".join(linhas_txt(proj[k_])))
        enc = encaixe_vpc(proj, d)
        if enc["total"]:
            b.append(f"- Encaixe: {enc['cobertos']} de {enc['total']} dores e ganhos têm resposta na oferta.")
    frase = frase_posicionamento(proj)
    if frase:
        b.append(f"\n## 12. Posicionamento\n{frase}")
    if not d["mapa_perceptual"].empty:
        b.append(f"Mapa perceptual (eixo X: {proj.get('mapa_eixo_x')}; eixo Y: {proj.get('mapa_eixo_y')}):")
        b.append(_md_tabela(d["mapa_perceptual"]))
    if not d["errc"].empty:
        b.append("\n## 12. Curva de valor e matriz ERRC (oceano azul)")
        b.append(_md_tabela(d["errc"]))
        b.append(leitura_errc(d))
    oc = orcamento_canais(proj, d, res)
    if not oc.empty:
        b.append("\n## 12. Orçamento de marketing por canal")
        b.append(_md_tabela(oc.assign(**{c: oc[c].map(lambda x: brl(x) if x == x else "n.a.") for c in ("Gasto mensal", "CAC", "Valor do cliente")},
                                      **{"Retorno por R$ 1": oc["Retorno por R$ 1"].map(lambda x: num_br(x, 2) if x == x else "n.a.")})))
    if not d["aarrr"].empty:
        b.append("Funil AARRR:")
        b.append(_md_tabela(d["aarrr"]))
    b.append(f"- Meta de NPS no primeiro ano: {num_br(proj.get('nps_meta') or 0, 0)}.")
    if res["anos"][1]["rec_bruta"] > 0:
        mc = monte_carlo(res, faixas_mc(d))
        b.append(f"\n## 12. Simulação de Monte Carlo ({mc['n']} cenários, distribuição triangular)")
        b.append(_md_tabela(faixas_mc(d).rename(columns={"variavel": "Variável", "minimo_pct": "Pior caso (%)", "maximo_pct": "Melhor caso (%)"})))
        b.append(f"- Chance de VPL positivo: {pct(mc['p_vpl_pos'], 0)}; VPL em 5% / 50% / 95%: {brl(mc['p5'])} / {brl(mc['p50'])} / "
                 f"{brl(mc['p95'])}; chance de caixa negativo no Ano 1: {pct(mc['p_caixa_neg'], 0)}. {leitura_mc(mc)}")
        pr = avaliar_prontidao(proj, d, res, mc, preco)
        perg = perguntas_banca(proj, d, res, mc, preco)
    else:
        pr, perg = None, []
    pa = perda_aceitavel(proj, d, res)
    b.append(f"- Perda potencial se o negócio fechar: {brl(pa['potencial'])}; perda aceitável informada pelos sócios: {brl(pa['aceitavel'])}.")
    bc = bass_comparacao(proj, res)
    if bc["ok"]:
        b.append(f"\n## 12. Curva de adoção de Bass (mercado potencial {num_br(bc['m'], 0)}, p = {num_br(bc['p'], 3)}, q = {num_br(bc['q'], 2)})")
        tbb = bc["tabela"].copy()
        for c in tbb.columns:
            if c != "Ano":
                tbb[c] = tbb[c].map(lambda x: num_br(x, 2))
        b.append(_md_tabela(tbb))
    if pr is not None:
        nome = next(n for k, n, _ in NIVEIS_PRONTIDAO if k == max(pr["nivel"], 1))
        prox = next(((n, c) for k, n, c in NIVEIS_PRONTIDAO if k == pr["nivel"] + 1), None)
        b.append(f"\n## 12. Régua de prontidão: nível {pr['nivel']} de 9 ({nome})"
                 + (f"; próximo degrau: {prox[0]} ({prox[1]})" if prox else ""))
    if perg:
        resp = {str(r["pergunta"]): str(r["resposta"]) for _, r in d["banca_respostas"].iterrows()}
        b.append("\n## 12. Banca simulada (perguntas geradas pelo programa e respostas do grupo)")
        for q in perg:
            b.append(f"- [{q['tema']}] {q['pergunta']} ({q['motivo']}) Resposta do grupo: {resp.get(q['pergunta']) or '(sem resposta)'}")
    ps = porte_sugerido(res, d)
    b.append(f"\n## 12. Formalização\n- Enquadramento sugerido pela receita do Ano 1: {ps['porte']}.")
    if not d["formalizacao"].empty:
        b.append(_md_tabela(d["formalizacao"][["item", "status"]]))
    return b


def _prompt_v8(proj: dict, d: dict, res: dict) -> list:
    """Seções do prompt com as ferramentas da versão 8."""
    b = []
    dp = d.get("dados_publicos", pd.DataFrame())
    if not dp.empty:
        b.append(f"\n## 12. Dados públicos ({proj.get('municipio_nome') or ''} - {proj.get('uf_sigla') or ''})")
        for _, r in dp.iterrows():
            b.append(f"- {r['rotulo']} ({r['periodo']}): {num_br(r['valor'], 2)}. Fonte: {r['fonte']}")
    if proj.get("setor_referencia") and res["anos"][1]["rec_bruta"] > 0:
        try:
            cs = comparar_setor(proj["setor_referencia"], proj, d, res)
            b.append(f"\n## 12. Comparação com o setor: {proj['setor_referencia']} (média Brasil, IBGE 2024)")
            b.append(_md_tabela(cs.assign(**{c: cs[c].map(lambda v: num_br(v, 3)) for c in ("Plano", "Setor (Brasil)", "Plano ÷ setor")})))
        except Exception:
            pass
    b.append("- Sobrevivência em 5 anos (Sebrae, 2026): " + "; ".join(f"{k}: {pct(v, 1)}" for k, v in SOBREVIVENCIA_5A.items()) + ".")
    est = d.get("swot_estrategias", pd.DataFrame())
    if not est.empty:
        b.append("\n## 12. SWOT cruzada (estratégias)")
        b.append(_md_tabela(est))
    lc = lean_canvas(proj, d, res)
    b.append("\n## 12. Lean Canvas")
    for k_, v_ in lc.items():
        if v_:
            b.append(f"- {k_}: " + "; ".join(str(x) for x in v_))
    emp = d["emprestimos"]
    if not emp.empty:
        b.append("\n## 12. Financiamento")
        for _, r in emp.iterrows():
            if float(r["valor"] or 0) > 0:
                rs = resumo_emprestimo(r["valor"], r["taxa_anual_pct"], r["prazo_meses"], r["carencia_meses"], r["sistema"])
                b.append(f"- {r['descricao']}: {brl(r['valor'])}, {num_br(r['taxa_anual_pct'], 2)}% a.a., {num_br(r['prazo_meses'], 0)} meses, "
                         f"carência de {num_br(r['carencia_meses'] or 0, 0)} meses, sistema {r['sistema'] or 'Price'}; juros totais {brl(rs['juros'])}; "
                         f"CET {pct(rs['cet'])} a.a.")
    vl = valuation(res, proj, d)
    b.append(f"\n## 12. Valuation\n- Fluxo de caixa descontado: {brl(vl['equity_dcf'])} (perpetuidade com crescimento de "
             f"{pct(vl['g'])}); faixa entre métodos: {brl(vl['faixa'][0])} a {brl(vl['faixa'][1])}.")
    if float(proj.get("rodada_investimento") or 0) > 0:
        pre, inv = float(proj.get("rodada_pre_money") or 0), float(proj["rodada_investimento"])
        b.append(f"- Rodada: aporte de {brl(inv)} com pre-money de {brl(pre)}; investidor fica com {pct(div(inv, pre + inv))}.")
    imp = [(rot, proj.get(k_)) for k_, rot, _ in ETAPAS_MUDANCA if proj.get(k_)]
    if imp or proj.get("ods"):
        b.append("\n## 12. Impacto social")
        for rot, v_ in imp:
            b.append(f"- {rot}: {v_}")
        if proj.get("ods"):
            b.append(f"- ODS: {str(proj['ods']).replace(';', '; ')}")
        if not d["impacto_indicadores"].empty:
            b.append(_md_tabela(d["impacto_indicadores"]))
    cm = d.get("comentarios", pd.DataFrame())
    if not cm.empty:
        abertos = cm[cm["resolvido"].fillna(0).astype(int) == 0]
        if not abertos.empty:
            b.append("\n## 12. Comentários do professor ainda não resolvidos")
            for _, r in abertos.iterrows():
                b.append(f"- [{r['pagina']}] {r['comentario']}")
    return b


def gerar_prompt_ia(proj: dict, d: dict, res: dict, tarefas: list, publico="estudante de Administração",
                    piv_conc: pd.DataFrame | None = None) -> str:
    a1 = res["anos"][1]
    sem = semaforo(res)
    cat = res["cat"]
    chk = tabela_checagens_v5(res, d, proj)
    mk = mercado(proj, res)
    diag = pontuacao_swot(d["swot"])
    b = []
    b.append(f"Você é um consultor de negócios e professor de finanças empresariais. Analise o plano de negócio abaixo, "
             f"elaborado por um {publico}, com linguagem clara, didática e em português do Brasil.")
    b.append("REGRAS: use SOMENTE os números fornecidos (não invente valores); quando fizer contas, mostre o cálculo; "
             "aponte dados incoerentes ou faltantes; separe fatos dos dados de suas sugestões; seja específico para este negócio.")
    b.append("\n## 1. Identificação")
    b.append(f"- Negócio: {proj['nome_negocio']} | Cidade: {proj['cidade_uf']} | Setor: {proj['setor_atuacao']}")
    b.append(f"- Enquadramento: {proj['enquadramento']} | Porte: {proj['porte']} | Objetivo: {proj['objetivo']}")
    b.append(f"- Tipo de negócio: {proj.get('tipo_negocio') or 'Produtos'} | Horizonte do plano: {res['N']} ano(s)")
    if any(proj.get(k) for k in ("missao", "visao", "valores")):
        b.append("\n## 1. Missão, visão e valores")
        b.append(f"- Missão: {proj.get('missao', '')}\n- Visão: {proj.get('visao', '')}\n- Valores: {proj.get('valores', '')}")
    ob = d.get("objetivos")
    if ob is not None and not ob.empty:
        b.append("\n## 1. Objetivos SMART (checagem automática do programa)")
        b.append(_md_tabela(pd.DataFrame({
            "Objetivo": ob["objetivo"], "Perspectiva": ob["perspectiva"], "Indicador": ob["indicador"],
            "Meta": [f"{num_br(m, 1)} {u}" for m, u in zip(ob["meta"], ob["unidade"])],
            "Prazo (mês)": ob["prazo_meses"].map(lambda v: num_br(v, 0)), "Iniciativa": ob["iniciativa"],
            "SMART": ["".join(L if ok else "-" for L, (ok, _) in checar_smart(o).items()) for _, o in ob.iterrows()]})))
    seg = d["segmentos"]
    if not seg.empty:
        s_ = seg.assign(Média=seg.iloc[:, 2:7].mean(axis=1).round(1)).sort_values("Média", ascending=False)
        b.append("\n## 2. Segmentos de clientes (notas 1 a 10)")
        b.append(_md_tabela(s_[["nome", "descricao", "Média"]].rename(columns={"nome": "Segmento", "descricao": "Descrição"})))
    pers = d["personas"]
    if not pers.empty:
        b.append("\n## 2. Personas")
        for _, p_ in pers.iterrows():
            b.append(f"- {p_['nome']}, {num_br(p_['idade'], 0)} anos, {p_['profissao']}: {p_['descricao']} Ganhos: {p_['ganhos']}. "
                     f"Dores: {p_['desafios']}. Objeções: {p_['duvidas_objecoes']}.")
    jor = d["jornada_cliente"]
    if not jor.empty:
        b.append("\n## 2. Jornada do cliente")
        b.append(_md_tabela(jor[["persona_nome", "macro_fase", "etapa", "estado_emocional", "oportunidades"]].rename(columns={
            "persona_nome": "Persona", "macro_fase": "Fase", "etapa": "Etapa", "estado_emocional": "Emoção",
            "oportunidades": "Oportunidade"}), 30))
    props = d.get("propostas")
    if props is not None and not props.empty:
        b.append("\n## 3. Propostas de valor")
        for _, p_ in props.iterrows():
            b.append(f"- Persona {p_['persona_nome']}: deseja {p_['que_desejam']}; por meio de {p_['por_meio']}; "
                     f"diferente de {p_['diferentemente']}; oferecemos {p_['oferecem']}. Valores: {p_['valores']}")
    if mk["preenchido"]:
        b.append("\n## 4. Mercado")
        b.append(f"- TAM {num_br(mk['tam_c'], 0)} clientes ({brl(mk['tam'])}/ano); SAM {num_br(mk['sam_c'], 0)} ({brl(mk['sam'])}); "
                 f"SOM {num_br(mk['som_c'], 0)} ({brl(mk['som'])}). Receita do Ano 1 = {pct(mk['pct_som'])} do SOM.")
    if piv_conc is not None and not piv_conc.empty:
        b.append("\n## 5. Matriz de concorrência (notas 0 a 10)")
        b.append(_md_tabela(piv_conc.round(1)))
    sw = d["swot"]
    if not sw.empty:
        b.append("\n## 6. SWOT")
        for nat in ("Força", "Fraqueza", "Oportunidade", "Ameaça"):
            itens = sw[sw["natureza"] == nat]
            if not itens.empty:
                b.append(f"- {nat}s: " + "; ".join(f"{r['fator']} ({r['impacto']})" for _, r in itens.iterrows()))
        if diag["predominante"]:
            b.append(f"- Quadrante predominante: {QUADRANTES[diag['predominante']][0]}")
    b.append("\n## 6. Análise estratégica")
    bl = canvas_blocos(proj, d, res)
    b.append("Business Model Canvas: " + " | ".join(f"{k}: {'; '.join(v) or '(vazio)'}" for k, v in bl.items()))
    pt = porter_tabela(d)
    if (pt["Intensidade"] > 0).any():
        b.append("Cinco Forças de Porter (1 = fraca, 5 = muito forte):")
        b.append(_md_tabela(pt.assign(Intensidade=pt["Intensidade"].map(lambda v: num_br(v, 0)))))
        b.append(f"- Leitura do programa: {porter_leitura(pt)}")
    ge_ = ge_tabela(d)
    if not ge_.empty:
        b.append("Matriz GE/McKinsey (0 a 10):")
        b.append(_md_tabela(ge_[["Segmento", "Atratividade", "Força competitiva", "Recomendação"]]))
    bcg_ = bcg_tabela(d, res, float(proj.get("bcg_limite_crescimento_pct") or 10))
    if not bcg_.empty:
        b.append(f"Matriz BCG (crescimento alto a partir de {num_br(proj.get('bcg_limite_crescimento_pct') or 10, 1)}% a.a.):")
        b.append(_md_tabela(bcg_.assign(**{"Receita anual": bcg_["Receita anual"].map(brl), "Margem": bcg_["Margem"].map(pct)})))
    ind_bsc = bsc_indicadores_plano(proj, d, res)
    b.append("Balanced Scorecard - indicadores do plano: " + " | ".join(
        f"{p}: " + "; ".join(f"{n} = {v}" for n, v in ind_bsc[p]) for p in PERSPECTIVAS if ind_bsc.get(p)))
    if any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
        b.append("\n## 7. Marketing (4Ps)")
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            b.append(f"- {rot}: {proj.get(k, '')}")
    if not d["canais"].empty:
        b.append("\n## 7. Canais e funil de vendas")
        b.append(_md_tabela(d["canais"].rename(columns={"canal": "Canal"})))
        if not d["funil_vendas"].empty:
            b.append(_md_tabela(d["funil_vendas"]))
    b.append("\n## 7. Plano operacional")
    b.append(f"- Localização: {proj.get('loc_endereco', '')}. {proj.get('loc_justificativa', '')}\n- Processos: {proj.get('op_processos', '')}"
             f"\n- Horário: {proj.get('op_horario', '')}")
    cd_ = capacidade_demanda(d, res)
    if not cd_.empty:
        b.append(_md_tabela(cd_.assign(**{"Uso Mês 1": cd_["Uso Mês 1"].map(pct), "Uso Mês 12": cd_["Uso Mês 12"].map(pct)})))
    pes_ = d["pessoal"]
    if not pes_.empty:
        b.append(_md_tabela(pd.DataFrame({"Cargo": pes_["cargo"], "Área": pes_["area"], "Nº": pes_["num_empregados"].map(lambda v: num_br(v, 0)),
                                          "Salário": pes_["salario_mensal"].map(brl), "Começa no mês": pes_["mes_inicio"].map(lambda v: num_br(v, 0))})))
    fo_ = d["fornecedores"]
    if not fo_.empty:
        b.append(_md_tabela(fo_[["nome", "itens", "prazo_pagamento_dias", "critico"]].rename(columns={
            "nome": "Fornecedor", "itens": "Itens", "prazo_pagamento_dias": "Prazo pagto (dias)", "critico": "Crítico"})))
    b.append("\n## 8. Investimento e fontes")
    b.append(f"- Investimento total {brl(res['investimento_total'])}: fixos {brl(res['inv_fixo'])}; pré-operacionais {brl(res['inv_pre'])}; "
             f"estoque {brl(res['inv_estoque'])}; capital de giro {brl(res['capital_giro'])}.")
    b.append(f"- Fontes: próprio {brl(res['capital_proprio'])}; terceiros {brl(res['capital_terceiros'])}.")
    if not cat.empty:
        b.append("\n## 9. Produtos (Mês 1)")
        t = pd.DataFrame({"Produto": cat["descricao"], "Categoria": cat["categoria"], "Preço": cat["preco_venda"].map(brl),
                          "Custo variável": cat["custo_unit"].map(brl), "Custo absorção": cat["custo_absorcao_unit"].map(brl),
                          "Margem": cat["margem_pct"].map(pct), "Qtd/mês": cat["qtd_mensal"].map(lambda v: num_br(v, 0)),
                          "Receita/mês": cat["receita_mensal"].map(brl)})
        b.append(_md_tabela(t, 60))
    b.append(f"\n- Tipo de negócio: {proj.get('tipo_negocio') or 'Produtos'}")
    ch = custo_hora(res, d, proj)
    if not ch["tabela"].empty:
        b.append("\n## 9b. Serviços e custo da hora")
        b.append(f"- Equipe da operação: {num_br(ch['pessoas'], 0)} pessoa(s); {num_br(ch['horas_cobraveis'], 0)} horas cobráveis/mês "
                 f"(utilização {pct(ch['utilizacao'], 0)}); custo da hora produtiva {brl(ch['custo_hora'])}; ocupação no Mês 12 "
                 f"{pct(ch['ocupacao_m12'])}.")
        t_ = ch["tabela"]
        b.append(_md_tabela(pd.DataFrame({"Serviço": t_["Serviço"], "Horas": t_["Horas por serviço"].map(lambda v: num_br(v, 1)), "Custo pleno": t_["Custo pleno"].map(brl),
                                          "Preço": t_["Preço atual"].map(brl),
                                          "Sugerido": t_["Preço sugerido"].map(lambda v: brl(v) if v == v else "n.c."),
                                          "R$/hora cobrada": t_["Valor da hora cobrada"].map(brl)})))
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        a_ = d["assinaturas"]
        b.append("\n## 9c. Assinaturas (receita recorrente)")
        b.append(_md_tabela(pd.DataFrame({"Plano": a_["plano"], "Público": a_["publico"], "Preço por ciclo": a_["preco_ciclo"].map(brl),
                                          "Ciclo (meses)": a_["ciclo_meses"].map(lambda v: num_br(v, 0)),
                                          "Novos no Mês 1": a_["novos_mes1"].map(lambda v: num_br(v, 0)),
                                          "Crescimento dos novos %/mês": a_["crescimento_novos_pct"].map(lambda v: num_br(v, 1)),
                                          "Churn %/mês": a_["churn_mensal_pct"].map(lambda v: num_br(v, 1)),
                                          "Custo de entrega/assinante": a_["custo_entrega_mensal"].map(brl),
                                          "Taxa de cobrança %": a_["taxa_cobranca_pct"].map(lambda v: num_br(v, 1))})))
        b.append(f"- MRR Mês 1 {brl(rc['mrr1'])}; MRR Mês 12 {brl(rc['mrr12'])}; ARR {brl(rc['arr12'])}; assinantes no Mês 12 "
                 f"{num_br(rc['ativos12'], 0)}; ARPU {brl(rc['arpu'])}; churn médio {pct(rc['churn'])}; margem bruta "
                 f"{pct(rc['margem_bruta'])}; LTV {brl(rc['ltv'])}; CAC {brl(rc['cac'])}; LTV/CAC {num_br(rc['ltv_cac'], 1)}x; "
                 f"payback do CAC {num_br(rc['payback_cac'], 1)} meses; receita antecipada no fim do Ano 1 {brl(rc['receita_diferida_a1'])}.")
    b.append("\n## 10. Estrutura de custos mensal")
    b.append(f"- Custos diretos {brl(res['cdir_m1'])}; impostos {pct(res['aliq'])} e despesas variáveis {pct(res['var_pct'])} da receita; "
             f"MOD {brl(res['mod_m'])}; custos indiretos de produção {brl(res['cif_m'])}; depreciação {brl(res['dep_m'])}; "
             f"pessoal adm./comercial {brl(res['pessoal_adm_m'] + res['pessoal_com_m'])}; despesas adm. {brl(res['desp_adm_m'])}, "
             f"comerciais {brl(res['desp_com_m'])}, financeiras {brl(res['desp_fin_m'])}.")
    cf = d["custos_fixos"]
    if not cf.empty:
        b.append(_md_tabela(pd.DataFrame({"Custo fixo": cf["descricao"], "Classificação": cf["classificacao"],
                                          "R$/mês": cf["valor_mensal"].map(brl)})))
    b.append(f"\n## 11. DRE (Anos 1 a {res['N']})")
    b.append(_md_tabela(tabela_dre(res)))
    b.append("\n## 12. Indicadores e semáforo")
    b.append(_md_tabela(pd.DataFrame({"Indicador": [s_["Indicador"] for s_ in sem], "Valor": [s_["Valor"] for s_ in sem],
                                      "Situação": [{"🟢": "bom", "🟡": "atenção", "🔴": "crítico"}[s_["Status"]] for s_ in sem]})))
    b.append(f"- VPL {brl(res['vpl'])}; TIR {pct(res['tir']) if res['tir'] is not None else 'n.a.'}; TMA {pct(res['tma'])}; "
             f"payback descontado {num_br(res['payback_desc'], 1) + ' meses' if res['payback_desc'] == res['payback_desc'] else '> ' + str(res['MESES']) + ' meses'}; "
             f"PE contábil {brl(a1['pe_contabil'])}/ano; margem de segurança {pct(div(a1['rec_bruta'] - a1['pe_contabil'], a1['rec_bruta']))}.")
    sens = sensibilidade(res)
    b.append("- Sensibilidade (efeito de ±10% no resultado do Ano 1): " +
             "; ".join(f"{r['Variável']} {brl(r['Impacto absoluto'])}" for _, r in sens.iterrows()))
    b.append("\n## 12. Cenários")
    b.append(_md_tabela(cenarios(res, proj)))
    if float(proj.get("reforma_b2b_pct") or 0) > 0 or proj.get("anexo_simples"):
        an = proj.get("anexo_simples") or "Anexo III - Serviços"
        ano_r = int(float(proj.get("reforma_ano") or 2027))
        cc = float(proj.get("reforma_compras_credito") or -1)
        rr = simular_reforma(a1["rec_bruta"], res["aliq"], an, ano_r, float(proj.get("reforma_cbs_pct") or 11.3) / 100,
                             float(proj.get("reforma_ibs_pct") or 0.1) / 100, float(proj.get("reforma_b2b_pct") or 0) / 100,
                             a1["custo_dir"] + a1["fixos"] if cc < 0 else cc)
        b.append(f"\n## 12. Reforma tributária (simulação didática, {ano_r}, {an})")
        b.append(f"- Simples puro: tributos {brl(rr['puro_total'])}, crédito aos clientes empresas {brl(rr['puro_credito_cliente'])}. "
                 f"Simples híbrido: DAS {brl(rr['das_hibrido'])} + IBS/CBS líquidos {brl(rr['iva_debito'] - rr['iva_credito'])} = "
                 f"{brl(rr['hib_total'])}, crédito aos clientes {brl(rr['hib_credito_cliente'])}. Vendas B2B: "
                 f"{num_br(proj.get('reforma_b2b_pct') or 0, 0)}%.")
    for tab_, tit, cols in (("riscos", "Riscos", ["risco", "probabilidade", "impacto", "mitigacao"]),
                            ("cronograma", "Cronograma de implantação", ["atividade", "mes_inicio", "duracao_meses", "custo"]),
                            ("experimentos", "Experimentos de validação", ["hipotese", "experimento", "metrica", "criterio_sucesso", "status"]),
                            ("realizado", "Orçado x realizado (meses já operados)", ["mes", "receita_real", "custos_variaveis_real", "custos_fixos_real"])):
        if d.get(tab_) is not None and not d[tab_].empty:
            b.append(f"\n## 12. {tit}")
            b.append(_md_tabela(d[tab_][cols]))
    for k_, tit in (("sumario_executivo", "Sumário executivo escrito pelo grupo"), ("conclusao", "Conclusão do grupo")):
        if str(proj.get(k_) or "").strip():
            b.append(f"\n## 12. {tit}\n{proj[k_]}")
    b.extend(_prompt_v71(proj, d, res))
    b.extend(_prompt_v8(proj, d, res))
    pend = chk[chk["Situação"] != "OK"]
    if not pend.empty:
        b.append("\n## 13. Inconsistências apontadas pelo programa")
        for _, r in pend.iterrows():
            b.append(f"- {r['Teste']}: {r['Valor apurado']}")
    fm = res["fluxo_mensal"]
    if not fm.empty:
        b.append(f"\n- Fluxo de caixa: menor saldo {brl(fm['Saldo final'].min())} em {fm.loc[fm['Saldo final'].idxmin(), 'Mês']}; "
                 f"saldo no fim do Ano 1 {brl(fm['Saldo final'].iloc[-1])}.")

    b.append("\n# O QUE EU PRECISO QUE VOCÊ FAÇA")
    pedidos = {
        "smart": "Avalie a missão, a visão e os valores (clareza, foco no cliente, prazo da visão) e reescreva-os se preciso. Para cada "
                 "objetivo da tabela de Objetivos SMART, diga quais letras falham e proponha uma versão SMART completa, coerente com os "
                 "números do plano.",
        "estrategia": "Faça uma análise integrada da seção Análise estratégica: o Canvas é coerente com os números? As Cinco Forças "
                      "justificam as margens previstas? Cruze a matriz GE (onde investir) com a BCG (que produtos financiam o "
                      "crescimento) e com a SWOT. Termine com 5 prioridades estratégicas, cada uma ligada a um objetivo do BSC.",
        "coerencia": "Verifique a coerência entre as ferramentas: a meta de vendas cabe no mercado (SOM)? A capacidade e a equipe "
                     "(com as contratações) atendem a demanda? Os canais e o funil sustentam o volume? Os custos e preços batem com "
                     "a DRE? Liste cada incoerência encontrada e como corrigi-la.",
        "reforma": "Explique, para este negócio, o que muda com a reforma tributária (IBS e CBS) e, usando a simulação fornecida, "
                   "quando o Simples híbrido pode valer a pena. Deixe claro que a decisão deve ser confirmada com um contador.",
        "indicadores": "Para CADA indicador da seção Indicadores e semáforo, explique em linguagem simples: (a) o que ele mede; (b) a situação atual deste "
                       "negócio com o valor exato; (c) o que isso significa para a saúde do negócio; (d) uma ação concreta para melhorar. "
                       "Termine com um parecer geral (viável, viável com ressalvas ou inviável) e justifique.",
        "concorrencia": "Com base na concorrência, na SWOT e nas propostas de valor, proponha 5 estratégias para driblar os concorrentes, "
                        "cada uma com: objetivo, ação prática, custo estimado (baixo/médio/alto), prazo e indicador de sucesso.",
        "custos": "Identifique os 5 itens de custo que mais pesam e proponha como reduzi-los sem prejudicar a qualidade. Para cada um, "
                  "estime o efeito no resultado anual em R$ usando os números fornecidos. Aponte também produtos com margem baixa.",
        "preco": "Analise preço x custo variável x custo por absorção de cada produto. Diga quais preços estão baixos ou altos demais, "
                 "sugira ajustes e explique o raciocínio (markup, concorrência, percepção de valor).",
        "marketing": "Monte um plano de marketing de 6 meses (mês a mês), com canais, ações, orçamento estimado e metas, coerente com a "
                     "persona e a proposta de valor.",
        "riscos": "Liste os 5 maiores riscos que você enxerga neste plano (financeiros, de mercado e operacionais), a probabilidade, o "
                  "impacto e uma ação de prevenção para cada um.",
        "graficos": "Gere gráficos interpretativos (se puder executar código, use Python/matplotlib; se não, descreva o gráfico e monte "
                    "uma tabela pronta para virar gráfico no Excel): (1) composição do investimento; (2) estrutura de custos; "
                    "(3) receita x custos x resultado em todos os anos do horizonte; (4) ponto de equilíbrio; (5) saldo de caixa mês a mês. "
                    "Abaixo de cada gráfico, escreva 2 ou 3 frases de interpretação.",
        "recorrencia": "Se o plano tiver assinaturas ou serviços (seções de Serviços e de Assinaturas), avalie: o churn é realista para o setor? O LTV/CAC "
                       "e o payback do CAC são saudáveis? Os preços dos planos e o desconto dos planos longos fazem sentido? O custo "
                       "da hora cobre a equipe? Proponha 5 ações para reduzir cancelamentos e aumentar a receita por cliente, com o "
                       "efeito estimado no MRR.",
        "banca": "Faça o papel de uma banca examinadora. Avalie as respostas que o grupo já deu na seção Banca simulada (nota de 0 a 10, "
                 "o que faltou e uma versão melhor usando os números do plano). Depois elabore mais 5 perguntas difíceis que o "
                 "programa não fez e indique o que uma boa resposta deveria conter.",
        "viabilidade_real": "TESTE DE VIABILIDADE REAL. Atue ao mesmo tempo como investidor-anjo cético e analista de crédito de um banco. "
                            "(1) Para cada premissa do Placar de evidências, compare o valor do plano com referências do setor no Brasil "
                            "e na região (Sebrae, IBGE, associações setoriais, preços praticados por concorrentes), citando a fonte e o seu "
                            "grau de confiança; quando não tiver uma referência confiável, diga isso claramente em vez de inventar um número. "
                            "Classifique cada premissa como realista, otimista ou pessimista. (2) Monte um CENÁRIO REALISTA trocando as "
                            "premissas otimistas pelos valores que considera mais prováveis e recalcule, mostrando as contas: receita do Ano 1, "
                            "margem de contribuição, resultado operacional, ponto de equilíbrio, payback e uma estimativa do VPL com a TMA "
                            "informada. (3) Use as entrevistas, o teste de preço, a intenção de compra ajustada, a curva de Bass e a "
                            "simulação de Monte Carlo para confirmar ou contestar esse cenário. (4) Dê um veredito único: VIÁVEL, VIÁVEL COM "
                            "AJUSTES ou INVIÁVEL NO FORMATO ATUAL, com as 3 razões principais. (5) Liste os ajustes que tornariam o negócio "
                            "viável na prática (preço, mix de produtos, custos, tamanho do investimento inicial, começar com um produto mínimo, "
                            "canais, equipe, financiamento), em ordem de impacto, com o efeito estimado em R$ no resultado anual ou no VPL. "
                            "(6) Defina os critérios de parada: que resultado, nos primeiros 90 dias de testes ou de operação, mostraria que o "
                            "negócio não deve continuar ou precisa mudar de rumo. (7) Diga se a perda potencial cabe na perda aceitável dos sócios.",
        "validacao": "Monte um plano de validação de 90 dias para as premissas de maior peso que ainda não foram testadas com clientes "
                     "(veja o Placar de evidências). Para cada uma: experimento barato (entrevista, página de pré-cadastro, pré-venda, "
                     "piloto, venda manual antes de automatizar), métrica, critério de sucesso definido antes do teste, custo, prazo e o "
                     "que fazer se o resultado for ruim.",
        "setor": "Compare as premissas do plano com a média do setor (tabela de referências do IBGE) e com os dados do município e da "
                 "POF. Aponte o que está fora da curva, explique se a diferença é plausível para um pequeno negócio local e "
                 "sugira valores mais realistas quando for o caso.",
        "investidor": "Avalie o plano como um investidor-anjo: o valuation é coerente com o risco e com o estágio do negócio? A "
                      "participação oferecida na rodada é razoável? O financiamento escolhido (sistema, carência, custo efetivo) cabe "
                      "no caixa? Diga o que faria você investir ou recusar.",
        "impacto": "Avalie a teoria da mudança (as atividades levam mesmo aos resultados e ao impacto declarados?), os indicadores de "
                   "impacto e a ligação com os ODS. Sugira como fortalecer a proposta para um edital de inovação ou de negócios de "
                   "impacto, sem exagerar os efeitos.",
        "posicionamento": "Avalie o Canvas da proposta de valor (as dores mais importantes têm resposta?), a declaração de posicionamento "
                          "(é específica e defensável?), a curva de valor e a matriz ERRC (há diferenciação real ou só mais custo?) e o "
                          "orçamento por canal (CAC, retorno e coerência com o gasto comercial da DRE). Proponha uma versão melhor do "
                          "posicionamento e uma redistribuição do orçamento entre os canais.",
    }
    for i, t in enumerate([t for t in tarefas if t in pedidos], start=1):
        b.append(f"{i}. {pedidos[t]}")
    if "viabilidade_real" in tarefas:
        b.append("\nComece a resposta pelo veredito do teste de viabilidade real, em uma frase, antes das demais análises.")
    b.append("\nOrganize a resposta com títulos, use tabelas quando ajudar e termine com uma lista de próximos passos.")
    texto = "\n".join(b)
    cont = iter(range(1, 200))
    return re.sub(r"(?m)^## \d+[a-z]?\. ", lambda m_: f"## {next(cont)}. ", texto)


# ------------------------------------------------------------------------------
# 4. PITCH DECK EM PDF (slides 16:9)
# ------------------------------------------------------------------------------
class Pitch(Relatorio):
    def __init__(self, nome):
        FPDF.__init__(self, orientation="L", unit="mm", format=(190.5, 338.7))
        self.nome = nome
        self.unicode = False
        for reg, neg, ita in FONTES_CANDIDATAS:
            if all(Path(p).exists() for p in (reg, neg, ita)):
                self.add_font("Base", "", reg)
                self.add_font("Base", "B", neg)
                self.add_font("Base", "I", ita)
                self.unicode = True
                break
        self.fam = "Base" if self.unicode else "Helvetica"
        self.set_auto_page_break(False)
        self.set_margins(18, 14, 18)

    def footer(self):
        self.set_y(-10)
        self.fonte("", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 5, self.t(f"{self.nome}  ·  {self.page_no()}"), align="R")
        self.set_text_color(0, 0, 0)

    def slide(self, titulo, sub=""):
        self.add_page()
        self.set_fill_color(*AZUL)
        self.rect(0, 0, self.w, 22, style="F")
        self.set_xy(18, 6)
        self.fonte("B", 20)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, self.t(titulo))
        self.set_text_color(0, 0, 0)
        self.set_xy(18, 28)
        if sub:
            self.fonte("I", 12)
            self.set_text_color(90, 90, 90)
            self.multi_cell(0, 6, self.t(sub), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.set_text_color(0, 0, 0)
            self.ln(2)

    def topicos(self, itens, tam=15):
        self.fonte("", tam)
        for it in itens:
            if not str(it).strip():
                continue
            self.set_x(self.l_margin)
            self.multi_cell(0, tam * 0.5, self.t(f"•  {it}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.ln(2)

    def numeros(self, pares):
        n = len(pares)
        w = (self.w - 36 - (n - 1) * 6) / n
        y = self.get_y() + 4
        for i, (rot, val) in enumerate(pares):
            x = 18 + i * (w + 6)
            self.set_fill_color(232, 240, 251)
            self.rect(x, y, w, 34, style="F")
            self.set_xy(x, y + 6)
            self.fonte("B", 20)
            self.set_text_color(*AZUL)
            self.cell(w, 10, self.t(val), align="C")
            self.set_xy(x, y + 20)
            self.fonte("", 11)
            self.set_text_color(60, 60, 60)
            self.cell(w, 6, self.t(rot), align="C")
        self.set_text_color(0, 0, 0)
        self.set_y(y + 40)


def gerar_pitch(proj: dict, d: dict, res: dict) -> bytes:
    a1 = res["anos"][1]
    pp = Pitch(proj["nome_negocio"])
    # 1. Capa
    pp.add_page()
    pp.set_fill_color(*AZUL)
    pp.rect(0, 0, pp.w, pp.h, style="F")
    pp.set_text_color(255, 255, 255)
    pp.set_xy(18, 60)
    pp.fonte("B", 40)
    pp.multi_cell(0, 18, pp.t(proj["nome_negocio"]), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.fonte("", 16)
    props = d.get("propostas")
    frase = props["oferecem"].iloc[0] if props is not None and not props.empty and props["oferecem"].iloc[0] else "Plano de negócio"
    pp.multi_cell(0, 9, pp.t(frase), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.ln(10)
    pp.fonte("", 12)
    pp.multi_cell(0, 7, pp.t(" · ".join(x for x in [proj["cidade_uf"], proj.get("turma", ""), proj.get("autores", "").replace("\n", ", ")] if x)),
                  align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.set_text_color(0, 0, 0)
    # 2. Problema
    pers = d["personas"]
    pp.slide("O problema", "Quem é o cliente e qual dor ele sente")
    itens = []
    for _, p_ in pers.head(2).iterrows():
        itens.append(f"{p_['nome']}, {int(p_['idade'])} anos, {p_['profissao']}: \"{p_['citacao']}\"")
        if p_["desafios"]:
            itens.append(f"Dores: {p_['desafios']}")
    pp.topicos(itens or ["Descreva as personas no programa para preencher este slide."])
    # 3. Solução
    pp.slide("A solução", "Proposta de valor")
    itens = [f"Para {p_['persona_nome']}: {p_['por_meio']} — oferecemos {p_['oferecem']}" for _, p_ in props.iterrows()] \
        if props is not None and not props.empty else []
    top = res["cat"].sort_values("receita_mensal", ascending=False)["descricao"].head(5).tolist()
    if top:
        itens.append("Principais produtos: " + ", ".join(top))
    pp.topicos(itens)
    # 4. Mercado
    mk = mercado(proj, res)
    pp.slide("Mercado", proj.get("sam_desc") or "Tamanho do mercado")
    if mk["preenchido"]:
        pp.numeros([("TAM - mercado total", brl(mk["tam"])), ("SAM - alcançável", brl(mk["sam"])), ("SOM - meta", brl(mk["som"]))])
        pp.topicos([f"{num_br(mk['som_c'], 0)} clientes conquistáveis; a receita do Ano 1 equivale a {pct(mk['pct_som'])} do SOM."], 13)
    else:
        pp.topicos(["Preencha a ferramenta Tamanho do mercado para completar este slide."])
    # 5. Concorrência
    pp.slide("Concorrência e diferencial", "Como nos posicionamos")
    notas = d["concorrencia_notas"]
    if not notas.empty:
        med = notas.groupby("empresa")["nota"].mean().sort_values(ascending=False)
        pp.topicos([f"{e}: média {num_br(v, 1)}" for e, v in med.items()] +
                   [f"Diferente de: {props['diferentemente'].iloc[0]}" if props is not None and not props.empty else ""], 14)
    # 6. Modelo de negócio / marketing
    pp.slide("Como vamos vender", "Plano de marketing (4Ps)")
    pp.topicos([f"Produto: {proj['mkt_produto']}", f"Preço: {proj['mkt_preco']}", f"Praça: {proj['mkt_praca']}",
                f"Promoção: {proj['mkt_promocao']}"], 13)
    # 7. Operação
    pp.slide("Operação", proj.get("loc_endereco") or "")
    cd = capacidade_demanda(d, res)
    itens = [proj.get("op_processos", "")] + [f"{r['Recurso']}: uso de {pct(r['Uso Mês 12'])} da capacidade no Mês 12"
                                             for _, r in cd.iterrows()]
    pp.topicos(itens, 13)
    # 8. Números
    pp.slide("Os números", "Projeção do primeiro ano")
    pp.numeros([("Investimento total", brl(res["investimento_total"])), ("Receita - Ano 1", brl(a1["rec_bruta"])),
                ("Lucro - Ano 1", brl(a1["resultado"])), ("Lucratividade", pct(div(a1["resultado"], a1["rec_bruta"])))])
    pp.numeros([(f"VPL ({res['N']} anos)", brl(res["vpl"])), ("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a."),
                ("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else f"> {res['MESES']} meses"),
                ("Ponto de equilíbrio", brl(a1["pe_contabil"]))])
    # 9. Gráficos
    imgs = graficos_png(composicoes(res, d))
    escolhidos = [i for i in imgs if i[0] in ("Composição do investimento total", "Ponto de equilíbrio")]
    if escolhidos:
        pp.slide("Investimento e equilíbrio")
        x = 18
        for tit, png in escolhidos[:2]:
            pp.image(io.BytesIO(png), x=x, y=30, h=140)
            x += 160
    # 10. Investimento pedido
    pp.slide("O que precisamos", "Uso e fontes dos recursos")
    pp.numeros([("Capital próprio", brl(res["capital_proprio"])), ("Capital de terceiros", brl(res["capital_terceiros"])),
                ("Total", brl(res["fontes"]))])
    pp.topicos([f"Investimentos fixos: {brl(res['inv_fixo'])}", f"Pré-operacionais: {brl(res['inv_pre'])}",
                f"Estoque inicial: {brl(res['inv_estoque'])}", f"Capital de giro: {brl(res['capital_giro'])}"], 14)
    # 11. Equipe
    pp.slide("Equipe", "Quem faz acontecer")
    pp.topicos([f"{r['nome']} — {r['funcao']}: {r['perfil']}" for _, r in d["socios"].iterrows()], 13)
    # 12. Encerramento
    pp.add_page()
    pp.set_fill_color(*AZUL)
    pp.rect(0, 0, pp.w, pp.h, style="F")
    pp.set_text_color(255, 255, 255)
    pp.set_xy(18, 75)
    pp.fonte("B", 34)
    pp.cell(0, 16, pp.t("Obrigado!"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pp.fonte("", 14)
    pp.cell(0, 10, pp.t(proj.get("contatos") or ""), align="C")
    return bytes(pp.output())


# ------------------------------------------------------------------------------
# 5. EXPORTAÇÃO EM WORD (.docx)
# ------------------------------------------------------------------------------
def gerar_docx(proj: dict, d: dict, res: dict, piv_conc=None) -> bytes | None:
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        return None
    doc = Document()
    st_ = doc.styles["Normal"]
    st_.font.name = "Arial"
    st_.font.size = Pt(11)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.5)
        s.top_margin = s.bottom_margin = Cm(2.5)

    def tabela(df: pd.DataFrame):
        if df is None or df.empty:
            doc.add_paragraph("(sem registros)")
            return
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = "Light Grid Accent 1"
        for i, c in enumerate(df.columns):
            t.rows[0].cells[i].text = str(c)
        for _, r in df.iterrows():
            cells = t.add_row().cells
            for i, c in enumerate(df.columns):
                cells[i].text = str(r[c])
        for row in t.rows:
            for c in row.cells:
                for p_ in c.paragraphs:
                    for run in p_.runs:
                        run.font.size = Pt(8.5)
        doc.add_paragraph()

    p_ = doc.add_paragraph()
    p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = p_.add_run(proj["nome_negocio"])
    r_.bold, r_.font.size = True, Pt(26)
    r_.font.color.rgb = RGBColor(0x1C, 0x31, 0x50)
    p_ = doc.add_paragraph("Plano de Negócio e Análise de Viabilidade Econômico-Financeira")
    p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for linha in [proj["cidade_uf"], proj.get("turma", ""), proj.get("autores", ""), proj.get("professor", ""),
                  proj.get("data_entrega", "") or date.today().strftime("%d/%m/%Y")]:
        if linha:
            q = doc.add_paragraph(str(linha))
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    doc.add_heading("Sumário executivo", 1)
    for par in str(proj.get("sumario_executivo") or gerar_sumario(proj, d, res)).split("\n"):
        doc.add_paragraph(par)
    if any(proj.get(k) for k in ("missao", "visao", "valores")):
        doc.add_heading("Missão, visão e valores", 1)
        for rot, k in (("Missão", "missao"), ("Visão", "visao"), ("Valores", "valores")):
            if proj.get(k):
                doc.add_paragraph(f"{rot}: {proj[k]}")
    ob = d.get("objetivos")
    if ob is not None and not ob.empty:
        doc.add_heading("Objetivos estratégicos SMART", 2)
        tabela(pd.DataFrame({"Objetivo": ob["objetivo"], "Perspectiva": ob["perspectiva"],
                             "Meta": [f"{i_}: {num_br(m, 1)} {u}" for i_, m, u in zip(ob["indicador"], ob["meta"], ob["unidade"])],
                             "Prazo": ob["prazo_meses"].map(lambda v: f"mês {num_br(v, 0)}"),
                             "SMART": ["".join(L if ok else "-" for L, (ok, _) in checar_smart(o).items()) for _, o in ob.iterrows()]}))
    doc.add_heading("1. Análise de mercado", 1)
    doc.add_heading("1.1 Segmentos de clientes", 2)
    seg = d["segmentos"]
    if not seg.empty:
        tabela(seg.assign(Média=seg.iloc[:, 2:7].mean(axis=1).round(1))[["nome", "descricao", "Média"]]
               .rename(columns={"nome": "Segmento", "descricao": "Descrição"}))
    doc.add_heading("1.2 Personas", 2)
    for _, p in d["personas"].iterrows():
        doc.add_paragraph(f"{p['nome']}, {int(p['idade'])} anos, {p['profissao']}", style="List Bullet")
        doc.add_paragraph(f"{p['descricao']} Ganhos: {p['ganhos']} Dores: {p['desafios']}")
    mk = mercado(proj, res)
    if mk["preenchido"]:
        doc.add_heading("1.3 Tamanho do mercado", 2)
        doc.add_paragraph(f"TAM: {num_br(mk['tam_c'], 0)} clientes ({brl(mk['tam'])}/ano). SAM: {num_br(mk['sam_c'], 0)} clientes "
                          f"({brl(mk['sam'])}/ano). SOM: {num_br(mk['som_c'], 0)} clientes ({brl(mk['som'])}/ano). "
                          f"Fontes: {proj.get('mercado_fontes', '')}")
    doc.add_heading("1.4 Concorrência", 2)
    if piv_conc is not None and not piv_conc.empty:
        tabela(piv_conc.round(1))
    doc.add_heading("2. Proposta de valor", 1)
    props = d.get("propostas")
    if props is not None:
        for _, pr in props.iterrows():
            doc.add_paragraph(f"{proj['nome_negocio']} ajuda {pr['persona_nome']}, que deseja {pr['que_desejam']}, por meio de "
                              f"{pr['por_meio']}. Diferentemente de {pr['diferentemente']}, oferecemos {pr['oferecem']}.")
    doc.add_heading("3. Análise estratégica (SWOT)", 1)
    sw = d["swot"]
    tabela(sw[["ambiente", "natureza", "fator", "impacto", "plano_acao"]].rename(columns={
        "ambiente": "Ambiente", "natureza": "Classificação", "fator": "Fator", "impacto": "Impacto", "plano_acao": "Ação"}))
    diag = pontuacao_swot(sw)
    if diag["predominante"]:
        doc.add_paragraph(f"Estratégia predominante: {QUADRANTES[diag['predominante']][0]}. {QUADRANTES[diag['predominante']][2]}")
    doc.add_heading("3.1 Business Model Canvas", 2)
    for k, v in canvas_blocos(proj, d, res).items():
        doc.add_paragraph(f"{k}: " + ("; ".join(v) if v else "(não preenchido)"))
    pt = porter_tabela(d)
    if (pt["Intensidade"] > 0).any():
        doc.add_heading("3.2 Cinco Forças de Porter", 2)
        tabela(pt.assign(Intensidade=pt["Intensidade"].map(lambda v: num_br(v, 0)),
                         Justificativa=pt["Justificativa"].str.replace("(sugestão) ", "", regex=False)))
        doc.add_paragraph(porter_leitura(pt))
    ge_ = ge_tabela(d)
    if not ge_.empty:
        doc.add_heading("3.3 Matriz GE/McKinsey", 2)
        tabela(ge_[["Segmento", "Atratividade", "Força competitiva", "Recomendação"]])
    bcg_ = bcg_tabela(d, res, float(proj.get("bcg_limite_crescimento_pct") or 10))
    if not bcg_.empty:
        doc.add_heading("3.4 Matriz BCG", 2)
        tabela(bcg_.assign(**{"Receita anual": bcg_["Receita anual"].map(brl), "Margem": bcg_["Margem"].map(pct)}))
    if ob is not None and not ob.empty:
        doc.add_heading("3.5 Balanced Scorecard", 2)
        ind = bsc_indicadores_plano(proj, d, res)
        for per in PERSPECTIVAS:
            sub = ob[ob["perspectiva"] == per]
            doc.add_paragraph(f"Perspectiva {per.lower()}: " + ("; ".join(f"{o} (iniciativa: {i_})" for o, i_ in
                                                                         zip(sub["objetivo"], sub["iniciativa"])) or "sem objetivos")
                              + (". Indicadores do plano: " + "; ".join(f"{n} = {v}" for n, v in ind[per]) if ind.get(per) else ""))
    doc.add_heading("4. Plano de marketing", 1)
    for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
        doc.add_paragraph(f"{rot}: {proj.get(k, '')}")
    doc.add_heading("5. Plano operacional", 1)
    for rot, k in [("Localização", "loc_endereco"), ("Justificativa", "loc_justificativa"), ("Layout", "op_layout"),
                   ("Processos", "op_processos"), ("Horário", "op_horario")]:
        if proj.get(k):
            doc.add_paragraph(f"{rot}: {proj[k]}")
    doc.add_heading("6. Plano financeiro", 1)
    doc.add_heading("6.1 Investimento total", 2)
    it = res["investimento_total"]
    tabela(pd.DataFrame({"Item": ["Investimentos fixos", "Pré-operacionais", "Estoque inicial", "Capital de giro", "Total"],
                         "Valor": [brl(x) for x in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)]}))
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        doc.add_heading("6.1b Assinaturas e receita recorrente", 2)
        a_ = d["assinaturas"]
        tabela(pd.DataFrame({"Plano": a_["plano"], "Preço por ciclo": a_["preco_ciclo"].map(brl),
                             "Ciclo (meses)": a_["ciclo_meses"].map(lambda v: num_br(v, 0)),
                             "Novos no Mês 1": a_["novos_mes1"].map(lambda v: num_br(v, 0)),
                             "Churn/mês": a_["churn_mensal_pct"].map(lambda v: pct(v / 100))}))
        doc.add_paragraph(f"MRR no Mês 12 de {brl(rc['mrr12'])} ({num_br(rc['ativos12'], 0)} assinantes), churn médio de "
                          f"{pct(rc['churn'])} ao mês, margem bruta recorrente de {pct(rc['margem_bruta'])}, LTV de {brl(rc['ltv'])}, "
                          f"CAC de {brl(rc['cac'])} e LTV/CAC de {num_br(rc['ltv_cac'], 1)}x.")
    ch = custo_hora(res, d, proj)
    if not ch["tabela"].empty:
        doc.add_heading("6.1c Serviços e custo da hora", 2)
        doc.add_paragraph(f"Custo da hora produtiva de {brl(ch['custo_hora'])}, com {num_br(ch['horas_cobraveis'], 0)} horas "
                          f"cobráveis por mês e ocupação prevista de {pct(ch['ocupacao_m12'])} no Mês 12.")
        t_ = ch["tabela"]
        tabela(pd.DataFrame({"Serviço": t_["Serviço"], "Custo pleno": t_["Custo pleno"].map(brl), "Preço": t_["Preço atual"].map(brl),
                             "Preço sugerido": t_["Preço sugerido"].map(lambda v: brl(v) if v == v else "n.c.")}))
    doc.add_heading("6.2 Demonstração do resultado", 2)
    tabela(tabela_dre(res))
    doc.add_heading("6.3 Indicadores de viabilidade", 2)
    sem = semaforo(res)
    tabela(pd.DataFrame({"Indicador": [s["Indicador"] for s in sem], "Valor": [s["Valor"] for s in sem],
                         "Leitura": [s["Leitura"] for s in sem]}))
    imgs = graficos_png(composicoes(res, d))
    if imgs:
        doc.add_heading("6.4 Gráficos", 2)
        for tit, png in imgs:
            doc.add_picture(io.BytesIO(png), width=Cm(14))
            doc.add_paragraph(f"Figura - {tit}. Fonte: elaborado pelos autores com o ViabilizaLab.").runs[0].italic = True
    doc.add_heading("6.5 Validação, incerteza e defesa", 2)
    preco = preco_referencia(proj, res)
    pl = placar_evidencias(proj, d, res)
    doc.add_paragraph(f"Placar de evidências: {pct(pl['score'], 0)} do resultado apoiado em premissas pesquisadas ou testadas. "
                      f"Premissa a validar primeiro: {pl['proxima'] or 'nenhuma'}.")
    tabela(pl["tabela"].assign(peso=pl["tabela"]["peso"].map(lambda v: pct(v, 0)), status=pl["tabela"]["status"].map(STATUS_CURTO))[
        ["premissa", "valor_plano", "peso", "status", "fonte"]].rename(columns={"premissa": "Premissa", "valor_plano": "Valor no plano",
                                                                              "peso": "Peso", "status": "Origem", "fonte": "Fonte"}))
    ent = resumo_entrevistas(d, preco)
    if ent["n"]:
        doc.add_paragraph(f"Entrevistas: {ent['n']}; pagariam {pct(ent['pct_sim'], 0)} (sim) e {pct(ent['pct_talvez'], 0)} (talvez); "
                          f"dor mais citada: {ent['dor_top'][0]} ({ent['dor_top'][1]} entrevistas).")
    vw = van_westendorp(d.get("precos_vw"))
    if vw and vw.get("ok"):
        doc.add_picture(io.BytesIO(png_vw(vw, preco)), width=Cm(14))
        doc.add_paragraph(f"Figura - Teste de preço Van Westendorp ({vw['n']} respostas): faixa aceitável de {brl(vw['pmc'])} a "
                          f"{brl(vw['pme'])}, preço ótimo {brl(vw['opp'])}. Fonte: elaborado pelos autores com o ViabilizaLab.").runs[0].italic = True
    frase = frase_posicionamento(proj)
    if frase:
        doc.add_paragraph("Posicionamento: " + frase)
    if not d["errc"].empty:
        tabela(d["errc"].rename(columns={"atributo": "Atributo", "acao": "Ação", "nova_nota": "Nova nota", "justificativa": "Justificativa"}))
    oc = orcamento_canais(proj, d, res)
    if not oc.empty:
        tabela(oc.assign(**{c: oc[c].map(lambda x: brl(x) if x == x else "n.a.") for c in ("Gasto mensal", "CAC", "Valor do cliente")},
                         **{"Retorno por R$ 1": oc["Retorno por R$ 1"].map(lambda x: num_br(x, 2) if x == x else "n.a.")}))
    if res["anos"][1]["rec_bruta"] > 0:
        mc = monte_carlo(res, faixas_mc(d))
        doc.add_picture(io.BytesIO(png_hist_mc(mc["vpl"])), width=Cm(14))
        doc.add_paragraph(f"Figura - Distribuição do VPL em {mc['n']} cenários (Monte Carlo). Chance de VPL positivo: "
                          f"{pct(mc['p_vpl_pos'], 0)}. Fonte: elaborado pelos autores com o ViabilizaLab.").runs[0].italic = True
        doc.add_paragraph(leitura_mc(mc))
        pr = avaliar_prontidao(proj, d, res, mc, preco)
        doc.add_paragraph(f"Régua de prontidão: nível {pr['nivel']} de 9.")
    pa = perda_aceitavel(proj, d, res)
    doc.add_paragraph(f"Perda potencial: {brl(pa['potencial'])}; perda aceitável dos sócios: {brl(pa['aceitavel'])}. "
                      f"Enquadramento sugerido: {porte_sugerido(res, d)['porte']}.")
    doc.add_heading("7. Conclusão", 1)
    doc.add_paragraph(str(proj.get("conclusao") or "Escreva aqui a conclusão do grupo sobre a viabilidade do negócio."))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------------------
# 6. MODO PROFESSOR: AVALIAÇÃO DE VÁRIOS PROJETOS
# ------------------------------------------------------------------------------
RUBRICA_PADRAO = [
    ("Cliente e mercado (segmentos, personas, jornada, TAM/SAM/SOM)", 20),
    ("Proposta de valor, concorrência e estratégia (SWOT)", 20),
    ("Marketing e plano operacional", 15),
    ("Plano financeiro completo e consistente", 25),
    ("Viabilidade e interpretação dos indicadores", 10),
    ("Sumário executivo e conclusão", 10),
]
GRUPOS_RUBRICA = [
    ["segmentos", "personas", "jornada", "mercado"],
    ["proposta", "concorrencia", "forcas", "oportunidades", "swot"],
    ["marketing", "canais", "operacional", "fornecedores"],
    ["investimentos", "produtos", "custos"],
    None,
    ["sumario"],
]


def carregar_projeto_arquivo(caminho) -> tuple:
    conn = conectar(caminho)
    proj = ler_unica(conn, "projeto")
    d = {t: ler(conn, t) for t in SCHEMA if t not in ("projeto", "proposta_valor")}
    d["proposta_valor"] = ler_unica(conn, "proposta_valor")
    conn.close()
    return proj, d


def avaliar(proj: dict, d: dict, rubrica=RUBRICA_PADRAO) -> dict:
    res = executar_motor(proj, d)
    chk = tabela_checagens_v5(res, d, proj)
    sem = semaforo(res)
    comp = completude_de(proj, d)
    rx = raio_x(res, chk, comp, sem)
    regras = regras_situacao(proj, d)

    def frac(chaves):
        v = [1 if regras[k][0] else (0.5 if regras[k][1] else 0) for k in chaves]
        return sum(v) / len(v)

    ok_chk = float((chk["Situação"] == "OK").mean()) if len(chk) else 0
    verdes = sum(1 for s in sem if s["Status"] == "🟢") / max(len(sem), 1)
    notas = []
    for (crit, peso), grupo in zip(rubrica, GRUPOS_RUBRICA + [None] * max(0, len(rubrica) - len(GRUPOS_RUBRICA))):
        if grupo is None and "consistente" in crit.lower():
            f = 0.5 * frac(["investimentos", "produtos", "custos"]) + 0.5 * ok_chk
        elif grupo is None:
            f = 0.5 * verdes + 0.5 * (1 if len(str(proj.get("conclusao", ""))) > 150 else 0.3)
        else:
            f = frac(grupo)
            if "consistente" in crit.lower():
                f = 0.5 * f + 0.5 * ok_chk
            if grupo == ["sumario"]:
                f = 0.6 * f + 0.4 * (1 if len(str(proj.get("conclusao", ""))) > 150 else 0)
        notas.append(round(peso * f, 1))
    a1 = res["anos"][1]
    return {"Projeto": proj["nome_negocio"], "Integrantes": str(proj.get("autores", "")).replace("\n", "; "),
            "Turma": proj.get("turma", ""), "Preenchimento": comp, "Checagens OK": ok_chk, "Raio-X": rx["nota"],
            "Investimento": res["investimento_total"], "Receita Ano 1": a1["rec_bruta"], "Lucro Ano 1": a1["resultado"],
            "VPL": res["vpl"], "TIR": res["tir"], "Payback desc. (meses)": res["payback_desc"],
            "Pendências": "; ".join(chk.loc[chk["Situação"] != "OK", "Teste"]), "notas": notas,
            "Nota sugerida": round(sum(notas), 1), "res": res}


def ficha_avaliacao_pdf(avaliacoes: list, rubrica=RUBRICA_PADRAO, professor="", disciplina="") -> bytes:
    pdf = Relatorio("Fichas de avaliação")
    pdf.alias_nb_pages()
    for av in avaliacoes:
        pdf.add_page()
        pdf.n_secao = 0
        pdf.fonte("B", 15)
        pdf.set_text_color(*AZUL)
        pdf.cell(0, 9, pdf.t("FICHA DE AVALIAÇÃO DE PLANO DE NEGÓCIO"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.fonte("", 9)
        pdf.cell(0, 5, pdf.t(f"{disciplina}  ·  Professor(a): {professor}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)
        pdf.campo("Projeto", av["Projeto"])
        pdf.campo("Integrantes", av["Integrantes"])
        pdf.campo("Turma", av["Turma"])
        pdf.campo("Indicadores", f"Investimento {brl(av['Investimento'])} · Receita Ano 1 {brl(av['Receita Ano 1'])} · "
                                 f"Lucro Ano 1 {brl(av['Lucro Ano 1'])} · VPL {brl(av['VPL'])} · "
                                 f"TIR {pct(av['TIR']) if av['TIR'] is not None else 'n.a.'}")
        pdf.campo("Preenchimento / checagens", f"{pct(av['Preenchimento'], 0)} preenchido · {pct(av['Checagens OK'], 0)} das checagens OK · "
                                               f"Raio-X {av['Raio-X']}/100")
        if av["Pendências"]:
            pdf.campo("Pendências apontadas pelo programa", av["Pendências"])
        pdf.ln(2)
        notas_fin = av.get("notas_finais") or av["notas"]
        tab = pd.DataFrame({"Critério": [c for c, _ in rubrica], "Peso": [num_br(p, 0) for _, p in rubrica],
                            "Nota sugerida": [num_br(n, 1) for n in av["notas"]],
                            "Nota da banca": [num_br(n, 1) if av.get("notas_finais") else "" for n in notas_fin]})
        tab.loc[len(tab)] = ["TOTAL", num_br(sum(p for _, p in rubrica), 0), num_br(sum(av["notas"]), 1),
                             num_br(sum(notas_fin), 1) if av.get("notas_finais") else ""]
        pdf.tabela(tab, larguras=(110, 15, 27, 28), alinhar=("LEFT", "CENTER", "CENTER", "CENTER"), destaque_total=True)
        pdf.subtitulo("Comentários da banca")
        y = pdf.get_y() + 2
        for i in range(8):
            pdf.line(pdf.l_margin, y + i * 8, pdf.w - pdf.r_margin, y + i * 8)
        pdf.set_y(y + 70)
        pdf.cell(0, 6, pdf.t("Assinatura: ____________________________________    Data: ____/____/______"),
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    return bytes(pdf.output())
