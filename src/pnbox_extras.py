"""
================================================================================
PNBOX ACADÊMICO - RECURSOS EXTRAS (pnbox_extras.py)
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
                         tabela_checagens_v5, raio_x, executar_motor)
from pnbox_db import SCHEMA, conectar, ler, ler_unica
from pnbox_pdf import Relatorio, graficos_png, AZUL, FONTES_CANDIDATAS


# ------------------------------------------------------------------------------
# 1. SITUAÇÃO DE PREENCHIMENTO DE CADA FERRAMENTA
# ------------------------------------------------------------------------------
def regras_situacao(proj: dict, d: dict) -> dict:
    def tem(t, n=1):
        return len(d[t]) >= n
    sw = d["swot"]
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
        "preco": (tem("catalogo"), tem("catalogo")),
        "dre": (tem("catalogo") and tem("custos_fixos"), tem("catalogo")),
        "fluxo": (tem("catalogo") and tem("capital_giro_prazos"), tem("catalogo")),
        "indicadores": (tem("catalogo") and tem("inv_fixos"), tem("catalogo")),
        "graficos": (tem("catalogo"), tem("catalogo")),
        "simulador": (tem("catalogo"), tem("catalogo")),
        "experimentos": (tem("experimentos"), tem("experimentos")),
        "riscos": (tem("riscos", 3), tem("riscos")),
        "cronograma": (tem("cronograma", 3), tem("cronograma")),
        "realizado": (tem("realizado"), tem("realizado")),
        "sumario": (len(str(proj["sumario_executivo"])) > 200, bool(proj["sumario_executivo"])),
    }


def completude_de(proj, d) -> float:
    r = regras_situacao(proj, d)
    # "Orçado x realizado" só faz sentido depois da abertura: não entra no cálculo
    vals = [1 if c else (0.5 if i else 0) for k, (c, i) in r.items() if k != "realizado"]
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
    p.append(f"{nome} é um empreendimento do setor de {proj['setor_atuacao'] or 'atuação a definir'}, localizado em "
             f"{proj['cidade_uf'] or 'cidade a definir'}, com o objetivo de {str(proj['objetivo']).lower()}. "
             f"O enquadramento tributário previsto é o {proj['enquadramento']} e o porte, {proj['porte']}.")
    if seg_top or valores:
        p.append(f"O segmento de clientes priorizado é {seg_top or 'a definir'}. A proposta de valor se apoia em "
                 f"{valores or 'benefícios a definir'}" + (f", com destaque para {', '.join(top)}." if top else "."))
    if mk["preenchido"]:
        p.append(f"O mercado alcançável foi estimado em {num_br(mk['sam_c'], 0)} clientes ({brl(mk['sam'])} por ano), e a meta "
                 f"de conquista (SOM) em {num_br(mk['som_c'], 0)} clientes. A receita prevista para o primeiro ano representa "
                 f"{pct(mk['pct_som'])} desse SOM.")
    p.append(f"A análise SWOT indica a {estrategia.lower()} como postura predominante.")
    p.append(f"O investimento total necessário é de {brl(res['investimento_total'])}, composto por {brl(res['inv_fixo'])} em "
             f"investimentos fixos, {brl(res['inv_pre'])} em gastos pré-operacionais, {brl(res['inv_estoque'])} em estoque inicial "
             f"e {brl(res['capital_giro'])} de capital de giro. As fontes são {brl(res['capital_proprio'])} de capital próprio"
             + (f" ({socios})" if socios else "") + f" e {brl(res['capital_terceiros'])} de capital de terceiros.")
    p.append(f"Para o primeiro ano, projeta-se receita bruta de {brl(a1['rec_bruta'])}, margem de contribuição de {pct(a1['imc'])} "
             f"e resultado líquido de {brl(a1['resultado'])} (lucratividade de {pct(div(a1['resultado'], a1['rec_bruta']))}). "
             f"O ponto de equilíbrio contábil é de {brl(a1['pe_contabil'])} por ano.")
    p.append(f"Na avaliação do investimento em três anos, à TMA de {pct(res['tma'])} ao ano, o VPL é de {brl(res['vpl'])}"
             + (f", a TIR é de {pct(tir)}" if tir is not None else "")
             + (f" e o payback descontado ocorre em {num_br(pb, 1)} meses." if pb == pb else " e o investimento não se paga em 36 meses."))
    p.append("Com base nesses resultados, o negócio se mostra " + ("economicamente viável" if viavel else
             "ainda não viável nas premissas atuais, sendo necessário rever preços, custos ou investimentos") + ".")
    return "\n".join(p)


# ------------------------------------------------------------------------------
# 3. TEXTO PARA ANÁLISE COM IA (copiar e colar no ChatGPT, Gemini, Claude...)
# ------------------------------------------------------------------------------
TAREFAS_IA = {
    "indicadores": "Diagnóstico de cada indicador: situação atual, o que significa para a saúde do negócio e sinais de alerta",
    "concorrencia": "Estratégias para se diferenciar e superar os concorrentes",
    "custos": "Plano de redução de custos e de melhoria de margens",
    "preco": "Revisão da política de preços por produto",
    "marketing": "Plano de marketing e de aquisição de clientes para os primeiros 6 meses",
    "riscos": "Principais riscos do plano e como mitigá-los",
    "graficos": "Gráficos interpretativos dos principais números",
    "banca": "Simulação de banca examinadora: perguntas difíceis sobre o plano",
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
    seg = d["segmentos"]
    if not seg.empty:
        s_ = seg.assign(Média=seg.iloc[:, 2:7].mean(axis=1).round(1)).sort_values("Média", ascending=False)
        b.append("\n## 2. Segmentos de clientes (notas 1 a 10)")
        b.append(_md_tabela(s_[["nome", "descricao", "Média"]].rename(columns={"nome": "Segmento", "descricao": "Descrição"})))
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
    if any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
        b.append("\n## 7. Marketing (4Ps)")
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            b.append(f"- {rot}: {proj.get(k, '')}")
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
    b.append("\n## 10. Estrutura de custos mensal")
    b.append(f"- Custos diretos {brl(res['cdir_m1'])}; impostos {pct(res['aliq'])} e despesas variáveis {pct(res['var_pct'])} da receita; "
             f"MOD {brl(res['mod_m'])}; custos indiretos de produção {brl(res['cif_m'])}; depreciação {brl(res['dep_m'])}; "
             f"pessoal adm./comercial {brl(res['pessoal_adm_m'] + res['pessoal_com_m'])}; despesas adm. {brl(res['desp_adm_m'])}, "
             f"comerciais {brl(res['desp_com_m'])}, financeiras {brl(res['desp_fin_m'])}.")
    cf = d["custos_fixos"]
    if not cf.empty:
        b.append(_md_tabela(pd.DataFrame({"Custo fixo": cf["descricao"], "Classificação": cf["classificacao"],
                                          "R$/mês": cf["valor_mensal"].map(brl)})))
    b.append("\n## 11. DRE (Anos 1 a 3)")
    b.append(_md_tabela(tabela_dre(res)))
    b.append("\n## 12. Indicadores e semáforo")
    b.append(_md_tabela(pd.DataFrame({"Indicador": [s_["Indicador"] for s_ in sem], "Valor": [s_["Valor"] for s_ in sem],
                                      "Situação": [{"🟢": "bom", "🟡": "atenção", "🔴": "crítico"}[s_["Status"]] for s_ in sem]})))
    b.append(f"- VPL {brl(res['vpl'])}; TIR {pct(res['tir']) if res['tir'] is not None else 'n.a.'}; TMA {pct(res['tma'])}; "
             f"payback descontado {num_br(res['payback_desc'], 1) + ' meses' if res['payback_desc'] == res['payback_desc'] else '> 36 meses'}; "
             f"PE contábil {brl(a1['pe_contabil'])}/ano; margem de segurança {pct(div(a1['rec_bruta'] - a1['pe_contabil'], a1['rec_bruta']))}.")
    sens = sensibilidade(res)
    b.append("- Sensibilidade (efeito de ±10% no resultado do Ano 1): " +
             "; ".join(f"{r['Variável']} {brl(r['Impacto absoluto'])}" for _, r in sens.iterrows()))
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
        "indicadores": "Para CADA indicador da seção 12, explique em linguagem simples: (a) o que ele mede; (b) a situação atual deste "
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
                    "(3) receita x custos x resultado nos 3 anos; (4) ponto de equilíbrio; (5) saldo de caixa mês a mês. "
                    "Abaixo de cada gráfico, escreva 2 ou 3 frases de interpretação.",
        "banca": "Faça o papel de uma banca examinadora: elabore 10 perguntas difíceis sobre este plano (mercado, números, premissas, "
                 "riscos) e, para cada uma, indique o que uma boa resposta deveria conter.",
    }
    for i, t in enumerate([t for t in tarefas if t in pedidos], start=1):
        b.append(f"{i}. {pedidos[t]}")
    b.append("\nOrganize a resposta com títulos, use tabelas quando ajudar e termine com uma lista de próximos passos.")
    return "\n".join(b)


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
    pp.numeros([("VPL (3 anos)", brl(res["vpl"])), ("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a."),
                ("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "> 36 meses"),
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
            doc.add_paragraph(f"Figura - {tit}. Fonte: elaborado pelos autores com o PNBOX Acadêmico.").runs[0].italic = True
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
