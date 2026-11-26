"""
================================================================================
VIABILIZALAB - RELATÓRIO ACADÊMICO NO PADRÃO ABNT (pnbox_abnt.py)
================================================================================
Capa, folha de rosto, sumário com páginas (NBR 6027), seções numeradas (NBR 6024),
tabelas com título e fonte, citações autor-data (NBR 10520) e referências (NBR 6023).
Margens de 3 cm (superior e esquerda) e 2 cm (inferior e direita), fonte 12, espaçamento 1,5.
"""
from datetime import date

import pandas as pd
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace

from pnbox_pdf import Relatorio
from pnbox_motor import brl, num_br, pct, div, tabela_dre, tabela_indicadores, cenarios
from pnbox_extras import gerar_sumario
from pnbox_validacao import placar_evidencias, resumo_entrevistas, monte_carlo, faixas_mc, preco_referencia
from pnbox_v8 import QUADRANTES_CRUZADA, comparar_setor, referencias_abnt_dados, valuation
from pnbox_mix import CITA_MKT, REFS_MKT, conteudo_plano_mkt

ENTRELINHA = 6.4      # 12 pt com espaçamento 1,5
RECUO = " " * 8  # recuo de primeira linha do parágrafo

REFS_METODO = {
    "swot": "WRIGHT, Peter; KROLL, Mark J.; PARNELL, John. **Administração estratégica**: conceitos. São Paulo: Atlas, 2000.",
    "canvas": "OSTERWALDER, Alexander; PIGNEUR, Yves. **Business model generation**: inovação em modelos de negócios. "
              "Rio de Janeiro: Alta Books, 2011.",
    "sebrae": "SERVIÇO BRASILEIRO DE APOIO ÀS MICRO E PEQUENAS EMPRESAS. **Aplicações de análise de sobrevivência em estudos do "
              "Sebrae**. Brasília: Sebrae, 2026. Disponível em: https://datasebrae.com.br. Acesso em: {hoje}.",
}


REFS_METODO.update(REFS_MKT)


class RelatorioABNT(Relatorio):
    def __init__(self, nome):
        super().__init__(nome)
        self.set_margins(30, 30, 20)
        self.set_auto_page_break(auto=True, margin=20)
        self.numerar_a_partir = 10 ** 6
        self.n_tab = 0
        self.toc = []

    def header(self):
        if self.page_no() >= self.numerar_a_partir:
            self.set_y(20)
            self.fonte("", 10)
            self.cell(0, 5, str(self.page_no() - 1), align="R")
            self.set_y(30)

    def footer(self):
        pass

    def secao_abnt(self, numero: str, titulo: str, nivel: int = 1, nova_pagina=False):
        if nova_pagina:
            self.add_page()
        elif self.get_y() > 245:
            self.add_page()
        self.ln(4 if nivel == 1 else 2)
        self.fonte("B" if nivel <= 2 else "", 12)
        txt = f"{numero} {titulo.upper() if nivel == 1 else titulo}".strip()
        self.multi_cell(0, ENTRELINHA, self.t(txt), align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)
        self.toc.append((nivel, numero, titulo.upper() if nivel == 1 else titulo, self.page_no() - 1))

    def texto(self, t: str):
        for par in [p.strip() for p in str(t or "").split("\n") if p.strip()]:
            self.fonte("", 12)
            self.multi_cell(0, ENTRELINHA, self.t(RECUO + par), align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def tabela_abnt(self, titulo: str, df: pd.DataFrame, larguras=None, alinhar=None, fonte="Elaborada pelos autores (2026).",
                    tam=10):
        if df is None or df.empty:
            return
        self.n_tab += 1
        if self.get_y() > 230:
            self.add_page()
        self.ln(2)
        self.fonte("", 10)
        self.multi_cell(0, 5, self.t(f"Tabela {self.n_tab} - {titulo}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        cols = list(df.columns)
        if alinhar is None:
            alinhar = ["LEFT"] + ["RIGHT"] * (len(cols) - 1)
        self.fonte("", tam)
        with self.table(col_widths=larguras, text_align=alinhar, line_height=tam * 0.5, borders_layout="HORIZONTAL_LINES",
                        headings_style=FontFace(emphasis="BOLD"), first_row_as_headings=True) as tb:
            r = tb.row()
            for c in cols:
                r.cell(self.t(c))
            for _, linha in df.iterrows():
                r = tb.row()
                for c in cols:
                    r.cell(self.t(linha[c]))
        self.fonte("", 10)
        self.multi_cell(0, 5, self.t(f"Fonte: {fonte}"), align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(3)


def _capa(pdf, proj, autores, titulo):
    pdf.add_page()
    pdf.fonte("B", 12)
    for linha in [proj.get("abnt_instituicao") or "INSTITUIÇÃO DE ENSINO", proj.get("abnt_curso") or ""]:
        if linha:
            pdf.multi_cell(0, ENTRELINHA, pdf.t(str(linha).upper()), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(18)
    pdf.fonte("", 12)
    for a in autores:
        pdf.multi_cell(0, ENTRELINHA, pdf.t(a.upper()), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_y(120)
    pdf.fonte("B", 14)
    pdf.multi_cell(0, 8, pdf.t(titulo.upper()), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_y(250)
    pdf.fonte("", 12)
    pdf.multi_cell(0, ENTRELINHA, pdf.t(proj.get("abnt_cidade") or proj.get("cidade_uf") or ""), align="C",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, ENTRELINHA, pdf.t(proj.get("abnt_ano") or str(date.today().year)), align="C",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT)


def _folha_rosto(pdf, proj, autores, titulo):
    pdf.add_page()
    pdf.fonte("", 12)
    for a in autores:
        pdf.multi_cell(0, ENTRELINHA, pdf.t(a.upper()), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_y(110)
    pdf.fonte("B", 14)
    pdf.multi_cell(0, 8, pdf.t(titulo.upper()), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(14)
    disc = proj.get("abnt_disciplina") or "disciplina"
    natureza = (f"Plano de negócio apresentado à {disc}" + (f" do curso de {proj['abnt_curso']}" if proj.get("abnt_curso") else "")
                + (f" da {proj['abnt_instituicao']}" if proj.get("abnt_instituicao") else "") + ", como requisito parcial para aprovação.")
    if proj.get("abnt_orientador"):
        natureza += f"\n\nOrientador(a): {proj['abnt_orientador']}"
    pdf.fonte("", 10)
    pdf.set_x(30 + 80)
    pdf.multi_cell(80, 5, pdf.t(natureza), align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_y(250)
    pdf.fonte("", 12)
    pdf.multi_cell(0, ENTRELINHA, pdf.t(proj.get("abnt_cidade") or proj.get("cidade_uf") or ""), align="C",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, ENTRELINHA, pdf.t(proj.get("abnt_ano") or str(date.today().year)), align="C",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT)


def _sumario(pdf, toc):
    pdf.add_page()
    pdf.fonte("B", 12)
    pdf.cell(0, ENTRELINHA, "SUMÁRIO" if pdf.unicode else "SUMARIO", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)
    w = pdf.w - pdf.l_margin - pdf.r_margin
    for nivel, num, tit, pag in toc:
        pdf.fonte("B" if nivel == 1 else "", 12)
        esq = pdf.t(f"{num} {tit}".strip())
        while pdf.get_string_width(esq) > w - 25 and len(esq) > 5:
            esq = esq[:-2]
        y = pdf.get_y()
        pdf.set_xy(pdf.l_margin, y)
        pdf.cell(pdf.get_string_width(esq) + 1, ENTRELINHA, esq)
        pdf.fonte("", 12)
        x_ini = pdf.get_x()
        num_w = pdf.get_string_width(str(pag)) + 1
        largura_pontos = w - (x_ini - pdf.l_margin) - num_w - 1
        n_pontos = max(int(largura_pontos / max(pdf.get_string_width("."), 0.1)), 0)
        pdf.cell(largura_pontos, ENTRELINHA, "." * n_pontos)
        pdf.set_xy(pdf.l_margin + w - num_w, y)
        pdf.cell(num_w, ENTRELINHA, str(pag), align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)


def _conteudo(pdf, proj, d, res, refs_extra: list) -> list:
    """Escreve as seções textuais. Retorna as referências metodológicas usadas."""
    usadas = []
    a1 = res["anos"][1]
    N = res["N"]
    pdf.numerar_a_partir = pdf.page_no() + 1
    # 1 INTRODUÇÃO
    pdf.secao_abnt("1", "Introdução", nova_pagina=True)
    pdf.texto(proj.get("sumario_executivo") or gerar_sumario(proj, d, res))
    # 2 APRESENTAÇÃO
    pdf.secao_abnt("2", "Apresentação do empreendimento")
    pdf.texto(f"O empreendimento {proj.get('nome_negocio') or ''} será instalado em {proj.get('cidade_uf') or 'local a definir'}, "
              f"no setor de {str(proj.get('setor_atuacao') or 'atuação a definir').lower()}, com enquadramento previsto no "
              f"{proj.get('enquadramento') or 'regime tributário a definir'}.")
    if any(proj.get(k) for k in ("missao", "visao", "valores")):
        pdf.secao_abnt("2.1", "Missão, visão e valores", 2)
        pdf.texto(f"Missão: {proj.get('missao') or '-'}\nVisão: {proj.get('visao') or '-'}\nValores: {proj.get('valores') or '-'}")
    soc = d["socios"]
    if not soc.empty:
        pdf.secao_abnt("2.2", "Quadro societário", 2)
        pdf.tabela_abnt("Quadro societário e capital investido", pd.DataFrame({
            "Sócio": soc["nome"], "Função": soc["funcao"], "Capital investido": soc["capital_integralizado"].map(brl)}),
            larguras=(60, 60, 40), alinhar=("LEFT", "LEFT", "RIGHT"))
    # 3 MERCADO
    pdf.secao_abnt("3", "Análise de mercado")
    seg = d["segmentos"]
    if not seg.empty:
        pdf.secao_abnt("3.1", "Segmentos de clientes", 2)
        pdf.tabela_abnt("Segmentos de clientes avaliados", pd.DataFrame({"Segmento": seg["nome"], "Descrição": seg["descricao"]}),
                        larguras=(50, 110), alinhar=("LEFT", "LEFT"))
    pdf.secao_abnt("3.2", "Tamanho do mercado", 2)
    tam, sam = float(proj.get("tam_clientes") or 0), float(proj.get("sam_clientes") or 0)
    txt = (f"O mercado total foi estimado em {num_br(tam, 0)} clientes potenciais e o mercado atendível em {num_br(sam, 0)}, "
           f"com meta de participação de {num_br(proj.get('som_pct') or 0, 1)}%.")
    dp = d.get("dados_publicos", pd.DataFrame())
    if not dp.empty:
        m = {r["chave"]: r for _, r in dp.iterrows()}
        if "populacao" in m:
            txt += (f" Segundo o Censo Demográfico, o município de {proj.get('municipio_nome') or ''} tinha "
                    f"{num_br(m['populacao']['valor'], 0)} habitantes em 2022 (IBGE, 2023)")
            txt += (f", distribuídos em {num_br(m['domicilios']['valor'], 0)} domicílios ocupados." if "domicilios" in m else ".")
        pofs = [r for k, r in m.items() if str(k).startswith("pof")]
        for r in pofs[:2]:
            txt += (f" De acordo com a Pesquisa de Orçamentos Familiares, a despesa média mensal das famílias com "
                    f"{str(r['rotulo']).lower()} era de {brl(r['valor'])} em 2017-2018 (IBGE, 2019).")
    pdf.texto(txt)
    conc = d["concorrentes"]
    if not conc.empty:
        pdf.secao_abnt("3.3", "Concorrência", 2)
        notas = d["concorrencia_notas"]
        if not notas.empty:
            piv = notas.pivot_table(index="atributo", columns="empresa", values="nota", aggfunc="first").reset_index()
            piv = piv.rename(columns={"atributo": "Atributo"})
            for c in piv.columns[1:]:
                piv[c] = piv[c].map(lambda v: num_br(v, 0))
            w = 160 / len(piv.columns)
            pdf.tabela_abnt("Notas atribuídas à empresa e aos concorrentes (0 a 10)", piv, larguras=[w] * len(piv.columns))
    ent = resumo_entrevistas(d, preco_referencia(proj, res))
    pl = placar_evidencias(proj, d, res)
    pdf.secao_abnt("3.4", "Validação das premissas", 2)
    t = f"O placar de evidências indica que {pct(pl['score'], 0)} do resultado projetado se apoia em premissas pesquisadas ou testadas."
    if ent["n"]:
        t += (f" Foram realizadas {ent['n']} entrevistas com potenciais clientes, das quais {pct(ent['pct_sim'], 0)} indicaram que "
              f"pagariam pelo produto ou serviço.")
    pdf.texto(t)
    # 4 ESTRATÉGIA
    pdf.secao_abnt("4", "Estratégia")
    sw = d["swot"]
    if not sw.empty:
        pdf.secao_abnt("4.1", "Análise SWOT", 2)
        pdf.texto("A análise SWOT confronta os fatores internos (forças e fraquezas) com os externos (oportunidades e ameaças) e deve "
                  "subsidiar diretamente a formulação estratégica (Wright; Kroll; Parnell, 2000).")
        usadas.append("swot")
        pdf.tabela_abnt("Fatores da análise SWOT", pd.DataFrame({"Natureza": sw["natureza"], "Fator": sw["fator"],
                                                                 "Impacto": sw["impacto"]}), larguras=(30, 105, 25),
                        alinhar=("LEFT", "LEFT", "CENTER"))
    est = d.get("swot_estrategias", pd.DataFrame())
    if not est.empty:
        pdf.secao_abnt("4.2", "Estratégias da SWOT cruzada", 2)
        pdf.tabela_abnt("Estratégias derivadas do cruzamento dos fatores", pd.DataFrame({
            "Tipo": est["quadrante"].str.split(" ").str[0], "Estratégia": est["estrategia"], "Ação": est["acao"]}),
            larguras=(28, 66, 66), alinhar=("LEFT", "LEFT", "LEFT"))
    # 5 MARKETING
    pdf.secao_abnt("5", "Plano de marketing")
    pdf.texto("O plano de marketing segue a sequência análise da situação, objetivos, estratégia, táticas, ações e controle "
              "(Smith, 2011; McDonald, 2007), também adotada para micro e pequenas empresas por Las Casas (2011).")
    usadas += ["smith", "mcdonald", "lascasas"]
    for i, sec in enumerate(conteudo_plano_mkt(proj, d, res), 1):
        pdf.secao_abnt(f"5.{i}", sec["titulo"], 2)
        cits = [CITA_MKT[r] for r in sec["refs"] if r not in ("smith", "mcdonald")]
        cit = f"({'; '.join(cits)})" if cits else ""
        pdf.texto(sec["texto"].rstrip(".") + f" {cit}." if cit else sec["texto"])
        usadas += sec["refs"]
        tb = sec["tabela"]
        if tb is not None and not tb.empty:
            n = len(tb.columns)
            larg = {2: (40, 120), 4: (70, 40, 30, 20), 5: (45, 25, 25, 40, 25)}.get(n)
            if n == 4 and "Mês" in tb.columns:
                larg = (15, 85, 35, 25)
            pdf.tabela_abnt(sec["titulo"], tb, larguras=larg, alinhar=tuple(["LEFT"] * n), tam=9)
    # 6 OPERACIONAL
    pdf.secao_abnt("6", "Plano operacional")
    pdf.texto(proj.get("op_processos") or "Processos a descrever.")
    pes = d["pessoal"]
    if not pes.empty:
        pdf.tabela_abnt("Equipe prevista", pd.DataFrame({"Cargo": pes["cargo"], "Pessoas": pes["num_empregados"].map(lambda v: num_br(v, 0)),
                                                         "Salário": pes["salario_mensal"].map(brl),
                                                         "Início (mês)": pes["mes_inicio"].map(lambda v: num_br(v, 0))}),
                        larguras=(70, 25, 35, 30))
    # 7 FINANCEIRO
    pdf.secao_abnt("7", "Plano financeiro")
    pdf.secao_abnt("7.1", "Investimento e fontes", 2)
    pdf.tabela_abnt("Investimento total", pd.DataFrame({
        "Item": ["Investimentos fixos", "Investimentos pré-operacionais", "Capital de giro", "Total"],
        "Valor": [brl(res["inv_fixo"]), brl(res["inv_pre"]), brl(res["capital_giro"]), brl(res["investimento_total"])]}),
        larguras=(110, 50))
    pdf.secao_abnt("7.2", "Demonstração do resultado", 2)
    dre = tabela_dre(res)
    cols = list(dre.columns)[: min(len(dre.columns), 5)]
    pdf.tabela_abnt(f"Demonstração do resultado projetada ({N} ano{'s' if N > 1 else ''})", dre[cols],
                    larguras=[160 - 30 * (len(cols) - 1)] + [30] * (len(cols) - 1), tam=8)
    pdf.secao_abnt("7.3", "Indicadores de viabilidade", 2)
    ind = tabela_indicadores(res)
    ci = ["Indicador", "Ano 1"] + ([f"Ano {N}"] if N > 1 and f"Ano {N}" in ind.columns else [])
    pdf.tabela_abnt("Indicadores de viabilidade", ind[ci], larguras=[90] + [70 / (len(ci) - 1)] * (len(ci) - 1), tam=9)
    pdf.texto(f"Com taxa mínima de atratividade de {pct(res['tma'])} ao ano, o valor presente líquido é de {brl(res['vpl'])}"
              + (f" e a taxa interna de retorno, de {pct(res['tir'])}" if res["tir"] is not None and res["tir"] == res["tir"] else "")
              + f". O payback descontado é {res['payback_desc']}.")
    if a1["rec_bruta"] > 0:
        mc = monte_carlo(res, faixas_mc(d))
        pdf.texto(f"Na simulação de Monte Carlo, com {mc['n']} cenários, a probabilidade de VPL positivo é de "
                  f"{pct(mc['p_vpl_pos'], 0)}, e em 5% dos cenários o VPL fica abaixo de {brl(mc['p5'])}.")
        cen = cenarios(res, proj)
        if cen is not None and not cen.empty:
            cen = cen.copy()
            for c in cen.columns[1:]:
                cen[c] = cen[c].astype(str)
            pdf.tabela_abnt("Cenários", cen.iloc[:, :4], larguras=[55] + [35] * 3, tam=9)
    if proj.get("setor_referencia"):
        try:
            cs = comparar_setor(proj["setor_referencia"], proj, d, res)
            fonte = ("IBGE (2026), Pesquisa Anual de " + ("Comércio" if proj["setor_referencia"].startswith("Comércio") else "Serviços")
                     + " 2024; elaborada pelos autores.")
            pdf.tabela_abnt(f"Comparação com a média do setor: {proj['setor_referencia'].split(' · ')[1]}", pd.DataFrame({
                "Indicador": cs["Indicador"],
                "Plano": [pct(v) if "÷" in i or "margem" in i.lower() else brl(v) for i, v in zip(cs["Indicador"], cs["Plano"])],
                "Setor": [pct(v) if "÷" in i or "margem" in i.lower() else brl(v) for i, v in zip(cs["Indicador"], cs["Setor (Brasil)"])]}),
                larguras=(90, 35, 35), fonte=fonte)
        except Exception:
            pass
    vl = valuation(res, proj, d)
    pdf.texto(f"Pelo método do fluxo de caixa descontado, o patrimônio dos sócios foi estimado em {brl(vl['equity_dcf'])}; "
              f"os métodos por múltiplos resultam em valores entre {brl(vl['faixa'][0])} e {brl(vl['faixa'][1])}.")
    # 8 RISCOS
    ri = d["riscos"]
    pdf.secao_abnt("8", "Riscos e contingência")
    pdf.texto("De acordo com o Sebrae (2026), apenas 41,5% dos microempreendedores individuais e 76% das micro e pequenas empresas "
              "continuam ativos após cinco anos, o que reforça a importância do planejamento de riscos.")
    usadas.append("sebrae")
    if not ri.empty:
        pdf.tabela_abnt("Matriz de riscos", pd.DataFrame({"Risco": ri["risco"], "Prob.": ri["probabilidade"].map(lambda v: num_br(v, 0)),
                                                          "Impacto": ri["impacto"].map(lambda v: num_br(v, 0)), "Mitigação": ri["mitigacao"]}),
                        larguras=(52, 16, 22, 70), alinhar=("LEFT", "CENTER", "CENTER", "LEFT"))
    # 9 CONSIDERAÇÕES
    pdf.secao_abnt("9", "Considerações finais")
    pdf.texto(proj.get("conclusao") or "Considerações finais a redigir pelo grupo.")
    return usadas


def gerar_relatorio_abnt(proj: dict, d: dict, res: dict) -> bytes:
    autores = [a.strip() for a in str(proj.get("autores") or "").replace(",", ";").split(";") if a.strip()] or ["Autor(a)"]
    titulo = f"Plano de negócio: {proj.get('nome_negocio') or 'empreendimento'}"
    hoje = date.today()
    meses = ["jan.", "fev.", "mar.", "abr.", "maio", "jun.", "jul.", "ago.", "set.", "out.", "nov.", "dez."]
    acesso = f"{hoje.day} {meses[hoje.month - 1]} {hoje.year}"
    toc_final = None
    paginas_sumario = 1
    for _ in range(2):   # 1ª passagem: descobre as páginas das seções; 2ª: escreve o sumário com os números certos
        pdf = RelatorioABNT(proj.get("nome_negocio") or "")
        _capa(pdf, proj, autores, titulo)
        _folha_rosto(pdf, proj, autores, titulo)
        if toc_final is not None:
            _sumario(pdf, toc_final)
            while pdf.page_no() < 2 + paginas_sumario:
                pdf.add_page()
        else:
            pdf.add_page()   # reserva de uma página para o sumário na primeira passagem
        usadas = _conteudo(pdf, proj, d, res, [])
        # referências
        pdf.add_page()
        pdf.fonte("B", 12)
        pdf.cell(0, ENTRELINHA, "REFERÊNCIAS" if pdf.unicode else "REFERENCIAS", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.toc.append((1, "", "REFERÊNCIAS", pdf.page_no() - 1))
        pdf.ln(4)
        refs = [str(r) for r in d.get("referencias", pd.DataFrame(columns=["referencia"]))["referencia"] if str(r).strip()]
        refs += [REFS_METODO[u].format(hoje=acesso) for u in dict.fromkeys(usadas)]
        refs += referencias_abnt_dados(proj, d, acesso)
        for r in sorted(set(refs), key=lambda s: s.upper()):
            pdf.fonte("", 12)
            pdf.multi_cell(0, 5, pdf.t(r), align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
            pdf.ln(5)
        if toc_final is None:
            toc_final = pdf.toc
            tmp = RelatorioABNT("")
            _sumario(tmp, toc_final)
            paginas_sumario = tmp.page_no()
            if paginas_sumario > 1:   # sumário maior que uma página desloca as seções
                toc_final = [(n, a, t, p + paginas_sumario - 1) for n, a, t, p in toc_final]
    return bytes(pdf.output())
