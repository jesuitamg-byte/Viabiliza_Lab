"""
================================================================================
PNBOX ACADÊMICO - RELATÓRIO EM PDF (pnbox_pdf.py)
================================================================================
Gera o relatório completo do plano de negócio, na ordem do PNBOX/Sebrae.
Requer fpdf2 >= 2.7 (pip install -U fpdf2).
================================================================================
"""

import io
import re
from pathlib import Path

import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace

AZUL = (28, 49, 80)
def _fontes_matplotlib():
    try:
        import matplotlib
        base = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
        return [(str(base / "DejaVuSans.ttf"), str(base / "DejaVuSans-Bold.ttf"), str(base / "DejaVuSans-Oblique.ttf"))]
    except Exception:
        return []


FONTES_CANDIDATAS = [
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/ariali.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf", "/Library/Fonts/Arial Italic.ttf"),
] + _fontes_matplotlib()
TROCAS = {"–": "-", "—": "-", "−": "-", "×": "x", "÷": "/", "Σ": "Soma", "“": '"', "”": '"',
          "’": "'", "‘": "'", "…": "...", "•": "-", "≥": ">=", "≤": "<="}


class Relatorio(FPDF):
    def __init__(self, nome):
        super().__init__(orientation="P", unit="mm", format="A4")
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
        self.set_auto_page_break(auto=True, margin=15)
        self.set_margins(15, 15, 15)

    def t(self, v) -> str:
        s = "" if v is None else str(v)
        if s == "nan":
            s = ""
        for a, b in TROCAS.items():
            s = s.replace(a, b)
        if not self.unicode:
            s = s.encode("latin-1", "replace").decode("latin-1")
        return s

    def fonte(self, estilo="", tam=9):
        self.set_font(self.fam, estilo, tam)

    def footer(self):
        self.set_y(-11)
        self.fonte("I", 7)
        self.set_text_color(120, 120, 120)
        self.cell(0, 5, self.t(f"Plano de Negócio - {self.nome}"), new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.cell(0, 5, f"Pág. {self.page_no()}/{{nb}}", align="R")
        self.set_text_color(0, 0, 0)

    # --- blocos de conteúdo -------------------------------------------------
    def secao(self, texto, nova_pagina=False):
        self.n_secao = getattr(self, "n_secao", 0) + 1
        texto = f"{self.n_secao}. {texto}"
        if nova_pagina:
            self.add_page()
        elif self.get_y() > 245:
            self.add_page()
        self.ln(3)
        self.set_fill_color(*AZUL)
        self.set_text_color(255, 255, 255)
        self.fonte("B", 11)
        self.cell(0, 8, self.t(f"  {texto}"), fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def subtitulo(self, texto):
        if self.get_y() > 260:
            self.add_page()
        self.ln(1)
        self.fonte("B", 10)
        self.set_text_color(*AZUL)
        self.cell(0, 6, self.t(texto), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(0, 0, 0)

    def paragrafo(self, texto, estilo="", tam=9):
        if not str(texto).strip():
            return
        self.fonte(estilo, tam)
        self.multi_cell(0, 4.8, self.t(texto), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(1)

    def campo(self, rotulo, valor):
        if not str(valor).strip():
            return
        self.set_x(self.l_margin)
        self.fonte("B", 9)
        self.write(4.8, self.t(f"{rotulo}: "))
        self.fonte("", 9)
        self.write(4.8, self.t(valor))
        self.ln(5.5)

    def tabela(self, df: pd.DataFrame, larguras=None, alinhar=None, tam=8, destaque_total=False):
        if df is None or df.empty:
            self.paragrafo("(sem registros)", "I", 8)
            return
        self.fonte("", tam)
        self.set_fill_color(255, 255, 255)
        cols = list(df.columns)
        if alinhar is None:
            alinhar = ["LEFT"] + ["RIGHT"] * (len(cols) - 1)
        cab = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(70, 95, 130))
        with self.table(col_widths=larguras, text_align=alinhar, line_height=tam * 0.55,
                        headings_style=cab, first_row_as_headings=True,
                        cell_fill_color=(242, 245, 250), cell_fill_mode="ROWS") as tb:
            linha = tb.row()
            for c in cols:
                linha.cell(self.t(c))
            n = len(df)
            for i, (_, r) in enumerate(df.iterrows()):
                linha = tb.row()
                estilo = FontFace(emphasis="BOLD") if destaque_total and i == n - 1 else None
                for c in cols:
                    linha.cell(self.t(r[c]), style=estilo)
        self.ln(2)

    def matriz_swot(self, sw: pd.DataFrame):
        """Desenha a matriz SWOT 2x2 colorida (Forças, Fraquezas, Oportunidades, Ameaças)."""
        if self.get_y() > 150:
            self.add_page()
        x0, y0 = self.l_margin, self.get_y() + 2
        lat, cab, w, h, gap = 10, 9, 84, 78, 1.5
        cores = {"Força": ((38, 84, 140), (255, 255, 255)), "Fraqueza": ((214, 0, 110), (255, 255, 255)),
                 "Oportunidade": ((226, 230, 235), (40, 50, 60)), "Ameaça": ((92, 102, 114), (255, 255, 255))}
        titulos = {"Força": "FORÇAS", "Fraqueza": "FRAQUEZAS", "Oportunidade": "OPORTUNIDADES", "Ameaça": "AMEAÇAS"}
        # cabeçalhos
        self.fonte("B", 9)
        self.set_text_color(255, 255, 255)
        self.set_fill_color(38, 84, 140)
        self.set_xy(x0 + lat + gap, y0)
        self.cell(w, cab, self.t("FATORES POSITIVOS"), fill=True, align="C")
        self.set_fill_color(214, 0, 110)
        self.set_xy(x0 + lat + gap * 2 + w, y0)
        self.cell(w, cab, self.t("FATORES NEGATIVOS"), fill=True, align="C")
        for j, (rot, cor) in enumerate([("INTERNOS", (38, 84, 140)), ("EXTERNOS", (200, 204, 210))]):
            yy = y0 + cab + gap + j * (h + gap)
            self.set_fill_color(*cor)
            self.rect(x0, yy, lat, h, style="F")
            self.set_text_color(255, 255, 255) if j == 0 else self.set_text_color(40, 50, 60)
            self.fonte("B", 9)
            y_txt = yy + (h - len(rot) * 5) / 2
            for i, letra in enumerate(self.t(rot)):
                self.set_xy(x0, y_txt + i * 5)
                self.cell(lat, 5, letra, align="C")
        pos = {"Força": (0, 0), "Fraqueza": (1, 0), "Oportunidade": (0, 1), "Ameaça": (1, 1)}
        for nat, (c, l) in pos.items():
            xx = x0 + lat + gap + c * (w + gap)
            yy = y0 + cab + gap + l * (h + gap)
            fundo, texto = cores[nat]
            self.set_fill_color(*fundo)
            self.rect(xx, yy, w, h, style="F")
            self.set_text_color(*texto)
            self.set_xy(xx + 3, yy + 3)
            self.fonte("B", 11)
            self.cell(w - 6, 6, self.t(titulos[nat]))
            itens = sw[sw["natureza"] == nat]
            self.set_xy(xx + 3, yy + 10)
            self.fonte("", 7.5)
            linhas = []
            for _, r in itens.iterrows():
                marca = "*" if int(r.get("priorizado", 0) or 0) else "-"
                linhas.append(f"{marca} {r['fator']} ({r['impacto']})")
            txt = "\n".join(linhas) if linhas else "(nenhum fator cadastrado)"
            self.multi_cell(w - 6, 3.8, self.t(txt), new_x=XPos.LEFT, new_y=YPos.NEXT, max_line_height=3.8)
        self.set_text_color(0, 0, 0)
        self.set_xy(x0, y0 + cab + 2 * h + 3 * gap + 1)
        self.fonte("I", 7)
        self.cell(0, 4, self.t("* fator priorizado. Entre parênteses, o impacto (Alto/Baixo)."),
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)


# ==============================================================================
# GRÁFICOS (matplotlib) PARA O PDF
# ==============================================================================
PALETA = ["#26548c", "#d6006e", "#5c6672", "#3a8fd1", "#f39c12", "#27ae60", "#8e44ad", "#c0392b",
          "#16a085", "#7f8c8d", "#2c3e50", "#e67e22"]


def _fmt_rs(v):
    return ("R$ " + f"{v:,.0f}".replace(",", "."))


def _fmt_curto(v, _=None):
    a = abs(v)
    if a >= 1e6:
        return f"R$ {v/1e6:.1f} mi".replace(".", ",")
    if a >= 1e3:
        return f"R$ {v/1e3:.0f} mil"
    return f"R$ {v:.0f}"


def graficos_png(comp: dict) -> list:
    """Devolve [(título, bytes PNG)]. Se o matplotlib não estiver instalado, devolve lista vazia."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return []
    import io as _io
    plt.rcParams.update({"font.size": 9, "font.family": "DejaVu Sans"})
    saida = []

    def salvar(fig, titulo):
        b = _io.BytesIO()
        fig.savefig(b, format="png", dpi=170, bbox_inches="tight")
        plt.close(fig)
        saida.append((titulo, b.getvalue()))

    def rosca(df, titulo):
        if df.empty:
            return
        fig, ax = plt.subplots(figsize=(5.2, 4.4))
        cores = PALETA[: len(df)] if len(df) <= len(PALETA) else None
        cunhas, _, autot = ax.pie(df["Valor"], colors=cores, startangle=90, counterclock=False,
                                  autopct=lambda p: f"{p:.1f}%".replace(".", ",") if p >= 4 else "",
                                  pctdistance=0.78, wedgeprops=dict(width=0.42, edgecolor="white"))
        for t in autot:
            t.set_color("white"); t.set_fontsize(8); t.set_fontweight("bold")
        ax.text(0, 0, _fmt_rs(df["Valor"].sum()), ha="center", va="center", fontsize=9, fontweight="bold")
        ax.legend(cunhas, [f"{i} ({p*100:.1f}%)".replace(".", ",") for i, p in zip(df["Item"], df["%"])],
                  loc="upper center", bbox_to_anchor=(0.5, -0.02), ncol=1, frameon=False, fontsize=8)
        ax.set_title(titulo, fontsize=10, fontweight="bold")
        ax.axis("equal")
        salvar(fig, titulo)

    def barras(df, titulo):
        if df.empty:
            return
        df = df.iloc[::-1]
        fig, ax = plt.subplots(figsize=(7.5, 0.38 * len(df) + 1.0))
        ax.barh([str(i)[:55] for i in df["Item"]], df["Valor"], color="#26548c")
        for y, (v, p) in enumerate(zip(df["Valor"], df["%"])):
            ax.text(v, y, f"  {p*100:.1f}%".replace(".", ","), va="center", fontsize=8)
        ax.set_title(titulo, fontsize=10, fontweight="bold", loc="left")
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_curto)); ax.xaxis.set_major_locator(matplotlib.ticker.MaxNLocator(5))
        ax.spines[["top", "right"]].set_visible(False)
        ax.set_xlim(0, df["Valor"].max() * 1.15)
        salvar(fig, titulo)

    rosca(comp["investimento"], "Composição do investimento total")
    rosca(comp["fontes"], "Fontes de recursos")
    rosca(comp["estoque"], "Estoque inicial por categoria")
    rosca(comp["custo_direto"], "Custo direto mensal por categoria")
    rosca(comp["receita_categoria"], "Receita por categoria")
    barras(comp["inv_fixos"], "Maiores itens do ativo imobilizado")
    barras(comp["top_produtos"], "Produtos/serviços com maior receita mensal")
    barras(comp["estrutura_receita"], "Custos, despesas e lucro em % da receita (Ano 1)")

    eq = comp["equilibrio"]
    pe, rec = comp["pe_ponto"]
    fig, ax = plt.subplots(figsize=(7.5, 3.8))
    ax.plot(eq["Receita"], eq["Receita total"], color="#26548c", lw=2.5, label="Receita total")
    ax.plot(eq["Receita"], eq["Custo total"], color="#d6006e", lw=2.5, label="Custo total")
    ax.plot(eq["Receita"], eq["Custos fixos"], color="#5c6672", lw=1.5, ls="--", label="Custos fixos")
    ax.axvline(pe, color="black", ls=":", lw=1)
    ax.annotate(f"PE: {_fmt_rs(pe)}", (pe, pe), xytext=(10, -25), textcoords="offset points", fontsize=8)
    ax.axvline(rec, color="#27ae60", ls=":", lw=1)
    ax.annotate(f"Receita projetada: {_fmt_rs(rec)}", (rec, rec), xytext=(-150, 15), textcoords="offset points", fontsize=8)
    fmt = matplotlib.ticker.FuncFormatter(_fmt_curto)
    ax.xaxis.set_major_formatter(fmt); ax.yaxis.set_major_formatter(fmt); ax.xaxis.set_major_locator(matplotlib.ticker.MaxNLocator(6))
    ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Ponto de equilíbrio contábil (Ano 1)", fontsize=10, fontweight="bold", loc="left")
    salvar(fig, "Ponto de equilíbrio")

    fa = comp["fluxo_acumulado"]
    fig, ax = plt.subplots(figsize=(7.5, 3.5))
    x = range(len(fa))
    ax.bar([i - 0.2 for i in x], fa["Acumulado"], 0.4, color="#26548c", label="Acumulado")
    ax.bar([i + 0.2 for i in x], fa["Acumulado descontado (TMA)"], 0.4, color="#d6006e", label="Acumulado descontado (TMA)")
    ax.axhline(0, color="black", lw=0.8)
    ax.set_xticks(list(x), fa["Ano"]); ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_curto))
    ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
    ax.set_title("Fluxo de caixa acumulado do projeto (payback)", fontsize=10, fontweight="bold", loc="left")
    salvar(fig, "Payback")
    return saida


# ==============================================================================
# MONTAGEM DO RELATÓRIO
# ==============================================================================
def gerar_pdf(proj, d, res, extras) -> bytes:
    """extras: dict com dre, indicadores, cenarios, checagens, swot_diag, quadrantes."""
    from pnbox_motor import brl, num_br, pct, div, QUADRANTES, mercado, semaforo, capacidade_demanda, custo_hora  # noqa (removido no arquivo único)
    pdf = Relatorio(proj["nome_negocio"])
    pdf.alias_nb_pages()
    pdf.add_page()

    # ---------------- CAPA ----------------
    pdf.ln(50)
    pdf.fonte("B", 24)
    pdf.set_text_color(*AZUL)
    pdf.multi_cell(0, 12, pdf.t(proj["nome_negocio"]), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.fonte("", 14)
    pdf.cell(0, 9, pdf.t("Plano de Negócio e Análise de Viabilidade Econômico-Financeira"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(15)
    pdf.fonte("", 10)
    for rot, k in [("Objetivo", "objetivo"), ("Cidade", "cidade_uf"), ("Setor de atuação", "setor_atuacao"),
                   ("Enquadramento tributário", "enquadramento"), ("Porte", "porte"), ("Contatos", "contatos")]:
        pdf.cell(0, 6, pdf.t(f"{rot}: {proj[k]}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 6, pdf.t(f"Período de planejamento: {int(proj['periodo_anos'])} ano(s)"), align="C",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    autores = [a.strip() for a in re.split(r"[;\n]", proj.get("autores") or "") if a.strip()]
    if autores or proj.get("turma") or proj.get("professor"):
        pdf.ln(12)
        pdf.fonte("B", 10)
        pdf.cell(0, 6, pdf.t("Elaborado por"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.fonte("", 10)
        for a in autores:
            pdf.cell(0, 5.5, pdf.t(a), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        for rot, k in [("Turma", "turma"), ("Professor(a)", "professor"), ("Entrega", "data_entrega")]:
            if proj.get(k):
                pdf.cell(0, 5.5, pdf.t(f"{rot}: {proj[k]}"), align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if str(proj.get("sumario_executivo") or "").strip():
        pdf.secao("SUMÁRIO EXECUTIVO", nova_pagina=True)
        for par in str(proj["sumario_executivo"]).split("\n"):
            pdf.paragrafo(par, "", 10)

    # ---------------- 1. SÓCIOS ----------------
    pdf.secao("INFORMAÇÕES DO PLANO E QUADRO SOCIETÁRIO", nova_pagina=True)
    s = d["socios"]
    pdf.tabela(pd.DataFrame({
        "Sócio": s["nome"], "Função": s["funcao"], "Minicurrículo": s["perfil"],
        "Capital (R$)": s["capital_integralizado"].map(brl)}), larguras=(40, 34, 76, 30),
        alinhar=("LEFT", "LEFT", "LEFT", "RIGHT"))

    # ---------------- 2. SEGMENTOS ----------------
    pdf.secao("SEGMENTAÇÃO DE MERCADO")
    sg = d["segmentos"].copy()
    notas = ["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]
    sg["media"] = sg[notas].mean(axis=1)
    sg = sg.sort_values("media", ascending=False)
    pdf.tabela(pd.DataFrame({
        "Segmento": sg["nome"], "Descrição": sg["descricao"],
        "Tam.": sg["nota_tamanho"].map(lambda v: num_br(v, 0)), "Lucro": sg["nota_lucro"].map(lambda v: num_br(v, 0)),
        "Long.": sg["nota_longevidade"].map(lambda v: num_br(v, 0)),
        "Conc.": sg["nota_concorrencia"].map(lambda v: num_br(v, 0)),
        "Ader.": sg["nota_aderencia"].map(lambda v: num_br(v, 0)), "Média": sg["media"].map(num_br)}),
        larguras=(40, 70, 11, 11, 11, 11, 11, 15), alinhar=("LEFT", "LEFT") + ("CENTER",) * 6)
    pdf.paragrafo("Notas de 1 a 10: tamanho do segmento, potencial de lucro, longevidade, pouca concorrência "
                  "e aderência da oferta às demandas do segmento.", "I", 7.5)
    mk = mercado(proj, res)
    if mk["preenchido"]:
        pdf.subtitulo("Tamanho do mercado (TAM, SAM e SOM)")
        pdf.tabela(pd.DataFrame({
            "Nível": ["TAM - mercado total", "SAM - mercado alcançável", "SOM - mercado conquistável"],
            "Descrição": [proj.get("tam_desc", ""), proj.get("sam_desc", ""), f"{num_br(proj.get('som_pct', 0), 1)}% do SAM"],
            "Clientes": [num_br(mk["tam_c"], 0), num_br(mk["sam_c"], 0), num_br(mk["som_c"], 0)],
            "R$ por ano": [brl(mk["tam"]), brl(mk["sam"]), brl(mk["som"])]}),
            larguras=(42, 78, 25, 35), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT"))
        pdf.paragrafo(f"A receita prevista no Ano 1 ({brl(mk['receita_ano1'])}) corresponde a {pct(mk['pct_som'])} do SOM "
                      f"e a {pct(mk['pct_sam'])} do SAM. Fontes: {proj.get('mercado_fontes', '')}", "I", 8)

    # ---------------- 3. PROPOSTA DE VALOR ----------------
    pdf.secao("PROPOSTA DE VALOR")
    props = d.get("propostas")
    if props is not None and not props.empty:
        for _, pr in props.iterrows():
            pdf.subtitulo(f"Persona: {pr['persona_nome']}")
            if pr["valores"]:
                pdf.campo("Valores entregues", pr["valores"])
            pdf.paragrafo(f"{proj['nome_negocio']} ajuda {pr['persona_nome']}, que deseja {pr['que_desejam']}, por meio de "
                          f"{pr['por_meio']}. Diferentemente de {pr['diferentemente']}, oferecemos {pr['oferecem']}.", "", 10)
    pv = d["proposta_valor"]
    if (props is None or props.empty) and pv.get("segmento_alvo"):
        pdf.paragrafo(f"Nossa empresa ({proj['nome_negocio']}) ajuda {pv['segmento_alvo']} que desejam "
                      f"{pv['beneficios_desejados']} por meio de {pv['meios_entrega']}. Diferentemente de "
                      f"{pv['diferencial_competitivo']}, os nossos produtos ou serviços oferecem "
                      f"{pv['entrega_produtos_servicos']}", "", 10)

    # ---------------- 4. PERSONAS ----------------
    pdf.secao("PERSONAS", nova_pagina=True)
    for _, p in d["personas"].iterrows():
        pdf.subtitulo(f"{p['nome']}, {int(p['idade'])} anos - {p['profissao']}")
        pdf.campo("Estado civil", p["estado_civil"])
        pdf.campo("Renda mensal", brl(p["renda_mensal"]) if p["renda_mensal"] else "")
        pdf.campo("Escolaridade", p["escolaridade"])
        pdf.campo("Cidade", p["cidade"])
        pdf.paragrafo(p["descricao"])
        if str(p["citacao"]).strip():
            pdf.paragrafo(f'"{p["citacao"]}"', "I")
        for rot, k in [("Hábitos e comportamentos", "habitos"), ("Características", "caracteristicas"),
                       ("Fluência digital", "fluencia_digital"), ("Redes sociais", "redes_sociais"),
                       ("Quem influencia", "quem_influencia")]:
            pdf.campo(rot, p[k])
        rot_df = pd.DataFrame({"Manhã": [p["rotina_manha"]], "Tarde": [p["rotina_tarde"]], "Noite": [p["rotina_noite"]]})
        pdf.tabela(rot_df, alinhar=("LEFT",) * 3)
        gd = pd.DataFrame({"Ganhos": [p["ganhos"]], "Desafios": [p["desafios"]]})
        pdf.tabela(gd, alinhar=("LEFT",) * 2)
        rf = pd.DataFrame({"Dúvidas ou objeções": [p["duvidas_objecoes"]],
                           "Aborrecimentos a evitar": [p["aborrecimentos"]],
                           "Experiência ideal e como o negócio ajuda": [p["experiencia_ideal"]]})
        pdf.tabela(rf, alinhar=("LEFT",) * 3)

    # ---------------- 5. JORNADA ----------------
    pdf.secao("JORNADA DO CLIENTE")
    j = d["jornada_cliente"]
    for persona, grupo in j.groupby("persona_nome", sort=False):
        pdf.subtitulo(f"Persona: {persona}")
        pdf.tabela(pd.DataFrame({"Fase": grupo["macro_fase"], "Etapa": grupo["etapa"],
                                 "Ações do cliente": grupo["acoes_cliente"], "Pensamentos": grupo["pensamentos"],
                                 "Pontos de contato": grupo["pontos_contato"], "Emoção": grupo["estado_emocional"],
                                 "Oportunidades": grupo["oportunidades"]}),
                   larguras=(14, 22, 34, 26, 28, 20, 36), alinhar=("LEFT",) * 7, tam=7)

    # ---------------- 6. CONCORRÊNCIA ----------------
    pdf.secao("ANÁLISE DA CONCORRÊNCIA", nova_pagina=True)
    piv = extras["concorrencia_pivot"]
    if not piv.empty:
        tab = piv.copy()
        media = tab.drop(columns=["Atributo"]).mean()
        tab = pd.concat([tab, pd.DataFrame([{"Atributo": "Média geral", **media.to_dict()}])], ignore_index=True)
        for c in tab.columns[1:]:
            tab[c] = tab[c].map(lambda v: num_br(v, 1))
        n = len(tab.columns) - 1
        pdf.tabela(tab, larguras=(50,) + (130 / n,) * n, alinhar=("LEFT",) + ("CENTER",) * n,
                   destaque_total=True, tam=7.5)
        nossa = piv.set_index("Atributo")
        if "Nossa Empresa" in nossa.columns and len(nossa.columns) > 1:
            outros = nossa.drop(columns=["Nossa Empresa"]).max(axis=1)
            fortes = nossa.index[nossa["Nossa Empresa"] > outros].tolist()
            fracos = nossa.index[nossa["Nossa Empresa"] < outros].tolist()
            pdf.campo("Forças frente à concorrência", ", ".join(fortes) or "nenhuma")
            pdf.campo("Fraquezas frente à concorrência", ", ".join(fracos) or "nenhuma")
    else:
        pdf.paragrafo("(sem registros)", "I", 8)

    # ---------------- 7. AMBIENTE E SWOT ----------------
    pdf.secao("ANÁLISE DE AMBIENTE E MATRIZ SWOT", nova_pagina=True)
    sw = d["swot"]
    pdf.matriz_swot(sw)
    diag = extras["swot_diag"]
    q = diag["quadrantes"]
    pdf.subtitulo("Matriz de estratégias (pontuação: impacto alto = 2; baixo = 1)")
    pdf.tabela(pd.DataFrame({
        "Quadrante": [f"{QUADRANTES[k][0]} ({QUADRANTES[k][1]})" for k in q],
        "Pontuação": [str(v) for v in q.values()]}), larguras=(150, 30))
    if diag["predominante"]:
        nome, comb, txt = QUADRANTES[diag["predominante"]]
        pdf.paragrafo(f"Situação do negócio: {nome} ({comb}). {txt}", "B")
    pdf.subtitulo("Detalhamento dos fatores")
    pdf.tabela(pd.DataFrame({"Ambiente": sw["ambiente"], "Dimensão": sw["dimensao"], "Fator": sw["fator"],
                             "Classif.": sw["natureza"] + " / " + sw["impacto"], "Diagnóstico": sw["diagnostico"],
                             "Plano de ação": sw["plano_acao"]}),
               larguras=(16, 20, 30, 24, 45, 45), alinhar=("LEFT",) * 6, tam=7)

    # ---------------- 8. CANAIS E FUNIL ----------------
    pdf.secao("CANAIS DE AQUISIÇÃO E FUNIL DE VENDAS", nova_pagina=True)
    cn = d["canais"].copy()
    if not cn.empty:
        cn["media"] = cn[["atratividade", "alcance", "cliente", "relacionamento", "viabilidade"]].mean(axis=1)
        cn = cn.sort_values("media", ascending=False)
        pdf.tabela(pd.DataFrame({"Canal": cn["canal"], "Atratividade": cn["atratividade"].map(lambda v: num_br(v, 0)),
                                 "Alcance": cn["alcance"].map(lambda v: num_br(v, 0)),
                                 "Cliente": cn["cliente"].map(lambda v: num_br(v, 0)),
                                 "Relacionamento": cn["relacionamento"].map(lambda v: num_br(v, 0)),
                                 "Viabilidade": cn["viabilidade"].map(lambda v: num_br(v, 0)),
                                 "Média": cn["media"].map(lambda v: num_br(v, 1))}),
                   larguras=(52, 22, 18, 18, 26, 22, 22), alinhar=("LEFT",) + ("CENTER",) * 6)
        pdf.campo("Canais de tração", ", ".join(cn["canal"].head(3)))
    fu = extras["funil"]
    pdf.subtitulo("Funil de vendas")
    pdf.tabela(fu, alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"),
               larguras=(22, 50, 22, 20, 20, 22, 24), tam=7.5)

    # ---------------- 9. INVESTIMENTOS ----------------
    # ---------------- MARKETING E OPERAÇÃO ----------------
    if any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
        pdf.secao("PLANO DE MARKETING (4Ps)")
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            if proj.get(k):
                pdf.subtitulo(rot)
                pdf.paragrafo(proj[k])
    pdf.secao("PLANO OPERACIONAL", nova_pagina=True)
    for rot, k in [("Localização", "loc_endereco"), ("Justificativa do ponto", "loc_justificativa"),
                   ("Fluxo de pessoas", "loc_fluxo"), ("Horário de funcionamento", "op_horario")]:
        pdf.campo(rot, proj.get(k, ""))
    if proj.get("loc_area_m2"):
        pdf.campo("Área", f"{num_br(proj['loc_area_m2'], 0)} m²")
    if proj.get("op_layout"):
        pdf.subtitulo("Layout")
        pdf.paragrafo(proj["op_layout"])
    if proj.get("op_processos"):
        pdf.subtitulo("Processos operacionais")
        pdf.paragrafo(proj["op_processos"])
    cd = capacidade_demanda(d, res)
    if not cd.empty:
        pdf.subtitulo("Capacidade produtiva x demanda prevista")
        pdf.tabela(pd.DataFrame({"Recurso": cd["Recurso"], "Atende": cd["Atende"],
                                 "Capacidade/mês": cd["Capacidade/mês"].map(lambda v: num_br(v, 0)),
                                 "Demanda Mês 1": cd["Demanda Mês 1"].map(lambda v: num_br(v, 0)),
                                 "Demanda Mês 12": cd["Demanda Mês 12"].map(lambda v: num_br(v, 0)),
                                 "Uso Mês 12": cd["Uso Mês 12"].map(pct)}), larguras=(48, 42, 24, 22, 22, 22), tam=7.5)
    pes = d["pessoal"]
    if not pes.empty:
        pdf.subtitulo("Necessidade de pessoal por área")
        pdf.tabela(pd.DataFrame({"Área": pes["area"], "Cargo": pes["cargo"], "Nº": pes["num_empregados"].map(lambda v: num_br(v, 0)),
                                 "Atribuições": pes["atribuicoes"]}), larguras=(38, 52, 10, 80),
                   alinhar=("LEFT", "LEFT", "CENTER", "LEFT"), tam=7.5)
    fo = d.get("fornecedores")
    if fo is not None and not fo.empty:
        pdf.subtitulo("Fornecedores")
        pdf.tabela(pd.DataFrame({"Fornecedor": fo["nome"], "Itens": fo["itens"],
                                 "Pagamento": fo["prazo_pagamento_dias"].map(lambda v: f"{num_br(v, 0)} dias"),
                                 "Entrega": fo["prazo_entrega_dias"].map(lambda v: f"{num_br(v, 0)} dias"),
                                 "Condições": fo["condicoes"], "Crítico": fo["critico"].map(lambda v: "Sim" if v else "Não")}),
                   larguras=(36, 46, 20, 18, 46, 14), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "LEFT", "CENTER"), tam=7.5)

    pdf.secao("INVESTIMENTOS, CAPITAL DE GIRO E FONTES DE RECURSOS", nova_pagina=True)
    inv = res["inv"]
    pdf.subtitulo("Investimentos fixos e depreciação")
    tinv = pd.DataFrame({"Descrição": inv["descricao"], "Qtd.": inv["quantidade"].map(lambda v: num_br(v, 0)),
                         "Valor unit.": inv["valor_unitario"].map(brl), "Total": inv["total"].map(brl),
                         "Vida útil": inv["vida_util_anos"].map(lambda v: f"{num_br(v, 0)} anos"),
                         "Deprec. anual": inv["dep_anual"].map(brl)})
    tinv.loc[len(tinv)] = ["Total", "", "", brl(res["inv_fixo"]), "", brl(res["dep_anual"])]
    pdf.tabela(tinv, larguras=(62, 12, 26, 28, 20, 32), destaque_total=True, tam=7.5)
    pdf.subtitulo("Investimentos pré-operacionais")
    tp = pd.DataFrame({"Descrição": d["inv_pre"]["descricao"], "Valor": d["inv_pre"]["valor"].map(brl)})
    tp.loc[len(tp)] = ["Total", brl(res["inv_pre"])]
    pdf.tabela(tp, larguras=(140, 40), destaque_total=True)
    pdf.subtitulo("Estoque inicial")
    m = d["materiais"]
    m = m[m["qtd_estoque_inicial"] > 0]
    rv = res["cat"][(res["cat"]["categoria"] == "Revenda") & (res["cat"]["qtd_estoque_inicial"] > 0)]
    te = pd.DataFrame({"Item": list(m["nome"]) + list(rv["descricao"]),
                       "Tipo": ["Material/insumo"] * len(m) + ["Mercadoria p/ revenda"] * len(rv),
                       "Qtd.": [num_br(v, 1) for v in list(m["qtd_estoque_inicial"]) + list(rv["qtd_estoque_inicial"])],
                       "Valor unit.": [brl(v) for v in list(m["valor_unitario"]) + list(rv["custo_unit"])],
                       "Total": [brl(a * b) for a, b in zip(list(m["qtd_estoque_inicial"]) + list(rv["qtd_estoque_inicial"]),
                                                            list(m["valor_unitario"]) + list(rv["custo_unit"]))]})
    te.loc[len(te)] = ["Total", "", "", "", brl(res["inv_estoque"])]
    pdf.tabela(te, larguras=(70, 36, 18, 26, 30), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT"),
               destaque_total=True, tam=7.5)
    pdf.subtitulo("Capital de giro")
    pdf.tabela(pd.DataFrame({
        "Item": ["Prazo médio de recebimento (PMR)", "Prazo médio de pagamento (PMP)", "Prazo médio de estoques (PME)",
                 "Ciclo operacional (PME + PMR)", "Ciclo financeiro (PME + PMR - PMP)",
                 "Desembolso médio mensal (Mês 1)", "Necessidade de capital de giro (NCG)",
                 "Reserva financeira", "Capital de giro total"],
        "Valor": [f"{num_br(res['pmr'], 1)} dias", f"{num_br(res['pmp'], 1)} dias", f"{num_br(res['pme'], 1)} dias",
                  f"{num_br(res['ciclo_op'], 1)} dias", f"{num_br(res['ciclo_fin'], 1)} dias",
                  brl(res["desembolso_m1"]), brl(res["ncg"]), brl(res["reserva"]), brl(res["capital_giro"])]}),
        larguras=(130, 50), destaque_total=True)
    pdf.subtitulo("Investimento total e fontes de recursos")
    it = res["investimento_total"]
    pdf.tabela(pd.DataFrame({
        "Aplicação de recursos": ["Investimentos fixos", "Investimentos pré-operacionais", "Estoque inicial",
                                  "Capital de giro", "Investimento total"],
        "Valor": [brl(res["inv_fixo"]), brl(res["inv_pre"]), brl(res["inv_estoque"]), brl(res["capital_giro"]), brl(it)],
        "%": [pct(div(v, it)) for v in (res["inv_fixo"], res["inv_pre"], res["inv_estoque"], res["capital_giro"], it)]}),
        larguras=(110, 45, 25), destaque_total=True)
    fontes = [(r["nome"], r["capital_integralizado"], "Próprio") for _, r in d["socios"].iterrows()] + \
             [(r["descricao"], r["valor"], f"Terceiros ({num_br(r['taxa_anual_pct'])}% a.a.; {int(r['prazo_meses'])} meses)")
              for _, r in d["emprestimos"].iterrows()]
    tf = pd.DataFrame({"Fonte": [f[0] for f in fontes], "Tipo": [f[2] for f in fontes],
                       "Valor": [brl(f[1]) for f in fontes], "%": [pct(div(f[1], res["fontes"])) for f in fontes]})
    tf.loc[len(tf)] = ["Total das fontes", "", brl(res["fontes"]), "100,00%"]
    pdf.tabela(tf, larguras=(65, 60, 35, 20), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT"), destaque_total=True)

    # ---------------- 10. PRODUTOS ----------------
    pdf.secao("PRODUTOS, SERVIÇOS E CUSTOS UNITÁRIOS", nova_pagina=True)
    c = res["cat"]
    for categoria in ("Fabricação Própria", "Revenda", "Serviço"):
        sub = c[c["categoria"] == categoria]
        if sub.empty:
            continue
        pdf.subtitulo(categoria)
        tc = pd.DataFrame({"Produto/serviço": sub["descricao"], "Preço": sub["preco_venda"].map(brl),
                           "Custo unit.": sub["custo_unit"].map(brl), "Margem": sub["margem_pct"].map(pct),
                           "Markup": sub["markup"].map(lambda v: f"{num_br(v)}x" if v == v else "n.a."),
                           "Qtd./mês": sub["qtd_mensal"].map(lambda v: num_br(v, 0)),
                           "Receita/mês": sub["receita_mensal"].map(brl)})
        pdf.tabela(tc, larguras=(58, 20, 20, 17, 15, 17, 33), tam=7.5)
    pdf.subtitulo("Fichas técnicas (composição do custo dos produtos)")
    ft = d["ficha_tecnica"]
    precos = dict(zip(d["materiais"]["nome"], d["materiais"]["valor_unitario"]))
    unids = dict(zip(d["materiais"]["nome"], d["materiais"]["unidade"]))
    for prod, g in ft.groupby("produto", sort=False):
        preco_v = c.loc[c["descricao"] == prod, "preco_venda"]
        rot = f"{prod}" + (f" - preço de venda {brl(preco_v.iloc[0])}" if len(preco_v) else " (produto não está no catálogo)")
        if pdf.get_y() > 240:
            pdf.add_page()
        pdf.fonte("B", 8.5)
        pdf.cell(0, 5, pdf.t(rot), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        vt = [q * precos.get(mm, 0) for mm, q in zip(g["material"], g["quantidade"])]
        tot = sum(vt)
        tf = pd.DataFrame({"Insumo": g["material"], "Quant.": g["quantidade"].map(lambda v: num_br(v, 3)),
                           "Unid.": [unids.get(mm, "?") for mm in g["material"]],
                           "Valor unit.": [brl(precos.get(mm, 0)) for mm in g["material"]],
                           "Total": [brl(v) for v in vt], "%": [pct(div(v, tot)) for v in vt]})
        tf.loc[len(tf)] = ["Custo unitário", "", "", "", brl(tot), "100,00%"]
        pdf.tabela(tf, larguras=(70, 20, 18, 26, 26, 20), destaque_total=True, tam=7.5)

    # ---------------- 11. CUSTOS ----------------
    pdf.secao("CUSTOS DE COMERCIALIZAÇÃO, PESSOAL E CUSTOS FIXOS", nova_pagina=True)
    a1 = res["anos"][1]
    pdf.subtitulo("Custos variáveis de comercialização (% sobre a receita)")
    cv = d["custos_variaveis"]
    tcv = pd.DataFrame({"Descrição": [f"Impostos sobre vendas ({proj['enquadramento']})"] + list(cv["descricao"]),
                        "%": [num_br(proj["aliquota_imposto_pct"])] + [num_br(v) for v in cv["percentual"]],
                        "Valor no Mês 1": [brl(res["receita_m1"] * res["aliq"])] +
                                          [brl(res["receita_m1"] * v / 100) for v in cv["percentual"]]})
    tcv.loc[len(tcv)] = ["Total", num_br((res["aliq"] + res["var_pct"]) * 100),
                         brl(res["receita_m1"] * (res["aliq"] + res["var_pct"]))]
    pdf.tabela(tcv, larguras=(110, 25, 45), destaque_total=True)
    pdf.subtitulo("Mão de obra")
    pe = d["pessoal"]
    sal = pe["num_empregados"] * pe["salario_mensal"]
    tpe = pd.DataFrame({"Cargo": pe["cargo"], "Nº": pe["num_empregados"].map(lambda v: num_br(v, 0)),
                        "Salário": pe["salario_mensal"].map(brl), "Encargos %": pe["encargos_pct"].map(num_br),
                        "Benefícios/func.": pe["beneficios_unit"].map(brl),
                        "Total mensal": (sal * (1 + pe["encargos_pct"] / 100) + pe["num_empregados"] * pe["beneficios_unit"]).map(brl)})
    tpe.loc[len(tpe)] = ["Total", num_br(pe["num_empregados"].sum(), 0), "", "", "", brl(res["pessoal_m"])]
    pdf.tabela(tpe, larguras=(62, 10, 25, 20, 28, 35), destaque_total=True, tam=7.5)
    pdf.subtitulo("Custos fixos mensais")
    tcf = pd.DataFrame({"Descrição": list(d["custos_fixos"]["descricao"]) +
                        ["Salários", "Encargos sociais", "Benefícios", "Depreciação"],
                        "Valor mensal": [brl(v) for v in d["custos_fixos"]["valor_mensal"]] +
                        [brl(res["salarios_m"]), brl(res["encargos_m"]), brl(res["beneficios_m"]), brl(res["dep_m"])]})
    tot_cf = res["fixos_m"] + res["pessoal_m"] + res["dep_m"]
    tcf.loc[len(tcf)] = ["Total", brl(tot_cf)]
    pdf.tabela(tcf, larguras=(130, 50), destaque_total=True)

    # ---------------- 12. PROJEÇÕES E DRE ----------------
    pdf.subtitulo("Classificação dos custos e despesas (mensal)")
    pdf.tabela(pd.DataFrame({
        "Classificação": ["Custos diretos (materiais e mercadorias) - Mês 1", "Mão de obra direta (produção)",
                          "Custos indiretos de produção", "Depreciação", "Pessoal administrativo e comercial",
                          "Despesas administrativas", "Despesas comerciais", "Despesas financeiras"],
        "Valor mensal": [brl(v) for v in (res["cdir_m1"], res["mod_m"], res["cif_m"], res["dep_m"],
                                          res["pessoal_adm_m"] + res["pessoal_com_m"], res["desp_adm_m"],
                                          res["desp_com_m"], res["desp_fin_m"])],
        "Vai para o produto (absorção)?": ["Sim (direto)", "Sim (rateio)", "Sim (rateio)", "Sim (rateio)",
                                           "Não (despesa do período)", "Não", "Não", "Não"]}),
        larguras=(90, 40, 50), alinhar=("LEFT", "RIGHT", "LEFT"))
    pdf.paragrafo(f"Critério de rateio dos custos indiretos: {res['criterio_rateio']}.", "I", 8)

    pdf.secao("FORMAÇÃO DE PREÇO E PONTO DE EQUILÍBRIO POR PRODUTO", nova_pagina=True)
    pdf.paragrafo(f"Markup divisor = 1 / [1 - (impostos {pct(res['aliq'])} + despesas variáveis {pct(res['var_pct'])} "
                  f"+ despesas fixas {pct(res['df_pct'])} + margem desejada {pct(res['margem_desejada'])})] = "
                  f"{num_br(res['markup_mult'], 3)}. Preço sugerido = custo direto x markup.", "", 9)
    cp = res["cat"]
    pdf.tabela(pd.DataFrame({"Produto/serviço": cp["descricao"], "Custo variável": cp["custo_unit"].map(brl),
                             "Custo absorção": cp["custo_absorcao_unit"].map(brl), "Preço": cp["preco_venda"].map(brl),
                             "Sugerido": cp["preco_sugerido"].map(brl), "Margem abs.": cp["margem_absorcao_pct"].map(pct),
                             "PE un./mês": cp["pe_unidades_mes"].map(lambda v: num_br(v, 0))}),
               larguras=(56, 21, 21, 20, 20, 21, 21), tam=7)

    # ---------------- SERVIÇOS E ASSINATURAS ----------------
    ch = custo_hora(res, d, proj)
    if not ch["tabela"].empty:
        pdf.secao("SERVIÇOS: CUSTO DA HORA E PREÇO", nova_pagina=True)
        pdf.paragrafo(f"Equipe da operação: {num_br(ch['pessoas'], 0)} pessoa(s) x {num_br(ch['horas_mes_pessoa'], 0)} h/mês x "
                      f"utilização de {pct(ch['utilizacao'], 0)} = {num_br(ch['horas_cobraveis'], 0)} horas cobráveis por mês. "
                      f"Custo da hora produtiva = {brl(ch['custo_base'])} ÷ {num_br(ch['horas_cobraveis'], 0)} h = "
                      f"{brl(ch['custo_hora'])}. Ocupação prevista: {pct(ch['ocupacao_m1'])} no Mês 1 e "
                      f"{pct(ch['ocupacao_m12'])} no Mês 12.", "", 9)
        t_ = ch["tabela"]
        pdf.tabela(pd.DataFrame({"Serviço": t_["Serviço"], "Horas": t_["Horas por serviço"].map(lambda v: num_br(v, 1)),
                                 "Insumos": t_["Insumos"].map(brl), "Mão de obra": t_["Mão de obra (horas x custo hora)"].map(brl),
                                 "Custo pleno": t_["Custo pleno"].map(brl), "Preço": t_["Preço atual"].map(brl),
                                 "Sugerido": t_["Preço sugerido"].map(lambda v: brl(v) if v == v else "n.c."),
                                 "R$/hora": t_["Valor da hora cobrada"].map(brl)}),
                   larguras=(52, 14, 19, 20, 20, 18, 19, 18), tam=7)
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        s_ = res["ass"]
        pdf.secao("ASSINATURAS E RECEITA RECORRENTE", nova_pagina=True)
        a_ = d["assinaturas"]
        pdf.tabela(pd.DataFrame({"Plano": a_["plano"], "Preço/ciclo": a_["preco_ciclo"].map(brl),
                                 "Ciclo": a_["ciclo_meses"].map(lambda v: f"{num_br(v, 0)} m"),
                                 "Novos Mês 1": a_["novos_mes1"].map(lambda v: num_br(v, 0)),
                                 "Cresc. novos": a_["crescimento_novos_pct"].map(lambda v: pct(v / 100)),
                                 "Churn/mês": a_["churn_mensal_pct"].map(lambda v: pct(v / 100)),
                                 "Custo entrega": a_["custo_entrega_mensal"].map(brl),
                                 "Taxa": a_["taxa_cobranca_pct"].map(lambda v: pct(v / 100))}),
                   larguras=(52, 20, 12, 18, 18, 18, 22, 20), tam=7)
        pdf.subtitulo("Projeção por ano")
        linhas = []
        for k in (1, 2, 3):
            b_ = s_.iloc[(k - 1) * 12:k * 12]
            r_, c_ = float(b_["Receita reconhecida"].sum()), float(b_["Cobranças (caixa)"].sum())
            ce, tx = float(b_["Custo de entrega"].sum()), float(b_["Taxas de cobrança"].sum())
            linhas.append({"Ano": f"Ano {k}", "Novos": num_br(b_["Novos"].sum(), 0),
                           "Cancelados": num_br(b_["Cancelamentos"].sum(), 0), "Ativos (fim)": num_br(b_["Ativos"].iloc[-1], 0),
                           "Receita reconhecida": brl(r_), "Cobranças": brl(c_), "Margem bruta": pct(div(r_ - ce - tx, r_))})
        pdf.tabela(pd.DataFrame(linhas), larguras=(14, 18, 22, 22, 38, 36, 30), tam=7.5)
        pdf.subtitulo("Indicadores de recorrência")
        pdf.tabela(pd.DataFrame({
            "Indicador": ["MRR no Mês 12", "ARR (MRR x 12)", "Assinantes no Mês 12", "ARPU", "Churn mensal médio",
                          "Margem bruta recorrente", "LTV", "CAC", "LTV/CAC", "Payback do CAC", "Receita antecipada (fim do Ano 1)"],
            "Valor": [brl(rc["mrr12"]), brl(rc["arr12"]), num_br(rc["ativos12"], 0), brl(rc["arpu"]), pct(rc["churn"]),
                      pct(rc["margem_bruta"]), brl(rc["ltv"]) if rc["ltv"] == rc["ltv"] else "n.a.",
                      brl(rc["cac"]) if rc["cac"] == rc["cac"] else "n.a.",
                      f"{num_br(rc['ltv_cac'], 1)}x" if rc["ltv_cac"] == rc["ltv_cac"] else "n.a.",
                      f"{num_br(rc['payback_cac'], 1)} meses" if rc["payback_cac"] == rc["payback_cac"] else "n.a.",
                      brl(rc["receita_diferida_a1"])],
            "Como se calcula": ["Valor mensal equivalente dos assinantes ativos", "", "", "MRR ÷ assinantes",
                                "Churn de cada plano ponderado pelos ativos", "(Receita − entrega − taxas) ÷ receita",
                                "ARPU x margem bruta ÷ churn", "Gasto comercial do Mês 1 ÷ novos do Mês 1", "Referência: 3x ou mais",
                                "CAC ÷ (ARPU x margem bruta)", "Cobrado e ainda não prestado (passivo)"]}),
            larguras=(55, 35, 90), alinhar=("LEFT", "RIGHT", "LEFT"), tam=8)
        pdf.paragrafo("A receita das assinaturas é reconhecida mês a mês; planos pagos por ciclo entram inteiros no caixa na "
                      "cobrança, e a parte ainda não prestada aparece no balanço como receita antecipada.", "I", 8)

    pdf.secao("PROJEÇÃO DE RESULTADOS E DRE", nova_pagina=True)
    pdf.paragrafo(f"Premissas: crescimento de {num_br(proj['crescimento_mensal_ano1_pct'])}% ao mês no Ano 1 e "
                  f"{num_br(proj['crescimento_anual_pct'])}% ao ano a partir do Ano 2; custos fixos corrigidos em "
                  f"{num_br(proj['inflacao_custos_anual_pct'])}% ao ano; TMA de {num_br(proj['tma_anual_pct'])}% ao ano.", "I", 8)
    pdf.subtitulo("Projeção mensal do Ano 1")
    mm = res["meses"].copy()
    for col in mm.columns[1:]:
        mm[col] = mm[col].map(brl)
    pdf.tabela(mm, larguras=(18, 33, 33, 33, 30, 33), tam=7.5)
    pdf.subtitulo("Demonstração do resultado do exercício (Anos 1 a 3)")
    pdf.tabela(extras["dre"], larguras=(70, 30, 30, 30, 20), tam=7.5)

    # ---------------- 13. INDICADORES ----------------
    pdf.secao("FLUXO DE CAIXA MENSAL (ANO 1)", nova_pagina=True)
    fm = res["fluxo_mensal"]
    if not fm.empty:
        tf_ = fm.copy()
        for col in tf_.columns[1:]:
            tf_[col] = tf_[col].map(brl)
        tf_ = tf_.rename(columns={"Impostos e desp. variáveis": "Impostos/var.", "Pessoal e custos fixos": "Pessoal/fixos",
                                  "Parcelas de empréstimos": "Empréstimos"})
        pdf.tabela(tf_, larguras=(14, 25, 25, 23, 22, 24, 22, 25), tam=6.5)
        pdf.paragrafo(f"Menor saldo de caixa: {brl(fm['Saldo final'].min())} ({fm.loc[fm['Saldo final'].idxmin(), 'Mês']}).", "B", 9)

    pdf.secao("INDICADORES DE VIABILIDADE", nova_pagina=True)
    sem = semaforo(res)
    pdf.subtitulo("Semáforo da saúde do negócio")
    pdf.tabela(pd.DataFrame({"Situação": [{"🟢": "Bom", "🟡": "Atenção", "🔴": "Crítico"}.get(s_["Status"], "") for s_ in sem],
                             "Indicador": [s_["Indicador"] for s_ in sem], "Valor": [s_["Valor"] for s_ in sem],
                             "Leitura": [s_["Leitura"] for s_ in sem]}), larguras=(18, 45, 27, 90),
               alinhar=("CENTER", "LEFT", "RIGHT", "LEFT"), tam=7.5)
    rs = res["resumo"]
    pdf.subtitulo("Resumo financeiro (Ano 1, padrão PNBOX)")
    pdf.tabela(pd.DataFrame({
        "Indicador": ["Receita anual", "Custos variáveis", "Margem de contribuição", "Custos fixos (com depreciação)",
                      "Resultado operacional", "Lucratividade", "Rentabilidade", "Payback simples",
                      "PE contábil", "PE financeiro", "PE econômico"],
        "Valor": [brl(rs["receita_anual"]), brl(rs["custos_variaveis"]), brl(rs["margem_contrib"]),
                  brl(rs["custos_fixos"]), brl(rs["resultado_operacional"]), pct(rs["lucratividade"]),
                  pct(rs["rentabilidade"]),
                  f"{num_br(rs['payback_meses'], 1)} meses" if rs["payback_meses"] == rs["payback_meses"] else "Não recupera",
                  brl(a1["pe_contabil"]), brl(a1["pe_financeiro"]), brl(a1["pe_economico"])]}),
        larguras=(120, 60))
    pdf.subtitulo("Avaliação do investimento (horizonte de 3 anos)")
    fl = res["fluxos"]
    pdf.tabela(pd.DataFrame({
        "Indicador": ["Fluxos de caixa do projeto (Ano 0 a 3)", "VPL à TMA", "TIR", "Índice de lucratividade",
                      "Payback descontado"],
        "Valor": [" | ".join(brl(v) for v in fl), brl(res["vpl"]),
                  pct(res["tir"]) if res["tir"] is not None else "n.a.",
                  f"{num_br(res['il'])}x" if res["il"] == res["il"] else "n.a.",
                  f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else "acima de 36 meses"]}),
        larguras=(55, 125), alinhar=("LEFT", "LEFT"))
    pdf.paragrafo("Fluxo do projeto: Ano 0 = investimento total; Anos 1 a 3 = EBITDA" +
                  (" (Ano 3 inclui a recuperação do capital de giro)." if proj.get("recuperar_giro") else "."), "I", 7.5)
    pdf.subtitulo("Matriz de indicadores em 5 blocos")
    ind = extras["indicadores"]
    for bloco, g in ind.groupby("Bloco", sort=False):
        pdf.fonte("B", 8.5)
        if pdf.get_y() > 250:
            pdf.add_page()
        pdf.cell(0, 5, pdf.t(bloco), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.tabela(g[["Indicador", "Fórmula", "Ano 1", "Ano 2", "Ano 3", "Var. Ano 2/Ano 1"]],
                   larguras=(42, 46, 25, 25, 25, 17), alinhar=("LEFT", "LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"), tam=7)

    # ---------------- 14. GRÁFICOS ----------------
    imgs = graficos_png(extras["comp"]) if extras.get("comp") else []
    if imgs:
        pdf.secao("GRÁFICOS", nova_pagina=True)
        roscas = imgs[:5]
        outros = imgs[5:]
        largura = (pdf.epw - 6) / 2
        for i in range(0, len(roscas), 2):
            par = roscas[i:i + 2]
            if pdf.get_y() > 175:
                pdf.add_page()
            y = pdf.get_y()
            alturas = []
            for j, (tit, png) in enumerate(par):
                info = pdf.image(io.BytesIO(png), x=pdf.l_margin + j * (largura + 6), y=y, w=largura)
                alturas.append(info.rendered_height)
            pdf.set_y(y + max(alturas) + 4)
        for tit, png in outros:
            info_h = 95
            if pdf.get_y() > 297 - 15 - info_h:
                pdf.add_page()
            info = pdf.image(io.BytesIO(png), x=pdf.l_margin, w=pdf.epw * 0.92)
            pdf.ln(4)

    # ---------------- 15. CENÁRIOS E CHECAGENS ----------------
    # ---------------- FERRAMENTAS COMPLEMENTARES ----------------
    rk, cr, ex = d.get("riscos"), d.get("cronograma"), d.get("experimentos")
    if any(x is not None and not x.empty for x in (rk, cr, ex)):
        pdf.secao("RISCOS, CRONOGRAMA E VALIDAÇÃO", nova_pagina=True)
        if rk is not None and not rk.empty:
            pdf.subtitulo("Riscos e plano de contingência")
            pdf.tabela(pd.DataFrame({"Risco": rk["risco"], "Prob.": rk["probabilidade"].map(lambda v: num_br(v, 0)),
                                     "Impacto": rk["impacto"].map(lambda v: num_br(v, 0)),
                                     "Nível": (rk["probabilidade"] * rk["impacto"]).map(lambda v: num_br(v, 0)),
                                     "Mitigação": rk["mitigacao"], "Responsável": rk["responsavel"]}),
                       larguras=(42, 12, 14, 12, 72, 28), alinhar=("LEFT", "CENTER", "CENTER", "CENTER", "LEFT", "LEFT"), tam=7.5)
        if cr is not None and not cr.empty:
            pdf.subtitulo("Cronograma de implantação")
            pdf.tabela(pd.DataFrame({"Atividade": cr["atividade"], "Início (mês)": cr["mes_inicio"].map(lambda v: num_br(v, 0)),
                                     "Duração (meses)": cr["duracao_meses"].map(lambda v: num_br(v, 1)),
                                     "Responsável": cr["responsavel"], "Custo": cr["custo"].map(brl)}),
                       larguras=(70, 22, 25, 35, 28), alinhar=("LEFT", "CENTER", "CENTER", "LEFT", "RIGHT"), tam=7.5)
        if ex is not None and not ex.empty:
            pdf.subtitulo("Quadro de experimentação")
            pdf.tabela(pd.DataFrame({"Hipótese": ex["hipotese"], "Experimento": ex["experimento"], "Métrica": ex["metrica"],
                                     "Critério": ex["criterio_sucesso"], "Resultado": ex["resultado"], "Situação": ex["status"]}),
                       larguras=(38, 38, 24, 26, 34, 20), alinhar=("LEFT",) * 6, tam=7)

    pdf.secao("SIMULADOR DE CENÁRIOS E CHECAGENS", nova_pagina=True)
    pdf.tabela(extras["cenarios"], larguras=(45, 45, 45, 45), tam=7.5)
    pdf.subtitulo("Checagens de consistência")
    ch = extras["checagens"]
    pdf.tabela(ch[["Teste", "Valor apurado", "Situação", "Observação"]], larguras=(50, 40, 20, 70),
               alinhar=("LEFT", "LEFT", "CENTER", "LEFT"), tam=7.5)

    if str(proj.get("conclusao") or "").strip():
        pdf.secao("CONCLUSÃO E PARECER FINAL", nova_pagina=True)
        for par in str(proj["conclusao"]).split("\n"):
            pdf.paragrafo(par, "", 10)

    out = pdf.output()
    return bytes(out)
