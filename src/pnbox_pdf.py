"""
================================================================================
VIABILIZALAB - RELATÓRIO EM PDF (pnbox_pdf.py)
================================================================================
Gera o relatório completo do plano de negócio, na ordem do plano de negócio do Sebrae.
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

from pnbox_validacao import *  # noqa
from pnbox_v8 import *  # noqa
from pnbox_mix import *  # noqa

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
            self.set_text_color(*((255, 255, 255) if j == 0 else (40, 50, 60)))
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


# ------------------------------------------------------------------------------
# FIGURAS DA ANÁLISE ESTRATÉGICA (matplotlib) E QUADRO DO CANVAS
# ------------------------------------------------------------------------------
def _fig_png(fig) -> bytes:
    import matplotlib.pyplot as plt
    b = io.BytesIO()
    fig.savefig(b, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    return b.getvalue()


def png_porter(tab: pd.DataFrame) -> bytes | None:
    try:
        import numpy as np
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return None
    vals = tab["Intensidade"].tolist()
    rot = ["Rivalidade", "Novos\nentrantes", "Substitutos", "Clientes", "Fornecedores"]
    ang = np.linspace(0, 2 * np.pi, len(vals), endpoint=False).tolist()
    fig, ax = plt.subplots(figsize=(4.2, 4.2), subplot_kw={"polar": True})
    ax.plot(ang + ang[:1], vals + vals[:1], color="#2f6fd0", linewidth=2)
    ax.fill(ang + ang[:1], vals + vals[:1], color="#2f6fd0", alpha=0.25)
    ax.set_xticks(ang)
    ax.set_xticklabels(rot, fontsize=9)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.tick_params(axis="y", labelsize=7, colors="#777")
    return _fig_png(fig)


def png_ge(tab: pd.DataFrame) -> bytes | None:
    """Matriz GE/McKinsey 3x3 com células coloridas (força alta à esquerda) e segmentos numerados."""
    try:
        plt = _plt()
        from matplotlib.patches import FancyBboxPatch
    except ImportError:
        return None
    cores = {"Seguro e crescimento": "#3fd0c0", "Seletivo ou cauteloso": "#ffd23f",
             "Expansão limitada ou colheita": "#ffa630", "Perigo ou colheita": "#ec4561"}
    grade = [["Seletivo ou cauteloso", "Expansão limitada ou colheita", "Perigo ou colheita"],
             ["Seguro e crescimento", "Seletivo ou cauteloso", "Expansão limitada ou colheita"],
             ["Seguro e crescimento", "Seguro e crescimento", "Seletivo ou cauteloso"]]  # linha 0 = atratividade baixa
    fig, ax = plt.subplots(figsize=(6.2, 5.4))
    s = 10 / 3
    for i in range(3):
        for j in range(3):
            n = grade[i][j]
            ax.add_patch(FancyBboxPatch((j * s + 0.12, i * s + 0.12), s - 0.24, s - 0.24, boxstyle="round,pad=0,rounding_size=0.35",
                                        color=cores[n], lw=0, zorder=1))
            ax.text(j * s + s / 2, i * s + s - 0.4, n.replace(" ou ", " ou\n"), ha="center", va="top", fontsize=8.5,
                    fontweight="bold", color="#253041", alpha=0.55, zorder=2)
    for k, (_, r) in enumerate(tab.iterrows(), start=1):
        x, y = 10 - float(r["Força competitiva"]), float(r["Atratividade"])
        ax.scatter([x], [y], s=380, color="#1f3f75", edgecolors="white", linewidths=2, zorder=3)
        ax.text(x, y, str(k), ha="center", va="center", color="white", fontsize=9, fontweight="bold", zorder=4)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.set_xticks([s / 2, 1.5 * s, 2.5 * s])
    ax.set_xticklabels(["Alta", "Média", "Baixa"])
    ax.set_yticks([s / 2, 1.5 * s, 2.5 * s])
    ax.set_yticklabels(["Baixa", "Média", "Alta"])
    ax.set_xlabel("Força competitiva da empresa no segmento")
    ax.set_ylabel("Atratividade do mercado")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(length=0)
    return _fig_png(fig)


def png_bcg(tab: pd.DataFrame, lim: float) -> bytes | None:
    """Matriz BCG com quadrantes coloridos, descrições e bolhas pela receita (participação em log, líder à esquerda)."""
    try:
        plt = _plt()
        import numpy as np
    except ImportError:
        return None
    ymax = max(2 * lim, 1.0)
    quad = {"Estrela": ((0.1, lim), "#ede3fb", "#7b3fe4", ["Mercado atrativo", "Gera bom faturamento", "Exige investimento"]),
            "Interrogação": ((-1, lim), "#fff4c7", "#b38600", ["Ainda gera pouco retorno", "Expectativa de crescimento", "Decisão de risco"]),
            "Vaca leiteira": ((0.1, 0), "#eefbd0", "#4f7a00", ["Gera muito caixa", "Exige pouco investimento", "Recursos realocáveis"]),
            "Abacaxi": ((-1, 0), "#fde2e5", "#c8102e", ["Gera pouco caixa", "Mercado sem crescimento", "Provável saída"])}
    fig, ax = plt.subplots(figsize=(6.6, 5.4))
    # eixo x em log10 invertido: posição = 1 - log10(p)  (10x -> 0; 1x -> 1; 0,1x -> 2)
    for q, ((_, y0), bg, cor, itens) in quad.items():
        x0 = 0 if q in ("Estrela", "Vaca leiteira") else 1
        ax.add_patch(plt.Rectangle((x0, y0), 1, lim if y0 == 0 else ymax - lim, color=bg, zorder=0))
        topo = y0 + (lim if y0 == 0 else ymax - lim)
        # títulos: no alto dos quadrantes de cima e no rodapé dos de baixo (como na matriz clássica)
        ax.text(x0 + 0.5, topo - 0.06 * ymax if y0 > 0 else y0 + 0.03 * ymax, q, ha="center", fontsize=12, fontweight="bold",
                color=cor, zorder=1)
        for n, it in enumerate(itens):
            yy = topo - (0.13 if y0 > 0 else 0.06) * ymax - n * 0.06 * ymax
            ax.text(x0 + 0.06, yy, "✓ " + it, fontsize=7.3, color="#334", alpha=0.75, zorder=1)
    mx = max(float(tab["Receita anual"].max()), 1.0)
    cores = {q: v[2] for q, v in quad.items()}
    for _, r in tab.iterrows():
        x = 1 - np.log10(min(max(float(r["Participação relativa"]), 0.1), 10))
        y = min(max(float(r["Crescimento do mercado (%)"]), 0), ymax * 0.97)
        ax.scatter([x], [y], s=150 + 1700 * float(r["Receita anual"]) / mx, color=cores[r["Quadrante"]], alpha=0.8,
                   edgecolors="white", linewidths=2, zorder=3)
        ax.annotate(str(r["Produto"])[:26], (x, y), textcoords="offset points", xytext=(0, -16), ha="center", fontsize=7,
                    fontweight="bold", zorder=4, bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.8))
    ax.axvline(1, color="#2d4f8a", lw=1)
    ax.axhline(lim, color="#2d4f8a", lw=1)
    ax.set_xlim(0, 2)
    ax.set_ylim(0, ymax)
    ax.set_xticks([0, 1 - np.log10(5), 1, 1 - np.log10(0.5), 2])
    ax.set_xticklabels(["10x", "5x", "1x", "0,5x", "0,1x"])
    ax.set_yticks([0, lim / 2, lim, 1.5 * lim, ymax])
    ax.set_yticklabels([f"{num_br(v, 0)}%" for v in (0, lim / 2, lim, 1.5 * lim, ymax)])
    ax.set_xlabel("Participação relativa de mercado (alta à esquerda)")
    ax.set_ylabel("Taxa de crescimento do mercado")
    for sp in ax.spines.values():
        sp.set_color("#2d4f8a")
    return _fig_png(fig)


def png_bsc(ob: pd.DataFrame) -> bytes | None:
    """Ciclo do Balanced Scorecard (aprendizado -> processos -> clientes -> financeira), com o número de objetivos."""
    try:
        plt = _plt()
        import numpy as np
        from matplotlib.patches import Wedge, Polygon
    except ImportError:
        return None
    seg = [("APRENDIZADO E\nCRESCIMENTO", "Aprendizado e crescimento", "#0d6fb8", 90, "Conhecimento e habilidades\nda equipe sustentam\nas melhorias."),
           ("PROCESSOS\nINTERNOS", "Processos internos", "#123a5e", 0, "Equipe capacitada\nmelhora os processos."),
           ("CLIENTES", "Clientes", "#d7262e", -90, "Processos melhores\nconduzem a clientes\nmais satisfeitos."),
           ("FINANCEIRA", "Financeira", "#b5191f", 180, "Clientes satisfeitos\nlevam a melhores\nresultados.")]
    R1, R0 = 1.0, 0.48
    Rm = (R1 + R0) / 2
    fig, ax = plt.subplots(figsize=(5.6, 5.6))
    for tit, nome, cor, a_ini, desc in seg:
        a_fim = a_ini - 86  # sentido horário
        ax.add_patch(Wedge((0, 0), R1, a_fim, a_ini, width=R1 - R0, color=cor))
        ang = np.radians(a_fim)
        ponta = np.radians(a_fim - 7)
        ax.add_patch(Polygon([(R1 * np.cos(ang), R1 * np.sin(ang)), (Rm * np.cos(ponta), Rm * np.sin(ponta)),
                              (R0 * np.cos(ang), R0 * np.sin(ang))], color=cor))
        am = np.radians(a_ini - 43)
        n = int((ob["perspectiva"] == nome).sum()) if ob is not None and not ob.empty else 0
        ax.text(Rm * np.cos(am), Rm * np.sin(am), f"{tit}\n\n{desc}\n{n} objetivo{'s' if n != 1 else ''}", ha="center",
                va="center", color="white", fontsize=6.8, fontweight="bold", linespacing=1.15)
    ax.add_patch(plt.Circle((0, 0), R0 - 0.03, color="white"))
    ax.text(0, 0, "BSC", ha="center", va="center", fontsize=34, fontweight="bold", color="#333")
    ax.set_xlim(-1.08, 1.08)
    ax.set_ylim(-1.08, 1.08)
    ax.set_aspect("equal")
    ax.axis("off")
    return _fig_png(fig)


CORES_CANVAS = {"Parcerias-chave": (106, 49, 132), "Atividades-chave": (240, 144, 31), "Recursos-chave": (229, 101, 62),
                "Proposta de valor": (68, 146, 58), "Relacionamento com clientes": (141, 184, 51), "Canais": (20, 164, 154),
                "Segmentos de clientes": (21, 146, 196), "Estrutura de custos": (68, 86, 160), "Fontes de receita": (34, 97, 122)}


def canvas_pdf(pdf, bl: dict):
    """Desenha o Business Model Canvas em uma página (retrato, 180 mm de largura), com um bloco colorido por campo."""
    x0, y0, w, h1, h2, g = pdf.l_margin, pdf.get_y() + 2, 36, 62, 45, 1.2
    caixas = [("Parcerias-chave", 0, 0, 1, 2), ("Atividades-chave", 1, 0, 1, 1), ("Recursos-chave", 1, 1, 1, 1),
              ("Proposta de valor", 2, 0, 1, 2), ("Relacionamento com clientes", 3, 0, 1, 1), ("Canais", 3, 1, 1, 1),
              ("Segmentos de clientes", 4, 0, 1, 2)]
    for nome, c, r, cw, rh in caixas:
        x, y = x0 + c * w, y0 + r * h1
        _texto_caixa(pdf, nome, bl.get(nome, []), x + g / 2, y + g / 2, w * cw - g, h1 * rh - g)
    yb = y0 + 2 * h1
    for i, nome in enumerate(["Estrutura de custos", "Fontes de receita"]):
        _texto_caixa(pdf, nome, bl.get(nome, []), x0 + i * 90 + g / 2, yb + g / 2, 90 - g, h2 - g)
    pdf.set_y(yb + h2 + 3)


def _texto_caixa(pdf, nome, itens, x, y, w, h):
    cor = CORES_CANVAS.get(nome, AZUL)
    pdf.set_fill_color(*cor)
    try:
        pdf.rect(x, y, w, h, style="F", round_corners=True, corner_radius=2.5)
    except TypeError:  # fpdf2 antigo, sem cantos arredondados
        pdf.rect(x, y, w, h, style="F")
    pdf.set_xy(x + 2, y + 2)
    pdf.fonte("B", 7.2)
    pdf.set_text_color(255, 255, 255)
    pdf.multi_cell(w - 4, 3.4, pdf.t(nome.upper()), align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.fonte("", 6.6)
    yy = pdf.get_y() + 1.5
    for it in itens:
        if yy > y + h - 4:
            break
        pdf.set_xy(x + 2, yy)
        pdf.multi_cell(w - 4, 3.1, pdf.t("• " + str(it))[:160], align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        yy = pdf.get_y() + 0.6
    pdf.set_text_color(0, 0, 0)


# ------------------------------------------------------------------------------
# FIGURAS E SEÇÕES DA VALIDAÇÃO, MARKETING, INCERTEZA E DEFESA (v7.1)
# ------------------------------------------------------------------------------
def _plt():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    return plt


def png_vw(vw: dict, preco: float) -> bytes:
    plt = _plt()
    cv = vw["curvas"]
    fig, ax = plt.subplots(figsize=(6.6, 3.6))
    ax.axvspan(vw["pmc"], vw["pme"], color="#27ae60", alpha=0.12, label="faixa aceitável")
    for c, cor in (("Muito barato", "#7fb3d5"), ("Barato", "#2471a3"), ("Caro", "#e59866"), ("Muito caro", "#c0392b")):
        ax.plot(cv["Preço"], cv[c] * 100, color=cor, lw=2, label=c)
    if preco == preco:
        ax.axvline(preco, color="black", ls="--", lw=1, label="preço do plano")
    ax.set_xlabel("Preço (R$)")
    ax.set_ylabel("% dos respondentes")
    ax.legend(fontsize=7, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.18), frameon=False)
    return _fig_png(fig)


def png_hist_mc(vpl) -> bytes:
    plt = _plt()
    fig, ax = plt.subplots(figsize=(6.6, 3.0))
    n_, bins, patches = ax.hist(vpl, bins=45)
    for b_, p_ in zip(bins, patches):
        p_.set_facecolor("#27ae60" if b_ >= 0 else "#c0392b")
    ax.axvline(0, color="black", lw=1)
    from matplotlib.ticker import FuncFormatter
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v / 1000:,.0f}".replace(",", ".")))
    ax.set_xlabel("VPL (R$ mil)")
    ax.set_ylabel("Cenários")
    return _fig_png(fig)


def png_curva_valor(cv: pd.DataFrame) -> bytes:
    plt = _plt()
    ordem = list(dict.fromkeys(cv["Atributo"]))
    fig, ax = plt.subplots(figsize=(6.8, 3.4))
    for emp, g in cv.groupby("Empresa", sort=False):
        g = g.set_index("Atributo").reindex(ordem)
        novo = emp == "Nova curva (ERRC)"
        ax.plot(range(len(ordem)), g["Nota"], marker="o", lw=3 if novo else 1.4, ls="-" if novo else "--",
                color="#e67e22" if novo else None, label=emp[:28])
    ax.set_xticks(range(len(ordem)))
    ax.set_xticklabels([o[:18] for o in ordem], rotation=35, ha="right", fontsize=7)
    ax.set_ylim(0, 10.5)
    ax.legend(fontsize=7, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.38), frameon=False)
    return _fig_png(fig)


def png_bass(tb: pd.DataFrame) -> bytes:
    plt = _plt()
    fig, ax = plt.subplots(figsize=(6.4, 2.8))
    ax.plot(tb["Ano"], tb["Adotantes acumulados"], marker="o", lw=2.5, label="Adoção prevista (Bass)")
    if "Clientes no plano" in tb:
        ax.plot(tb["Ano"], tb["Clientes no plano"], marker="s", lw=2.5, label="Clientes no plano")
    ax.set_xlabel("Ano")
    ax.set_ylabel("Clientes")
    ax.set_xticks(list(tb["Ano"]))
    ax.legend(fontsize=8, frameon=False)
    return _fig_png(fig)


def secoes_v71(pdf, proj, d, res):
    preco = preco_referencia(proj, res)
    # ---- Validação do mercado
    pl = placar_evidencias(proj, d, res)
    pdf.secao("VALIDAÇÃO: PLACAR DE EVIDÊNCIAS", nova_pagina=True)
    pdf.paragrafo(f"{pct(pl['score'], 0)} do resultado está apoiado em premissas pesquisadas ou testadas com clientes. "
                  f"Premissa a validar primeiro: {pl['proxima'] or 'nenhuma'}.", "B", 9)
    tb = pl["tabela"]
    pdf.tabela(pd.DataFrame({"Premissa": tb["premissa"], "Valor no plano": tb["valor_plano"], "Peso": tb["peso"].map(lambda v: pct(v, 0)),
                             "Origem": tb["status"].map(STATUS_CURTO), "Fonte": tb["fonte"]}),
               larguras=(46, 46, 14, 20, 54), alinhar=("LEFT", "LEFT", "RIGHT", "CENTER", "LEFT"), tam=7.2)
    ent = resumo_entrevistas(d, preco)
    if ent["n"]:
        pdf.subtitulo(f"Entrevistas com clientes ({ent['n']})")
        pdf.paragrafo(f"Pagariam: {pct(ent['pct_sim'], 0)} sim e {pct(ent['pct_talvez'], 0)} talvez. Valor mediano citado: "
                      f"{brl(ent['mediana_valor'])} (preço de referência do plano: {brl(preco)}).", "", 8.5)
        pdf.tabela(ent["dores"].assign(**{"% das entrevistas": ent["dores"]["% das entrevistas"].map(lambda v: pct(v, 0))}),
                   larguras=(120, 25, 35), alinhar=("LEFT", "RIGHT", "RIGHT"), tam=7.5)
    vw = van_westendorp(d.get("precos_vw"))
    if vw and vw.get("ok"):
        pdf.subtitulo(f"Teste de preço Van Westendorp ({vw['n']} respostas)")
        pdf.image(io.BytesIO(png_vw(vw, preco)), x=30, w=150)
        pdf.paragrafo(f"Faixa aceitável: {brl(vw['pmc'])} a {brl(vw['pme'])}. Preço ótimo: {brl(vw['opp'])}. Indiferença: "
                      f"{brl(vw['ipp'])}. Preço do plano: {brl(preco)}.", "", 8.5)
    ia = intencao_ajustada(proj, res)
    if ia["ok"]:
        pdf.paragrafo(f"Intenção de compra ajustada: {num_br(ia['ajustados'], 0)} clientes prováveis x "
                      f"{num_br(ia['plano'], 0)} clientes supostos pelo plano no Ano 1.", "", 8.5)
    # ---- Marketing estratégico
    tem_vpc = any(proj.get(k) for k, _, _ in BLOCOS_VPC)
    frase = frase_posicionamento(proj)
    if tem_vpc or frase or not d["errc"].empty:
        pdf.secao("PROPOSTA DE VALOR, POSICIONAMENTO E CURVA DE VALOR", nova_pagina=True)
        if proj.get("vpc_job"):
            pdf.campo("Trabalho do cliente (job)", proj["vpc_job"])
        if tem_vpc:
            lin = [(rot, "; ".join(linhas_txt(proj.get(k_)))) for k_, rot, _ in BLOCOS_VPC]
            pdf.tabela(pd.DataFrame(lin, columns=["Bloco", "Itens"]), larguras=(45, 135), alinhar=("LEFT", "LEFT"), tam=7.5)
            enc = encaixe_vpc(proj, d)
            if enc["total"]:
                pdf.paragrafo(f"Encaixe: {enc['cobertos']} de {enc['total']} dores e ganhos têm resposta na oferta.", "I", 8)
        if frase:
            pdf.subtitulo("Declaração de posicionamento")
            pdf.paragrafo(frase, "", 9)
        if not d["errc"].empty:
            pdf.subtitulo("Curva de valor e matriz ERRC")
            pdf.image(io.BytesIO(png_curva_valor(curva_valor(d))), x=28, w=154)
            pdf.tabela(d["errc"].assign(nova_nota=d["errc"]["nova_nota"].map(lambda v: num_br(v, 1))).rename(columns={
                "atributo": "Atributo", "acao": "Ação", "nova_nota": "Nova nota", "justificativa": "Justificativa"}),
                larguras=(45, 22, 20, 93), alinhar=("LEFT", "CENTER", "CENTER", "LEFT"), tam=7.5)
            pdf.paragrafo(leitura_errc(d), "I", 8)
    oc = orcamento_canais(proj, d, res)
    if not oc.empty or not d["aarrr"].empty:
        pdf.secao("ORÇAMENTO DE MARKETING E FUNIL DE CRESCIMENTO", nova_pagina=True)
        if not oc.empty:
            pdf.tabela(pd.DataFrame({"Canal": oc["Canal"], "Gasto mensal": oc["Gasto mensal"].map(brl),
                                     "Clientes/mês": oc["Clientes novos/mês"].map(lambda v: num_br(v, 0)),
                                     "CAC": oc["CAC"].map(lambda v: brl(v) if v == v else "n.a."),
                                     "Retorno por R$ 1": oc["Retorno por R$ 1"].map(lambda v: f"R$ {num_br(v, 2)}" if v == v else "n.a.")}),
                       larguras=(62, 30, 24, 30, 34), alinhar=("LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"), tam=7.5)
            pdf.paragrafo(f"Valor do cliente usado no retorno: {oc.attrs.get('rotulo_valor', '')} ({brl(float(oc['Valor do cliente'].iloc[0]))}).", "I", 7.5)
        if not d["aarrr"].empty:
            pdf.subtitulo("Funil AARRR")
            pdf.tabela(d["aarrr"].rename(columns={"etapa": "Etapa", "metrica": "Métrica", "meta": "Meta", "acao": "Ação"}),
                       larguras=(24, 50, 34, 72), alinhar=("LEFT",) * 4, tam=7.5)
        pdf.paragrafo(f"Meta de NPS no primeiro ano: {num_br(proj.get('nps_meta') or 0, 0)}.", "", 8.5)
    # ---- Incerteza
    if res["anos"][1]["rec_bruta"] > 0:
        mc = monte_carlo(res, faixas_mc(d))
        pdf.secao("INCERTEZA: SIMULAÇÃO DE MONTE CARLO E PERDA ACEITÁVEL", nova_pagina=True)
        pdf.paragrafo(f"{mc['n']} cenários sorteados (distribuição triangular). Chance de VPL positivo: {pct(mc['p_vpl_pos'], 0)}. "
                      f"VPL nos 5% piores casos: {brl(mc['p5'])}; mediano: {brl(mc['p50'])}; nos 5% melhores: {brl(mc['p95'])}. "
                      f"Chance de caixa negativo no Ano 1: {pct(mc['p_caixa_neg'], 0)}.", "", 8.5)
        pdf.image(io.BytesIO(png_hist_mc(mc["vpl"])), x=30, w=150)
        fx = faixas_mc(d)
        pdf.tabela(pd.DataFrame({"Variável": fx["variavel"], "Pior caso": fx["minimo_pct"].map(lambda v: f"{num_br(v, 0)}%"),
                                 "Melhor caso": fx["maximo_pct"].map(lambda v: f"{num_br(v, 0)}%")}),
                   larguras=(100, 40, 40), alinhar=("LEFT", "RIGHT", "RIGHT"), tam=7.5)
        pdf.paragrafo(leitura_mc(mc), "I", 8)
        pa = perda_aceitavel(proj, d, res)
        pdf.paragrafo(f"Perda potencial se o negócio fechar: {brl(pa['potencial'])} (capital dos sócios + empréstimos − venda dos bens). "
                      f"Perda aceitável informada pelos sócios: {brl(pa['aceitavel'])}.", "", 8.5)
        bc = bass_comparacao(proj, res)
        if bc["ok"]:
            pdf.subtitulo(f"Curva de adoção de Bass (p = {num_br(bc['p'], 3)}; q = {num_br(bc['q'], 2)}; mercado potencial "
                          f"{num_br(bc['m'], 0)})")
            pdf.image(io.BytesIO(png_bass(bc["tabela"])), x=32, w=146)
        # ---- Defesa
        pr = avaliar_prontidao(proj, d, res, mc, preco)
        pdf.secao("PRONTIDÃO, BANCA SIMULADA E FORMALIZAÇÃO", nova_pagina=True)
        pdf.tabela(pd.DataFrame({"Nível": [str(k) for k, _, _ in NIVEIS_PRONTIDAO], "Degrau": [n for _, n, _ in NIVEIS_PRONTIDAO],
                                 "Critério": [c for _, _, c in NIVEIS_PRONTIDAO],
                                 "Situação": ["atingido" if k <= pr["nivel"] else ("próximo" if k == pr["nivel"] + 1 else "-")
                                              for k, _, _ in NIVEIS_PRONTIDAO]}),
                   larguras=(12, 48, 100, 20), alinhar=("CENTER", "LEFT", "LEFT", "CENTER"), tam=7.2)
        pdf.paragrafo(f"Nível de prontidão do negócio: {pr['nivel']} de 9.", "B", 9)
        perg = perguntas_banca(proj, d, res, mc, preco)
        resp = {str(r["pergunta"]): str(r["resposta"]) for _, r in d["banca_respostas"].iterrows()}
        pdf.subtitulo("Banca simulada")
        for i, q in enumerate(perg, start=1):
            pdf.paragrafo(f"{i}. [{q['tema']}] {q['pergunta']}", "B", 8.3)
            pdf.paragrafo("Resposta do grupo: " + (resp.get(q["pergunta"]) or "(sem resposta)"), "", 8.3)
    ps = porte_sugerido(res, d)
    pdf.subtitulo("Formalização")
    pdf.paragrafo(f"Enquadramento sugerido pela receita do Ano 1 ({brl(ps['receita'])}): {ps['porte']}.", "", 8.5)
    if not d["formalizacao"].empty:
        pdf.tabela(d["formalizacao"].rename(columns={"item": "Item", "status": "Situação", "observacao": "Observação"}),
                   larguras=(100, 26, 54), alinhar=("LEFT", "CENTER", "LEFT"), tam=7.5)


def secoes_v8(pdf, proj, d, res):
    """Seções da versão 8 no relatório completo."""
    if any(abs(x - 1) > 1e-9 for x in res.get("saz", [1.0] * 12)):
        pdf.secao("SAZONALIDADE DAS VENDAS", nova_pagina=True)
        ini = int(proj.get("saz_mes_inicio") or 1)
        nomes = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
        mm = res["meses"]
        pdf.tabela(pd.DataFrame({"Mês do plano": [f"{i + 1} ({nomes[(ini - 1 + i) % 12]})" for i in range(12)],
                                 "Fator": [num_br(x, 2) for x in res["saz"]], "Receita": mm["Receita"].map(brl),
                                 "Resultado operacional": mm["Resultado operacional"].map(brl),
                                 "Saldo de caixa": res["fluxo_mensal"]["Saldo final"].map(brl)}),
                   larguras=(35, 20, 40, 45, 40), alinhar=("LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"), tam=7.8)
    la = d.get("leituras_aluno", pd.DataFrame())
    if la is not None and not la.empty and (la["texto"].astype(str).str.strip() != "").any():
        pdf.secao("LEITURAS DOS RESULTADOS ESCRITAS PELO GRUPO", nova_pagina=True)
        la = la[la["texto"].astype(str).str.strip() != ""]
        pdf.tabela(pd.DataFrame({"Tabela": la["tabela"].str.replace("_", " ").str.capitalize(), "Leitura do grupo": la["texto"]}),
                   larguras=(40, 140), alinhar=("LEFT", "LEFT"), tam=8)
    est = d.get("swot_estrategias", pd.DataFrame())
    if not est.empty:
        pdf.secao("SWOT CRUZADA: ESTRATÉGIAS", nova_pagina=True)
        pdf.tabela(est.rename(columns={"quadrante": "Tipo", "estrategia": "Estratégia", "acao": "Ação", "prazo_meses": "Prazo (meses)"})
                   .assign(**{"Prazo (meses)": est["prazo_meses"].map(lambda v: num_br(v, 0))}),
                   larguras=(44, 60, 60, 16), alinhar=("LEFT", "LEFT", "LEFT", "CENTER"), tam=7.5)
    lc = lean_canvas(proj, d, res)
    pdf.secao("LEAN CANVAS", nova_pagina=True)
    pdf.tabela(pd.DataFrame([(k, "; ".join(str(x) for x in v) or "-") for k, v in lc.items()], columns=["Bloco", "Conteúdo"]),
               larguras=(45, 135), alinhar=("LEFT", "LEFT"), tam=7.8)
    dp = d.get("dados_publicos", pd.DataFrame())
    if not dp.empty or proj.get("setor_referencia"):
        pdf.secao("DADOS PÚBLICOS E REFERÊNCIAS DO SETOR", nova_pagina=True)
        if not dp.empty:
            pdf.subtitulo(f"{proj.get('municipio_nome') or ''} ({proj.get('uf_sigla') or ''})")
            pdf.tabela(pd.DataFrame({"Indicador": dp["rotulo"], "Valor": dp["valor"].map(lambda v: num_br(v, 2)), "Período": dp["periodo"],
                                     "Fonte": dp["fonte"]}), larguras=(55, 28, 20, 77), alinhar=("LEFT", "RIGHT", "CENTER", "LEFT"), tam=7.5)
        if proj.get("setor_referencia") and res["anos"][1]["rec_bruta"] > 0:
            try:
                cs = comparar_setor(proj["setor_referencia"], proj, d, res)
                pdf.subtitulo(f"Comparação com a média do setor: {proj['setor_referencia'].split(' · ')[1]}")
                fm = lambda i, v: pct(v) if ("÷" in i or "margem" in i.lower()) else brl(v)
                pdf.tabela(pd.DataFrame({"Indicador": cs["Indicador"], "Plano": [fm(i, v) for i, v in zip(cs["Indicador"], cs["Plano"])],
                                         "Setor (Brasil)": [fm(i, v) for i, v in zip(cs["Indicador"], cs["Setor (Brasil)"])],
                                         "Plano ÷ setor": cs["Plano ÷ setor"].map(lambda v: f"{num_br(v, 2)}x" if v == v else "n.a.")}),
                           larguras=(80, 35, 35, 30), tam=7.8)
                for a in leitura_setor(cs):
                    pdf.paragrafo(a, "I", 8)
                pdf.paragrafo(FONTE_PAC if proj["setor_referencia"].startswith("Comércio") else FONTE_PAS, "I", 7.5)
            except Exception:
                pass
        pdf.paragrafo("Sobrevivência em 5 anos (Sebrae, 2026): " + "; ".join(f"{k}: {pct(v, 1)}" for k, v in SOBREVIVENCIA_5A.items()) + ".",
                      "", 8.5)
    emp = d["emprestimos"]
    if not emp.empty and float(emp["valor"].sum()) > 0:
        pdf.secao("FINANCIAMENTO", nova_pagina=True)
        linhas_ = []
        for _, r in emp.iterrows():
            if float(r["valor"] or 0) > 0:
                rs = resumo_emprestimo(r["valor"], r["taxa_anual_pct"], r["prazo_meses"], r["carencia_meses"], r["sistema"])
                linhas_.append((r["descricao"], brl(r["valor"]), f"{num_br(r['taxa_anual_pct'], 2)}%", num_br(r["prazo_meses"], 0),
                                num_br(r["carencia_meses"] or 0, 0), r["sistema"] or "Price", brl(rs["juros"]), pct(rs["cet"])))
        pdf.tabela(pd.DataFrame(linhas_, columns=["Linha", "Valor", "Juros a.a.", "Prazo", "Carência", "Sistema", "Juros totais", "CET a.a."]),
                   larguras=(38, 22, 16, 13, 15, 16, 24, 18), tam=7.3)
    vl = valuation(res, proj, d)
    pdf.secao("VALUATION E RODADA DE INVESTIMENTO", nova_pagina=True)
    pdf.tabela(vl["metodos"].assign(**{"Valor do patrimônio dos sócios": vl["metodos"]["Valor do patrimônio dos sócios"].map(brl)}),
               larguras=(130, 50), alinhar=("LEFT", "RIGHT"))
    pre, inv = float(proj.get("rodada_pre_money") or 0), float(proj.get("rodada_investimento") or 0)
    if inv > 0 and pre > 0:
        ct = cap_table(d, pre, inv)
        pdf.subtitulo(f"Rodada: aporte de {brl(inv)} com pre-money de {brl(pre)}")
        pdf.tabela(ct.assign(**{"Participação antes": ct["Participação antes"].map(pct), "Participação depois": ct["Participação depois"].map(pct),
                                "Valor da participação": ct["Valor da participação"].map(brl)}), larguras=(60, 40, 40, 40))
    imp = [(rot, proj.get(k_)) for k_, rot, _ in ETAPAS_MUDANCA if proj.get(k_)]
    if imp or not d["impacto_indicadores"].empty:
        pdf.secao("IMPACTO SOCIAL E ODS", nova_pagina=True)
        for rot, v_ in imp:
            pdf.campo(rot, v_)
        if proj.get("ods"):
            pdf.campo("ODS", str(proj["ods"]).replace(";", "; "))
        ii = d["impacto_indicadores"]
        if not ii.empty:
            pdf.tabela(ii.rename(columns={"indicador": "Indicador", "tipo": "Tipo", "linha_base": "Situação atual", "meta": "Meta", "ods": "ODS"}),
                       larguras=(55, 28, 30, 30, 37), alinhar=("LEFT",) * 5, tam=7.5)


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
    from pnbox_motor import brl, num_br, pct, div, QUADRANTES, mercado, semaforo, capacidade_demanda, custo_hora  # noqa
    from pnbox_estrategia import canvas_blocos, porter_tabela, porter_leitura, ge_tabela, bcg_tabela  # noqa
    from pnbox_estrategia import QUADRANTES_BCG, checar_smart, PERSPECTIVAS, bsc_indicadores_plano, simular_reforma  # noqa
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

    ob = d.get("objetivos")
    if any(proj.get(k) for k in ("missao", "visao", "valores")) or (ob is not None and not ob.empty):
        pdf.secao("MISSÃO, VISÃO, VALORES E OBJETIVOS", nova_pagina=True)
        pdf.campo("Missão", proj.get("missao", ""))
        pdf.campo("Visão", proj.get("visao", ""))
        pdf.campo("Valores", proj.get("valores", ""))
        if ob is not None and not ob.empty:
            pdf.subtitulo("Objetivos estratégicos (critérios SMART)")
            pdf.tabela(pd.DataFrame({
                "Objetivo": ob["objetivo"], "Perspectiva": ob["perspectiva"], "Indicador e meta": [
                    f"{i_}: {num_br(m, 1)} {u}" for i_, m, u in zip(ob["indicador"], ob["meta"], ob["unidade"])],
                "Prazo": ob["prazo_meses"].map(lambda v: f"mês {num_br(v, 0)}"), "Responsável": ob["responsavel"],
                "SMART": ["".join(L if ok else "-" for L, (ok, _) in checar_smart(o).items()) for _, o in ob.iterrows()]}),
                larguras=(54, 26, 46, 14, 26, 14), alinhar=("LEFT", "LEFT", "LEFT", "CENTER", "LEFT", "CENTER"), tam=7)
            pdf.paragrafo("SMART: S = específico, M = mensurável, A = atingível, R = relevante, T = temporal. Um hífen indica o "
                          "critério que ainda precisa ser ajustado.", "I", 7.5)

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

    # ---------------- ANÁLISE ESTRATÉGICA ----------------
    pdf.secao("BUSINESS MODEL CANVAS", nova_pagina=True)
    canvas_pdf(pdf, canvas_blocos(proj, d, res))
    pt = porter_tabela(d)
    if (pt["Intensidade"] > 0).any():
        pdf.secao("CINCO FORÇAS DE PORTER", nova_pagina=True)
        img = png_porter(pt)
        if img:
            pdf.image(io.BytesIO(img), x=60, w=90)
        pdf.tabela(pd.DataFrame({"Força": pt["Força"], "Nota (1-5)": pt["Intensidade"].map(lambda v: num_br(v, 0)),
                                 "Justificativa": pt["Justificativa"].str.replace("(sugestão) ", "", regex=False),
                                 "Estratégia": pt["Estratégia"]}), larguras=(40, 16, 62, 62),
                   alinhar=("LEFT", "CENTER", "LEFT", "LEFT"), tam=7.5)
        pdf.paragrafo(porter_leitura(pt), "I", 8.5)
    ge_ = ge_tabela(d)
    if not ge_.empty:
        pdf.secao("MATRIZ GE / McKINSEY", nova_pagina=True)
        img = png_ge(ge_)
        if img:
            pdf.image(io.BytesIO(img), x=35, w=140)
        pdf.tabela(ge_[["Segmento", "Atratividade", "Força competitiva", "Recomendação"]].assign(
            Segmento=[f"{k}. {s}" for k, s in enumerate(ge_["Segmento"], start=1)],
            Atratividade=ge_["Atratividade"].map(lambda v: num_br(v, 1)),
            **{"Força competitiva": ge_["Força competitiva"].map(lambda v: num_br(v, 1))}),
            larguras=(70, 30, 35, 45), alinhar=("LEFT", "CENTER", "CENTER", "LEFT"), tam=7.5)
        pdf.paragrafo("Atratividade = média das notas de tamanho, lucro, longevidade e concorrência do segmento. Força competitiva = "
                      "informada pelo grupo ou média entre a aderência ao segmento e a nota da empresa frente aos concorrentes.", "I", 7.5)
    lim = float(proj.get("bcg_limite_crescimento_pct") or 10)
    bcg_ = bcg_tabela(d, res, lim)
    if not bcg_.empty:
        pdf.secao("MATRIZ BCG", nova_pagina=True)
        img = png_bcg(bcg_, lim)
        if img:
            pdf.image(io.BytesIO(img), x=35, w=140)
        pdf.tabela(pd.DataFrame({"Produto": bcg_["Produto"], "Cresc. mercado": bcg_["Crescimento do mercado (%)"].map(lambda v: f"{num_br(v, 1)}%"),
                                 "Part. relativa": bcg_["Participação relativa"].map(lambda v: f"{num_br(v, 2)}x"),
                                 "Receita anual": bcg_["Receita anual"].map(brl), "Quadrante": bcg_["Quadrante"]}),
                   larguras=(66, 24, 24, 34, 32), alinhar=("LEFT", "RIGHT", "RIGHT", "RIGHT", "LEFT"), tam=7.5)
        for q, (_, txt) in QUADRANTES_BCG.items():
            itens = bcg_.loc[bcg_["Quadrante"] == q, "Produto"].tolist()
            if itens:
                pdf.paragrafo(f"{q} ({', '.join(itens)}): {txt}", "", 8)
    if ob is not None and not ob.empty:
        pdf.secao("BALANCED SCORECARD", nova_pagina=True)
        img = png_bsc(ob)
        if img:
            pdf.image(io.BytesIO(img), x=55, w=100)
        ind = bsc_indicadores_plano(proj, d, res)
        for per in PERSPECTIVAS:
            sub = ob[ob["perspectiva"] == per]
            pdf.subtitulo(f"Perspectiva {per.lower()}")
            if not sub.empty:
                pdf.tabela(pd.DataFrame({"Objetivo": sub["objetivo"], "Indicador": sub["indicador"],
                                         "Meta": [f"{num_br(m, 1)} {u}" for m, u in zip(sub["meta"], sub["unidade"])],
                                         "Prazo": sub["prazo_meses"].map(lambda v: f"mês {num_br(v, 0)}"),
                                         "Iniciativa": sub["iniciativa"]}), larguras=(55, 38, 28, 16, 43),
                           alinhar=("LEFT", "LEFT", "RIGHT", "CENTER", "LEFT"), tam=7.2)
            if ind.get(per):
                pdf.paragrafo("Indicadores do plano: " + "; ".join(f"{n} = {v}" for n, v in ind[per]), "I", 7.5)

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
    if any(s for _, s, _ in etapas_plano_mkt(proj, d)) or proj.get("mix_nucleo") or proj.get("mix_estrategia_preco"):
        pdf.secao("PLANO DE MARKETING", nova_pagina=True)
        for i, sec in enumerate(conteudo_plano_mkt(proj, d, res), 1):
            pdf.subtitulo(f"{i}. {sec['titulo']}")
            pdf.paragrafo(sec["texto"])
            tb = sec["tabela"]
            if tb is not None and not tb.empty:
                n = len(tb.columns)
                larg = {2: (40, 140), 4: (75, 45, 35, 25), 5: (50, 30, 25, 45, 30)}.get(n)
                if n == 4 and "Mês" in tb.columns:
                    larg = (15, 95, 40, 30)
                pdf.tabela(tb, larguras=larg, alinhar=tuple(["LEFT"] * n), tam=7.5)
        co = coerencia_mix(proj, d, res)
        if not co.empty:
            pdf.subtitulo("Coerência do mix")
            pdf.tabela(co.assign(Situação=co["Situação"].str.replace("✅ ", "").str.replace("⚠️ ", "")),
                       larguras=(55, 25, 100), alinhar=("LEFT", "LEFT", "LEFT"), tam=7.5)
        for rot, k in [("Produto", "mkt_produto"), ("Preço", "mkt_preco"), ("Praça", "mkt_praca"), ("Promoção", "mkt_promocao")]:
            if proj.get(k):
                pdf.subtitulo(f"Justificativa: {rot.lower()}")
                pdf.paragrafo(proj[k])
    elif any(proj.get(k) for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")):
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
        for k in res["ANOS"]:
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
    N = res["N"]
    pdf.subtitulo(f"Demonstração do resultado do exercício (Anos 1 a {N})" if N > 1 else "Demonstração do resultado do exercício (Ano 1)")
    dre = extras["dre"]
    for bloco in [res["ANOS"][i:i + 5] for i in range(0, N, 5)]:
        cols = ["Conta"] + [f"Ano {a}" for a in bloco] + (["% Receita (Ano 1)"] if bloco[0] == 1 else [])
        lw = 62 + (0 if bloco[0] == 1 else 18)
        larg = (lw,) + tuple((180 - lw - (18 if bloco[0] == 1 else 0)) / len(bloco) for _ in bloco) + ((18,) if bloco[0] == 1 else ())
        pdf.tabela(dre[cols], larguras=larg, tam=7.5 if len(bloco) <= 3 else 6.8)

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
    pdf.subtitulo("Resumo financeiro (Ano 1, padrão Sebrae)")
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
    pdf.subtitulo(f"Avaliação do investimento (horizonte de {res['N']} ano{'s' if res['N'] > 1 else ''})")
    fl = res["fluxos"]
    pdf.tabela(pd.DataFrame({
        "Indicador": [f"Fluxos de caixa do projeto (Ano 0 a {res['N']})", "VPL à TMA", "TIR", "Índice de lucratividade",
                      "Payback descontado"],
        "Valor": [" | ".join(brl(v) for v in fl), brl(res["vpl"]),
                  pct(res["tir"]) if res["tir"] is not None else "n.a.",
                  f"{num_br(res['il'])}x" if res["il"] == res["il"] else "n.a.",
                  f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else f"acima de {res['MESES']} meses"]}),
        larguras=(55, 125), alinhar=("LEFT", "LEFT"))
    pdf.paragrafo(f"Fluxo do projeto: Ano 0 = investimento total; Anos 1 a {res['N']} = EBITDA" +
                  (f" (Ano {res['N']} inclui a recuperação do capital de giro)." if proj.get("recuperar_giro") else "."), "I", 7.5)
    pdf.subtitulo("Matriz de indicadores em 5 blocos")
    ind = extras["indicadores"]
    for bloco, g in ind.groupby("Bloco", sort=False):
        pdf.fonte("B", 8.5)
        if pdf.get_y() > 250:
            pdf.add_page()
        pdf.cell(0, 5, pdf.t(bloco), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        for bloco_a in [res["ANOS"][i:i + 4] for i in range(0, res["N"], 4)]:
            cols = ["Indicador"] + (["Fórmula"] if bloco_a[0] == 1 else []) + [f"Ano {a}" for a in bloco_a] + \
                   (["Var. Ano 2/Ano 1"] if bloco_a[0] == 1 and res["N"] >= 2 else [])
            fixa = 42 + (46 if bloco_a[0] == 1 else 0) + (17 if "Var. Ano 2/Ano 1" in cols else 0)
            larg = [42] + ([46] if bloco_a[0] == 1 else []) + [(180 - fixa) / len(bloco_a)] * len(bloco_a) + \
                   ([17] if "Var. Ano 2/Ano 1" in cols else [])
            pdf.tabela(g[cols], larguras=tuple(larg), alinhar=tuple(["LEFT"] * (2 if bloco_a[0] == 1 else 1) +
                                                                     ["RIGHT"] * (len(cols) - (2 if bloco_a[0] == 1 else 1))), tam=7)

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

    if float(proj.get("reforma_b2b_pct") or 0) > 0 or proj.get("anexo_simples"):
        an = proj.get("anexo_simples") or "Anexo III - Serviços"
        ano_r = int(float(proj.get("reforma_ano") or 2027))
        cc = float(proj.get("reforma_compras_credito") or -1)
        rr = simular_reforma(a1["rec_bruta"], res["aliq"], an, ano_r, float(proj.get("reforma_cbs_pct") or 11.3) / 100,
                             float(proj.get("reforma_ibs_pct") or 0.1) / 100, float(proj.get("reforma_b2b_pct") or 0) / 100,
                             a1["custo_dir"] + a1["fixos"] if cc < 0 else cc)
        pdf.secao("REFORMA TRIBUTÁRIA: SIMULAÇÃO", nova_pagina=True)
        pdf.paragrafo(f"Ano simulado: {ano_r}; {an}; CBS {num_br(proj.get('reforma_cbs_pct') or 11.3, 2)}% e IBS "
                      f"{num_br(proj.get('reforma_ibs_pct') or 0.1, 2)}%; vendas para empresas que aproveitam crédito: "
                      f"{num_br(proj.get('reforma_b2b_pct') or 0, 0)}%.", "", 8.5)
        pdf.tabela(pd.DataFrame({
            "Item": ["DAS", "IBS + CBS sobre as vendas", "(-) Crédito das compras", "Total (DAS + IBS/CBS)", "Crédito aos clientes empresas"],
            "Simples puro": [brl(rr["das_puro"]), "no DAS", "-", brl(rr["puro_total"]), brl(rr["puro_credito_cliente"])],
            "Simples híbrido": [brl(rr["das_hibrido"]), brl(rr["iva_debito"]), brl(-rr["iva_credito"]), brl(rr["hib_total"]),
                                brl(rr["hib_credito_cliente"])]}), larguras=(80, 50, 50), alinhar=("LEFT", "RIGHT", "RIGHT"))
        pdf.paragrafo("Simulação didática e simplificada. Fontes: LC 214/2025, art. 41; partilha do DAS da LC 123/2006; CBS de "
                      "referência de 11,3% para 2027 calculada pelo TCU, sujeita à definição do Senado. Confirme com um contador.", "I", 7.5)

    secoes_v71(pdf, proj, d, res)
    secoes_v8(pdf, proj, d, res)

    if str(proj.get("conclusao") or "").strip():
        pdf.secao("CONCLUSÃO E PARECER FINAL", nova_pagina=True)
        for par in str(proj["conclusao"]).split("\n"):
            pdf.paragrafo(par, "", 10)

    out = pdf.output()
    return bytes(out)
