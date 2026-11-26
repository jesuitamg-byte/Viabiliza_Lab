"""
Funções das versões 8.6 a 9.0 (sem interface): sazonalidade, sensibilidade (tornado),
taxas do Banco Central, pitch em PowerPoint, modelos por setor, mapa de concorrentes,
árvore de decisão e acessibilidade.
"""
import pandas as pd

from pnbox_mix import _n, _p, comparar_concorrentes, indice_preco_medio, resumo_mix_linhas
from pnbox_motor import executar_motor, fatores_sazonais

MESES_NOME = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
              "Novembro", "Dezembro"]

# ------------------------------------------------------------------------------
# SAZONALIDADE (v8.6)
# ------------------------------------------------------------------------------
# Perfis ILUSTRATIVOS (100 = mês médio). Servem de ponto de partida: o grupo deve ajustar com dados do
# próprio negócio, de concorrentes, de fornecedores ou de pesquisas do IBGE.
PERFIS_SAZONAIS = {
    "Sem sazonalidade": [100] * 12,
    "Varejo de presentes e vestuário (pico no Natal)": [80, 78, 92, 90, 105, 95, 98, 100, 92, 98, 105, 167],
    "Alimentação e cafeteria (mais fraco nas férias)": [92, 95, 102, 100, 103, 98, 95, 102, 100, 103, 102, 108],
    "Sorvetes e bebidas geladas (verão)": [140, 130, 115, 95, 80, 70, 70, 75, 90, 105, 120, 110],
    "Cursos e educação (fraco nas férias escolares)": [70, 115, 120, 110, 105, 85, 75, 110, 110, 105, 100, 95],
    "Serviços a empresas (fraco em janeiro e dezembro)": [80, 95, 105, 105, 105, 100, 100, 105, 105, 105, 105, 90],
    "Turismo e lazer (férias e feriados)": [150, 115, 85, 90, 80, 90, 130, 85, 85, 95, 90, 105],
}


def tabela_perfil(nome: str) -> pd.DataFrame:
    return pd.DataFrame({"mes": list(range(1, 13)), "indice": [float(x) for x in PERFIS_SAZONAIS[nome]]})


def resumo_sazonal(proj: dict, res: dict) -> dict:
    """Meses do plano mais fortes e mais fracos, amplitude e o mês de caixa mais apertado."""
    mm = res["meses"]
    fm = res["fluxo_mensal"]
    ini = int(proj.get("saz_mes_inicio") or 1)
    nome = lambda m: f"mês {m + 1} ({MESES_NOME[(ini - 1 + m) % 12].lower()})"
    rec = mm["Receita"].tolist()
    i_max, i_min = rec.index(max(rec)), rec.index(min(rec))
    j = int(fm["Saldo final"].idxmin()) if not fm.empty else 0
    return {"forte": nome(i_max), "fraco": nome(i_min), "amplitude": (max(rec) / min(rec) - 1) if min(rec) > 0 else float("nan"),
            "caixa_min": float(fm["Saldo final"].min()) if not fm.empty else float("nan"), "mes_caixa": nome(j),
            "neg": [nome(m) for m in range(12) if mm["Resultado operacional"].iloc[m] < 0]}


def leitura_sazonalidade(proj: dict, res: dict) -> list[str]:
    if all(abs(x - 1) < 1e-9 for x in res.get("saz", [1.0] * 12)):
        return ["Sem sazonalidade: as vendas crescem de forma regular mês a mês."]
    r = resumo_sazonal(proj, res)
    out = [f"O mês mais forte é o {r['forte']} e o mais fraco, o {r['fraco']}: a receita do melhor mês supera a do pior em "
           f"{_p(r['amplitude'], 0)}."]
    if r["neg"]:
        out.append("Dá prejuízo operacional em: " + ", ".join(r["neg"]) + ". É nesses meses que a reserva de caixa é usada.")
    out.append(f"O caixa fica mais apertado no {r['mes_caixa']}, com saldo de R$ {_n(r['caixa_min'], 2)}"
               + (": falta dinheiro e o capital de giro precisa aumentar." if r["caixa_min"] < 0 else "."))
    out.append("O total do ano não muda com a sazonalidade; muda quando o dinheiro entra.")
    return out


# ------------------------------------------------------------------------------
# SENSIBILIDADE DO VPL: TORNADO, PONTO DE RUPTURA E MAPA PREÇO × VOLUME (v8.7)
# ------------------------------------------------------------------------------
VARS_TORNADO = [("Preço de venda", "preco"), ("Volume de vendas", "volume"), ("Custo de materiais e mercadorias", "materiais"),
                ("Salários e encargos", "salarios"), ("Demais custos fixos", "fixos"),
                ("Crescimento anual (pontos percentuais)", "cresc")]
# premissa do placar de evidências correspondente a cada variável
PREMISSA_DE = {"Preço de venda": "Preço de venda", "Volume de vendas": "Volume de vendas do Ano 1",
               "Custo de materiais e mercadorias": "Custo de materiais e mercadorias", "Salários e encargos": "Salários e encargos",
               "Demais custos fixos": "Demais custos fixos", "Crescimento anual (pontos percentuais)": "Crescimento das vendas"}


def _rs(x: float) -> str:
    return ("-R$ " if x < 0 else "R$ ") + _n(abs(x), 0)


def vpl_variado(res: dict, preco=0.0, volume=0.0, materiais=0.0, salarios=0.0, fixos=0.0, cresc=0.0) -> float:
    """VPL com variações (0,10 = +10%), pela mesma lógica do Monte Carlo (fluxo do projeto = EBITDA)."""
    fp, fv, fm, fs, ff = 1 + preco, 1 + volume, 1 + materiais, 1 + salarios, 1 + fixos
    g_a, tma, N = res["g_a"], res["tma"], res["N"]
    adj = (1 + g_a + cresc) / (1 + g_a) if (1 + g_a) else 1.0
    vpl = -res["investimento_total"]
    for k in res["ANOS"]:
        a = res["anos"][k]
        gk = adj ** (k - 1)
        rec = a["rec_bruta"] * fp * fv * gk
        var_ratio = (a["impostos"] + a["outras_var"]) / a["rec_bruta"] if a["rec_bruta"] else 0.0
        ebitda = rec - a["custo_dir"] * fv * fm * gk - rec * var_ratio - a["pessoal"] * fs - a["fixos"] * ff
        if k == N and res.get("recupera_giro"):
            ebitda += res.get("capital_giro", 0.0)
        vpl += ebitda / (1 + tma) ** k
    return float(vpl)


def tornado(res: dict, faixas: pd.DataFrame) -> pd.DataFrame:
    """VPL no pior e no melhor valor de cada variável (faixas do Monte Carlo), do maior para o menor efeito."""
    fx = {r["variavel"]: (float(r["minimo_pct"]) / 100, float(r["maximo_pct"]) / 100) for _, r in faixas.iterrows()}
    base = vpl_variado(res)
    linhas = []
    for nome, ch in VARS_TORNADO:
        lo, hi = fx.get(nome, (-0.10, 0.10))
        v_lo, v_hi = vpl_variado(res, **{ch: lo}), vpl_variado(res, **{ch: hi})
        linhas.append({"Variável": nome, "Variação mínima": lo, "Variação máxima": hi, "VPL no mínimo": v_lo,
                       "VPL no máximo": v_hi, "Amplitude": abs(v_hi - v_lo), "VPL base": base})
    return pd.DataFrame(linhas).sort_values("Amplitude", ascending=False).reset_index(drop=True)


def ponto_ruptura(res: dict, ch: str) -> float:
    """Variação (fração) da variável que zera o VPL; nan se não houver dentro de -100% a +500%."""
    base = vpl_variado(res)
    sinal_bom = vpl_variado(res, **{ch: 0.01}) - base       # efeito de aumentar 1%
    if abs(sinal_bom) < 1e-9:
        return float("nan")
    lo, hi = (-1.0, 0.0) if (base > 0) == (sinal_bom > 0) else (0.0, 5.0)
    f = lambda x: vpl_variado(res, **{ch: x})
    if f(lo) * f(hi) > 0:
        return float("nan")
    for _ in range(60):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def tabela_rupturas(res: dict) -> pd.DataFrame:
    linhas = []
    for nome, ch in VARS_TORNADO[:5]:
        x = ponto_ruptura(res, ch)
        linhas.append({"Variável": nome, "Variação que zera o VPL": x})
    return pd.DataFrame(linhas)


def mapa_preco_volume(res: dict, passos=(-0.20, -0.15, -0.10, -0.05, 0.0, 0.05, 0.10, 0.15, 0.20)) -> pd.DataFrame:
    linhas = []
    for p in passos:
        for v in passos:
            linhas.append({"Preço": p, "Volume": v, "VPL": vpl_variado(res, preco=p, volume=v)})
    return pd.DataFrame(linhas)


def valide_primeiro(tor: pd.DataFrame, placar_tab: pd.DataFrame) -> list[tuple[str, str]]:
    """Cruza as variáveis que mais mexem no VPL com o grau de evidência de cada premissa."""
    st_ = dict(zip(placar_tab["premissa"], placar_tab["status"])) if placar_tab is not None and not placar_tab.empty else {}
    out = []
    for _, r in tor.head(3).iterrows():
        prem = PREMISSA_DE.get(r["Variável"], r["Variável"])
        status = st_.get(prem, "Estimada (sem fonte)")
        if status.startswith("Estimada"):
            out.append(("critico", f"{r['Variável']} está entre as variáveis que mais mexem no VPL e ainda é só estimativa: "
                        "valide isto primeiro (entrevista, pré-venda, orçamento ou dado de fonte confiável)."))
        elif status.startswith("Pesquisa"):
            out.append(("atencao", f"{r['Variável']} mexe muito no VPL e se apoia em pesquisa secundária; um teste com clientes "
                        "ou fornecedores daria mais segurança."))
        else:
            out.append(("ok", f"{r['Variável']} mexe muito no VPL, mas essa premissa já foi testada com clientes."))
    return out


def leitura_tornado(tor: pd.DataFrame, rup: pd.DataFrame) -> list[str]:
    if tor.empty:
        return []
    t = tor.iloc[0]
    out = [f"A variável que mais mexe no VPL é {t['Variável'].lower()}: entre o pior e o melhor caso o VPL vai de "
           f"{_rs(t['VPL no mínimo'])} a {_rs(t['VPL no máximo'])}."]
    neg = tor[(tor["VPL no mínimo"] < 0) | (tor["VPL no máximo"] < 0)]["Variável"].tolist()
    if neg:
        out.append("Na faixa informada, o VPL fica negativo por conta de: " + ", ".join(x.lower() for x in neg) + ".")
    ok = rup.dropna()
    if not ok.empty:
        r = ok.iloc[ok["Variação que zera o VPL"].abs().argmin()]
        out.append(f"A margem de segurança mais estreita está em {r['Variável'].lower()}: uma variação de "
                   f"{_p(r['Variação que zera o VPL'], 0)} já zera o VPL.")
    return out


# ------------------------------------------------------------------------------
# TAXAS DO BANCO CENTRAL E TMA EM DEGRAUS (v8.8)
# ------------------------------------------------------------------------------
API_BCB = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados/ultimos/1?formato=json"
SERIES_BCB = {
    "bcb_selic": (432, "Meta da taxa Selic (% ao ano)"),
    "bcb_ipca12": (13522, "IPCA acumulado em 12 meses (%)"),
}
# Últimos valores conhecidos, usados quando não há internet (conferidos em nov. 2026)
BCB_RESERVA = {
    "bcb_selic": (13.75, "04/11/2026", "Copom, 282ª reunião"),
    "bcb_ipca12": (4.63, "10/2026", "IBGE, divulgado em 12/11/2026"),
}
PREMIOS_RISCO = {
    "Baixo (negócio conhecido, demanda estável, pouco investimento)": 4.0,
    "Médio (negócio comum, alguma concorrência)": 6.0,
    "Alto (produto novo, mercado não testado, muita concorrência)": 9.0,
    "Muito alto (inovação, tecnologia, dependência de um cliente)": 12.0,
}


def buscar_bcb(chave: str) -> tuple:
    """(valor, data, fonte) da série do SGS; se a consulta falhar, devolve o último valor conhecido."""
    from pnbox_v8 import baixar_json
    codigo, _ = SERIES_BCB[chave]
    try:
        js = baixar_json(API_BCB.format(codigo=codigo))
        return float(str(js[-1]["valor"]).replace(",", ".")), js[-1]["data"], f"Banco Central do Brasil, SGS série {codigo}"
    except Exception:
        v, dt, fonte = BCB_RESERVA[chave]
        return v, dt, f"{fonte} (valor guardado no programa)"


def tma_degraus(selic: float, premio_risco: float, premio_pequeno: float) -> dict:
    """TMA pelo método de acumulação (build-up): taxa livre de risco + prêmios, em pontos percentuais ao ano."""
    total = selic + premio_risco + premio_pequeno
    return {"livre": selic, "risco": premio_risco, "pequeno": premio_pequeno, "tma": total}


def taxa_real(nominal_pct: float, inflacao_pct: float) -> float:
    """Taxa real pela equação de Fisher, em %."""
    return ((1 + nominal_pct / 100) / (1 + inflacao_pct / 100) - 1) * 100


def leitura_taxas(proj: dict, d: dict, res: dict, selic: float, ipca: float, tma_sug: float) -> list[str]:
    tma = float(proj.get("tma_anual_pct") or 0)
    out = [f"Com Selic de {_n(selic, 2)}% e inflação de {_n(ipca, 2)}% em 12 meses, o juro real da renda fixa está em torno de "
           f"{_n(taxa_real(selic, ipca), 1)}% ao ano: é o mínimo que o dinheiro dos sócios renderia sem risco."]
    if tma < selic:
        out.append(f"A TMA do plano ({_n(tma, 1)}%) está abaixo da Selic: o projeto está sendo comparado com uma alternativa pior "
                   "do que deixar o dinheiro aplicado. O VPL fica otimista.")
    elif tma < tma_sug - 0.5:
        out.append(f"A TMA do plano ({_n(tma, 1)}%) está abaixo da sugerida pelos degraus ({_n(tma_sug, 1)}%): o risco do negócio "
                   "pode estar subestimado.")
    else:
        out.append(f"A TMA do plano ({_n(tma, 1)}%) cobre a Selic e os prêmios de risco escolhidos.")
    inf = float(proj.get("inflacao_custos_anual_pct") or 0)
    if abs(inf - ipca) > 1.5:
        out.append(f"A inflação de custos do plano ({_n(inf, 1)}%) está longe do IPCA atual ({_n(ipca, 2)}%); vale justificar a escolha.")
    emp = d.get("emprestimos")
    if emp is not None and not emp.empty and float(emp["valor"].sum()) > 0:
        tx = float((emp["valor"] * emp["taxa_anual_pct"]).sum() / emp["valor"].sum())
        if tx < selic:
            out.append(f"Os empréstimos estão com taxa média de {_n(tx, 1)}% ao ano, abaixo da Selic: só é realista em linhas "
                       "subsidiadas (como Pronampe ou fundos públicos). Confira a condição real.")
    return out


# ------------------------------------------------------------------------------
# PITCH EM POWERPOINT (v8.9)
# ------------------------------------------------------------------------------
ROTEIROS_PITCH = {
    "Banca acadêmica": ["capa", "problema", "solucao", "mercado", "concorrencia", "marketing", "operacao", "numeros",
                        "sensibilidade", "riscos", "fechamento"],
    "Investidor": ["capa", "problema", "solucao", "mercado", "tracao", "concorrencia", "numeros", "investimento", "equipe",
                   "fechamento"],
    "Edital de fomento": ["capa", "problema", "solucao", "impacto", "mercado", "equipe", "investimento", "numeros", "fechamento"],
}
TEMPO_SLIDE = {"capa": 20, "problema": 60, "solucao": 60, "mercado": 50, "concorrencia": 50, "marketing": 60, "operacao": 40,
               "numeros": 70, "sensibilidade": 50, "riscos": 40, "tracao": 50, "investimento": 50, "equipe": 40, "impacto": 60,
               "fechamento": 30}


def _rs2(x):
    return ("-R$ " if x < 0 else "R$ ") + _n(abs(x), 0)


def conteudo_pitch(proj: dict, d: dict, res: dict) -> dict:
    """Texto de cada slide (título, tópicos e notas do apresentador), independente do formato."""
    from pnbox_validacao import mercado
    a1 = res["anos"][1]
    N = res["N"]
    pers, props = d["personas"], d.get("propostas")
    c = {}
    c["capa"] = (proj.get("nome_negocio") or "Plano de negócio",
                 [x for x in [proj.get("cidade_uf"), proj.get("turma"), str(proj.get("autores") or "").replace("\n", ", ")] if x],
                 "Apresente o grupo e diga em uma frase o que o negócio faz e para quem.")
    itens = []
    for _, p_ in pers.head(2).iterrows():
        itens.append(f"{p_['nome']}, {_n(p_['idade'])} anos: {p_['desafios'] or p_['citacao']}")
    c["problema"] = ("O problema", itens or ["Descreva quem é o cliente e que dor ele sente."],
                     "Conte uma situação concreta vivida pela persona. A banca precisa sentir o problema antes de ouvir a solução.")
    itens = []
    if str(proj.get("mix_nucleo") or "").strip():
        itens.append(f"O que o cliente compra: {proj['mix_nucleo']}")
    if props is not None and not props.empty:
        itens.append(f"Oferecemos {props['oferecem'].iloc[0]}")
    top = res["cat"].sort_values("receita_mensal", ascending=False)["descricao"].head(4).tolist() if not res["cat"].empty else []
    if top:
        itens.append("Principais itens: " + ", ".join(top))
    c["solucao"] = ("A solução", itens or ["Descreva a proposta de valor."],
                    "Mostre como a solução alivia exatamente a dor apresentada no slide anterior.")
    mk = mercado(proj, res)
    c["mercado"] = ("O mercado", [f"Mercado total: {_rs2(mk['tam'])}", f"Mercado alcançável: {_rs2(mk['sam'])}",
                                  f"Meta de participação: {_rs2(mk['som'])}",
                                  f"A receita do Ano 1 equivale a {_p(mk['pct_som'], 0) if mk['pct_som'] == mk['pct_som'] else '-'} da meta"]
                    if mk["preenchido"] else ["Preencha o tamanho do mercado."],
                    "Diga de onde vieram os números (IBGE, pesquisa, concorrentes). Número sem fonte é a primeira pergunta da banca.")
    comp = comparar_concorrentes(res["cat"], d.get("precos_concorrentes")) if "precos_concorrentes" in d else pd.DataFrame()
    itens = [f"Concorrentes analisados: {', '.join(d['concorrentes']['nome'].head(4))}"] if not d["concorrentes"].empty else []
    if not comp.empty:
        idx = indice_preco_medio(res["cat"], comp)
        itens.append(f"Nossos preços: {_p(abs(idx - 1), 0)} {'acima' if idx >= 1 else 'abaixo'} da média")
    if str(proj.get("posic_diferencial") or "").strip():
        itens.append(f"Diferencial: {proj['posic_diferencial']}")
    c["concorrencia"] = ("Concorrência e diferencial", itens or ["Cadastre os concorrentes."],
                         "Não diga que não há concorrentes: mostre que conhece as alternativas e por que o cliente trocaria.")
    lm = resumo_mix_linhas(proj, d, res)
    c["marketing"] = ("Como vamos vender", [f"{r}: {t}" for r, t in lm[:5]] or ["Preencha o mix de marketing."],
                      "Explique o caminho do cliente: como conhece, como compra, como paga e por que volta.")
    c["operacao"] = ("Operação", [x for x in [proj.get("loc_endereco"), proj.get("op_processos")] if str(x or "").strip()][:3]
                     or ["Descreva a localização e os processos."], "Mostre que o grupo sabe como o serviço ou produto é entregue no dia a dia.")
    c["numeros"] = ("Os números", [f"Investimento total: {_rs2(res['investimento_total'])}", f"Receita no Ano 1: {_rs2(a1['rec_bruta'])}",
                                   f"Resultado no Ano 1: {_rs2(a1['resultado'])}", f"VPL em {N} anos: {_rs2(res['vpl'])}",
                                   f"TIR: {_p(res['tir'], 1) if res['tir'] is not None and res['tir'] == res['tir'] else 'n.a.'}",
                                   f"Payback descontado: {_n(res['payback_desc'], 1) + ' meses' if res['payback_desc'] == res['payback_desc'] else 'acima do horizonte'}"],
                    "Leia os três números principais em voz alta (investimento, VPL e payback) e diga o que significam.")
    try:
        from pnbox_validacao import faixas_mc
        tor = tornado(res, faixas_mc(d))
        t = tor.iloc[0]
        c["sensibilidade"] = ("E se der errado?", [f"Variável mais sensível: {t['Variável'].lower()}",
                                                  f"VPL entre {_rs2(t['VPL no mínimo'])} e {_rs2(t['VPL no máximo'])} na faixa testada"]
                              + [f"{r['Variável']}: zera o VPL com {_p(r['Variação que zera o VPL'], 0)}"
                                 for _, r in tabela_rupturas(res).dropna().head(2).iterrows()],
                              "Mostrar o pior caso transmite segurança. Diga o que o grupo fará se essa variável vier abaixo do esperado.")
    except Exception:
        c["sensibilidade"] = ("E se der errado?", ["Veja a ferramenta Sensibilidade do VPL."], "")
    rk = d["riscos"]
    itens = []
    if not rk.empty:
        rk2 = rk.assign(sev=rk["probabilidade"] * rk["impacto"]).sort_values("sev", ascending=False).head(3)
        itens = [f"{r['risco']}: {r['mitigacao'] or 'mitigação a definir'}" for _, r in rk2.iterrows()]
    c["riscos"] = ("Riscos e como lidamos com eles", itens or ["Cadastre os riscos."], "Escolha os dois riscos mais graves e a resposta para cada um.")
    ev = d.get("entrevistas")
    itens = [f"{len(ev)} entrevistas com clientes"] if ev is not None and not ev.empty else []
    if d.get("precos_vw") is not None and not d["precos_vw"].empty:
        itens.append(f"{len(d['precos_vw'])} respostas na pesquisa de preço")
    c["tracao"] = ("Validação e tração", itens or ["Mostre as evidências: entrevistas, pré-vendas, pilotos."],
                   "Investidor compra evidência. Mostre o que já foi testado com clientes reais.")
    c["investimento"] = ("Quanto precisamos e para quê", [f"Investimento fixo: {_rs2(res['inv_fixo'])}", f"Pré-operacional: {_rs2(res['inv_pre'])}",
                                                         f"Capital de giro: {_rs2(res['capital_giro'])}",
                                                         f"Capital próprio: {_rs2(res['capital_proprio'])}; terceiros: {_rs2(res['capital_terceiros'])}"],
                         "Diga exatamente o valor pedido e em que ele será usado.")
    soc = d["socios"]
    c["equipe"] = ("Equipe", [f"{r['nome']}: {r['funcao']}" for _, r in soc.head(5).iterrows()] or ["Apresente os sócios."],
                   "Diga por que este grupo é capaz de executar o plano.")
    imp = [proj.get(k) for k in ("imp_problema", "imp_resultados", "imp_impacto")]
    c["impacto"] = ("Impacto", [x for x in imp if str(x or "").strip()] + ([f"ODS: {proj['ods']}"] if proj.get("ods") else [])
                    or ["Preencha a ferramenta Impacto social e ODS."], "Editais valorizam o impacto: diga quem é beneficiado e como medir.")
    c["fechamento"] = ("Obrigado", [proj.get("contatos") or "Contato do grupo"],
                       "Termine repetindo em uma frase o que o negócio faz e o principal número do plano.")
    return c


def gerar_pitch_pptx(proj: dict, d: dict, res: dict, roteiro: str = "Banca acadêmica", perguntas: list | None = None) -> bytes:
    """Apresentação 16:9 editável no PowerPoint, LibreOffice ou Google Apresentações, com notas do apresentador."""
    import io
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.util import Inches, Pt

    AZUL, CINZA = RGBColor(0x1F, 0x4F, 0x8F), RGBColor(0x44, 0x4B, 0x55)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    branco = prs.slide_layouts[6]
    cont = conteudo_pitch(proj, d, res)

    def caixa(sl, x, y, w, h, texto, tam=20, cor=CINZA, negrito=False):
        tf = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = texto
        p.font.size, p.font.bold, p.font.color.rgb = Pt(tam), negrito, cor
        return tf

    def faixa(sl):
        sh = sl.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.25))
        sh.fill.solid()
        sh.fill.fore_color.rgb = AZUL
        sh.line.fill.background()

    def grafico_anos(sl):
        cd = CategoryChartData()
        cd.categories = [f"Ano {k}" for k in res["ANOS"]]
        cd.add_series("Receita", [round(res["anos"][k]["rec_bruta"], 2) for k in res["ANOS"]])
        cd.add_series("Resultado líquido", [round(res["anos"][k]["resultado"], 2) for k in res["ANOS"]])
        ch = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.0), Inches(1.6), Inches(5.8), Inches(5.2), cd).chart
        ch.has_legend, ch.legend.position, ch.legend.include_in_layout = True, XL_LEGEND_POSITION.BOTTOM, False
        ch.value_axis.tick_labels.font.size = Pt(11)
        ch.category_axis.tick_labels.font.size = Pt(12)

    def grafico_investimento(sl):
        cd = CategoryChartData()
        cd.categories = ["Fixo", "Pré-operacional", "Estoque", "Capital de giro"]
        cd.add_series("Investimento", [round(res["inv_fixo"], 2), round(res["inv_pre"], 2), round(res["inv_estoque"], 2),
                                       round(res["capital_giro"], 2)])
        ch = sl.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(7.3), Inches(1.6), Inches(5.5), Inches(5.2), cd).chart
        ch.has_legend, ch.legend.position = True, XL_LEGEND_POSITION.RIGHT
        ch.plots[0].has_data_labels = True
        ch.plots[0].data_labels.number_format, ch.plots[0].data_labels.number_format_is_linked = "0%", False
        ch.plots[0].data_labels.show_percentage, ch.plots[0].data_labels.show_value = True, False

    for i, chave_ in enumerate(ROTEIROS_PITCH.get(roteiro, ROTEIROS_PITCH["Banca acadêmica"])):
        titulo, itens, nota = cont[chave_]
        sl = prs.slides.add_slide(branco)
        if chave_ in ("capa", "fechamento"):
            bg = sl.background.fill
            bg.solid()
            bg.fore_color.rgb = AZUL
            caixa(sl, 0.8, 2.4, 11.7, 1.4, titulo, 48, RGBColor(255, 255, 255), True)
            caixa(sl, 0.8, 4.0, 11.7, 2.0, " · ".join(str(x) for x in itens if x), 20, RGBColor(0xDD, 0xE6, 0xF3))
        else:
            faixa(sl)
            caixa(sl, 0.6, 0.45, 12, 0.9, titulo, 34, AZUL, True)
            larg = 6.2 if chave_ in ("numeros", "investimento") else 12.0
            tf = caixa(sl, 0.7, 1.6, larg, 5.4, "", 20)
            for j, it in enumerate(x for x in itens if str(x).strip()):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.text = "• " + str(it)
                p.font.size, p.font.color.rgb = Pt(20 if len(itens) <= 5 else 17), CINZA
                p.space_after = Pt(10)
            if chave_ == "numeros":
                grafico_anos(sl)
            if chave_ == "investimento":
                grafico_investimento(sl)
            caixa(sl, 11.6, 6.9, 1.5, 0.4, str(i + 1), 12, CINZA)
        sl.notes_slide.notes_text_frame.text = f"Tempo sugerido: {TEMPO_SLIDE.get(chave_, 45)} segundos.\n{nota}"
    if perguntas:
        sl = prs.slides.add_slide(branco)
        faixa(sl)
        caixa(sl, 0.6, 0.45, 12, 0.9, "Reserva: perguntas prováveis da banca", 30, AZUL, True)
        tf = caixa(sl, 0.7, 1.5, 12, 5.6, "", 16)
        for j, q in enumerate(perguntas[:7]):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = f"• {q['pergunta']}"
            p.font.size, p.font.color.rgb = Pt(16), CINZA
            p.space_after = Pt(8)
        sl.notes_slide.notes_text_frame.text = "Slide de reserva: use só se a pergunta aparecer. " + " | ".join(
            f"{q['tema']}: {q['motivo']}" for q in perguntas[:7])
    total = sum(TEMPO_SLIDE.get(k, 45) for k in ROTEIROS_PITCH.get(roteiro, []))
    prs.core_properties.title = f"{proj.get('nome_negocio') or 'Plano'} - pitch ({roteiro}, cerca de {total // 60} min)"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------------------
# PONTOS DE PARTIDA POR SETOR (v8.9)
# Estruturas típicas, SEM valores: o grupo preenche preços, custos e quantidades com orçamentos reais.
# ------------------------------------------------------------------------------
SETORES = {
    "Padaria e lanchonete": {
        "tipo": "Produtos", "ref": "Serviços · Serviços de alimentação (restaurantes, lanchonetes)",
        "saz": "Alimentação e cafeteria (mais fraco nas férias)",
        "catalogo": [("Fabricação Própria", "Pão francês (kg)"), ("Fabricação Própria", "Salgado assado"),
                     ("Fabricação Própria", "Bolo (fatia)"), ("Revenda", "Refrigerante lata"), ("Revenda", "Leite (litro)")],
        "inv": ["Forno turbo", "Masseira", "Cilindro", "Vitrine refrigerada", "Balcão e prateleiras", "Freezer", "Computador e sistema de caixa"],
        "fixos": [("Aluguel", "Despesa administrativa"), ("Energia elétrica", "Custo indireto de produção"), ("Gás", "Custo indireto de produção"),
                  ("Água", "Custo indireto de produção"), ("Contador", "Despesa administrativa"), ("Internet e telefone", "Despesa administrativa")],
        "formal": ["Alvará de funcionamento", "Licença da vigilância sanitária", "Auto de vistoria do Corpo de Bombeiros",
                   "Curso de boas práticas de manipulação de alimentos"],
        "riscos": [("Alta do preço da farinha e do trigo", 4, 3, "Contrato com dois fornecedores e revisão mensal de preços"),
                   ("Perdas de produção no fim do dia", 4, 2, "Produção em lotes menores e promoção no fim da tarde")],
        "erros": ["Esquecer as perdas e sobras do dia no custo do produto.", "Calcular o preço do pão sem a energia e o gás do forno.",
                  "Prever o mesmo movimento todos os dias da semana."],
    },
    "Salão de beleza": {
        "tipo": "Serviços", "ref": "Serviços · Serviços pessoais (beleza, lavanderia etc.)", "saz": "Alimentação e cafeteria (mais fraco nas férias)",
        "catalogo": [("Serviço", "Corte"), ("Serviço", "Escova"), ("Serviço", "Coloração"), ("Serviço", "Manicure e pedicure"),
                     ("Revenda", "Produto de cuidado capilar")],
        "inv": ["Cadeiras de corte", "Lavatório", "Secadores e chapinhas", "Espelhos e bancadas", "Esterilizador", "Recepção e sofá"],
        "fixos": [("Aluguel", "Despesa administrativa"), ("Energia elétrica", "Custo indireto de produção"), ("Água", "Custo indireto de produção"),
                  ("Sistema de agendamento", "Despesa comercial"), ("Contador", "Despesa administrativa")],
        "formal": ["Alvará de funcionamento", "Licença da vigilância sanitária", "Contrato de parceria com profissionais (Lei do Salão-Parceiro)"],
        "riscos": [("Saída de profissional com a carteira de clientes", 3, 4, "Contrato de parceria e relacionamento direto com o cliente"),
                   ("Agenda vazia em dias de semana", 4, 3, "Pacotes e preço diferenciado de segunda a quarta")],
        "erros": ["Esquecer a comissão ou parceria do profissional no custo do serviço.", "Ignorar o tempo de cada serviço na capacidade do salão.",
                  "Não contar produtos usados em cada atendimento (custo por aplicação)."],
    },
    "Oficina mecânica": {
        "tipo": "Serviços", "ref": "Serviços · Manutenção e reparação de veículos", "saz": "Sem sazonalidade",
        "catalogo": [("Serviço", "Troca de óleo"), ("Serviço", "Revisão de freios"), ("Serviço", "Alinhamento e balanceamento"),
                     ("Revenda", "Peças de reposição")],
        "inv": ["Elevador automotivo", "Compressor", "Ferramentas manuais", "Scanner de diagnóstico", "Bancada e armários"],
        "fixos": [("Aluguel do galpão", "Despesa administrativa"), ("Energia elétrica", "Custo indireto de produção"),
                  ("Descarte de óleo e resíduos", "Custo indireto de produção"), ("Contador", "Despesa administrativa")],
        "formal": ["Alvará de funcionamento", "Licença ambiental municipal", "Auto de vistoria do Corpo de Bombeiros"],
        "riscos": [("Garantia de serviço mal executado", 2, 4, "Checklist de qualidade e teste final"),
                   ("Inadimplência de clientes", 3, 3, "Pagamento na retirada do veículo")],
        "erros": ["Somar a margem sobre peças e mão de obra como se fossem a mesma coisa.", "Esquecer o descarte correto de óleo e peças."],
    },
    "Loja de roupas": {
        "tipo": "Produtos", "ref": "Comércio · Vestuário e complementos (varejo)", "saz": "Varejo de presentes e vestuário (pico no Natal)",
        "catalogo": [("Revenda", "Blusa"), ("Revenda", "Calça"), ("Revenda", "Vestido"), ("Revenda", "Acessórios")],
        "inv": ["Araras e manequins", "Provadores", "Balcão e caixa", "Iluminação da vitrine", "Computador e sistema de caixa"],
        "fixos": [("Aluguel", "Despesa administrativa"), ("Energia elétrica", "Despesa administrativa"), ("Marketing digital", "Despesa comercial"),
                  ("Contador", "Despesa administrativa")],
        "formal": ["Alvará de funcionamento", "Auto de vistoria do Corpo de Bombeiros"],
        "riscos": [("Estoque parado de coleções passadas", 4, 3, "Compras menores e liquidação ao fim de cada estação"),
                   ("Vendas parceladas sem juros apertando o caixa", 4, 3, "Limite de parcelas e capital de giro para o prazo")],
        "erros": ["Comprar estoque demais na abertura.", "Ignorar o prazo do cartão parcelado no capital de giro.",
                  "Prever vendas iguais em todos os meses."],
    },
    "Produção rural e agricultura familiar": {
        "tipo": "Produtos", "ref": "Comércio · Produtos alimentícios, bebidas e fumo (varejo)", "saz": "Sem sazonalidade",
        "catalogo": [("Fabricação Própria", "Hortaliças (maço)"), ("Fabricação Própria", "Queijo artesanal (kg)"),
                     ("Fabricação Própria", "Doces e conservas (pote)")],
        "inv": ["Sistema de irrigação", "Estufa", "Câmara fria ou freezer", "Utensílios de produção", "Embalagens e rotulagem"],
        "fixos": [("Energia elétrica", "Custo indireto de produção"), ("Transporte até a feira", "Despesa comercial"),
                  ("Taxa da cooperativa ou associação", "Despesa administrativa")],
        "formal": ["Inscrição de produtor rural", "Cadastro da Agricultura Familiar (CAF)", "Registro sanitário dos produtos de origem animal",
                   "Habilitação para vender ao PNAE e ao PAA"],
        "riscos": [("Perda da safra por clima", 3, 5, "Diversificação de culturas e irrigação"),
                   ("Dependência de um único comprador", 3, 4, "Vender em feiras, cooperativa e compras públicas")],
        "erros": ["Não contar o trabalho da família como custo.", "Esquecer as perdas entre a colheita e a venda.",
                  "Ignorar as compras públicas (PNAE e PAA) como canal de venda."],
    },
    "Aplicativo ou serviço digital": {
        "tipo": "Assinatura", "ref": "Serviços · Tecnologia da informação", "saz": "Sem sazonalidade",
        "catalogo": [("Serviço", "Implantação para cliente empresarial")],
        "inv": ["Computadores da equipe", "Desenvolvimento da primeira versão", "Identidade visual"],
        "fixos": [("Servidores e nuvem", "Custo indireto de produção"), ("Ferramentas de software", "Despesa administrativa"),
                  ("Marketing digital", "Despesa comercial"), ("Contador", "Despesa administrativa")],
        "formal": ["Termos de uso e política de privacidade (LGPD)", "Registro da marca no INPI"],
        "riscos": [("Cancelamento alto nos primeiros meses", 4, 4, "Acompanhar uso e conversar com quem cancela"),
                   ("Custo de aquisição maior que o valor do cliente", 3, 4, "Testar canais pequenos antes de escalar")],
        "erros": ["Supor crescimento sem custo de aquisição.", "Esquecer o cancelamento mensal (churn) na receita.",
                  "Ignorar o custo de nuvem que cresce com os usuários."],
    },
}


def criar_projeto_setor(caminho, nome: str, setor: str) -> None:
    """Projeto em branco com a estrutura típica do setor (sem valores)."""
    from pnbox_db import conectar, criar_projeto_em_branco, salvar_unica
    criar_projeto_em_branco(caminho, nome)
    s = SETORES[setor]
    conn = conectar(caminho)
    salvar_unica(conn, "projeto", {"tipo_negocio": s["tipo"], "setor_referencia": s["ref"], "setor_atuacao": setor,
                                   "saz_perfil": s["saz"] if s["saz"] != "Sem sazonalidade" else ""})
    conn.executemany("INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial, horas_unit) "
                     "VALUES (?,?,0,0,0,0,0)", s["catalogo"])
    conn.executemany("INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct) VALUES (?,1,0,10,0)",
                     [(x,) for x in s["inv"]])
    conn.executemany("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,0)", s["fixos"])
    conn.executemany("INSERT INTO formalizacao (item, status, observacao) VALUES (?, 'Pendente', '')", [(x,) for x in s["formal"]])
    conn.executemany("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,'')", s["riscos"])
    if s["saz"] != "Sem sazonalidade":
        conn.executemany("INSERT INTO sazonalidade (mes, indice) VALUES (?,?)",
                         [(m + 1, float(v)) for m, v in enumerate(PERFIS_SAZONAIS[s["saz"]])])
    conn.commit()
    conn.close()
