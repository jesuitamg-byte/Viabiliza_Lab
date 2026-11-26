"""
Funções das versões 8.6 a 9.0 (sem interface): sazonalidade, sensibilidade (tornado),
taxas do Banco Central, pitch em PowerPoint, modelos por setor, mapa de concorrentes,
árvore de decisão e acessibilidade.
"""
import pandas as pd

from pnbox_mix import _n, _p
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
