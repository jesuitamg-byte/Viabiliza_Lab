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
