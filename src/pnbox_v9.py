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
