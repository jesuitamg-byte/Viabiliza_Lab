"""
================================================================================
VIABILIZALAB 8 - CÁLCULOS DAS NOVAS FERRAMENTAS (pnbox_v8.py)
================================================================================
Glossário, trilha guiada, financiamento, dados públicos (IBGE), referências do setor,
SWOT cruzada, Lean Canvas, valuation e rodada, impacto social/ODS, editais de fomento,
importação de planilhas e apoio ao modo professor.
"""
import json
import re
import unicodedata
from datetime import date

import numpy as np
import pandas as pd

from pnbox_motor import brl, num_br, pct, div, parcelas_emprestimo, calcular_tir

# ------------------------------------------------------------------------------
# 1. GLOSSÁRIO E LEITURA DOS INDICADORES
# ------------------------------------------------------------------------------
GLOSSARIO = {
    "Capital de giro": "Dinheiro necessário para manter a operação entre pagar fornecedores e salários e receber dos clientes.",
    "CAC": "Custo de aquisição de cliente: quanto se gasta em marketing e vendas para conquistar um cliente novo.",
    "Churn": "Percentual de clientes que cancelam a assinatura em um período, normalmente por mês.",
    "CMV / CPV / CSV": "Custo das mercadorias vendidas, dos produtos vendidos ou dos serviços vendidos: o custo direto do que foi vendido.",
    "Custo fixo": "Gasto que não muda com o volume de vendas no curto prazo, como aluguel, salários administrativos e contador.",
    "Custo variável": "Gasto que acompanha as vendas, como matéria-prima, comissões, impostos sobre a venda e taxas de cartão.",
    "Depreciação": "Perda de valor dos bens (equipamentos, móveis, veículos) pelo uso e pelo tempo, lançada como despesa sem saída de caixa.",
    "DRE": "Demonstração do Resultado do Exercício: receitas menos custos e despesas, até o lucro ou prejuízo do período.",
    "EBITDA": "Resultado antes de juros, impostos sobre o lucro, depreciação e amortização: mostra a geração de caixa da operação.",
    "Fator R": "No Simples Nacional, a razão entre a folha de pagamento e a receita dos últimos 12 meses; acima de 28% alguns serviços pagam menos imposto.",
    "Fluxo de caixa": "Registro das entradas e saídas de dinheiro mês a mês; mostra se vai faltar caixa mesmo com lucro.",
    "Índice de lucratividade (IL)": "Valor presente das entradas dividido pelo investimento; acima de 1 o projeto cria valor.",
    "Investimento inicial": "Tudo o que se gasta antes de abrir: equipamentos, reformas, legalização, estoque inicial e capital de giro.",
    "LTV": "Valor do tempo de vida do cliente: quanto um cliente deixa de margem enquanto continua comprando ou assinando.",
    "Margem de contribuição": "Receita menos custos variáveis: o que sobra de cada venda para pagar os custos fixos e gerar lucro.",
    "Markup": "Multiplicador aplicado ao custo para chegar ao preço de venda, cobrindo despesas, impostos e lucro desejado.",
    "MRR": "Receita recorrente mensal das assinaturas ativas.",
    "Payback": "Tempo para recuperar o investimento com o caixa gerado. O descontado considera o valor do dinheiro no tempo.",
    "Ponto de equilíbrio": "Receita em que o lucro é zero: abaixo dela há prejuízo. O financeiro desconsidera a depreciação; o econômico inclui o custo de oportunidade.",
    "Pró-labore": "Remuneração dos sócios pelo trabalho na empresa, diferente da distribuição de lucros.",
    "Ponto de equilíbrio em unidades": "Quantidade que precisa ser vendida para cobrir todos os custos.",
    "Rentabilidade": "Lucro em relação ao investimento feito (retorno sobre o capital investido).",
    "Lucratividade": "Lucro em relação à receita (quanto sobra de cada real vendido).",
    "Simples Nacional": "Regime tributário de micro e pequenas empresas, com um documento único de arrecadação (DAS) e alíquota por faixa de receita.",
    "SOM / SAM / TAM": "Mercado total (TAM), mercado que a empresa consegue atender (SAM) e a fatia que pretende conquistar (SOM).",
    "TIR": "Taxa interna de retorno: a rentabilidade anual que o projeto oferece. Compare com a TMA.",
    "TMA": "Taxa mínima de atratividade: o retorno mínimo exigido pelos sócios, em geral o que renderia uma aplicação segura mais um prêmio pelo risco.",
    "VPL": "Valor presente líquido: soma dos fluxos de caixa futuros trazidos a valor de hoje pela TMA, menos o investimento. Positivo = o projeto cria valor.",
    "Valuation": "Estimativa de quanto vale o negócio, por fluxo de caixa descontado ou por múltiplos de mercado.",
    "Pre-money / post-money": "Valor da empresa antes e depois de receber o dinheiro de um investidor.",
    "SAC": "Sistema de amortização constante: o principal é pago em partes iguais e as parcelas diminuem com o tempo.",
    "Tabela Price": "Sistema de parcelas iguais: no início paga-se mais juros e no fim mais principal.",
    "Carência": "Período inicial do empréstimo em que se paga só os juros, sem amortizar o principal.",
    "CET": "Custo efetivo total do empréstimo: a taxa que iguala o valor recebido a todas as parcelas pagas.",
    "Teoria da mudança": "Cadeia que liga as atividades de um negócio de impacto aos resultados e às transformações sociais pretendidas.",
}


def explicar_indicadores(res: dict) -> list:
    """Leitura em linguagem simples dos principais números do projeto aberto."""
    a1 = res["anos"][1]
    N = res["N"]
    L = []
    v = res["vpl"]
    L.append(("VPL", brl(v), f"Trazendo os {N} anos de caixa para valores de hoje pela TMA de {pct(res['tma'])}, o projeto "
              + (f"devolve o investimento e ainda cria {brl(v)} de valor." if v > 0 else f"não recupera o investimento: faltam {brl(-v)}.")))
    tir = res["tir"]
    if tir is not None and tir == tir:
        L.append(("TIR", pct(tir), "O projeto rende " + pct(tir) + " ao ano; " +
                  ("mais" if tir > res["tma"] else "menos") + f" que a TMA de {pct(res['tma'])}."))
    L.append(("Payback descontado", str(res["payback_desc"]), "É o tempo para o caixa gerado, já descontado pela TMA, pagar o investimento inicial."))
    if a1["rec_bruta"]:
        L.append(("Margem de contribuição (Ano 1)", pct(a1["imc"]),
                  f"De cada R$ 100 vendidos, R$ {num_br(a1['imc'] * 100, 2)} sobram para pagar os custos fixos e gerar lucro."))
        L.append(("Lucratividade (Ano 1)", pct(div(a1["resultado"], a1["rec_bruta"])),
                  f"De cada R$ 100 vendidos, o lucro do Ano 1 é de R$ {num_br(div(a1['resultado'], a1['rec_bruta']) * 100, 2)}."))
        L.append(("Ponto de equilíbrio contábil", brl(a1["pe_contabil"]),
                  f"É preciso vender {brl(a1['pe_contabil'] / 12)} por mês, em média, para não ter prejuízo. A receita prevista é "
                  f"{pct(div(a1['rec_bruta'], a1['pe_contabil']) - 1, 0)} {'acima' if a1['rec_bruta'] >= a1['pe_contabil'] else 'abaixo'} disso."))
    minimo = float(res["fluxo_mensal"]["Saldo final"].min())
    L.append(("Menor saldo de caixa (Ano 1)", brl(minimo),
              "O caixa nunca fica negativo no primeiro ano." if minimo >= 0 else "Em algum mês falta dinheiro: aumente o capital de giro ou a reserva."))
    return L


# ------------------------------------------------------------------------------
# 2. FINANCIAMENTO: SIMULAÇÃO E COMPARAÇÃO DE PROPOSTAS
# ------------------------------------------------------------------------------
def resumo_emprestimo(valor, taxa_aa, prazo, carencia, sistema) -> dict:
    tab = parcelas_emprestimo(float(valor), float(taxa_aa), int(prazo), int(carencia or 0), sistema)
    p = tab["Parcela"]
    pos = p[p > 0]
    fluxo = [float(valor)] + [-x for x in p]
    tir_m = calcular_tir(fluxo)
    cet = (1 + tir_m) ** 12 - 1 if tir_m is not None else float("nan")
    return {"tabela": tab, "total_pago": float(p.sum()), "juros": float(tab["Juros"].sum()),
            "primeira": float(pos.iloc[0]) if len(pos) else 0.0, "maior": float(p.max()), "menor": float(pos.min()) if len(pos) else 0.0,
            "cet": cet, "carencia_parcela": float(p.iloc[0]) if int(carencia or 0) > 0 else float("nan")}


# ------------------------------------------------------------------------------
# 3. DADOS PÚBLICOS: IBGE (CENSO 2022, PIB, CEMPRE) E POF 2017-2018
# ------------------------------------------------------------------------------
UFS = {"AC": (12, "Acre"), "AL": (27, "Alagoas"), "AP": (16, "Amapá"), "AM": (13, "Amazonas"), "BA": (29, "Bahia"),
       "CE": (23, "Ceará"), "DF": (53, "Distrito Federal"), "ES": (32, "Espírito Santo"), "GO": (52, "Goiás"),
       "MA": (21, "Maranhão"), "MT": (51, "Mato Grosso"), "MS": (50, "Mato Grosso do Sul"), "MG": (31, "Minas Gerais"),
       "PA": (15, "Pará"), "PB": (25, "Paraíba"), "PR": (41, "Paraná"), "PE": (26, "Pernambuco"), "PI": (22, "Piauí"),
       "RJ": (33, "Rio de Janeiro"), "RN": (24, "Rio Grande do Norte"), "RS": (43, "Rio Grande do Sul"), "RO": (11, "Rondônia"),
       "RR": (14, "Roraima"), "SC": (42, "Santa Catarina"), "SP": (35, "São Paulo"), "SE": (28, "Sergipe"), "TO": (17, "Tocantins")}
API_IBGE = "https://servicodados.ibge.gov.br/api"


def baixar_json(url: str):
    """Baixa JSON da internet na versão on-line (Pyodide) ou instalada (Python comum)."""
    try:
        from pyodide.http import open_url  # versão on-line: requisição feita pelo próprio navegador
        return json.loads(open_url(url).read())
    except ImportError:
        import urllib.request
        with urllib.request.urlopen(url, timeout=20) as r:
            return json.loads(r.read().decode("utf-8"))


def municipios_da_uf(uf: str) -> list:
    js = baixar_json(f"{API_IBGE}/v1/localidades/estados/{uf}/municipios")
    return sorted(((str(m["id"]), m["nome"]) for m in js), key=lambda x: x[1])


def _serie_ibge(agregado: int, variavel: int, cod: str, periodo: str = "-1") -> tuple:
    js = baixar_json(f"{API_IBGE}/v3/agregados/{agregado}/periodos/{periodo}/variaveis/{variavel}?localidades=N6[{cod}]")
    serie = js[0]["resultados"][0]["series"][0]["serie"]
    per = sorted(serie)[-1]
    v = serie[per]
    return (float(v) if v not in ("-", "...", "X", None) else float("nan")), per


INDICADORES_MUNICIPIO = [
    ("populacao", "População residente (Censo 2022)", 4709, 93, "2022", "pessoas",
     "IBGE. Censo Demográfico 2022, tabela 4709 (Sidra)."),
    ("domicilios", "Domicílios particulares permanentes ocupados (Censo 2022)", 4712, 381, "2022", "domicílios",
     "IBGE. Censo Demográfico 2022, tabela 4712 (Sidra)."),
    ("pib", "PIB a preços correntes (mil R$)", 5938, 37, "-1", "mil R$", "IBGE. Produto Interno Bruto dos Municípios, tabela 5938 (Sidra)."),
    ("unidades_locais", "Unidades locais de empresas e organizações", 1685, 706, "-1", "unidades",
     "IBGE. Cadastro Central de Empresas (Cempre), tabela 1685 (Sidra)."),
]


def dados_municipio(cod: str) -> list:
    """Busca os indicadores do município na API do IBGE. Cada item: (chave, rótulo, valor, período, fonte)."""
    saida = []
    for chave, rot, ag, var, per, _, fonte in INDICADORES_MUNICIPIO:
        try:
            v, p = _serie_ibge(ag, var, cod, per)
            saida.append((chave, rot, v, p, fonte))
        except Exception:
            continue
    m = {c: (v, p) for c, _, v, p, _ in saida}
    if "pib" in m and "populacao" in m and m["populacao"][0]:
        saida.append(("pib_per_capita", "PIB per capita (R$)", m["pib"][0] * 1000 / m["populacao"][0], m["pib"][1],
                      "Calculado: PIB ÷ população do Censo 2022 (IBGE)."))
    if "populacao" in m and "domicilios" in m and m["domicilios"][0]:
        saida.append(("moradores_dom", "Moradores por domicílio", m["populacao"][0] / m["domicilios"][0], "2022",
                      "Calculado com o Censo 2022 (IBGE)."))
    return saida


# Fallback da POF 2017-2018 (despesa monetária e não monetária média mensal familiar, R$; tabela 6715 do Sidra)
POF_FALLBACK = {
    "Brasil": {"1 Despesa total": 4649.03, "2.1 Despesas de consumo": 3764.51, "2.1.1 Alimentação": 658.23,
               "2.1.2 Habitação": 1377.14, "2.1.3 Vestuário": 160.25, "2.1.4 Transporte": 679.76,
               "2.1.5 Higiene e cuidados pessoais": 136.82, "2.1.6 Assistência à saude": 302.06, "2.1.7 Educação": 175.60,
               "2.1.8 Recreação e cultura": 96.16, "2.1.9 Fumo": 17.40, "2.1.10 Serviços pessoais": 48.55,
               "2.1.10.1 Cabeleireiro": 29.16, "2.1.10.2 Manicuro e pedicuro": 9.13, "2.1.11 Despesas diversas": 112.53},
    "Minas Gerais": {"1 Despesa total": 4390.72, "2.1 Despesas de consumo": 3635.92, "2.1.1 Alimentação": 688.51,
                     "2.1.2 Habitação": 1230.81, "2.1.3 Vestuário": 165.41, "2.1.4 Transporte": 662.29,
                     "2.1.5 Higiene e cuidados pessoais": 126.94, "2.1.6 Assistência à saude": 324.49, "2.1.7 Educação": 163.56,
                     "2.1.8 Recreação e cultura": 95.09, "2.1.9 Fumo": 15.11, "2.1.10 Serviços pessoais": 50.81,
                     "2.1.10.1 Cabeleireiro": 29.63, "2.1.10.2 Manicuro e pedicuro": 10.11, "2.1.11 Despesas diversas": 112.90},
}


def pof_uf(uf: str) -> tuple:
    """Despesa média mensal familiar por tipo de gasto (POF 2017-2018) na UF. Retorna (DataFrame, origem)."""
    cod, nome = UFS.get(uf, (None, "Brasil"))
    try:
        loc = f"N3[{cod}]" if cod else "N1[all]"
        js = baixar_json(f"{API_IBGE}/v3/agregados/6715/periodos/2018/variaveis/1201?localidades={loc}"
                         "&classificacao=339[7999]|12190[all]")
        linhas = []
        for r in js[0]["resultados"]:
            cat = list(next(c for c in r["classificacoes"] if c["id"] == "12190")["categoria"].values())[0]
            v = r["series"][0]["serie"]["2018"]
            linhas.append((cat, float(v) if v not in ("-", "...", "X") else 0.0))
        return pd.DataFrame(linhas, columns=["Tipo de despesa", "R$ por mês por família"]), f"API do IBGE ({nome})"
    except Exception:
        base = POF_FALLBACK.get(nome, POF_FALLBACK["Brasil"])
        rot = nome if nome in POF_FALLBACK else "Brasil"
        return (pd.DataFrame(list(base.items()), columns=["Tipo de despesa", "R$ por mês por família"]),
                f"tabela embutida no programa ({rot}; sem conexão com o IBGE)")


# ------------------------------------------------------------------------------
# 4. REFERÊNCIAS DO SETOR (IBGE PAC 2024 E PAS 2024) E SOBREVIVÊNCIA (SEBRAE)
# ------------------------------------------------------------------------------
# PAC 2024 (tabela 10645): receita operacional líquida e salários em mil R$, pessoal ocupado, empresas, taxa de margem (%)
PAC_2024 = [
    ("Comércio varejista (total)", 3196219376, 229175228, 7647251, 1175862, 38.6),
    ("Hipermercados e supermercados", 924967521, 50717945, 1609883, 19434, 26.5),
    ("Minimercados, mercearias e armazéns", 148833942, 11026259, 526752, 113355, 32.8),
    ("Produtos alimentícios, bebidas e fumo (varejo)", 156338185, 17654808, 736302, 167786, 52.3),
    ("Material de construção (varejo)", 288251891, 26750623, 858513, 166575, 54.9),
    ("Informática e comunicação (varejo)", 42907443, 5093579, 162032, 45099, 53.2),
    ("Eletrodomésticos, áudio e vídeo (varejo)", 84546506, 5775333, 139352, 10728, 51.3),
    ("Móveis e artigos de uso doméstico (varejo)", 101785615, 10667868, 362219, 72325, 68.5),
    ("Tecidos e armarinho (varejo)", 24404688, 2802899, 127716, 30696, 77.9),
    ("Artigos culturais, recreativos e esportivos (varejo)", 62914009, 5612408, 207344, 49431, 72.1),
    ("Farmácia, perfumaria e cosméticos (varejo)", 314671039, 31665392, 887305, 133786, 66.2),
    ("Vestuário e complementos (varejo)", 127724083, 18816230, 653824, 123163, 86.1),
    ("Calçados, artigos de couro e viagem (varejo)", 41230762, 5393552, 190949, 27246, 70.1),
    ("Combustíveis e lubrificantes (varejo)", 569223037, 12982495, 424940, 38620, 15.8),
    ("Peças e acessórios para veículos", 232531527, 19400443, 568790, 103533, 42.8),
    ("Comércio por atacado (total)", 3753962262, 102806917, 2014081, 270636, 22.0),
]
# PAS 2024 (tabela 10758): empresas, pessoal ocupado, receita operacional líquida (mil R$), salários e retiradas (mil R$)
PAS_2024 = [
    ("Serviços de alimentação (restaurantes, lanchonetes)", 220098, 1856632, 257709554, 44068802),
    ("Serviços de alojamento", 31828, 386535, 59173475, 10594948),
    ("Atividades culturais, recreativas e esportivas", 57579, 253500, 33676457, 5054736),
    ("Serviços pessoais (beleza, lavanderia etc.)", 56993, 254839, 28389907, 5521780),
    ("Ensino continuado (cursos livres)", 89232, 348550, 32517587, 7361978),
    ("Tecnologia da informação", 120332, 888225, 397283053, 74780082),
    ("Serviços audiovisuais", 29931, 129679, 52818481, 7581899),
    ("Serviços técnico-profissionais (consultoria, contabilidade, engenharia)", 430594, 1768914, 446082202, 63468515),
    ("Agências de viagens e turismo", 16735, 76859, 20135883, 2797727),
    ("Serviços de escritório e apoio administrativo", 189964, 1317782, 135359231, 31424511),
    ("Serviços para edifícios e paisagismo (limpeza, jardinagem)", 27986, 1251112, 87352375, 30309006),
    ("Transporte rodoviário de cargas", 111658, 1286209, 431203673, 49867401),
    ("Transporte rodoviário de passageiros", 28495, 587512, 101062930, 21968911),
    ("Intermediação imobiliária", 68634, 245300, 27940901, 5483068),
    ("Manutenção e reparação de veículos", 83924, 338492, 34852163, 7378712),
    ("Manutenção e reparação de informática e comunicação", 15265, 65003, 10636573, 1909897),
    ("Manutenção e reparação de objetos pessoais e domésticos", 14199, 53944, 5315556, 1142785),
]
FONTE_PAC = "IBGE, Pesquisa Anual de Comércio 2024 (Sidra, tabela 10645)"
FONTE_PAS = "IBGE, Pesquisa Anual de Serviços 2024 (Sidra, tabela 10758)"
# Sebrae (DataSebrae, seminário "Aplicações de análise de sobrevivência", jul. 2026): sobrevivência em 5 anos
SOBREVIVENCIA_5A = {"MEI": 0.415, "Pequenos negócios (MEI + ME + EPP)": 0.525, "Micro e pequenas empresas (ME + EPP)": 0.76}
CAUSAS_FECHAMENTO = [("Falta de capital de giro", 0.22), ("Baixo volume de vendas e falta de clientes", 0.20)]


def atividades_referencia() -> list:
    return [f"Comércio · {a[0]}" for a in PAC_2024] + [f"Serviços · {a[0]}" for a in PAS_2024]


def comparar_setor(atividade: str, proj: dict, d: dict, res: dict) -> pd.DataFrame:
    """Compara os números do plano com a média do setor no Brasil. Retorna linhas (indicador, plano, setor, leitura)."""
    a1 = res["anos"][1]
    rec = a1["rec_bruta"]
    pessoas = float(d["pessoal"]["num_empregados"].sum()) + len(d["socios"])
    linhas = []
    if atividade.startswith("Comércio · "):
        nome = atividade.split(" · ", 1)[1]
        _, rol, sal, po, emp, tm = next(x for x in PAC_2024 if x[0] == nome)
        cmv = a1["custo_dir"]
        if cmv > 0:
            margem_plano = rec / cmv - 1
            linhas.append(("Taxa de margem sobre o custo da mercadoria", margem_plano, tm / 100,
                           "Margem = (receita − custo das mercadorias) ÷ custo das mercadorias."))
        linhas.append(("Salários e retiradas ÷ receita", div(res["salarios_m"] * 12 + _prolabore(d) * 12, rec), sal / rol,
                       "O setor não inclui encargos; o plano usa só salários e pró-labore."))
        linhas.append(("Receita por pessoa ocupada (R$/ano)", div(rec, pessoas), rol * 1000 / po, "Produtividade da equipe."))
        linhas.append(("Receita média por empresa (R$/ano)", rec, rol * 1000 / emp, "Inclui empresas grandes: a média puxa para cima."))
    else:
        nome = atividade.split(" · ", 1)[1]
        _, emp, po, rol, sal = next(x for x in PAS_2024 if x[0] == nome)
        linhas.append(("Salários e retiradas ÷ receita", div(res["salarios_m"] * 12 + _prolabore(d) * 12, rec), sal / rol,
                       "O setor não inclui encargos; o plano usa só salários e pró-labore."))
        linhas.append(("Receita por pessoa ocupada (R$/ano)", div(rec, pessoas), rol * 1000 / po, "Produtividade da equipe."))
        linhas.append(("Receita média por empresa (R$/ano)", rec, rol * 1000 / emp, "Inclui empresas grandes: a média puxa para cima."))
    df = pd.DataFrame(linhas, columns=["Indicador", "Plano", "Setor (Brasil)", "Observação"])
    df["Plano ÷ setor"] = [div(a, b) if b else float("nan") for a, b in zip(df["Plano"], df["Setor (Brasil)"])]
    return df


def _prolabore(d: dict) -> float:
    cf = d["custos_fixos"]
    if cf.empty:
        return 0.0
    return float(cf.loc[cf["descricao"].astype(str).str.lower().str.contains("pró-labore|pro-labore|prolabore", regex=True),
                        "valor_mensal"].sum())


def leitura_setor(df: pd.DataFrame) -> list:
    alertas = []
    for _, r in df.iterrows():
        x = r["Plano ÷ setor"]
        if x != x:
            continue
        if r["Indicador"].startswith("Taxa de margem") and x > 1.5:
            alertas.append("A margem sobre o custo está bem acima da média do setor: confira se o preço é aceito pelos clientes.")
        if r["Indicador"].startswith("Taxa de margem") and x < 0.6:
            alertas.append("A margem sobre o custo está bem abaixo da média do setor: o preço pode estar baixo ou o custo alto.")
        if r["Indicador"].startswith("Receita por pessoa") and x > 2:
            alertas.append("A receita por pessoa está muito acima do setor: a equipe pode estar pequena para o volume previsto.")
        if r["Indicador"].startswith("Salários") and x < 0.5:
            alertas.append("O peso dos salários está muito abaixo do setor: confira se toda a equipe e o pró-labore foram incluídos.")
    return alertas


# ------------------------------------------------------------------------------
# 5. SWOT CRUZADA
# ------------------------------------------------------------------------------
QUADRANTES_CRUZADA = {
    "Ofensiva (forças × oportunidades)": ("Força", "Oportunidade", "Use as forças para aproveitar as oportunidades."),
    "Confronto (forças × ameaças)": ("Força", "Ameaça", "Use as forças para neutralizar as ameaças."),
    "Reforço (fraquezas × oportunidades)": ("Fraqueza", "Oportunidade", "Corrija fraquezas para não perder oportunidades."),
    "Defensiva (fraquezas × ameaças)": ("Fraqueza", "Ameaça", "Reduza fraquezas que deixam a empresa exposta às ameaças."),
}


def fatores_swot(d: dict, natureza: str) -> list:
    s = d["swot"]
    s = s[s["natureza"] == natureza]
    pri = s[s["priorizado"].fillna(0).astype(int) == 1]
    return list((pri if not pri.empty else s)["fator"].astype(str))


def sugestoes_cruzada(d: dict) -> dict:
    out = {}
    for q, (a, b, dica) in QUADRANTES_CRUZADA.items():
        fa, fb = fatores_swot(d, a), fatores_swot(d, b)
        out[q] = {"dica": dica, "a": fa, "b": fb,
                  "pares": [f"{x} + {y}" for x in fa[:3] for y in fb[:2]]}
    return out


# ------------------------------------------------------------------------------
# 6. LEAN CANVAS
# ------------------------------------------------------------------------------
def linhas(t) -> list:
    return [x.strip(" -•\t") for x in re.split(r"[\n;]", str(t or "")) if x.strip(" -•\t")]


def lean_canvas(proj: dict, d: dict, res: dict) -> dict:
    """Lean Canvas (Maurya, 2012) montado com o que o grupo já preencheu nas outras ferramentas."""
    dores = linhas(proj.get("vpc_dores"))
    if not dores and not d["entrevistas"].empty:
        from collections import Counter
        cont = Counter(x.strip().lower() for t in d["entrevistas"]["dores"] for x in str(t).split(";") if x.strip())
        dores = [k for k, _ in cont.most_common(3)]
    segs = list(d["segmentos"]["nome"].astype(str))
    pers = list(d["personas"]["nome"].astype(str)) if "nome" in d["personas"] else []
    prop = ""
    if str(proj.get("posic_beneficio") or "").strip():
        prop = f"{proj.get('posic_marca') or proj.get('nome_negocio')}: {proj.get('posic_beneficio')}"
    elif not d["propostas"].empty:
        r = d["propostas"].iloc[0]
        prop = f"Ajudamos a {r['que_desejam']} por meio de {r['por_meio']}"
    solucao = linhas(proj.get("vpc_aliviadores")) or list(res["cat"]["descricao"].astype(str))[:4]
    canais = list(d["canais"]["canal"].astype(str))[:5]
    cat = res["cat"]
    receitas = [f"{r['descricao']}: {brl(r['preco_venda'])}" for _, r in cat.head(5).iterrows()] if not cat.empty else []
    if res.get("recorrencia", {}).get("tem"):
        receitas.append(f"Assinaturas (MRR do mês 12: {brl(res['recorrencia'].get('mrr12', 0))})")
    a1 = res["anos"][1]
    custos = [f"Pessoal: {brl(a1['pessoal'])}/ano", f"Custos fixos: {brl(a1['fixos'])}/ano",
              f"Custos diretos: {brl(a1['custo_dir'])}/ano", f"Investimento inicial: {brl(res['investimento_total'])}"]
    metricas = linhas(proj.get("lean_metricas"))
    if not metricas and not d["aarrr"].empty:
        metricas = [f"{r['etapa']}: {r['metrica']} (meta {r['meta']})" for _, r in d["aarrr"].iterrows() if str(r["metrica"]).strip()]
    if not metricas:
        metricas = list(d["objetivos"]["indicador"].astype(str))[:4]
    alternativas = linhas(proj.get("lean_alternativas")) or list(d["concorrentes"]["nome"].astype(str))[:4]
    return {"Problema": dores[:3], "Alternativas existentes": alternativas, "Segmentos de clientes": segs[:3],
            "Primeiros clientes (early adopters)": pers[:2], "Proposta única de valor": [prop] if prop else [],
            "Solução": solucao[:4], "Canais": canais, "Fontes de receita": receitas, "Estrutura de custos": custos,
            "Métricas-chave": metricas[:5], "Vantagem injusta": linhas(proj.get("lean_vantagem"))}


# ------------------------------------------------------------------------------
# 7. VALUATION E RODADA DE INVESTIMENTO
# ------------------------------------------------------------------------------
def valuation(res: dict, proj: dict, d: dict) -> dict:
    tma = res["tma"]
    g = float(proj.get("val_crescimento_perp_pct") if proj.get("val_crescimento_perp_pct") is not None else 3) / 100
    N = res["N"]
    ebitdas = [res["anos"][k]["ebitda"] for k in res["ANOS"]]
    vp = sum(e / (1 + tma) ** k for k, e in zip(res["ANOS"], ebitdas))
    terminal = ebitdas[-1] * (1 + g) / (tma - g) if tma > g and ebitdas[-1] > 0 else 0.0
    vp_term = terminal / (1 + tma) ** N
    divida = float(d["emprestimos"]["valor"].sum())
    ev = vp + vp_term
    m_eb = float(proj.get("val_multiplo_ebitda") or 4)
    m_rec = float(proj.get("val_multiplo_receita") or 1)
    a1 = res["anos"][1]
    eb_ref = res["anos"][min(2, N)]["ebitda"]
    met = pd.DataFrame([
        ("Fluxo de caixa descontado (EBITDA + perpetuidade)", ev - divida),
        (f"Múltiplo de EBITDA ({num_br(m_eb, 1)}x o EBITDA do Ano {min(2, N)})", eb_ref * m_eb - divida),
        (f"Múltiplo de receita ({num_br(m_rec, 1)}x a receita do Ano 1)", a1["rec_bruta"] * m_rec - divida),
    ], columns=["Método", "Valor do patrimônio dos sócios"])
    return {"vp_fluxos": vp, "terminal": terminal, "vp_terminal": vp_term, "ev": ev, "divida": divida,
            "equity_dcf": ev - divida, "metodos": met, "g": g, "faixa": (float(met.iloc[:, 1].min()), float(met.iloc[:, 1].max()))}


def cap_table(d: dict, pre_money: float, investimento: float) -> pd.DataFrame:
    soc = d["socios"]
    total = float(soc["capital_integralizado"].sum())
    linhas_ = []
    post = pre_money + investimento
    for _, r in soc.iterrows():
        antes = div(float(r["capital_integralizado"]), total) if total else div(1, len(soc))
        depois = antes * div(pre_money, post) if post > 0 else antes
        linhas_.append((str(r["nome"]), antes, depois, depois * post))
    if investimento > 0 and post > 0:
        linhas_.append(("Investidor", 0.0, investimento / post, investimento))
    return pd.DataFrame(linhas_, columns=["Sócio", "Participação antes", "Participação depois", "Valor da participação"])


# ------------------------------------------------------------------------------
# 8. IMPACTO SOCIAL E ODS
# ------------------------------------------------------------------------------
ODS = ["1 Erradicação da pobreza", "2 Fome zero e agricultura sustentável", "3 Saúde e bem-estar", "4 Educação de qualidade",
       "5 Igualdade de gênero", "6 Água potável e saneamento", "7 Energia limpa e acessível",
       "8 Trabalho decente e crescimento econômico", "9 Indústria, inovação e infraestrutura", "10 Redução das desigualdades",
       "11 Cidades e comunidades sustentáveis", "12 Consumo e produção responsáveis", "13 Ação contra a mudança global do clima",
       "14 Vida na água", "15 Vida terrestre", "16 Paz, justiça e instituições eficazes", "17 Parcerias e meios de implementação"]
TIPOS_INDICADOR_IMPACTO = ["Produto (output)", "Resultado (outcome)", "Impacto"]
ETAPAS_MUDANCA = [("imp_problema", "Problema social ou ambiental", "Qual problema da comunidade o negócio enfrenta?"),
                  ("imp_publico", "Público beneficiado", "Quem é afetado e quantas pessoas?"),
                  ("imp_atividades", "Atividades", "O que a empresa faz no dia a dia que gera o efeito?"),
                  ("imp_produtos", "Produtos imediatos (outputs)", "O que é entregue e pode ser contado?"),
                  ("imp_resultados", "Resultados (outcomes)", "Que mudança acontece na vida do público?"),
                  ("imp_impacto", "Impacto de longo prazo", "Que transformação se espera na comunidade?")]


# ------------------------------------------------------------------------------
# 9. EDITAIS DE FOMENTO (estrutura usual do Programa Centelha e editais de inovação)
# ------------------------------------------------------------------------------
SECOES_EDITAL = [("Título da proposta", 150), ("Problema ou oportunidade", 2000), ("Solução proposta", 2000),
                 ("Caráter inovador e diferencial", 2000), ("Mercado e clientes", 2000), ("Modelo de negócio", 2000),
                 ("Estágio de desenvolvimento", 1500), ("Equipe", 1500), ("Impactos esperados", 1500),
                 ("Cronograma e próximos passos", 1500)]
RUBRICAS_EDITAL = ["Bolsas", "Material de consumo", "Serviços de terceiros", "Equipamentos e material permanente",
                   "Passagens e diárias", "Outros"]
TRL = {1: "Princípios básicos observados", 2: "Conceito tecnológico formulado", 3: "Prova de conceito",
       4: "Validação em laboratório", 5: "Validação em ambiente relevante", 6: "Protótipo demonstrado em ambiente relevante",
       7: "Protótipo demonstrado em ambiente real", 8: "Sistema completo e qualificado", 9: "Sistema em operação"}


def rascunho_edital(proj: dict, d: dict, res: dict) -> dict:
    nome = proj.get("nome_negocio") or "o negócio"
    seg = ", ".join(d["segmentos"]["nome"].astype(str).head(2)) or "o público-alvo"
    dores = linhas(proj.get("vpc_dores"))
    prod = ", ".join(res["cat"]["descricao"].astype(str).head(4))
    soc = d["socios"]
    m_tam, m_sam = float(proj.get("tam_clientes") or 0), float(proj.get("sam_clientes") or 0)
    a1 = res["anos"][1]
    trl = int(proj.get("edital_trl") or 3)
    errc = d.get("errc", pd.DataFrame())
    criar = list(errc.loc[errc["acao"] == "Criar", "atributo"]) if not errc.empty else []
    cron = d["cronograma"]
    return {
        "Título da proposta": f"{nome}: {proj.get('posic_categoria') or 'novo negócio'} para {seg}",
        "Problema ou oportunidade": (f"{seg} enfrenta: " + "; ".join(dores[:3]) + ".") if dores else
        f"Descreva o problema que {seg} enfrenta, com dados (entrevistas, pesquisa, IBGE).",
        "Solução proposta": f"{nome} oferece {prod}. " + (proj.get("vpc_job") or ""),
        "Caráter inovador e diferencial": (frase_pos(proj) + (" Atributos novos em relação aos concorrentes: " + ", ".join(criar) + "."
                                                              if criar else "")).strip(),
        "Mercado e clientes": f"Mercado total estimado em {num_br(m_tam, 0)} clientes e mercado atendível de {num_br(m_sam, 0)}. "
                              f"Receita prevista no Ano 1: {brl(a1['rec_bruta'])}.",
        "Modelo de negócio": f"Receita com {prod}. Margem de contribuição de {pct(a1['imc'])}; investimento inicial de "
                             f"{brl(res['investimento_total'])}; VPL de {brl(res['vpl'])} em {res['N']} anos.",
        "Estágio de desenvolvimento": f"Nível de maturidade tecnológica (TRL) {trl}: {TRL.get(trl, '')}. "
                                      f"{len(d['entrevistas'])} entrevistas com clientes e {len(d['experimentos'])} experimento(s) planejado(s).",
        "Equipe": "; ".join(f"{r['nome']} ({r['funcao']}): {r['perfil']}" for _, r in soc.iterrows()),
        "Impactos esperados": "; ".join(str(proj.get(k) or "") for k in ("imp_resultados", "imp_impacto") if proj.get(k)) or
        "Descreva os impactos econômicos, sociais e ambientais.",
        "Cronograma e próximos passos": "; ".join(f"mês {num_br(r['mes_inicio'], 0)}: {r['atividade']}" for _, r in cron.iterrows()),
    }


def frase_pos(proj: dict) -> str:
    if not str(proj.get("posic_beneficio") or "").strip():
        return ""
    return (f"Para {proj.get('posic_segmento') or '...'}, {proj.get('posic_marca') or proj.get('nome_negocio')} "
            f"{proj.get('posic_beneficio')}. Diferentemente de {proj.get('posic_concorrente') or '...'}, {proj.get('posic_diferencial') or '...'}.")


# ------------------------------------------------------------------------------
# 10. IMPORTAÇÃO DE PLANILHAS (EXCEL OU CSV)
# ------------------------------------------------------------------------------
def _norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


def ler_planilha(nome: str, dados: bytes) -> pd.DataFrame:
    import io as _io
    if nome.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(_io.BytesIO(dados))
    texto = dados.decode("utf-8-sig", errors="replace")
    sep = ";" if texto.count(";") > texto.count(",") else ","
    df = pd.read_csv(_io.StringIO(texto), sep=sep)
    return df


def mapear_colunas(df: pd.DataFrame, colunas_destino: list, rotulos: dict | None = None) -> dict:
    """Liga as colunas da planilha às do programa pelo nome (ignora acentos, maiúsculas e espaços)."""
    rotulos = rotulos or {}
    alvo = {_norm(c): c for c in colunas_destino}
    alvo.update({_norm(r): c for c, r in rotulos.items()})
    return {c: alvo[_norm(c)] for c in df.columns if _norm(c) in alvo}


def converter_numero(v):
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace("%", "").strip()
    if not s or s.lower() == "nan":
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ------------------------------------------------------------------------------
# 11. MODO PROFESSOR: SEMELHANÇA ENTRE PLANOS E RODADA DE INVESTIMENTO DA TURMA
# ------------------------------------------------------------------------------
def _assinatura_texto(proj: dict, d: dict) -> set:
    partes = [str(proj.get(k) or "") for k in ("nome_negocio", "mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao",
                                                "sumario_executivo", "conclusao", "missao", "visao", "vpc_dores")]
    for t in ("segmentos", "swot", "riscos", "catalogo"):
        df = d.get(t)
        if df is not None and not df.empty:
            partes.extend(df.astype(str).fillna("").agg(" ".join, axis=1).tolist())
    palavras = re.findall(r"[a-zà-ú0-9]{3,}", " ".join(partes).lower())
    return {" ".join(palavras[i:i + 3]) for i in range(max(len(palavras) - 2, 0))}


def _assinatura_numeros(d: dict) -> set:
    nums = set()
    for t, cols in (("catalogo", ("preco_venda", "custo_aquisicao", "qtd_mensal")), ("custos_fixos", ("valor_mensal",)),
                    ("inv_fixos", ("valor_unitario",))):
        df = d.get(t)
        if df is not None and not df.empty:
            for c in cols:
                nums.update(round(float(x), 2) for x in df[c] if float(x or 0) > 0)
    return nums


def semelhanca_planos(projetos: list) -> pd.DataFrame:
    """projetos: lista de (rotulo, proj, d). Retorna pares com semelhança de texto (trigramas) e de números (Jaccard)."""
    ass = [(r, _assinatura_texto(p, d), _assinatura_numeros(d)) for r, p, d in projetos]
    out = []
    for i in range(len(ass)):
        for j in range(i + 1, len(ass)):
            a, b = ass[i], ass[j]
            jt = div(len(a[1] & b[1]), len(a[1] | b[1])) if (a[1] | b[1]) else 0.0
            jn = div(len(a[2] & b[2]), len(a[2] | b[2])) if (a[2] | b[2]) else 0.0
            out.append((a[0], b[0], jt, jn))
    df = pd.DataFrame(out, columns=["Grupo A", "Grupo B", "Textos em comum", "Números em comum"])
    return df.sort_values(["Textos em comum", "Números em comum"], ascending=False).reset_index(drop=True)


def resultado_rodada(aloc: pd.DataFrame) -> pd.DataFrame:
    """aloc: linhas = investidores (grupos), colunas = projetos. Soma o capital recebido, sem contar o próprio grupo."""
    m = aloc.copy()
    for c in m.columns:
        if c in m.index:
            m.loc[c, c] = 0.0
    tot = m.sum(axis=0)
    inv = (m > 0).sum(axis=0)
    df = pd.DataFrame({"Projeto": tot.index, "Capital recebido": tot.values, "Investidores": inv.values})
    return df.sort_values("Capital recebido", ascending=False).reset_index(drop=True)


# ------------------------------------------------------------------------------
# 12. TRILHA GUIADA
# ------------------------------------------------------------------------------
ESSENCIAIS = ["projeto", "identidade", "segmentos", "personas", "mercado", "proposta", "concorrencia", "swot", "marketing",
              "canais", "investimentos", "produtos", "custos", "preco", "dre", "fluxo", "indicadores", "riscos", "sumario"]


def referencias_abnt_dados(proj: dict, d: dict, acesso: str | None = None) -> list:
    """Referências (NBR 6023:2025) das bases de dados usadas pelo programa, para o relatório ABNT."""
    hoje = acesso or date.today().strftime("%d %b. %Y").lower()
    meses = {"jan.": "jan.", "feb.": "fev.", "mar.": "mar.", "apr.": "abr.", "may.": "maio", "jun.": "jun.", "jul.": "jul.",
             "aug.": "ago.", "sep.": "set.", "oct.": "out.", "nov.": "nov.", "dec.": "dez."}
    for a, b in meses.items():
        hoje = hoje.replace(a, b)
    refs = []
    chaves = set(d["dados_publicos"]["chave"]) if "dados_publicos" in d and not d["dados_publicos"].empty else set()
    ibge = "INSTITUTO BRASILEIRO DE GEOGRAFIA E ESTATÍSTICA."
    if chaves & {"populacao", "domicilios", "moradores_dom"}:
        refs.append(f"{ibge} **Censo demográfico 2022**: tabelas 4709 e 4712. Rio de Janeiro: IBGE, 2023. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/4709. Acesso em: {hoje}.")
    if chaves & {"pib", "pib_per_capita"}:
        refs.append(f"{ibge} **Produto interno bruto dos municípios**: tabela 5938. Rio de Janeiro: IBGE, [2025]. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/5938. Acesso em: {hoje}.")
    if "unidades_locais" in chaves:
        refs.append(f"{ibge} **Cadastro central de empresas**: tabela 1685. Rio de Janeiro: IBGE, [2023]. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/1685. Acesso em: {hoje}.")
    if any(str(c).startswith("pof") for c in chaves):
        refs.append(f"{ibge} **Pesquisa de orçamentos familiares 2017-2018**: tabela 6715. Rio de Janeiro: IBGE, 2019. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/6715. Acesso em: {hoje}.")
    setor = str(proj.get("setor_referencia") or "")
    if setor.startswith("Comércio"):
        refs.append(f"{ibge} **Pesquisa anual de comércio 2024**: tabela 10645. Rio de Janeiro: IBGE, [2026]. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/10645. Acesso em: {hoje}.")
    if setor.startswith("Serviços"):
        refs.append(f"{ibge} **Pesquisa anual de serviços 2024**: tabela 10758. Rio de Janeiro: IBGE, [2026]. "
                    f"Disponível em: https://sidra.ibge.gov.br/tabela/10758. Acesso em: {hoje}.")
    return refs


# ------------------------------------------------------------------------------
# 13. SITUAÇÃO DE PREENCHIMENTO DAS FERRAMENTAS DA V8
# ------------------------------------------------------------------------------
OPCIONAIS_V8 = {"financiamento", "valuation", "impacto", "editais"}


def regras_v8(proj: dict, d: dict) -> dict:
    def n(t):
        x = d.get(t)
        return 0 if x is None else len(x)
    est = d.get("swot_estrategias", pd.DataFrame(columns=["quadrante"]))
    lean = [proj.get(k) for k in ("lean_alternativas", "lean_metricas", "lean_vantagem")]
    imp = [proj.get(k) for k, _, _ in ETAPAS_MUDANCA]
    return {
        "dados_publicos": (n("dados_publicos") >= 3, n("dados_publicos") >= 1),
        "swot_cruzada": (n("swot_estrategias") >= 4 and est["quadrante"].nunique() >= 3, n("swot_estrategias") >= 1),
        "lean": (all(lean), any(lean)),
        "referencias": (bool(proj.get("setor_referencia")), bool(proj.get("setor_referencia"))),
        "financiamento": (n("emprestimos") >= 1, n("emprestimos") >= 1),
        "valuation": (float(proj.get("rodada_investimento") or 0) > 0, float(proj.get("rodada_pre_money") or 0) > 0),
        "impacto": (all(imp) and n("impacto_indicadores") >= 1, any(imp)),
        "editais": (n("edital_textos") >= 1 and n("edital_orcamento") >= 1, n("edital_textos") >= 1),
    }
