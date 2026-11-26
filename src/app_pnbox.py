"""
================================================================================
VIABILIZALAB v8.6 (antigo PNBOX Acadêmico) - PLANO DE NEGÓCIO E VIABILIDADE ECONÔMICO-FINANCEIRA
================================================================================
Programa único (um só arquivo). Formas de abrir:
    - dois cliques no arquivo "Abrir ViabilizaLab.bat" (na mesma pasta); ou
    - dois cliques neste arquivo .py (se o Windows abrir .py com o Python); ou
    - no terminal:  py app_pnbox.py
Na primeira vez ele instala sozinho as bibliotecas que faltarem
(streamlit, pandas, openpyxl, fpdf2, matplotlib) e já abre o navegador,
sem pedir e-mail, numa porta livre escolhida na hora.
Ao fechar a aba do navegador, o programa se encerra sozinho (após ~15 s).

Os projetos ficam na pasta "projetos", criada ao lado deste arquivo.
Partes do código: (1) banco de dados, (2) modelo FitBox, (3) motor financeiro,
(4) relatório PDF, (5) extras (IA, pitch, Word, professor), (6) interface.
================================================================================
"""
from __future__ import annotations

# ==============================================================================
# INICIALIZAÇÃO AUTOMÁTICA (executada antes de tudo)
# ==============================================================================
import importlib.util as _ilu
import subprocess as _sp
import sys as _sys
from pathlib import Path as _Path


def _no_streamlit() -> bool:
    """True quando o arquivo já está rodando dentro do Streamlit."""
    try:
        from streamlit.runtime import exists
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return exists() or get_script_run_ctx(suppress_warning=True) is not None
    except Exception:
        return False


def _preparar_ambiente() -> None:
    pacotes = {"streamlit": "streamlit", "pandas": "pandas", "openpyxl": "openpyxl",
               "matplotlib": "matplotlib", "altair": "altair", "docx": "python-docx"}
    faltando = [pip for mod, pip in pacotes.items() if _ilu.find_spec(mod) is None]
    try:
        import fpdf
        if int(fpdf.__version__.split(".")[0]) < 2 or tuple(map(int, fpdf.__version__.split(".")[:2])) < (2, 7):
            faltando.append("fpdf2")
    except Exception:
        faltando.append("fpdf2")
    if faltando:
        print("Instalando bibliotecas necessárias:", ", ".join(faltando))
        _sp.check_call([_sys.executable, "-m", "pip", "install", "-U", *faltando])
    # Evita a pergunta "Email:" do Streamlit na primeira execução
    cred = _Path.home() / ".streamlit" / "credentials.toml"
    if not cred.exists():
        cred.parent.mkdir(parents=True, exist_ok=True)
        cred.write_text('[general]\nemail = ""\n', encoding="utf-8")


if __name__ == "__main__" and not _no_streamlit():
    _preparar_ambiente()
    import os as _os
    import socket as _socket
    # Porta livre escolhida pelo sistema (evita conflito com outro programa)
    with _socket.socket() as _s:
        _s.bind(("127.0.0.1", 0))
        _porta = _s.getsockname()[1]
    _os.environ["PNBOX_AUTOFECHAR"] = "1"
    print(f"Abrindo o ViabilizaLab em http://localhost:{_porta}  (feche a aba do navegador para encerrar)")
    from streamlit.web import cli as _stcli
    _sys.argv = ["streamlit", "run", str(_Path(__file__).resolve()),
                 f"--server.port={_porta}", "--server.address=localhost",
                 "--browser.gatherUsageStats=false", "--server.headless=false"]
    _sys.exit(_stcli.main())
from collections import Counter
from datetime import date
from datetime import datetime
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from fpdf.fonts import FontFace
from pathlib import Path
import altair as alt
import hashlib
import html as _html
import io
import json
import math
import matplotlib.pyplot as plt
import numpy as np
import os
import pandas as pd
import re
import shutil
import sqlite3
import streamlit as st
import sys
import tempfile
import threading
import time
import unicodedata


# Historico de versoes (tela Novidades), gerado de CHANGELOG.md
NOVIDADES_MD = r'''# Histórico de versões do ViabilizaLab

Programa gratuito de plano de negócio e análise de viabilidade para o ensino de Administração (antigo PNBOX Acadêmico).
Cada versão publicada tem uma etiqueta (*tag*) e uma página em **Releases** no GitHub. A numeração segue o padrão
`maior.menor.correção`: o primeiro número muda quando entra um conjunto grande de ferramentas, o segundo quando entram
funções novas e o terceiro quando há só correções.

---

## v8.6 (25/11/2026): sazonalidade das vendas

**Sazonalidade das vendas**
- Nova ferramenta **Sazonalidade das vendas**: índice de cada mês do calendário (100 = mês médio) e o mês em que o negócio abre.
- **Perfis de partida ilustrativos** para ajustar: varejo de presentes, alimentação, sorvetes e bebidas geladas, cursos, serviços a empresas e turismo.
- O motor financeiro distribui a receita e os custos diretos do Ano 1 pelos meses conforme os índices. O total do ano é preservado: a DRE anual, o VPL e a TIR não mudam; mudam a projeção mensal e o **fluxo de caixa**, que passa a mostrar os meses apertados.
- Gráfico de receita e resultado mês a mês e leitura dos resultados com o mês mais forte, o mais fraco, os meses com prejuízo operacional e o ponto mais baixo do caixa.
- Nova seção no relatório PDF com fator, receita, resultado e saldo de caixa de cada mês.
- Projetos sem sazonalidade informada continuam com os mesmos números.

---

## v8.5 (25/11/2026): leitura automática dos resultados e modo exercício

**Leitura automática dos resultados**
- Ao final das tabelas do programa aparece a caixa **📌 Leitura dos resultados**, com duas a quatro frases que dizem o que os números mostram.
- Há regras para 37 tabelas: produtos, custos, equipe, investimentos, empréstimos, sócios, prazos, segmentos, SWOT, Porter, riscos, pesquisas de preço e entrevistas, canais, funil, orçamento de marketing, fornecedores, cronograma, BCG, objetivos, mix de marketing e outras. Também há leitura para a DRE e para o fluxo de caixa.
- As regras são fixas e não usam IA: o resultado é o mesmo em qualquer computador e o professor sabe de onde veio cada frase.
- **Modo exercício** (menu lateral): o aluno escreve a própria leitura de cada tabela e só depois vê a do programa. As leituras ficam salvas no projeto e saem no relatório PDF, numa seção própria, para o professor avaliar.
- A leitura automática pode ser desligada no menu lateral.

---

## v8.4 (25/11/2026): plano de marketing estruturado

**Plano de marketing estruturado**
- Nova ferramenta **Plano de marketing**, que amarra as ferramentas de marketing na sequência da literatura: análise da situação, objetivos, estratégia, táticas, ações e controle (Smith, 2011; McDonald, 2007; Las Casas, 2011).
- **Situação de cada etapa** (completa, em andamento ou vazia), com atalhos para a ferramenta que alimenta cada uma.
- **Situação:** síntese do diagnóstico escrita pelo grupo, somada ao tamanho do mercado e à comparação de preços com os concorrentes.
- **Objetivos:** os objetivos SMART da perspectiva Clientes.
- **Estratégia:** declaração de posicionamento, direção de crescimento pela matriz de Ansoff e estratégia competitiva de Porter, com alertas quando contradizem a estratégia de preço.
- **Táticas:** resumo do mix montado a partir das escolhas estruturadas, com a coerência do mix.
- **Ações:** o calendário de marketing.
- **Controle:** indicadores com meta, frequência, fonte do dado e responsável, com sugestões prontas.
- O relatório PDF e o **relatório ABNT** ganham o plano de marketing completo, com citações e referências (Smith, McDonald, Las Casas, Kotler e Keller, Booms e Bitner, Ansoff).
- A antiga página "Plano de marketing (4Ps)" passa a se chamar **Mix de marketing (4Ps e 7Ps)**.
- Os modelos Café Mucuri e TecVale trazem o plano de marketing preenchido como exemplo.

---

## v8.3 (25/11/2026): 7Ps de serviços, coerência do mix e 4Cs

**Mix de marketing: 7Ps de serviços, coerência do mix e 4Cs**
- **Pessoas, processos e evidência física** (Booms; Bitner, 1981), nova aba do mix: quem atende e como o atendimento é padronizado, o que o cliente vê e toca, e o **mapa do serviço** (Shostack, 1984), com etapas visíveis ao cliente, bastidores e apoio, tempo de cada etapa e pontos de falha.
- **Coerência entre os Ps:** tabela de combinações que se reforçam ou se contradizem (preço premium com produto ampliado e distribuição seletiva, penetração com comunicação de alcance, desnatação no lançamento, objetivo da comunicação conforme a fase do ciclo de vida, canais cadastrados conforme a intensidade de distribuição).
- **Autoavaliação pelos 4Cs** (Lauterborn, 1990): cliente, custo para o cliente, conveniência e comunicação, com nota, evidência e indicação do ponto mais fraco.
- Os modelos TecVale e Café Mucuri trazem exemplos do mapa do serviço.

---

## v8.2 (25/11/2026): mix de marketing integrado: Praça e Promoção

**Plano de marketing (4Ps): Praça e Promoção**
- **Praça:** intensidade da distribuição (intensiva, seletiva ou exclusiva) e tabela de canais de venda com a parte das vendas e a comissão de cada canal. Um botão aplica a comissão média aos custos variáveis, sem contar duas vezes a comissão de aplicativos que já existia.
- Ligação com o plano operacional para a localização.
- **Promoção:** objetivo principal da comunicação (tornar conhecido, fazer experimentar, converter ou fidelizar) com as ações típicas de cada um.
- **Volume do plano × promoção:** clientes ativos que o primeiro mês exige (receita ÷ gasto mensal por cliente) comparados aos clientes novos que o orçamento de marketing traz, com o tempo para formar a base.
- **CAC e retorno:** custo por cliente novo, meses para recuperar esse custo em margem de contribuição e verba de promoção sobre a receita.
- O diagnóstico do mix ganha alertas de canais que não somam 100%, comissão não aplicada, dependência de um canal de terceiros, posição premium em canais de comparação de preço e volume sem promoção que o sustente.

---

## v8.1 (25/11/2026): mix de marketing integrado: Produto e Preço

**Plano de marketing (4Ps): Produto e Preço**
- A página do 4Ps passa a ser um painel com abas por P e um diagnóstico do mix. Os textos já escritos continuam como justificativa.
- **Produto:** níveis de produto (benefício central, produto em si e produto ampliado), fase do ciclo de vida com a orientação de cada fase, concentração do portfólio (itens que fazem 80% da receita) e aviso quando o diferencial não aparece na proposta de valor nem no posicionamento.
- **Preço:** estratégia de preço (penetração, desnatação, valor percebido, custo mais margem, referência na concorrência) e posição de preço (econômico, intermediário, premium).
- **Preços dos concorrentes** por produto, com o índice de preço (seu preço ÷ média dos concorrentes).
- **Margem de contribuição de cada item** e a conta de equilíbrio: quanto o volume precisa subir para compensar um desconto e quanto pode cair depois de um aumento.
- **Condições de pagamento** (Pix, dinheiro, débito, crédito, parcelado) com taxa e prazo de cada forma. Um botão aplica a taxa média aos custos variáveis e os prazos ao capital de giro.
- **Diagnóstico do mix:** alertas de margem negativa ou estreita, preço incoerente com a estratégia ou com a posição declarada, preço fora da faixa da pesquisa Van Westendorp e dependência de um único item.

---

## v8.0 (25/11/2026): dados públicos, ABNT, financiamento e modo professor completo

**Funções básicas**
- **Trilha guiada e o que falta:** ordem recomendada de preenchimento, próximo passo e lista de pendências com atalho.
- **Glossário e leitura dos números:** 35 termos explicados e a interpretação, em linguagem simples, dos indicadores do projeto aberto.
- **Importar e exportar planilha:** produtos, custos, investimentos, materiais, equipe, entrevistas, teste de preço,
  concorrentes e riscos vindos de Excel ou CSV, com planilha-modelo.
- **Relatório acadêmico ABNT:** capa, folha de rosto, sumário com páginas (NBR 6027), seções numeradas (NBR 6024),
  tabelas com título e fonte, citações autor-data (NBR 10520) e referências (NBR 6023), margens de 3 e 2 cm e fonte 12.
- **Financiamento:** sistemas SAC e Price, carência, custo efetivo total e comparação entre propostas de crédito.
- **Histórico de versões do projeto:** comparação dos indicadores entre as versões salvas.

**Funções completas**
- **Dados públicos:** população, domicílios, PIB e empresas do município pela API do IBGE, e gastos das famílias por
  categoria da POF 2017-2018, com botões para usar os valores no tamanho do mercado.
- **Referências do setor:** comparação com a média nacional da atividade (IBGE, Pesquisa Anual de Comércio e Pesquisa
  Anual de Serviços 2024) e taxas de sobrevivência das empresas (Sebrae).
- **SWOT cruzada:** estratégias ofensivas, de confronto, de reforço e defensivas.
- **Lean Canvas**, montado a partir das outras ferramentas.
- **Valuation e rodada de investimento:** fluxo de caixa descontado com perpetuidade, múltiplos de EBITDA e receita,
  tabela de participação dos sócios antes e depois do aporte.
- **Impacto social e ODS:** teoria da mudança, indicadores de impacto e os 17 Objetivos de Desenvolvimento Sustentável.
- **Editais de fomento:** rascunho da proposta no formato usual do Programa Centelha, com limite de caracteres e
  orçamento por rubrica.
- **Modo professor:** comentários do professor em cada ferramenta (modo de correção), painel de prontidão e evidências
  da turma, detecção de planos parecidos e rodada de investimento da turma.

**Controle de versões**
- Este arquivo, etiquetas e páginas de *release* para todas as versões e a tela **Novidades** dentro do programa.

---

## v7.2.1 (25/11/2026): correção
- O rascunho automático do sumário executivo aparece na hora também na versão on-line.

## v7.2 (25/11/2026): novo visual das matrizes e modelo de serviços
- Matriz GE/McKinsey 3x3 colorida, BCG com quadrantes descritos e bolhas, BSC em ciclo e Canvas com blocos coloridos, na
  tela e no PDF.
- Quarto exemplo: **TecVale Soluções**, empresa de serviços de TI (link `?versao=servico`).
- Correção do texto "None" que aparecia ao gerar o PDF.

## v7.1 (25/11/2026): validação, marketing estratégico, incerteza e defesa
- Placar de evidências, entrevistas com clientes, teste de preço de Van Westendorp e intenção de compra ajustada.
- Canvas da proposta de valor, curva de valor e matriz ERRC, posicionamento e mapa perceptual, orçamento por canal e funil AARRR.
- Simulação de Monte Carlo, perda aceitável dos sócios e curva de adoção de Bass.
- Régua de prontidão, banca simulada e checklist de formalização.
- Teste de viabilidade real no prompt de IA.

## v7.0 (25/11/2026): ViabilizaLab
- Novo nome (antes PNBOX Acadêmico); arquivos `.pnbox` continuam compatíveis.
- Horizonte de projeção de 1 a 10 anos, plano de contratações mês a mês e CAC pela verba de aquisição.
- Missão, visão, valores e objetivos SMART; Business Model Canvas, cinco forças de Porter, matrizes GE e BCG e Balanced Scorecard.
- Simulação da reforma tributária (Simples puro x híbrido).

## v6.0 (24/11/2026): serviços e assinaturas
- Custo da hora de serviço e ocupação da equipe; receita recorrente com safras, churn, MRR, CAC, LTV e receita antecipada.
- Calculadora do Simples Nacional (anexos I a V e fator R).
- Exemplo **Conecta Vale** (aplicativo por assinatura) e layout adaptado para celular.

## v5.0.1 (24/11/2026): ajustes
- Cartões clicáveis com descrição e selo de andamento; organograma legível; instalador sem a janela do Python.

## v5.0 (24/11/2026): plano completo
- Interface em cartões, banco de dados e motor financeiro ampliados.
- Análise com IA, sumário executivo, modo apresentação, relatório em PDF, pitch e Word, modo professor com rubrica,
  configurações, pontos de restauração e o exemplo **Café Mucuri**.

## v4.1 (24/11/2026): primeira versão publicada no GitHub
- Segmentação, proposta de valor, personas, jornada do cliente, concorrência, matriz SWOT, canais e funil, investimentos,
  materiais e fichas técnicas, custos, DRE de 3 anos, indicadores (VPL, TIR, payback, pontos de equilíbrio), cenários,
  gráficos, checagens de consistência e relatório em PDF e Excel.
- Versão on-line (GitHub Pages) e instaladores para Windows (modelo FitBox e em branco).

## v1 a v4.0: desenvolvimento local
- Versões construídas no computador do autor, antes da criação do repositório. Não há registro detalhado de cada etapa;
  o conjunto de funções que elas reuniram é o descrito na v4.1.
'''


# ##############################################################################
# PARTE: pnbox_db.py
# ##############################################################################



# ------------------------------------------------------------------------------
# 1. ESTRUTURA DAS TABELAS
# ------------------------------------------------------------------------------
# Tabelas "únicas" têm apenas uma linha (id = 1). As demais são listas.
TABELAS_UNICAS = {"projeto", "proposta_valor"}
# Tabelas que não são editadas como planilha comum
TABELAS_INTERNAS = {"projeto", "proposta_valor", "versoes"}

SCHEMA = {
    "projeto": [
        ("nome_negocio", "TEXT"), ("objetivo", "TEXT"), ("enquadramento", "TEXT"),
        ("periodo_anos", "INTEGER"), ("porte", "TEXT"), ("setor_atuacao", "TEXT"),
        ("cidade_uf", "TEXT"), ("contatos", "TEXT"),
        ("aliquota_imposto_pct", "REAL"),          # Impostos sobre vendas (ex.: Simples)
        ("crescimento_mensal_ano1_pct", "REAL"),
        ("crescimento_anual_pct", "REAL"),
        ("inflacao_custos_anual_pct", "REAL"),
        ("tma_anual_pct", "REAL"),
        ("pme_dias", "REAL"),                      # Prazo médio de estoques
        ("reserva_caixa_extra", "REAL"),
        ("recuperar_giro", "INTEGER"),             # Recupera capital de giro no último ano do VPL
        ("cen_otim_receita_pct", "REAL"), ("cen_otim_custo_pct", "REAL"),
        ("cen_pess_receita_pct", "REAL"), ("cen_pess_custo_pct", "REAL"),
        # Identificação acadêmica
        ("autores", "TEXT"), ("turma", "TEXT"), ("professor", "TEXT"), ("data_entrega", "TEXT"),
        # Mercado (TAM / SAM / SOM)
        ("tam_desc", "TEXT"), ("tam_clientes", "REAL"), ("sam_desc", "TEXT"), ("sam_clientes", "REAL"),
        ("som_pct", "REAL"), ("gasto_anual_cliente", "REAL"), ("mercado_fontes", "TEXT"),
        # Marketing (4Ps) e localização
        ("mkt_produto", "TEXT"), ("mkt_preco", "TEXT"), ("mkt_praca", "TEXT"), ("mkt_promocao", "TEXT"),
        ("loc_endereco", "TEXT"), ("loc_justificativa", "TEXT"), ("loc_area_m2", "REAL"),
        ("loc_fluxo", "TEXT"),
        # Plano operacional
        ("op_layout", "TEXT"), ("op_processos", "TEXT"), ("op_capacidade_txt", "TEXT"),
        ("op_horario", "TEXT"),
        # Custos e preço
        ("rateio_criterio", "TEXT"), ("margem_desejada_pct", "REAL"),
        # Textos finais
        ("sumario_executivo", "TEXT"), ("conclusao", "TEXT"),
        # Serviços e assinaturas (v6)
        ("tipo_negocio", "TEXT"),                  # Produtos | Serviços | Assinatura | Misto
        ("horas_mes_pessoa", "REAL"),              # horas contratadas por pessoa da operação
        ("utilizacao_pct", "REAL"),                # % dessas horas que viram trabalho cobrável
        ("anexo_simples", "TEXT"),
        ("verba_aquisicao_mensal", "REAL"),         # parte do gasto comercial dedicada a conquistar assinantes
        # Identidade e estratégia (v7)
        ("missao", "TEXT"), ("visao", "TEXT"), ("valores", "TEXT"),
        ("canvas_relacionamento", "TEXT"), ("canvas_parcerias", "TEXT"),
        ("bcg_limite_crescimento_pct", "REAL"),
        # Reforma tributária (simulação)
        ("reforma_b2b_pct", "REAL"), ("reforma_cbs_pct", "REAL"), ("reforma_ibs_pct", "REAL"),
        ("reforma_ano", "REAL"), ("reforma_compras_credito", "REAL"),
        # Validação, marketing e incerteza (v7.1)
        ("intencao_certamente_pct", "REAL"), ("intencao_provavel_pct", "REAL"),
        ("fator_certamente_pct", "REAL"), ("fator_provavel_pct", "REAL"), ("vw_produto", "TEXT"),
        ("vpc_job", "TEXT"), ("vpc_tarefas", "TEXT"), ("vpc_dores", "TEXT"), ("vpc_ganhos", "TEXT"),
        ("vpc_produtos", "TEXT"), ("vpc_aliviadores", "TEXT"), ("vpc_criadores", "TEXT"),
        ("posic_segmento", "TEXT"), ("posic_necessidade", "TEXT"), ("posic_marca", "TEXT"), ("posic_categoria", "TEXT"),
        ("posic_beneficio", "TEXT"), ("posic_concorrente", "TEXT"), ("posic_diferencial", "TEXT"),
        ("mapa_eixo_x", "TEXT"), ("mapa_eixo_y", "TEXT"), ("nps_meta", "REAL"),
        ("bass_mercado", "REAL"), ("bass_p", "REAL"), ("bass_q", "REAL"), ("revenda_ativos_pct", "REAL"),
        # v8: relatório ABNT, dados públicos, Lean Canvas, valuation, impacto e edital
        ("abnt_instituicao", "TEXT"), ("abnt_curso", "TEXT"), ("abnt_disciplina", "TEXT"), ("abnt_orientador", "TEXT"),
        ("abnt_cidade", "TEXT"), ("abnt_ano", "TEXT"), ("municipio_ibge", "TEXT"), ("municipio_nome", "TEXT"),
        ("uf_sigla", "TEXT"), ("setor_referencia", "TEXT"), ("lean_vantagem", "TEXT"), ("lean_alternativas", "TEXT"),
        ("lean_metricas", "TEXT"), ("val_crescimento_perp_pct", "REAL"), ("val_multiplo_ebitda", "REAL"),
        ("val_multiplo_receita", "REAL"), ("rodada_investimento", "REAL"), ("rodada_pre_money", "REAL"),
        ("imp_problema", "TEXT"), ("imp_publico", "TEXT"), ("imp_atividades", "TEXT"), ("imp_produtos", "TEXT"),
        ("imp_resultados", "TEXT"), ("imp_impacto", "TEXT"), ("ods", "TEXT"), ("edital_trl", "REAL"),
        # v8.1: mix de marketing integrado (Produto e Preço)
        ("mix_nucleo", "TEXT"), ("mix_real", "TEXT"), ("mix_ampliado", "TEXT"), ("mix_ciclo", "TEXT"),
        ("mix_estrategia_preco", "TEXT"), ("mix_posicao_preco", "TEXT"),
        # v8.2: Praça e Promoção
        ("mix_distribuicao", "TEXT"), ("mix_objetivo_com", "TEXT"), ("mix_ticket", "REAL"), ("mix_freq", "REAL"),
        # v8.3: 7Ps de serviços e 4Cs
        ("mix_pessoas", "TEXT"), ("mix_processos", "TEXT"), ("mix_evidencia", "TEXT"),
        ("mix_4c_cliente", "TEXT"), ("mix_4c_custo", "TEXT"), ("mix_4c_conveniencia", "TEXT"), ("mix_4c_comunicacao", "TEXT"),
        ("mix_4c_n1", "REAL"), ("mix_4c_n2", "REAL"), ("mix_4c_n3", "REAL"), ("mix_4c_n4", "REAL"),
        # v8.4: plano de marketing
        ("pm_situacao", "TEXT"), ("pm_ansoff", "TEXT"), ("pm_generica", "TEXT"),
        # v8.6: sazonalidade
        ("saz_mes_inicio", "REAL"), ("saz_perfil", "TEXT"),
    ],
    "socios": [("nome", "TEXT"), ("funcao", "TEXT"), ("contato", "TEXT"),
               ("perfil", "TEXT"), ("capital_integralizado", "REAL"), ("perda_aceitavel", "REAL")],
    "segmentos": [("nome", "TEXT"), ("descricao", "TEXT"), ("nota_tamanho", "REAL"),
                  ("nota_lucro", "REAL"), ("nota_longevidade", "REAL"),
                  ("nota_concorrencia", "REAL"), ("nota_aderencia", "REAL")],
    "proposta_valor": [("segmento_alvo", "TEXT"), ("beneficios_desejados", "TEXT"),
                       ("meios_entrega", "TEXT"), ("diferencial_competitivo", "TEXT"),
                       ("entrega_produtos_servicos", "TEXT")],
    "personas": [
        ("nome", "TEXT"), ("idade", "INTEGER"), ("profissao", "TEXT"), ("estado_civil", "TEXT"),
        ("renda_mensal", "REAL"), ("escolaridade", "TEXT"), ("cidade", "TEXT"),
        ("descricao", "TEXT"), ("citacao", "TEXT"), ("habitos", "TEXT"),
        ("caracteristicas", "TEXT"), ("fluencia_digital", "TEXT"), ("redes_sociais", "TEXT"),
        ("quem_influencia", "TEXT"), ("rotina_manha", "TEXT"), ("rotina_tarde", "TEXT"),
        ("rotina_noite", "TEXT"), ("ganhos", "TEXT"), ("desafios", "TEXT"),
        ("duvidas_objecoes", "TEXT"), ("aborrecimentos", "TEXT"), ("experiencia_ideal", "TEXT"),
    ],
    "jornada_cliente": [("persona_nome", "TEXT"), ("macro_fase", "TEXT"), ("etapa", "TEXT"),
                        ("acoes_cliente", "TEXT"), ("pensamentos", "TEXT"),
                        ("pontos_contato", "TEXT"), ("estado_emocional", "TEXT"),
                        ("oportunidades", "TEXT")],
    "concorrentes": [("nome", "TEXT"), ("observacoes", "TEXT")],
    "concorrencia_notas": [("atributo", "TEXT"), ("empresa", "TEXT"), ("nota", "REAL")],
    "swot": [("ambiente", "TEXT"), ("dimensao", "TEXT"), ("fator", "TEXT"),
             ("natureza", "TEXT"), ("impacto", "TEXT"), ("diagnostico", "TEXT"),
             ("plano_acao", "TEXT"), ("priorizado", "INTEGER")],
    "canais": [("canal", "TEXT"), ("atratividade", "REAL"), ("alcance", "REAL"),
               ("cliente", "REAL"), ("relacionamento", "REAL"), ("viabilidade", "REAL")],
    "funil_vendas": [("etapa", "TEXT"), ("acao", "TEXT"), ("orcamento", "REAL"),
                     ("alcancados", "REAL"), ("convertidos", "REAL")],
    "inv_fixos": [("descricao", "TEXT"), ("quantidade", "REAL"), ("valor_unitario", "REAL"),
                  ("vida_util_anos", "REAL"), ("valor_residual_pct", "REAL")],
    "inv_pre": [("descricao", "TEXT"), ("valor", "REAL")],
    "emprestimos": [("descricao", "TEXT"), ("valor", "REAL"), ("taxa_anual_pct", "REAL"),
                    ("prazo_meses", "REAL"), ("carencia_meses", "REAL"), ("sistema", "TEXT")],
    "capital_giro_prazos": [("tipo", "TEXT"), ("prazo_dias", "REAL"), ("percentual", "REAL")],
    "materiais": [("nome", "TEXT"), ("categoria", "TEXT"), ("unidade", "TEXT"), ("valor_unitario", "REAL"),
                  ("qtd_estoque_inicial", "REAL")],
    "catalogo": [("categoria", "TEXT"), ("descricao", "TEXT"), ("preco_venda", "REAL"),
                 ("custo_aquisicao", "REAL"), ("qtd_mensal", "REAL"),
                 ("qtd_estoque_inicial", "REAL"), ("horas_unit", "REAL")],
    "ficha_tecnica": [("produto", "TEXT"), ("material", "TEXT"), ("quantidade", "REAL")],
    "custos_variaveis": [("descricao", "TEXT"), ("percentual", "REAL")],
    "pessoal": [("cargo", "TEXT"), ("area", "TEXT"), ("num_empregados", "REAL"), ("salario_mensal", "REAL"),
                ("encargos_pct", "REAL"), ("beneficios_unit", "REAL"), ("atribuicoes", "TEXT"),
                ("mes_inicio", "REAL")],                 # mês do plano em que o cargo começa (contratações futuras)
    "custos_fixos": [("descricao", "TEXT"), ("classificacao", "TEXT"), ("valor_mensal", "REAL")],
    # --- novas tabelas (v5) ---
    "propostas": [("persona_nome", "TEXT"), ("valores", "TEXT"), ("que_desejam", "TEXT"),
                  ("por_meio", "TEXT"), ("diferentemente", "TEXT"), ("oferecem", "TEXT")],
    "capacidade": [("recurso", "TEXT"), ("produto_ou_grupo", "TEXT"), ("capacidade_mensal", "REAL"),
                   ("observacao", "TEXT")],
    "fornecedores": [("nome", "TEXT"), ("itens", "TEXT"), ("prazo_pagamento_dias", "REAL"),
                     ("prazo_entrega_dias", "REAL"), ("condicoes", "TEXT"), ("critico", "INTEGER")],
    "riscos": [("risco", "TEXT"), ("probabilidade", "REAL"), ("impacto", "REAL"), ("mitigacao", "TEXT"),
               ("responsavel", "TEXT")],
    "experimentos": [("hipotese", "TEXT"), ("experimento", "TEXT"), ("metrica", "TEXT"),
                     ("criterio_sucesso", "TEXT"), ("resultado", "TEXT"), ("status", "TEXT")],
    "cronograma": [("atividade", "TEXT"), ("mes_inicio", "REAL"), ("duracao_meses", "REAL"),
                   ("responsavel", "TEXT"), ("custo", "REAL")],
    "realizado": [("mes", "REAL"), ("receita_real", "REAL"), ("custos_variaveis_real", "REAL"),
                  ("custos_fixos_real", "REAL")],
    "listas": [("tipo", "TEXT"), ("valor", "TEXT")],
    # --- v6: receita recorrente (planos de assinatura) ---
    "assinaturas": [("plano", "TEXT"), ("publico", "TEXT"), ("preco_ciclo", "REAL"), ("ciclo_meses", "REAL"),
                    ("novos_mes1", "REAL"), ("crescimento_novos_pct", "REAL"), ("churn_mensal_pct", "REAL"),
                    ("custo_entrega_mensal", "REAL"), ("taxa_cobranca_pct", "REAL")],
    "versoes": [("nome", "TEXT"), ("data", "TEXT"), ("dados", "TEXT")],
    # --- v7: objetivos SMART, matrizes estratégicas ---
    "objetivos": [("objetivo", "TEXT"), ("perspectiva", "TEXT"), ("indicador", "TEXT"), ("meta", "REAL"),
                  ("unidade", "TEXT"), ("prazo_meses", "REAL"), ("responsavel", "TEXT"), ("por_que_atingivel", "TEXT"),
                  ("relevancia", "TEXT"), ("iniciativa", "TEXT"), ("contribui_para", "TEXT")],
    "bcg": [("produto", "TEXT"), ("crescimento_mercado_pct", "REAL"), ("participacao_relativa", "REAL")],
    "ge": [("unidade", "TEXT"), ("forca", "REAL"), ("justificativa", "TEXT")],
    "porter": [("forca", "TEXT"), ("intensidade", "REAL"), ("justificativa", "TEXT"), ("estrategia", "TEXT")],
    # --- v7.1: validação, marketing, incerteza e defesa ---
    "evidencias": [("premissa", "TEXT"), ("status", "TEXT"), ("fonte", "TEXT"), ("observacao", "TEXT")],
    "entrevistas": [("data", "TEXT"), ("entrevistado", "TEXT"), ("perfil", "TEXT"), ("dores", "TEXT"),
                    ("solucao_atual", "TEXT"), ("pagaria", "TEXT"), ("valor_pagaria", "REAL"), ("aprendizado", "TEXT")],
    "precos_vw": [("respondente", "TEXT"), ("muito_barato", "REAL"), ("barato", "REAL"), ("caro", "REAL"),
                  ("muito_caro", "REAL")],
    "vpc_encaixe": [("item", "TEXT"), ("tipo", "TEXT"), ("resposta", "TEXT")],
    "errc": [("atributo", "TEXT"), ("acao", "TEXT"), ("nova_nota", "REAL"), ("justificativa", "TEXT")],
    "mapa_perceptual": [("marca", "TEXT"), ("eixo_x", "REAL"), ("eixo_y", "REAL")],
    "orcamento_canais": [("canal", "TEXT"), ("gasto_mensal", "REAL"), ("clientes_mes", "REAL"), ("observacao", "TEXT")],
    "aarrr": [("etapa", "TEXT"), ("metrica", "TEXT"), ("meta", "TEXT"), ("acao", "TEXT")],
    "calendario_mkt": [("mes", "REAL"), ("acao", "TEXT"), ("canal", "TEXT"), ("custo", "REAL")],
    "mc_faixas": [("variavel", "TEXT"), ("minimo_pct", "REAL"), ("maximo_pct", "REAL")],
    "formalizacao": [("item", "TEXT"), ("status", "TEXT"), ("observacao", "TEXT")],
    "banca_respostas": [("pergunta", "TEXT"), ("resposta", "TEXT")],
    # --- v8 ---
    "referencias": [("referencia", "TEXT")],
    "swot_estrategias": [("quadrante", "TEXT"), ("estrategia", "TEXT"), ("acao", "TEXT"), ("prazo_meses", "REAL")],
    "impacto_indicadores": [("indicador", "TEXT"), ("tipo", "TEXT"), ("linha_base", "TEXT"), ("meta", "TEXT"), ("ods", "TEXT")],
    "edital_textos": [("secao", "TEXT"), ("texto", "TEXT"), ("limite", "REAL")],
    "edital_orcamento": [("rubrica", "TEXT"), ("item", "TEXT"), ("valor", "REAL")],
    "comentarios": [("pagina", "TEXT"), ("comentario", "TEXT"), ("autor", "TEXT"), ("data", "TEXT"), ("resolvido", "INTEGER")],
    # --- v8.1: mix de marketing integrado ---
    "precos_concorrentes": [("concorrente", "TEXT"), ("produto", "TEXT"), ("preco", "REAL"), ("fonte", "TEXT")],
    "meios_pagamento": [("forma", "TEXT"), ("pct_vendas", "REAL"), ("taxa_pct", "REAL"), ("prazo_dias", "REAL")],
    "canais_venda": [("canal", "TEXT"), ("pct_vendas", "REAL"), ("comissao_pct", "REAL"), ("observacao", "TEXT")],
    "blueprint": [("etapa", "TEXT"), ("linha", "TEXT"), ("tempo_min", "REAL"), ("ponto_falha", "TEXT")],
    "kpis_mkt": [("indicador", "TEXT"), ("meta", "TEXT"), ("frequencia", "TEXT"), ("fonte", "TEXT"), ("responsavel", "TEXT")],
    "sazonalidade": [("mes", "REAL"), ("indice", "REAL")],
    "leituras_aluno": [("tabela", "TEXT"), ("texto", "TEXT")],
    "dados_publicos": [("chave", "TEXT"), ("rotulo", "TEXT"), ("valor", "REAL"), ("periodo", "TEXT"), ("fonte", "TEXT")],
}

# Valores-padrão de um projeto em branco
PROJETO_PADRAO = {
    "nome_negocio": "Novo Projeto", "objetivo": "Criar um novo negócio",
    "enquadramento": "Simples Nacional", "periodo_anos": 3, "porte": "Microempresa (ME)",
    "setor_atuacao": "", "cidade_uf": "", "contatos": "",
    "aliquota_imposto_pct": 6.0, "crescimento_mensal_ano1_pct": 1.0,
    "crescimento_anual_pct": 8.0, "inflacao_custos_anual_pct": 4.5, "tma_anual_pct": 12.0,
    "pme_dias": 15.0, "reserva_caixa_extra": 0.0, "recuperar_giro": 1,
    "cen_otim_receita_pct": 10.0, "cen_otim_custo_pct": 5.0,
    "cen_pess_receita_pct": 15.0, "cen_pess_custo_pct": 5.0,
    "rateio_criterio": "Custo direto", "margem_desejada_pct": 15.0,
    "tipo_negocio": "Produtos", "horas_mes_pessoa": 168.0, "utilizacao_pct": 75.0, "anexo_simples": "",
    "verba_aquisicao_mensal": 0.0, "bcg_limite_crescimento_pct": 10.0,
    "reforma_b2b_pct": 0.0, "reforma_cbs_pct": 11.3, "reforma_ibs_pct": 0.1, "reforma_ano": 2027.0,
    "reforma_compras_credito": -1.0,
    "intencao_certamente_pct": 0.0, "intencao_provavel_pct": 0.0, "fator_certamente_pct": 70.0, "fator_provavel_pct": 30.0,
    "nps_meta": 50.0, "bass_mercado": 0.0, "bass_p": 0.03, "bass_q": 0.38, "revenda_ativos_pct": 40.0,
    "mapa_eixo_x": "Preço (baixo → alto)", "mapa_eixo_y": "Conveniência (baixa → alta)",
    "val_crescimento_perp_pct": 3.0, "val_multiplo_ebitda": 4.0, "val_multiplo_receita": 1.0,
    "rodada_investimento": 0.0, "rodada_pre_money": 0.0, "edital_trl": 3.0, "saz_mes_inicio": 1.0,
}

# Custos variáveis sugeridos para um projeto em branco (percentuais zerados)
CUSTOS_VARIAVEIS_PADRAO = [
    "Comissões de vendas", "Taxas de cartões", "Propaganda variável",
    "Previsão de inadimplência", "Comissões de aplicativos de delivery",
    "Descontos e cupons promocionais", "Frete subsidiado em entregas",
]


# ------------------------------------------------------------------------------
# 2. CONEXÃO E CRIAÇÃO / ATUALIZAÇÃO DA ESTRUTURA
# ------------------------------------------------------------------------------
def conectar(caminho: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(caminho), check_same_thread=False)
    garantir_estrutura(conn)
    return conn


def _colunas(conn, tabela):
    return [r[1] for r in conn.execute(f"PRAGMA table_info({tabela})")]


def _existe(conn, tabela):
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tabela,)
    ).fetchone() is not None


def garantir_estrutura(conn: sqlite3.Connection) -> None:
    """Cria as tabelas que faltam e acrescenta colunas novas em arquivos antigos."""
    for tabela, cols in SCHEMA.items():
        pk = "id INTEGER PRIMARY KEY CHECK (id = 1)" if tabela in TABELAS_UNICAS \
            else "id INTEGER PRIMARY KEY AUTOINCREMENT"
        defs = ", ".join(f"{c} {t}" for c, t in cols)
        conn.execute(f"CREATE TABLE IF NOT EXISTS {tabela} ({pk}, {defs})")
        existentes = _colunas(conn, tabela)
        for c, t in cols:
            if c not in existentes:
                conn.execute(f"ALTER TABLE {tabela} ADD COLUMN {c} {t}")

    # Linha única do projeto
    if conn.execute("SELECT COUNT(*) FROM projeto").fetchone()[0] == 0:
        cols = list(PROJETO_PADRAO)
        conn.execute(
            f"INSERT INTO projeto (id, {', '.join(cols)}) VALUES (1, {', '.join('?' * len(cols))})",
            [PROJETO_PADRAO[c] for c in cols],
        )
    else:  # preenche campos novos que vieram vazios de versões anteriores
        for c, v in PROJETO_PADRAO.items():
            conn.execute(f"UPDATE projeto SET {c}=? WHERE id=1 AND {c} IS NULL", (v,))
    if conn.execute("SELECT COUNT(*) FROM proposta_valor").fetchone()[0] == 0:
        conn.execute("INSERT INTO proposta_valor (id) VALUES (1)")

    _migrar_versao_anterior(conn)
    conn.commit()


def _migrar_versao_anterior(conn):
    """Converte dados de arquivos .pnbox gerados por versões anteriores do programa."""
    conn.execute("UPDATE custos_fixos SET classificacao='Despesa administrativa' "
                 "WHERE classificacao IS NULL OR classificacao=''")
    conn.execute("UPDATE pessoal SET area='Produção / operação' WHERE area IS NULL OR area=''")
    conn.execute("UPDATE catalogo SET horas_unit=0 WHERE horas_unit IS NULL")
    conn.execute("UPDATE pessoal SET mes_inicio=1 WHERE mes_inicio IS NULL OR mes_inicio<1")
    conn.execute("UPDATE emprestimos SET carencia_meses=0 WHERE carencia_meses IS NULL")
    conn.execute("UPDATE emprestimos SET sistema='Price' WHERE sistema IS NULL OR sistema=''")
    conn.execute("UPDATE assinaturas SET ciclo_meses=1 WHERE ciclo_meses IS NULL OR ciclo_meses<1")
    # Proposta de valor única (v1-v4) -> lista de propostas
    if conn.execute("SELECT COUNT(*) FROM propostas").fetchone()[0] == 0:
        pv = conn.execute("SELECT segmento_alvo, beneficios_desejados, meios_entrega, diferencial_competitivo, "
                          "entrega_produtos_servicos FROM proposta_valor WHERE id=1").fetchone()
        if pv and any(pv):
            conn.execute("INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, "
                         "oferecem) VALUES (?,?,?,?,?,?)", (pv[0], "", pv[1], pv[2], pv[3], pv[4]))
    # Catálogo: custo_unitario -> custo_aquisicao
    if "custo_unitario" in _colunas(conn, "catalogo"):
        conn.execute("UPDATE catalogo SET custo_aquisicao = custo_unitario "
                     "WHERE custo_aquisicao IS NULL")
    # Investimentos em tabela única -> tabelas separadas
    if _existe(conn, "investimentos") and \
            conn.execute("SELECT COUNT(*) FROM inv_fixos").fetchone()[0] == 0 and \
            conn.execute("SELECT COUNT(*) FROM inv_pre").fetchone()[0] == 0:
        for tipo, desc, qtd, vu, vida in conn.execute(
                "SELECT tipo, descricao, quantidade, valor_unitario, vida_util_anos FROM investimentos"):
            if tipo == "Fixo":
                conn.execute("INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, "
                             "vida_util_anos, valor_residual_pct) VALUES (?,?,?,?,0)",
                             (desc, qtd, vu, vida))
            elif tipo == "Pré-Operacional":
                conn.execute("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)",
                             (desc, (qtd or 0) * (vu or 0)))
            else:
                conn.execute("INSERT INTO materiais (nome, unidade, valor_unitario, "
                             "qtd_estoque_inicial) VALUES (?, 'unidade', ?, ?)", (desc, vu, qtd))
    # Dívidas da v2 -> tabela de empréstimos
    cols_proj = _colunas(conn, "projeto")
    if "divida_cp_inicial" in cols_proj and \
            conn.execute("SELECT COUNT(*) FROM emprestimos").fetchone()[0] == 0:
        cp, lp, tx = conn.execute(
            "SELECT divida_cp_inicial, divida_lp_inicial, taxa_juros_divida_pct FROM projeto").fetchone()
        if cp:
            conn.execute("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) "
                         "VALUES ('Empréstimo de curto prazo', ?, ?, 12)", (cp, tx or 0))
        if lp:
            conn.execute("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) "
                         "VALUES ('Empréstimo de longo prazo', ?, ?, 48)", (lp, tx or 0))
    # Concorrência em formato largo -> formato longo
    if "concorrente_1" in _colunas(conn, "concorrencia") if _existe(conn, "concorrencia") else False:
        if conn.execute("SELECT COUNT(*) FROM concorrencia_notas").fetchone()[0] == 0:
            nomes = {"nossa_empresa": "Nossa Empresa", "concorrente_1": "Concorrente 1",
                     "concorrente_2": "Concorrente 2", "concorrente_3": "Concorrente 3"}
            for n in list(nomes.values())[1:]:
                conn.execute("INSERT INTO concorrentes (nome) VALUES (?)", (n,))
            for row in conn.execute("SELECT atributo, nossa_empresa, concorrente_1, "
                                    "concorrente_2, concorrente_3 FROM concorrencia").fetchall():
                for (col, nome), nota in zip(nomes.items(), row[1:]):
                    conn.execute("INSERT INTO concorrencia_notas (atributo, empresa, nota) "
                                 "VALUES (?,?,?)", (row[0], nome, nota))


# ------------------------------------------------------------------------------
# 3. LEITURA E GRAVAÇÃO
# ------------------------------------------------------------------------------
def colunas_de(tabela):
    return [c for c, _ in SCHEMA[tabela]]


def ler(conn, tabela) -> pd.DataFrame:
    cols = colunas_de(tabela)
    df = pd.read_sql(f"SELECT {', '.join(cols)} FROM {tabela} ORDER BY id", conn)
    return limpar(df, tabela)


def limpar(df: pd.DataFrame, tabela) -> pd.DataFrame:
    """Garante colunas, troca vazios numéricos por 0 e textos vazios por ''."""
    df = df.copy()
    for c, t in SCHEMA[tabela]:
        if c not in df.columns:
            df[c] = None
        if t in ("REAL", "INTEGER"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
            if c in ("priorizado", "critico"):
                df[c] = df[c].astype(bool)
        else:
            df[c] = df[c].astype(object).where(df[c].notna(), "")
    return df[colunas_de(tabela)].reset_index(drop=True)


def salvar(conn, tabela, df: pd.DataFrame, chave: str | None = None) -> None:
    """Substitui o conteúdo da tabela pelo DataFrame editado."""
    df = df.copy()
    if chave:
        df = df[df[chave].astype(str).str.strip() != ""]
        df = df[df[chave].notna()]
    df = limpar(df, tabela)
    conn.execute(f"DELETE FROM {tabela}")
    if not df.empty:
        df.to_sql(tabela, conn, if_exists="append", index=False)
    conn.commit()


def ler_unica(conn, tabela) -> dict:
    cols = colunas_de(tabela)
    row = conn.execute(f"SELECT {', '.join(cols)} FROM {tabela} WHERE id=1").fetchone()
    d = dict(zip(cols, row))
    for c, t in SCHEMA[tabela]:
        if t in ("REAL", "INTEGER"):
            d[c] = float(d[c] or 0)
        else:
            d[c] = d[c] or ""
    return d


def salvar_unica(conn, tabela, dados: dict) -> None:
    cols = [c for c in colunas_de(tabela) if c in dados]
    conn.execute(f"UPDATE {tabela} SET {', '.join(c + '=?' for c in cols)} WHERE id=1",
                 [dados[c] for c in cols])
    conn.commit()


# ------------------------------------------------------------------------------
# 4. CRIAÇÃO DE PROJETOS
# ------------------------------------------------------------------------------
def criar_projeto_em_branco(caminho: Path, nome: str) -> None:
    conn = conectar(caminho)
    salvar_unica(conn, "projeto", {"nome_negocio": nome})
    conn.executemany("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?, 0)",
                     [(d,) for d in CUSTOS_VARIAVEIS_PADRAO])
    conn.executemany("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)",
                     [("Recebimento", 0, 100.0), ("Pagamento", 0, 100.0)])
    conn.commit()
    conn.close()


def criar_projeto_modelo(caminho: Path) -> None:
    """Cria o estudo de caso FitBox (inspirado no relatório PNBOX do Sebrae)."""
    conn = conectar(caminho)
    popular_modelo_fitbox(conn)
    popular_estrategia(conn, "fitbox")
    conn.commit()
    conn.close()


def criar_projeto_cafe(caminho: Path) -> None:
    """Cria o segundo estudo de caso: Café Mucuri (cafeteria e padaria artesanal)."""
    conn = conectar(caminho)
    popular_modelo_cafe(conn)
    popular_estrategia(conn, "cafe")
    conn.commit()
    conn.close()


def criar_projeto_servico(caminho: Path) -> None:
    """Cria o quarto estudo de caso: TecVale Soluções (serviços de TI por hora, por projeto e por contrato)."""
    conn = conectar(caminho)
    popular_modelo_servico(conn)
    popular_estrategia(conn, "servico")
    conn.commit()
    conn.close()


def criar_projeto_assinatura(caminho: Path) -> None:
    """Cria o terceiro estudo de caso: Conecta Vale (aplicativo com planos de assinatura)."""
    conn = conectar(caminho)
    popular_modelo_assinatura(conn)
    popular_estrategia(conn, "assinatura")
    conn.commit()
    conn.close()

# ##############################################################################
# PARTE: pnbox_modelo.py
# ##############################################################################


def popular_modelo_fitbox(conn) -> None:
    ins = conn.executemany

    conn.execute("""
        UPDATE projeto SET nome_negocio='FitBox', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Pequena Empresa (EPP)',
            setor_atuacao='Comércio; Serviço; Alimentos e bebidas; Bares e restaurantes; Varejo; Moda',
            cidade_uf='Teófilo Otoni - MG', contatos='@fitbox.br',
            aliquota_imposto_pct=10.0, crescimento_mensal_ano1_pct=1.0, crescimento_anual_pct=8.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=12.0, pme_dias=15, reserva_caixa_extra=150000,
            recuperar_giro=1, cen_otim_receita_pct=10, cen_otim_custo_pct=5,
            cen_pess_receita_pct=15, cen_pess_custo_pct=5
        WHERE id=1""")

    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Alessandra Ribeiro Gonçalves", "Sócia-administradora", "",
         "Graduanda em Administração; planejamento operacional e financeiro, controle de custos e fornecedores.", 100000.0),
        ("Paulo Henrique Batista", "Sócio-administrador", "",
         "Gestão estratégica, marketing, posicionamento da marca e relacionamento com clientes.", 80000.0),
        ("Camila Gomes de Souza", "Sócia comercial", "",
         "Desenvolvimento comercial, parcerias com academias e varejo fitness.", 60000.0),
        ("Cauan Ribeiro de Souza", "Sócio operacional", "",
         "Acompanhamento de indicadores, expansão e consolidação da empresa no mercado.", 43975.0),
    ])

    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Alimentação saudável / Conveniência",
         "Trabalhadores, estudantes e consumidores que buscam refeições equilibradas, saborosas e práticas.", 10, 9, 10, 6, 9),
        ("Praticantes de atividades físicas",
         "Pessoas que praticam academia ou esportes e buscam alimentação equilibrada, suplementos e roupas fitness.", 8, 9, 9, 7, 10),
        ("Iniciantes no estilo de vida saudável",
         "Pessoas que desejam melhorar hábitos sem dietas restritivas.", 9, 7, 9, 7, 8),
        ("Consumidores de lifestyle fitness",
         "Público engajado em bem-estar, moda esportiva e experiências.", 7, 8, 8, 6, 9),
    ])

    conn.execute("""UPDATE proposta_valor SET
        segmento_alvo='trabalhadores, estudantes e praticantes de atividades físicas',
        beneficios_desejados='conforto, conveniência, economia de tempo, transparência nutricional e alimentação equilibrada',
        meios_entrega='refeições saborosas e personalizáveis, planos semanais, delivery ágil e loja integrada de suplementos e moda fitness',
        diferencial_competitivo='estabelecimentos focados apenas em marmitas tradicionais ou somente em artigos esportivos',
        entrega_produtos_servicos='praticidade, variedade, sabor e controle nutricional em uma experiência acessível e acolhedora.'
        WHERE id=1""")

    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, caracteristicas, fluencia_digital, redes_sociais, quem_influencia,
           rotina_manha, rotina_tarde, rotina_noite, ganhos, desafios, duvidas_objecoes, aborrecimentos,
           experiencia_ideal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Mariana", 24, "Analista Administrativa", "Solteira", 3500.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Trabalha em horário comercial e busca conciliar carreira, vida pessoal e alimentação equilibrada. "
         "Valoriza praticidade, sabor, variedade e informações nutricionais confiáveis.",
         "Quero cuidar da minha saúde, mas minha rotina nem sempre permite cozinhar.",
         "Pratica exercícios algumas vezes por semana; pesquisa preços, cardápios e avaliações pelo celular.",
         "#Organizada #Pragmática #Responsável #Sociável #Consciente #Exigente",
         "Computador: alta; Smartphone: alta; Aplicativos: alta; Compras on-line: média",
         "Instagram, WhatsApp, TikTok",
         "Amigos, colegas de trabalho, especialistas, avaliações de outros clientes",
         "Acorda às 6h30, toma café rápido, verifica mensagens e inicia o expediente às 8h.",
         "Reuniões e tarefas administrativas; no almoço prefere refeições equilibradas e sem longa espera.",
         "Vai à academia; nem sempre tem disposição para cozinhar e considera pedir refeições prontas.",
         "Economizar tempo; manter alimentação saudável; conhecer calorias e nutrientes; personalizar refeições.",
         "Falta de tempo; preços elevados; poucas opções práticas; demora na entrega; falta de informação sobre ingredientes.",
         "As refeições têm boa qualidade? Os preços cabem no orçamento? O delivery chega no horário?",
         "Informações nutricionais incorretas, cobranças inesperadas, atrasos sem aviso, cardápio confuso.",
         "Acessa um cardápio intuitivo com preços e macros, personaliza o pedido e recebe a refeição fresca no horário."),
        ("Rafael", 26, "Técnico em Informática", "Solteiro", 3200.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Pratica musculação diariamente e busca refeições proteicas, porções adequadas e suplementos confiáveis.",
         "Treinar faz parte da minha rotina; preciso de proteína na medida certa.",
         "Treina após o trabalho; acompanha influenciadores fitness; compra suplementos on-line.",
         "#Disciplinado #Focado #Curioso #Exigente",
         "Computador: alta; Smartphone: alta; Aplicativos: alta; Compras on-line: alta",
         "Instagram, YouTube, WhatsApp",
         "Personal trainer, nutricionista, amigos da academia",
         "Acorda às 6h30, organiza lanches proteicos e acompanha conteúdos de treino.",
         "Suporte técnico; busca refeições com alto teor proteico e porções bem dimensionadas.",
         "Musculação; compra refeição pós-treino e pesquisa suplementos.",
         "Aumentar performance; saber gramatura de proteínas; comprar suplementos originais.",
         "Porções pequenas; falta de transparência nutricional; suplementos de origem duvidosa.",
         "Os suplementos são originais? As porções atendem à minha meta de proteína?",
         "Produtos falsificados, porções menores que o anunciado, falta de tabela nutricional.",
         "Encontra em um só lugar refeição pós-treino, suplementos originais e roupas para treinar."),
    ])

    etapas = [
        ("Antes", "Descoberta", "Busca alternativas práticas para almoçar bem no trabalho.", "Onde como bem perto do trabalho?", "Instagram, Google e indicações", "Neutro", "Conteúdos sobre rotina saudável e fotos reais dos pratos."),
        ("Antes", "Consideração", "Compara preços, porções e informações nutricionais.", "Vale a pena pelo preço?", "Cardápio digital e WhatsApp", "Neutro", "Exibir macros, ingredientes e descontos de pacotes semanais."),
        ("Antes", "Decisão", "Verifica prazo de entrega e escolhe a refeição.", "Vai chegar a tempo?", "Canal de pedidos e Google Avaliações", "Satisfeito", "Pedidos agendados para o horário exato do intervalo."),
        ("Durante", "Compra e Pagamento", "Personaliza acompanhamentos e paga via Pix/cartão.", "Foi fácil pedir.", "Cardápio digital e checkout", "Satisfeito", "Checkout rápido sem taxas ocultas."),
        ("Durante", "Entrega", "Recebe a marmita lacrada e identificada no trabalho.", "Chegou quentinha.", "Entregador e embalagem térmica", "Satisfeito", "Embalagem sustentável com lacre e etiqueta nutricional."),
        ("Durante", "Uso", "Consome a refeição e avalia sabor, temperatura e saciedade.", "Está muito bom!", "Produto e apresentação", "Entusiasmado", "Padronização rigorosa de temperos e gramaturas."),
        ("Depois", "Avaliação", "Responde pesquisa rápida e avalia no Google.", "Vou recomendar.", "WhatsApp Business e Google", "Satisfeito", "Cupom de retorno para quem avalia o pedido."),
        ("Depois", "Suporte", "Tira dúvidas sobre ingredientes ou conservação.", "Posso congelar?", "Atendimento WhatsApp", "Neutro", "Atendimento ágil e política clara de troca."),
        ("Depois", "Lealdade e advocacia", "Assina plano mensal e indica colegas de trabalho.", "Virou parte da rotina.", "Programa de fidelidade", "Entusiasmado", "Benefícios progressivos para assinantes."),
    ]
    ins("""INSERT INTO jornada_cliente (persona_nome, macro_fase, etapa, acoes_cliente, pensamentos,
           pontos_contato, estado_emocional, oportunidades) VALUES ('Mariana',?,?,?,?,?,?,?)""", etapas)

    concs = ["Concorrente A (marmitaria fitness)", "Concorrente B (restaurante natural)",
             "Concorrente C (loja de suplementos)"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {
        "Atendimento": (9, 8, 7, 8), "Canais de venda": (9, 7, 8, 6),
        "Conveniência": (9, 8, 7, 6), "Design": (8, 7, 8, 6),
        "Experiência do cliente": (9, 8, 7, 7), "Inovação": (9, 6, 7, 6),
        "Preço": (7, 8, 7, 8), "Presença digital": (9, 7, 6, 8),
        "Qualidade": (9, 8, 9, 8), "Variedade": (10, 7, 7, 5),
        "Personalização": (9, 6, 6, 3), "Transparência nutricional": (10, 5, 4, 6),
        "Alimentação equilibrada": (10, 9, 8, 4), "Sabor dos produtos": (9, 8, 8, 6),
    }
    linhas = []
    for atr, vals in notas.items():
        for emp, v in zip(["Nossa Empresa"] + concs, vals):
            linhas.append((atr, emp, v))
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)", linhas)

    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Pessoas", "Treinamento da equipe", "Força", "Alto", "Programa de capacitação em higiene, atendimento e vendas.", "Calendário trimestral de capacitação.", 0),
        ("Interno", "Pessoas", "Liderança visionária", "Força", "Alto", "Proposta estratégica definida integrando alimentação e produtos fitness.", "Metas, indicadores e reuniões periódicas entre os sócios.", 1),
        ("Interno", "Marketing", "Linhas de produtos e serviços", "Força", "Alto", "Refeições, suplementos, roupas e acessórios em um só lugar.", "Ampliar portfólio conforme a demanda e criar combos.", 1),
        ("Interno", "Marketing", "Presença digital", "Força", "Alto", "Redes sociais e canais de atendimento estruturados.", "Calendário editorial e anúncios segmentados.", 0),
        ("Interno", "Finanças", "Controles financeiros", "Força", "Alto", "Controle de receitas, despesas, fluxo de caixa e indicadores.", "Auditar CMV semanalmente e revisar margens.", 0),
        ("Interno", "Operação", "Controle de qualidade", "Força", "Alto", "Procedimentos de higiene, segurança alimentar e padronização.", "Checklists diários de temperatura e validade.", 1),
        ("Interno", "Operação", "Nível de automação", "Força", "Baixo", "Automação gradual de pedidos e pagamentos.", "Implantar sistema de pedidos integrado ao PDV.", 0),
        ("Interno", "Marketing", "Reputação e imagem", "Fraqueza", "Alto", "Empreendimento novo, sem avaliações de clientes.", "Ações de lançamento e incentivo a avaliações autênticas.", 1),
        ("Interno", "Marketing", "Participação de mercado", "Fraqueza", "Alto", "Início sem participação consolidada.", "Parcerias com academias e empresas locais.", 1),
        ("Interno", "Operação", "Instalações e equipamentos", "Fraqueza", "Alto", "Estrutura física ainda a viabilizar.", "Planejar estrutura conforme capacidade produtiva.", 1),
        ("Interno", "Operação", "Relação com fornecedores", "Fraqueza", "Alto", "Fornecedores ainda não consolidados.", "Homologar ao menos 2 fornecedores por insumo crítico.", 0),
        ("Externo", "Social", "Preocupação com saúde", "Oportunidade", "Alto", "Maior procura por alimentação equilibrada.", "Divulgar benefícios com informações nutricionais confiáveis.", 1),
        ("Externo", "Social", "Busca por qualidade de vida", "Oportunidade", "Alto", "Interesse por bem-estar e hábitos saudáveis.", "Planos de alimentação e campanhas de bem-estar.", 1),
        ("Externo", "Tecnológico", "Hábitos de consumo digitais", "Oportunidade", "Alto", "Valorização da praticidade e das compras digitais.", "Pedidos pelo WhatsApp e programa de fidelidade.", 1),
        ("Externo", "Econômico", "Inflação", "Ameaça", "Alto", "Alta de alimentos, embalagens e energia pressiona custos.", "Negociar com fornecedores e revisar preços periodicamente.", 1),
        ("Externo", "Econômico", "Renda dos consumidores", "Ameaça", "Alto", "Renda limitada reduz compras recorrentes.", "Refeições econômicas e combos promocionais.", 1),
        ("Externo", "Político-Legal", "Fiscalização e regulação", "Ameaça", "Alto", "Exigências sanitárias e risco de penalidades.", "Licenças atualizadas e consultoria técnica.", 0),
        ("Externo", "Econômico", "Concorrência de aplicativos", "Ameaça", "Baixo", "Restaurantes em apps de delivery com promoções.", "Diferenciar pela transparência nutricional.", 0),
    ])

    ins("""INSERT INTO canais (canal, atratividade, alcance, cliente, relacionamento, viabilidade)
           VALUES (?,?,?,?,?,?)""", [
        ("Marketing de Conteúdo", 9, 9, 9, 8, 10), ("Publicidade em Redes Sociais", 9, 9, 10, 8, 10),
        ("Marketing Viral", 9, 9, 9, 8, 9), ("Comunidade", 8, 8, 8, 8, 7),
        ("Plataformas Existentes", 8, 7, 7, 7, 8), ("Acordos e Parcerias", 7, 6, 6, 7, 7),
        ("Eventos", 7, 6, 5, 7, 7),
    ])
    ins("INSERT INTO funil_vendas (etapa, acao, orcamento, alcancados, convertidos) VALUES (?,?,?,?,?)", [
        ("Topo do Funil", "Anúncios Instagram/Meta Ads + parcerias com academias", 1200.0, 25000, 3200),
        ("Meio do Funil", "Cardápio digital interativo + cupom de primeira compra", 500.0, 3200, 950),
        ("Fundo do Funil", "Atendimento consultivo no WhatsApp + oferta de assinatura", 300.0, 950, 410),
    ])

    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,10)""", [
        ("Forno combinado profissional 6 GN", 1, 34000, 10), ("Fogão industrial inox 6 bocas", 2, 4000, 10),
        ("Chapa profissional para grelhados", 1, 4500, 10), ("Processador industrial de alimentos", 1, 5500, 10),
        ("Liquidificador industrial 8 litros", 2, 1800, 10), ("Extrator profissional de sucos", 1, 3600, 10),
        ("Seladora a vácuo profissional", 1, 5500, 10), ("Balança comercial 30 kg", 3, 750, 10),
        ("Refrigerador industrial inox 4 portas", 2, 8500, 10), ("Freezer vertical comercial", 2, 6500, 10),
        ("Balcão refrigerado de preparo", 1, 8500, 10), ("Expositor refrigerado de bebidas", 1, 11000, 10),
        ("Máquina de gelo profissional", 1, 7000, 10), ("Coifa e exaustão industrial", 1, 11000, 10),
        ("Lava-louças profissional", 1, 18000, 10), ("Máquina de café expresso", 1, 7500, 10),
        ("Bancada inox de manipulação", 3, 3000, 10), ("Panelas e utensílios profissionais (conjunto)", 2, 4000, 10),
        ("Mesas de atendimento para 4 lugares", 12, 900, 10), ("Cadeiras de atendimento", 48, 320, 10),
        ("Balcão de atendimento e caixa", 1, 7000, 10), ("Móveis expositores para suplementos e roupas", 5, 1700, 10),
        ("Sistema de climatização instalado", 1, 18000, 10), ("Instalações elétrica, hidráulica e de gás", 1, 35000, 10),
        ("Computadores, PDV e tablets", 1, 17800, 5), ("Impressoras, nobreak e rede", 1, 9500, 5),
        ("Motocicleta para entregas", 2, 16000, 5),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Reformas, obras e adequação sanitária", 75000), ("Legalização, alvarás e licenças sanitárias", 12000),
        ("Projeto arquitetônico e de cozinha", 15000), ("Marca e identidade visual", 8000),
        ("Softwares de gestão e cardápio digital", 9500), ("Marketing de inauguração", 25000),
        ("Treinamento inicial da equipe", 10000), ("Taxas, cauções e despesas de implantação", 10000),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Capital de giro (Pronampe)", 120000, 14.0, 36),
        ("Financiamento de equipamentos (BDMG)", 280000, 12.0, 60),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 55.0), ("Recebimento", 15, 25.0), ("Recebimento", 30, 15.0), ("Recebimento", 60, 5.0),
        ("Pagamento", 0, 15.0), ("Pagamento", 15, 20.0), ("Pagamento", 30, 40.0),
        ("Pagamento", 45, 15.0), ("Pagamento", 60, 10.0),
    ])

    ins("INSERT INTO materiais (nome, categoria, unidade, valor_unitario, qtd_estoque_inicial) VALUES (?,?,?,?,?)", [
        ("Peito de frango", "Matéria-prima", "kg", 22.00, 120), ("Carne bovina (patinho)", "Matéria-prima", "kg", 42.00, 60),
        ("Salmão", "Matéria-prima", "kg", 85.00, 15), ("Ovos", "Matéria-prima", "dúzia", 12.00, 60),
        ("Arroz integral", "Matéria-prima", "kg", 8.00, 80), ("Feijão carioca", "Matéria-prima", "kg", 9.00, 50),
        ("Batata-doce", "Matéria-prima", "kg", 6.00, 80), ("Brócolis", "Matéria-prima", "kg", 14.00, 40),
        ("Cenoura", "Matéria-prima", "kg", 6.00, 45), ("Abobrinha", "Matéria-prima", "kg", 7.00, 35),
        ("Rúcula", "Matéria-prima", "maço", 4.00, 60), ("Tomate", "Matéria-prima", "kg", 8.00, 65),
        ("Cebola", "Matéria-prima", "kg", 6.00, 50), ("Alho", "Matéria-prima", "kg", 30.00, 12),
        ("Azeite extra virgem", "Matéria-prima", "L", 55.00, 25), ("Sal", "Matéria-prima", "kg", 4.00, 15),
        ("Wrap integral", "Matéria-prima", "pacote", 12.00, 65), ("Banana", "Matéria-prima", "kg", 6.00, 75),
        ("Morango", "Matéria-prima", "kg", 22.00, 25), ("Leite desnatado", "Matéria-prima", "L", 6.00, 110),
        ("Iogurte natural", "Matéria-prima", "kg", 13.00, 65), ("Whey protein para preparo", "Matéria-prima", "kg", 95.00, 15),
        ("Polpa de fruta", "Matéria-prima", "kg", 17.00, 90), ("Marmita térmica 750 ml", "Embalagem", "unidade", 0.85, 2500),
        ("Copo 500 ml com tampa", "Embalagem", "unidade", 0.85, 1200), ("Etiqueta lacre", "Embalagem", "unidade", 0.10, 4000),
        ("Kit talher descartável", "Embalagem", "unidade", 0.45, 2000),
        ("Detergente e sanitizante", "Limpeza e higiene", "galão", 36.00, 20), ("Luva e touca descartáveis", "Limpeza e higiene", "caixa", 30.00, 30),
    ])

    F, S, R = "Fabricação Própria", "Serviço", "Revenda"
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal,
           qtd_estoque_inicial) VALUES (?,?,?,?,?,?)""", [
        (S, "Assinatura semanal - 5 marmitas tradicionais", 129.90, 42.20, 78, 0),
        (S, "Assinatura mensal - 20 marmitas tradicionais", 499.90, 168.80, 59, 0),
        (S, "Personalização de pedido (adicional)", 5.90, 1.70, 244, 0),
        (S, "Coffee break saudável para eventos (por pessoa)", 34.90, 16.98, 116, 0),
        (S, "Delivery até 3 km", 6.90, 4.20, 438, 0),
        (S, "Delivery de 3 a 10 km", 9.90, 6.50, 1168, 0),
        (F, "Marmita fitness tradicional", 29.90, 0, 876, 0),
        (F, "Marmita fitness premium", 39.90, 12.80, 291, 0),
        (F, "Marmita fitness econômica", 24.90, 6.90, 487, 0),
        (F, "Marmita low carb", 34.90, 11.20, 340, 0),
        (F, "Marmita hiperproteica", 39.90, 0, 291, 0),
        (F, "Marmita vegetariana", 29.90, 8.90, 98, 0),
        (F, "Prato executivo saudável", 36.90, 11.90, 389, 0),
        (F, "Bowl de frango e vegetais", 34.90, 10.80, 175, 0),
        (F, "Bowl de salmão", 54.90, 0, 59, 0),
        (F, "Salada completa com frango", 32.90, 9.80, 155, 0),
        (F, "Wrap integral de frango", 24.90, 0, 210, 0),
        (F, "Tapioca de frango", 18.90, 5.20, 224, 0),
        (F, "Açaí proteico 300 ml", 26.90, 9.20, 168, 0),
        (F, "Suco natural 500 ml", 15.90, 0, 560, 0),
        (F, "Smoothie de morango", 19.90, 0, 252, 0),
        (F, "Shake proteico com whey", 24.90, 0, 280, 0),
        (F, "Café expresso", 6.90, 1.40, 700, 0),
        (R, "Whey protein concentrado 900 g", 149.90, 90.00, 42, 40),
        (R, "Whey protein isolado 900 g", 229.90, 145.00, 17, 15),
        (R, "Creatina monohidratada 300 g", 79.90, 45.00, 49, 30),
        (R, "Pré-treino 300 g", 119.90, 70.00, 17, 15),
        (R, "Barra de proteína", 12.90, 6.00, 308, 200),
        (R, "Pasta de amendoim 500 g", 24.90, 14.00, 63, 40),
        (R, "Mix de castanhas 200 g", 29.90, 17.00, 56, 40),
        (R, "Água mineral 500 ml", 4.90, 1.30, 560, 300),
        (R, "Legging fitness feminina", 119.90, 60.00, 25, 20),
        (R, "Camiseta dry fit masculina", 79.90, 38.00, 28, 25),
        (R, "Coqueteleira 600 ml", 34.90, 15.00, 35, 30),
    ])

    ficha = {
        "Marmita fitness tradicional": [("Peito de frango", 0.15), ("Arroz integral", 0.10), ("Feijão carioca", 0.08),
                                        ("Brócolis", 0.06), ("Cenoura", 0.05), ("Cebola", 0.02), ("Alho", 0.005),
                                        ("Azeite extra virgem", 0.01), ("Sal", 0.003), ("Marmita térmica 750 ml", 1),
                                        ("Etiqueta lacre", 1), ("Kit talher descartável", 1)],
        "Marmita hiperproteica": [("Peito de frango", 0.25), ("Batata-doce", 0.15), ("Brócolis", 0.08),
                                  ("Ovos", 0.17), ("Azeite extra virgem", 0.01), ("Sal", 0.003),
                                  ("Marmita térmica 750 ml", 1), ("Etiqueta lacre", 1), ("Kit talher descartável", 1)],
        "Bowl de salmão": [("Salmão", 0.15), ("Arroz integral", 0.10), ("Abobrinha", 0.05), ("Rúcula", 0.1),
                           ("Azeite extra virgem", 0.01), ("Marmita térmica 750 ml", 1), ("Etiqueta lacre", 1),
                           ("Kit talher descartável", 1)],
        "Wrap integral de frango": [("Wrap integral", 0.1), ("Peito de frango", 0.10), ("Rúcula", 0.1),
                                    ("Tomate", 0.05), ("Iogurte natural", 0.03), ("Etiqueta lacre", 1)],
        "Suco natural 500 ml": [("Polpa de fruta", 0.15), ("Copo 500 ml com tampa", 1)],
        "Smoothie de morango": [("Morango", 0.15), ("Iogurte natural", 0.15), ("Banana", 0.10),
                                ("Copo 500 ml com tampa", 1)],
        "Shake proteico com whey": [("Whey protein para preparo", 0.03), ("Leite desnatado", 0.30),
                                    ("Banana", 0.10), ("Copo 500 ml com tampa", 1)],
    }
    ins("INSERT INTO ficha_tecnica (produto, material, quantidade) VALUES (?,?,?)",
        [(p, m, q) for p, itens in ficha.items() for m, q in itens])

    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [
        ("Comissões de vendas", 0.5), ("Taxas de cartões", 1.5), ("Propaganda variável", 0.8),
        ("Previsão de inadimplência", 0.3), ("Comissões de aplicativos de delivery", 0.5),
        ("Descontos e cupons promocionais", 0.5), ("Frete subsidiado em entregas", 0.5),
    ])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Gerente operacional", A, 1, 3800, 20, 200, "Coordena equipe, compras, escalas e indicadores."),
        ("Chef de cozinha / líder", P, 1, 3200, 20, 200, "Cardápio, fichas técnicas e padrão de qualidade."),
        ("Cozinheiros", P, 2, 2400, 20, 200, "Preparo das refeições conforme fichas técnicas."),
        ("Auxiliares de cozinha", P, 4, 1850, 20, 200, "Pré-preparo, porcionamento, embalagem e limpeza."),
        ("Atendentes / barista / caixa", C, 6, 1870, 20, 200, "Atendimento, bebidas, caixa e pedidos on-line."),
        ("Vendedores / estoquista / limpeza / entregadores", C, 8, 1880, 20, 200, "Loja, estoque, limpeza e entregas."),
    ])
    CIF, DA, DC, DF = ("Custo indireto de produção", "Despesa administrativa", "Despesa comercial",
                       "Despesa financeira")
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Aluguel da área de produção (cozinha)", CIF, 4200), ("Aluguel da loja e escritório", DA, 2800),
        ("Pró-labore", DA, 6000), ("Energia elétrica da cozinha", CIF, 3600), ("Energia elétrica da loja", DA, 900),
        ("Água", CIF, 2000), ("Honorários contábeis", DA, 1800), ("Manutenção de equipamentos", CIF, 1300),
        ("Seguros", DA, 1200), ("Serviços de terceiros", DA, 750), ("Marketing e propaganda (fixo)", DC, 600),
        ("Materiais de limpeza", CIF, 600), ("Telefone e internet", DA, 500), ("Material de escritório", DA, 450),
        ("Tarifas bancárias", DF, 250), ("Mensalidade de máquinas de cartão", DC, 150),
    ])

    conn.execute("""UPDATE projeto SET
        autores='Alessandra Ribeiro Gonçalves; Paulo Henrique Batista; Camila Gomes de Souza; Cauan Ribeiro de Souza',
        turma='Administração - Empreendedorismo (exemplo)', professor='', data_entrega='',
        tam_desc='Moradores de Teófilo Otoni e região que consomem refeições fora de casa (estimativa ilustrativa)',
        tam_clientes=60000,
        sam_desc='Trabalhadores, estudantes e praticantes de atividade física no raio de entrega de 10 km',
        sam_clientes=18000, som_pct=12, gasto_anual_cliente=1600,
        mercado_fontes='Estimativas ilustrativas para fins didáticos; substitua por dados do IBGE, Sebrae e pesquisa de campo.',
        mkt_produto='Refeições saudáveis com informação nutricional, planos semanais/mensais, suplementos e moda fitness em um só lugar.',
        mkt_preco='Preço intermediário: marmitas de R$ 24,90 a R$ 54,90; descontos progressivos em assinaturas.',
        mkt_praca='Loja física com consumo no local, retirada, delivery próprio até 10 km e pedidos via WhatsApp e aplicativo.',
        mkt_promocao='Instagram e anúncios segmentados, parcerias com academias, programa de fidelidade e degustações.',
        loc_endereco='Região central de Teófilo Otoni - MG (próximo a academias e centros comerciais)',
        loc_justificativa='Alto fluxo de trabalhadores e estudantes, proximidade de academias e facilidade para entregas.',
        loc_area_m2=180, loc_fluxo='Fluxo intenso no almoço (11h-14h) e no fim da tarde (17h-20h).',
        op_layout='Cozinha industrial com áreas separadas (pré-preparo, cocção, porcionamento, expedição), salão com 48 lugares e área de loja.',
        op_processos='Compras semanais -> recebimento e conferência -> pré-preparo -> cocção -> porcionamento e etiquetagem -> venda/entrega -> pós-venda.',
        op_capacidade_txt='Capacidade estimada pela equipe de cozinha e pelos equipamentos de cocção e refrigeração.',
        op_horario='Segunda a sábado, das 7h às 21h.',
        rateio_criterio='Custo direto', margem_desejada_pct=15
        WHERE id=1""")

    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Mariana", "Conforto e conveniência; Economia de tempo; Alimentação equilibrada; Transparência nutricional; Rapidez",
         "manter uma alimentação equilibrada mesmo com a rotina corrida",
         "refeições prontas com informações nutricionais, planos semanais e entrega no horário do almoço",
         "restaurantes por quilo e marmitarias tradicionais sem informação nutricional",
         "praticidade, sabor e controle nutricional por um preço acessível"),
        ("Rafael", "Desempenho; Confiança; Integração de produtos; Personalização",
         "atingir sua meta de proteína e comprar suplementos confiáveis",
         "marmitas hiperproteicas com gramatura informada e loja de suplementos originais",
         "lojas que vendem só suplementos ou só refeições",
         "refeição pós-treino, suplementos e roupas de treino em um só lugar"),
    ])
    ins("INSERT INTO capacidade (recurso, produto_ou_grupo, capacidade_mensal, observacao) VALUES (?,?,?,?)", [
        ("Cozinha e bar (equipe + equipamentos)", "Fabricação Própria", 7500, "≈ 290 itens por dia em 26 dias"),
        ("Equipe de atendimento e serviços", "Serviço", 2600, "Assinaturas, eventos e entregas"),
        ("Entregadores (2 motos)", "Delivery de 3 a 10 km", 1400, "≈ 54 entregas longas por dia"),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Distribuidora de carnes (exemplo)", "Frango, carne bovina, peixe", 28, 2, "Boleto 28 dias; pedido mínimo R$ 800", 1),
        ("Ceasa / hortifrúti local (exemplo)", "Verduras, legumes e frutas", 7, 1, "Pagamento semanal", 1),
        ("Atacadista de alimentos (exemplo)", "Arroz, feijão, grãos, óleos", 30, 3, "Boleto 30 dias", 0),
        ("Fornecedor de embalagens (exemplo)", "Marmitas, copos, etiquetas", 30, 7, "Boleto 30/60 dias", 1),
        ("Distribuidor de suplementos (exemplo)", "Whey, creatina, barras", 45, 5, "Nota fiscal com rastreabilidade", 0),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Alta no preço das proteínas", 4, 4, "Contratos de compra programada e cardápio com proteínas alternativas", "Sócio operacional"),
        ("Vendas abaixo do previsto nos primeiros meses", 3, 5, "Campanha de lançamento, parcerias e reserva de caixa", "Sócia comercial"),
        ("Interdição sanitária", 1, 5, "Boas práticas, checklists diários e consultoria técnica", "Gerente operacional"),
        ("Rotatividade de funcionários", 3, 3, "Treinamento, plano de carreira e bom clima de trabalho", "Sócia-administradora"),
        ("Atraso nas entregas em horário de pico", 3, 3, "Pedidos agendados e terceira moto em períodos de pico", "Gerente operacional"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Trabalhadores do centro pagariam R$ 29,90 por marmita com informação nutricional",
         "Venda-teste de 100 marmitas em 2 semanas via Instagram", "Unidades vendidas e recompra",
         "Vender 80 unidades e 30% de recompra", "", "Planejado"),
        ("Alunos de academia assinariam um plano mensal", "Oferta de plano em 3 academias parceiras",
         "Número de assinaturas", "15 assinaturas no primeiro mês", "", "Planejado"),
    ])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Abertura da empresa e registro de marca", 1, 1, "Sócia-administradora", 12000),
        ("Reforma e adequação sanitária", 1, 3, "Sócio operacional", 75000),
        ("Compra e instalação de equipamentos", 2, 2, "Sócio operacional", 330910),
        ("Contratação e treinamento da equipe", 3, 1, "Sócia-administradora", 10000),
        ("Campanha de pré-lançamento", 3, 1, "Sócia comercial", 25000),
        ("Inauguração", 4, 1, "Todos os sócios", 0),
    ])


# ==============================================================================
# SEGUNDO ESTUDO DE CASO: CAFÉ MUCURI (cafeteria e padaria artesanal)
# ==============================================================================
def popular_modelo_cafe(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='Café Mucuri', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Alimentos e bebidas; Padaria e confeitaria; Bares e restaurantes',
            cidade_uf='Teófilo Otoni - MG', contatos='@cafemucuri',
            aliquota_imposto_pct=8.0, crescimento_mensal_ano1_pct=1.5, crescimento_anual_pct=7.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=12.0, pme_dias=7, reserva_caixa_extra=25000,
            recuperar_giro=1, cen_otim_receita_pct=10, cen_otim_custo_pct=5,
            cen_pess_receita_pct=15, cen_pess_custo_pct=5,
            autores='Integrante 1; Integrante 2; Integrante 3', turma='Exemplo didático - setor de alimentação',
            tam_desc='Moradores de Teófilo Otoni que consomem café e lanches fora de casa (estimativa ilustrativa)',
            tam_clientes=45000, sam_desc='Público do centro e bairros vizinhos num raio de 2 km', sam_clientes=9000,
            som_pct=10, gasto_anual_cliente=900,
            mercado_fontes='Estimativas ilustrativas; substitua por IBGE, Sebrae, ABIC e pesquisa de campo.',
            mkt_produto='Cafés especiais de produtores do Vale do Mucuri e do Jequitinhonha, pães de fermentação natural e quitandas mineiras.',
            mkt_preco='Preço intermediário; combos de café da manhã e cartão fidelidade (10º café grátis).',
            mkt_praca='Loja de esquina no centro, retirada no balcão e encomendas pelo WhatsApp.',
            mkt_promocao='Instagram com a história dos produtores, degustações, parcerias com coworkings e eventos culturais.',
            loc_endereco='Centro de Teófilo Otoni - MG, próximo a escritórios, bancos e escolas',
            loc_justificativa='Fluxo de pedestres pela manhã e no fim da tarde, poucos cafés especiais na região.',
            loc_area_m2=90, loc_fluxo='Picos das 7h às 9h e das 15h às 18h.',
            op_layout='Padaria compacta nos fundos (forno de lastro, masseira, bancada), balcão de café e salão com 28 lugares.',
            op_processos='Fermentação noturna -> forneamento às 5h -> abertura às 6h30 -> produção de quitandas -> vendas -> fechamento e limpeza.',
            op_horario='Segunda a sábado, das 6h30 às 19h.', rateio_criterio='Receita', margem_desejada_pct=12
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Administração e finanças", "", "Gestão financeira, compras e pessoas.", 70000.0),
        ("Sócio(a) 2", "Produção e barista", "", "Panificação artesanal e preparo de cafés especiais.", 50000.0),
    ])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Trabalhadores do centro", "Pessoas que tomam café da manhã ou lanche perto do trabalho.", 9, 7, 9, 6, 9),
        ("Estudantes e professores", "Público das escolas e faculdades próximas.", 8, 6, 8, 7, 8),
        ("Apreciadores de cafés especiais", "Consumidores que valorizam origem e método de preparo.", 5, 9, 8, 8, 10),
        ("Famílias no fim de semana", "Compra de pães e quitandas para casa.", 7, 7, 9, 5, 7),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Luciana", 34, "Bancária", "Casada", 5200.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Chega ao centro às 7h30 e quer um café bem feito antes do expediente.",
         "Um bom café muda o meu dia, mas não tenho tempo de esperar.",
         "Usa Instagram, paga por aproximação, valoriza produtos locais.",
         "Rapidez; qualidade constante; ambiente agradável.", "Filas; café requentado; pouco lugar para sentar.",
         "O café especial cabe no meu orçamento diário?", "Pede pelo WhatsApp e retira pronto no balcão em 3 minutos."),
    ])
    ins("""INSERT INTO jornada_cliente (persona_nome, macro_fase, etapa, acoes_cliente, pensamentos,
           pontos_contato, estado_emocional, oportunidades) VALUES ('Luciana',?,?,?,?,?,?,?)""", [
        ("Antes", "Descoberta", "Vê post de colega no Instagram.", "Que lugar é esse?", "Instagram", "Neutro", "Conteúdo sobre os produtores."),
        ("Antes", "Consideração", "Compara com a padaria do bairro.", "Será caro?", "Cardápio na vitrine", "Neutro", "Combo café + pão de queijo."),
        ("Antes", "Decisão", "Passa na porta e sente o cheiro do pão.", "Vou experimentar.", "Fachada e vitrine", "Satisfeito", "Degustação na calçada."),
        ("Durante", "Compra e Pagamento", "Pede no balcão e paga por aproximação.", "Foi rápido.", "Caixa", "Satisfeito", "Pedido antecipado pelo WhatsApp."),
        ("Durante", "Entrega", "Recebe o café em xícara ou copo para viagem.", "Está quentinho.", "Balcão", "Satisfeito", "Copo reutilizável com desconto."),
        ("Durante", "Uso", "Consome no salão ou a caminho do trabalho.", "Que sabor!", "Produto", "Entusiasmado", "Cartão com a origem do grão."),
        ("Depois", "Avaliação", "Posta foto nos stories.", "Vou indicar.", "Instagram", "Entusiasmado", "Repostar clientes e marcar produtores."),
        ("Depois", "Suporte", "Pergunta sobre encomendas para festa.", "Fazem para 30 pessoas?", "WhatsApp", "Neutro", "Catálogo de encomendas."),
        ("Depois", "Lealdade e advocacia", "Usa o cartão fidelidade.", "Já é minha parada diária.", "Cartão fidelidade", "Entusiasmado", "Assinatura mensal de café."),
    ])
    concs = ["Padaria tradicional do centro", "Rede de cafeteria"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Qualidade do café": (9, 5, 7), "Pães e quitandas": (9, 8, 5), "Preço": (6, 9, 5),
             "Atendimento": (8, 6, 7), "Ambiente": (9, 5, 8), "Rapidez": (7, 8, 7), "Presença digital": (8, 3, 9)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Operação", "Panificação artesanal", "Força", "Alto", "Fermentação natural e receitas mineiras.", "Padronizar fichas técnicas.", 1),
        ("Interno", "Marketing", "Café de origem regional", "Força", "Alto", "Parceria direta com produtores locais.", "Contar a história dos produtores.", 1),
        ("Interno", "Finanças", "Capital limitado", "Fraqueza", "Alto", "Pouca reserva para os primeiros meses.", "Controle semanal de caixa.", 1),
        ("Interno", "Operação", "Espaço pequeno", "Fraqueza", "Baixo", "Salão de 28 lugares.", "Estimular o para viagem.", 0),
        ("Externo", "Social", "Valorização do consumo local", "Oportunidade", "Alto", "Clientes buscam produtos da região.", "Selo de origem nos produtos.", 1),
        ("Externo", "Econômico", "Alta do preço do café", "Ameaça", "Alto", "Café arábica subiu fortemente.", "Contratos com produtores e revisão de preços.", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,10)""", [
        ("Forno de lastro elétrico", 1, 22000, 10), ("Masseira espiral 15 kg", 1, 9500, 10),
        ("Máquina de café expresso 2 grupos", 1, 28000, 10), ("Moedor de café profissional", 2, 4500, 10),
        ("Vitrine refrigerada", 1, 7800, 10), ("Geladeira e freezer comerciais", 2, 5200, 10),
        ("Mesas e cadeiras", 7, 1400, 10), ("Balcão e mobiliário", 1, 15000, 10),
        ("Computador, PDV e impressora", 1, 6500, 5), ("Utensílios de padaria e café", 1, 6000, 5),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Reforma e adequação sanitária", 30000), ("Legalização e alvarás", 4000),
        ("Identidade visual e fachada", 6000), ("Marketing de inauguração", 5000), ("Treinamento da equipe", 2500),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 60000, 14.0, 36)])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 75.0), ("Recebimento", 30, 25.0), ("Pagamento", 0, 40.0), ("Pagamento", 30, 60.0)])
    ins("INSERT INTO materiais (nome, categoria, unidade, valor_unitario, qtd_estoque_inicial) VALUES (?,?,?,?,?)", [
        ("Café especial em grão", "Matéria-prima", "kg", 95.00, 15), ("Leite integral", "Matéria-prima", "L", 5.50, 60),
        ("Farinha de trigo", "Matéria-prima", "kg", 5.00, 100), ("Polvilho azedo", "Matéria-prima", "kg", 12.00, 30),
        ("Queijo minas meia cura", "Matéria-prima", "kg", 48.00, 10), ("Ovos", "Matéria-prima", "dúzia", 12.00, 20),
        ("Manteiga", "Matéria-prima", "kg", 45.00, 8), ("Açúcar", "Matéria-prima", "kg", 4.50, 30),
        ("Copo para viagem 200 ml", "Embalagem", "unidade", 0.40, 1000), ("Saco de papel kraft", "Embalagem", "unidade", 0.25, 1000),
    ])
    F, R = "Fabricação Própria", "Revenda"
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial)
           VALUES (?,?,?,?,?,?)""", [
        (F, "Café expresso", 6.50, 0, 2200, 0), (F, "Cappuccino", 11.00, 0, 900, 0),
        (F, "Café coado (método)", 12.00, 0, 450, 0), (F, "Pão de queijo (unidade)", 4.50, 0, 3000, 0),
        (F, "Pão de fermentação natural (500 g)", 22.00, 0, 520, 0), (F, "Bolo caseiro (fatia)", 9.00, 2.40, 800, 0),
        (F, "Biscoito de polvilho (pacote)", 8.00, 2.10, 400, 0),
        (R, "Café em grão 250 g (revenda)", 42.00, 26.00, 120, 30), (R, "Suco engarrafado", 9.00, 4.50, 300, 60),
    ])
    ficha = {
        "Café expresso": [("Café especial em grão", 0.009), ("Copo para viagem 200 ml", 0.4)],
        "Cappuccino": [("Café especial em grão", 0.009), ("Leite integral", 0.15), ("Copo para viagem 200 ml", 0.5)],
        "Café coado (método)": [("Café especial em grão", 0.018), ("Copo para viagem 200 ml", 0.3)],
        "Pão de queijo (unidade)": [("Polvilho azedo", 0.025), ("Queijo minas meia cura", 0.012), ("Leite integral", 0.01),
                                    ("Ovos", 0.008), ("Saco de papel kraft", 0.2)],
        "Pão de fermentação natural (500 g)": [("Farinha de trigo", 0.32), ("Saco de papel kraft", 1)],
    }
    ins("INSERT INTO ficha_tecnica (produto, material, quantidade) VALUES (?,?,?)",
        [(p, m, q) for p, itens in ficha.items() for m, q in itens])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [
        ("Taxas de cartões", 2.0), ("Perdas e quebras", 1.5), ("Propaganda variável", 0.5)])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Padeiro(a)", P, 1, 2300, 20, 200, "Massas, fermentação e forneamento."),
        ("Auxiliar de produção", P, 1, 1700, 20, 200, "Quitandas, limpeza e apoio."),
        ("Barista / atendente", C, 2, 1750, 20, 200, "Preparo de cafés, atendimento e caixa."),
    ])
    CIF, DA, DC, DF = "Custo indireto de produção", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Aluguel", DA, 3200), ("Energia elétrica (forno e máquinas)", CIF, 1500), ("Água e gás", CIF, 450),
        ("Pró-labore", DA, 4000), ("Contador", DA, 700), ("Internet e telefone", DA, 180),
        ("Manutenção de equipamentos", CIF, 300), ("Marketing fixo", DC, 400), ("Tarifas bancárias", DF, 90),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Luciana", "Rapidez; Qualidade; Origem regional; Ambiente acolhedor",
         "começar o dia com um café de qualidade sem perder tempo",
         "cafés especiais da região, pão de queijo saindo do forno e pedido antecipado pelo WhatsApp",
         "padarias tradicionais com café comum e redes sem identidade local",
         "sabor mineiro autêntico com rapidez e preço justo"),
    ])
    ins("INSERT INTO capacidade (recurso, produto_ou_grupo, capacidade_mensal, observacao) VALUES (?,?,?,?)", [
        ("Máquina de expresso 2 grupos", "Café expresso", 6000, "≈ 230 bebidas por dia"),
        ("Forno de lastro", "Fabricação Própria", 7000, "Fornadas das 5h às 11h"),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Produtor de café do Vale (exemplo)", "Café especial em grão", 30, 7, "Contrato trimestral de preço", 1),
        ("Laticínio regional (exemplo)", "Leite, queijo e manteiga", 14, 1, "Entrega 3x por semana", 1),
        ("Atacadista (exemplo)", "Farinha, açúcar, polvilho", 28, 3, "Boleto 28 dias", 0),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Alta do preço do café", 4, 4, "Contrato com produtor e ajuste trimestral de preços", "Sócio(a) 1"),
        ("Quebra do forno", 2, 5, "Manutenção preventiva e parceria com padaria vizinha", "Sócio(a) 2"),
        ("Vendas abaixo do previsto", 3, 4, "Degustações, combos e reserva de caixa", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Clientes do centro pagariam R$ 12 por café coado especial", "Barraca de degustação em feira por 2 sábados",
         "Cafés vendidos", "Vender 150 cafés", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Abertura da empresa e alvarás", 1, 1, "Sócio(a) 1", 4000), ("Reforma", 1, 2, "Sócio(a) 2", 30000),
        ("Equipamentos", 2, 1, "Sócio(a) 2", 120000), ("Contratação e treinamento", 3, 1, "Sócio(a) 1", 2500),
        ("Inauguração", 3, 1, "Todos", 5000),
    ])


# ==============================================================================
# TERCEIRO ESTUDO DE CASO: CONECTA VALE (aplicativo com planos de assinatura)
# Negócio fictício, para ensinar receita recorrente, churn, CAC/LTV e custos de tecnologia.
# ==============================================================================
def popular_modelo_assinatura(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='Conecta Vale', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Tecnologia e aplicativos; Turismo, cultura e lazer', cidade_uf='Teófilo Otoni - MG',
            contatos='@conectavale', tipo_negocio='Assinatura (receita recorrente)',
            horas_mes_pessoa=168, utilizacao_pct=70, anexo_simples='Anexo III - Serviços',
            aliquota_imposto_pct=6.0, crescimento_mensal_ano1_pct=3.0, crescimento_anual_pct=10.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=14.0, pme_dias=0, reserva_caixa_extra=50000,
            recuperar_giro=1, cen_otim_receita_pct=15, cen_otim_custo_pct=5, cen_pess_receita_pct=25,
            cen_pess_custo_pct=10, autores='Integrante 1; Integrante 2', turma='Exemplo didático - negócio por assinatura',
            tam_desc='Empresas de comércio, alimentação, cultura e turismo do Vale do Mucuri (estimativa ilustrativa)',
            tam_clientes=6000, sam_desc='Empresas de Teófilo Otoni com presença digital', sam_clientes=2000, som_pct=8,
            gasto_anual_cliente=2400,
            mercado_fontes='Estimativas ilustrativas; substitua por dados da Junta Comercial, IBGE e pesquisa com empresas.',
            mkt_produto='Aplicativo que reúne eventos, cultura, gastronomia e lazer da região. Empresas pagam planos para aparecer em  destaque; usuários usam grátis ou assinam o Premium com descontos.',
            mkt_preco='Três planos para empresas (mensal, trimestral e anual, com desconto para quem paga adiantado) e um plano  Premium barato para usuários. Primeiro mês grátis para empresas.',
            mkt_praca='Aplicativo nas lojas Google Play e App Store; venda consultiva às empresas, de porta em porta e pelo WhatsApp.',
            mkt_promocao='Parcerias com secretarias de cultura e turismo, influenciadores locais e semana de descontos mensal.',
            loc_endereco='Coworking no centro de Teófilo Otoni - MG (trabalho principalmente remoto)',
            loc_justificativa='Custo baixo e perto das empresas clientes para reuniões presenciais.',
            op_processos='Cadastro da empresa -> criação do perfil e das publicações -> aprovação -> divulgação no app ->  relatório mensal de visualizações -> renovação.',
            op_horario='Atendimento às empresas em horário comercial; aplicativo disponível 24 horas.',
            rateio_criterio='Receita', margem_desejada_pct=15
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Produto e tecnologia", "", "Desenvolvimento de aplicativos e gestão do produto.", 50000.0),
        ("Sócio(a) 2", "Comercial e marketing", "", "Vendas para empresas, parcerias e conteúdo.", 50000.0),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 21200, 14.0, 36)])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Bares, restaurantes e cafés", "Querem atrair clientes em dias de pouco movimento.", 9, 7, 8, 6, 9),
        ("Produtores de eventos e cultura", "Precisam divulgar shows, peças e festas.", 6, 6, 7, 7, 10),
        ("Hotéis e pousadas", "Buscam turistas e visitantes de negócios.", 4, 8, 8, 5, 7),
        ("Moradores e turistas (usuários)", "Querem saber o que fazer na cidade.", 10, 3, 8, 6, 8),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Marcos", 41, "Dono de restaurante", "Casado", 9000.0, "Ensino Médio", "Teófilo Otoni - MG",
         "Tem um restaurante com música ao vivo e faz a própria divulgação no Instagram.",
         "Eu pago impulsionamento, mas não sei se vem cliente por causa disso.",
         "Usa WhatsApp o dia todo; pouco tempo para marketing.", "Mesas cheias às quartas e quintas.",
         "Não sabe medir o retorno do marketing; caixa apertado.", "Vale a pena pagar todo mês?",
         "Ver no relatório quantas pessoas abriram a oferta e usaram o cupom."),
    ])
    concs = ["Plataforma nacional de eventos", "Perfis locais no Instagram"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Conteúdo local": (9, 4, 8), "Preço para pequenas empresas": (8, 5, 9), "Relatórios de resultado": (8, 7, 2),
             "Alcance de público": (5, 9, 7), "Facilidade de uso": (8, 8, 7)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Operação", "Equipe técnica própria", "Força", "Alto", "Sócio desenvolvedor e equipe enxuta.", "Lançar melhorias a cada mês.", 1),
        ("Interno", "Finanças", "Receita depende de renovações", "Fraqueza", "Alto", "Churn alto derruba a receita.", "Programa de sucesso do cliente.", 1),
        ("Externo", "Social", "Crescimento do turismo regional", "Oportunidade", "Alto", "Mais visitantes buscando o que fazer.", "Parcerias com o poder público.", 1),
        ("Externo", "Tecnológico", "Mudança nas taxas das lojas de aplicativos", "Ameaça", "Médio", "Comissões e regras mudam com frequência.", "Cobrar empresas fora da loja (boleto/Pix).", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,?)""", [
        ("Notebooks para desenvolvimento", 4, 6500, 5, 10), ("Celulares para testes (Android e iPhone)", 2, 3000, 3, 10),
        ("Mobiliário e equipamentos de escritório", 1, 3500, 10, 10),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Desenvolvimento da primeira versão do app (MVP)", 25000), ("Legalização e contratos", 2500),
        ("Registro de marca", 1500), ("Marketing de lançamento", 6000), ("Contas de desenvolvedor nas lojas", 700),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 100.0), ("Pagamento", 0, 100.0)])
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial, horas_unit)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Serviço", "Cobertura de evento patrocinado", 900, 120, 2, 0, 16),
        ("Serviço", "Produção de conteúdo avulso (post e vídeo)", 180, 10, 8, 0, 3),
    ])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [("Comissões de vendas", 5.0)])
    P, A, C = "Produção / operação", "Administrativo", "Comercial / vendas"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Desenvolvedor(a)", P, 1, 5500, 28, 300, "Manter e evoluir o aplicativo, corrigir falhas."),
        ("Atendimento e conteúdo", P, 1, 2300, 28, 300, "Montar publicações, cobrir eventos, suporte às empresas."),
        ("Vendedor(a) externo(a)", C, 1, 2200, 28, 300, "Visitar empresas e fechar planos."),
    ])
    TEC, DA, DC, DF = "Custo de tecnologia / infraestrutura", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Servidores em nuvem (base fixa)", TEC, 700), ("Ferramentas e licenças de software", TEC, 450),
        ("Contas de desenvolvedor (rateio mensal)", TEC, 60), ("Coworking", DA, 900), ("Pró-labore", DA, 4000),
        ("Contador", DA, 500), ("Internet e telefone", DA, 200), ("Marketing digital (anúncios)", DC, 2000),
        ("Tarifas bancárias", DF, 80),
    ])
    ins("""INSERT INTO assinaturas (plano, publico, preco_ciclo, ciclo_meses, novos_mes1, crescimento_novos_pct,
           churn_mensal_pct, custo_entrega_mensal, taxa_cobranca_pct) VALUES (?,?,?,?,?,?,?,?,?)""", [
        ("Parceiro Essencial (mensal)", "Empresas (B2B)", 149, 1, 8, 4, 6, 10, 3.5),
        ("Parceiro Destaque (trimestral)", "Empresas (B2B)", 600, 3, 5, 4, 4, 14, 3.5),
        ("Parceiro Premium (anual)", "Empresas (B2B)", 3000, 12, 2, 4, 2, 18, 3.5),
        ("Usuário Premium (no app)", "Consumidores (B2C)", 6.90, 1, 50, 5, 8, 0.6, 15),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Marcos", "Clientes novos; Resultado medido; Preço acessível",
         "lotar o restaurante nos dias fracos sem perder tempo com marketing",
         "divulgação no aplicativo da cidade, cupons e relatório mensal de resultados",
         "impulsionamentos genéricos em redes sociais", "divulgação local com resultado que dá para medir"),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Churn maior que o previsto", 4, 5, "Acompanhamento mensal das empresas e relatórios de resultado", "Sócio(a) 2"),
        ("Falha ou queda do aplicativo", 2, 4, "Monitoramento, cópias de segurança e plano de contingência", "Sócio(a) 1"),
        ("Aumento do custo de nuvem com o crescimento", 3, 3, "Otimizar o sistema e negociar créditos de nuvem", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Restaurantes pagariam R$ 149/mês por divulgação no app", "Oferecer 30 dias grátis a 20 restaurantes",
         "Conversão para pago", "Pelo menos 6 assinam após o teste", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Desenvolvimento do MVP", 1, 3, "Sócio(a) 1", 25000), ("Teste com empresas piloto", 3, 1, "Sócio(a) 2", 0),
        ("Publicação nas lojas", 4, 1, "Sócio(a) 1", 700), ("Lançamento", 4, 1, "Todos", 6000),
    ])


# ==============================================================================
# QUARTO EXEMPLO: TECVALE SOLUÇÕES (SERVIÇOS DE TI POR HORA E POR PROJETO) - fictício
# ==============================================================================
def popular_modelo_servico(conn) -> None:
    ins = conn.executemany
    conn.execute("""
        UPDATE projeto SET nome_negocio='TecVale Soluções', objetivo='Criar um novo negócio',
            enquadramento='Simples Nacional', periodo_anos=3, porte='Microempresa (ME)',
            setor_atuacao='Tecnologia e aplicativos; Serviço', cidade_uf='Teófilo Otoni - MG',
            contatos='@tecvalesolucoes', tipo_negocio='Serviços (hora ou projeto)',
            horas_mes_pessoa=168, utilizacao_pct=70, anexo_simples='Anexo III - Serviços',
            aliquota_imposto_pct=6.0, crescimento_mensal_ano1_pct=2.0, crescimento_anual_pct=10.0,
            inflacao_custos_anual_pct=4.5, tma_anual_pct=14.0, pme_dias=0, reserva_caixa_extra=25000,
            recuperar_giro=1, cen_otim_receita_pct=15, cen_otim_custo_pct=5, cen_pess_receita_pct=20,
            cen_pess_custo_pct=10, autores='Integrante 1; Integrante 2', turma='Exemplo didático - empresa de serviços',
            tam_desc='Micro e pequenas empresas do Vale do Mucuri (estimativa ilustrativa)',
            tam_clientes=9000, sam_desc='Micro e pequenas empresas de Teófilo Otoni com computador e presença digital',
            sam_clientes=3500, som_pct=5, gasto_anual_cliente=2400,
            mercado_fontes='Estimativas ilustrativas; substitua por dados do Sebrae (Data MPE), da Junta Comercial e de pesquisa com empresas.',
            mkt_produto='Suporte de informática, manutenção de computadores, criação de sites e lojas virtuais e treinamento para pequenas empresas. O carro-chefe é o contrato mensal de manutenção, que gera receita previsível.',
            mkt_preco='Hora técnica avulsa, pacotes fechados por projeto (sites e lojas virtuais) e contrato mensal com horas incluídas, mais barato que a hora avulsa.',
            mkt_praca='Atendimento no cliente, na bancada da empresa e remoto. Venda direta pelo WhatsApp, indicação e parcerias com contadores.',
            mkt_promocao='Diagnóstico gratuito de segurança e backup, conteúdo educativo no Instagram e programa de indicação.',
            loc_endereco='Sala comercial no centro de Teófilo Otoni - MG', loc_area_m2=45,
            loc_justificativa='Perto dos clientes do comércio e com espaço para a bancada de manutenção.',
            op_processos='Chamado pelo WhatsApp -> triagem e orçamento -> atendimento remoto ou no local -> registro das horas -> faturamento mensal.',
            op_horario='Segunda a sexta, 8h às 18h; plantão para clientes com contrato.',
            rateio_criterio='Horas', margem_desejada_pct=15
        WHERE id=1""")
    ins("INSERT INTO socios (nome, funcao, contato, perfil, capital_integralizado) VALUES (?,?,?,?,?)", [
        ("Sócio(a) 1", "Técnico(a) responsável", "", "Técnico em informática, 6 anos de experiência em suporte.", 30000.0),
        ("Sócio(a) 2", "Comercial e projetos web", "", "Desenvolvimento de sites e atendimento a clientes.", 30000.0),
    ])
    ins("INSERT INTO emprestimos (descricao, valor, taxa_anual_pct, prazo_meses) VALUES (?,?,?,?)", [
        ("Pronampe", 25000, 14.0, 36)])
    ins("""INSERT INTO segmentos (nome, descricao, nota_tamanho, nota_lucro, nota_longevidade,
           nota_concorrencia, nota_aderencia) VALUES (?,?,?,?,?,?,?)""", [
        ("Comércio varejista", "Lojas com sistema de vendas, impressoras e redes que não podem parar.", 9, 6, 8, 5, 9),
        ("Escritórios e clínicas", "Contabilidade, advocacia e saúde: dados sensíveis e muitos computadores.", 7, 8, 9, 6, 8),
        ("Novos negócios sem site", "Precisam de presença digital rápida e barata.", 8, 5, 6, 4, 7),
    ])
    ins("""INSERT INTO personas (nome, idade, profissao, estado_civil, renda_mensal, escolaridade, cidade,
           descricao, citacao, habitos, ganhos, desafios, duvidas_objecoes, experiencia_ideal)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [
        ("Cláudia", 47, "Dona de loja de roupas", "Casada", 8000.0, "Ensino Superior", "Teófilo Otoni - MG",
         "Tem uma loja com duas filiais e depende do sistema de vendas e da maquininha.",
         "Quando o computador para, eu paro de vender.", "Resolve tudo pelo WhatsApp; chama o sobrinho para consertar o PC.",
         "Sistema funcionando e site que traga clientes.", "Atendimento demorado e sem garantia; medo de perder dados.",
         "Contrato mensal compensa?", "Um técnico que atende no mesmo dia e avisa antes de o problema acontecer."),
    ])
    concs = ["Técnicos autônomos", "Loja de informática do centro", "Agência digital da capital"]
    ins("INSERT INTO concorrentes (nome) VALUES (?)", [(c,) for c in concs])
    notas = {"Rapidez no atendimento": (9, 6, 5, 4), "Preço": (7, 9, 6, 3), "Garantia do serviço": (8, 3, 6, 7),
             "Conhecimento técnico": (8, 5, 6, 9), "Atendimento presencial": (9, 8, 8, 2), "Sites e lojas virtuais": (7, 3, 2, 9)}
    ins("INSERT INTO concorrencia_notas (atributo, empresa, nota) VALUES (?,?,?)",
        [(a, e, v) for a, vals in notas.items() for e, v in zip(["Nossa Empresa"] + concs, vals)])
    ins("""INSERT INTO swot (ambiente, dimensao, fator, natureza, impacto, diagnostico, plano_acao, priorizado)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Interno", "Pessoas", "Sócios com experiência técnica e em sites", "Força", "Alto", "Atende do computador ao site.", "Pacotes que combinam os dois serviços.", 1),
        ("Interno", "Operação", "Receita depende das horas da equipe", "Fraqueza", "Alto", "Faturamento limitado pela capacidade.", "Contratos mensais e serviços remotos.", 1),
        ("Externo", "Tecnológico", "Digitalização das pequenas empresas", "Oportunidade", "Alto", "Pix, nota fiscal eletrônica e vendas on-line exigem estrutura.", "Pacote \"empresa digital\".", 1),
        ("Externo", "Legal", "LGPD exige cuidado com dados de clientes", "Oportunidade", "Médio", "Clínicas e escritórios precisam se adequar.", "Diagnóstico de segurança e backup.", 1),
        ("Externo", "Mercado", "Técnicos informais cobram muito pouco", "Ameaça", "Médio", "Concorrência por preço.", "Vender garantia, prazo e contrato.", 1),
    ])
    ins("""INSERT INTO inv_fixos (descricao, quantidade, valor_unitario, vida_util_anos, valor_residual_pct)
           VALUES (?,?,?,?,?)""", [
        ("Notebooks para a equipe", 3, 4500, 5, 10), ("Bancada, ferramentas e estação de solda", 1, 4500, 5, 10),
        ("Mobiliário da sala", 1, 4000, 10, 10), ("Moto para atendimento externo", 1, 14000, 5, 40),
    ])
    ins("INSERT INTO inv_pre (descricao, valor) VALUES (?,?)", [
        ("Legalização e contrato social", 2500), ("Licenças de software (primeiro ano)", 2500),
        ("Reforma e adequação da sala", 3000), ("Marketing de lançamento", 3000),
    ])
    ins("INSERT INTO capital_giro_prazos (tipo, prazo_dias, percentual) VALUES (?,?,?)", [
        ("Recebimento", 0, 50.0), ("Recebimento", 30, 50.0), ("Pagamento", 0, 100.0)])
    ins("""INSERT INTO catalogo (categoria, descricao, preco_venda, custo_aquisicao, qtd_mensal, qtd_estoque_inicial, horas_unit)
           VALUES (?,?,?,?,?,?,?)""", [
        ("Serviço", "Contrato mensal de manutenção (6 h incluídas)", 690, 20, 14, 0, 6),
        ("Serviço", "Hora técnica avulsa", 130, 0, 40, 0, 1),
        ("Serviço", "Formatação e configuração de computador", 190, 15, 16, 0, 2.5),
        ("Serviço", "Site institucional", 2900, 250, 2, 0, 40),
        ("Serviço", "Loja virtual", 5800, 400, 0.5, 0, 80),
        ("Serviço", "Treinamento em ferramentas digitais (turma)", 1300, 80, 1, 0, 12),
    ])
    ins("INSERT INTO custos_variaveis (descricao, percentual) VALUES (?,?)", [("Taxas de cartão e Pix", 2.0)])
    P = "Produção / operação"
    ins("""INSERT INTO pessoal (cargo, area, num_empregados, salario_mensal, encargos_pct, beneficios_unit, atribuicoes, mes_inicio)
           VALUES (?,?,?,?,?,?,?,?)""", [
        ("Técnico(a) de suporte", P, 2, 2400, 32, 350, "Atendimento remoto e presencial, manutenção de computadores e redes.", 1),
        ("Desenvolvedor(a) web", P, 1, 3500, 32, 350, "Criação de sites e lojas virtuais.", 1),
        ("Técnico(a) de suporte júnior", P, 1, 2000, 32, 350, "Contratação prevista para quando a ocupação passar de 90%.", 25),
    ])
    TEC, DA, DC, DF = "Custo de tecnologia / infraestrutura", "Despesa administrativa", "Despesa comercial", "Despesa financeira"
    ins("INSERT INTO custos_fixos (descricao, classificacao, valor_mensal) VALUES (?,?,?)", [
        ("Softwares de acesso remoto e gestão de chamados", TEC, 350), ("Hospedagem e domínios dos clientes (base)", TEC, 150),
        ("Combustível e manutenção da moto", TEC, 450), ("Aluguel da sala", DA, 1500), ("Pró-labore", DA, 5000),
        ("Contador", DA, 500), ("Internet e telefone", DA, 300), ("Energia elétrica", DA, 300),
        ("Anúncios e conteúdo", DC, 700), ("Tarifas bancárias", DF, 80),
    ])
    ins("""INSERT INTO fornecedores (nome, itens, prazo_pagamento_dias, prazo_entrega_dias, condicoes, critico)
           VALUES (?,?,?,?,?,?)""", [
        ("Distribuidora de peças de informática", "Memórias, SSDs, fontes, cabos", 28, 5, "Boleto 28 dias", 1),
        ("Provedor de hospedagem", "Hospedagem e domínios", 0, 0, "Cartão, mensal", 1),
        ("Software de acesso remoto", "Licenças por técnico", 0, 0, "Anual", 0),
    ])
    ins("""INSERT INTO propostas (persona_nome, valores, que_desejam, por_meio, diferentemente, oferecem)
           VALUES (?,?,?,?,?,?)""", [
        ("Cláudia", "Rapidez; Garantia; Prevenção", "manter o sistema de vendas funcionando sem sustos",
         "contrato mensal com atendimento no mesmo dia, backup e manutenção preventiva",
         "técnicos avulsos que só aparecem quando o problema já aconteceu", "tranquilidade para vender todos os dias"),
    ])
    ins("INSERT INTO riscos (risco, probabilidade, impacto, mitigacao, responsavel) VALUES (?,?,?,?,?)", [
        ("Equipe sem horas livres para novos clientes", 3, 4, "Acompanhar a ocupação mensal e contratar no mês 25", "Sócio(a) 1"),
        ("Inadimplência dos contratos mensais", 3, 3, "Cobrança por boleto recorrente e suspensão após 30 dias", "Sócio(a) 2"),
        ("Perda de dados de um cliente durante o serviço", 2, 5, "Backup antes de cada intervenção e termo de responsabilidade", "Sócio(a) 1"),
    ])
    ins("""INSERT INTO experimentos (hipotese, experimento, metrica, criterio_sucesso, resultado, status)
           VALUES (?,?,?,?,?,?)""", [
        ("Lojas fecham contrato mensal após um diagnóstico gratuito", "Oferecer diagnóstico a 25 lojas do centro",
         "Contratos fechados", "Pelo menos 5 contratos", "", "Planejado")])
    ins("INSERT INTO cronograma (atividade, mes_inicio, duracao_meses, responsavel, custo) VALUES (?,?,?,?,?)", [
        ("Formalização e sala", 1, 1, "Sócio(a) 2", 5500), ("Diagnósticos gratuitos", 1, 2, "Sócio(a) 1", 0),
        ("Lançamento do site e das redes", 1, 1, "Sócio(a) 2", 3000), ("Contratação do técnico júnior", 25, 1, "Sócio(a) 1", 0),
    ])



# ==============================================================================
# IDENTIDADE, OBJETIVOS SMART E MATRIZES ESTRATÉGICAS DOS TRÊS EXEMPLOS
# ==============================================================================
_ESTRATEGIA = {
    "fitbox": {
        "missao": "Oferecer refeições saudáveis, saborosas e práticas a quem tem pouco tempo, com ingredientes frescos e entrega rápida.",
        "visao": "Ser a principal referência em alimentação saudável de Teófilo Otoni até 2030.",
        "valores": "Qualidade; Respeito ao cliente; Transparência; Sustentabilidade; Valorização das pessoas",
        "rel": "Atendimento personalizado pelo WhatsApp\nCardápio semanal enviado aos clientes\nPrograma de fidelidade",
        "par": "Nutricionista parceira\nAcademias do bairro\nFornecedores de hortifrúti locais",
        "obj": [
            ("Atingir a receita prevista para o primeiro ano", "Financeira", "Receita bruta acumulada", 680000, "R$", 12, "Sócia administradora",
             "Meta baseada no volume e nos preços calculados no plano.", "Sustenta a viabilidade projetada.", "Acompanhamento semanal das vendas", ""),
            ("Fidelizar clientes com o programa de pontos", "Clientes", "Clientes que compram toda semana", 150, "clientes", 9, "Sócio comercial",
             "Base de clientes prevista e programa de baixo custo.", "Recompra reduz o custo de aquisição.", "Cartão fidelidade digital",
             "Atingir a receita prevista para o primeiro ano"),
            ("Reduzir o desperdício na produção", "Processos internos", "Perdas sobre a produção", 3, "% (máximo)", 6, "Chefe de cozinha",
             "Fichas técnicas já padronizadas.", "Desperdício corrói a margem de contribuição.", "Controle diário de sobras",
             "Atingir a receita prevista para o primeiro ano"),
            ("Treinar toda a equipe em boas práticas de manipulação", "Aprendizado e crescimento", "Funcionários certificados", 100, "%", 3,
             "Sócia administradora", "Curso gratuito do Sebrae/Anvisa disponível.", "Garante a segurança alimentar prometida.",
             "Curso de boas práticas", "Reduzir o desperdício na produção"),
        ],
        "bcg": [("Marmita fitness tradicional", 6, 1.4), ("Bowl de salmão", 15, 0.3), ("Suco natural 500 ml", 4, 0.6),
                ("Shake proteico com whey", 14, 1.1)],
        "porter": [(4, "Muitos restaurantes e marmitarias no centro.", "Diferenciar pela nutrição balanceada e pela entrega pontual."),
                   (4, "Investimento inicial baixo e receitas fáceis de copiar.", "Construir marca e fidelizar antes da concorrência."),
                   (3, "Cliente pode cozinhar em casa ou comprar congelados.", "Mostrar economia de tempo e variedade semanal."),
                   (3, "Cliente compara preço pelo aplicativo.", "Combos e planos semanais com desconto."),
                   (2, "Vários fornecedores de alimentos na região.", "Manter dois fornecedores para cada insumo crítico.")],
    },
    "cafe": {
        "missao": "Servir cafés especiais e quitandas mineiras feitas na hora, valorizando os produtores do Vale do Mucuri.",
        "visao": "Ser o café de referência do centro de Teófilo Otoni até 2029.",
        "valores": "Compromisso com a região; Qualidade; Excelência no atendimento; Confiança",
        "rel": "Atendimento pelo nome no balcão\nCartão fidelidade (10º café grátis)\nEventos de degustação",
        "par": "Produtores de café do Vale\nLaticínio regional\nCoworkings do centro",
        "obj": [
            ("Vender a média diária de cafés prevista no plano", "Financeira", "Cafés vendidos por dia", 140, "cafés/dia", 6, "Sócio(a) 2",
             "Fluxo de pedestres do centro e capacidade da máquina.", "O café é o produto de maior margem.", "Combos de café da manhã", ""),
            ("Conquistar clientes recorrentes do centro", "Clientes", "Clientes com cartão fidelidade ativo", 400, "clientes", 12, "Sócio(a) 1",
             "Cerca de 9.000 pessoas no público alcançável.", "Recompra diária sustenta a receita.", "Parcerias com escritórios",
             "Vender a média diária de cafés prevista no plano"),
            ("Padronizar o preparo de todas as bebidas", "Processos internos", "Bebidas com ficha de preparo", 100, "%", 3, "Sócio(a) 2",
             "Cardápio enxuto e barista experiente.", "Qualidade constante é a promessa da marca.", "Fichas de preparo e treinamento",
             "Conquistar clientes recorrentes do centro"),
            ("Formar a equipe em atendimento e barismo", "Aprendizado e crescimento", "Horas de treinamento por atendente", 16, "horas", 4,
             "Sócio(a) 2", "Treinamento interno com o sócio barista.", "Atendimento é diferencial frente às redes.",
             "Treinamentos mensais", "Padronizar o preparo de todas as bebidas"),
        ],
        "canais": ["Loja física (balcão)", "Instagram", "WhatsApp (encomendas)", "Indicação (boca a boca)"],
        "bcg": [("Café expresso", 8, 1.1), ("Pão de queijo (unidade)", 5, 1.5), ("Café coado (método)", 14, 0.5)],
        "porter": [(3, "Padarias tradicionais e uma rede de cafeteria por perto.", "Foco em café especial e origem regional."),
                   (3, "Abrir uma cafeteria exige investimento moderado em equipamentos.", "Fidelizar e criar marca local forte."),
                   (3, "Café de máquina no trabalho e cafés de padaria.", "Experiência e qualidade que justificam o preço."),
                   (3, "Cliente individual, compra pequena, mas compara com a padaria.", "Combos e programa de fidelidade."),
                   (4, "Café especial depende de poucos produtores e o preço do grão oscila.", "Contratos trimestrais e mais de um produtor.")],
    },
    "assinatura": {
        "missao": "Conectar moradores e visitantes ao que acontece no Vale do Mucuri e ajudar pequenas empresas locais a atrair clientes.",
        "visao": "Ser o principal aplicativo de cultura, lazer e consumo local do Vale do Mucuri até 2030.",
        "valores": "Compromisso com a região; Transparência; Inovação; Simplicidade",
        "rel": "Gerente de conta para empresas\nRelatório mensal de resultados\nSuporte pelo WhatsApp\nComunidade de usuários no app",
        "par": "Secretarias de cultura e turismo\nProdutores de eventos\nInfluenciadores locais\nProvedor de nuvem",
        "obj": [
            ("Atingir a receita recorrente mensal prevista para o fim do primeiro ano", "Financeira", "MRR", 37000, "R$/mês", 12,
             "Sócio(a) 2", "Projeção por safras com churn e captação realistas.", "MRR sustenta a viabilidade do aplicativo.",
             "Metas mensais de vendas por plano", ""),
            ("Manter o cancelamento de empresas abaixo do previsto", "Clientes", "Churn mensal das empresas", 5, "% (máximo)", 12,
             "Sócio(a) 2", "Relatórios mensais mostram resultado ao cliente.", "Cada ponto de churn reduz o LTV.",
             "Programa de sucesso do cliente", "Atingir a receita recorrente mensal prevista para o fim do primeiro ano"),
            ("Publicar as divulgações das empresas em até 24 horas", "Processos internos", "Publicações no prazo", 95, "%", 6,
             "Atendimento e conteúdo", "Fluxo de aprovação simples e equipe dedicada.", "Agilidade é o que as empresas cobram.",
             "Fila de publicação com prazos", "Manter o cancelamento de empresas abaixo do previsto"),
            ("Lançar uma melhoria do aplicativo por mês", "Aprendizado e crescimento", "Versões publicadas", 12, "versões/ano", 12,
             "Sócio(a) 1", "Equipe técnica própria e ciclo curto de desenvolvimento.", "Produto melhor reduz o churn.",
             "Roteiro mensal de melhorias", "Publicar as divulgações das empresas em até 24 horas"),
        ],
        "canais": ["Aplicativo (Google Play e App Store)", "Venda consultiva às empresas", "Instagram", "Acordos e Parcerias"],
        "bcg": [("Cobertura de evento patrocinado", 18, 0.3), ("Produção de conteúdo avulso (post e vídeo)", 12, 0.2),
                ("Assinaturas (todos os planos)", 25, 0.6)],
        "porter": [(3, "Plataformas nacionais e perfis locais de Instagram disputam a atenção.", "Conteúdo hiperlocal e relatórios de resultado."),
                   (4, "Criar um aplicativo é relativamente barato; a barreira é a base de usuários.", "Crescer rápido a comunidade e as parcerias."),
                   (4, "Redes sociais gratuitas cumprem parte da função.", "Mostrar resultado medido (cupons usados, visitas)."),
                   (3, "Empresas pequenas cancelam com facilidade.", "Plano anual com desconto e acompanhamento mensal."),
                   (3, "Dependência de lojas de aplicativos e provedor de nuvem.", "Cobrar empresas fora da loja e evitar dependência de um só provedor.")],
    },
    "servico": {
        "missao": "Manter a tecnologia das pequenas empresas do Vale do Mucuri funcionando, com atendimento rápido, garantia e prevenção.",
        "visao": "Ser a empresa de suporte de TI mais lembrada pelo comércio de Teófilo Otoni até 2030.",
        "valores": "Agilidade; Honestidade no diagnóstico; Segurança dos dados; Aprendizado contínuo",
        "rel": "Atendimento pelo WhatsApp com prazo de resposta\nVisita preventiva mensal aos clientes com contrato\nRelatório de chamados",
        "par": "Escritórios de contabilidade\nDistribuidora de peças\nProvedor de hospedagem\nSebrae (capacitação de clientes)",
        "obj": [
            ("Chegar a 25 contratos mensais de manutenção", "Financeira", "Contratos ativos", 25, "contratos", 12, "Sócio(a) 2",
             "Começa com 14 e o diagnóstico gratuito gera novos contratos.", "Contrato dá receita previsível e reduz ociosidade.",
             "Diagnóstico gratuito e indicação de contadores", ""),
            ("Atender os chamados no mesmo dia", "Clientes", "Chamados resolvidos em até 8 horas", 90, "%", 6, "Sócio(a) 1",
             "Equipe de dois técnicos e acesso remoto.", "Rapidez é o principal motivo de escolha.", "Fila de chamados com prazo",
             "Chegar a 25 contratos mensais de manutenção"),
            ("Manter a ocupação da equipe entre 70% e 90%", "Processos internos", "Horas vendidas ÷ horas cobráveis", 85, "%", 12,
             "Sócio(a) 1", "Ocupação calculada na ferramenta de serviços.", "Ocupação baixa é custo; alta demais derruba a qualidade.",
             "Registro de horas por chamado", "Atender os chamados no mesmo dia"),
            ("Certificar os técnicos em redes e segurança", "Aprendizado e crescimento", "Técnicos certificados", 2, "técnicos", 9,
             "Sócio(a) 1", "Cursos gratuitos e de baixo custo disponíveis.", "Abre o serviço de adequação à LGPD.",
             "Plano de cursos", "Manter a ocupação da equipe entre 70% e 90%"),
        ],
        "canais": ["Indicação (boca a boca)", "WhatsApp Business", "Instagram", "Acordos e Parcerias"],
        "bcg": [("Contrato mensal de manutenção (6 h incluídas)", 12, 1.5), ("Hora técnica avulsa", 3, 0.8),
                ("Site institucional", 9, 0.4), ("Loja virtual", 16, 0.2), ("Formatação e configuração de computador", 1, 1.1)],
        "porter": [(4, "Muitos técnicos autônomos e lojas de informática.", "Vender contrato, garantia e prazo, não só a hora."),
                   (4, "Basta conhecimento e um notebook para abrir.", "Construir carteira de contratos e reputação."),
                   (3, "Suporte do próprio fabricante e sobrinho que entende de computador.", "Mostrar o custo de uma parada do sistema."),
                   (3, "Pequenas empresas comparam preço e negociam.", "Planos com horas incluídas e preço fechado."),
                   (2, "Peças e softwares têm vários fornecedores.", "Dois fornecedores para peças críticas.")],
    },
}


def popular_estrategia(conn, qual: str) -> None:
    e = _ESTRATEGIA[qual]
    conn.execute("UPDATE projeto SET missao=?, visao=?, valores=?, canvas_relacionamento=?, canvas_parcerias=? WHERE id=1",
                 (e["missao"], e["visao"], e["valores"], e["rel"], e["par"]))
    conn.executemany("""INSERT INTO objetivos (objetivo, perspectiva, indicador, meta, unidade, prazo_meses, responsavel,
                        por_que_atingivel, relevancia, iniciativa, contribui_para) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", e["obj"])
    conn.executemany("INSERT INTO bcg (produto, crescimento_mercado_pct, participacao_relativa) VALUES (?,?,?)", e["bcg"])
    if e.get("canais") and conn.execute("SELECT COUNT(*) FROM canais").fetchone()[0] == 0:
        conn.executemany("INSERT INTO canais (canal, atratividade, alcance, cliente, relacionamento, viabilidade) VALUES (?,?,?,?,?,?)",
                         [(c, 8, 7, 8, 7, 8) for c in e["canais"]])
    forcas = ["Rivalidade entre concorrentes", "Ameaça de novos entrantes", "Ameaça de produtos substitutos",
              "Poder de barganha dos clientes", "Poder de barganha dos fornecedores"]
    conn.executemany("INSERT INTO porter (forca, intensidade, justificativa, estrategia) VALUES (?,?,?,?)",
                     [(f, *v) for f, v in zip(forcas, e["porter"])])
    popular_validacao(conn, qual)


# ------------------------------------------------------------------------------
# DADOS DE EXEMPLO DAS FERRAMENTAS DE VALIDAÇÃO E MARKETING (v7.1) - fictícios
# ------------------------------------------------------------------------------
def _respostas_vw(preco: float, n: int, semente: int) -> list:
    """Respostas fictícias e coerentes do teste de Van Westendorp em torno de um preço."""
    import random
    rnd = random.Random(semente)
    linhas = []
    for i in range(n):
        base = preco * rnd.uniform(0.75, 1.25)
        mb = round(base * rnd.uniform(0.45, 0.65), 0)
        ba = round(base * rnd.uniform(0.75, 0.9), 0)
        ca = round(base * rnd.uniform(1.05, 1.25), 0)
        mc = round(base * rnd.uniform(1.35, 1.7), 0)
        linhas.append((f"R{i + 1:02d}", mb, max(ba, mb), max(ca, ba), max(mc, ca)))
    return linhas


_VALIDACAO = {
    "fitbox": {
        "vw_produto": "Assinatura semanal - 5 marmitas tradicionais", "vw_preco": 129.9, "vw_n": 32,
        "intencao": (18.0, 31.0),
        "entrevistas": [
            ("03/02", "E01", "Trabalhadora de escritório", "falta de tempo; comida pouco saudável perto do trabalho", "marmita de casa", "Sim", 130, "Quer variar o cardápio na semana"),
            ("03/02", "E02", "Praticante de academia", "dificuldade de controlar macros; falta de tempo", "prepara no domingo", "Sim", 150, "Valoriza tabela nutricional"),
            ("04/02", "E03", "Estudante universitário", "preço alto; comida pouco saudável perto do trabalho", "restaurante a quilo", "Talvez", 90, "Sensível a preço"),
            ("04/02", "E04", "Profissional liberal", "falta de tempo; enjoa da mesma comida", "aplicativo de delivery", "Sim", 140, "Pediu entrega até 11h"),
            ("05/02", "E05", "Mãe que trabalha fora", "falta de tempo; dificuldade de controlar macros", "congela comida", "Sim", 120, "Quer opção para a família"),
            ("05/02", "E06", "Praticante de corrida", "dificuldade de controlar macros", "marmitaria fitness", "Talvez", 110, "Já usa concorrente"),
            ("06/02", "E07", "Servidor público", "comida pouco saudável perto do trabalho; falta de tempo", "restaurante a quilo", "Sim", 130, "Almoça no trabalho"),
            ("06/02", "E08", "Estudante universitário", "preço alto", "lanche", "Não", 60, "Não é o público prioritário"),
            ("07/02", "E09", "Trabalhadora de escritório", "falta de tempo; enjoa da mesma comida", "marmita de casa", "Sim", 125, "Indicaria a colegas"),
            ("07/02", "E10", "Personal trainer", "dificuldade de controlar macros", "prepara em casa", "Talvez", 140, "Possível parceiro de indicação"),
            ("08/02", "E11", "Profissional de saúde", "falta de tempo; comida pouco saudável perto do trabalho", "aplicativo de delivery", "Sim", 135, "Plantões longos"),
            ("08/02", "E12", "Praticante de academia", "falta de tempo; dificuldade de controlar macros", "marmitaria fitness", "Sim", 145, "Quer suplementos junto"),
        ],
        "evidencias": {"Preço de venda": ("Testada com clientes (entrevista, pré-venda, piloto)", "32 respostas do teste de preço e 12 entrevistas (fev.)"),
                       "Custo de materiais e mercadorias": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Orçamentos de 3 fornecedores"),
                       "Salários e encargos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Piso da categoria e tabela do contador"),
                       "Demais custos fixos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Aluguel e contas cotados")},
        "vpc": ("Quando tenho pouco tempo no almoço, quero comer algo saudável e saboroso, para manter a dieta sem perder o horário.",
                "Almoçar em pouco tempo\nSeguir a dieta da nutricionista\nPlanejar a alimentação da semana",
                "Falta de tempo\nComida pouco saudável perto do trabalho\nDificuldade de controlar macros\nEnjoa da mesma comida",
                "Sentir mais disposição\nVer resultado no treino\nEconomizar com delivery caro",
                "Marmitas fitness\nAssinatura semanal e mensal\nCoffee break saudável",
                "Entrega no trabalho até 11h\nTabela nutricional em cada marmita\nCardápio que muda toda semana",
                "Acompanhamento com nutricionista parceira\nPreço de assinatura menor que o delivery"),
        "encaixe": [("Falta de tempo", "Dor", "Entrega no trabalho até 11h"),
                    ("Comida pouco saudável perto do trabalho", "Dor", "Tabela nutricional em cada marmita"),
                    ("Dificuldade de controlar macros", "Dor", "Tabela nutricional em cada marmita"),
                    ("Enjoa da mesma comida", "Dor", "Cardápio que muda toda semana"),
                    ("Ver resultado no treino", "Ganho", "Acompanhamento com nutricionista parceira"),
                    ("Economizar com delivery caro", "Ganho", "Preço de assinatura menor que o delivery")],
        "errc": [("Variedade", "Reduzir", 6, "Cardápio enxuto e rotativo reduz custo e desperdício"),
                 ("Design", "Reduzir", 5, "Embalagem simples e reciclável"),
                 ("Transparência nutricional", "Elevar", 10, "Tabela completa e macros em cada marmita"),
                 ("Conveniência", "Elevar", 9, "Entrega programada no trabalho"),
                 ("Personalização", "Criar", 8, "Ajuste de porções conforme o objetivo do cliente")],
        "posic": ("trabalhadores e praticantes de atividade física de Teófilo Otoni", "não têm tempo de preparar refeições saudáveis",
                  "FitBox", "uma assinatura de refeições saudáveis", "entrega no trabalho refeições balanceadas com controle de macros",
                  "marmitarias comuns e aplicativos de delivery", "cada marmita traz a tabela nutricional e o cardápio é planejado por nutricionista"),
        "mapa": ("Preço (baixo → alto)", "Saudabilidade (baixa → alta)",
                 [("FitBox", 6, 9), ("Concorrente A (marmitaria fitness)", 7, 7), ("Concorrente B (restaurante natural)", 8, 8),
                  ("Concorrente C (loja de suplementos)", 5, 6), ("Restaurante a quilo", 4, 4)]),
        "orc": [("Publicidade em Redes Sociais", 1500, 45, "Cupom por anúncio"), ("Acordos e Parcerias", 600, 25, "Código de academia parceira"),
                ("Marketing de Conteúdo", 400, 12, "Link na bio"), ("Eventos", 800, 10, "Inscrições no estande")],
        "aarrr": [("Aquisição", "Visitas ao perfil e cadastros", "300 cadastros/mês", "Anúncios e parcerias com academias"),
                  ("Ativação", "Primeiro pedido em 7 dias", "40% dos cadastros", "Cupom de primeira semana"),
                  ("Retenção", "Assinantes que renovam no mês", "75%", "Cardápio rotativo e avaliação semanal"),
                  ("Receita", "Ticket médio mensal", "R$ 450", "Upgrade para plano mensal"),
                  ("Indicação", "Clientes vindos de indicação", "20%", "Uma marmita grátis por indicação")],
        "cal": [(1, "Lançamento com degustação em academias", "Eventos", 1200), (1, "Campanha de anúncios de abertura", "Redes sociais", 1500),
                (3, "Parceria com nutricionistas", "Parcerias", 600), (6, "Campanha de indicação", "Indicação", 500)],
        "perda": [40000, 30000, 25000, 15000],
        "formal": {"Contratar contador ou escritório contábil": "Concluído",
                   "Consultar a viabilidade do endereço na prefeitura (zoneamento)": "Em andamento",
                   "Escolher a natureza jurídica (MEI, empresário individual, SLU ou LTDA)": "Concluído"},
        "banca": [("Por que este é o momento certo para abrir este negócio?",
                   "A procura por alimentação saudável cresce na cidade e nenhum concorrente entrega no local de trabalho com controle de macros. "
                   "As 12 entrevistas mostraram que falta de tempo é a dor mais citada.")],
    },
    "cafe": {
        "vw_produto": "Cappuccino", "vw_preco": 11.0, "vw_n": 20, "intencao": (12.0, 25.0), "entrevistas": [],
        "evidencias": {"Demais custos fixos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Aluguel cotado")},
        "vpc": ("Quando quero uma pausa no meio da tarde, quero um café de qualidade num lugar agradável, para descansar e encontrar amigos.",
                "Fazer uma pausa no trabalho\nEncontrar amigos\nComprar pão para casa",
                "Café fraco ou requentado\nFila na padaria\nPoucos lugares para sentar", "Ambiente acolhedor\nProdutos artesanais",
                "Cafés especiais\nPães de fermentação natural\nBolos caseiros", "Café moído na hora\nAtendimento rápido no balcão",
                "Ambiente com mesas e Wi-Fi\nProdutos da região"),
        "errc": [("Rapidez", "Manter", 0, ""), ("Ambiente", "Elevar", 9, "Mesas, tomadas e Wi-Fi"),
                 ("Presença digital", "Elevar", 8, "Encomendas pelo WhatsApp")],
        "posic": ("moradores e trabalhadores do centro", "querem um café de qualidade sem pagar preço de rede", "Café Mucuri",
                  "uma cafeteria de bairro", "serve café especial e pães artesanais com preço justo",
                  "padarias tradicionais e redes de cafeteria", "usamos grãos e farinhas de produtores da região"),
        "mapa": ("Preço (baixo → alto)", "Qualidade do café (baixa → alta)",
                 [("Café Mucuri", 5, 8), ("Padaria tradicional do centro", 3, 4), ("Rede de cafeteria", 8, 8)]),
        "orc": [("Instagram", 600, 30, "Cupom do perfil"), ("Indicação (boca a boca)", 200, 20, "Cartão fidelidade")],
        "aarrr": [], "cal": [], "perda": [35000, 25000], "formal": {}, "banca": [],
    },
    "assinatura": {
        "vw_produto": "Assinatura (ticket médio)", "vw_preco": 0, "vw_n": 0, "intencao": (8.0, 20.0), "entrevistas": [],
        "evidencias": {"Cancelamento mensal (churn)": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Referências de SaaS para pequenas empresas")},
        "vpc": ("Quando organizo um evento na cidade, quero que o público certo saiba dele, para encher o evento sem gastar muito com anúncios.",
                "Divulgar eventos\nAtrair público local\nMedir o resultado da divulgação",
                "Anúncio caro e pouco segmentado\nNão sabe se a divulgação deu resultado", "Público certo no evento\nRelatório para os patrocinadores",
                "Aplicativo de agenda local\nPlanos de parceiro\nCobertura de eventos", "Público segmentado por cidade e interesse\nRelatório mensal de visualizações",
                "Selo de parceiro destaque\nCobertura em vídeo"),
        "errc": [("Conteúdo local", "Elevar", 10, "Agenda completa do Vale"), ("Relatórios de resultado", "Criar", 8, "Relatório mensal ao parceiro")],
        "posic": ("produtores de eventos e pequenos comércios do Vale do Mucuri", "querem divulgar para o público local", "Conecta Vale",
                  "um aplicativo de agenda e divulgação local", "leva o evento ao público certo com relatório de resultado",
                  "perfis locais no Instagram", "o parceiro recebe todo mês o relatório de visualizações e cliques"),
        "mapa": ("Preço (baixo → alto)", "Alcance local (baixo → alto)",
                 [("Conecta Vale", 4, 8), ("Plataforma nacional de eventos", 7, 4), ("Perfis locais no Instagram", 3, 6)]),
        "orc": [("Venda consultiva às empresas", 1200, 8, "Contratos fechados"), ("Instagram", 500, 4, "Cadastros pelo link")],
        "aarrr": [], "cal": [], "perda": [30000, 30000], "formal": {}, "banca": [],
    },
    "servico": {
        "vw_produto": "Contrato mensal de manutenção (6 h incluídas)", "vw_preco": 690, "vw_n": 24, "intencao": (10.0, 22.0),
        "entrevistas": [
            ("10/03", "S01", "Loja de roupas", "computador para de funcionar; técnico demora", "sobrinho", "Sim", 600, "Perde vendas quando o sistema para"),
            ("10/03", "S02", "Escritório de contabilidade", "medo de perder dados; técnico demora", "técnico autônomo", "Sim", 900, "Quer backup diário"),
            ("11/03", "S03", "Clínica odontológica", "medo de perder dados; LGPD", "empresa da capital", "Talvez", 800, "Pede contrato com garantia"),
            ("11/03", "S04", "Mercearia", "computador para de funcionar", "loja de informática", "Não", 250, "Prefere pagar só quando quebra"),
            ("12/03", "S05", "Loja de calçados", "técnico demora; computador para de funcionar", "técnico autônomo", "Sim", 650, "Indica outros lojistas"),
            ("12/03", "S06", "Imobiliária", "não tem site; medo de perder dados", "nenhuma", "Talvez", 500, "Interesse maior no site"),
        ],
        "encaixe": [],
        "evidencias": {"Salários e encargos": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Piso da categoria e tabela do contador"),
                       "Preço de venda": ("Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)", "Preços de 4 concorrentes locais")},
        "vpc": ("Quando o computador ou o sistema da loja para, quero que alguém resolva no mesmo dia, para não perder vendas.",
                "Manter o sistema de vendas funcionando\nGuardar os dados com segurança\nTer presença na internet",
                "Técnico demora\nComputador para de funcionar\nMedo de perder dados\nNão tem site",
                "Previsibilidade de gastos\nSaber que os dados estão seguros",
                "Contrato mensal de manutenção\nHora técnica\nSites e lojas virtuais",
                "Atendimento no mesmo dia\nManutenção preventiva\nBackup automático\nSite em pacote fechado",
                "Preço fixo mensal\nRelatório de backup"),
        "errc": [("Preço", "Reduzir", 6, "Não competir com o técnico informal pelo menor preço"),
                 ("Garantia do serviço", "Elevar", 10, "Garantia por escrito de 90 dias"),
                 ("Rapidez no atendimento", "Elevar", 10, "Atendimento no mesmo dia para contratos"),
                 ("Manutenção preventiva", "Criar", 8, "Visita mensal antes de o problema acontecer")],
        "posic": ("pequenas empresas do comércio e de serviços de Teófilo Otoni", "não podem parar quando o computador falha",
                  "TecVale Soluções", "um serviço de suporte de TI por contrato", "resolve no mesmo dia e evita que o problema aconteça",
                  "técnicos avulsos e lojas de informática", "o contrato inclui visita preventiva mensal, backup e garantia por escrito"),
        "mapa": ("Preço (baixo → alto)", "Rapidez e garantia (baixa → alta)",
                 [("TecVale Soluções", 6, 9), ("Técnicos autônomos", 3, 4), ("Loja de informática do centro", 5, 5),
                  ("Agência digital da capital", 9, 5)]),
        "orc": [("Indicação (boca a boca)", 200, 3, "Desconto de um mês por indicação"), ("Instagram", 500, 2, "Mensagens pelo link"),
                ("Acordos e Parcerias", 300, 2, "Contadores parceiros")],
        "aarrr": [], "cal": [], "perda": [20000, 20000], "formal": {}, "banca": [],
    },
}


# Dados reais de Teófilo Otoni (MG) consultados na API do IBGE em 25/11/2026
_DADOS_TEOFILO_OTONI = [
    ("populacao", "População residente (Censo 2022)", 137418, "2022", "IBGE. Censo Demográfico 2022, tabela 4709 (Sidra)."),
    ("domicilios", "Domicílios particulares permanentes ocupados (Censo 2022)", 49858, "2022", "IBGE. Censo Demográfico 2022, tabela 4712 (Sidra)."),
    ("pib", "PIB a preços correntes (mil R$)", 3540248, "2023", "IBGE. Produto Interno Bruto dos Municípios, tabela 5938 (Sidra)."),
    ("unidades_locais", "Unidades locais de empresas e organizações", 3880, "2021", "IBGE. Cadastro Central de Empresas (Cempre), tabela 1685 (Sidra)."),
    ("pib_per_capita", "PIB per capita (R$)", 3540248 * 1000 / 137418, "2023", "Calculado: PIB ÷ população do Censo 2022 (IBGE)."),
    ("moradores_dom", "Moradores por domicílio", 137418 / 49858, "2022", "Calculado com o Censo 2022 (IBGE)."),
]
_V8 = {
    "fitbox": {"setor": "Serviços · Serviços de alimentação (restaurantes, lanchonetes)",
               "pof": ("pof_alimentacao", "Alimentação", 688.51),
               "swot": [("Ofensiva (forças × oportunidades)", "Usar o cardápio planejado por nutricionista para atender a procura por alimentação saudável",
                         "Parceria com 5 academias no primeiro trimestre", 3),
                        ("Confronto (forças × ameaças)", "Destacar a tabela nutricional para se diferenciar das marmitarias comuns",
                         "Selo de informação nutricional em todas as embalagens", 2),
                        ("Reforço (fraquezas × oportunidades)", "Reduzir a dependência de um só canal com vendas por assinatura",
                         "Lançar o plano mensal com desconto", 6),
                        ("Defensiva (fraquezas × ameaças)", "Proteger o caixa contra a alta dos insumos", "Contrato de preço com dois fornecedores", 4)],
               "lean": ("Marmita feita em casa; restaurante a quilo; aplicativos de delivery",
                        "Assinantes ativos; renovação mensal; custo por refeição; avaliação semanal",
                        "Nutricionista parceira e base de clientes das academias da cidade"),
               "imp": None, "ods": "3 Saúde e bem-estar", "rodada": (0, 0), "emp": None},
    "cafe": {"setor": "Comércio · Produtos alimentícios, bebidas e fumo (varejo)", "pof": ("pof_alimentacao", "Alimentação", 688.51),
             "swot": [("Ofensiva (forças × oportunidades)", "Valorizar os produtos da região para atrair quem busca experiência", "Cardápio com origem dos grãos", 3),
                      ("Defensiva (fraquezas × ameaças)", "Evitar falta de caixa nos meses fracos", "Reserva equivalente a dois meses de custos fixos", 6)],
             "lean": ("Padarias do centro; café em casa", "Ticket médio; clientes por dia; recompra no cartão fidelidade",
                      "Relação com produtores de café da região"),
             "imp": None, "ods": "8 Trabalho decente e crescimento econômico", "rodada": (0, 0), "emp": None},
    "assinatura": {"setor": "Serviços · Tecnologia da informação", "pof": None,
                   "swot": [("Ofensiva (forças × oportunidades)", "Ser a agenda oficial de eventos do Vale", "Parceria com a secretaria de cultura", 6)],
                   "lean": ("Perfis locais no Instagram; grupos de WhatsApp", "Parceiros pagantes; churn; usuários ativos por mês",
                            "Base de eventos locais que nenhum concorrente nacional cobre"),
                   "imp": ("Pouca divulgação da cultura e do comércio local", "Produtores culturais e pequenos comerciantes do Vale do Mucuri",
                           "Agenda gratuita de eventos e planos de divulgação", "Eventos divulgados por mês",
                           "Mais público nos eventos locais e mais renda para os produtores", "Economia criativa mais forte no Vale do Mucuri"),
                   "ods": "8 Trabalho decente e crescimento econômico;11 Cidades e comunidades sustentáveis", "rodada": (300000, 60000), "emp": None},
    "servico": {"setor": "Serviços · Manutenção e reparação de informática e comunicação", "pof": None,
                "swot": [("Confronto (forças × ameaças)", "Vender garantia e prazo contra a concorrência por preço dos técnicos informais",
                          "Garantia por escrito de 90 dias", 1),
                         ("Reforço (fraquezas × oportunidades)", "Aumentar a receita previsível com contratos para clínicas que precisam se adequar à LGPD",
                          "Pacote de backup e segurança", 6)],
                "lean": ("Sobrinho que entende de computador; técnico avulso; loja de informática",
                         "Contratos ativos; horas vendidas; tempo de resposta", "Atendimento no mesmo dia na cidade"),
                "imp": None, "ods": "", "rodada": (0, 0), "emp": ("Pronampe", 6, "SAC")},
}


def popular_v8(conn, qual: str) -> None:
    v = _V8[qual]
    conn.execute("UPDATE projeto SET setor_referencia=?, municipio_nome='Teófilo Otoni', municipio_ibge='3168606', uf_sigla='MG', "
                 "lean_alternativas=?, lean_metricas=?, lean_vantagem=?, ods=?, rodada_pre_money=?, rodada_investimento=? WHERE id=1",
                 (v["setor"], v["lean"][0].replace("; ", "\n"), v["lean"][1].replace("; ", "\n"), v["lean"][2], v["ods"], *v["rodada"]))
    dados = list(_DADOS_TEOFILO_OTONI)
    if v["pof"]:
        dados.append((v["pof"][0], v["pof"][1], v["pof"][2], "2017-2018",
                      "IBGE. Pesquisa de Orçamentos Familiares 2017-2018, tabela 6715 (Minas Gerais)."))
    conn.executemany("INSERT INTO dados_publicos (chave, rotulo, valor, periodo, fonte) VALUES (?,?,?,?,?)", dados)
    conn.executemany("INSERT INTO swot_estrategias (quadrante, estrategia, acao, prazo_meses) VALUES (?,?,?,?)", v["swot"])
    if v["imp"]:
        conn.execute("UPDATE projeto SET imp_problema=?, imp_publico=?, imp_atividades=?, imp_produtos=?, imp_resultados=?, imp_impacto=? "
                     "WHERE id=1", v["imp"])
        conn.execute("INSERT INTO impacto_indicadores (indicador, tipo, linha_base, meta, ods) VALUES (?,?,?,?,?)",
                     ("Eventos locais divulgados por mês", "Produto (output)", "0", "60", "11 Cidades e comunidades sustentáveis"))
    if v["emp"]:
        conn.execute("UPDATE emprestimos SET carencia_meses=?, sistema=? WHERE descricao=?", (v["emp"][1], v["emp"][2], v["emp"][0]))
    conn.execute("UPDATE emprestimos SET carencia_meses=0 WHERE carencia_meses IS NULL")
    conn.execute("UPDATE emprestimos SET sistema='Price' WHERE sistema IS NULL OR sistema=''")


_MIX = {
    "fitbox": ("Comer bem sem ter de pensar no cardápio nem cozinhar", "Marmitas com cardápio de nutricionista, embalagem própria e tabela nutricional",
               "Entrega agendada, troca de prato sem custo e plano mensal", "Crescimento", "Valor percebido", "Intermediário", (1.08, 0.92)),
    "cafe": ("Um lugar agradável para uma pausa com café da região", "Cafés especiais de produtores locais, pães e salgados da casa",
             "Wi-fi, cartão fidelidade e encomendas para festas", "Maturidade", "Referência na concorrência", "Intermediário", (1.05, 0.97)),
    "assinatura": ("Saber o que acontece na cidade e ser encontrado pelo público", "Aplicativo de agenda local com planos para parceiros",
                   "Suporte por WhatsApp e relatório mensal de visualizações", "Introdução", "Penetração", "Econômico", (1.30, 1.15)),
    "servico": ("Computador funcionando de novo, rápido e sem dor de cabeça", "Manutenção e suporte de TI com técnico certificado",
                "Garantia por escrito de 90 dias, coleta e entrega", "Maturidade", "Valor percebido", "Premium", (0.85, 0.90)),
}


_MIX2 = {   # distribuição, objetivo de comunicação, compras por mês, canais (canal, % vendas, comissão %)
    "fitbox": ("Seletiva", "Fazer o cliente experimentar", 4.0,
               [("Entrega própria", 60, 0), ("WhatsApp e redes sociais", 25, 0), ("Aplicativo de delivery", 15, 20)]),
    "cafe": ("Seletiva", "Fidelizar e fazer voltar", 6.0,
             [("Loja física", 85, 0), ("Aplicativo de delivery", 15, 20)]),
    "assinatura": ("Intensiva", "Tornar a marca conhecida", 1.0,
                   [("Site próprio", 70, 0), ("Venda direta a empresas", 30, 0)]),
    "servico": ("Exclusiva", "Converter em venda", 1.0,
                [("Loja física", 50, 0), ("Venda direta a empresas", 40, 0), ("WhatsApp e redes sociais", 10, 0)]),
}


_MIX3 = {   # pessoas, evidência física, etapas do mapa do serviço (etapa, onde, minutos, ponto de falha)
    "servico": ("Dois técnicos com certificação e roteiro padrão de atendimento; o sócio faz o contato comercial",
                "Ordem de serviço impressa, relatório do que foi feito, etiqueta de garantia no equipamento, avaliações no Google",
                [("Cliente pede atendimento pelo WhatsApp", "Visível ao cliente", 5, "Demora em responder"),
                 ("Agendamento e coleta do equipamento", "Visível ao cliente", 30, "Atraso na coleta"),
                 ("Diagnóstico na bancada", "Bastidores", 60, ""),
                 ("Orçamento enviado e aprovado", "Visível ao cliente", 15, "Orçamento sem detalhes"),
                 ("Compra de peças", "Apoio (sistemas e fornecedores)", 120, "Peça em falta no fornecedor"),
                 ("Reparo e testes", "Bastidores", 90, ""),
                 ("Entrega com relatório e garantia", "Visível ao cliente", 20, "")]),
    "cafe": ("Atendentes treinadas em preparo de café e no cardápio regional", "Ambiente com madeira e fotos dos produtores, "
             "xícaras com a marca, cardápio com a origem dos grãos",
             [("Cliente chega e escolhe no balcão", "Visível ao cliente", 3, "Fila no horário de pico"),
              ("Pedido e pagamento", "Visível ao cliente", 2, "Maquininha sem sinal"),
              ("Preparo do café", "Bastidores", 4, "Moagem errada"),
              ("Entrega na mesa", "Visível ao cliente", 1, "")]),
}


_MIX4 = {   # síntese da situação, Ansoff, estratégia competitiva, objetivos de clientes, indicadores (indicador, meta, frequência, fonte)
    "cafe": ("O centro tem fluxo alto de pessoas e poucas opções de café especial. A concorrência é de padarias com preço parecido. "
             "A principal fraqueza é depender do movimento de rua nos meses fracos.",
             "Penetração de mercado", "Diferenciação",
             [("Conquistar clientes que voltam toda semana", "Clientes no cartão fidelidade", 400, "clientes", 6),
              ("Aumentar o gasto médio por visita", "Ticket médio", 22, "reais", 9)],
             [("Clientes novos por mês", "50", "Mensal", "Cartão fidelidade"),
              ("Gasto médio por compra (ticket médio)", "R$ 22", "Semanal", "Sistema de caixa"),
              ("Nota de satisfação (NPS)", "60", "Trimestral", "Pesquisa no balcão")]),
    "servico": ("Empresas e profissionais liberais da cidade dependem de técnicos informais sem garantia. "
                "O concorrente mais forte cobra menos, mas não dá prazo nem garantia por escrito.",
                "Penetração de mercado", "Foco em diferenciação",
                [("Fechar contratos mensais com empresas", "Contratos ativos", 25, "contratos", 12)],
                [("Contratos mensais ativos", "25 até o mês 12", "Mensal", "Planilha de contratos"),
                 ("Tempo de resposta ao chamado", "até 2 horas", "Semanal", "WhatsApp"),
                 ("Custo por cliente novo (CAC)", "até R$ 150", "Mensal", "Gasto de marketing ÷ clientes novos")]),
}


_CAL_MIX = {
    "cafe": [(1, "Inauguração com degustação de cafés da região", "Eventos", 600), (1, "Perfil no Instagram e Google Meu Negócio", "Redes sociais", 0),
             (2, "Lançamento do cartão fidelidade", "Loja", 150), (4, "Parceria com o produtor de café do mês", "Parcerias", 200),
             (6, "Campanha do Dia dos Pais com kits de presente", "Redes sociais", 300)],
    "servico": [(1, "Visitas a escritórios do centro com proposta de contrato", "Venda direta", 300),
                (1, "Anúncio no Google para conserto de computador", "Busca paga", 400),
                (3, "Palestra gratuita sobre segurança digital na associação comercial", "Eventos", 250),
                (6, "Campanha de indicação: um mês grátis para quem indicar", "Indicação", 500)],
}


def popular_mix(conn, qual: str) -> None:
    """Exemplo do mix integrado (v8.1): níveis de produto, estratégia de preço e preços de concorrentes."""
    if qual not in _MIX:
        return
    nuc, real, amp, ciclo, est, pos, fatores = _MIX[qual]
    conn.execute("UPDATE projeto SET mix_nucleo=?, mix_real=?, mix_ampliado=?, mix_ciclo=?, mix_estrategia_preco=?, "
                 "mix_posicao_preco=? WHERE id=1", (nuc, real, amp, ciclo, est, pos))
    prods = conn.execute("SELECT descricao, preco_venda FROM catalogo WHERE preco_venda > 0 "
                         "ORDER BY preco_venda * qtd_mensal DESC LIMIT 2").fetchall()
    concs = [r[0] for r in conn.execute("SELECT nome FROM concorrentes LIMIT 2").fetchall()]
    for c, f in zip(concs, fatores):
        for prod, preco in prods:
            conn.execute("INSERT INTO precos_concorrentes (concorrente, produto, preco, fonte) VALUES (?,?,?,?)",
                         (c, prod, round(preco * f, 2), "Exemplo didático"))
    if qual in _MIX2:
        dist, obj, freq, canais = _MIX2[qual]
        conn.execute("UPDATE projeto SET mix_distribuicao=?, mix_objetivo_com=?, mix_freq=? WHERE id=1", (dist, obj, freq))
        conn.executemany("INSERT INTO canais_venda (canal, pct_vendas, comissao_pct, observacao) VALUES (?,?,?,?)",
                         [(c, p, k, "Exemplo didático: confira a taxa real do canal") for c, p, k in canais])
    if qual in _MIX3:
        pessoas, evid, etapas = _MIX3[qual]
        conn.execute("UPDATE projeto SET mix_pessoas=?, mix_evidencia=? WHERE id=1", (pessoas, evid))
        conn.executemany("INSERT INTO blueprint (etapa, linha, tempo_min, ponto_falha) VALUES (?,?,?,?)", etapas)
    if qual in _MIX4:
        sit, ans, gen, objs, kpis = _MIX4[qual]
        conn.execute("UPDATE projeto SET pm_situacao=?, pm_ansoff=?, pm_generica=? WHERE id=1", (sit, ans, gen))
        conn.executemany("INSERT INTO objetivos (objetivo, perspectiva, indicador, meta, unidade, prazo_meses) VALUES (?,'Clientes',?,?,?,?)",
                         objs)
        conn.executemany("INSERT INTO kpis_mkt (indicador, meta, frequencia, fonte, responsavel) VALUES (?,?,?,?,'Sócio-gerente')", kpis)
        if not conn.execute("SELECT COUNT(*) FROM calendario_mkt").fetchone()[0]:
            conn.executemany("INSERT INTO calendario_mkt (mes, acao, canal, custo) VALUES (?,?,?,?)", _CAL_MIX[qual])


def popular_validacao(conn, qual: str) -> None:
    v = _VALIDACAO[qual]
    vpc = v["vpc"]
    conn.execute("""UPDATE projeto SET vw_produto=?, intencao_certamente_pct=?, intencao_provavel_pct=?, vpc_job=?, vpc_tarefas=?,
                    vpc_dores=?, vpc_ganhos=?, vpc_produtos=?, vpc_aliviadores=?, vpc_criadores=?, posic_segmento=?,
                    posic_necessidade=?, posic_marca=?, posic_categoria=?, posic_beneficio=?, posic_concorrente=?,
                    posic_diferencial=?, mapa_eixo_x=?, mapa_eixo_y=? WHERE id=1""",
                 (v["vw_produto"], *v["intencao"], *vpc, *v["posic"], v["mapa"][0], v["mapa"][1]))
    if v["vw_n"]:
        conn.executemany("INSERT INTO precos_vw (respondente, muito_barato, barato, caro, muito_caro) VALUES (?,?,?,?,?)",
                         _respostas_vw(v["vw_preco"], v["vw_n"], 7))
    conn.executemany("""INSERT INTO entrevistas (data, entrevistado, perfil, dores, solucao_atual, pagaria, valor_pagaria, aprendizado)
                        VALUES (?,?,?,?,?,?,?,?)""", v["entrevistas"])
    conn.executemany("INSERT INTO evidencias (premissa, status, fonte, observacao) VALUES (?,?,?,'')",
                     [(k, s, f) for k, (s, f) in v["evidencias"].items()])
    conn.executemany("INSERT INTO vpc_encaixe (item, tipo, resposta) VALUES (?,?,?)", v.get("encaixe", []))
    conn.executemany("INSERT INTO errc (atributo, acao, nova_nota, justificativa) VALUES (?,?,?,?)", v["errc"])
    conn.executemany("INSERT INTO mapa_perceptual (marca, eixo_x, eixo_y) VALUES (?,?,?)", v["mapa"][2])
    conn.executemany("INSERT INTO orcamento_canais (canal, gasto_mensal, clientes_mes, observacao) VALUES (?,?,?,?)", v["orc"])
    conn.executemany("INSERT INTO aarrr (etapa, metrica, meta, acao) VALUES (?,?,?,?)", v["aarrr"])
    conn.executemany("INSERT INTO calendario_mkt (mes, acao, canal, custo) VALUES (?,?,?,?)", v["cal"])
    ids = [r[0] for r in conn.execute("SELECT id FROM socios ORDER BY id")]
    conn.executemany("UPDATE socios SET perda_aceitavel=? WHERE id=?", list(zip(v["perda"], ids)))
    conn.executemany("INSERT INTO formalizacao (item, status, observacao) VALUES (?,?,'')", list(v["formal"].items()))
    conn.executemany("INSERT INTO banca_respostas (pergunta, resposta) VALUES (?,?)", v["banca"])
    popular_v8(conn, qual)
    popular_mix(conn, qual)

# ##############################################################################
# PARTE: pnbox_motor.py
# ##############################################################################



CATEGORIAS = ["Fabricação Própria", "Revenda", "Serviço"]


# ------------------------------------------------------------------------------
# Formatação no padrão brasileiro
# ------------------------------------------------------------------------------
def brl(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def num_br(v, casas=2) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "n.a."
    if math.isnan(v) or math.isinf(v):
        return "n.a."
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, casas=2) -> str:
    return "n.a." if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))) \
        else f"{num_br(v * 100, casas)}%"


def div(a, b):
    return a / b if b not in (0, 0.0) and b is not None else float("nan")


# ------------------------------------------------------------------------------
# 1. CUSTO UNITÁRIO DOS PRODUTOS (FICHA TÉCNICA x CUSTO INFORMADO)
# ------------------------------------------------------------------------------
def custos_unitarios(cat: pd.DataFrame, mat: pd.DataFrame, ficha: pd.DataFrame) -> pd.DataFrame:
    precos = dict(zip(mat["nome"], mat["valor_unitario"]))
    f = ficha.copy()
    f["valor_unitario"] = f["material"].map(precos).fillna(0.0)
    f["custo"] = f["quantidade"] * f["valor_unitario"]
    custo_ficha = f.groupby("produto")["custo"].sum().to_dict()

    df = cat.copy()
    custos, origens = [], []
    for _, r in df.iterrows():
        if r["categoria"] != "Revenda" and r["descricao"] in custo_ficha:
            custos.append(custo_ficha[r["descricao"]])
            origens.append("Ficha técnica")
        else:
            custos.append(float(r["custo_aquisicao"]))
            origens.append("Valor de aquisição" if r["categoria"] == "Revenda" else "Custo informado")
    df["custo_unit"] = custos
    df["origem_custo"] = origens
    df["margem_unit"] = df["preco_venda"] - df["custo_unit"]
    df["margem_pct"] = [div(m, p) for m, p in zip(df["margem_unit"], df["preco_venda"])]
    df["markup"] = [div(p, c) for p, c in zip(df["preco_venda"], df["custo_unit"])]
    df["receita_mensal"] = df["preco_venda"] * df["qtd_mensal"]
    df["custo_mensal"] = df["custo_unit"] * df["qtd_mensal"]
    return df


# ------------------------------------------------------------------------------
# 2. CRONOGRAMA DOS EMPRÉSTIMOS (TABELA PRICE, PARCELAS MENSAIS)
# ------------------------------------------------------------------------------
def parcelas_emprestimo(pv: float, taxa_anual_pct: float, n: int, carencia: int = 0, sistema: str = "Price",
                        meses: int | None = None) -> pd.DataFrame:
    """
    Tabela mês a mês de um empréstimo.
    Price: parcelas iguais após a carência. SAC: amortização constante (parcelas decrescentes).
    Carência: nos primeiros meses paga-se só os juros; o prazo total inclui a carência.
    """
    n = max(int(round(n)), 1)
    carencia = 0 if carencia is None or carencia != carencia else carencia
    sistema = "Price" if not isinstance(sistema, str) or not sistema.strip() else sistema
    carencia = int(min(max(round(carencia), 0), n - 1))
    i = (1 + float(taxa_anual_pct or 0) / 100.0) ** (1 / 12) - 1
    na = n - carencia
    sac = str(sistema or "Price").upper().startswith("SAC")
    pmt = (pv / na if i == 0 else pv * i / (1 - (1 + i) ** (-na))) if not sac else 0.0
    amort_sac = pv / na
    saldo = pv
    linhas = []
    for m in range(meses if meses is not None else n):
        if m < n and saldo > 1e-9:
            j = saldo * i
            if m < carencia:
                a = 0.0
            elif sac:
                a = min(amort_sac, saldo)
            else:
                a = min(pmt - j, saldo)
            saldo -= a
        else:
            j = a = 0.0
        linhas.append((m + 1, j, a, j + a, max(saldo, 0.0)))
    return pd.DataFrame(linhas, columns=["Mês", "Juros", "Amortização", "Parcela", "Saldo devedor"])


def cronograma_emprestimos(emp: pd.DataFrame, anos: int = 4) -> dict:
    juros = [0.0] * anos
    amort = [0.0] * anos
    saldo_fim = [0.0] * anos
    juros_m = [0.0] * (anos * 12)
    amort_m = [0.0] * (anos * 12)
    for _, r in emp.iterrows():
        pv, n = float(r["valor"]), int(round(r["prazo_meses"]))
        if pv <= 0:
            continue
        tab = parcelas_emprestimo(pv, r["taxa_anual_pct"], n, r.get("carencia_meses", 0) or 0, r.get("sistema", "Price") or "Price",
                                  meses=anos * 12)
        for m, (j, a, s) in enumerate(zip(tab["Juros"], tab["Amortização"], tab["Saldo devedor"])):
            ano = m // 12
            juros[ano] += j
            amort[ano] += a
            juros_m[m] += j
            amort_m[m] += a
            if m % 12 == 11:
                saldo_fim[ano] += s
    return {"juros": juros, "amort": amort, "saldo_fim": saldo_fim, "juros_m": juros_m, "amort_m": amort_m}


# ------------------------------------------------------------------------------
# 3. TIR (bisseção)
# ------------------------------------------------------------------------------
def calcular_tir(fluxos):
    if not fluxos or all(f >= 0 for f in fluxos) or all(f <= 0 for f in fluxos):
        return None
    vpl = lambda r: sum(f / (1 + r) ** t for t, f in enumerate(fluxos))
    lo, hi = -0.99, 10.0
    if vpl(lo) * vpl(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if vpl(lo) * vpl(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


# ------------------------------------------------------------------------------
# 3b. RECEITA RECORRENTE (ASSINATURAS) POR SAFRAS DE CLIENTES
# ------------------------------------------------------------------------------
COLS_ASS = ["Mês", "Novos", "Cancelamentos", "Ativos", "Receita reconhecida", "Cobranças (caixa)",
            "Custo de entrega", "Taxas de cobrança"]


def simular_assinaturas(ass: pd.DataFrame | None, g_a: float, inf: float, meses: int = 36):
    """
    Cada mês de entrada forma uma safra (coorte) de assinantes. A safra encolhe a cada renovação
    conforme o churn; quem paga por ciclo (trimestral, anual...) só pode sair na renovação.
    - Cobrança (caixa): no início de cada ciclo, preço do ciclo x assinantes ativos.
    - Receita reconhecida: preço do ciclo / meses do ciclo, todos os meses em que o assinante está ativo.
    - Custo de entrega: custo mensal por assinante ativo (nuvem, APIs, suporte), corrigido pela inflação.
    - Taxas de cobrança: % sobre o valor cobrado (meio de pagamento, loja de aplicativos).
    - Novos assinantes: crescem à taxa mensal do plano no Ano 1; nos Anos 2 e 3 partem do Mês 12
      e crescem à taxa anual de crescimento do projeto.
    """
    vazio = pd.DataFrame(columns=COLS_ASS), pd.DataFrame(columns=["Mês", "Plano", "Ativos", "Receita reconhecida"])
    if ass is None or ass.empty:
        return vazio
    tot = {c: [0.0] * meses for c in COLS_ASS[1:]}
    por_plano = []
    for _, p in ass.iterrows():
        preco, c = float(p["preco_ciclo"] or 0), max(1, int(round(float(p["ciclo_meses"] or 1))))
        n1, g = float(p["novos_mes1"] or 0), float(p["crescimento_novos_pct"] or 0) / 100
        churn_m = min(max(float(p["churn_mensal_pct"] or 0) / 100, 0.0), 1.0)
        sobra_ciclo = (1 - churn_m) ** c              # fração que renova a cada ciclo
        custo, taxa = float(p["custo_entrega_mensal"] or 0), float(p["taxa_cobranca_pct"] or 0) / 100
        if preco <= 0 and n1 <= 0:
            continue
        novos = [n1 * (1 + g) ** t if t < 12 else n1 * (1 + g) ** 11 * (1 + g_a) ** (t // 12) for t in range(meses)]
        ativos, rec, cob, canc = [0.0] * meses, [0.0] * meses, [0.0] * meses, [0.0] * meses
        for t0 in range(meses):
            for t in range(t0, meses):
                a = t - t0
                ren = a // c
                vivos = novos[t0] * sobra_ciclo ** ren
                ativos[t] += vivos
                rec[t] += vivos * preco / c
                if a % c == 0:
                    cob[t] += vivos * preco
                    if a > 0:
                        canc[t] += novos[t0] * sobra_ciclo ** (ren - 1) - vivos
        for t in range(meses):
            tot["Novos"][t] += novos[t]
            tot["Cancelamentos"][t] += canc[t]
            tot["Ativos"][t] += ativos[t]
            tot["Receita reconhecida"][t] += rec[t]
            tot["Cobranças (caixa)"][t] += cob[t]
            tot["Custo de entrega"][t] += ativos[t] * custo * (1 + inf) ** (t // 12)
            tot["Taxas de cobrança"][t] += cob[t] * taxa
            por_plano.append({"Mês": t + 1, "Plano": p["plano"], "Ativos": ativos[t], "Receita reconhecida": rec[t]})
    if not por_plano:
        return vazio
    df = pd.DataFrame({"Mês": list(range(1, meses + 1)), **tot})
    return df, pd.DataFrame(por_plano)


def metricas_recorrencia(s: pd.DataFrame, planos: pd.DataFrame, ass: pd.DataFrame, gasto_aquisicao_m1: float) -> dict:
    """Indicadores de negócios por assinatura (Ano 1)."""
    if s is None or s.empty or float(s["Receita reconhecida"].sum()) <= 0:
        return {"tem": False}
    a1 = s.iloc[:12]
    mrr1, mrr12 = float(s["Receita reconhecida"].iloc[0]), float(s["Receita reconhecida"].iloc[11])
    ativos12 = float(s["Ativos"].iloc[11])
    arpu = div(mrr12, ativos12)
    # churn mensal médio: churn informado em cada plano, ponderado pelos assinantes ativos no Mês 12
    p12 = planos[planos["Mês"] == 12].groupby("Plano")["Ativos"].sum()
    taxas = dict(zip(ass["plano"], ass["churn_mensal_pct"] / 100))
    churn = div(sum(p12.get(k, 0.0) * taxas.get(k, 0.0) for k in p12.index), float(p12.sum())) if p12.sum() > 0 else 0.0
    rec_a1 = float(a1["Receita reconhecida"].sum())
    margem_bruta = div(rec_a1 - float(a1["Custo de entrega"].sum()) - float(a1["Taxas de cobrança"].sum()), rec_a1)
    lucro_bruto_cliente = arpu * margem_bruta
    ltv = div(lucro_bruto_cliente, churn) if churn and churn > 0 else float("nan")
    vida_meses = div(1, churn) if churn and churn > 0 else float("nan")
    novos1 = float(s["Novos"].iloc[0])
    cac = div(gasto_aquisicao_m1, novos1) if novos1 > 0 else float("nan")
    ltv_cac = div(ltv, cac) if cac == cac and cac and cac > 0 and ltv == ltv else float("nan")
    payback_cac = div(cac, lucro_bruto_cliente) if lucro_bruto_cliente > 0 and cac == cac else float("nan")
    diferida = float(a1["Cobranças (caixa)"].sum()) - rec_a1
    return {"tem": True, "mrr1": mrr1, "mrr12": mrr12, "arr12": mrr12 * 12, "ativos12": ativos12, "arpu": arpu,
            "churn": churn, "vida_meses": vida_meses, "margem_bruta": margem_bruta, "ltv": ltv, "cac": cac,
            "gasto_aquisicao_m1": gasto_aquisicao_m1, "novos1": novos1, "ltv_cac": ltv_cac,
            "payback_cac": payback_cac, "receita_diferida_a1": diferida, "rec_a1": rec_a1,
            "cobrancas_a1": float(a1["Cobranças (caixa)"].sum()),
            "crescimento_mrr": div(mrr12, mrr1) - 1 if mrr1 > 0 else float("nan")}


# ------------------------------------------------------------------------------
# 3c. SIMPLES NACIONAL (Anexos I a V, LC 123/2006 com redação da LC 155/2016)
# ------------------------------------------------------------------------------
TABELAS_SIMPLES = {
    "Anexo I - Comércio": [(180000, .04, 0), (360000, .073, 5940), (720000, .095, 13860), (1800000, .107, 22500),
                           (3600000, .143, 87300), (4800000, .19, 378000)],
    "Anexo II - Indústria": [(180000, .045, 0), (360000, .078, 5940), (720000, .10, 13860), (1800000, .112, 22500),
                             (3600000, .147, 85500), (4800000, .30, 720000)],
    "Anexo III - Serviços": [(180000, .06, 0), (360000, .112, 9360), (720000, .135, 17640), (1800000, .16, 35640),
                             (3600000, .21, 125640), (4800000, .33, 648000)],
    "Anexo IV - Serviços": [(180000, .045, 0), (360000, .09, 8100), (720000, .102, 12420), (1800000, .14, 39780),
                            (3600000, .22, 183780), (4800000, .33, 828000)],
    "Anexo V - Serviços": [(180000, .155, 0), (360000, .18, 4500), (720000, .195, 9900), (1800000, .205, 17100),
                           (3600000, .23, 62100), (4800000, .305, 540000)],
}


def aliquota_simples(anexo: str, rbt12: float) -> dict:
    """Alíquota efetiva = (RBT12 x alíquota nominal - parcela a deduzir) / RBT12."""
    tab = TABELAS_SIMPLES.get(anexo)
    if not tab or rbt12 <= 0:
        return {"faixa": 0, "nominal": float("nan"), "deduzir": 0.0, "efetiva": float("nan")}
    for i, (lim, nom, ded) in enumerate(tab, start=1):
        if rbt12 <= lim:
            return {"faixa": i, "nominal": nom, "deduzir": ded, "efetiva": (rbt12 * nom - ded) / rbt12}
    lim, nom, ded = tab[-1]
    return {"faixa": 7, "nominal": nom, "deduzir": ded, "efetiva": (rbt12 * nom - ded) / rbt12, "acima_limite": True}


def fator_r(res: dict, d: dict) -> dict:
    """Fator R = folha de salários dos últimos 12 meses (com pró-labore e encargos) ÷ receita bruta."""
    cf = d["custos_fixos"]
    pro = float(cf.loc[cf["descricao"].astype(str).str.lower().str.replace("-", "").str.replace(" ", "")
                       .str.contains("prolabore|prólabore"), "valor_mensal"].sum())
    folha12 = res.get("folha_ano1", (res["salarios_m"] + res["encargos_m"]) * 12) + pro * 12
    rec = res["anos"][1]["rec_bruta"]
    fr = div(folha12, rec) if rec > 0 else float("nan")
    return {"folha12": folha12, "pro_labore_m": pro, "receita12": rec, "fator_r": fr,
            "anexo_por_fator_r": "Anexo III - Serviços" if fr == fr and fr >= 0.28 else "Anexo V - Serviços"}


# ------------------------------------------------------------------------------
# 4. MOTOR PRINCIPAL
# ------------------------------------------------------------------------------
def fatores_sazonais(tab, mes_inicio, g_m: float = 0.0) -> list:
    """
    Fatores de sazonalidade dos 12 meses do Ano 1 (v8.6). A tabela traz um índice por mês do calendário
    (100 = mês médio); o plano começa no mês de abertura. Os fatores são ajustados para que a soma do
    ano, já com o crescimento mensal, seja igual à de um ano sem sazonalidade: a DRE anual não muda,
    só a distribuição dos meses e o fluxo de caixa.
    """
    if tab is None or len(tab) == 0:
        return [1.0] * 12
    idx = {int(m): float(v) for m, v in zip(tab["mes"], tab["indice"]) if 1 <= int(m or 0) <= 12 and float(v or 0) > 0}
    if len(idx) < 12 or max(idx.values()) - min(idx.values()) < 1e-9:
        return [1.0] * 12
    ini = int(mes_inicio or 1) if 1 <= int(mes_inicio or 1) <= 12 else 1
    bruto = [idx[(ini - 1 + m) % 12 + 1] for m in range(12)]
    cresc = [(1 + g_m) ** m for m in range(12)]
    k = sum(cresc) / sum(c * b for c, b in zip(cresc, bruto))
    return [b * k for b in bruto]


def executar_motor(proj: dict, d: dict) -> dict:
    """d = dicionário com os DataFrames do projeto (nomes das tabelas)."""
    cat = custos_unitarios(d["catalogo"], d["materiais"], d["ficha_tecnica"])
    aliq = proj["aliquota_imposto_pct"] / 100.0
    var_pct = float(d["custos_variaveis"]["percentual"].sum()) / 100.0

    # --- Receita e custos diretos do Mês 1, por categoria
    receita_m1 = float(cat["receita_mensal"].sum())
    cdir_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "custo_mensal"].sum()) for c in CATEGORIAS}
    rec_cat_m1 = {c: float(cat.loc[cat["categoria"] == c, "receita_mensal"].sum()) for c in CATEGORIAS}
    cdir_m1 = sum(cdir_cat_m1.values())

    # --- Horizonte de planejamento escolhido pelo usuário (1 a 10 anos)
    N = int(min(max(round(float(proj.get("periodo_anos") or 3)), 1), 10))
    ANOS = list(range(1, N + 1))
    MESES = 12 * N

    # --- Assinaturas (receita recorrente), todos os meses do horizonte
    g_m = proj["crescimento_mensal_ano1_pct"] / 100
    g_a = proj["crescimento_anual_pct"] / 100
    inf = proj["inflacao_custos_anual_pct"] / 100
    ass_df, ass_planos = simular_assinaturas(d.get("assinaturas"), g_a, inf, meses=MESES)
    tem_ass = not ass_df.empty

    def serie(col):
        return [float(v) for v in ass_df[col]] if tem_ass else [0.0] * MESES
    s_rec, s_cob, s_custo, s_taxa = (serie("Receita reconhecida"), serie("Cobranças (caixa)"),
                                     serie("Custo de entrega"), serie("Taxas de cobrança"))

    # --- Pessoal, custos fixos e depreciação (mensais)
    pes_todos = d["pessoal"]
    inicio = (pes_todos["mes_inicio"] if "mes_inicio" in pes_todos else pd.Series([1.0] * len(pes_todos))).fillna(1).clip(lower=1)
    linha_total = (pes_todos["num_empregados"] * pes_todos["salario_mensal"] * (1 + pes_todos["encargos_pct"] / 100) +
                   pes_todos["num_empregados"] * pes_todos["beneficios_unit"])
    linha_sal_enc = pes_todos["num_empregados"] * pes_todos["salario_mensal"] * (1 + pes_todos["encargos_pct"] / 100)
    # custo de pessoal mês a mês (sem inflação): cada cargo entra no seu mês de início
    pes_serie = [float(linha_total[inicio <= t + 1].sum()) for t in range(MESES)]
    sal_enc_serie = [float(linha_sal_enc[inicio <= t + 1].sum()) for t in range(MESES)]
    pes = pes_todos[inicio <= 1]                                    # equipe do Mês 1 (custeio unitário e custo da hora)
    sal = pes["num_empregados"] * pes["salario_mensal"]
    salarios_m = float(sal.sum())
    encargos_m = float((sal * pes["encargos_pct"] / 100).sum())
    beneficios_m = float((pes["num_empregados"] * pes["beneficios_unit"]).sum())
    pessoal_m = salarios_m + encargos_m + beneficios_m
    fixos_m = float(d["custos_fixos"]["valor_mensal"].sum())
    pes_total_linha = sal * (1 + pes["encargos_pct"] / 100) + pes["num_empregados"] * pes["beneficios_unit"]
    area = pes["area"].fillna("") if "area" in pes else pd.Series([""] * len(pes))
    mod_m = float(pes_total_linha[area.str.startswith("Produ")].sum())          # mão de obra direta (produção)
    pessoal_adm_m = float(pes_total_linha[area.str.startswith("Admin")].sum())
    pessoal_com_m = pessoal_m - mod_m - pessoal_adm_m
    cfx = d["custos_fixos"]
    cls = cfx["classificacao"].fillna("") if "classificacao" in cfx else pd.Series([""] * len(cfx))
    cif_m = float(cfx.loc[cls.str.startswith("Custo"), "valor_mensal"].sum())   # indiretos de produção e de tecnologia
    desp_adm_m = float(cfx.loc[cls.str.contains("administrativa") | (cls == ""), "valor_mensal"].sum())
    desp_com_m = float(cfx.loc[cls.str.contains("comercial"), "valor_mensal"].sum())
    desp_fin_m = float(cfx.loc[cls.str.contains("financeira"), "valor_mensal"].sum())

    inv = d["inv_fixos"].copy()
    inv["total"] = inv["quantidade"] * inv["valor_unitario"]
    inv["dep_anual"] = [
        (t * (1 - res / 100) / v) if v > 0 else 0.0
        for t, res, v in zip(inv["total"], inv["valor_residual_pct"], inv["vida_util_anos"])
    ]
    inv["dep_mensal"] = inv["dep_anual"] / 12
    dep_anual = float(inv["dep_anual"].sum())
    dep_m = dep_anual / 12
    cf_desemb_m = pessoal_m + fixos_m

    # --- Investimentos
    inv_fixo = float(inv["total"].sum())
    inv_pre = float(d["inv_pre"]["valor"].sum())
    mat = d["materiais"]
    est_mat = float((mat["qtd_estoque_inicial"] * mat["valor_unitario"]).sum())
    rev = cat[cat["categoria"] == "Revenda"]
    est_rev = float((rev["qtd_estoque_inicial"] * rev["custo_unit"]).sum())
    inv_estoque = est_mat + est_rev

    # --- Prazos médios e capital de giro
    pz = d["capital_giro_prazos"]
    rec_pz, pag_pz = pz[pz["tipo"] == "Recebimento"], pz[pz["tipo"] == "Pagamento"]
    pmr = float((rec_pz["prazo_dias"] * rec_pz["percentual"] / 100).sum())
    pmp = float((pag_pz["prazo_dias"] * pag_pz["percentual"] / 100).sum())
    pme = float(proj["pme_dias"])
    ciclo_op = pme + pmr
    ciclo_fin = ciclo_op - pmp
    desembolso_m1 = (cdir_m1 + (receita_m1 + s_rec[0]) * (aliq + var_pct) + cf_desemb_m + s_custo[0] + s_taxa[0])
    ncg = max(0.0, desembolso_m1 / 30.0 * ciclo_fin)
    reserva = float(proj["reserva_caixa_extra"])
    capital_giro = ncg + reserva
    investimento_total = inv_fixo + inv_pre + inv_estoque + capital_giro

    # --- Fontes
    capital_proprio = float(d["socios"]["capital_integralizado"].sum())
    capital_terceiros = float(d["emprestimos"]["valor"].sum())
    fontes = capital_proprio + capital_terceiros
    crono = cronograma_emprestimos(d["emprestimos"], anos=N + 1)

    # --- Projeção mensal do Ano 1 (com sazonalidade, se informada; o total do ano não muda)
    tma = proj["tma_anual_pct"] / 100
    saz = fatores_sazonais(d.get("sazonalidade"), proj.get("saz_mes_inicio"), g_m)
    meses = []
    for m in range(12):
        f = (1 + g_m) ** m * saz[m]
        rec = receita_m1 * f + s_rec[m]
        cv = cdir_m1 * f + rec * (aliq + var_pct) + s_custo[m] + s_taxa[m]
        cf = pes_serie[m] + fixos_m + dep_m
        meses.append({"Mês": f"Mês {m + 1}", "Receita": rec, "Custos variáveis": cv,
                      "Margem de contribuição": rec - cv, "Custos fixos": cf,
                      "Resultado operacional": rec - cv - cf})
    df_meses = pd.DataFrame(meses)
    fator_a1 = sum((1 + g_m) ** m for m in range(12))

    # --- DRE e balanço dos Anos 1 a 3
    anos = {}
    lucros_acum = 0.0
    for k in ANOS:
        fr = fator_a1 * (1 + g_a) ** (k - 1)
        fi = (1 + inf) ** (k - 1)
        meses_k = slice((k - 1) * 12, k * 12)
        ass_rec, ass_custo, ass_taxa = sum(s_rec[meses_k]), sum(s_custo[meses_k]), sum(s_taxa[meses_k])
        rec_prod = receita_m1 * fr
        rec_bruta = rec_prod + ass_rec
        rec_cat = {c: v * fr for c, v in rec_cat_m1.items()}
        rec_cat["Assinatura"] = ass_rec
        cmv, cpv, csv = (cdir_cat_m1["Revenda"] * fr, cdir_cat_m1["Fabricação Própria"] * fr,
                         cdir_cat_m1["Serviço"] * fr + ass_custo)
        custo_dir = cmv + cpv + csv
        impostos = rec_bruta * aliq
        outras_var = rec_bruta * var_pct + ass_taxa
        rec_liquida = rec_bruta - impostos
        cv_tot = custo_dir + impostos + outras_var
        mc = rec_bruta - cv_tot
        imc = div(mc, rec_bruta)
        pessoal = sum(pes_serie[meses_k]) * fi
        fixos = fixos_m * 12 * fi
        cf_desemb = pessoal + fixos
        ebitda = mc - cf_desemb
        ebit = ebitda - dep_anual
        juros = crono["juros"][k - 1]
        resultado = ebit - juros
        lucros_acum += resultado
        fco = resultado + dep_anual
        amort = crono["amort"][k - 1]
        fluxo_socio = fco - amort

        # Balanço simplificado (caixa calculado por diferença: Ativo = Passivo + PL)
        receber = rec_bruta / 360 * pmr
        estoques = inv_estoque * (rec_prod / (receita_m1 * fator_a1)) if receita_m1 > 0 else inv_estoque
        imobilizado = max(0.0, inv_fixo - dep_anual * k)
        fornecedores = custo_dir / 360 * pmp
        obrigacoes = (pessoal + impostos) / 12
        saldo_div = crono["saldo_fim"][k - 1]
        div_cp = crono["amort"][k] if k < len(crono["amort"]) else 0.0
        div_cp = min(div_cp, saldo_div)
        div_lp = saldo_div - div_cp
        # Receita antecipada (assinaturas cobradas e ainda não prestadas): passivo circulante
        receita_diferida = sum(s_cob[:k * 12]) - sum(s_rec[:k * 12])
        pc = fornecedores + obrigacoes + div_cp + receita_diferida
        pnc = div_lp
        pl = capital_proprio + lucros_acum
        caixa = pc + pnc + pl - (receber + estoques + imobilizado + inv_pre)
        ac = caixa + receber + estoques
        ativo_total = ac + imobilizado + inv_pre

        cf_total = cf_desemb + dep_anual
        pe_cont = div(cf_total, imc) if imc and imc > 0 else float("nan")
        pe_fin = div(cf_desemb, imc) if imc and imc > 0 else float("nan")
        pe_eco = div(cf_total + investimento_total * tma, imc) if imc and imc > 0 else float("nan")

        anos[k] = dict(
            rec_bruta=rec_bruta, rec_cat=rec_cat, impostos=impostos, rec_liquida=rec_liquida,
            cmv=cmv, cpv=cpv, csv=csv, custo_dir=custo_dir, outras_var=outras_var,
            cv_tot=cv_tot, mc=mc, imc=imc, pessoal=pessoal, fixos=fixos, cf_desemb=cf_desemb,
            dep=dep_anual, cf_total=cf_total, ebitda=ebitda, ebit=ebit, juros=juros,
            resultado=resultado, fco=fco, amort=amort, fluxo_socio=fluxo_socio,
            caixa=caixa, receber=receber, estoques=estoques, ac=ac, imobilizado=imobilizado,
            ativo_total=ativo_total, fornecedores=fornecedores, obrigacoes=obrigacoes, receita_diferida=receita_diferida,
            div_cp=div_cp, div_lp=div_lp, pc=pc, pnc=pnc, pl=pl, divida_bruta=saldo_div,
            divida_liquida=saldo_div - caixa, pe_contabil=pe_cont, pe_financeiro=pe_fin,
            pe_economico=pe_eco,
        )

    # --- Avaliação do projeto (fluxo do projeto = EBITDA, investimento no ano 0)
    fluxos = [-investimento_total] + [anos[k]["ebitda"] for k in ANOS]
    if proj.get("recuperar_giro"):
        fluxos[N] += capital_giro
    vpl = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos))
    tir = calcular_tir(fluxos)
    vp_entradas = sum(f / (1 + tma) ** t for t, f in enumerate(fluxos) if t > 0)
    il = div(vp_entradas, investimento_total)

    lucro_mensal_a1 = anos[1]["resultado"] / 12
    payback_simples = div(investimento_total, lucro_mensal_a1) if lucro_mensal_a1 > 0 else float("nan")
    acum, payback_desc = -investimento_total, float("nan")
    for t in ANOS:
        fd = fluxos[t] / (1 + tma) ** t
        if fd > 0 and acum + fd >= 0:
            payback_desc = ((t - 1) + (-acum) / fd) * 12
            break
        acum += fd

    a1 = anos[1]
    resumo = {
        "receita_anual": a1["rec_bruta"], "custos_variaveis": a1["cv_tot"], "margem_contrib": a1["mc"],
        "custos_fixos": a1["cf_total"], "resultado_operacional": a1["mc"] - a1["cf_total"],
        "lucratividade": div(a1["mc"] - a1["cf_total"], a1["rec_bruta"]),
        "rentabilidade": div(a1["mc"] - a1["cf_total"], investimento_total),
        "payback_meses": div(investimento_total, (a1["mc"] - a1["cf_total"]) / 12)
        if a1["mc"] - a1["cf_total"] > 0 else float("nan"),
    }

    # --- Custeio por absorção: MOD + custos indiretos de produção + depreciação rateados
    overhead_m = mod_m + cif_m + dep_m
    criterio = proj.get("rateio_criterio") or "Custo direto"
    prod_mask = cat["categoria"] != "Revenda"
    base = {"Custo direto": cat["custo_mensal"], "Receita": cat["receita_mensal"],
            "Quantidade": cat["qtd_mensal"]}.get(criterio, cat["custo_mensal"]).where(prod_mask, 0.0)
    tot_base = float(base.sum())
    cat["rateio_mensal"] = base / tot_base * overhead_m if tot_base > 0 else 0.0
    cat["custo_absorcao_unit"] = [cu + (r / q if q > 0 else 0.0) for cu, r, q in
                                  zip(cat["custo_unit"], cat["rateio_mensal"], cat["qtd_mensal"])]
    cat["margem_absorcao_pct"] = [div(p - c, p) for p, c in zip(cat["preco_venda"], cat["custo_absorcao_unit"])]

    # --- Formação de preço (markup divisor)
    margem_desej = float(proj.get("margem_desejada_pct") or 0) / 100
    df_pct = div(cf_desemb_m + dep_m, receita_m1 + s_rec[0]) if receita_m1 + s_rec[0] > 0 else 0.0
    df_pct = 0.0 if df_pct != df_pct else df_pct
    divisor = 1 - (aliq + var_pct + df_pct + margem_desej)
    markup_mult = div(1, divisor) if divisor > 0 else float("nan")
    cat["preco_sugerido"] = [c * markup_mult if divisor > 0 else float("nan") for c in cat["custo_unit"]]
    cat["dif_preco_pct"] = [div(p - s_, s_) for p, s_ in zip(cat["preco_venda"], cat["preco_sugerido"])]

    # --- Ponto de equilíbrio em unidades por produto (mix atual de vendas)
    pe_m = anos[1]["pe_contabil"] / 12 if anos[1]["pe_contabil"] == anos[1]["pe_contabil"] else float("nan")
    cat["pe_unidades_mes"] = [div(pe_m * div(r, receita_m1), p) if receita_m1 > 0 else float("nan")
                              for r, p in zip(cat["receita_mensal"], cat["preco_venda"])]

    # --- Fluxo de caixa mensal do Ano 1
    def distribuir(valores, prazos_df):
        """Distribui cada valor mensal no tempo conforme os prazos (dias) e percentuais."""
        saida = [0.0] * 14
        regras = [(float(p), float(q) / 100) for p, q in zip(prazos_df["prazo_dias"], prazos_df["percentual"])] \
            or [(0.0, 1.0)]
        for m, v in enumerate(valores):
            for dias, frac in regras:
                lag, resto = int(dias // 30), (dias % 30) / 30
                if m + lag < 14:
                    saida[m + lag] += v * frac * (1 - resto)
                if m + lag + 1 < 14:
                    saida[m + lag + 1] += v * frac * resto
        return saida[:12]

    rec_mes = [receita_m1 * (1 + g_m) ** m * saz[m] for m in range(12)]
    cdir_mes = [cdir_m1 * (1 + g_m) ** m * saz[m] for m in range(12)]
    entradas = [e + c for e, c in zip(distribuir(rec_mes, rec_pz), s_cob[:12])]   # assinaturas: caixa na cobrança
    pag_forn = distribuir(cdir_mes, pag_pz)
    saldo = fontes - inv_fixo - inv_pre - inv_estoque
    fluxo_m = []
    for m in range(12):
        sai_var = (rec_mes[m] + s_rec[m]) * (aliq + var_pct) + s_custo[m] + s_taxa[m]
        parcela = crono["juros_m"][m] + crono["amort_m"][m]
        sai = pag_forn[m] + sai_var + pes_serie[m] + fixos_m + parcela
        ini = saldo
        saldo = saldo + entradas[m] - sai
        fluxo_m.append({"Mês": f"Mês {m + 1}", "Saldo inicial": ini, "Recebimentos": entradas[m],
                        "Fornecedores": pag_forn[m], "Impostos e desp. variáveis": sai_var,
                        "Pessoal e custos fixos": pes_serie[m] + fixos_m, "Parcelas de empréstimos": parcela,
                        "Saldo final": saldo})
    df_fluxo_m = pd.DataFrame(fluxo_m)

    # --- Indicadores de recorrência (CAC = gasto comercial fixo do Mês 1 ÷ novos assinantes do Mês 1)
    verba = float(proj.get("verba_aquisicao_mensal") or 0)
    rec_ind = metricas_recorrencia(ass_df, ass_planos, d.get("assinaturas"), verba if verba > 0 else desp_com_m + pessoal_com_m)
    if rec_ind.get("tem"):
        rec_ind["origem_cac"] = "verba de aquisição informada" if verba > 0 else "gasto comercial fixo do Mês 1"

    return dict(
        mod_m=mod_m, pessoal_adm_m=pessoal_adm_m, pessoal_com_m=pessoal_com_m, cif_m=cif_m,
        desp_adm_m=desp_adm_m, desp_com_m=desp_com_m, desp_fin_m=desp_fin_m, overhead_m=overhead_m,
        criterio_rateio=criterio, margem_desejada=margem_desej, df_pct=df_pct, markup_mult=markup_mult,
        divisor_markup=divisor, fluxo_mensal=df_fluxo_m, g_m=g_m,
        cat=cat, inv=inv, receita_m1=receita_m1, cdir_m1=cdir_m1, var_pct=var_pct, aliq=aliq,
        salarios_m=salarios_m, encargos_m=encargos_m, beneficios_m=beneficios_m, pessoal_m=pessoal_m,
        fixos_m=fixos_m, dep_m=dep_m, dep_anual=dep_anual, inv_fixo=inv_fixo, inv_pre=inv_pre,
        est_mat=est_mat, est_rev=est_rev, inv_estoque=inv_estoque, pmr=pmr, pmp=pmp, pme=pme,
        ciclo_op=ciclo_op, ciclo_fin=ciclo_fin, desembolso_m1=desembolso_m1, ncg=ncg,
        reserva=reserva, capital_giro=capital_giro, investimento_total=investimento_total,
        capital_proprio=capital_proprio, capital_terceiros=capital_terceiros, fontes=fontes,
        crono=crono, meses=df_meses, anos=anos, fluxos=fluxos, vpl=vpl, tir=tir, il=il,
        payback_simples=payback_simples, payback_desc=payback_desc, resumo=resumo, tma=tma,
        ass=ass_df, ass_planos=ass_planos, recorrencia=rec_ind, g_a=g_a, inf=inf,
        N=N, ANOS=ANOS, MESES=MESES, pes_serie=pes_serie, folha_ano1=sum(sal_enc_serie[:12]),
        recupera_giro=bool(proj.get("recuperar_giro")), saz=saz,
    )


# ------------------------------------------------------------------------------
# 5. CENÁRIOS (estilo "Simulador de Resultados" do PNBOX)
# ------------------------------------------------------------------------------
def cenarios(res: dict, proj: dict) -> pd.DataFrame:
    a1 = res["anos"][1]
    inv, tma = res["investimento_total"], res["tma"]

    def calc(fr, fc):
        rec, cv, cf = a1["rec_bruta"] * fr, a1["cv_tot"] * fc, a1["cf_total"] * fc
        mc = rec - cv
        r = mc - cf
        imc = div(mc, rec)
        ok = imc == imc and imc > 0
        return [brl(rec), brl(cv), brl(mc), brl(cf), brl(r), pct(div(r, rec)), pct(div(r, inv)),
                f"{num_br(div(inv, r / 12), 1)} meses" if r > 0 else "Não recupera",
                brl(cf / imc) if ok else "n.a.", brl((cf - a1["dep"] * fc) / imc) if ok else "n.a.",
                brl((cf + inv * tma) / imc) if ok else "n.a."]

    rotulos = ["Receita anual", "Custos variáveis", "Margem de contribuição", "Custos fixos",
               "Resultado operacional", "Lucratividade", "Rentabilidade", "Payback simples",
               "PE contábil", "PE financeiro", "PE econômico"]
    ro, co = proj["cen_otim_receita_pct"] / 100, proj["cen_otim_custo_pct"] / 100
    rp, cp = proj["cen_pess_receita_pct"] / 100, proj["cen_pess_custo_pct"] / 100
    return pd.DataFrame({
        "Indicador (Ano 1)": rotulos,
        f"Pessimista (-{num_br(rp*100,0)}% rec. / +{num_br(cp*100,0)}% custos)": calc(1 - rp, 1 + cp),
        "Provável": calc(1, 1),
        f"Otimista (+{num_br(ro*100,0)}% rec. / -{num_br(co*100,0)}% custos)": calc(1 + ro, 1 - co),
    })


# ------------------------------------------------------------------------------
# 6. DRE EM TABELA
# ------------------------------------------------------------------------------
def tabela_dre(res: dict) -> pd.DataFrame:
    A = res["anos"]
    linhas = [
        ("1. Receita operacional bruta", "rec_bruta", 1),
        ("   (-) Impostos sobre vendas", "impostos", -1),
        ("2. (=) Receita líquida", "rec_liquida", 1),
        ("   (-) CMV - mercadorias de revenda", "cmv", -1),
        ("   (-) CPV - produtos de fabricação própria", "cpv", -1),
        ("   (-) CSV - serviços vendidos", "csv", -1),
        ("   (-) Outras despesas variáveis", "outras_var", -1),
        ("3. (=) Margem de contribuição", "mc", 1),
        ("   (-) Pessoal (salários, encargos e benefícios)", "pessoal", -1),
        ("   (-) Demais custos fixos", "fixos", -1),
        ("4. (=) EBITDA", "ebitda", 1),
        ("   (-) Depreciação", "dep", -1),
        ("5. (=) Resultado operacional (EBIT)", "ebit", 1),
        ("   (-) Juros dos empréstimos", "juros", -1),
        ("6. (=) Resultado líquido do exercício", "resultado", 1),
    ]
    rows = []
    for nome, k, s in linhas:
        lin = {"Conta": nome}
        for a in res["ANOS"]:
            lin[f"Ano {a}"] = brl(A[a][k] * s)
        lin["% Receita (Ano 1)"] = pct(div(A[1][k], A[1]["rec_bruta"]))
        rows.append(lin)
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------------
# 7. INDICADORES EM 5 BLOCOS
# ------------------------------------------------------------------------------
def tabela_indicadores(res: dict) -> pd.DataFrame:
    A = res["anos"]
    inv = res["investimento_total"]
    linhas = []

    AN = res["ANOS"]

    def add(bloco, nome, formula, vals, fmt, leitura, var="pct"):
        f = {"brl": brl, "pct": pct, "x": lambda v: "n.a." if v != v else f"{num_br(v)}x"}[fmt]
        lin = {"Bloco": bloco, "Indicador": nome, "Fórmula": formula}
        for i, a in enumerate(AN):
            lin[f"Ano {a}"] = f(vals[i])
        if len(AN) >= 2:
            v1, v2 = vals[0], vals[1]
            if var == "pp":
                vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br((v2 - v1) * 100):s} p.p."
            elif var == "dif":
                vr = "n.a." if v1 != v1 or v2 != v2 else f"{num_br(v2 - v1)}x"
            else:
                vr = "n.a." if not v1 or v1 != v1 else f"{num_br((v2 - v1) / abs(v1) * 100, 1)}%"
            lin["Var. Ano 2/Ano 1"] = vr
        lin["Leitura"] = leitura
        linhas.append(lin)

    g = lambda k: [A[a][k] for a in AN]
    r = lambda n, d_: [div(A[a][n], A[a][d_]) for a in AN]

    b = "1. Rentabilidade e resultado"
    add(b, "Receita operacional bruta", "Σ preço × quantidade", g("rec_bruta"), "brl", "Faturamento antes das deduções.")
    add(b, "Margem de contribuição", "Receita − custos e despesas variáveis", g("mc"), "brl", "Valor que sobra para pagar os custos fixos.")
    add(b, "Índice de margem de contribuição", "MC ÷ Receita bruta", g("imc"), "pct", "Parcela de cada R$ 1,00 vendido que cobre a estrutura fixa.", "pp")
    add(b, "EBITDA", "MC − custos fixos desembolsáveis", g("ebitda"), "brl", "Geração operacional de caixa antes de juros e depreciação.")
    add(b, "Margem EBITDA", "EBITDA ÷ Receita líquida", r("ebitda", "rec_liquida"), "pct", "Eficiência operacional.", "pp")
    add(b, "Resultado líquido", "EBIT − juros", g("resultado"), "brl", "Lucro (ou prejuízo) do exercício.")
    add(b, "Lucratividade líquida", "Resultado líquido ÷ Receita bruta", r("resultado", "rec_bruta"), "pct", "Ganho para cada real faturado.", "pp")
    add(b, "Rentabilidade (ROI)", "Resultado líquido ÷ Investimento total", [div(A[a]["resultado"], inv) for a in AN], "pct", "Retorno contábil sobre o capital investido.", "pp")
    add(b, "ROE", "Resultado líquido ÷ Patrimônio líquido", r("resultado", "pl"), "pct", "Retorno sobre o capital dos sócios.", "pp")

    b = "2. Endividamento e estrutura de capital"
    add(b, "Dívida bruta", "Saldo dos empréstimos no fim do ano", g("divida_bruta"), "brl", "Obrigações financeiras onerosas.")
    add(b, "Dívida líquida", "Dívida bruta − caixa", g("divida_liquida"), "brl", "Negativa indica caixa maior que a dívida.")
    add(b, "Endividamento geral", "(PC + PNC) ÷ Ativo total", [div(A[a]["pc"] + A[a]["pnc"], A[a]["ativo_total"]) for a in AN], "pct", "Parcela do ativo financiada por terceiros.", "pp")
    add(b, "Dívida líquida ÷ EBITDA", "Dívida líquida ÷ EBITDA", [div(A[a]["divida_liquida"], A[a]["ebitda"]) if A[a]["ebitda"] > 0 else float("nan") for a in AN], "x", "Anos de geração operacional para quitar a dívida.", "dif")
    add(b, "Cobertura de juros", "EBITDA ÷ juros", [div(A[a]["ebitda"], A[a]["juros"]) if A[a]["juros"] > 0 else float("nan") for a in AN], "x", "Abaixo de 1,0x a operação não paga os juros.", "dif")

    b = "3. Liquidez e capital de giro"
    add(b, "Liquidez corrente", "AC ÷ PC", r("ac", "pc"), "x", "Acima de 1,0x indica folga de curto prazo.", "dif")
    add(b, "Liquidez seca", "(AC − estoques) ÷ PC", [div(A[a]["ac"] - A[a]["estoques"], A[a]["pc"]) for a in AN], "x", "Capacidade de pagamento sem vender estoques.", "dif")
    add(b, "Liquidez imediata", "Caixa ÷ PC", r("caixa", "pc"), "x", "Disponibilidade imediata.", "dif")
    add(b, "Capital circulante líquido", "AC − PC", [A[a]["ac"] - A[a]["pc"] for a in AN], "brl", "Folga (ou falta) de recursos de curto prazo.")
    add(b, "Caixa projetado", "Ativo − demais contas (por diferença)", g("caixa"), "brl", "Negativo indica necessidade de mais recursos.")

    b = "4. Viabilidade do investimento"
    add(b, "Custo com pessoal ÷ Receita líquida", "Folha total ÷ Receita líquida", r("pessoal", "rec_liquida"), "pct", "Peso da mão de obra.", "pp")
    add(b, "Fluxo de caixa operacional (FCO)", "Resultado líquido + depreciação", g("fco"), "brl", "Caixa gerado pelas operações.")
    add(b, "Fluxo de caixa do sócio", "FCO − amortização dos empréstimos", g("fluxo_socio"), "brl", "Caixa que sobra após pagar o principal das dívidas.")

    b = "5. Pontos de equilíbrio"
    add(b, "PE contábil (anual)", "Custos fixos totais ÷ IMC", g("pe_contabil"), "brl", "Receita para lucro zero.")
    add(b, "PE financeiro (anual)", "(Custos fixos − depreciação) ÷ IMC", g("pe_financeiro"), "brl", "Receita para cobrir os desembolsos.")
    add(b, "PE econômico (anual)", "(Custos fixos + TMA × investimento) ÷ IMC", g("pe_economico"), "brl", "Receita para remunerar o capital à TMA.")
    add(b, "Margem de segurança", "(Receita − PE contábil) ÷ Receita", [div(A[a]["rec_bruta"] - A[a]["pe_contabil"], A[a]["rec_bruta"]) for a in AN], "pct", "Quanto a receita pode cair antes do prejuízo.", "pp")
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 8. CHECAGENS DE CONSISTÊNCIA
# ------------------------------------------------------------------------------
def tabela_checagens(res: dict, d: dict) -> pd.DataFrame:
    cat = res["cat"]
    A = res["anos"]
    chk = []

    def add(t, valor, ok, obs):
        chk.append({"Teste": t, "Valor apurado": valor, "Situação": "OK" if ok else "VERIFICAR", "Observação": obs})

    dif = res["fontes"] - res["investimento_total"]
    add("Fontes de recursos x investimento total", f"Diferença: {brl(dif)}", abs(dif) <= 1,
        "Capital dos sócios + empréstimos deve igualar o investimento total.")
    neg = cat[cat["margem_unit"] <= 0]["descricao"].tolist()
    add("Margem unitária positiva", f"{len(neg)} item(ns) com margem <= 0", not neg,
        ("Itens: " + ", ".join(neg[:5])) if neg else "Todos os itens têm preço maior que o custo.")
    sem = cat[cat["custo_unit"] <= 0]["descricao"].tolist()
    add("Produtos com custo cadastrado", f"{len(sem)} item(ns) sem custo", not sem,
        ("Sem ficha técnica/custo: " + ", ".join(sem[:5])) if sem else "Todos os itens têm custo.")
    mats = set(d["materiais"]["nome"])
    orf = sorted(set(d["ficha_tecnica"]["material"]) - mats)
    add("Fichas técnicas com materiais válidos", f"{len(orf)} material(is) não cadastrado(s)", not orf,
        ", ".join(orf[:5]) if orf else "Todos os insumos existem no cadastro de materiais.")
    for tipo in ("Recebimento", "Pagamento"):
        s = d["capital_giro_prazos"].query("tipo == @tipo")["percentual"].sum()
        add(f"Prazos de {tipo.lower()} somam 100%", f"{num_br(s)}%", abs(s - 100) < 0.01,
            "A soma dos percentuais de cada tipo de prazo deve ser 100%.")
    sw = d["swot"]
    incoer = sw[((sw["ambiente"] == "Interno") & sw["natureza"].isin(["Oportunidade", "Ameaça"])) |
                ((sw["ambiente"] == "Externo") & sw["natureza"].isin(["Força", "Fraqueza"]))]
    add("SWOT: ambiente coerente com a classificação", f"{len(incoer)} fator(es) incoerente(s)", incoer.empty,
        "Interno = Força/Fraqueza; Externo = Oportunidade/Ameaça.")
    neg_cx = [k for k in res["ANOS"] if A[k]["caixa"] < 0]
    add("Caixa projetado não negativo", "Anos com caixa negativo: " + (", ".join(map(str, neg_cx)) or "nenhum"),
        not neg_cx, "Caixa negativo indica que as fontes não cobrem a operação.")
    lc = div(A[1]["ac"], A[1]["pc"])
    add("Liquidez corrente do Ano 1 >= 1,00", f"{num_br(lc)}x", lc == lc and lc >= 1, "Abaixo de 1,00 há aperto de curto prazo.")
    add("VPL positivo", brl(res["vpl"]), res["vpl"] > 0, "VPL positivo indica criação de valor acima da TMA.")
    return pd.DataFrame(chk)


# ------------------------------------------------------------------------------
# 9. DIAGNÓSTICO SWOT (matriz de estratégias)
# ------------------------------------------------------------------------------
QUADRANTES = {
    "FO": ("Estratégia de Desenvolvimento", "Forças + Oportunidades",
           "Cenário mais favorável: aproveitar oportunidades e expandir (novos mercados, produtos e parcerias)."),
    "FrO": ("Estratégia de Crescimento", "Fraquezas + Oportunidades",
            "Há oportunidades, mas fraquezas internas limitam o aproveitamento: corrigir pontos fracos para crescer."),
    "FA": ("Estratégia de Manutenção", "Forças + Ameaças",
           "Usar as forças para enfrentar as ameaças e manter a posição conquistada."),
    "FrA": ("Estratégia de Sobrevivência", "Fraquezas + Ameaças",
            "Situação mais delicada: reduzir custos, rever o modelo e proteger o caixa."),
}


def pontuacao_swot(sw: pd.DataFrame) -> dict:
    peso = lambda nat: int(sum(2 if i == "Alto" else 1 for i in sw.loc[sw["natureza"] == nat, "impacto"]))
    F, Fr, O, A = peso("Força"), peso("Fraqueza"), peso("Oportunidade"), peso("Ameaça")
    q = {"FO": F + O, "FrO": Fr + O, "FA": F + A, "FrA": Fr + A}
    pred = max(q, key=q.get) if any(q.values()) else None
    return {"F": F, "Fr": Fr, "O": O, "A": A, "quadrantes": q, "predominante": pred}


# ------------------------------------------------------------------------------
# 10. COMPOSIÇÕES PARA GRÁFICOS
# ------------------------------------------------------------------------------
def _serie(dic: dict) -> pd.DataFrame:
    df = pd.DataFrame({"Item": list(dic.keys()), "Valor": [float(v) for v in dic.values()]})
    df = df[df["Valor"] > 0.005].sort_values("Valor", ascending=False).reset_index(drop=True)
    tot = df["Valor"].sum()
    df["%"] = df["Valor"] / tot if tot else 0.0
    return df


def composicoes(res: dict, d: dict) -> dict:
    mat = d["materiais"]
    cat_mat = {n: (c or "Não classificado") for n, c in zip(mat["nome"], mat["categoria"])}
    preco_mat = dict(zip(mat["nome"], mat["valor_unitario"]))
    out = {}

    out["investimento"] = _serie({
        "Ativo imobilizado (investimentos fixos)": res["inv_fixo"], "Investimentos pré-operacionais": res["inv_pre"],
        "Estoque inicial": res["inv_estoque"], "Capital de giro": res["capital_giro"]})

    est = {}
    for _, r in mat.iterrows():
        k = r["categoria"] or "Não classificado"
        est[k] = est.get(k, 0) + r["qtd_estoque_inicial"] * r["valor_unitario"]
    est["Mercadorias para revenda"] = res["est_rev"]
    out["estoque"] = _serie(est)

    ficha = d["ficha_tecnica"]
    custo = {}
    for _, p in res["cat"].iterrows():
        if p["qtd_mensal"] <= 0:
            continue
        if p["origem_custo"] == "Ficha técnica":
            for _, f in ficha[ficha["produto"] == p["descricao"]].iterrows():
                k = cat_mat.get(f["material"], "Não classificado")
                custo[k] = custo.get(k, 0) + p["qtd_mensal"] * f["quantidade"] * preco_mat.get(f["material"], 0)
        elif p["categoria"] == "Revenda":
            custo["Mercadorias para revenda"] = custo.get("Mercadorias para revenda", 0) + p["custo_mensal"]
        else:
            custo["Custo informado (sem ficha técnica)"] = custo.get("Custo informado (sem ficha técnica)", 0) + p["custo_mensal"]
    out["custo_direto"] = _serie(custo)

    out["fontes"] = _serie({"Capital próprio (sócios)": res["capital_proprio"],
                            "Capital de terceiros (empréstimos)": res["capital_terceiros"]})
    out["receita_categoria"] = _serie(res["anos"][1]["rec_cat"])
    top = res["cat"].sort_values("receita_mensal", ascending=False).head(10)
    out["top_produtos"] = _serie(dict(zip(top["descricao"], top["receita_mensal"])))
    inv = res["inv"].sort_values("total", ascending=False)
    fixos = dict(zip(inv["descricao"].head(10), inv["total"].head(10)))
    resto = inv["total"].iloc[10:].sum()
    if resto > 0:
        fixos["Demais itens"] = resto
    out["inv_fixos"] = _serie(fixos)

    a1 = res["anos"][1]
    estrutura = {"CMV (revenda)": a1["cmv"], "CPV (fabricação)": a1["cpv"], "CSV (serviços)": a1["csv"],
                 "Impostos sobre vendas": a1["impostos"], "Outras despesas variáveis": a1["outras_var"],
                 "Pessoal": a1["pessoal"], "Demais custos fixos": a1["fixos"], "Depreciação": a1["dep"],
                 "Juros": a1["juros"], "Lucro líquido": max(a1["resultado"], 0)}
    df = pd.DataFrame({"Item": list(estrutura), "Valor": list(estrutura.values())})
    df["%"] = df["Valor"] / a1["rec_bruta"] if a1["rec_bruta"] else 0.0
    out["estrutura_receita"] = df[df["Valor"] > 0.005].reset_index(drop=True)

    # Gráfico do ponto de equilíbrio (receita x custo total, Ano 1)
    imc = a1["imc"] if a1["imc"] == a1["imc"] else 0.0
    pe = a1["pe_contabil"] if a1["pe_contabil"] == a1["pe_contabil"] else a1["rec_bruta"]
    xmax = max(a1["rec_bruta"], pe) * 1.3 or 1.0
    xs = [xmax * i / 20 for i in range(21)]
    out["equilibrio"] = pd.DataFrame({"Receita": xs, "Receita total": xs,
                                      "Custo total": [a1["cf_total"] + x * (1 - imc) for x in xs],
                                      "Custos fixos": [a1["cf_total"]] * len(xs)})
    out["pe_ponto"] = (pe, a1["rec_bruta"])

    fl = res["fluxos"]
    acum, acum_d, s, sd = [], [], 0.0, 0.0
    for t, f in enumerate(fl):
        s += f
        sd += f / (1 + res["tma"]) ** t
        acum.append(s)
        acum_d.append(sd)
    out["fluxo_acumulado"] = pd.DataFrame({"Ano": [f"Ano {t}" for t in range(len(fl))], "Fluxo do ano": fl,
                                           "Acumulado": acum, "Acumulado descontado (TMA)": acum_d})
    return out


# ------------------------------------------------------------------------------
# 11. SIMULADOR "E SE?" E ANÁLISE DE SENSIBILIDADE (Ano 1)
# ------------------------------------------------------------------------------
def simular(res: dict, preco=0.0, volume=0.0, materiais=0.0, salarios=0.0, fixos=0.0) -> dict:
    """Recalcula o resultado do Ano 1 aplicando variações percentuais (0.10 = +10%)."""
    a1 = res["anos"][1]
    fp, fv, fm, fs, ff = 1 + preco, 1 + volume, 1 + materiais, 1 + salarios, 1 + fixos
    rec = a1["rec_bruta"] * fp * fv
    cdir = a1["custo_dir"] * fv * fm
    var = rec * (res["aliq"] + res["var_pct"])
    mc = rec - cdir - var
    cf = a1["pessoal"] * fs + a1["fixos"] * ff + a1["dep"]
    resultado = mc - cf - a1["juros"]
    imc = div(mc, rec)
    return {"receita": rec, "custos_variaveis": cdir + var, "mc": mc, "custos_fixos": cf,
            "resultado": resultado, "lucratividade": div(resultado, rec),
            "pe": div(cf, imc) if imc == imc and imc > 0 else float("nan")}


def sensibilidade(res: dict, delta: float = 0.10) -> pd.DataFrame:
    base = simular(res)["resultado"]
    linhas = []
    for nome, chave_, sinal in [("Preço de venda", "preco", 1), ("Volume de vendas", "volume", 1),
                                ("Custo de materiais e mercadorias", "materiais", -1),
                                ("Salários e encargos", "salarios", -1), ("Demais custos fixos", "fixos", -1)]:
        up = simular(res, **{chave_: delta})["resultado"] - base
        dn = simular(res, **{chave_: -delta})["resultado"] - base
        linhas.append({"Variável": nome, f"+{int(delta*100)}%": up, f"-{int(delta*100)}%": dn,
                       "Impacto absoluto": max(abs(up), abs(dn))})
    return pd.DataFrame(linhas).sort_values("Impacto absoluto", ascending=False).reset_index(drop=True)


# ------------------------------------------------------------------------------
# 12. MERCADO (TAM / SAM / SOM) E CAPACIDADE x DEMANDA
# ------------------------------------------------------------------------------
def mercado(proj: dict, res: dict) -> dict:
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    tam_c, sam_c = float(proj.get("tam_clientes") or 0), float(proj.get("sam_clientes") or 0)
    som_c = sam_c * float(proj.get("som_pct") or 0) / 100
    tam, sam, som = tam_c * gasto, sam_c * gasto, som_c * gasto
    rec1 = res["anos"][1]["rec_bruta"]
    return {"tam": tam, "sam": sam, "som": som, "tam_c": tam_c, "sam_c": sam_c, "som_c": som_c,
            "receita_ano1": rec1, "pct_som": div(rec1, som), "pct_sam": div(rec1, sam),
            "preenchido": tam > 0 and sam > 0 and som > 0}


def capacidade_demanda(d: dict, res: dict) -> pd.DataFrame:
    cat = res["cat"]
    fator12 = (1 + res["g_m"]) ** 11
    linhas = []
    for _, r in d["capacidade"].iterrows():
        alvo = str(r["produto_ou_grupo"])
        if alvo in ("Fabricação Própria", "Serviço", "Revenda"):
            q = cat.loc[cat["categoria"] == alvo, "qtd_mensal"].sum()
        elif alvo == "Todos os produtos":
            q = cat["qtd_mensal"].sum()
        else:
            q = cat.loc[cat["descricao"] == alvo, "qtd_mensal"].sum()
        dem1, dem12 = float(q), float(q) * fator12
        cap = float(r["capacidade_mensal"])
        linhas.append({"Recurso": r["recurso"], "Atende": alvo, "Capacidade/mês": cap,
                       "Demanda Mês 1": dem1, "Demanda Mês 12": dem12,
                       "Uso Mês 1": div(dem1, cap), "Uso Mês 12": div(dem12, cap)})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 13. SEMÁFORO DOS INDICADORES COM "EXPLIQUE ESTE NÚMERO"
# ------------------------------------------------------------------------------
def semaforo(res: dict) -> list:
    a1 = res["anos"][1]
    it, tma = res["investimento_total"], res["tma"]
    out = []

    def add(nome, valor, txt, status, leitura, conta):
        out.append({"Indicador": nome, "Valor": txt, "Status": status, "Leitura": leitura,
                    "Como foi calculado": conta, "valor_num": valor})

    luc = div(a1["resultado"], a1["rec_bruta"])
    add("Lucratividade líquida (Ano 1)", luc, pct(luc),
        "🟢" if luc >= 0.10 else ("🟡" if luc > 0 else "🔴"),
        "Quanto sobra de lucro para cada R$ 1,00 vendido. Acima de 10% é bom para pequenos negócios de alimentação e varejo; "
        "negativo significa prejuízo.",
        f"{brl(a1['resultado'])} ÷ {brl(a1['rec_bruta'])} = {pct(luc)}")
    imc = a1["imc"]
    add("Margem de contribuição", imc, pct(imc),
        "🟢" if imc >= 0.40 else ("🟡" if imc >= 0.25 else "🔴"),
        "Parte de cada venda que sobra, depois dos custos variáveis, para pagar a estrutura fixa. Margens baixas "
        "exigem volume muito alto.",
        f"{brl(a1['mc'])} ÷ {brl(a1['rec_bruta'])} = {pct(imc)}")
    ms = div(a1["rec_bruta"] - a1["pe_contabil"], a1["rec_bruta"])
    add("Margem de segurança", ms, pct(ms),
        "🟢" if ms >= 0.20 else ("🟡" if ms > 0 else "🔴"),
        "Quanto as vendas podem cair antes de o negócio ter prejuízo. Abaixo de 0% significa que a receita prevista "
        "não cobre o ponto de equilíbrio.",
        f"({brl(a1['rec_bruta'])} − PE {brl(a1['pe_contabil'])}) ÷ {brl(a1['rec_bruta'])} = {pct(ms)}")
    add(f"VPL ({res['N']} anos)", res["vpl"], brl(res["vpl"]),
        "🟢" if res["vpl"] > 0 else "🔴",
        "Valor criado pelo projeto acima do retorno mínimo exigido (TMA). Positivo: o investimento compensa.",
        "Soma dos fluxos descontados: " + " + ".join(f"{brl(f)}/(1+{num_br(tma*100,1)}%)^{t}" for t, f in enumerate(res["fluxos"])))
    tir = res["tir"]
    add("TIR", tir if tir is not None else float("nan"), pct(tir) if tir is not None else "n.a.",
        "🟢" if tir is not None and tir > tma else "🔴",
        f"Rentabilidade anual do projeto. Deve superar a TMA de {pct(tma)}.",
        "Taxa que faz o VPL ser zero nos fluxos: " + "; ".join(brl(f) for f in res["fluxos"]))
    pb = res["payback_desc"]
    add("Payback descontado", pb, f"{num_br(pb,1)} meses" if pb == pb else f"> {res['MESES']} meses",
        "🟢" if pb == pb and pb <= 24 else ("🟡" if pb == pb else "🔴"),
        "Tempo para recuperar o investimento considerando o valor do dinheiro no tempo.",
        f"Investimento de {brl(it)} recuperado pelos fluxos descontados à TMA de {pct(tma)}")
    lc = div(a1["ac"], a1["pc"])
    add("Liquidez corrente (Ano 1)", lc, f"{num_br(lc)}x",
        "🟢" if lc >= 1.5 else ("🟡" if lc >= 1 else "🔴"),
        "Quantos reais de bens e direitos de curto prazo existem para cada R$ 1,00 de dívida de curto prazo.",
        f"{brl(a1['ac'])} ÷ {brl(a1['pc'])} = {num_br(lc)}x")
    pes = div(a1["pessoal"], a1["rec_liquida"])
    add("Peso da folha de pagamento", pes, pct(pes),
        "🟢" if pes <= 0.25 else ("🟡" if pes <= 0.35 else "🔴"),
        "Quanto da receita líquida vai para salários, encargos e benefícios. Acima de 35% costuma apertar o resultado.",
        f"{brl(a1['pessoal'])} ÷ {brl(a1['rec_liquida'])} = {pct(pes)}")
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        lc = rc["ltv_cac"]
        add("LTV/CAC (assinaturas)", lc, f"{num_br(lc, 1)}x" if lc == lc else "n.a.",
            "🟢" if lc == lc and lc >= 3 else ("🟡" if lc == lc and lc >= 1 else "🔴"),
            "Quantas vezes o lucro bruto que um assinante deixa ao longo da vida cobre o custo de conquistá-lo. "
            "Abaixo de 1 cada cliente novo dá prejuízo; 3 ou mais é a referência usada por investidores.",
            f"LTV {brl(rc['ltv'])} ÷ CAC {brl(rc['cac'])}" if lc == lc else "Sem gasto comercial ou sem novos assinantes no Mês 1")
        ch_ = rc["churn"]
        add("Churn mensal efetivo", ch_, pct(ch_),
            "🟢" if ch_ <= 0.03 else ("🟡" if ch_ <= 0.07 else "🔴"),
            "Percentual de assinantes que cancelam por mês. Cada ponto de churn encurta a vida do cliente e o LTV.",
            f"Churn de cada plano ponderado pelos assinantes ativos no Mês 12 = {pct(ch_)}; vida média ≈ 1 ÷ churn = "
            f"{num_br(rc['vida_meses'], 0)} meses")
    fm = res["fluxo_mensal"]
    minimo = float(fm["Saldo final"].min()) if not fm.empty else 0.0
    mes_min = fm.loc[fm["Saldo final"].idxmin(), "Mês"] if not fm.empty else "-"
    add("Menor saldo de caixa no Ano 1", minimo, brl(minimo),
        "🟢" if minimo > res["desembolso_m1"] * 0.5 else ("🟡" if minimo >= 0 else "🔴"),
        f"O caixa fica mais apertado em {mes_min}. Negativo indica que faltará dinheiro nesse mês.",
        f"Saldo inicial + recebimentos − pagamentos, mês a mês (menor valor em {mes_min})")
    return out


# ------------------------------------------------------------------------------
# 14. CHECAGENS COM INDICAÇÃO DE ONDE CORRIGIR
# ------------------------------------------------------------------------------
def tabela_checagens_v5(res: dict, d: dict, proj: dict) -> pd.DataFrame:
    base = tabela_checagens(res, d)
    destinos = {
        "Fontes de recursos x investimento total": ("investimentos", "Investimentos → Fontes de recursos",
            "Ajuste o capital dos sócios (aba Projeto) ou os empréstimos para que a soma seja igual ao investimento total."),
        "Margem unitária positiva": ("produtos", "Produtos e materiais → Produtos e serviços",
            "Revise o preço de venda ou o custo dos itens listados."),
        "Produtos com custo cadastrado": ("produtos", "Produtos e materiais → Ficha técnica",
            "Monte a ficha técnica ou informe o custo dos itens listados."),
        "Fichas técnicas com materiais válidos": ("produtos", "Produtos e materiais → Ficha técnica",
            "Cadastre os materiais que faltam ou troque o insumo na ficha técnica."),
        "Prazos de recebimento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de recebimento para somarem 100%."),
        "Prazos de pagamento somam 100%": ("investimentos", "Investimentos → Capital de giro",
            "Ajuste os percentuais de pagamento para somarem 100%."),
        "SWOT: ambiente coerente com a classificação": ("forcas", "Estratégia → Fatores internos e externos",
            "Fatores internos são Forças/Fraquezas; externos são Oportunidades/Ameaças."),
        "Caixa projetado não negativo": ("fluxo", "Finanças → Fluxo de caixa",
            "Aumente a reserva/capital de giro, renegocie prazos ou reduza custos."),
        "Liquidez corrente do Ano 1 >= 1,00": ("investimentos", "Investimentos → Capital de giro",
            "Aumente o capital de giro ou alongue o prazo das dívidas."),
        "VPL positivo": ("simulador", "Simulador e sensibilidade",
            "Veja quais variáveis mais afetam o resultado e teste ajustes de preço, volume e custos."),
    }
    base["destino"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[0])
    base["Onde corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[1])
    base["Como corrigir"] = base["Teste"].map(lambda t: destinos.get(t, ("", "", ""))[2])
    extras = []
    mk = mercado(proj, res)
    if mk["preenchido"]:
        ok = mk["pct_som"] <= 1
        extras.append({"Teste": "Meta de vendas cabe no mercado (SOM)", "Valor apurado": f"Receita = {pct(mk['pct_som'])} do SOM",
                       "Situação": "OK" if ok else "VERIFICAR",
                       "Observação": "A receita do Ano 1 não deve superar o mercado que a empresa consegue conquistar.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Revise as quantidades vendidas ou a estimativa de mercado."})
    else:
        extras.append({"Teste": "Tamanho do mercado estimado", "Valor apurado": "não preenchido", "Situação": "VERIFICAR",
                       "Observação": "Estime TAM, SAM e SOM para verificar se a meta de vendas é realista.",
                       "destino": "mercado", "Onde corrigir": "Cliente e mercado → Tamanho do mercado",
                       "Como corrigir": "Preencha clientes, percentual de participação e gasto anual por cliente."})
    cd = capacidade_demanda(d, res)
    if not cd.empty:
        exced = cd[cd["Uso Mês 12"] > 1]
        extras.append({"Teste": "Capacidade produtiva suficiente", "Valor apurado": f"{len(exced)} recurso(s) acima de 100%",
                       "Situação": "OK" if exced.empty else "VERIFICAR",
                       "Observação": ", ".join(exced["Recurso"]) or "A capacidade atende a demanda prevista até o Mês 12.",
                       "destino": "operacional", "Onde corrigir": "Marketing e operação → Plano operacional",
                       "Como corrigir": "Aumente a capacidade (equipamentos, pessoas, turnos) ou reveja a meta de vendas."})
    # --- Serviços: horas demandadas x horas disponíveis
    ch = custo_hora(res, d, proj)
    if ch["horas_demandadas_m1"] > 0:
        ok = ch["ocupacao_m12"] == ch["ocupacao_m12"] and ch["ocupacao_m12"] <= 1
        extras.append({"Teste": "Horas da equipe suficientes para os serviços",
                       "Valor apurado": f"Ocupação no Mês 12: {pct(ch['ocupacao_m12'])}" if ch["horas_cobraveis"] > 0
                       else "nenhuma pessoa na área de produção/operação",
                       "Situação": "OK" if ok else "VERIFICAR",
                       "Observação": "As horas vendidas não podem passar das horas cobráveis da equipe.",
                       "destino": "servicos", "Onde corrigir": "Finanças → Serviços e custo da hora",
                       "Como corrigir": "Contrate mais pessoas, aumente a utilização ou reduza as horas vendidas."})
    # --- Assinaturas
    ass = d.get("assinaturas")
    rc = res.get("recorrencia", {"tem": False})
    if ass is not None and not ass.empty:
        sem_churn = ass[(ass["churn_mensal_pct"] <= 0) & (ass["preco_ciclo"] > 0)]
        extras.append({"Teste": "Assinaturas com taxa de cancelamento (churn)",
                       "Valor apurado": f"{len(sem_churn)} plano(s) com churn zero",
                       "Situação": "OK" if sem_churn.empty else "VERIFICAR",
                       "Observação": ", ".join(sem_churn["plano"].astype(str)) or
                       "Todos os planos preveem cancelamentos, como acontece na prática.",
                       "destino": "assinaturas", "Onde corrigir": "Finanças → Assinaturas e receita recorrente",
                       "Como corrigir": "Informe um churn mensal realista (pesquisas citam algo entre 3% e 8% ao mês para pequenas "
                                        "empresas de software e aplicativos)."})
        if rc.get("tem"):
            lc = rc["ltv_cac"]
            ok = lc == lc and lc >= 3
            extras.append({"Teste": "LTV/CAC de pelo menos 3", "Valor apurado": f"{num_br(lc, 1)}x" if lc == lc else "sem CAC",
                           "Situação": "OK" if ok else "VERIFICAR",
                           "Observação": "Cada cliente deve gerar, ao longo da vida, ao menos 3 vezes o que custou para conquistá-lo.",
                           "destino": "assinaturas", "Onde corrigir": "Finanças → Assinaturas → Indicadores",
                           "Como corrigir": "Reduza o churn, aumente o preço/ARPU ou diminua o gasto comercial por novo cliente."})
    if str(proj.get("tipo_negocio", "")).startswith(("Serviço", "Assinatura")) and res["inv_estoque"] > 0:
        extras.append({"Teste": "Estoque inicial em negócio de serviço", "Valor apurado": brl(res["inv_estoque"]),
                       "Situação": "VERIFICAR",
                       "Observação": "Serviços e assinaturas costumam não ter estoque; confira se os materiais são mesmo necessários.",
                       "destino": "produtos", "Onde corrigir": "Produtos e materiais → Materiais",
                       "Como corrigir": "Zere o estoque inicial dos itens que não serão comprados antes da abertura."})
    return pd.concat([base, pd.DataFrame(extras)], ignore_index=True)


# ------------------------------------------------------------------------------
# 15. RAIO-X DO PLANO (nota de 0 a 100 e principais correções)
# ------------------------------------------------------------------------------
def raio_x(res: dict, chk: pd.DataFrame, completude: float, sem: list) -> dict:
    pontos_chk = 30 * (chk["Situação"] == "OK").mean() if len(chk) else 0
    pontos_comp = 30 * completude
    verdes = sum(1 for s in sem if s["Status"] == "🟢")
    amarelos = sum(1 for s in sem if s["Status"] == "🟡")
    pontos_fin = 40 * (verdes + 0.5 * amarelos) / max(len(sem), 1)
    nota = round(pontos_chk + pontos_comp + pontos_fin)
    acoes = [f"{r['Teste']}: {r['Como corrigir'] or r['Observação']}" for _, r in chk[chk["Situação"] != "OK"].iterrows()]
    acoes += [f"{s['Indicador']} em {s['Valor']}: {s['Leitura']}" for s in sem if s["Status"] == "🔴"]
    return {"nota": nota, "acoes": acoes[:5], "pontos": (round(pontos_comp), round(pontos_chk), round(pontos_fin))}


# ------------------------------------------------------------------------------
# 15. SERVIÇOS: CUSTO DA HORA PRODUTIVA, CAPACIDADE EM HORAS E PREÇO DO SERVIÇO
# ------------------------------------------------------------------------------
def custo_hora(res: dict, d: dict, proj: dict) -> dict:
    """
    Custo da hora produtiva = (mão de obra da operação + custos indiretos + depreciação) ÷ horas cobráveis.
    Horas cobráveis = pessoas da operação x horas contratadas por mês x taxa de utilização.
    Custo pleno do serviço = insumos + horas x custo da hora.
    Preço sugerido = custo pleno ÷ [1 − (impostos + despesas variáveis + despesas de estrutura % + margem)].
    """
    pes = d["pessoal"]
    area = pes["area"].fillna("").astype(str)
    pessoas = float(pes.loc[area.str.startswith("Produ"), "num_empregados"].sum())
    h_mes = float(proj.get("horas_mes_pessoa") or 168)
    util = float(proj.get("utilizacao_pct") or 75) / 100
    horas_contr = pessoas * h_mes
    horas_fat = horas_contr * util
    custo_base = res["mod_m"] + res["cif_m"] + res["dep_m"]
    ch = div(custo_base, horas_fat) if horas_fat > 0 else float("nan")
    rec_a1 = res["anos"][1]["rec_bruta"]
    estrutura = res["pessoal_adm_m"] + res["pessoal_com_m"] + res["desp_adm_m"] + res["desp_com_m"] + res["desp_fin_m"]
    desp_pct = div(estrutura * 12, rec_a1) if rec_a1 > 0 else 0.0       # média do Ano 1
    margem = res["margem_desejada"]
    divisor = 1 - (res["aliq"] + res["var_pct"] + desp_pct + margem)
    cat = res["cat"]
    sv = cat[cat["horas_unit"] > 0].copy() if "horas_unit" in cat else cat.iloc[0:0].copy()
    linhas = []
    for _, r in sv.iterrows():
        h = float(r["horas_unit"])
        mao = h * ch if ch == ch else float("nan")
        pleno = float(r["custo_unit"]) + (mao if mao == mao else 0.0)
        preco = float(r["preco_venda"])
        sug = div(pleno, divisor) if divisor > 0.05 else float("nan")
        margem_real = div(preco - pleno - preco * (res["aliq"] + res["var_pct"]), preco) if preco > 0 else float("nan")
        linhas.append({"Serviço": r["descricao"], "Horas por serviço": h, "Qtd/mês": float(r["qtd_mensal"]),
                       "Horas/mês": h * float(r["qtd_mensal"]), "Insumos": float(r["custo_unit"]),
                       "Mão de obra (horas x custo hora)": mao, "Custo pleno": pleno, "Preço atual": preco,
                       "Preço sugerido": sug, "Margem após custos diretos": margem_real,
                       "Valor da hora cobrada": div(preco, h)})
    tab = pd.DataFrame(linhas)
    horas_dem = float(tab["Horas/mês"].sum()) if not tab.empty else 0.0
    return {"pessoas": pessoas, "horas_mes_pessoa": h_mes, "utilizacao": util, "horas_contratadas": horas_contr,
            "horas_cobraveis": horas_fat, "custo_base": custo_base, "custo_hora": ch, "desp_pct": desp_pct,
            "divisor": divisor, "tabela": tab, "horas_demandadas_m1": horas_dem,
            "horas_demandadas_m12": horas_dem * (1 + res["g_m"]) ** 11,
            "ocupacao_m1": div(horas_dem, horas_fat) if horas_fat > 0 else float("nan"),
            "ocupacao_m12": div(horas_dem * (1 + res["g_m"]) ** 11, horas_fat) if horas_fat > 0 else float("nan")}

# ##############################################################################
# PARTE: pnbox_estrategia.py
# ##############################################################################





# ------------------------------------------------------------------------------
# 1. OBJETIVOS SMART
# ------------------------------------------------------------------------------
PERSPECTIVAS = ["Financeira", "Clientes", "Processos internos", "Aprendizado e crescimento"]
VERBOS_VAGOS = {"melhorar", "otimizar", "fortalecer", "buscar", "tentar", "procurar", "desenvolver", "trabalhar"}

SMART = {
    "S": ("Específico", "Diz exatamente o que será alcançado, com um verbo de ação e um foco claro."),
    "M": ("Mensurável", "Tem um indicador e uma meta numérica."),
    "A": ("Atingível", "Explica por que é possível com os recursos do plano."),
    "R": ("Relevante", "Está ligado à estratégia (SWOT, proposta de valor ou viabilidade)."),
    "T": ("Temporal", "Tem prazo definido (mês do plano)."),
}


def checar_smart(o: dict) -> dict:
    """Verifica cada critério SMART de um objetivo e devolve {letra: (ok, dica)}."""
    txt = str(o.get("objetivo") or "").strip()
    palavras = re.findall(r"\w+", txt.lower())
    vago = bool(palavras) and palavras[0] in VERBOS_VAGOS and len(palavras) < 8
    s_ok = len(palavras) >= 5 and not vago
    m_ok = bool(str(o.get("indicador") or "").strip()) and float(o.get("meta") or 0) != 0
    a_ok = len(str(o.get("por_que_atingivel") or "").strip()) >= 15
    r_ok = len(str(o.get("relevancia") or "").strip()) >= 10 and bool(str(o.get("perspectiva") or "").strip())
    t_ok = float(o.get("prazo_meses") or 0) > 0
    return {
        "S": (s_ok, "Escreva com verbo de ação e o que muda (ex.: \"Conquistar 300 assinantes pagantes\")." if not s_ok else ""),
        "M": (m_ok, "Informe o indicador e a meta numérica." if not m_ok else ""),
        "A": (a_ok, "Explique em uma frase por que a meta é possível (equipe, verba, capacidade, histórico)." if not a_ok else ""),
        "R": (r_ok, "Escolha a perspectiva e diga a que estratégia ou problema ele responde." if not r_ok else ""),
        "T": (t_ok, "Defina o prazo em meses." if not t_ok else ""),
    }


def frase_smart(o: dict) -> str:
    meta = num_br(o.get("meta") or 0, 2 if abs(float(o.get("meta") or 0)) < 100 else 0)
    un = str(o.get("unidade") or "").strip()
    prazo = int(float(o.get("prazo_meses") or 0))
    resp = str(o.get("responsavel") or "").strip()
    return (f"{str(o.get('objetivo') or '').strip()}: atingir {meta} {un} em {o.get('indicador') or 'indicador'} "
            f"até o mês {prazo}" + (f", sob responsabilidade de {resp}" if resp else "") + ".")


def objetivos_sugeridos(proj: dict, d: dict, res: dict) -> list:
    """Rascunhos de objetivos SMART a partir dos números do plano."""
    a1 = res["anos"][1]
    sug = [
        {"objetivo": "Alcançar o faturamento previsto no primeiro ano", "perspectiva": "Financeira",
         "indicador": "Receita bruta acumulada", "meta": round(a1["rec_bruta"], 0), "unidade": "R$", "prazo_meses": 12,
         "por_que_atingivel": "Meta calculada no plano a partir do volume e dos preços previstos.",
         "relevancia": "Garante a viabilidade financeira projetada.", "iniciativa": "Acompanhar vendas semanalmente."},
        {"objetivo": "Operar acima do ponto de equilíbrio contábil todos os meses", "perspectiva": "Financeira",
         "indicador": "Receita mensal acima do PE", "meta": round(a1["pe_contabil"] / 12, 0), "unidade": "R$/mês", "prazo_meses": 12,
         "por_que_atingivel": "O ponto de equilíbrio foi calculado com os custos do plano.",
         "relevancia": "Evita prejuízo e consumo do capital de giro.", "iniciativa": "Controle mensal de custos fixos."},
    ]
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        sug.append({"objetivo": "Manter o cancelamento de assinantes sob controle", "perspectiva": "Clientes",
                    "indicador": "Churn mensal", "meta": round(rc["churn"] * 100, 1), "unidade": "% ao mês (máximo)",
                    "prazo_meses": 12, "por_que_atingivel": "Churn usado na projeção dos planos.",
                    "relevancia": "Cada ponto de churn reduz o LTV e o MRR.", "iniciativa": "Programa de sucesso do cliente."})
        sug.append({"objetivo": "Chegar ao número de assinantes previsto no fim do primeiro ano", "perspectiva": "Clientes",
                    "indicador": "Assinantes ativos", "meta": round(rc["ativos12"], 0), "unidade": "assinantes", "prazo_meses": 12,
                    "por_que_atingivel": "Resultado da captação e do churn projetados.",
                    "relevancia": "Base de assinantes sustenta a receita recorrente.", "iniciativa": "Campanhas mensais de captação."})
    else:
        sug.append({"objetivo": "Conquistar e manter clientes satisfeitos", "perspectiva": "Clientes",
                    "indicador": "Nota de satisfação (NPS)", "meta": 70, "unidade": "pontos", "prazo_meses": 12,
                    "por_que_atingivel": "Pesquisa simples após a compra, com ajustes mensais.",
                    "relevancia": "Clientes satisfeitos voltam e indicam, reduzindo o custo de aquisição.",
                    "iniciativa": "Pesquisa de satisfação pelo WhatsApp."})
    ch = custo_hora(res, d, proj)
    if ch["horas_cobraveis"] > 0 and ch["horas_demandadas_m1"] > 0:
        sug.append({"objetivo": "Usar bem as horas da equipe de serviços", "perspectiva": "Processos internos",
                    "indicador": "Taxa de utilização", "meta": round(ch["utilizacao"] * 100, 0), "unidade": "%", "prazo_meses": 6,
                    "por_que_atingivel": "Taxa prevista na ferramenta de serviços.",
                    "relevancia": "Hora ociosa é custo sem receita.", "iniciativa": "Agenda semanal e controle de horas."})
    else:
        sug.append({"objetivo": "Padronizar os processos de produção e atendimento", "perspectiva": "Processos internos",
                    "indicador": "Processos documentados", "meta": 5, "unidade": "processos", "prazo_meses": 6,
                    "por_que_atingivel": "Equipe pequena e processos já descritos no plano operacional.",
                    "relevancia": "Reduz erros e desperdício, protegendo a margem.", "iniciativa": "Fichas e checklists."})
    sug.append({"objetivo": "Capacitar a equipe nas competências críticas do negócio", "perspectiva": "Aprendizado e crescimento",
                "indicador": "Horas de capacitação por pessoa", "meta": 20, "unidade": "horas/ano", "prazo_meses": 12,
                "por_que_atingivel": "Cursos gratuitos do Sebrae e treinamentos internos.",
                "relevancia": "Equipe preparada sustenta a qualidade prometida na proposta de valor.",
                "iniciativa": "Plano anual de treinamento."})
    return sug


# ------------------------------------------------------------------------------
# 2. BUSINESS MODEL CANVAS (montado a partir das outras ferramentas)
# ------------------------------------------------------------------------------
def _lista(x, n=6):
    return [str(v).strip() for v in x if str(v).strip()][:n]


def canvas_blocos(proj: dict, d: dict, res: dict) -> dict:
    seg = d["segmentos"]
    seg_l = []
    if not seg.empty:
        s_ = seg.assign(m=seg[["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]].mean(axis=1))
        seg_l = _lista(s_.sort_values("m", ascending=False)["nome"], 5)
    props = d.get("propostas")
    prop_l = _lista(props["oferecem"], 4) if props is not None and not props.empty else []
    if props is not None and not props.empty:
        vals = [v.strip() for x in props["valores"] for v in str(x).split(";") if v.strip()]
        if vals:
            prop_l.append("Valores: " + ", ".join(dict.fromkeys(vals)))
    canais = _lista(d["canais"]["canal"], 6) if "canal" in d["canais"] else []
    rel = [x.strip() for x in str(proj.get("canvas_relacionamento") or "").split("\n") if x.strip()]
    if not rel:
        etapas = d["jornada_cliente"]
        rel = _lista(etapas.loc[etapas["macro_fase"] == "Depois", "oportunidades"], 4)
    rec = []
    cat = res["cat"]
    if not cat.empty:
        for c_, v_ in cat.groupby("categoria")["receita_mensal"].sum().sort_values(ascending=False).items():
            rec.append(f"{c_}: {pct(div(v_ * 12, res['anos'][1]['rec_bruta']), 0)} da receita")
    if res.get("recorrencia", {}).get("tem"):
        a_ = d["assinaturas"]
        rec.append("Assinaturas: " + ", ".join(f"{p} ({brl(v)})" for p, v in zip(a_["plano"], a_["preco_ciclo"])))
    recursos = _lista(d["inv_fixos"].sort_values("valor_unitario", ascending=False)["descricao"], 3) + \
        [f"Equipe: {num_br(d['pessoal']['num_empregados'].sum(), 0)} pessoa(s)"] if not d["pessoal"].empty else \
        _lista(d["inv_fixos"]["descricao"], 3)
    atividades = [x.strip() for x in re.split(r"->|→|;|\n", str(proj.get("op_processos") or "")) if x.strip()][:6]
    parc = [x.strip() for x in str(proj.get("canvas_parcerias") or "").split("\n") if x.strip()]
    if not parc and not d["fornecedores"].empty:
        fo = d["fornecedores"]
        parc = _lista(fo.sort_values("critico", ascending=False)["nome"], 5)
    a1 = res["anos"][1]
    custos = [f"Pessoal: {pct(div(a1['pessoal'], a1['rec_bruta']), 0)} da receita",
              f"Custos diretos: {pct(div(a1['custo_dir'], a1['rec_bruta']), 0)}",
              f"Demais custos fixos: {pct(div(a1['fixos'], a1['rec_bruta']), 0)}",
              f"Investimento inicial: {brl(res['investimento_total'])}"]
    return {"Parcerias-chave": parc, "Atividades-chave": atividades, "Recursos-chave": recursos,
            "Proposta de valor": prop_l, "Relacionamento com clientes": rel, "Canais": canais,
            "Segmentos de clientes": seg_l, "Estrutura de custos": custos, "Fontes de receita": rec}


# ------------------------------------------------------------------------------
# 3. MATRIZ BCG (crescimento do mercado x participação relativa)
# ------------------------------------------------------------------------------
QUADRANTES_BCG = {
    "Estrela": ("⭐", "Mercado em alta e posição forte. Investir para manter a liderança; consome caixa, mas tende a virar vaca leiteira."),
    "Vaca leiteira": ("🐄", "Mercado maduro e posição forte. Gera caixa: manter eficiência e usar o dinheiro para financiar estrelas e interrogações."),
    "Interrogação": ("❓", "Mercado em alta, posição fraca. Decidir: investir para ganhar participação ou abandonar."),
    "Abacaxi": ("🍍", "Mercado parado e posição fraca. Rever preço e custo, reposicionar ou descontinuar."),
}


def bcg_tabela(d: dict, res: dict, limite_cresc: float = 10.0) -> pd.DataFrame:
    cat = res["cat"]
    b = d["bcg"]
    linhas = []
    for _, r in b.iterrows():
        prod = str(r["produto"]).strip()
        if not prod:
            continue
        linha = cat[cat["descricao"] == prod]
        receita = float(linha["receita_mensal"].sum()) * 12 if not linha.empty else 0.0
        margem = float(linha["margem_pct"].iloc[0]) if not linha.empty else float("nan")
        if prod.startswith("Assinatura") and res.get("recorrencia", {}).get("tem"):
            receita = res["anos"][1]["rec_cat"].get("Assinatura", 0.0)
        cres, part = float(r["crescimento_mercado_pct"]), float(r["participacao_relativa"])
        alto, forte = cres >= limite_cresc, part >= 1.0
        quad = "Estrela" if alto and forte else "Vaca leiteira" if forte else "Interrogação" if alto else "Abacaxi"
        linhas.append({"Produto": prod, "Crescimento do mercado (%)": cres, "Participação relativa": part,
                       "Receita anual": receita, "Margem": margem, "Quadrante": quad})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 4. MATRIZ GE / McKINSEY (atratividade do mercado x força competitiva)
# ------------------------------------------------------------------------------
def forca_competitiva_padrao(d: dict) -> float:
    """Nota da 'Nossa Empresa' na matriz de concorrência comparada à média dos concorrentes, em escala 0-10."""
    n = d["concorrencia_notas"]
    if n.empty:
        return 5.0
    nossa = n.loc[n["empresa"] == "Nossa Empresa", "nota"].mean()
    outras = n.loc[n["empresa"] != "Nossa Empresa", "nota"].mean()
    if nossa != nossa:
        return 5.0
    if outras != outras or outras == 0:
        return float(nossa)
    return float(min(10, max(0, 5 * nossa / outras)))


CELULAS_GE = {3: "Seguro e crescimento", 2: "Seletivo ou cauteloso", 1: "Expansão limitada ou colheita", 0: "Perigo ou colheita"}
ACOES_GE = {"Seguro e crescimento": "Investir para crescer e defender a posição.",
            "Seletivo ou cauteloso": "Investir com cautela, só onde houver vantagem clara.",
            "Expansão limitada ou colheita": "Manter com pouco investimento e aproveitar o caixa.",
            "Perigo ou colheita": "Colher o que for possível ou sair do segmento."}


def ge_tabela(d: dict) -> pd.DataFrame:
    seg = d["segmentos"]
    ge = d["ge"]
    padrao = forca_competitiva_padrao(d)
    forcas = dict(zip(ge["unidade"], ge["forca"]))
    linhas = []
    for _, s_ in seg.iterrows():
        atr = float(pd.Series([s_["nota_tamanho"], s_["nota_lucro"], s_["nota_longevidade"], s_["nota_concorrencia"]]).mean())
        f_inf = forcas.get(s_["nome"], 0) or 0
        forca = float(f_inf) if f_inf > 0 else round((padrao + float(s_["nota_aderencia"])) / 2, 1)
        def nivel(v):
            return "Alta" if v >= 6.67 else ("Média" if v >= 3.33 else "Baixa")
        na, nf = nivel(atr), nivel(forca)
        pontos = {"Alta": 2, "Média": 1, "Baixa": 0}
        soma = pontos[na] + pontos[nf]
        rec = CELULAS_GE[min(soma, 3) if soma < 3 else 3]
        linhas.append({"Segmento": s_["nome"], "Atratividade": round(atr, 1), "Força competitiva": round(forca, 1),
                       "Nível atratividade": na, "Nível força": nf, "Recomendação": rec,
                       "Origem da força": "informada" if f_inf > 0 else "automática (aderência + concorrência)"})
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 5. CINCO FORÇAS DE PORTER
# ------------------------------------------------------------------------------
FORCAS_PORTER = [
    ("Rivalidade entre concorrentes", "Quantos concorrentes, quão parecidos e quão agressivos em preço?"),
    ("Ameaça de novos entrantes", "É fácil abrir um negócio igual? Exige muito capital, licença, marca ou tecnologia?"),
    ("Ameaça de produtos substitutos", "O cliente resolve o mesmo problema de outro jeito?"),
    ("Poder de barganha dos clientes", "O cliente troca de fornecedor com facilidade? Compra em grande volume? Compara preço?"),
    ("Poder de barganha dos fornecedores", "Poucos fornecedores? Insumo crítico ou difícil de substituir?"),
]


def porter_sugestoes(proj: dict, d: dict, res: dict) -> dict:
    """Nota sugerida (1 a 5) e motivo para cada força, a partir dos dados do plano."""
    n_conc = int(d["concorrentes"]["nome"].astype(str).str.strip().ne("").sum())
    nota_conc = forca_competitiva_padrao(d)
    riv = 2 if n_conc <= 1 else (3 if n_conc <= 3 else 4)
    if nota_conc < 5:
        riv = min(5, riv + 1)
    inv = res["investimento_total"]
    ent = 4 if inv < 50000 else (3 if inv < 200000 else 2)
    imc = res["anos"][1]["imc"]
    subs = 3
    cli = 4 if str(proj.get("tipo_negocio", "")).startswith("Produtos") and imc < 0.4 else 3
    fo = d["fornecedores"]
    criticos = int(fo["critico"].sum()) if not fo.empty else 0
    forn = 2 if fo.empty or criticos == 0 else (3 if criticos <= 2 else 4)
    return {
        FORCAS_PORTER[0][0]: (riv, f"{n_conc} concorrente(s) cadastrado(s); nota da empresa frente à média deles: {num_br(nota_conc, 1)}/10."),
        FORCAS_PORTER[1][0]: (ent, f"Investimento inicial de {brl(inv)}: quanto menor, mais fácil alguém copiar o negócio."),
        FORCAS_PORTER[2][0]: (subs, "Avalie os substitutos citados na proposta de valor (o que o cliente usa hoje)."),
        FORCAS_PORTER[3][0]: (cli, f"Margem de contribuição de {pct(imc, 0)}: margens baixas indicam clientes sensíveis a preço."),
        FORCAS_PORTER[4][0]: (forn, f"{criticos} fornecedor(es) marcado(s) como crítico(s)."),
    }


def porter_tabela(d: dict) -> pd.DataFrame:
    p = d["porter"]
    base = pd.DataFrame({"Força": [f for f, _ in FORCAS_PORTER]})
    if p.empty:
        base["Intensidade"] = 0.0
        base["Justificativa"] = ""
        base["Estratégia"] = ""
        return base
    m = p.set_index("forca")
    base["Intensidade"] = base["Força"].map(lambda f: float(m["intensidade"].get(f, 0)) if f in m.index else 0.0)
    base["Justificativa"] = base["Força"].map(lambda f: str(m["justificativa"].get(f, "")) if f in m.index else "")
    base["Estratégia"] = base["Força"].map(lambda f: str(m["estrategia"].get(f, "")) if f in m.index else "")
    return base


def porter_leitura(tab: pd.DataFrame) -> str:
    v = tab["Intensidade"]
    v = v[v > 0]
    if v.empty:
        return "Atribua uma intensidade de 1 a 5 a cada força."
    media = float(v.mean())
    forte = tab.loc[tab["Intensidade"].idxmax(), "Força"]
    nivel = "alta" if media >= 3.5 else ("moderada" if media >= 2.5 else "baixa")
    return (f"Intensidade média de {num_br(media, 1)} em 5: a pressão competitiva do setor é {nivel}. "
            f"A força mais intensa é \"{forte}\"; ela deve orientar as primeiras ações estratégicas. "
            + ("Setores de pressão alta exigem diferenciação clara ou custos muito baixos para sustentar margens."
               if nivel == "alta" else "Há espaço para margens saudáveis se o negócio entregar a proposta de valor."))


# ------------------------------------------------------------------------------
# 6. BALANCED SCORECARD
# ------------------------------------------------------------------------------
def bsc_indicadores_plano(proj: dict, d: dict, res: dict) -> dict:
    """Indicadores que o próprio plano já calcula, organizados pelas quatro perspectivas."""
    a1 = res["anos"][1]
    fin = [("Lucratividade líquida (Ano 1)", pct(div(a1["resultado"], a1["rec_bruta"]))),
           (f"VPL ({res['N']} anos)", brl(res["vpl"])),
           ("Margem de contribuição", pct(a1["imc"]))]
    cli = []
    rc = res.get("recorrencia", {"tem": False})
    if rc.get("tem"):
        fin.append(("MRR no Mês 12", brl(rc["mrr12"])))
        cli += [("Churn mensal", pct(rc["churn"])), ("LTV/CAC", f"{num_br(rc['ltv_cac'], 1)}x" if rc["ltv_cac"] == rc["ltv_cac"] else "n.a.")]
    mk = mercado(proj, res)
    if mk["preenchido"]:
        cli.append(("Participação no SOM (Ano 1)", pct(mk["pct_som"])))
    proc = []
    ch = custo_hora(res, d, proj)
    if ch["horas_demandadas_m1"] > 0:
        proc.append(("Ocupação da equipe (Mês 12)", pct(ch["ocupacao_m12"])))
    proc.append(("Ciclo financeiro", f"{num_br(res['ciclo_fin'], 0)} dias"))
    apr = [("Peso da folha na receita líquida", pct(div(a1["pessoal"], a1["rec_liquida"])))]
    return {"Financeira": fin, "Clientes": cli, "Processos internos": proc, "Aprendizado e crescimento": apr}


# ------------------------------------------------------------------------------
# 7. REFORMA TRIBUTÁRIA: SIMPLES "PURO" x IBS/CBS POR FORA ("HÍBRIDO")
# ------------------------------------------------------------------------------
# Parcela do DAS (1ª faixa) que corresponde aos tributos substituídos pela reforma, pelas tabelas da LC 123/2006.
# PIS + Cofins são substituídos pela CBS (a partir de 2027); ICMS/ISS pelo IBS (2029 a 2033); IPI zerado.
PARTILHA_DAS = {
    "Anexo I - Comércio": {"federais_consumo": 0.1274 + 0.0276, "icms_iss": 0.34},
    "Anexo II - Indústria": {"federais_consumo": 0.1151 + 0.0249 + 0.075, "icms_iss": 0.32},
    "Anexo III - Serviços": {"federais_consumo": 0.1282 + 0.0278, "icms_iss": 0.335},
    "Anexo IV - Serviços": {"federais_consumo": 0.1767 + 0.0383, "icms_iss": 0.445},
    "Anexo V - Serviços": {"federais_consumo": 0.1410 + 0.0305, "icms_iss": 0.14},
}


def simular_reforma(receita: float, aliq_simples: float, anexo: str, ano: int, cbs: float, ibs: float,
                    b2b: float, compras_credito: float) -> dict:
    """
    Comparação didática e simplificada (alíquotas por dentro, sobre a receita):
    - Simples "puro": paga o DAS cheio; o cliente empresa só aproveita como crédito a parcela de IBS/CBS do DAS.
    - Simples "híbrido": tira do DAS a parcela dos tributos substituídos e recolhe IBS e CBS pelo regime regular,
      com crédito sobre as compras; o cliente empresa aproveita o crédito integral.
    """
    part = PARTILHA_DAS.get(anexo, PARTILHA_DAS["Anexo III - Serviços"])
    # Em 2027-2028 só PIS/Cofins já foram substituídos (CBS); ICMS/ISS seguem no DAS até a transição do IBS
    # ICMS e ISS são reduzidos em 10% ao ano de 2029 a 2032 e extintos em 2033 (EC 132/2023)
    frac_icms = 0.0 if ano <= 2028 else (1.0 if ano >= 2033 else 0.1 * (ano - 2028))
    substituida = part["federais_consumo"] + part["icms_iss"] * frac_icms
    das = receita * aliq_simples
    iva = cbs + ibs
    puro_total = das
    puro_credito_cliente = das * substituida * b2b
    das_hibrido = das * (1 - substituida)
    iva_debito = receita * iva
    iva_credito = compras_credito * iva
    hib_total = das_hibrido + max(0.0, iva_debito - iva_credito)
    hib_credito_cliente = iva_debito * b2b
    return {"das_puro": das, "puro_total": puro_total, "puro_credito_cliente": puro_credito_cliente,
            "das_hibrido": das_hibrido, "iva_debito": iva_debito, "iva_credito": iva_credito, "hib_total": hib_total,
            "hib_credito_cliente": hib_credito_cliente, "parcela_substituida": substituida, "iva": iva,
            "diferenca": hib_total - puro_total,
            "custo_liquido_cliente_puro": -puro_credito_cliente, "custo_liquido_cliente_hib": -hib_credito_cliente}

# ##############################################################################
# PARTE: pnbox_validacao.py
# ##############################################################################




# ------------------------------------------------------------------------------
# 1. PLACAR DE EVIDÊNCIAS
# ------------------------------------------------------------------------------
STATUS_EVIDENCIA = {
    "Estimada (sem fonte)": 0.0,
    "Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)": 0.5,
    "Testada com clientes (entrevista, pré-venda, piloto)": 1.0,
}
STATUS_CURTO = {"Estimada (sem fonte)": "estimada",
                "Pesquisa secundária (IBGE, Sebrae, artigos, concorrentes)": "pesquisada",
                "Testada com clientes (entrevista, pré-venda, piloto)": "testada"}


def premissas_plano(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    """Premissas que movem o resultado, com o valor usado no plano e o peso de cada uma (soma 100%)."""
    a1 = res["anos"][1]
    sens = sensibilidade(res) if a1["rec_bruta"] > 0 else pd.DataFrame(columns=["Variável", "Impacto absoluto"])
    imp = dict(zip(sens["Variável"], sens["Impacto absoluto"]))
    base = float(np.mean(list(imp.values()))) if imp else 1.0
    base = base if base > 0 else 1.0
    cat = res.get("cat", pd.DataFrame())
    qtd = float(cat["qtd_mensal"].sum()) if "qtd_mensal" in cat else 0.0
    ticket = div(float((cat["preco_venda"] * cat["qtd_mensal"]).sum()), qtd) if qtd else float("nan")
    linhas = [
        ("Preço de venda", f"preço médio {brl(ticket)}" if ticket == ticket else "assinaturas", imp.get("Preço de venda", base)),
        ("Volume de vendas do Ano 1", f"receita {brl(a1['rec_bruta'])}", imp.get("Volume de vendas", base)),
        ("Custo de materiais e mercadorias", brl(a1["custo_dir"]), imp.get("Custo de materiais e mercadorias", base)),
        ("Salários e encargos", brl(a1["pessoal"]), imp.get("Salários e encargos", base)),
        ("Demais custos fixos", brl(a1["fixos"]), imp.get("Demais custos fixos", base)),
        ("Crescimento das vendas", f"{num_br(proj.get('crescimento_mensal_ano1_pct') or 0, 1)}% a.m. no Ano 1; "
                                   f"{num_br(proj.get('crescimento_anual_pct') or 0, 1)}% a.a. depois",
         0.5 * imp.get("Volume de vendas", base)),
    ]
    m = mercado(proj, res)
    if m["preenchido"]:
        linhas.append(("Participação de mercado (SOM)", f"{num_br(proj.get('som_pct') or 0, 1)}% do SAM", 0.4 * base))
    if not d["funil_vendas"].empty:
        linhas.append(("Conversão do funil de vendas", "taxas do funil", 0.4 * base))
    rec = res.get("recorrencia", {})
    if rec.get("tem"):
        linhas.append(("Cancelamento mensal (churn)", pct(rec.get("churn", 0)), 0.8 * base))
        linhas.append(("Custo de aquisição (CAC)", brl(rec.get("cac", 0)), 0.4 * base))
    linhas.append(("Prazo de recebimento das vendas", "prazos do capital de giro", 0.2 * base))
    df = pd.DataFrame(linhas, columns=["premissa", "valor_plano", "peso"])
    df["peso"] = df["peso"].clip(lower=0) / max(df["peso"].clip(lower=0).sum(), 1e-9)
    return df


def placar_evidencias(proj: dict, d: dict, res: dict) -> dict:
    pr = premissas_plano(proj, d, res)
    ev = d.get("evidencias", pd.DataFrame(columns=["premissa", "status", "fonte", "observacao"]))
    mapa = {str(r["premissa"]): r for _, r in ev.iterrows()}
    st_, fo = [], []
    for p in pr["premissa"]:
        r = mapa.get(p)
        s = str(r["status"]) if r is not None and str(r["status"]) in STATUS_EVIDENCIA else "Estimada (sem fonte)"
        st_.append(s)
        fo.append(str(r["fonte"]) if r is not None else "")
    pr["status"], pr["fonte"] = st_, fo
    pr["nota"] = pr["status"].map(STATUS_EVIDENCIA)
    score = float((pr["peso"] * pr["nota"]).sum())
    pend = pr[pr["nota"] < 1].sort_values("peso", ascending=False)
    prox = pend.iloc[0]["premissa"] if not pend.empty else ""
    return {"tabela": pr, "score": score, "proxima": prox,
            "testadas": int((pr["nota"] == 1).sum()), "n": len(pr)}


# ------------------------------------------------------------------------------
# 2. ENTREVISTAS, VAN WESTENDORP E INTENÇÃO DE COMPRA
# ------------------------------------------------------------------------------
PAGARIA = ["Sim", "Talvez", "Não"]

ROTEIRO_ENTREVISTA = """**Antes:** converse com pessoas do segmento-alvo, uma de cada vez, sem apresentar a sua ideia no começo.
Pergunte sobre o **passado e o comportamento real**, não sobre opiniões e intenções futuras (Fitzpatrick, 2013).

1. Conte a última vez que você precisou resolver [o problema]. O que aconteceu?
2. O que foi mais difícil ou chato nisso? Por quê?
3. O que você já tentou para resolver? Quanto gastou com isso (dinheiro e tempo)?
4. O que você usa hoje? O que gosta e o que não gosta nessa solução?
5. Com que frequência isso acontece?
6. *(Só agora apresente a ideia em uma frase.)* O que achou? O que faltaria para você usar?
7. Quanto pagaria por isso? Pagaria hoje, com um pré-cadastro ou sinal?
8. Quem mais você conhece que tem esse problema? Pode me apresentar?

**Sinais fortes:** a pessoa já gasta dinheiro ou tempo com o problema, pede para ser avisada, aceita pagar um sinal ou indica
outras pessoas. **Sinais fracos:** elogios ("que ideia legal!") e promessas sobre o futuro ("eu compraria, com certeza")."""


def _itens(txt) -> list:
    return [re.sub(r"\s+", " ", x).strip().lower() for x in re.split(r"[;\n,]", str(txt or "")) if x.strip()]


def resumo_entrevistas(d: dict, preco_ref: float = float("nan")) -> dict:
    e = d.get("entrevistas", pd.DataFrame())
    e = e[e["entrevistado"].astype(str).str.strip() != ""] if not e.empty else e
    n = len(e)
    if n == 0:
        return {"n": 0}
    dores = Counter(x for t in e["dores"] for x in set(_itens(t)))
    tab_dores = pd.DataFrame([{"Dor citada": k, "Entrevistas": v, "% das entrevistas": v / n}
                              for k, v in dores.most_common(12)])
    pag = e["pagaria"].astype(str)
    vals = pd.to_numeric(e["valor_pagaria"], errors="coerce")
    vals = vals[vals > 0]
    med = float(vals.median()) if len(vals) else float("nan")
    return {"n": n, "dores": tab_dores, "dor_top": dores.most_common(1)[0] if dores else ("", 0),
            "pct_sim": float((pag == "Sim").mean()), "pct_talvez": float((pag == "Talvez").mean()),
            "pct_nao": float((pag == "Não").mean()), "mediana_valor": med,
            "valor_vs_preco": div(med, preco_ref) if preco_ref == preco_ref and preco_ref > 0 else float("nan")}


def van_westendorp(df: pd.DataFrame) -> dict | None:
    """Price Sensitivity Meter (Van Westendorp, 1976). Retorna as curvas e os quatro pontos de preço."""
    if df is None or df.empty:
        return None
    x = df[["muito_barato", "barato", "caro", "muito_caro"]].apply(pd.to_numeric, errors="coerce")
    x = x[(x > 0).all(axis=1)]
    # descarta respostas incoerentes (a ordem precisa ser muito barato <= barato <= caro <= muito caro)
    x = x[(x["muito_barato"] <= x["barato"]) & (x["barato"] <= x["caro"]) & (x["caro"] <= x["muito_caro"])]
    n = len(x)
    if n < 5:
        return {"n": n, "ok": False}
    lo, hi = float(x.values.min()) * 0.8, float(x.values.max()) * 1.1
    grade = np.linspace(lo, hi, 400)
    mb = np.array([(x["muito_barato"] >= p).mean() for p in grade])
    ba = np.array([(x["barato"] >= p).mean() for p in grade])
    ca = np.array([(x["caro"] <= p).mean() for p in grade])
    mc = np.array([(x["muito_caro"] <= p).mean() for p in grade])
    nao_barato, nao_caro = 1 - ba, 1 - ca

    def cruz(a, b):
        dif = a - b
        idx = np.where(np.diff(np.sign(dif)) != 0)[0]
        if len(idx) == 0:
            return float("nan")
        i = idx[0]
        t = dif[i] / (dif[i] - dif[i + 1]) if dif[i] != dif[i + 1] else 0.0
        return float(grade[i] + t * (grade[i + 1] - grade[i]))

    curvas = pd.DataFrame({"Preço": grade, "Muito barato": mb, "Barato": ba, "Caro": ca, "Muito caro": mc,
                           "Não é barato": nao_barato, "Não é caro": nao_caro})
    return {"n": n, "ok": True, "curvas": curvas,
            "pmc": cruz(mb, nao_barato),      # preço marginal barato (limite inferior)
            "pme": cruz(mc, nao_caro),        # preço marginal caro (limite superior)
            "opp": cruz(mb, mc),              # preço ótimo
            "ipp": cruz(nao_barato, nao_caro)}  # preço de indiferença


def intencao_ajustada(proj: dict, res: dict) -> dict:
    """Converte intenção de compra declarada em clientes prováveis (fatores editáveis)."""
    base = float(proj.get("sam_clientes") or 0)
    c, p = float(proj.get("intencao_certamente_pct") or 0) / 100, float(proj.get("intencao_provavel_pct") or 0) / 100
    fc, fp = float(proj.get("fator_certamente_pct") or 70) / 100, float(proj.get("fator_provavel_pct") or 30) / 100
    declarados = base * (c + p)
    ajustados = base * (c * fc + p * fp)
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    plano = div(res["anos"][1]["rec_bruta"], gasto) if gasto > 0 else float("nan")
    return {"base": base, "declarados": declarados, "ajustados": ajustados, "plano": plano,
            "cobertura": div(ajustados, plano) if plano == plano and plano > 0 else float("nan"),
            "ok": base > 0 and (c + p) > 0}


# ------------------------------------------------------------------------------
# 3. MARKETING: PROPOSTA DE VALOR, CURVA DE VALOR, POSICIONAMENTO, CANAIS
# ------------------------------------------------------------------------------
BLOCOS_VPC = [("vpc_tarefas", "Tarefas do cliente (jobs)", "O que o cliente tenta resolver ou realizar"),
              ("vpc_dores", "Dores", "O que incomoda, atrapalha ou gera risco antes, durante e depois da tarefa"),
              ("vpc_ganhos", "Ganhos", "Resultados e benefícios que o cliente quer ou que o surpreenderiam"),
              ("vpc_produtos", "Produtos e serviços", "O que a empresa oferece"),
              ("vpc_aliviadores", "Aliviadores de dores", "Como a oferta elimina ou reduz cada dor"),
              ("vpc_criadores", "Criadores de ganhos", "Como a oferta produz cada ganho")]


def linhas_txt(t) -> list:
    return [x.strip(" -•\t") for x in str(t or "").splitlines() if x.strip(" -•\t")]


def encaixe_vpc(proj: dict, d: dict) -> dict:
    dores, ganhos = linhas_txt(proj.get("vpc_dores")), linhas_txt(proj.get("vpc_ganhos"))
    enc = d.get("vpc_encaixe", pd.DataFrame(columns=["item", "tipo", "resposta"]))
    resp = {(str(r["tipo"]), str(r["item"])): str(r["resposta"]) for _, r in enc.iterrows()}
    linhas = [{"tipo": "Dor", "item": x, "resposta": resp.get(("Dor", x), "")} for x in dores] + \
             [{"tipo": "Ganho", "item": x, "resposta": resp.get(("Ganho", x), "")} for x in ganhos]
    tab = pd.DataFrame(linhas, columns=["tipo", "item", "resposta"])
    cobertos = int((tab["resposta"].astype(str).str.strip() != "").sum()) if not tab.empty else 0
    return {"tabela": tab, "cobertos": cobertos, "total": len(tab), "encaixe": div(cobertos, len(tab)) if len(tab) else float("nan")}


ACOES_ERRC = ["Manter", "Eliminar", "Reduzir", "Elevar", "Criar"]


def curva_valor(d: dict) -> pd.DataFrame:
    """Notas da concorrência (formato longo) + nova curva proposta pelo ERRC."""
    notas = d["concorrencia_notas"].copy()
    notas = notas[notas["atributo"].astype(str).str.strip() != ""]
    errc = d.get("errc", pd.DataFrame(columns=["atributo", "acao", "nova_nota", "justificativa"]))
    novos = errc[errc["atributo"].astype(str).str.strip() != ""]
    linhas = [{"Atributo": r["atributo"], "Empresa": r["empresa"], "Nota": float(r["nota"])} for _, r in notas.iterrows()]
    nossa = {r["atributo"]: float(r["nota"]) for _, r in notas[notas["empresa"] == "Nossa Empresa"].iterrows()}
    for _, r in novos.iterrows():
        a = r["atributo"]
        if str(r["acao"]) == "Eliminar":
            v = 0.0
        elif float(r["nova_nota"] or 0) > 0 or str(r["acao"]) in ("Criar", "Reduzir"):
            v = float(r["nova_nota"] or 0)
        else:
            v = nossa.get(a, 0.0)
        linhas.append({"Atributo": a, "Empresa": "Nova curva (ERRC)", "Nota": v})
    for a, v in nossa.items():
        if a not in set(novos["atributo"]):
            linhas.append({"Atributo": a, "Empresa": "Nova curva (ERRC)", "Nota": v})
    df = pd.DataFrame(linhas, columns=["Atributo", "Empresa", "Nota"])
    ordem = list(dict.fromkeys(list(notas["atributo"]) + list(novos["atributo"])))
    df["ordem"] = df["Atributo"].map({a: i for i, a in enumerate(ordem)})
    return df.sort_values(["ordem", "Empresa"]).drop(columns="ordem").reset_index(drop=True)


def leitura_errc(d: dict) -> str:
    errc = d.get("errc", pd.DataFrame())
    if errc.empty:
        return ""
    cont = Counter(str(a) for a in errc["acao"])
    if cont.get("Criar", 0) == 0 and cont.get("Eliminar", 0) + cont.get("Reduzir", 0) == 0:
        return ("A nova curva só eleva ou mantém atributos: isso tende a aumentar custos sem diferenciar. "
                "Procure o que eliminar ou reduzir (o que o cliente não valoriza) e algo novo para criar.")
    if cont.get("Criar", 0) == 0:
        return "Há cortes de atributos, mas nada novo foi criado: que fator, que o setor não oferece, poderia atrair não clientes?"
    return (f"A nova curva elimina ou reduz {cont.get('Eliminar', 0) + cont.get('Reduzir', 0)} atributo(s), eleva "
            f"{cont.get('Elevar', 0)} e cria {cont.get('Criar', 0)}: é o movimento de diferenciação com custo controlado "
            "proposto pela estratégia do oceano azul.")


def frase_posicionamento(proj: dict) -> str:
    c = {k: str(proj.get(k) or "").strip() for k in ("posic_segmento", "posic_necessidade", "posic_marca", "posic_categoria",
                                                       "posic_beneficio", "posic_concorrente", "posic_diferencial")}
    if not (c["posic_segmento"] and c["posic_beneficio"]):
        return ""
    return (f"Para {c['posic_segmento']} que {c['posic_necessidade'] or '...'}, {c['posic_marca'] or proj.get('nome_negocio', '')} "
            f"é {c['posic_categoria'] or '...'} que {c['posic_beneficio']}. Diferentemente de {c['posic_concorrente'] or '...'}, "
            f"{c['posic_diferencial'] or '...'}.")


def orcamento_canais(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    oc = d.get("orcamento_canais", pd.DataFrame())
    oc = oc[oc["canal"].astype(str).str.strip() != ""] if not oc.empty else oc
    if oc.empty:
        return pd.DataFrame()
    rec = res.get("recorrencia", {})
    valor = float(rec.get("ltv") or 0) if rec.get("tem") and rec.get("ltv", 0) == rec.get("ltv", 0) else 0.0
    rotulo = "LTV"
    if valor <= 0:
        valor, rotulo = float(proj.get("gasto_anual_cliente") or 0), "gasto anual"
    out = pd.DataFrame({"Canal": oc["canal"], "Gasto mensal": oc["gasto_mensal"], "Clientes novos/mês": oc["clientes_mes"]})
    out["CAC"] = [div(g, c) if c > 0 else float("nan") for g, c in zip(out["Gasto mensal"], out["Clientes novos/mês"])]
    out["Valor do cliente"] = valor
    out["Retorno por R$ 1"] = [div(c * valor, g) if g > 0 else float("nan") for g, c in zip(out["Gasto mensal"], out["Clientes novos/mês"])]
    out.attrs["rotulo_valor"] = rotulo
    return out


ETAPAS_AARRR = [("Aquisição", "Como as pessoas descobrem o negócio?"),
                ("Ativação", "Qual é a primeira boa experiência do cliente?"),
                ("Retenção", "O que faz o cliente voltar ou continuar assinando?"),
                ("Receita", "Como e quando o cliente paga?"),
                ("Indicação", "Por que o cliente indicaria o negócio a outras pessoas?")]


# ------------------------------------------------------------------------------
# 4. INCERTEZA: MONTE CARLO, BASS E PERDA ACEITÁVEL
# ------------------------------------------------------------------------------
VARIAVEIS_MC = [("Preço de venda", -10.0, 5.0), ("Volume de vendas", -30.0, 15.0),
                ("Custo de materiais e mercadorias", -5.0, 15.0), ("Salários e encargos", 0.0, 10.0),
                ("Demais custos fixos", -5.0, 15.0), ("Crescimento anual (pontos percentuais)", -5.0, 3.0)]


def faixas_mc(d: dict) -> pd.DataFrame:
    f = d.get("mc_faixas", pd.DataFrame(columns=["variavel", "minimo_pct", "maximo_pct"]))
    mapa = {str(r["variavel"]): (float(r["minimo_pct"]), float(r["maximo_pct"])) for _, r in f.iterrows()}
    return pd.DataFrame([{"variavel": v, "minimo_pct": mapa.get(v, (a, b))[0], "maximo_pct": mapa.get(v, (a, b))[1]}
                         for v, a, b in VARIAVEIS_MC])


def monte_carlo(res: dict, faixas: pd.DataFrame, n: int = 3000, seed: int = 42) -> dict:
    """
    Sorteia variações (distribuição triangular: mínimo, 0 = provável, máximo) e recalcula o VPL do projeto
    e o menor saldo de caixa do Ano 1 com a mesma lógica do motor (fluxo do projeto = EBITDA).
    """
    rng = np.random.default_rng(seed)
    fx = {r["variavel"]: (float(r["minimo_pct"]) / 100, float(r["maximo_pct"]) / 100) for _, r in faixas.iterrows()}

    def sorteio(nome):
        a, b = fx.get(nome, (0.0, 0.0))
        a, b = min(a, 0.0), max(b, 0.0)
        if a == b:
            return np.zeros(n)
        return rng.triangular(a, 0.0, b, n)

    fp, fv = 1 + sorteio("Preço de venda"), 1 + sorteio("Volume de vendas")
    fm, fs = 1 + sorteio("Custo de materiais e mercadorias"), 1 + sorteio("Salários e encargos")
    ff, dg = 1 + sorteio("Demais custos fixos"), sorteio("Crescimento anual (pontos percentuais)")
    g_a, tma, N = res["g_a"], res["tma"], res["N"]
    adj = (1 + g_a + dg) / (1 + g_a)
    vpl = -res["investimento_total"] * np.ones(n)
    fluxo_anos = []
    for k in res["ANOS"]:
        a = res["anos"][k]
        gk = adj ** (k - 1)
        rec = a["rec_bruta"] * fp * fv * gk
        var_ratio = div(a["impostos"] + a["outras_var"], a["rec_bruta"]) if a["rec_bruta"] else 0.0
        ebitda = rec - a["custo_dir"] * fv * fm * gk - rec * var_ratio - a["pessoal"] * fs - a["fixos"] * ff
        if k == N and res.get("recupera_giro"):
            ebitda = ebitda + res.get("capital_giro", 0.0)
        fluxo_anos.append(ebitda)
        vpl = vpl + ebitda / (1 + tma) ** k
    # caixa mês a mês do Ano 1
    fm_df = res["fluxo_mensal"]
    a1 = res["anos"][1]
    w_s = div(a1["pessoal"], a1["pessoal"] + a1["fixos"]) if (a1["pessoal"] + a1["fixos"]) else 0.5
    saldo = float(fm_df["Saldo inicial"].iloc[0]) * np.ones(n)
    minimo = np.full(n, np.inf)
    for _, r in fm_df.iterrows():
        saldo = saldo + r["Recebimentos"] * fp * fv - r["Fornecedores"] * fv * fm - r["Impostos e desp. variáveis"] * fp * fv \
            - r["Pessoal e custos fixos"] * (w_s * fs + (1 - w_s) * ff) - r["Parcelas de empréstimos"]
        minimo = np.minimum(minimo, saldo)
    return {"vpl": vpl, "caixa_min": minimo, "n": n,
            "p_vpl_pos": float((vpl > 0).mean()), "p_caixa_neg": float((minimo < 0).mean()),
            "p5": float(np.percentile(vpl, 5)), "p50": float(np.percentile(vpl, 50)), "p95": float(np.percentile(vpl, 95)),
            "caixa_p5": float(np.percentile(minimo, 5)), "vpl_medio": float(vpl.mean())}


def leitura_mc(mc: dict) -> str:
    p = mc["p_vpl_pos"]
    if p >= 0.8:
        txt = "O projeto é robusto: na grande maioria dos cenários sorteados ele cria valor."
    elif p >= 0.6:
        txt = "O projeto tende a criar valor, mas há uma chance relevante de fracasso: proteja-se das variáveis mais sensíveis."
    elif p >= 0.4:
        txt = "Resultado incerto, quase uma aposta: valide as premissas críticas antes de investir."
    else:
        txt = "Na maior parte dos cenários o projeto destrói valor: o modelo de negócio precisa ser revisto."
    if mc["p_caixa_neg"] > 0.2:
        txt += f" Em {pct(mc['p_caixa_neg'], 0)} dos cenários o caixa fica negativo no Ano 1: aumente a reserva ou o capital de giro."
    return txt


def bass_curva(m: float, p: float, q: float, anos: int) -> pd.DataFrame:
    """Modelo de difusão de Bass (1969): adotantes acumulados ao fim de cada ano."""
    linhas = []
    for t in range(1, anos + 1):
        F = (1 - math.exp(-(p + q) * t)) / (1 + (q / p) * math.exp(-(p + q) * t)) if p > 0 else 0.0
        linhas.append({"Ano": t, "Adotantes acumulados": m * F, "% do mercado potencial": F})
    df = pd.DataFrame(linhas)
    df["Novos adotantes no ano"] = df["Adotantes acumulados"].diff().fillna(df["Adotantes acumulados"])
    return df


def bass_comparacao(proj: dict, res: dict) -> dict:
    m = float(proj.get("bass_mercado") or 0) or float(proj.get("sam_clientes") or 0)
    p, q = float(proj.get("bass_p") or 0.03), float(proj.get("bass_q") or 0.38)
    gasto = float(proj.get("gasto_anual_cliente") or 0)
    if m <= 0:
        return {"ok": False}
    cv = bass_curva(m, p, q, res["N"])
    if gasto > 0:
        cv["Clientes no plano"] = [div(res["anos"][k]["rec_bruta"], gasto) for k in res["ANOS"]]
        cv["Plano ÷ Bass"] = [div(a, b) for a, b in zip(cv["Clientes no plano"], cv["Adotantes acumulados"])]
    return {"ok": True, "m": m, "p": p, "q": q, "tabela": cv, "tem_plano": gasto > 0}


def perda_aceitavel(proj: dict, d: dict, res: dict) -> dict:
    soc = d["socios"]
    aceit = float(soc["perda_aceitavel"].sum()) if "perda_aceitavel" in soc else 0.0
    proprio = float(soc["capital_integralizado"].sum())
    divida = float(d["emprestimos"]["valor"].sum())
    revenda = float(proj.get("revenda_ativos_pct") if proj.get("revenda_ativos_pct") is not None else 40) / 100
    inv_fixo = float((d["inv_fixos"]["quantidade"] * d["inv_fixos"]["valor_unitario"]).sum())
    potencial = max(proprio + divida - inv_fixo * revenda, 0.0)
    return {"aceitavel": aceit, "proprio": proprio, "divida": divida, "revenda": inv_fixo * revenda,
            "potencial": potencial, "ok": aceit > 0, "cabe": aceit >= potencial}


# ------------------------------------------------------------------------------
# 5. DEFESA: FORMALIZAÇÃO, PRONTIDÃO E BANCA SIMULADA
# ------------------------------------------------------------------------------
LIMITE_MEI, LIMITE_ME, LIMITE_EPP = 81_000.0, 360_000.0, 4_800_000.0
STATUS_FORMAL = ["Pendente", "Em andamento", "Concluído", "Não se aplica"]


def itens_formalizacao(proj: dict, d: dict) -> list:
    tipo = str(proj.get("tipo_negocio") or "")
    setor = (str(proj.get("setor_atuacao") or "") + " " + str(proj.get("nome_negocio") or "")).lower()
    itens = ["Escolher a natureza jurídica (MEI, empresário individual, SLU ou LTDA)",
             "Definir as atividades (CNAE) e conferir se são permitidas no Simples/MEI",
             "Consultar a viabilidade do endereço na prefeitura (zoneamento)",
             "Registrar a empresa (Junta Comercial / Redesim) e obter o CNPJ",
             "Inscrição municipal e alvará de funcionamento",
             "Optar pelo Simples Nacional (prazo no início da atividade)",
             "Contratar contador ou escritório contábil",
             "Abrir conta bancária PJ e meios de pagamento",
             "Contrato social com acordo de sócios (entrada, saída, divisão de lucros)",
             "Registrar a marca no INPI"]
    if any(w in setor for w in ("aliment", "café", "cafe", "restaurante", "lanche", "refei", "fit", "padaria", "bebida")):
        itens.append("Licença sanitária (Vigilância Sanitária) e boas práticas de manipulação")
    if float(proj.get("loc_area_m2") or 0) > 0 or proj.get("loc_endereco"):
        itens.append("Auto de vistoria / licença do Corpo de Bombeiros")
    if not d["pessoal"].empty:
        itens.append("Registro de empregados (eSocial), exames admissionais e segurança do trabalho")
    if tipo.startswith(("Assinatura", "Serviço", "Misto")) or "app" in setor or "aplicativo" in setor:
        itens += ["LGPD: política de privacidade, bases legais e encarregado de dados",
                  "Termos de uso e contrato de assinatura (renovação, cancelamento, reembolso)"]
    return itens


def porte_sugerido(res: dict, d: dict) -> dict:
    rec = res["anos"][1]["rec_bruta"]
    n_soc = len(d["socios"])
    n_emp = float(d["pessoal"]["num_empregados"].sum()) if not d["pessoal"].empty else 0
    if rec <= LIMITE_MEI and n_soc <= 1 and n_emp <= 1:
        porte = "MEI (se a atividade for permitida)"
    elif rec <= LIMITE_ME:
        porte = "Microempresa (ME)"
    elif rec <= LIMITE_EPP:
        porte = "Empresa de Pequeno Porte (EPP)"
    else:
        porte = "Acima do limite do Simples Nacional (Lucro Presumido ou Real)"
    obs = []
    if rec <= LIMITE_MEI and n_soc > 1:
        obs.append("O MEI não admite sócios: com mais de um sócio, o caminho é ME (LTDA).")
    return {"porte": porte, "receita": rec, "obs": obs}


NIVEIS_PRONTIDAO = [
    (1, "Ideia descrita", "Missão, proposta de valor e segmento definidos."),
    (2, "Mercado dimensionado", "TAM, SAM e SOM estimados e meta de vendas coerente com o SOM."),
    (3, "Hipóteses mapeadas", "Canvas, premissas listadas e status de evidência informado."),
    (4, "Problema validado", "Pelo menos 10 entrevistas e uma dor citada por metade ou mais dos entrevistados."),
    (5, "Solução e preço validados", "Preço dentro da faixa aceitável (Van Westendorp) ou 40% ou mais dispostos a pagar."),
    (6, "Modelo econômico viável", "VPL positivo e caixa do Ano 1 sem ficar negativo."),
    (7, "Viabilidade robusta", "Chance de VPL positivo de 70% ou mais (Monte Carlo) e placar de evidências de 50% ou mais."),
    (8, "Plano de execução", "Objetivos SMART, cronograma, riscos e 60% ou mais do checklist de formalização encaminhado."),
    (9, "Pronto para lançar ou buscar investimento", "Placar de 75% ou mais, perda aceitável coberta e perguntas da banca respondidas."),
]


def avaliar_prontidao(proj: dict, d: dict, res: dict, mc: dict | None, preco_ref: float) -> dict:
    a1 = res["anos"][1]
    pl = placar_evidencias(proj, d, res)
    ent = resumo_entrevistas(d, preco_ref)
    vw = van_westendorp(d.get("precos_vw"))
    m = mercado(proj, res)
    fo = d.get("formalizacao", pd.DataFrame())
    fo_ok = div(float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).sum()), len(fo)) if len(fo) else 0.0
    perda = perda_aceitavel(proj, d, res)
    resp = d.get("banca_respostas", pd.DataFrame())
    n_resp = int((resp["resposta"].astype(str).str.len() > 20).sum()) if not resp.empty else 0
    preco_ok = False
    if vw and vw.get("ok") and preco_ref == preco_ref:
        preco_ok = vw["pmc"] <= preco_ref <= vw["pme"]
    cond = {
        1: bool(proj.get("missao")) and not d["propostas"].empty and not d["segmentos"].empty,
        2: m["preenchido"] and (m["pct_som"] != m["pct_som"] or m["pct_som"] <= 1.0),
        3: (d.get("evidencias") is not None and len(d["evidencias"]) >= 3),
        4: ent["n"] >= 10 and ent.get("dor_top", ("", 0))[1] >= ent["n"] / 2,
        5: preco_ok or (ent["n"] >= 10 and ent.get("pct_sim", 0) >= 0.4),
        6: res["vpl"] > 0 and float(res["fluxo_mensal"]["Saldo final"].min()) >= 0,
        7: mc is not None and mc["p_vpl_pos"] >= 0.7 and pl["score"] >= 0.5,
        8: len(d["objetivos"]) >= 3 and len(d["cronograma"]) >= 3 and len(d["riscos"]) >= 3 and fo_ok >= 0.6,
        9: pl["score"] >= 0.75 and perda["ok"] and perda["cabe"] and n_resp >= 5,
    }
    nivel = 0
    for k in range(1, 10):
        if cond[k]:
            nivel = k
        else:
            break
    return {"nivel": nivel, "cond": cond, "a_rec": a1["rec_bruta"]}


def perguntas_banca(proj: dict, d: dict, res: dict, mc: dict | None, preco_ref: float) -> list:
    """Perguntas que uma banca faria, geradas a partir das fragilidades do próprio plano."""
    P = []

    def add(tema, pergunta, motivo):
        P.append({"tema": tema, "pergunta": pergunta, "motivo": motivo})

    a1 = res["anos"][1]
    rec = a1["rec_bruta"]
    m = mercado(proj, res)
    pl = placar_evidencias(proj, d, res)
    ent = resumo_entrevistas(d, preco_ref)
    vw = van_westendorp(d.get("precos_vw"))
    rr = res.get("recorrencia", {})
    if pl["score"] < 0.5:
        add("Evidências", f"Quais números do plano foram testados com clientes reais? Por que devemos acreditar em "
                          f"\"{pl['proxima']}\"?", f"Placar de evidências em {pct(pl['score'], 0)}.")
    if ent["n"] < 10:
        add("Cliente", "Com quantos clientes do segmento-alvo vocês conversaram? O que aprenderam que mudou o plano?",
            f"{ent['n']} entrevista(s) registrada(s).")
    if vw and vw.get("ok") and preco_ref == preco_ref and not (vw["pmc"] <= preco_ref <= vw["pme"]):
        lado = "acima" if preco_ref > vw["pme"] else "abaixo"
        add("Preço", f"O preço médio ({brl(preco_ref)}) está {lado} da faixa que os clientes consideram aceitável "
                     f"({brl(vw['pmc'])} a {brl(vw['pme'])}). Como justificam?", "Teste de Van Westendorp.")
    if m["preenchido"] and m["pct_som"] == m["pct_som"] and m["pct_som"] > 1:
        add("Mercado", "A receita do Ano 1 é maior que o mercado que vocês dizem conseguir conquistar (SOM). De onde vêm esses clientes?",
            f"Receita = {pct(m['pct_som'], 0)} do SOM.")
    if res["vpl"] <= 0:
        add("Viabilidade", "O VPL é negativo. O que mudaria no modelo de negócio para torná-lo viável, e por que ainda vale a pena?",
            f"VPL de {brl(res['vpl'])}.")
    tir = res.get("tir")
    if tir == tir and tir is not None and res["vpl"] > 0 and tir < res["tma"] * 1.3:
        add("Retorno", "A TIR é pouco maior que a TMA. Por que não deixar o dinheiro aplicado, com menos risco?",
            f"TIR {pct(tir)} x TMA {pct(res['tma'])}.")
    minimo = float(res["fluxo_mensal"]["Saldo final"].min())
    if minimo < 0:
        add("Caixa", "O caixa fica negativo no primeiro ano. Quem cobre esse buraco e em que mês?", f"Menor saldo: {brl(minimo)}.")
    if a1["imc"] == a1["imc"] and a1["imc"] < 0.3:
        add("Margem", "A margem de contribuição é baixa. O que acontece se o fornecedor aumentar o preço em 10%?",
            f"Margem de contribuição de {pct(a1['imc'])}.")
    if rec > 0 and div(a1["pessoal"], rec) > 0.45:
        add("Equipe", "A folha consome quase metade da receita. A equipe não está grande demais para o início?",
            f"Pessoal = {pct(div(a1['pessoal'], rec), 0)} da receita.")
    if rr.get("tem"):
        if rr.get("churn", 0) > 0.05:
            add("Retenção", "O churn mensal passa de 5%. O que fará o assinante continuar depois do terceiro mês?",
                f"Churn de {pct(rr['churn'])} ao mês.")
        lc = rr.get("ltv_cac", float("nan"))
        if lc == lc and lc < 3:
            add("Aquisição", "O LTV é menor que três vezes o CAC. O crescimento não vai queimar caixa?", f"LTV/CAC = {num_br(lc, 1)}.")
    if mc is not None and mc["p_vpl_pos"] < 0.6:
        add("Risco", "Em quase metade das simulações o negócio não se paga. Qual é o plano B?",
            f"Chance de VPL positivo: {pct(mc['p_vpl_pos'], 0)}.")
    perda = perda_aceitavel(proj, d, res)
    if perda["ok"] and not perda["cabe"]:
        add("Sócios", "Se der errado, vocês podem perder mais do que disseram aceitar. Como reduzir o investimento inicial?",
            f"Perda potencial {brl(perda['potencial'])} x aceitável {brl(perda['aceitavel'])}.")
    errc = d.get("errc", pd.DataFrame())
    if errc.empty or not (errc["acao"] == "Criar").any():
        add("Diferenciação", "O que vocês oferecem que nenhum concorrente oferece? Por que um concorrente grande não copiaria amanhã?",
            "Nenhum atributo novo na curva de valor.")
    forn = d["fornecedores"]
    if not forn.empty and int(forn["critico"].sum()) >= 1 and len(forn) <= 2:
        add("Operação", "Há dependência de poucos fornecedores críticos. O que acontece se o principal falhar?", "Fornecedores críticos.")
    add("Timing", "Por que este é o momento certo para abrir este negócio?", "Pergunta clássica de banca.")
    add("Equipe", "Por que este grupo é o mais capaz de executar este plano?", "Pergunta clássica de banca.")
    return P


def preco_referencia(proj: dict, res: dict) -> float:
    """Preço do produto escolhido no teste de preço; sem escolha, preço médio ponderado (ou ARPU das assinaturas)."""
    cat = res.get("cat", pd.DataFrame())
    alvo = str(proj.get("vw_produto") or "")
    if alvo and not cat.empty and alvo in set(cat["descricao"]):
        return float(cat.loc[cat["descricao"] == alvo, "preco_venda"].iloc[0])
    if alvo.startswith("Assinatura") or (cat.empty and res.get("recorrencia", {}).get("tem")):
        return float(res["recorrencia"].get("arpu", float("nan")))
    if cat.empty or float(cat["qtd_mensal"].sum()) <= 0:
        return float("nan")
    return div(float((cat["preco_venda"] * cat["qtd_mensal"]).sum()), float(cat["qtd_mensal"].sum()))


def regras_v71(proj: dict, d: dict) -> dict:
    """Situação de preenchimento das ferramentas novas: (completo, iniciado)."""
    def n(t):
        x = d.get(t)
        return 0 if x is None else len(x)
    pos = [proj.get(k) for k in ("posic_segmento", "posic_beneficio", "posic_diferencial")]
    vpc = [proj.get(k) for k, _, _ in BLOCOS_VPC]
    ev = d.get("evidencias", pd.DataFrame(columns=["status"]))
    fo = d.get("formalizacao", pd.DataFrame(columns=["status"]))
    resp = d.get("banca_respostas", pd.DataFrame(columns=["resposta"]))
    n_resp = int((resp["resposta"].astype(str).str.len() > 20).sum()) if len(resp) else 0
    errc = d.get("errc", pd.DataFrame(columns=["acao"]))
    rec = n("catalogo") > 0 or n("assinaturas") > 0
    return {
        "entrevistas": (n("entrevistas") >= 10, n("entrevistas") >= 1),
        "vpc": (all(vpc), any(vpc)),
        "oceano": (n("errc") >= 3 and bool((errc["acao"] == "Criar").any()), n("errc") >= 1),
        "posicionamento": (all(pos) and n("mapa_perceptual") >= 2, any(pos) or n("mapa_perceptual") >= 1),
        "orcamento_mkt": (n("orcamento_canais") >= 1 and n("aarrr") >= 5, n("orcamento_canais") >= 1 or n("calendario_mkt") >= 1),
        "teste_preco": (n("precos_vw") >= 5 or float(proj.get("intencao_certamente_pct") or 0) > 0,
                        n("precos_vw") >= 1 or float(proj.get("intencao_certamente_pct") or 0) > 0),
        "montecarlo": (rec and float(d["socios"]["perda_aceitavel"].sum()) > 0, rec),
        "adocao": (float(proj.get("sam_clientes") or 0) > 0 and float(proj.get("gasto_anual_cliente") or 0) > 0,
                   float(proj.get("sam_clientes") or 0) > 0 or float(proj.get("bass_mercado") or 0) > 0),
        "evidencias": (len(ev) >= 3 and bool((ev["status"] != "Estimada (sem fonte)").any()), len(ev) >= 1),
        "banca": (n_resp >= 5, n_resp >= 1),
        "formalizacao": (len(fo) > 0 and float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).mean()) >= 0.6,
                         len(fo) >= 1),
    }

# ##############################################################################
# PARTE: pnbox_v8.py
# ##############################################################################




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

# ##############################################################################
# PARTE: pnbox_mix.py
# ##############################################################################






def _n(x: float, casas: int = 0) -> str:
    """Número no formato brasileiro (1234,5 -> 1.234,5)."""
    t = f"{x:,.{casas}f}"
    return t.replace(",", "X").replace(".", ",").replace("X", ".")


def _p(x: float, casas: int = 1) -> str:
    """Percentual no formato brasileiro (0,257 -> 25,7%)."""
    return f"{x * 100:.{casas}f}%".replace(".", ",")


# ------------------------------------------------------------------------------
# PRODUTO
# ------------------------------------------------------------------------------
CICLO_VIDA = {
    "Introdução": "Vendas baixas e custo alto de conquistar cliente. Foco em fazer o produto ser conhecido e "
                  "experimentado; o preço costuma ser de penetração ou de desnatação.",
    "Crescimento": "As vendas sobem e surgem concorrentes. Foco em ampliar a distribuição e reforçar a marca.",
    "Maturidade": "O mercado para de crescer e a disputa é por participação. Foco em diferenciar, "
                  "fidelizar e controlar custos; guerra de preço corrói a margem.",
    "Declínio": "As vendas caem. Foco em reduzir custos, enxugar a linha ou reposicionar o produto.",
}


def pareto_produtos(cat: pd.DataFrame) -> dict:
    """Quantos itens respondem por 80% da receita (concentração do portfólio)."""
    if cat is None or cat.empty or cat["receita_mensal"].sum() <= 0:
        return {"n": 0, "total": 0, "itens": []}
    c = cat.sort_values("receita_mensal", ascending=False)
    acum = c["receita_mensal"].cumsum() / c["receita_mensal"].sum()
    n = int((acum < 0.8).sum()) + 1
    return {"n": min(n, len(c)), "total": len(c), "itens": c["descricao"].head(n).tolist(),
            "share_top": float(c["receita_mensal"].iloc[0] / c["receita_mensal"].sum())}


# ------------------------------------------------------------------------------
# PREÇO
# ------------------------------------------------------------------------------
ESTRATEGIAS_PRECO = {
    "Penetração": ("Preço abaixo do mercado para ganhar clientes depressa. Exige volume alto e custo baixo.", "abaixo"),
    "Desnatação": ("Preço alto no lançamento, para quem valoriza a novidade; cai depois.", "acima"),
    "Valor percebido": ("Preço definido pelo benefício que o cliente enxerga, medido em pesquisa.", None),
    "Custo mais margem": ("Custo unitário mais uma margem (markup). Simples, mas ignora o cliente e a concorrência.", None),
    "Referência na concorrência": ("Preço alinhado ao praticado pelos concorrentes diretos.", "igual"),
}
POSICOES_PRECO = {
    "Econômico": (0.0, 0.95),
    "Intermediário": (0.90, 1.10),
    "Premium": (1.05, 99.0),
}


def mc_produtos(res: dict) -> pd.DataFrame:
    """Margem de contribuição percentual de cada item: preço − custo direto − impostos − outros variáveis."""
    cat = res.get("cat")
    if cat is None or cat.empty:
        return pd.DataFrame(columns=["Produto", "Preço", "Custo direto", "MC %", "Peso na receita"])
    aliq, var = float(res.get("aliq") or 0), float(res.get("var_pct") or 0)
    tot = float(cat["receita_mensal"].sum()) or 1.0
    return pd.DataFrame({
        "Produto": cat["descricao"], "Preço": cat["preco_venda"], "Custo direto": cat["custo_unit"],
        "MC %": [div(p - c, p) - aliq - var if p > 0 else float("nan")
                 for p, c in zip(cat["preco_venda"], cat["custo_unit"])],
        "Peso na receita": cat["receita_mensal"] / tot,
    })


def mc_media(res: dict) -> float:
    mc = mc_produtos(res)
    if mc.empty:
        return float("nan")
    ok = mc.dropna(subset=["MC %"])
    return float((ok["MC %"] * ok["Peso na receita"]).sum() / ok["Peso na receita"].sum()) if ok["Peso na receita"].sum() > 0 else float("nan")


def volume_para_desconto(mc: float, desconto: float) -> float:
    """Aumento de volume necessário para manter a margem total após um desconto (frações)."""
    if mc != mc or mc <= 0 or desconto >= mc:
        return float("inf")
    return desconto / (mc - desconto)


def queda_tolerada_aumento(mc: float, aumento: float) -> float:
    """Quanto o volume pode cair, após um aumento de preço, sem reduzir a margem total (frações)."""
    if mc != mc or mc + aumento <= 0:
        return float("nan")
    return aumento / (mc + aumento)


def comparar_concorrentes(cat: pd.DataFrame, precos: pd.DataFrame) -> pd.DataFrame:
    """Preço próprio x média dos concorrentes, produto a produto."""
    if cat is None or cat.empty or precos is None or precos.empty:
        return pd.DataFrame(columns=["Produto", "Seu preço", "Média dos concorrentes", "Menor", "Maior", "Índice"])
    p = precos[(precos["preco"] > 0) & (precos["produto"].astype(str).str.strip() != "")]
    linhas = []
    for prod, g in p.groupby("produto"):
        meu = cat.loc[cat["descricao"] == prod, "preco_venda"]
        if meu.empty:
            continue
        m = float(meu.iloc[0])
        media = float(g["preco"].mean())
        linhas.append({"Produto": prod, "Seu preço": m, "Média dos concorrentes": media,
                       "Menor": float(g["preco"].min()), "Maior": float(g["preco"].max()),
                       "Índice": div(m, media)})
    return pd.DataFrame(linhas)


def indice_preco_medio(cat: pd.DataFrame, comp: pd.DataFrame) -> float:
    """Índice de preço (seu preço ÷ média dos concorrentes), ponderado pela receita dos itens comparados."""
    if comp.empty:
        return float("nan")
    pesos = [float(cat.loc[cat["descricao"] == p, "receita_mensal"].sum()) for p in comp["Produto"]]
    if sum(pesos) <= 0:
        return float(comp["Índice"].mean())
    return float(sum(i * w for i, w in zip(comp["Índice"], pesos)) / sum(pesos))


def checar_produto_preco(proj: dict, d: dict, res: dict, vw: dict | None = None) -> list[tuple[str, str]]:
    """Alertas de Produto e Preço. Cada item é (nível, mensagem) com nível 'ok', 'atencao' ou 'critico'."""
    out = []
    cat = res.get("cat")
    if cat is None or cat.empty:
        return [("atencao", "Cadastre os produtos ou serviços para que o programa confira preço e margem.")]
    # Margens
    mc = mc_produtos(res)
    neg = mc[mc["MC %"] <= 0]["Produto"].tolist()
    if neg:
        out.append(("critico", "Itens vendidos abaixo do custo variável: " + ", ".join(neg) +
                    ". Cada venda deles aumenta o prejuízo."))
    mm = mc_media(res)
    if mm == mm:
        nivel = "ok" if mm >= 0.30 else ("atencao" if mm >= 0.15 else "critico")
        out.append((nivel, f"Margem de contribuição média de {_p(mm)}. "
                    + ("Há folga para custos fixos e descontos." if nivel == "ok" else
                       "Pouca folga: qualquer desconto ou aumento de custo pesa muito." if nivel == "atencao" else
                       "Margem muito estreita: o negócio depende de volume muito alto para pagar os custos fixos.")))
    # Concorrência e posição de preço
    comp = comparar_concorrentes(cat, d.get("precos_concorrentes"))
    idx = indice_preco_medio(cat, comp)
    pos = proj.get("mix_posicao_preco") or ""
    est = proj.get("mix_estrategia_preco") or ""
    if idx == idx:
        txt = f"Seus preços estão, em média, {_p(abs(idx - 1), 0)} {'acima' if idx >= 1 else 'abaixo'} dos concorrentes pesquisados."
        if pos in POSICOES_PRECO:
            lo, hi = POSICOES_PRECO[pos]
            if not (lo <= idx <= hi):
                out.append(("atencao", txt + f" Isso não combina com o posicionamento {pos.lower()} declarado."))
            else:
                out.append(("ok", txt + f" Coerente com o posicionamento {pos.lower()}."))
        else:
            out.append(("ok", txt))
        esperado = ESTRATEGIAS_PRECO.get(est, ("", None))[1]
        if esperado == "abaixo" and idx >= 1:
            out.append(("atencao", "A estratégia é de penetração, mas os preços não estão abaixo da concorrência."))
        if esperado == "acima" and idx <= 1:
            out.append(("atencao", "A estratégia é de desnatação, mas os preços não estão acima da concorrência."))
        if esperado == "igual" and not (0.9 <= idx <= 1.1):
            out.append(("atencao", "A estratégia é seguir a concorrência, mas a diferença passa de 10%."))
    else:
        out.append(("atencao", "Sem preços de concorrentes cadastrados: a comparação de mercado não pôde ser feita."))
    if est == "Penetração" and mm == mm and mm < 0.25:
        out.append(("atencao", "Penetração com margem estreita é arriscada: confira se o volume previsto é realista."))
    # Van Westendorp
    if vw and vw.get("ok"):
        alvo = proj.get("vw_produto") or ""
        linha = cat[cat["descricao"] == alvo]
        preco = float(linha["preco_venda"].iloc[0]) if not linha.empty else float("nan")
        if preco == preco and vw["pmc"] == vw["pmc"] and vw["pme"] == vw["pme"]:
            if preco < vw["pmc"]:
                out.append(("atencao", f"O preço de {alvo} está abaixo da faixa aceitável da pesquisa: pode passar imagem de baixa qualidade."))
            elif preco > vw["pme"]:
                out.append(("critico", f"O preço de {alvo} está acima da faixa aceitável da pesquisa: muitos clientes o acham caro demais."))
            else:
                out.append(("ok", f"O preço de {alvo} está dentro da faixa aceitável da pesquisa de preço."))
    elif est == "Valor percebido":
        out.append(("atencao", "A estratégia é de valor percebido, mas não há pesquisa de preço (Van Westendorp) com respostas suficientes."))
    # Ciclo de vida
    ciclo = proj.get("mix_ciclo") or ""
    if ciclo == "Maturidade" and est == "Desnatação":
        out.append(("atencao", "Desnatação funciona no lançamento; num mercado maduro o cliente já conhece as alternativas."))
    if ciclo == "Declínio" and est == "Penetração":
        out.append(("atencao", "Penetração num mercado em declínio dificilmente compensa: o volume total está caindo."))
    # Concentração
    pt = pareto_produtos(cat)
    if pt["total"] >= 3 and pt["n"] == 1:
        out.append(("atencao", f"Um único item ({pt['itens'][0]}) responde por {_p(pt['share_top'], 0)} da receita: "
                    "o plano fica muito dependente dele."))
    return out


# ------------------------------------------------------------------------------
# PREÇO: CONDIÇÕES DE PAGAMENTO
# ------------------------------------------------------------------------------
PAGAMENTO_EXEMPLO = pd.DataFrame([
    {"forma": "Pix", "pct_vendas": 35.0, "taxa_pct": 0.0, "prazo_dias": 0.0},
    {"forma": "Dinheiro", "pct_vendas": 10.0, "taxa_pct": 0.0, "prazo_dias": 0.0},
    {"forma": "Cartão de débito", "pct_vendas": 20.0, "taxa_pct": 1.5, "prazo_dias": 1.0},
    {"forma": "Cartão de crédito à vista", "pct_vendas": 25.0, "taxa_pct": 3.5, "prazo_dias": 30.0},
    {"forma": "Cartão de crédito parcelado", "pct_vendas": 10.0, "taxa_pct": 6.0, "prazo_dias": 60.0},
])


def resumo_pagamento(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"soma": 0.0, "taxa": 0.0, "prazo": 0.0}
    w = df["pct_vendas"].astype(float)
    soma = float(w.sum())
    taxa = float((w * df["taxa_pct"]).sum() / 100) if soma else 0.0          # % sobre a receita total
    prazo = float((w * df["prazo_dias"]).sum() / soma) if soma else 0.0
    return {"soma": soma, "taxa": taxa, "prazo": prazo}


def aplicar_pagamento(df_pag: pd.DataFrame, cv: pd.DataFrame, prazos: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Atualiza 'Taxas de cartões' nos custos variáveis e substitui os prazos de recebimento."""
    r = resumo_pagamento(df_pag)
    cv = cv.copy()
    mask = cv["descricao"].astype(str).str.lower().str.startswith("taxas de cart")
    if mask.any():
        cv.loc[mask, "percentual"] = 0.0
        cv.loc[cv.index[mask][0], "percentual"] = round(r["taxa"], 4)
    else:
        cv = pd.concat([cv, pd.DataFrame([{"descricao": "Taxas de cartões", "percentual": round(r["taxa"], 4)}])],
                       ignore_index=True)
    g = df_pag[df_pag["pct_vendas"] > 0].groupby("prazo_dias", as_index=False)["pct_vendas"].sum()
    novos = pd.DataFrame({"tipo": "Recebimento", "prazo_dias": g["prazo_dias"], "percentual": g["pct_vendas"]})
    prazos = pd.concat([prazos[prazos["tipo"] != "Recebimento"], novos], ignore_index=True)
    return cv, prazos


# ------------------------------------------------------------------------------
# PRAÇA (v8.2)
# ------------------------------------------------------------------------------
CANAIS_VENDA = ["Loja física", "Entrega própria", "Aplicativo de delivery", "Marketplace", "Site próprio",
                "WhatsApp e redes sociais", "Revendedores", "Venda direta a empresas", "Feiras e eventos"]
DISTRIBUICAO = {
    "Intensiva": "Estar no maior número possível de pontos de venda. Combina com produtos de compra frequente e preço baixo.",
    "Seletiva": "Poucos pontos ou canais escolhidos. Combina com produtos em que o cliente compara antes de comprar.",
    "Exclusiva": "Um ou poucos pontos, com controle total da experiência. Combina com posicionamento premium.",
}
_LINHA_COMISSAO = "Comissões de canais de venda"
TERCEIROS = {"Aplicativo de delivery", "Marketplace", "Revendedores"}


def resumo_canais(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {"soma": 0.0, "comissao": 0.0, "maior": "", "maior_pct": 0.0}
    w = df["pct_vendas"].astype(float)
    soma = float(w.sum())
    i = int(w.values.argmax())
    return {"soma": soma, "comissao": float((w * df["comissao_pct"]).sum() / 100) if soma else 0.0,
            "maior": str(df["canal"].iloc[i]), "maior_pct": float(w.iloc[i])}


def aplicar_canais(df: pd.DataFrame, cv: pd.DataFrame) -> pd.DataFrame:
    """Grava a comissão média ponderada nos custos variáveis, zerando a linha antiga de aplicativos para não contar duas vezes."""
    r = resumo_canais(df)
    cv = cv.copy()
    antiga = cv["descricao"].astype(str).str.lower().str.startswith("comissões de aplicativos")
    cv.loc[antiga, "percentual"] = 0.0
    m = cv["descricao"] == _LINHA_COMISSAO
    if m.any():
        cv.loc[m, "percentual"] = round(r["comissao"], 4)
    else:
        cv = pd.concat([cv, pd.DataFrame([{"descricao": _LINHA_COMISSAO, "percentual": round(r["comissao"], 4)}])],
                       ignore_index=True)
    return cv


# ------------------------------------------------------------------------------
# PROMOÇÃO (v8.2)
# ------------------------------------------------------------------------------
OBJETIVOS_COMUNICACAO = {
    "Tornar a marca conhecida": "Alcance e frequência: anúncios, parcerias, presença em eventos e conteúdo.",
    "Fazer o cliente experimentar": "Amostras, degustação, primeira compra com vantagem e indicação.",
    "Converter em venda": "Oferta clara, prova social, facilidade de pagamento e contato rápido.",
    "Fidelizar e fazer voltar": "Programa de fidelidade, pós-venda, relacionamento por WhatsApp ou e-mail.",
}


def clientes_necessarios(proj: dict, res: dict) -> dict:
    """Clientes ativos por mês que o volume do plano exige (receita do mês 1 ÷ gasto mensal por cliente)."""
    cat = res.get("cat")
    receita = float(res.get("receita_m1") or 0)
    ticket = float(proj.get("mix_ticket") or 0)
    if ticket <= 0 and cat is not None and not cat.empty and cat["qtd_mensal"].sum() > 0:
        ticket = receita / float(cat["qtd_mensal"].sum())
    freq = float(proj.get("mix_freq") or 0)
    if receita <= 0 or ticket <= 0 or freq <= 0:
        return {"ok": False, "ticket": ticket, "freq": freq}
    return {"ok": True, "ticket": ticket, "freq": freq, "clientes": receita / (ticket * freq),
            "gasto_cliente_mes": ticket * freq}


def promocao_vs_volume(proj: dict, d: dict, res: dict) -> dict:
    """Confronta os clientes que o plano exige com os clientes novos que a promoção promete trazer."""
    cn = clientes_necessarios(proj, res)
    oc = d.get("orcamento_canais")
    novos = float(oc["clientes_mes"].sum()) if oc is not None and not oc.empty else 0.0
    gasto = float(oc["gasto_mensal"].sum()) if oc is not None and not oc.empty else 0.0
    out = {"cn": cn, "novos": novos, "gasto": gasto, "cac": div(gasto, novos) if novos > 0 else float("nan")}
    if cn["ok"] and novos > 0:
        out["meses_base"] = cn["clientes"] / novos
    mm = mc_media(res)
    if cn.get("ok") and mm == mm and mm > 0 and out["cac"] == out["cac"]:
        margem_cliente = cn["gasto_cliente_mes"] * mm
        out["payback_cac"] = div(out["cac"], margem_cliente)
    receita = float(res.get("receita_m1") or 0)
    out["verba_receita"] = div(gasto, receita) if receita > 0 else float("nan")
    return out


def checar_praca_promocao(proj: dict, d: dict, res: dict) -> list[tuple[str, str]]:
    out = []
    cv = d.get("canais_venda")
    if cv is None or cv.empty:
        out.append(("atencao", "Os canais de venda (onde o cliente compra e quanto cada canal cobra) ainda não foram informados."))
    else:
        r = resumo_canais(cv)
        if abs(r["soma"] - 100) > 0.5:
            out.append(("atencao", f"Os canais de venda somam {_n(r['soma'])}% das vendas; precisam somar 100%."))
        mm = mc_media(res)
        if r["comissao"] > 0 and mm == mm:
            cvv = d["custos_variaveis"]
            desc = cvv["descricao"].astype(str)
            aplicado = float(cvv.loc[(desc == _LINHA_COMISSAO) | desc.str.lower().str.startswith("comissões de aplicativos"),
                                     "percentual"].sum())
            if abs(aplicado - r["comissao"]) > 0.01:
                out.append(("atencao", f"A comissão média dos canais ({_p(r['comissao'] / 100, 2)} da receita) ainda não foi "
                            "aplicada aos custos variáveis: a margem do plano está maior do que a real."))
        if r["maior_pct"] >= 70 and r["maior"] in TERCEIROS:
            out.append(("atencao", f"{r['maior']} concentra {_n(r['maior_pct'])}% das vendas. Se esse canal mudar as regras ou a taxa, "
                        "o negócio sente direto."))
        pos, dist = proj.get("mix_posicao_preco") or "", proj.get("mix_distribuicao") or ""
        mkt = float(cv.loc[cv["canal"].isin(["Marketplace", "Aplicativo de delivery"]), "pct_vendas"].sum())
        if pos == "Premium" and (dist == "Intensiva" or mkt >= 50):
            out.append(("atencao", "Posição premium com distribuição intensiva ou vendas concentradas em marketplace e aplicativos: "
                        "o cliente compara preço lado a lado e a diferenciação some."))
    pv = promocao_vs_volume(proj, d, res)
    if pv["novos"] <= 0:
        out.append(("atencao", "Não há orçamento de promoção por canal com clientes novos por mês: não dá para saber de onde vem o volume do plano."))
    elif not pv["cn"]["ok"]:
        out.append(("atencao", "Informe quantas vezes por mês o cliente compra para comparar o volume do plano com a promoção."))
    else:
        mb = pv.get("meses_base", float("nan"))
        nivel = "ok" if mb <= 3 else ("atencao" if mb <= 6 else "critico")
        cvn = d.get("canais_venda")
        loja = float(cvn.loc[cvn["canal"] == "Loja física", "pct_vendas"].sum()) if cvn is not None and not cvn.empty else 0.0
        if nivel == "critico" and loja >= 50:
            nivel = "atencao"
        out.append((nivel, f"O plano precisa de cerca de {_n(pv['cn']['clientes'])} clientes ativos por mês, e a promoção traz "
                    f"{_n(pv['novos'])} novos por mês: são uns {_n(mb, 1)} meses para formar essa base, sem contar quem deixa de comprar."
                    + (" O volume do primeiro mês parece otimista." if nivel != "ok" else "")
                    + (" Numa loja de rua parte dos clientes vem do fluxo de pessoas e não da promoção paga; "
                       "estime esse fluxo no plano operacional para sustentar o volume." if loja >= 50 and nivel != "ok" else "")))
        pb = pv.get("payback_cac", float("nan"))
        if pb == pb:
            nivel = "ok" if pb <= 3 else ("atencao" if pb <= 12 else "critico")
            out.append((nivel, f"Cada cliente novo custa R$ {_n(pv['cac'], 2)} e devolve isso em margem de contribuição em cerca de {_n(pb, 1)} meses."))
    return out


# ------------------------------------------------------------------------------
# 7Ps DE SERVIÇOS (v8.3) - Booms; Bitner (1981)
# ------------------------------------------------------------------------------
LINHAS_BLUEPRINT = ["Visível ao cliente", "Bastidores", "Apoio (sistemas e fornecedores)"]


def resumo_blueprint(df: pd.DataFrame) -> dict:
    """Mapa do serviço (Shostack, 1984): tempo total, tempo visível ao cliente e pontos de falha."""
    if df is None or df.empty:
        return {"etapas": 0, "tempo": 0.0, "visivel": 0.0, "falhas": []}
    t = df["tempo_min"].astype(float)
    vis = df["linha"] == "Visível ao cliente"
    falhas = df.loc[df["ponto_falha"].astype(str).str.strip() != "", "etapa"].tolist()
    return {"etapas": len(df), "tempo": float(t.sum()), "visivel": float(t[vis].sum()), "falhas": falhas}


def e_servico(proj: dict, res: dict) -> bool:
    tipo = str(proj.get("tipo_negocio") or "")
    if tipo in ("Serviços", "Assinatura", "Misto"):
        return True
    cat = res.get("cat")
    if cat is not None and not cat.empty and "categoria" in cat:
        return bool((cat["categoria"] == "Serviço").any())
    return False


# ------------------------------------------------------------------------------
# COERÊNCIA DO MIX (v8.3)
# ------------------------------------------------------------------------------
def coerencia_mix(proj: dict, d: dict, res: dict) -> pd.DataFrame:
    """Combinações entre os Ps que se reforçam ou se contradizem. Cada linha: combinação, situação e comentário."""
    pos = proj.get("mix_posicao_preco") or ""
    est = proj.get("mix_estrategia_preco") or ""
    dist = proj.get("mix_distribuicao") or ""
    obj = proj.get("mix_objetivo_com") or ""
    ciclo = proj.get("mix_ciclo") or ""
    ampliado = str(proj.get("mix_ampliado") or "").strip()
    mm = mc_media(res)
    linhas = []

    def add(comb, ok, bom, ruim):
        linhas.append({"Combinação": comb, "Situação": "✅ coerente" if ok else "⚠️ contradição",
                       "Comentário": bom if ok else ruim})

    if pos == "Premium":
        add("Preço premium × produto ampliado", bool(ampliado),
            "O preço mais alto vem acompanhado de serviços e garantias que o justificam.",
            "Preço premium sem nada além do produto básico: o cliente não vê por que pagar mais.")
        if dist:
            add("Preço premium × distribuição", dist != "Intensiva",
                "Distribuição seletiva ou exclusiva protege a imagem e o controle da experiência.",
                "Distribuição intensiva espalha o produto em pontos onde ele disputa preço lado a lado.")
        if est:
            add("Preço premium × estratégia de preço", est != "Penetração",
                "A estratégia de preço sustenta a posição premium.",
                "Penetração é preço baixo para ganhar volume: contradiz a posição premium.")
    if pos == "Econômico":
        if dist:
            add("Preço econômico × distribuição", dist != "Exclusiva",
                "Preço baixo pede volume, e a distribuição ajuda a alcançá-lo.",
                "Preço baixo em poucos pontos exclusivos: falta volume para compensar a margem menor.")
        if mm == mm:
            add("Preço econômico × margem", mm >= 0.20,
                "Mesmo com preço baixo, a margem de contribuição deixa folga para os custos fixos.",
                "Preço econômico com margem estreita: o negócio depende de um volume muito alto.")
    if est == "Penetração" and obj:
        add("Penetração × objetivo de comunicação", obj in ("Tornar a marca conhecida", "Fazer o cliente experimentar"),
            "A comunicação busca alcance e experimentação, que é o que a penetração precisa.",
            "Penetração depende de muita gente conhecendo e experimentando; a comunicação está focada em outra coisa.")
    if est == "Desnatação" and ciclo:
        add("Desnatação × ciclo de vida", ciclo == "Introdução",
            "Preço alto no lançamento, quando o produto ainda é novidade.",
            "Desnatação perde força depois do lançamento, quando o cliente já conhece as alternativas.")
    if ciclo and obj:
        add("Ciclo de vida × objetivo de comunicação",
            not (ciclo == "Introdução" and obj == "Fidelizar e fazer voltar"),
            "O objetivo da comunicação combina com a fase do produto.",
            "Na introdução quase ninguém conhece o produto: antes de fidelizar é preciso ser conhecido e experimentado.")
    cv = d.get("canais_venda")
    if cv is not None and not cv.empty and dist:
        n = int((cv["pct_vendas"] > 0).sum())
        add("Distribuição declarada × canais usados", not (dist == "Intensiva" and n <= 1) and not (dist == "Exclusiva" and n >= 4),
            "Os canais cadastrados batem com a intensidade de distribuição escolhida.",
            "O número de canais cadastrados não bate com a intensidade de distribuição escolhida.")
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------------------
# 4Cs (v8.3) - Lauterborn (1990)
# ------------------------------------------------------------------------------
QUATRO_CS = [
    ("mix_4c_n1", "mix_4c_cliente", "Cliente (em vez de produto)",
     "O produto resolve uma necessidade que o cliente reconhece? Que dor das personas ele alivia?"),
    ("mix_4c_n2", "mix_4c_custo", "Custo para o cliente (em vez de preço)",
     "Quanto custa ao cliente comprar, somando preço, deslocamento, tempo e o risco de não gostar?"),
    ("mix_4c_n3", "mix_4c_conveniencia", "Conveniência (em vez de praça)",
     "É fácil encontrar, comprar e pagar? Horários, entrega, formas de pagamento e atendimento ajudam?"),
    ("mix_4c_n4", "mix_4c_comunicacao", "Comunicação (em vez de promoção)",
     "Há conversa com o cliente, e não só anúncio? Ele consegue perguntar, reclamar e indicar?"),
]


# ------------------------------------------------------------------------------
# PLANO DE MARKETING (v8.4): situação, objetivos, estratégia, táticas, ações e controle
# Smith (2011), McDonald (2007) e Las Casas (2011)
# ------------------------------------------------------------------------------
ANSOFF = {
    "Penetração de mercado": ("Produto atual, mercado atual", "Vender mais do mesmo para o mesmo público: frequência, fidelidade e "
                              "conquista de clientes dos concorrentes. É o caminho de menor risco."),
    "Desenvolvimento de produto": ("Produto novo, mercado atual", "Oferecer novos produtos ou serviços a quem já é cliente. "
                                   "Risco moderado: exige criar e testar o novo produto."),
    "Desenvolvimento de mercado": ("Produto atual, mercado novo", "Levar o produto a outros bairros, cidades, públicos ou canais. "
                                   "Risco moderado: exige conhecer o novo público."),
    "Diversificação": ("Produto novo, mercado novo", "Entrar com produto novo em mercado novo. É o caminho de maior risco, "
                       "porque nada do que a empresa já sabe se aplica diretamente."),
}
GENERICAS = {
    "Liderança em custo": "Ser o mais barato de produzir e vender, com preço competitivo.",
    "Diferenciação": "Oferecer algo que o cliente valoriza e pelo qual aceita pagar mais.",
    "Foco em custo": "Atender um nicho específico com o menor custo.",
    "Foco em diferenciação": "Atender um nicho específico com uma oferta feita sob medida para ele.",
}
FREQUENCIAS = ["Diária", "Semanal", "Mensal", "Trimestral", "Semestral", "Anual"]
KPIS_SUGERIDOS = [
    ("Clientes novos por mês", "Mensal", "Cadastro de clientes ou vendas"),
    ("Custo por cliente novo (CAC)", "Mensal", "Gasto de marketing ÷ clientes novos"),
    ("Gasto médio por compra (ticket médio)", "Semanal", "Sistema de caixa"),
    ("Taxa de recompra em 30 dias", "Mensal", "Cadastro de clientes ou cartão fidelidade"),
    ("Conversão de contatos em vendas", "Semanal", "WhatsApp, redes sociais ou CRM"),
    ("Nota de satisfação (NPS)", "Trimestral", "Pesquisa com clientes"),
    ("Receita por canal de venda", "Mensal", "Relatórios dos canais e do caixa"),
]
ETAPAS_PLANO = [
    ("Situação", "Onde estamos? Mercado, clientes, concorrentes e fatores internos e externos."),
    ("Objetivos", "Aonde queremos chegar? Objetivos de marketing mensuráveis e com prazo."),
    ("Estratégia", "Como chegar lá? Público-alvo, posicionamento e direção de crescimento."),
    ("Táticas", "Com que ferramentas? O mix de marketing (4Ps ou 7Ps)."),
    ("Ações", "Quem faz o quê e quando? Calendário, responsáveis e custos."),
    ("Controle", "Chegamos lá? Indicadores, metas e frequência de acompanhamento."),
]


def _tem_txt(x) -> bool:
    return bool(str(x or "").strip())


def objetivos_marketing(d: dict) -> pd.DataFrame:
    ob = d.get("objetivos")
    if ob is None or ob.empty:
        return pd.DataFrame(columns=["objetivo", "indicador", "meta", "unidade", "prazo_meses"])
    return ob[ob["perspectiva"] == "Clientes"]


def etapas_plano_mkt(proj: dict, d: dict) -> list[tuple[str, int, str]]:
    """Situação de cada etapa: (etapa, 0 = vazia / 1 = parcial / 2 = completa, o que falta)."""
    def n(t):
        x = d.get(t)
        return 0 if x is None else len(x)
    s = [proj.get("tam_clientes") or 0, n("concorrentes"), n("swot"), _tem_txt(proj.get("pm_situacao"))]
    sit = sum(bool(x) for x in s)
    obm = objetivos_marketing(d)
    est = [_tem_txt(proj.get("posic_segmento")), _tem_txt(proj.get("pm_ansoff")), _tem_txt(proj.get("pm_generica"))]
    tat = [_tem_txt(proj.get("mix_nucleo")), _tem_txt(proj.get("mix_estrategia_preco")), n("canais_venda") > 0,
           _tem_txt(proj.get("mix_objetivo_com"))]
    acoes = n("calendario_mkt")
    kp = n("kpis_mkt")
    return [
        ("Situação", 2 if sit >= 4 else (1 if sit else 0), "mercado, concorrentes, SWOT e a síntese do diagnóstico"),
        ("Objetivos", 2 if len(obm) >= 2 else (1 if len(obm) else 0), "ao menos dois objetivos da perspectiva Clientes"),
        ("Estratégia", 2 if all(est) else (1 if any(est) else 0), "público-alvo, direção de crescimento e estratégia competitiva"),
        ("Táticas", 2 if all(tat) else (1 if any(tat) else 0), "produto, estratégia de preço, canais de venda e objetivo da comunicação"),
        ("Ações", 2 if acoes >= 4 else (1 if acoes else 0), "ao menos quatro ações no calendário de marketing"),
        ("Controle", 2 if kp >= 3 else (1 if kp else 0), "ao menos três indicadores com meta e frequência"),
    ]


def regras_mix(proj: dict, d: dict) -> dict:
    et = etapas_plano_mkt(proj, d)
    completas = sum(1 for _, s, _ in et if s == 2)
    saz = d.get("sazonalidade")
    tem_saz = saz is not None and len(saz) >= 12 and float(saz["indice"].max() - saz["indice"].min()) > 0
    return {"plano_mkt": (completas >= 5, any(s for _, s, _ in et)),
            "sazonalidade": (tem_saz, saz is not None and len(saz) > 0)}


def frase_posicionamento(proj: dict) -> str:
    p = {k: str(proj.get(k) or "").strip() for k in ("posic_segmento", "posic_necessidade", "posic_marca", "posic_categoria",
                                                     "posic_beneficio", "posic_concorrente", "posic_diferencial")}
    if not (p["posic_segmento"] and p["posic_marca"]):
        return ""
    return (f"Para {p['posic_segmento']}, que {p['posic_necessidade'] or '...'}, {p['posic_marca']} é {p['posic_categoria'] or '...'} "
            f"que {p['posic_beneficio'] or '...'}. Diferentemente de {p['posic_concorrente'] or '...'}, "
            f"{p['posic_diferencial'] or '...'}.")


def resumo_mix_linhas(proj: dict, d: dict, res: dict) -> list[tuple[str, str]]:
    """Uma linha por P, montada a partir das escolhas estruturadas (para o plano e os relatórios)."""
    cv = d.get("canais_venda")
    canais = ", ".join(f"{c} ({_n(p)}%)" for c, p in zip(cv["canal"], cv["pct_vendas"])) if cv is not None and not cv.empty else ""
    mm = mc_media(res)
    linhas = [
        ("Produto", "; ".join(x for x in [proj.get("mix_nucleo"), proj.get("mix_ampliado")] if _tem_txt(x))
         + (f". Fase do ciclo de vida: {proj['mix_ciclo'].lower()}." if _tem_txt(proj.get("mix_ciclo")) else "")),
        ("Preço", ", ".join(x for x in [f"estratégia de {proj['mix_estrategia_preco'].lower()}" if _tem_txt(proj.get("mix_estrategia_preco")) else "",
                                         f"posição {proj['mix_posicao_preco'].lower()}" if _tem_txt(proj.get("mix_posicao_preco")) else "",
                                         f"margem de contribuição média de {_p(mm)}" if mm == mm else ""] if x)),
        ("Praça", ", ".join(x for x in [f"distribuição {proj['mix_distribuicao'].lower()}" if _tem_txt(proj.get("mix_distribuicao")) else "",
                                         f"canais: {canais}" if canais else ""] if x)),
        ("Promoção", proj.get("mix_objetivo_com") or ""),
    ]
    for rot, campo in (("Pessoas", "mix_pessoas"), ("Processos", "mix_processos"), ("Evidência física", "mix_evidencia")):
        if _tem_txt(proj.get(campo)):
            linhas.append((rot, proj[campo]))
    texto = {"Produto": "mkt_produto", "Preço": "mkt_preco", "Praça": "mkt_praca", "Promoção": "mkt_promocao"}
    def cap(x):
        x = str(x).strip()
        return x[:1].upper() + x[1:]
    return [(r, cap(t if t else str(proj.get(texto.get(r, ""), "") or ""))) for r, t in linhas
            if t or _tem_txt(proj.get(texto.get(r, "")))]


def conteudo_plano_mkt(proj: dict, d: dict, res: dict) -> list[dict]:
    """Seções do plano de marketing para a tela e os relatórios: titulo, texto, tabela (DataFrame ou None), refs."""
    out = []
    # 1 Situação
    t = []
    if proj.get("tam_clientes"):
        t.append(f"O mercado atendível foi estimado em {_n(float(proj.get('sam_clientes') or 0))} clientes, com meta de "
                 f"participação de {_n(float(proj.get('som_pct') or 0), 1)}%.")
    nc = len(d.get("concorrentes", []))
    if nc:
        comp = comparar_concorrentes(res.get("cat"), d.get("precos_concorrentes"))
        idx = indice_preco_medio(res.get("cat"), comp) if not comp.empty else float("nan")
        t.append(f"Foram analisados {nc} concorrentes" + (f", cujos preços ficam, em média, {_p(abs(idx - 1), 0)} "
                 f"{'abaixo' if idx >= 1 else 'acima'} dos do projeto." if idx == idx else "."))
    if _tem_txt(proj.get("pm_situacao")):
        t.append(str(proj["pm_situacao"]))
    out.append({"titulo": "Análise da situação", "texto": " ".join(t) or "Diagnóstico a elaborar.", "tabela": None, "refs": ["smith"]})
    # 2 Objetivos
    obm = objetivos_marketing(d)
    tab = None
    if not obm.empty:
        tab = pd.DataFrame({"Objetivo": obm["objetivo"], "Indicador": obm["indicador"],
                            "Meta": [f"{_n(m, 0)} {u}".strip() for m, u in zip(obm["meta"], obm["unidade"])],
                            "Prazo (mês)": obm["prazo_meses"].map(lambda v: _n(v))})
    out.append({"titulo": "Objetivos de marketing", "texto": "Os objetivos seguem os critérios SMART: específicos, mensuráveis, "
                "atingíveis, relevantes e com prazo." if tab is not None else "Objetivos a definir.", "tabela": tab, "refs": []})
    # 3 Estratégia
    t = []
    fp = frase_posicionamento(proj)
    if fp:
        t.append("Declaração de posicionamento: " + fp)
    if _tem_txt(proj.get("pm_ansoff")):
        t.append(f"Direção de crescimento (matriz de Ansoff): {proj['pm_ansoff'].lower()}, isto é, "
                 f"{ANSOFF.get(proj['pm_ansoff'], ('', ''))[0].lower()}.")
    if _tem_txt(proj.get("pm_generica")):
        t.append(f"Estratégia competitiva: {proj['pm_generica'].lower()}. {GENERICAS.get(proj['pm_generica'], '')}")
    out.append({"titulo": "Estratégia de marketing", "texto": " ".join(t) or "Estratégia a definir.", "tabela": None,
                "refs": ["ansoff"] if _tem_txt(proj.get("pm_ansoff")) else []})
    # 4 Táticas
    lm = resumo_mix_linhas(proj, d, res)
    refs = ["kotler"] + (["booms"] if any(r in ("Pessoas", "Processos", "Evidência física") for r, _ in lm) else [])
    out.append({"titulo": "Táticas: mix de marketing", "texto": "O mix de marketing traduz a estratégia em decisões sobre cada P.",
                "tabela": pd.DataFrame(lm, columns=["Elemento", "Decisão"]) if lm else None, "refs": refs})
    # 5 Ações
    cal = d.get("calendario_mkt")
    tab = None
    if cal is not None and not cal.empty:
        c = cal.sort_values("mes")
        tab = pd.DataFrame({"Mês": c["mes"].map(lambda v: _n(v)), "Ação": c["acao"], "Canal": c["canal"],
                            "Custo (R$)": c["custo"].map(lambda v: _n(v, 2))})
    out.append({"titulo": "Plano de ação", "texto": (f"O calendário prevê {len(cal)} ações, com custo total de R$ "
                f"{_n(float(cal['custo'].sum()), 2)}." if tab is not None else "Ações a programar."), "tabela": tab, "refs": []})
    # 6 Controle
    kp = d.get("kpis_mkt")
    tab = None
    if kp is not None and not kp.empty:
        tab = pd.DataFrame({"Indicador": kp["indicador"], "Meta": kp["meta"], "Frequência": kp["frequencia"],
                            "Fonte do dado": kp["fonte"], "Responsável": kp["responsavel"]})
    out.append({"titulo": "Controle", "texto": "Os indicadores abaixo serão acompanhados e comparados com as metas, e os desvios "
                "levam à revisão das táticas, fechando o ciclo de planejamento." if tab is not None else "Indicadores a definir.",
                "tabela": tab, "refs": ["mcdonald"]})
    return out


REFS_MKT = {
    "smith": "SMITH, P. R. **The SOSTAC guide**: to writing the perfect plan. [S. l.]: PR Smith, 2011.",
    "mcdonald": "McDONALD, Malcolm. **Marketing plans**: how to prepare them, how to use them. 6. ed. Oxford: Butterworth-Heinemann, 2007.",
    "kotler": "KOTLER, Philip; KELLER, Kevin Lane. **Administração de marketing**. 14. ed. São Paulo: Pearson Education do Brasil, 2012.",
    "booms": "BOOMS, Bernard H.; BITNER, Mary Jo. Marketing strategies and organization structures for service firms. In: DONNELLY, "
             "James H.; GEORGE, William R. (ed.). **Marketing of services**. Chicago: American Marketing Association, 1981. p. 47-51.",
    "ansoff": "ANSOFF, H. Igor. Strategies for diversification. **Harvard Business Review**, v. 35, n. 5, p. 113-124, 1957.",
    "lascasas": "LAS CASAS, Alexandre Luzzi. **Plano de marketing para micro e pequena empresa**. 6. ed. São Paulo: Atlas, 2011.",
}
CITA_MKT = {"smith": "Smith, 2011", "mcdonald": "McDonald, 2007", "kotler": "Kotler; Keller, 2012",
            "booms": "Booms; Bitner, 1981", "ansoff": "Ansoff, 1957", "lascasas": "Las Casas, 2011"}

# ##############################################################################
# PARTE: pnbox_leituras.py
# ##############################################################################




def _q(n, singular: str, plural: str) -> str:
    """Quantidade com a palavra no singular ou no plural: _q(1, "item", "itens") -> "1 item"."""
    n = float(n)
    return f"{_n(n)} {singular if abs(n) == 1 else plural}"


PLURAL_SWOT = {"Força": "forças", "Fraqueza": "fraquezas", "Oportunidade": "oportunidades", "Ameaça": "ameaças"}


def _brl(x: float) -> str:
    return "R$ " + _n(float(x), 2)


def _validas(df: pd.DataFrame, col: str) -> pd.DataFrame:
    return df[df[col].astype(str).str.strip() != ""] if col in df else df


def _top(df, col_valor, col_nome):
    if df.empty or df[col_valor].sum() <= 0:
        return None, 0.0
    i = df[col_valor].astype(float).idxmax()
    return str(df.loc[i, col_nome]), float(df.loc[i, col_valor]) / float(df[col_valor].sum())


# ------------------------------------------------------------------------------
# FINANÇAS
# ------------------------------------------------------------------------------
def _l_catalogo(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    res = R()
    cat = res["cat"]
    out = [f"{_q(len(df), 'item soma', 'itens somam')} receita de {_brl(res['receita_m1'])} no primeiro mês."]
    nome, share = _top(cat, "receita_mensal", "descricao")
    if nome:
        out.append(f"O item que mais fatura é {nome}, com {_p(share, 0)} da receita."
                   + (" O plano depende muito dele." if share >= 0.5 and len(cat) > 2 else ""))
    neg = cat.loc[cat["margem_unit"] <= 0, "descricao"].tolist()
    if neg:
        out.append("Vendidos sem margem (preço menor ou igual ao custo): " + ", ".join(neg) + ".")
    mm = mc_media(res)
    if mm == mm:
        out.append(f"Depois de impostos e custos variáveis, sobra em média {_p(mm)} de cada venda para pagar os custos fixos.")
    return out


def _l_custos_fixos(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["valor_mensal"].sum())
    out = [f"Os custos fixos (sem pessoal) somam {_brl(tot)} por mês, {_brl(tot * 12)} por ano."]
    nome, share = _top(df, "valor_mensal", "descricao")
    if nome:
        out.append(f"O maior é {nome}, com {_p(share, 0)} do total.")
    res = R()
    if res["receita_m1"] > 0:
        r = (tot + res["pessoal_m"]) / res["receita_m1"]
        out.append(f"Somados à equipe, os custos fixos equivalem a {_p(r, 0)} da receita do primeiro mês"
                   + (": o negócio começa sem cobrir a estrutura e precisa crescer para empatar." if r > 1 else "."))
    return out


def _l_custos_variaveis(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["percentual"].sum())
    out = [f"Os custos variáveis sobre a venda somam {_n(tot, 2)}%: de cada R$ 100 vendidos, R$ {_n(tot, 2)} saem para eles,"
           " além dos impostos."]
    nome, share = _top(df, "percentual", "descricao")
    if nome:
        out.append(f"O mais pesado é {nome}.")
    return out


def _l_pessoal(df, proj, d, R):
    df = _validas(df, "cargo")
    if df.empty:
        return []
    res = R()
    n = float(df["num_empregados"].sum())
    out = [f"{_q(n, 'pessoa', 'pessoas')} na equipe, com custo de {_brl(res['pessoal_m'])} por mês no primeiro mês, contando encargos e benefícios."]
    fut = df[df["mes_inicio"] > 1]
    if not fut.empty:
        out.append(f"{_q(len(fut), 'contratação prevista', 'contratações previstas')} para depois da abertura, o que aumenta o custo ao longo do plano.")
    if res["receita_m1"] > 0:
        out.append(f"A equipe consome {_p(res['pessoal_m'] / res['receita_m1'], 0)} da receita do primeiro mês.")
    return out


def _l_inv_fixos(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float((df["quantidade"] * df["valor_unitario"]).sum())
    res = R()
    out = [f"Os investimentos fixos somam {_brl(tot)}, {_p(tot / res['investimento_total'], 0) if res['investimento_total'] else '-'} "
           "do investimento total."]
    tmp = df.assign(total=df["quantidade"] * df["valor_unitario"])
    nome, share = _top(tmp, "total", "descricao")
    if nome:
        out.append(f"O item mais caro é {nome} ({_p(share, 0)}).")
    out.append(f"A depreciação, que entra como custo sem sair do caixa, é de {_brl(res['dep_anual'])} por ano.")
    return out


def _l_inv_pre(df, proj, d, R):
    df = _validas(df, "descricao")
    if df.empty:
        return []
    tot = float(df["valor"].sum())
    return [f"Os gastos antes de abrir somam {_brl(tot)}. São gastos que não voltam se o negócio não sair do papel, "
            "por isso convém conferir orçamentos reais de cada um."]


def _l_emprestimos(df, proj, d, R):
    df = _validas(df, "descricao")
    df = df[df["valor"] > 0]
    if df.empty:
        return []
    tot = float(df["valor"].sum())
    taxa = float((df["valor"] * df["taxa_anual_pct"]).sum() / tot)
    res = R()
    out = [f"{_q(len(df), 'empréstimo soma', 'empréstimos somam')} {_brl(tot)}, com taxa média de {_n(taxa, 1)}% ao ano."]
    if res["fontes"] > 0:
        out.append(f"O dinheiro de terceiros cobre {_p(tot / res['fontes'], 0)} das fontes de recursos.")
    if taxa / 100 > res["tma"] and res["tma"] > 0:
        out.append("A taxa do empréstimo é maior que a TMA do projeto: o crédito custa mais do que o retorno mínimo exigido.")
    return out


def _l_socios(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    tot = float(df["capital_integralizado"].sum())
    out = [f"{_q(len(df), 'sócio investe', 'sócios investem')} {_brl(tot)} de capital próprio."]
    if tot > 0 and len(df) > 1:
        nome, share = _top(df, "capital_integralizado", "nome")
        out.append(f"{nome} entra com {_p(share, 0)} do capital" + (", o que garante o controle da sociedade." if share > 0.5 else "."))
    return out


def _l_prazos(df, proj, d, R):
    res = R()
    out = [f"O cliente paga, em média, em {_n(res['pmr'])} dias e a empresa paga fornecedores em {_n(res['pmp'])} dias; "
           f"o estoque gira em {_n(res['pme'])} dias."]
    cf = res["ciclo_fin"]
    out.append((f"O ciclo financeiro é de {_n(cf)} dias" if cf >= 0 else f"O ciclo financeiro é negativo ({_n(cf)} dias)") + (": a empresa financia a operação nesse período, e é isso que pede capital de giro."
                                                         if cf > 0 else ": os fornecedores financiam a operação."))
    return out


def _l_materiais(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    est = float((df["qtd_estoque_inicial"] * df["valor_unitario"]).sum())
    return [f"{_q(len(df), 'material cadastrado', 'materiais cadastrados')}; o estoque inicial deles custa {_brl(est)}."]


def _l_assinaturas(df, proj, d, R):
    df = _validas(df, "plano")
    if df.empty:
        return []
    out = []
    for _, r in df.iterrows():
        ch = float(r["churn_mensal_pct"] or 0)
        if ch > 0:
            out.append(f"No plano {r['plano']}, com cancelamento de {_n(ch, 1)}% ao mês, o assinante fica em média {_n(100 / ch, 0)} meses.")
    return out[:3]


def _l_realizado(df, proj, d, R):
    df = df[df["mes"] > 0]
    if df.empty:
        return []
    res = R()
    plan = res["meses"]["Receita"].tolist()
    reais = [(int(m), float(r)) for m, r in zip(df["mes"], df["receita_real"]) if 1 <= m <= 12]
    if not reais:
        return []
    p = sum(plan[m - 1] for m, _ in reais)
    r = sum(v for _, v in reais)
    return [f"Em {_q(len(reais), 'mês lançado', 'meses lançados')}, a receita real foi {_brl(r)}, contra {_brl(p)} planejados "
            f"({_p(r / p - 1, 0) if p else '-'} de diferença).",
            "Diferença acima de 10% para baixo pede revisão do volume ou do preço no plano." if p and r / p < 0.9 else
            "O realizado está próximo ou acima do planejado."]


# ------------------------------------------------------------------------------
# MERCADO, ESTRATÉGIA E MARKETING
# ------------------------------------------------------------------------------
def _l_segmentos(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    notas = ["nota_tamanho", "nota_lucro", "nota_longevidade", "nota_concorrencia", "nota_aderencia"]
    m = df.assign(media=df[notas].mean(axis=1)).sort_values("media", ascending=False)
    out = [f"O segmento mais atraente pelas notas é {m['nome'].iloc[0]} (média {_n(m['media'].iloc[0], 1)})."]
    if len(m) > 1 and m["media"].iloc[0] - m["media"].iloc[1] < 0.5:
        out.append(f"A diferença para {m['nome'].iloc[1]} é pequena: vale olhar qual dos dois o grupo consegue atender melhor.")
    return out


def _l_swot(df, proj, d, R):
    df = _validas(df, "fator")
    if df.empty:
        return []
    c = df["natureza"].value_counts()
    out = ["Fatores: " + ", ".join(_q(v, k.lower(), PLURAL_SWOT.get(k, k.lower())) for k, v in c.items()) + "."]
    if c.get("Ameaça", 0) > c.get("Oportunidade", 0):
        out.append("Há mais ameaças que oportunidades no ambiente: as estratégias defensivas da SWOT cruzada ganham peso.")
    if c.get("Fraqueza", 0) > c.get("Força", 0):
        out.append("As fraquezas superam as forças: o plano precisa mostrar como cada uma será tratada.")
    return out


def _l_porter(df, proj, d, R):
    df = _validas(df, "forca")
    if df.empty:
        return []
    m = float(df["intensidade"].mean())
    i = df["intensidade"].idxmax()
    return [f"Intensidade média das forças: {_n(m, 1)}. A mais forte é {df.loc[i, 'forca'].lower()}, "
            "que é onde a estratégia precisa responder primeiro."]


def _l_riscos(df, proj, d, R):
    df = _validas(df, "risco")
    if df.empty:
        return []
    sev = df["probabilidade"] * df["impacto"]
    alto = df[sev >= 15]
    i = sev.idxmax()
    out = [f"O risco mais grave é {df.loc[i, 'risco']} (probabilidade × impacto = {_n(sev[i])})."]
    if not alto.empty:
        out.append(f"{_q(len(alto), 'risco está', 'riscos estão')} na zona alta (15 ou mais) e pedem plano de contingência antes da abertura.")
    sem = df[df["mitigacao"].astype(str).str.strip() == ""]
    if not sem.empty:
        out.append(f"{_q(len(sem), 'risco', 'riscos')} ainda sem mitigação.")
    return out


def _l_precos_vw(df, proj, d, R):
    df = _validas(df, "respondente")
    n = len(df)
    if not n:
        return []
    return [f"{_q(n, 'resposta', 'respostas')}. " + ("Com menos de 5 respostas coerentes a faixa de preço não é calculada; o recomendado é passar de 30."
                                    if n < 5 else "A faixa aceitável e o preço ótimo aparecem no gráfico; quanto mais respostas, mais confiável."
                                    if n < 30 else "Amostra razoável para uma primeira faixa de preço.")]


def _l_entrevistas(df, proj, d, R):
    df = _validas(df, "entrevistado")
    if df.empty:
        return []
    sim = float((df["pagaria"] == "Sim").mean())
    out = [f"{_q(len(df), 'entrevista', 'entrevistas')}; {_p(sim, 0)} disseram que pagariam."]
    if len(df) < 10:
        out.append("Com menos de 10 conversas é cedo para conclusões: as respostas ainda podem mudar muito.")
    if sim > 0.8:
        out.append("Intenção muito alta costuma ser otimista: quem diz \"pagaria\" nem sempre compra.")
    return out


def _l_canais(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    m = df.assign(media=df.iloc[:, 1:6].mean(axis=1)).sort_values("media", ascending=False)
    return [f"Os canais com maior nota média são {', '.join(m['canal'].head(3))}. Concentrar a verba neles costuma render mais "
            "do que espalhar em muitos canais."]


def _l_funil(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    topo = float(df.loc[df["etapa"] == "Topo do Funil", "alcancados"].sum())
    fundo = float(df.loc[df["etapa"] == "Fundo do Funil", "convertidos"].sum())
    if topo <= 0:
        return []
    return [f"De {_n(topo)} pessoas alcançadas no topo, {_n(fundo)} compram: conversão total de {_p(fundo / topo, 1)}."]


def _l_orcamento_canais(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    g, c = float(df["gasto_mensal"].sum()), float(df["clientes_mes"].sum())
    out = [f"Orçamento de {_brl(g)} por mês para trazer {_n(c)} clientes novos: custo médio de {_brl(g / c) if c else '-'} por cliente."]
    x = df[df["clientes_mes"] > 0].assign(cac=lambda t: t["gasto_mensal"] / t["clientes_mes"]).sort_values("cac")
    if len(x) > 1:
        out.append(f"O canal mais barato por cliente é {x['canal'].iloc[0]}; o mais caro, {x['canal'].iloc[-1]}.")
    return out


def _l_fornecedores(df, proj, d, R):
    df = _validas(df, "nome")
    if df.empty:
        return []
    crit = df[df["critico"].astype(bool)]
    out = [f"{_q(len(df), 'fornecedor', 'fornecedores')}, prazo médio de pagamento de {_n(float(df['prazo_pagamento_dias'].mean()))} dias."]
    if not crit.empty:
        out.append(f"{_q(len(crit), 'fornecedor está marcado como crítico', 'fornecedores estão marcados como críticos')}: vale ter um segundo fornecedor para cada um.")
    return out


def _l_cronograma(df, proj, d, R):
    df = _validas(df, "atividade")
    if df.empty:
        return []
    fim = float((df["mes_inicio"] + df["duracao_meses"] - 1).max())
    out = [f"{_q(len(df), 'atividade', 'atividades')}, terminando por volta do mês {_n(fim)}, com custo previsto de {_brl(df['custo'].sum())}."]
    sem = df[df["responsavel"].astype(str).str.strip() == ""]
    if not sem.empty:
        out.append(f"{_q(len(sem), 'atividade', 'atividades')} sem responsável.")
    return out


def _l_calendario(df, proj, d, R):
    df = _validas(df, "acao")
    if df.empty:
        return []
    meses = df["mes"].nunique()
    return [f"{_q(len(df), 'ação', 'ações')} em {_q(meses, 'mês', 'meses diferentes')}, custo total de {_brl(df['custo'].sum())}."
            + (" As ações estão concentradas em poucos meses; depois do lançamento a comunicação some." if meses <= 2 and len(df) >= 3 else "")]


def _l_bcg(df, proj, d, R):
    df = _validas(df, "produto")
    if df.empty:
        return []
    lim = float(proj.get("bcg_limite_crescimento_pct") or 10)
    q = {"Estrela": 0, "Vaca leiteira": 0, "Interrogação": 0, "Abacaxi": 0}
    for g, p in zip(df["crescimento_mercado_pct"], df["participacao_relativa"]):
        q["Estrela" if g >= lim and p >= 1 else "Interrogação" if g >= lim else "Vaca leiteira" if p >= 1 else "Abacaxi"] += 1
    out = ["; ".join(f"{k}: {v}" for k, v in q.items() if v) + "."]
    if q["Vaca leiteira"] == 0:
        out.append("Sem vaca leiteira, não há item que gere caixa folgado para financiar os outros.")
    return out


def _l_errc(df, proj, d, R):
    df = _validas(df, "atributo")
    if df.empty:
        return []
    c = df["acao"].value_counts()
    return [", ".join(f"{k}: {v}" for k, v in c.items()) + ". A nova curva de valor só se diferencia se houver algo em criar ou elevar."]


def _l_experimentos(df, proj, d, R):
    df = _validas(df, "hipotese")
    if df.empty:
        return []
    c = df["status"].value_counts()
    return [", ".join(f"{k or 'sem status'}: {v}" for k, v in c.items()) + "."]


def _l_formalizacao(df, proj, d, R):
    df = _validas(df, "item")
    if df.empty:
        return []
    c = df["status"].value_counts()
    return [", ".join(f"{k or 'sem status'}: {v}" for k, v in c.items()) + "."]


def _l_aarrr(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    sem = df[(df["metrica"].astype(str).str.strip() == "") | (df["meta"].astype(str).str.strip() == "")]
    return [f"{len(df) - len(sem)} de {len(df)} etapas têm métrica e meta." + (" Etapa sem meta não pode ser acompanhada." if len(sem) else "")]


def _l_objetivos(df, proj, d, R):
    df = _validas(df, "objetivo")
    if df.empty:
        return []
    c = df["perspectiva"].replace("", "sem perspectiva").value_counts()
    out = ["Objetivos por perspectiva: " + ", ".join(f"{k} {v}" for k, v in c.items()) + "."]
    falta = [p for p in ("Financeira", "Clientes", "Processos internos", "Aprendizado e crescimento") if p not in c]
    if falta:
        out.append("Sem objetivo em: " + ", ".join(falta).lower() + ". O mapa estratégico fica incompleto.")
    return out


def _l_precos_concorrentes(df, proj, d, R):
    cat = R()["cat"]
    comp = comparar_concorrentes(cat, df)
    if comp.empty:
        return []
    idx = indice_preco_medio(cat, comp)
    return [f"Considerando {_q(len(comp), 'produto comparado', 'produtos comparados')}, seus preços estão {_p(abs(idx - 1), 0)} "
            f"{'acima' if idx >= 1 else 'abaixo'} da média dos concorrentes."]


def _l_meios_pagamento(df, proj, d, R):
    df = _validas(df, "forma")
    if df.empty:
        return []
    r = resumo_pagamento(df)
    return [f"As taxas de pagamento consomem {_n(r['taxa'], 2)}% da receita e o dinheiro entra, em média, {_n(r['prazo'])} dias "
            "depois da venda. Quanto maior esse prazo, mais capital de giro o negócio precisa."]


def _l_canais_venda(df, proj, d, R):
    df = _validas(df, "canal")
    if df.empty:
        return []
    r = resumo_canais(df)
    return [f"{r['maior']} responde por {_n(r['maior_pct'])}% das vendas; a comissão média dos canais é de {_n(r['comissao'], 2)}% da receita."]


def _l_blueprint(df, proj, d, R):
    df = _validas(df, "etapa")
    if df.empty:
        return []
    b = resumo_blueprint(df)
    return [f"O serviço tem {b['etapas']} etapas e leva {_n(b['tempo'])} minutos, dos quais {_n(b['visivel'])} na frente do cliente.",
            f"{_q(len(b['falhas']), 'ponto de falha identificado', 'pontos de falha identificados')}." if b["falhas"] else "Nenhum ponto de falha anotado; vale perguntar onde o cliente costuma reclamar."]


def _l_kpis(df, proj, d, R):
    df = _validas(df, "indicador")
    if df.empty:
        return []
    sem = int((df["meta"].astype(str).str.strip() == "").sum())
    return [f"{_q(len(df), 'indicador', 'indicadores')} de controle" + (f", {sem} sem meta." if sem else ", todos com meta.")]


def _l_swot_estrategias(df, proj, d, R):
    df = _validas(df, "estrategia")
    if df.empty:
        return []
    q = df["quadrante"].str.split(" ").str[0].value_counts()
    return ["Estratégias por tipo: " + ", ".join(f"{k} {v}" for k, v in q.items()) + "."]


def _l_capacidade(df, proj, d, R):
    df = _validas(df, "recurso")
    if df.empty:
        return []
    cat = R()["cat"]
    out = []
    for _, r in df.iterrows():
        q = cat.loc[cat["descricao"] == r["produto_ou_grupo"], "qtd_mensal"].sum() if not cat.empty else 0
        if q and r["capacidade_mensal"]:
            u = q / r["capacidade_mensal"]
            out.append(f"{r['recurso']} trabalha a {_p(u, 0)} da capacidade no primeiro mês"
                       + (": acima de 85% sobra pouca folga para crescer." if u > 0.85 else "."))
    return out[:3] or [f"{_q(len(df), 'recurso de capacidade cadastrado', 'recursos de capacidade cadastrados')}."]


LEITURAS = {
    "catalogo": _l_catalogo, "custos_fixos": _l_custos_fixos, "custos_variaveis": _l_custos_variaveis, "pessoal": _l_pessoal,
    "inv_fixos": _l_inv_fixos, "inv_pre": _l_inv_pre, "emprestimos": _l_emprestimos, "socios": _l_socios,
    "capital_giro_prazos": _l_prazos, "materiais": _l_materiais, "assinaturas": _l_assinaturas, "realizado": _l_realizado,
    "segmentos": _l_segmentos, "swot": _l_swot, "porter": _l_porter, "riscos": _l_riscos, "precos_vw": _l_precos_vw,
    "entrevistas": _l_entrevistas, "canais": _l_canais, "funil_vendas": _l_funil, "orcamento_canais": _l_orcamento_canais,
    "fornecedores": _l_fornecedores, "cronograma": _l_cronograma, "calendario_mkt": _l_calendario, "bcg": _l_bcg,
    "errc": _l_errc, "experimentos": _l_experimentos, "formalizacao": _l_formalizacao, "aarrr": _l_aarrr,
    "objetivos": _l_objetivos, "precos_concorrentes": _l_precos_concorrentes, "meios_pagamento": _l_meios_pagamento,
    "canais_venda": _l_canais_venda, "blueprint": _l_blueprint, "kpis_mkt": _l_kpis, "swot_estrategias": _l_swot_estrategias,
    "capacidade": _l_capacidade,
}


# ------------------------------------------------------------------------------
# TABELAS DE RESULTADO (DRE, fluxo de caixa)
# ------------------------------------------------------------------------------
def leitura_dre(res: dict) -> list[str]:
    a = res["anos"]
    N = res["N"]
    a1, an = a[1], a[N]
    out = [f"No Ano 1 a receita é de {_brl(a1['rec_bruta'])} e o resultado líquido de {_brl(a1['resultado'])} "
           f"({_p(a1['resultado'] / a1['rec_bruta'], 1) if a1['rec_bruta'] else '-'} da receita)."]
    if N > 1:
        out.append(f"No Ano {N} o resultado chega a {_brl(an['resultado'])}"
                   + (", saindo do prejuízo." if a1["resultado"] < 0 <= an["resultado"] else "."))
    prim = next((k for k in res["ANOS"] if a[k]["resultado"] > 0), None)
    if prim is None:
        out.append("O resultado não fica positivo em nenhum ano do horizonte: o plano, como está, não se paga.")
    m = res["meses"]
    pos = m.index[m["Resultado operacional"] > 0]
    if len(pos):
        out.append("O resultado operacional já é positivo no primeiro mês." if pos[0] == m.index[0] else
                   f"Mês a mês, o resultado operacional fica positivo a partir do {m.loc[pos[0], 'Mês'].lower()}.")
    return out


def leitura_fluxo(res: dict) -> list[str]:
    fm = res["fluxo_mensal"]
    if fm.empty:
        return []
    i = fm["Saldo final"].idxmin()
    mn, mes = float(fm.loc[i, "Saldo final"]), fm.loc[i, "Mês"]
    out = [f"O caixa chega ao ponto mais baixo no {str(mes).lower()}, com {_brl(mn)}."]
    out.append("Como fica negativo, falta dinheiro nesse período: é preciso mais capital de giro, prazo maior de fornecedores "
               "ou recebimento mais rápido." if mn < 0 else "Como não fica negativo, o capital de giro previsto cobre o período mais apertado.")
    return out

# ##############################################################################
# PARTE: pnbox_v9.py
# ##############################################################################




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

# ##############################################################################
# PARTE: pnbox_pdf.py
# ##############################################################################







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

# ##############################################################################
# PARTE: pnbox_extras.py
# ##############################################################################








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
    fora = {"realizado", "adocao", "plano_mkt", "sazonalidade"} | OPCIONAIS_V8
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

# ##############################################################################
# PARTE: pnbox_abnt.py
# ##############################################################################









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

# ##############################################################################
# PARTE: pnbox_ui_base.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: BASE (configuração, projetos, navegação, salvamento)
# ==============================================================================





st.set_page_config(page_title="ViabilizaLab - Plano de Negócio", page_icon="📊", layout="wide",
                   initial_sidebar_state="auto")      # aberto no computador, recolhido no celular
ss = st.session_state
VERSAO_PROGRAMA = "8.6"

# ------------------------------------------------------------------------------
# ENCERRAMENTO AUTOMÁTICO: quando todas as abas do navegador forem fechadas,
# o programa termina sozinho (só quando aberto pelo próprio arquivo / .bat / .exe).
# ------------------------------------------------------------------------------
TEMPO_ESPERA_FECHAR = 15


@st.cache_resource
def _vigia_de_encerramento():
    if os.environ.get("PNBOX_AUTOFECHAR") != "1":
        return False

    def vigiar():
        from streamlit.runtime import Runtime
        ja_conectou, desde = False, None
        while True:
            time.sleep(2)
            try:
                abertas = Runtime.instance()._session_mgr.num_active_sessions()
            except Exception:
                continue
            if abertas > 0:
                ja_conectou, desde = True, None
            elif ja_conectou:
                desde = desde or time.time()
                if time.time() - desde > TEMPO_ESPERA_FECHAR:
                    print("Nenhuma aba aberta. Encerrando o ViabilizaLab...")
                    os._exit(0)

    threading.Thread(target=vigiar, daemon=True, name="pnbox-vigia").start()
    return True


_vigia_de_encerramento()

# ------------------------------------------------------------------------------
# ONDE O PROGRAMA ESTÁ RODANDO E QUAL VERSÃO ABRIR
# ------------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
NAVEGADOR = sys.platform == "emscripten"                      # GitHub Pages (stlite)
MODO_WEB = NAVEGADOR or Path("/mount/src").exists() or os.environ.get("PNBOX_WEB") == "1"


def _versao_inicial() -> str:
    """'modelo' abre com o FitBox; 'branco', com um projeto vazio; 'assinatura', com o Conecta Vale."""
    try:
        q = st.query_params.get("versao", "")
    except Exception:
        q = ""
    if q in ("modelo", "branco", "assinatura", "servico"):
        return q
    if os.environ.get("PNBOX_VERSAO") in ("modelo", "branco", "assinatura", "servico"):
        return os.environ["PNBOX_VERSAO"]
    arq = APP_DIR / "versao.txt"
    if arq.exists():
        v = arq.read_text(encoding="utf-8").strip().lower()
        if v in ("modelo", "branco", "assinatura", "servico"):
            return v
    return "modelo"


VERSAO = _versao_inicial()

if NAVEGADOR:
    # Pasta guardada no próprio navegador (IndexedDB): sobrevive a atualizações do programa
    PASTA = Path("/home/pyodide/pnbox")
elif MODO_WEB:
    if "sessao_web" not in ss:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            ss.sessao_web = get_script_run_ctx().session_id
        except Exception:
            import uuid
            ss.sessao_web = uuid.uuid4().hex
    PASTA = Path(tempfile.gettempdir()) / "pnbox" / ss.sessao_web
elif (APP_DIR / "instalado.txt").exists():
    PASTA = Path.home() / "Documents" / "ViabilizaLab"
    _antiga = Path.home() / "Documents" / "PNBOX"
    if _antiga.is_dir() and not PASTA.exists():
        # migração da versão anterior (PNBOX Acadêmico): copia os projetos para a nova pasta
        try:
            shutil.copytree(_antiga, PASTA)
        except Exception:
            PASTA = _antiga
else:
    PASTA = APP_DIR / "projetos"
PASTA.mkdir(parents=True, exist_ok=True)

for k, v in {"ver": {}, "buf": {}, "snap": {}, "hash": {}, "pagina": "inicio", "gen": 0,
             "alerta": None, "ultimo_salvo": None, "tema": "claro"}.items():
    ss.setdefault(k, v)

# ------------------------------------------------------------------------------
# LISTAS PADRÃO (o aluno pode acrescentar outras em Configurações)
# ------------------------------------------------------------------------------
LISTAS_PADRAO = {
    "unidade": ["kg", "g", "L", "ml", "unidade", "dúzia", "pacote", "maço", "caixa", "kit", "galão",
                "lata", "par", "m", "m²", "hora", "serviço"],
    "categoria_material": ["Matéria-prima", "Embalagem", "Limpeza e higiene", "Material de consumo",
                           "Mercadoria para revenda", "Outros insumos"],
    "dimensao_interna": ["Pessoas", "Marketing", "Finanças", "Operação", "Tecnologia", "Gestão"],
    "dimensao_externa": ["Político-Legal", "Econômico", "Social", "Tecnológico", "Ambiental",
                         "Demográfico", "Concorrencial"],
    "valor_proposta": ["Conforto e conveniência", "Economia de dinheiro", "Economia de tempo", "Esforço reduzido",
                       "Rapidez", "Segurança", "Variedade", "Personalização", "Qualidade", "Design e estética",
                       "Durabilidade", "Bem-estar", "Status", "Exclusividade", "Comunidade e pertencimento",
                       "Realização pessoal", "Sustentabilidade", "Confiança", "Desempenho", "Simplicidade",
                       "Aprendizado e autodesenvolvimento", "Integração de produtos"],
    "area_pessoal": ["Produção / operação", "Administrativo", "Comercial / vendas"],
    "classe_custo_fixo": ["Custo indireto de produção", "Custo de tecnologia / infraestrutura",
                          "Despesa administrativa", "Despesa comercial", "Despesa financeira"],
    "setor": ["Comércio", "Serviço", "Indústria", "Agronegócio", "Alimentos e bebidas", "Bares e restaurantes",
              "Varejo", "Moda", "Saúde e bem-estar", "Tecnologia", "Educação", "Turismo", "Beleza e estética"],
    "canal": ["Marketing de Conteúdo", "Publicidade em Redes Sociais", "Marketing Viral", "Comunidade",
              "Plataformas Existentes", "Acordos e Parcerias", "Eventos", "Indicação (boca a boca)",
              "Vendas Diretas", "E-mail e WhatsApp"],
}
NOMES_LISTAS = {
    "unidade": "Unidades de medida", "categoria_material": "Categorias de materiais",
    "dimensao_interna": "Dimensões internas (forças e fraquezas)",
    "dimensao_externa": "Dimensões externas (oportunidades e ameaças)",
    "valor_proposta": "Valores da proposta de valor", "area_pessoal": "Áreas da mão de obra",
    "classe_custo_fixo": "Classificação dos custos fixos", "setor": "Setores de atuação",
    "canal": "Canais de aquisição",
}
ETAPAS_FASE = {
    "Antes": ["Descoberta", "Consideração", "Decisão"],
    "Durante": ["Compra e Pagamento", "Entrega", "Uso"],
    "Depois": ["Avaliação", "Suporte", "Lealdade e advocacia"],
}
ETAPAS = [e for v in ETAPAS_FASE.values() for e in v]
EMOCOES = {"Entusiasmado": "🤩", "Satisfeito": "🙂", "Neutro": "😐", "Insatisfeito": "🙁", "Frustrado": "😠"}

# ------------------------------------------------------------------------------
# PÁGINAS (ferramentas) DO PLANO
# ------------------------------------------------------------------------------
PAGINAS = [
    ("trilha", "Identificação", "🚦", "Trilha guiada e o que falta",
     "A ordem recomendada de preenchimento, o próximo passo e tudo o que ainda falta, com atalho para cada ajuste."),
    ("projeto", "Identificação", "🏢", "Informações e sócios",
     "Dados do negócio, integrantes do grupo, premissas financeiras e quadro societário."),
    ("identidade", "Identificação", "🧭", "Missão, visão, valores e objetivos",
     "Por que a empresa existe, aonde quer chegar, em que acredita e objetivos SMART que transformam a visão em metas."),
    ("segmentos", "Cliente e mercado", "🎯", "Segmentação de mercado",
     "Liste os grupos de clientes possíveis e dê notas para descobrir qual atender primeiro."),
    ("personas", "Cliente e mercado", "🧑", "Gerador de personas",
     "Crie personagens que representam o cliente ideal: dados, características, rotina e insights."),
    ("jornada", "Cliente e mercado", "🧭", "Jornada do cliente",
     "Descreva o que a persona faz, pensa e sente antes, durante e depois da compra."),
    ("entrevistas", "Cliente e mercado", "🎙️", "Entrevistas com clientes",
     "Roteiro de entrevista de descoberta e registro das conversas: quais dores se repetem e quanto pagariam."),
    ("mercado", "Cliente e mercado", "📐", "Tamanho do mercado",
     "Estime o mercado total (TAM), o alcançável (SAM) e o conquistável (SOM) e confira se a meta de vendas é realista."),
    ("dados_publicos", "Cliente e mercado", "🏙️", "Dados do município e das famílias (IBGE)",
     "População, domicílios, PIB e empresas do município, e quanto as famílias gastam em cada categoria (POF)."),
    ("proposta", "Problema e solução", "❤️", "Proposta de valor",
     "Defina, para cada persona, os benefícios que o negócio entrega e o que o diferencia."),
    ("vpc", "Problema e solução", "🎁", "Canvas da proposta de valor",
     "Tarefas, dores e ganhos do cliente x produtos, aliviadores e criadores de ganho, e o encaixe entre eles."),
    ("concorrencia", "Problema e solução", "🔍", "Análise da concorrência",
     "Compare o seu negócio com os concorrentes em notas de 0 a 10 por atributo."),
    ("oceano", "Problema e solução", "🌊", "Curva de valor e ERRC",
     "Compare a sua curva com a dos concorrentes e decida o que eliminar, reduzir, elevar e criar (oceano azul)."),
    ("posicionamento", "Problema e solução", "📍", "Posicionamento e mapa perceptual",
     "A frase que diz para quem, o que e por que é diferente, e onde a marca fica frente aos concorrentes."),
    ("forcas", "Estratégia", "💪", "Forças e fraquezas",
     "Fatores internos: o que o negócio faz bem e o que precisa melhorar."),
    ("oportunidades", "Estratégia", "💡", "Oportunidades e ameaças",
     "Fatores externos: tendências que ajudam e riscos que vêm do ambiente."),
    ("swot", "Estratégia", "🧩", "Análise SWOT",
     "Matriz SWOT gerada automaticamente e a estratégia indicada para o negócio."),
    ("swot_cruzada", "Estratégia", "🔀", "SWOT cruzada",
     "Transforme o diagnóstico em estratégias ofensivas, de confronto, de reforço e defensivas."),
    ("canvas", "Análise estratégica", "🗂️", "Business Model Canvas",
     "O modelo de negócio em uma página, montado com o que você já preencheu."),
    ("lean", "Análise estratégica", "🪧", "Lean Canvas",
     "A versão do Canvas para negócios novos: problema, solução, métricas-chave e vantagem injusta."),
    ("porter", "Análise estratégica", "⚔️", "Cinco Forças de Porter",
     "Pressão competitiva do setor: rivalidade, novos entrantes, substitutos, clientes e fornecedores."),
    ("ge", "Análise estratégica", "🧮", "Matriz GE / McKinsey",
     "Atratividade de cada segmento x força competitiva da empresa: onde investir, ser seletivo ou colher."),
    ("bcg", "Análise estratégica", "🫧", "Matriz BCG",
     "Linhas de produto por crescimento do mercado e participação relativa: estrelas, vacas leiteiras, interrogações e abacaxis."),
    ("bsc", "Análise estratégica", "🗺️", "Balanced Scorecard",
     "Objetivos e indicadores nas quatro perspectivas e o mapa estratégico que liga tudo."),
    ("plano_mkt", "Marketing e operação", "🧭", "Plano de marketing",
     "Situação, objetivos, estratégia, táticas, ações e controle, amarrando as ferramentas de marketing."),
    ("marketing", "Marketing e operação", "📣", "Mix de marketing (4Ps e 7Ps)",
     "Produto, preço, praça e promoção (e, em serviços, pessoas, processos e evidência física) ligados aos números do plano."),
    ("canais", "Marketing e operação", "🛒", "Canais e funil de vendas",
     "Avalie os canais de aquisição e acompanhe a conversão do funil."),
    ("orcamento_mkt", "Marketing e operação", "💸", "Orçamento de marketing e funil AARRR",
     "Gasto e CAC por canal, retorno de cada canal, funil de crescimento, calendário de ações e meta de NPS."),
    ("operacional", "Marketing e operação", "🏭", "Plano operacional",
     "Localização, layout, processos, capacidade produtiva e organograma."),
    ("fornecedores", "Marketing e operação", "🚚", "Fornecedores",
     "Quem fornece cada item, prazos de pagamento e de entrega."),
    ("investimentos", "Finanças", "💰", "Investimentos e fontes",
     "Investimentos fixos, pré-operacionais, capital de giro e de onde vem o dinheiro."),
    ("financiamento", "Finanças", "🏦", "Financiamento (SAC, Price e carência)",
     "Parcelas, juros e custo efetivo de cada empréstimo, e comparação entre propostas de crédito."),
    ("produtos", "Finanças", "📦", "Produtos e materiais",
     "Materiais, produtos e serviços, ficha técnica e custo de cada item."),
    ("sazonalidade", "Finanças", "📅", "Sazonalidade das vendas",
     "Peso de cada mês nas vendas: meses fortes e fracos, e o efeito no fluxo de caixa."),
    ("servicos", "Finanças", "🧑‍💻", "Serviços e custo da hora",
     "Para quem vende tempo e conhecimento: custo da hora produtiva, horas da equipe e preço de cada serviço."),
    ("assinaturas", "Finanças", "💳", "Assinaturas e receita recorrente",
     "Planos mensais, trimestrais ou anuais: assinantes, cancelamentos (churn), MRR, CAC, LTV e custos de tecnologia."),
    ("custos", "Finanças", "🧾", "Custos e despesas",
     "Custos variáveis, mão de obra (produção e administrativa) e custos fixos classificados."),
    ("reforma", "Finanças", "🏛️", "Reforma tributária (simulação)",
     "Simples puro ou IBS/CBS por fora a partir de 2027: imposto da empresa e crédito para os clientes empresas."),
    ("preco", "Finanças", "🏷️", "Formação de preço",
     "Custeio variável e por absorção, preço sugerido pelo markup e ponto de equilíbrio por produto."),
    ("teste_preco", "Finanças", "🙋", "Teste de preço e demanda",
     "Faixa de preço aceitável pelo método Van Westendorp e demanda ajustada pela intenção de compra."),
    ("dre", "Finanças", "📄", "DRE e projeções",
     "Demonstração do resultado dos Anos 1 a 3 e a projeção mês a mês."),
    ("fluxo", "Finanças", "💵", "Fluxo de caixa mensal",
     "Entradas e saídas de dinheiro em cada mês do Ano 1."),
    ("indicadores", "Finanças", "📊", "Indicadores financeiros",
     "Semáforo de saúde, VPL, TIR, payback, pontos de equilíbrio e checagens."),
    ("referencias", "Finanças", "📏", "Referências do setor (IBGE e Sebrae)",
     "Compare margem, peso da folha e produtividade com a média do setor e veja a taxa de sobrevivência das empresas."),
    ("graficos", "Finanças", "📈", "Gráficos",
     "Composição do investimento, dos custos e da receita, ponto de equilíbrio e payback."),
    ("simulador", "Ferramentas complementares", "🎚️", "Simulador de resultados",
     "Cenários otimista e pessimista, simulador “e se?” e análise de sensibilidade."),
    ("montecarlo", "Ferramentas complementares", "🎲", "Monte Carlo e perda aceitável",
     "Milhares de cenários sorteados: a chance de o projeto dar certo e quanto os sócios podem perder."),
    ("adocao", "Ferramentas complementares", "🌱", "Curva de adoção (Bass)",
     "Teste se o crescimento de clientes do plano é plausível pela difusão de produtos novos."),
    ("valuation", "Ferramentas complementares", "💎", "Valuation e rodada de investimento",
     "Quanto vale o negócio por três métodos e quanto um investidor levaria em troca do aporte."),
    ("impacto", "Ferramentas complementares", "🌍", "Impacto social e ODS",
     "Teoria da mudança, indicadores de impacto e Objetivos de Desenvolvimento Sustentável."),
    ("experimentos", "Ferramentas complementares", "🧪", "Quadro de experimentação",
     "Hipóteses e testes para validar a ideia antes de investir."),
    ("riscos", "Ferramentas complementares", "⚠️", "Riscos e contingência",
     "Mapa de riscos (probabilidade x impacto) e o que fazer em cada caso."),
    ("cronograma", "Ferramentas complementares", "🗓️", "Cronograma de implantação",
     "O que precisa ser feito antes de abrir as portas, e quando."),
    ("realizado", "Ferramentas complementares", "📏", "Orçado x realizado",
     "Depois de abrir, lance os números reais e compare com o planejado."),
    ("ia", "Ferramentas complementares", "🤖", "Analisar com IA",
     "Monte o texto pronto para colar no ChatGPT, Gemini ou Claude e receber análises e estratégias."),
    ("evidencias", "Validação e defesa", "🔬", "Placar de evidências",
     "Quais números do plano são palpite, pesquisa ou teste com clientes, e qual premissa validar primeiro."),
    ("prontidao", "Validação e defesa", "🪜", "Régua de prontidão",
     "Em que degrau, de 1 a 9, o negócio está: da ideia descrita ao pronto para lançar ou buscar investimento."),
    ("banca", "Validação e defesa", "🧑‍⚖️", "Banca simulada",
     "Perguntas difíceis geradas a partir dos pontos fracos do plano, para treinar a defesa."),
    ("formalizacao", "Validação e defesa", "📋", "Checklist de formalização",
     "Enquadramento (MEI, ME, EPP), registros, licenças, LGPD e marca."),
    ("editais", "Validação e defesa", "📝", "Editais de fomento (modelo Centelha)",
     "Rascunho da proposta para editais de inovação, com limite de caracteres e orçamento por rubrica."),
    ("sumario", "Conclusão e entrega", "📝", "Sumário executivo",
     "Resumo do plano gerado automaticamente, para revisar e completar."),
    ("apresentacao", "Conclusão e entrega", "🎤", "Modo apresentação",
     "Telas grandes com os principais números para apresentar o plano."),
    ("relatorio", "Conclusão e entrega", "📤", "Relatório e exportação",
     "PDF completo, pitch deck, Word, Excel e arquivo do projeto."),
    ("professor", "Conclusão e entrega", "🎓", "Modo professor",
     "Compare os projetos da turma, aplique a rubrica e gere fichas de avaliação."),
    ("glossario", "Conclusão e entrega", "📖", "Glossário e leitura dos números",
     "O que significa cada termo e o que os números deste projeto querem dizer, em linguagem simples."),
    ("planilha", "Conclusão e entrega", "📑", "Importar e exportar planilha",
     "Traga produtos, custos, entrevistas e respostas de pesquisa de uma planilha e exporte tudo para o Excel."),
    ("historico", "Conclusão e entrega", "🕑", "Histórico de versões do projeto",
     "Salve versões, compare os indicadores entre elas e restaure uma versão anterior."),
    ("novidades", "Conclusão e entrega", "🆕", "Novidades do programa",
     "O que mudou em cada versão do ViabilizaLab."),
    ("config", "Conclusão e entrega", "⚙️", "Configurações e listas",
     "Listas personalizadas (unidades, categorias...), versões salvas e dados deste computador."),
]
PAG = {p[0]: p for p in PAGINAS}
GRUPOS = list(dict.fromkeys(p[1] for p in PAGINAS))


def ir(pagina: str, alerta: str | None = None):
    ss.pagina = pagina
    ss.alerta = (pagina, alerta) if alerta else None


# ------------------------------------------------------------------------------
# TEMA (CLARO / ESCURO) E ESTILO VISUAL
# ------------------------------------------------------------------------------
def aplicar_tema():
    base = "dark" if ss.tema == "escuro" else "light"
    try:
        if st.get_option("theme.base") != base:
            st._config.set_option("theme.base", base)
    except Exception:
        pass
    escuro = ss.tema == "escuro"
    fundo_card = "linear-gradient(135deg,#1f4f8f 0%,#2f6fd0 100%)" if not escuro else \
        "linear-gradient(135deg,#1b2d4f 0%,#24467d 100%)"
    css_escuro = """
      .stApp, [data-testid="stHeader"] {background-color:#0f1724 !important; color:#e6e9ef !important;}
      [data-testid="stSidebar"] {background-color:#151f30 !important;}
      .stApp p, .stApp span, .stApp label, .stApp li, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
      .stApp div[data-testid="stMarkdownContainer"] {color:#e6e9ef;}
      [data-testid="stMetricValue"] {color:#ffffff !important;}
      .pn-card-info {background:#1a2638 !important; border-color:#2b3a52 !important;}
    """ if escuro else ""
    st.markdown(f"""
    <style>
      [class*="st-key-card_"] {{
        position: relative; min-height: 186px; height: 100%; border-radius: 16px; padding: 16px 18px 14px 18px;
        background: {fundo_card}; box-shadow: 0 4px 14px rgba(20,40,80,.18); gap: .25rem;
        transition: transform .12s ease, box-shadow .12s ease; cursor: pointer;
      }}
      [class*="st-key-card_"]:hover {{ transform: translateY(-3px); box-shadow: 0 10px 24px rgba(20,40,80,.30); }}
      [class*="st-key-card_"] p, [class*="st-key-card_"] div {{ color: #fff !important; }}
      [class*="st-key-card_"] .pn-ic {{ font-size: 2.1rem; line-height: 1.1; }}
      [class*="st-key-card_"] .pn-tit {{ font-size: 1.08rem; font-weight: 700; margin-top: .35rem; line-height: 1.25; }}
      [class*="st-key-card_"] .pn-desc {{ font-size: .84rem; opacity: .88; margin-top: .3rem; line-height: 1.3; }}
      [class*="st-key-card_"] .pn-selo {{ font-size: .8rem; margin-top: .45rem; opacity: .95; }}
      [class*="st-key-cardbtn_"] {{ position: absolute !important; inset: 0; margin: 0 !important; z-index: 3; }}
      [class*="st-key-cardbtn_"] > div, [class*="st-key-cardbtn_"] .stButton {{ height: 100%; width: 100%; }}
      [class*="st-key-cardbtn_"] button {{ width: 100%; height: 100% !important; opacity: 0; cursor: pointer; }}
      [class*="st-key-cardbtn_"] button:focus-visible {{ opacity: .25; }}
      [class*="st-key-nav_"] button {{ justify-content: flex-start; text-align: left; padding: .25rem .6rem;
        min-height: 2rem; }}
      [class*="st-key-voltar"] button {{ font-weight: 600; }}
      .pn-grupo {{ font-size: 1.15rem; font-weight: 700; margin: 1.2rem 0 .4rem 0; }}
      .pn-card-info {{ border: 1px solid #dfe5ee; border-radius: 12px; padding: 14px 16px; background: #f7f9fc; }}
      .pn-face {{ font-size: 1.8rem; line-height: 1; }}
      .pn-etapa {{ border-left: 4px solid #2f6fd0; padding: 6px 12px; margin: 8px 0; }}
      [data-testid="stAppDeployButton"] {{ display: none; }}
      /* ---------- Celular e telas estreitas ---------- */
      @media (max-width: 700px) {{
        [data-testid="stMainBlockContainer"], .block-container {{ padding: 2.6rem .7rem 5rem .7rem !important; }}
        h1 {{ font-size: 1.55rem !important; }}
        h2 {{ font-size: 1.3rem !important; }}
        h3 {{ font-size: 1.12rem !important; }}
        /* grade de cartões: 2 por linha, compactos */
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) {{ flex-wrap: wrap !important; gap: .55rem !important; }}
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) > [data-testid="stColumn"] {{
          flex: 1 1 calc(50% - .55rem) !important; min-width: calc(50% - .55rem) !important; width: calc(50% - .55rem) !important; }}
        [data-testid="stHorizontalBlock"]:has([class*="st-key-card_"]) > [data-testid="stColumn"]:not(:has([class*="st-key-card_"])) {{
          display: none !important; }}
        [class*="st-key-card_"] {{ min-height: 118px; padding: 11px 12px 10px 12px; border-radius: 14px; }}
        [class*="st-key-card_"] .pn-ic {{ font-size: 1.6rem; }}
        [class*="st-key-card_"] .pn-tit {{ font-size: .92rem; margin-top: .2rem; }}
        [class*="st-key-card_"] .pn-desc {{ display: none; }}
        [class*="st-key-card_"] .pn-selo {{ font-size: .72rem; margin-top: .3rem; }}
        .pn-grupo {{ font-size: 1.02rem; margin: .9rem 0 .3rem 0; }}
        /* números e botões maiores para o dedo */
        [data-testid="stMetricValue"] {{ font-size: 1.25rem !important; }}
        [data-testid="stMetricLabel"] p {{ font-size: .8rem !important; }}
        .stButton button, .stDownloadButton button {{ min-height: 44px; }}
        [data-baseweb="tab-list"] {{ overflow-x: auto; scrollbar-width: none; }}
        [data-baseweb="tab"] {{ padding-left: .55rem !important; padding-right: .55rem !important; white-space: nowrap; }}
        .pn-card-info {{ padding: 11px 12px; font-size: .92rem; }}
      }}
      {css_escuro}
    </style>""", unsafe_allow_html=True)


aplicar_tema()


# ------------------------------------------------------------------------------
# GERENCIAMENTO DE PROJETOS (cada projeto = um arquivo .pnbox)
# ------------------------------------------------------------------------------
def listar_projetos():
    return sorted(PASTA.glob("*.pnbox"), key=lambda p: p.stem.lower())


def caminho_livre(nome: str) -> Path:
    base = re.sub(r"[^\w\- ()]", "", nome, flags=re.UNICODE).strip() or "Projeto"
    p, i = PASTA / f"{base}.pnbox", 2
    while p.exists():
        p, i = PASTA / f"{base} ({i}).pnbox", i + 1
    return p


def trocar_projeto(caminho):
    ss.db = str(caminho)
    ss.gen += 1
    for k in ("buf", "snap", "hash"):
        ss[k] = {}
    ss.pagina = "inicio"
    ss.alerta = None


if not listar_projetos():
    antigo = APP_DIR / "plano_negocio_completo.pnbox"
    if antigo.exists():
        shutil.copy(antigo, PASTA / "Projeto da versão anterior.pnbox")
    if VERSAO == "branco":
        criar_projeto_em_branco(PASTA / "Meu Projeto.pnbox", "Meu Projeto")
    elif VERSAO == "assinatura":
        criar_projeto_assinatura(PASTA / "Conecta Vale (modelo assinatura).pnbox")
    elif VERSAO == "servico":
        criar_projeto_servico(PASTA / "TecVale Soluções (modelo serviços).pnbox")
    else:
        criar_projeto_modelo(PASTA / "FitBox (modelo).pnbox")

projetos = listar_projetos()
if "db" not in ss or not Path(ss.db).exists():
    trocar_projeto(projetos[0])

# ------------------------------------------------------------------------------
# CARREGAMENTO DOS DADOS DO PROJETO ATIVO
# ------------------------------------------------------------------------------
conn = conectar(Path(ss.db))
proj = ler_unica(conn, "projeto")
d = {t: ler(conn, t) for t in SCHEMA if t not in ("projeto", "proposta_valor")}
d["proposta_valor"] = ler_unica(conn, "proposta_valor")
PID = f"{Path(ss.db).stem}_{ss.gen}"


def opcoes(tipo: str, extras=()) -> list:
    """Lista padrão + itens personalizados do projeto + valores já usados nos dados."""
    pers = d["listas"].loc[d["listas"]["tipo"] == tipo, "valor"].tolist() if "listas" in d else []
    out = []
    for v in list(LISTAS_PADRAO.get(tipo, [])) + pers + [x for x in extras if x]:
        if v and v not in out:
            out.append(v)
    return out


def marcar_salvo():
    ss.ultimo_salvo = datetime.now().strftime("%H:%M:%S")


# ------------------------------------------------------------------------------
# MOTOR FINANCEIRO SOB DEMANDA
# ------------------------------------------------------------------------------
_CACHE = {}


def R() -> dict:
    """Resultado do motor com os dados atuais (recalcula quando algo muda)."""
    if "res" not in _CACHE:
        _CACHE["res"] = executar_motor(proj, d)
    return _CACHE["res"]


def invalidar():
    _CACHE.clear()


def CHK():
    if "chk" not in _CACHE:
        _CACHE["chk"] = tabela_checagens_v5(R(), d, proj)
    return _CACHE["chk"]


# ------------------------------------------------------------------------------
# TABELAS EDITÁVEIS COM SALVAMENTO AUTOMÁTICO
# ------------------------------------------------------------------------------
def _h(df: pd.DataFrame) -> str:
    try:
        return hashlib.md5(df.to_csv(index=False).encode("utf-8")).hexdigest()
    except Exception:
        return ""


def chave(tabela, extra=""):
    return f"{PID}_{tabela}_{extra}_{ss.ver.get((PID, tabela, extra), 0)}"


def _nova_versao(tabela, extra=""):
    ss.ver[(PID, tabela, extra)] = ss.ver.get((PID, tabela, extra), 0) + 1


def tabela_editavel(df_base, config, tabela, extra="", normalizar=None, salvar_fn=None, altura=None,
                    adicionar=True):
    """
    Tabela editável com botões "Adicionar linha" e "Apagar marcadas" e SALVAMENTO AUTOMÁTICO:
    toda alteração válida é gravada no projeto na hora, sem precisar clicar em Salvar.
    """
    kb = (PID, tabela, extra)
    normalizar = normalizar or (lambda x: x)
    wkey = chave(tabela, extra)
    if wkey not in ss.snap:
        ss.snap[wkey] = normalizar(ss.buf.get(kb, df_base).copy().reset_index(drop=True))
        ss.hash.setdefault(kb, _h(normalizar(df_base.copy().reset_index(drop=True))))
    base = ss.snap[wkey].copy()
    base.insert(0, "Excluir", False)
    cfg = {"Excluir": st.column_config.CheckboxColumn("🗑️", help="Marque as linhas que deseja apagar", width="small")}
    cfg.update(config)
    ed = st.data_editor(base, num_rows="dynamic" if adicionar else "fixed", width="stretch", hide_index=True,
                        column_config=cfg, key=wkey, **({"height": altura} if altura else {}))
    marcadas = ed["Excluir"].fillna(False).astype(bool)
    dados = normalizar(ed.drop(columns="Excluir"))

    def _gravar(df_novo):
        if salvar_fn:
            salvar_fn(df_novo)
        ss.hash[kb] = _h(df_novo)
        marcar_salvo()

    # Salvamento automático
    if _h(dados) != ss.hash.get(kb):
        _gravar(dados)
        invalidar()

    b1, b2, b3 = st.columns([1.2, 1.6, 3])
    if adicionar and b1.button("➕ Adicionar linha", key="add_" + wkey, width="stretch"):
        vazia = normalizar(pd.DataFrame([{c: None for c in dados.columns}]))
        ss.buf[kb] = pd.concat([normalizar(ed[~marcadas].drop(columns="Excluir")), vazia], ignore_index=True)
        _nova_versao(tabela, extra)
        st.rerun()
    if b2.button(f"🗑️ Apagar marcadas ({int(marcadas.sum())})", key="del_" + wkey,
                 disabled=not marcadas.any(), width="stretch"):
        novo = normalizar(ed[~marcadas].drop(columns="Excluir"))
        ss.buf[kb] = novo
        _gravar(novo)
        _nova_versao(tabela, extra)
        st.toast("Linhas apagadas.", icon="🗑️")
        st.rerun()
    b3.caption(f"💾 Salvo automaticamente{' às ' + ss.ultimo_salvo if ss.ultimo_salvo else ''}")
    return dados


def editor(tabela, config, campo_chave, altura=None, filtro=None):
    """
    Tabela editável ligada diretamente a uma tabela do projeto.
    filtro={coluna: valor}: edita só as linhas daquele grupo, mantendo as demais intactas.
    """
    filtro = filtro or {}
    if filtro:
        mask = pd.Series(True, index=d[tabela].index)
        for col, val in filtro.items():
            mask &= d[tabela][col] == val
        df_base = d[tabela][mask].reset_index(drop=True)
        extra = "|".join(f"{c}={v}" for c, v in filtro.items())
    else:
        df_base, extra = d[tabela], ""

    def norm(x):
        x = limpar(x, tabela)
        for col, val in filtro.items():
            x[col] = val
        return x

    def gravar(df_novo):
        if filtro:
            m = pd.Series(True, index=d[tabela].index)
            for col, val in filtro.items():
                m &= d[tabela][col] == val
            completo = pd.concat([d[tabela][~m], df_novo], ignore_index=True)
        else:
            completo = df_novo
        salvar(conn, tabela, completo, campo_chave)
        validos = completo[completo[campo_chave].astype(str).str.strip() != ""] if campo_chave else completo
        d[tabela] = limpar(validos.reset_index(drop=True), tabela)

    cfg = {k: v for k, v in config.items() if k not in filtro}
    cfg.update({k: None for k in filtro})          # esconde as colunas do filtro
    ed = tabela_editavel(df_base, cfg, tabela, extra, normalizar=norm, salvar_fn=gravar, altura=altura)
    if not filtro:
        d[tabela] = ed[ed[campo_chave].astype(str).str.strip() != ""].reset_index(drop=True) if campo_chave else ed
    invalidar()
    if not filtro and tabela in LEITURAS:
        try:
            frases = LEITURAS[tabela](limpar(d[tabela], tabela), proj, d, R)
        except Exception:
            frases = []
        mostrar_leitura(tabela, frases)
    return ed


def mostrar_leitura(chave_: str, frases: list):
    """Caixa "Leitura dos resultados" ao final de uma tabela. No modo exercício, o aluno escreve antes de ver."""
    if not frases or not ss.get("tg_leit", True):
        return
    with st.container(border=True):
        st.markdown("**📌 Leitura dos resultados**")
        if ss.get("tg_exerc", False):
            la = d["leituras_aluno"]
            atual = la.loc[la["tabela"] == chave_, "texto"]
            atual = str(atual.iloc[0]) if not atual.empty else ""
            txt = st.text_area("Antes de ver a leitura do programa, escreva a sua: o que esta tabela mostra?", atual,
                               key=wk("la_" + chave_), height=90)
            if txt != atual:
                novo = pd.concat([la[la["tabela"] != chave_], pd.DataFrame([{"tabela": chave_, "texto": txt}])], ignore_index=True)
                salvar(conn, "leituras_aluno", novo, "tabela")
                d["leituras_aluno"] = limpar(novo, "leituras_aluno")
                marcar_salvo()
            if len(txt.strip()) < 20:
                st.caption("A leitura do programa aparece depois que você escrever ao menos uma frase.")
                return
            st.caption("Compare com a leitura do programa:")
        st.markdown("\n".join(f"- {f}" for f in frases))


def auto_unica(tabela: str, novos: dict):
    """Salva automaticamente campos de tabelas de linha única (projeto, proposta) quando mudam."""
    atual = proj if tabela == "projeto" else d["proposta_valor"]
    mudou = {k: v for k, v in novos.items() if str(atual.get(k, "")) != str(v)}
    if mudou:
        salvar_unica(conn, tabela, mudou)
        atual.update(mudou)
        marcar_salvo()
        invalidar()


def wk(nome: str) -> str:
    """Chave de widget única por projeto (evita mostrar dados de outro projeto)."""
    return f"{PID}_w_{nome}"


T = st.column_config.TextColumn
N = st.column_config.NumberColumn
S = st.column_config.SelectboxColumn
CK = st.column_config.CheckboxColumn


def moeda(rot, **kw):
    return N(rot, min_value=0.0, format="R$ %.2f", **kw)


def nota(rot, maximo=10):
    return N(rot, min_value=0, max_value=maximo, step=1, format="%d")


# ------------------------------------------------------------------------------
# SITUAÇÃO DE CADA FERRAMENTA (selo nos cartões)
# ------------------------------------------------------------------------------
def _tem(tab, minimo=1):
    return len(d[tab]) >= minimo


def _receita() -> bool:
    return _tem("catalogo") or _tem("assinaturas")


def situacao(p: str) -> str:
    """✅ completo · ◐ em andamento · ○ não iniciado"""
    sw = d["swot"]
    regras = {
        "projeto": (bool(proj["nome_negocio"]) and bool(proj["cidade_uf"]), _tem("socios")),
        "segmentos": (_tem("segmentos", 2), _tem("segmentos")),
        "personas": (_tem("personas") and bool(d["personas"]["desafios"].astype(str).str.strip().ne("").any()), _tem("personas")),
        "jornada": (d["jornada_cliente"]["etapa"].nunique() >= 6, _tem("jornada_cliente")),
        "mercado": (proj["tam_clientes"] > 0 and proj["sam_clientes"] > 0 and proj["som_pct"] > 0 and proj["gasto_anual_cliente"] > 0,
                    proj["tam_clientes"] > 0 or proj["gasto_anual_cliente"] > 0),
        "proposta": (_tem("propostas"), bool(d["proposta_valor"]["segmento_alvo"])),
        "concorrencia": (_tem("concorrentes") and _tem("concorrencia_notas", 3), _tem("concorrencia_notas")),
        "forcas": ((sw["natureza"] == "Força").any() and (sw["natureza"] == "Fraqueza").any(),
                   sw["natureza"].isin(["Força", "Fraqueza"]).any()),
        "oportunidades": ((sw["natureza"] == "Oportunidade").any() and (sw["natureza"] == "Ameaça").any(),
                          sw["natureza"].isin(["Oportunidade", "Ameaça"]).any()),
        "swot": (sw["natureza"].nunique() >= 4, len(sw) > 0),
        "marketing": (all(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao")),
                      any(proj[k] for k in ("mkt_produto", "mkt_preco", "mkt_praca", "mkt_promocao"))),
        "canais": (_tem("canais") and _tem("funil_vendas"), _tem("canais") or _tem("funil_vendas")),
        "operacional": (bool(proj["loc_endereco"]) and _tem("capacidade"), bool(proj["loc_endereco"]) or _tem("capacidade")),
        "fornecedores": (_tem("fornecedores", 2), _tem("fornecedores")),
        "investimentos": (_tem("inv_fixos") and (_tem("socios") or _tem("emprestimos")), _tem("inv_fixos")),
        "produtos": (_tem("catalogo") and _tem("materiais"), _tem("catalogo") or _tem("materiais")),
        "custos": (_tem("pessoal") and _tem("custos_fixos"), _tem("pessoal") or _tem("custos_fixos")),
        "identidade": (bool(proj["missao"]) and bool(proj["visao"]) and bool(proj["valores"]) and _tem("objetivos", 3) and
                       all(all(ok for ok, _ in checar_smart(o).values()) for _, o in d["objetivos"].iterrows()),
                       bool(proj["missao"]) or bool(proj["visao"]) or _tem("objetivos")),
        "canvas": (bool(proj["canvas_relacionamento"]) and bool(proj["canvas_parcerias"]), _tem("segmentos") or _tem("propostas")),
        "porter": (len(d["porter"]) >= 5 and bool((d["porter"]["intensidade"] > 0).all()) and
                   not d["porter"]["justificativa"].astype(str).str.startswith("(sugestão)").any(), _tem("porter")),
        "ge": (_tem("segmentos", 2) and _tem("concorrencia_notas"), _tem("segmentos")),
        "bcg": (_tem("bcg", 2), _tem("bcg")),
        "bsc": (set(PERSPECTIVAS) <= set(d["objetivos"]["perspectiva"]), _tem("objetivos")),
        "servicos": (bool((d["catalogo"]["horas_unit"] > 0).any()) and float(proj["utilizacao_pct"]) > 0,
                     bool((d["catalogo"]["horas_unit"] > 0).any())),
        "assinaturas": (_tem("assinaturas") and bool((d["assinaturas"]["churn_mensal_pct"] > 0).all()), _tem("assinaturas")),
        "preco": (_tem("catalogo"), _tem("catalogo")),
        "dre": (_receita() and _tem("custos_fixos"), _receita()),
        "fluxo": (_receita() and _tem("capital_giro_prazos"), _receita()),
        "indicadores": (_receita() and _tem("inv_fixos"), _receita()),
        "graficos": (_receita(), _receita()),
        "simulador": (_receita(), _receita()),
        "experimentos": (_tem("experimentos"), _tem("experimentos")),
        "riscos": (_tem("riscos", 3), _tem("riscos")),
        "cronograma": (_tem("cronograma", 3), _tem("cronograma")),
        "realizado": (_tem("realizado", 1), _tem("realizado")),
        "sumario": (len(proj["sumario_executivo"]) > 200, bool(proj["sumario_executivo"])),
    }
    regras.update(regras_v71(proj, d))
    regras.update(regras_v8(proj, d))
    regras.update(regras_mix(proj, d))
    if p not in regras:
        return ""
    completo, iniciado = regras[p]
    tipo = str(proj.get("tipo_negocio") or "Produtos")
    opcional = (p == "servicos" and not tipo.startswith(("Serviço", "Misto"))) or \
               (p == "assinaturas" and not tipo.startswith(("Assinatura", "Misto"))) or p in ("reforma", "adocao") or p in OPCIONAIS_V8 or p in ("plano_mkt", "sazonalidade")
    if opcional and not iniciado:
        return ""            # ferramenta não se aplica ao tipo de negócio escolhido
    return "✅" if completo else ("◐" if iniciado else "○")


def completude() -> float:
    return completude_de(proj, d)


# ------------------------------------------------------------------------------
# MENU LATERAL
# ------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📊 ViabilizaLab")
    st.caption("Laboratório de plano de negócio e viabilidade (antigo PNBOX Acadêmico)")
    if st.button("🏠 Painel do plano", key="nav_inicio", width="stretch",
                 type="primary" if ss.pagina == "inicio" else "secondary"):
        ir("inicio")
        st.rerun()

    nomes = [p.stem for p in projetos]
    atual = Path(ss.db).stem
    escolha = st.selectbox("📁 Projeto ativo", nomes, index=nomes.index(atual) if atual in nomes else 0)
    if escolha != atual:
        trocar_projeto(PASTA / f"{escolha}.pnbox")
        st.rerun()

    with st.expander("➕ Novo projeto"):
        novo_nome = st.text_input("Nome do novo projeto", placeholder="Ex.: Padaria Pão Nosso")
        tipo = st.radio("Começar a partir de", ["Em branco (do zero)", "Modelo FitBox (alimentação saudável)",
                                                "Modelo Café Mucuri (cafeteria e padaria)",
                                                "Modelo TecVale Soluções (serviços de TI)",
                                                "Modelo Conecta Vale (aplicativo por assinatura)", "Cópia do projeto atual"])
        if st.button("Criar projeto", type="primary", width="stretch"):
            if not novo_nome.strip():
                st.error("Informe um nome.")
            else:
                destino = caminho_livre(novo_nome)
                if tipo.startswith("Em branco"):
                    criar_projeto_em_branco(destino, novo_nome.strip())
                elif tipo.startswith("Modelo FitBox"):
                    criar_projeto_modelo(destino)
                elif tipo.startswith("Modelo Café"):
                    criar_projeto_cafe(destino)
                elif tipo.startswith("Modelo Conecta"):
                    criar_projeto_assinatura(destino)
                elif tipo.startswith("Modelo TecVale"):
                    criar_projeto_servico(destino)
                else:
                    conn.commit()
                    shutil.copy(ss.db, destino)
                c = conectar(destino)
                salvar_unica(c, "projeto", {"nome_negocio": novo_nome.strip()})
                c.close()
                trocar_projeto(destino)
                st.rerun()

    with st.expander("📂 Abrir arquivo .pnbox"):
        arq = st.file_uploader("Selecione um arquivo .pnbox", type=["pnbox"], key=f"up_{ss.gen}")
        if arq is not None and st.button("Abrir este arquivo", width="stretch"):
            dados_arq = arq.getvalue()
            if not dados_arq.startswith(b"SQLite format 3"):
                st.error("Arquivo inválido: não é um projeto .pnbox.")
            else:
                destino = caminho_livre(Path(arq.name).stem)
                destino.write_bytes(dados_arq)
                conectar(destino).close()      # atualiza a estrutura, se for de versão antiga
                trocar_projeto(destino)
                st.rerun()

    with st.expander("🗑️ Excluir projeto atual"):
        confirma = st.checkbox(f"Confirmo a exclusão de “{atual}”")
        if st.button("Excluir", disabled=not confirma, width="stretch"):
            conn.close()
            Path(ss.db).unlink(missing_ok=True)
            ss.pop("db")
            st.rerun()

    st.divider()
    tema_escuro = st.toggle("🌙 Tema escuro", value=ss.tema == "escuro", key="tg_tema")
    if tema_escuro != (ss.tema == "escuro"):
        ss.tema = "escuro" if tema_escuro else "claro"
        st.rerun()

    st.toggle("📌 Leitura automática das tabelas", value=True, key="tg_leit",
              help="Ao final de cada tabela, frases que dizem o que os números mostram.")
    st.toggle("✍️ Modo exercício", value=False, key="tg_exerc",
              help="O aluno escreve a própria leitura de cada tabela antes de ver a do programa. As leituras ficam salvas no projeto.")

    with st.expander("🧭 Ir direto para...", expanded=False):
        for g in GRUPOS:
            st.caption(g)
            for p in [x for x in PAGINAS if x[1] == g]:
                if st.button(f"{p[2]} {p[3]}", key=f"nav_{p[0]}", width="stretch",
                             type="primary" if ss.pagina == p[0] else "tertiary"):
                    ir(p[0])
                    st.rerun()

    st.divider()
    if NAVEGADOR:
        st.info("**Versão on-line:** seus projetos ficam guardados **neste navegador, neste computador**. "
                "Para levar a outro computador ou entregar ao professor, baixe o **.pnbox** em "
                "📤 Relatório e exportação.")
    elif MODO_WEB:
        st.warning("**Versão on-line:** o projeto fica só nesta aba. Ao terminar, baixe o **.pnbox** em "
                   "📤 Relatório e exportação.")
    else:
        st.caption(f"Projetos salvos em:\n`{PASTA}`")
    if not MODO_WEB and st.button("⏻ Encerrar o programa", width="stretch"):
        st.success("Programa encerrado. Pode fechar esta aba.")
        threading.Timer(1.0, lambda: os._exit(0)).start()
        st.stop()
    st.caption(f"Versão {VERSAO_PROGRAMA}")


# ------------------------------------------------------------------------------
# CABEÇALHO DAS FERRAMENTAS
# ------------------------------------------------------------------------------
def rolar_para_topo():
    """Ao trocar de página, volta a tela para o início (útil sobretudo no celular)."""
    try:
        import streamlit.components.v1 as components
    except Exception:          # recurso opcional: se a versão do Streamlit não tiver, apenas não rola a tela
        return
    import logging
    logging.getLogger("streamlit").setLevel(logging.ERROR)
    components.html("""<script>
      try { const d = window.parent.document;
        for (const s of ['[data-testid="stMain"]', '[data-testid="stAppViewContainer"]', 'section.main', '.main']) {
          const el = d.querySelector(s); if (el) el.scrollTo({top: 0}); }
        window.parent.scrollTo({top: 0});
      } catch (e) {}
    </script>""", height=0)


def cabecalho(p: str):
    _, grupo, icone, titulo, descricao = PAG[p]
    if ss.get("_pag_vista") != p:
        ss["_pag_vista"] = p
        rolar_para_topo()
    c1, c2 = st.columns([1, 5])
    if c1.button("← Voltar ao plano", key="voltar_" + p, width="stretch"):
        ir("inicio")
        st.rerun()
    c2.caption(f"{grupo}  ›  {titulo}")
    st.title(f"{icone} {titulo}")
    st.markdown(f"<div class='pn-card-info'>ℹ️ {descricao}</div>", unsafe_allow_html=True)
    if ss.alerta and ss.alerta[0] == p:
        st.warning(f"**O que ajustar aqui:** {ss.alerta[1]}")
        if st.button("Entendi, ocultar aviso", key="ok_alerta"):
            ss.alerta = None
            st.rerun()
    comentarios_da_pagina(p)
    st.write("")


def proxima(p: str):
    """Botões de navegação para a ferramenta anterior e a próxima."""
    chaves = [x[0] for x in PAGINAS if x[0] not in ("professor", "config")]
    if p not in chaves:
        return
    i = chaves.index(p)
    st.divider()
    c1, _, c3 = st.columns([1.3, 3, 1.3])
    if i > 0 and c1.button(f"← {PAG[chaves[i - 1]][3]}", key="ant_" + p, width="stretch"):
        ir(chaves[i - 1])
        st.rerun()
    if i < len(chaves) - 1 and c3.button(f"{PAG[chaves[i + 1]][3]} →", key="prox_" + p, width="stretch", type="primary"):
        ir(chaves[i + 1])
        st.rerun()

# ##############################################################################
# PARTE: pnbox_ui_paginas1.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: PAINEL E FERRAMENTAS DE ESTRATÉGIA E MERCADO
# ==============================================================================

# ------------------------------------------------------------------------------
# Registros editados um a um (personas, propostas): salvamento automático por id
# ------------------------------------------------------------------------------
def ler_com_id(tabela):
    cols = [c for c, _ in SCHEMA[tabela]]
    return pd.read_sql(f"SELECT id, {', '.join(cols)} FROM {tabela} ORDER BY id", conn)


def auto_registro(tabela, rid, atual: dict, novos: dict):
    mudou = {k: v for k, v in novos.items() if str(atual.get(k) if atual.get(k) is not None else "") != str(v)}
    if mudou:
        conn.execute(f"UPDATE {tabela} SET {', '.join(k + '=?' for k in mudou)} WHERE id=?",
                     list(mudou.values()) + [rid])
        conn.commit()
        d[tabela] = ler(conn, tabela)
        marcar_salvo()
        invalidar()


def novo_registro(tabela, valores: dict):
    cols = list(valores)
    cur = conn.execute(f"INSERT INTO {tabela} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                       [valores[c] for c in cols])
    conn.commit()
    return cur.lastrowid


def apagar_registro(tabela, rid):
    conn.execute(f"DELETE FROM {tabela} WHERE id=?", (rid,))
    conn.commit()


def salvo_txt():
    st.caption(f"💾 Salvo automaticamente{' às ' + ss.ultimo_salvo if ss.ultimo_salvo else ''}")


# ==============================================================================
# PAINEL DO PLANO (tela inicial)
# ==============================================================================
def pg_inicio():
    if ss.get("_pag_vista") != "inicio":
        ss["_pag_vista"] = "inicio"
        rolar_para_topo()
    if not ss.get("bemvindo_ok"):
        with st.container(border=True):
            st.markdown(f"""
### 👋 Bem-vindo ao ViabilizaLab
Este programa ajuda você a construir um **plano de negócio completo**, do cliente aos números, e a descobrir
se a ideia é **viável**. Como usar:

1. **Clique nos cartões** abaixo, de cima para baixo. Cada cartão é uma ferramenta, com explicação de como preencher.
2. **Tudo é salvo automaticamente** enquanto você digita. Não existe botão "Salvar".
3. O selo do cartão mostra o andamento: ✅ concluído · ◐ em andamento · ○ não iniciado.
4. Os **resultados financeiros** (DRE, indicadores, gráficos) são calculados sozinhos a partir do que você preencher.
5. No final, vá em **📤 Relatório e exportação** para baixar o PDF, o pitch, o Word e o arquivo **.pnbox** do projeto.

{"Você está na **versão on-line**: os dados ficam neste navegador. Para entregar ou levar a outro computador, baixe o arquivo .pnbox." if MODO_WEB else ""}
""")
            if st.button("Entendi, começar", type="primary"):
                ss.bemvindo_ok = True
                st.rerun()

    st.title(f"📊 {proj['nome_negocio'] or 'Meu plano de negócio'}")
    linha = " · ".join(x for x in [proj["cidade_uf"], proj["turma"], proj["autores"]] if x)
    if linha:
        st.caption(linha)

    comp = completude()
    try:
        res = R()
        chk = CHK()
        sem = semaforo(res)
        rx = raio_x(res, chk, comp, sem)
    except Exception as e:
        res, chk, sem, rx = None, None, [], None
        st.info(f"Os resultados financeiros aparecem quando houver produtos, custos e investimentos. ({e})")

    c1, c2, c3, c4 = st.columns([2.2, 1, 1, 1])
    with c1:
        st.markdown(f"**Plano {int(comp * 100)}% preenchido**")
        st.progress(comp)
        prox = proximo_passo()
        if prox and st.button(f"👉 Próximo passo: {PAG[prox][3]}", key="ini_prox"):
            ir(prox)
            st.rerun()
    if rx:
        c2.metric("Raio-X do plano", f"{rx['nota']}/100", help="Preenchimento (30) + checagens OK (30) + saúde financeira (40)")
        c3.metric("Investimento total", brl(res["investimento_total"]).split(",")[0],
                  help=brl(res["investimento_total"]))
        c4.metric("Lucro no Ano 1", brl(res["anos"][1]["resultado"]).split(",")[0],
                  help=brl(res["anos"][1]["resultado"]))
        verdes = sum(1 for s in sem if s["Status"] == "🟢")
        if res["anos"][1]["rec_bruta"] <= 0:
            st.caption("Saúde financeira: aparece quando houver produtos com preço e quantidade cadastrados.")
        else:
            st.caption(f"Saúde financeira: {' '.join(s['Status'] for s in sem)}  ({verdes} de {len(sem)} indicadores no verde)")

    if chk is not None:
        pend = chk[chk["Situação"] != "OK"]
        if not pend.empty:
            with st.expander(f"⚠️ {len(pend)} ponto(s) para corrigir — clique para ver e ir direto ao ajuste", expanded=False):
                for i, r in pend.iterrows():
                    a, b = st.columns([4, 1.3])
                    a.markdown(f"**{r['Teste']}** — {r['Valor apurado']}  \n<small>{r['Como corrigir'] or r['Observação']}</small>",
                               unsafe_allow_html=True)
                    if r["destino"] and b.button("Ir para o ajuste →", key=f"ini_chk_{i}", width="stretch"):
                        ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir'] or r['Observação']}")
                        st.rerun()

    for g in GRUPOS:
        itens = [p for p in PAGINAS if p[1] == g]
        st.markdown(f"<div class='pn-grupo'>{g}</div>", unsafe_allow_html=True)
        for i in range(0, len(itens), 4):
            cols = st.columns(4)
            for col, p in zip(cols, itens[i:i + 4]):
                selo = situacao(p[0])
                txt_selo = {"✅": "✅ concluído", "◐": "◐ em andamento", "○": "○ não iniciado"}.get(selo, "")
                with col.container(key=f"card_{p[0]}"):
                    st.markdown(f"<div class='pn-ic'>{p[2]}</div><div class='pn-tit'>{p[3]}</div>"
                                f"<div class='pn-desc'>{p[4]}</div><div class='pn-selo'>{txt_selo}</div>",
                                unsafe_allow_html=True)
                    if st.button(f"Abrir {p[3]}", key=f"cardbtn_{p[0]}", width="stretch"):
                        ir(p[0])
                        st.rerun()


# ==============================================================================
# IDENTIFICAÇÃO
# ==============================================================================
TIPOS_NEGOCIO = {
    "Produtos": "Fabrica ou revende mercadorias. Use Produtos e materiais, ficha técnica e estoques.",
    "Serviços (hora ou projeto)": "Vende tempo e conhecimento (consultoria, manutenção, design, desenvolvimento sob encomenda, "
                                  "estética...). Use também 🧑‍💻 Serviços e custo da hora.",
    "Assinatura (receita recorrente)": "Cobra planos que se renovam (aplicativos, software, clubes, academias, mensalidades). "
                                       "Use 💳 Assinaturas e receita recorrente.",
    "Misto": "Combina produtos, serviços e/ou assinaturas.",
}


def calculadora_simples():
    res = R()
    fr = fator_r(res, d)
    st.markdown("""
No Simples Nacional a alíquota depende do **anexo** da atividade e da **receita bruta dos últimos 12 meses (RBT12)**:
alíquota efetiva = (RBT12 × alíquota nominal − parcela a deduzir) ÷ RBT12. Serviços de tecnologia, engenharia e outros de
natureza intelectual ficam no **Anexo V**, mas passam ao **Anexo III**, mais barato, quando o **fator R** (folha de pagamento,
incluindo pró-labore, ÷ receita bruta) é de 28% ou mais. Confirme sempre o CNAE e o anexo com um contador.
""")
    c1, c2, c3 = st.columns(3)
    anexos = list(TABELAS_SIMPLES)
    sugerido = proj["anexo_simples"] if proj["anexo_simples"] in anexos else (
        "Anexo III - Serviços" if str(proj["tipo_negocio"]).startswith(("Serviço", "Assinatura")) else "Anexo I - Comércio")
    anexo = c1.selectbox("Anexo da atividade", anexos, index=anexos.index(sugerido), key=wk("anexo"))
    rbt = c2.number_input("Receita bruta de 12 meses (RBT12)", 0.0, 1e9, float(round(res["anos"][1]["rec_bruta"], 2)), 1000.0,
                          key=wk("rbt12"), help="Sugestão: a receita prevista para o Ano 1.")
    usar_r = anexo in ("Anexo III - Serviços", "Anexo V - Serviços")
    anexo_final = fr["anexo_por_fator_r"] if usar_r else anexo
    r = aliquota_simples(anexo_final, rbt)
    c3.metric("Fator R previsto", pct(fr["fator_r"]) if fr["fator_r"] == fr["fator_r"] else "n.a.",
              help=f"Folha de 12 meses {brl(fr['folha12'])} (salários, encargos e pró-labore de {brl(fr['pro_labore_m'])}/mês) "
                   f"÷ receita {brl(fr['receita12'])}")
    if usar_r:
        st.info(f"Com fator R de {pct(fr['fator_r'])}, a atividade é tributada pelo **{anexo_final}**.")
    if r["efetiva"] == r["efetiva"]:
        a, b, c = st.columns(3)
        a.metric("Faixa", "acima do limite" if r.get("acima_limite") else f"{r['faixa']}ª")
        b.metric("Alíquota nominal", pct(r["nominal"]))
        c.metric("Alíquota efetiva", pct(r["efetiva"]))
        if st.button(f"Usar {pct(r['efetiva'])} como imposto sobre vendas", key=wk("usar_aliq")):
            salvar_unica(conn, "projeto", {"aliquota_imposto_pct": round(r["efetiva"] * 100, 2), "anexo_simples": anexo})
            ss.pop(wk("aliq"), None)
            invalidar()
            marcar_salvo()
            st.rerun()
    st.caption("Tabelas da LC 123/2006 com a redação da LC 155/2016. A reforma tributária (EC 132/2023 e LC 214/2025) mantém o "
               "Simples, mas a partir de 2027 a empresa pode recolher IBS e CBS dentro da guia única ou por fora, no regime "
               "regular. Para o plano, use a regra vigente, cite a fonte e confirme com um contador.")


def pg_projeto():
    cabecalho("projeto")
    t1, t2, t3 = st.tabs(["🏢 Negócio", "👥 Grupo e entrega", "📈 Premissas financeiras"])
    v = {}
    with t1:
        tipos = list(TIPOS_NEGOCIO)
        atual_t = proj["tipo_negocio"] if proj["tipo_negocio"] in tipos else tipos[0]
        v["tipo_negocio"] = st.radio("Como o negócio ganha dinheiro?", tipos, index=tipos.index(atual_t), horizontal=True,
                                     key=wk("tipo_neg"), help="Define quais ferramentas o programa recomenda e quais checagens aplica.")
        st.caption(TIPOS_NEGOCIO[v["tipo_negocio"]])
        c1, c2, c3 = st.columns(3)
        v["nome_negocio"] = c1.text_input("Nome do negócio", proj["nome_negocio"], key=wk("nome"), max_chars=60)
        objs = ["Criar um novo negócio", "Expandir um negócio existente", "Captar recursos", "Outro"]
        v["objetivo"] = c2.selectbox("Objetivo do plano", objs,
                                     index=objs.index(proj["objetivo"]) if proj["objetivo"] in objs else 0, key=wk("obj"))
        v["periodo_anos"] = c3.number_input("Horizonte de planejamento (anos)", 1, 10, int(proj["periodo_anos"] or 3), key=wk("per"),
                                            help="Quantos anos o plano projeta: DRE, fluxos, VPL, TIR e payback usam este período. "
                                                 "Planos de bancos e do Sebrae costumam usar de 3 a 5 anos.")
        regs = ["Simples Nacional", "MEI", "Lucro Presumido", "Lucro Real"]
        v["enquadramento"] = c1.selectbox("Enquadramento tributário", regs, key=wk("enq"),
                                          index=regs.index(proj["enquadramento"]) if proj["enquadramento"] in regs else 0)
        portes = ["MEI", "Microempresa (ME)", "Pequena Empresa (EPP)", "Médio/Grande Porte"]
        v["porte"] = c2.selectbox("Porte", portes, key=wk("porte"),
                                  index=portes.index(proj["porte"]) if proj["porte"] in portes else 1)
        atuais = [x.strip() for x in proj["setor_atuacao"].split(";") if x.strip()]
        v["setor_atuacao"] = "; ".join(c3.multiselect("Setores de atuação", opcoes("setor", atuais), atuais, key=wk("setor"),
                                                      accept_new_options=True))
        v["cidade_uf"] = c1.text_input("Cidade / UF", proj["cidade_uf"], key=wk("cid"))
        v["contatos"] = c2.text_input("Contatos (site, redes, e-mail)", proj["contatos"], key=wk("cont"))
    with t2:
        st.caption("Estes dados aparecem na capa do relatório e no modo professor.")
        c1, c2 = st.columns(2)
        v["autores"] = c1.text_area("Integrantes do grupo (um por linha ou separados por ;)", proj["autores"], key=wk("aut"))
        v["turma"] = c2.text_input("Curso / disciplina / turma", proj["turma"], key=wk("turma"))
        v["professor"] = c2.text_input("Professor(a)", proj["professor"], key=wk("prof"))
        v["data_entrega"] = c2.text_input("Data de entrega", proj["data_entrega"], key=wk("data"), placeholder="dd/mm/aaaa")
    with t3:
        st.caption("Premissas usadas em todas as projeções. Passe o mouse no ícone ? para ver a explicação de cada uma.")
        p1, p2, p3 = st.columns(3)
        v["aliquota_imposto_pct"] = p1.number_input("Impostos sobre vendas (%)", 0.0, 40.0, proj["aliquota_imposto_pct"], 0.1, key=wk("aliq"),
                                                    help="Alíquota efetiva do Simples Nacional ou soma de PIS/COFINS/ISS/ICMS.")
        v["crescimento_mensal_ano1_pct"] = p2.number_input("Crescimento mensal das vendas no Ano 1 (%)", -10.0, 30.0,
                                                           proj["crescimento_mensal_ano1_pct"], 0.5, key=wk("gm"))
        v["crescimento_anual_pct"] = p3.number_input("Crescimento anual nos Anos 2 e 3 (%)", -50.0, 200.0,
                                                     proj["crescimento_anual_pct"], 1.0, key=wk("ga"))
        v["inflacao_custos_anual_pct"] = p1.number_input("Inflação dos custos fixos (% ao ano)", 0.0, 50.0,
                                                         proj["inflacao_custos_anual_pct"], 0.5, key=wk("inf"))
        v["tma_anual_pct"] = p2.number_input("TMA - taxa mínima de atratividade (% ao ano)", 0.0, 60.0, proj["tma_anual_pct"], 0.5,
                                             key=wk("tma"), help="Retorno mínimo que os sócios exigem (ex.: rendimento de uma aplicação segura + prêmio de risco).")
        v["pme_dias"] = p3.number_input("Prazo médio de estoques (dias)", 0.0, 365.0, proj["pme_dias"], 1.0, key=wk("pme"))
        v["reserva_caixa_extra"] = p1.number_input("Reserva financeira de segurança (R$)", 0.0, 1e8, proj["reserva_caixa_extra"], 1000.0, key=wk("res"))
        v["recuperar_giro"] = int(p2.checkbox("No VPL, recuperar o capital de giro no último ano", bool(proj["recuperar_giro"]), key=wk("rg")))
        with st.expander("🧮 Calcular a alíquota do Simples Nacional (anexos e fator R)"):
            calculadora_simples()
    auto_unica("projeto", v)
    salvo_txt()

    st.subheader("👥 Sócios")
    editor("socios", {"nome": T("Nome", required=True), "funcao": T("Função"), "contato": T("Contato"),
                      "perfil": T("Minicurrículo / atribuições", width="large"),
                      "capital_integralizado": moeda("Capital investido (R$)"),
                      "perda_aceitavel": moeda("Perda aceitável (R$)", help="Quanto este sócio aceita perder se o negócio não "
                                               "der certo, sem comprometer a vida pessoal (Sarasvathy, 2001).")}, "nome")
    st.caption(f"Capital próprio total: **{brl(d['socios']['capital_integralizado'].sum())}** · perda aceitável somada: "
               f"**{brl(d['socios']['perda_aceitavel'].sum())}** (usada em Monte Carlo e perda aceitável)")
    proxima("projeto")


# ==============================================================================
# CLIENTE E MERCADO
# ==============================================================================
def pg_segmentos():
    cabecalho("segmentos")
    t1, t2 = st.tabs(["🎯 Segmentos de clientes", "🏆 Priorização"])
    with t1:
        st.caption("Escreva até 4 a 6 grupos de clientes possíveis e dê notas de 1 a 10 para cada pergunta.")
        sg = editor("segmentos", {
            "nome": T("Segmento (máx. 40 caracteres)", required=True, max_chars=40),
            "descricao": T("Descrição do perfil", width="large"),
            "nota_tamanho": nota("É grande o suficiente?"), "nota_lucro": nota("Tem potencial de lucro?"),
            "nota_longevidade": nota("Não tende a desaparecer?"), "nota_concorrencia": nota("Tem poucos concorrentes?"),
            "nota_aderencia": nota("Seus produtos atendem esse segmento?")}, "nome")
    with t2:
        sg = d["segmentos"]
        if sg.empty:
            st.info("Cadastre os segmentos na primeira aba.")
        else:
            rk = sg.assign(Média=sg.iloc[:, 2:7].mean(axis=1)).sort_values("Média", ascending=False).reset_index(drop=True)
            for i, r in rk.iterrows():
                with st.container(border=True):
                    a, b = st.columns([5, 1])
                    a.markdown(f"### {i + 1}º  {r['nome']}\n{r['descricao']}")
                    b.metric("Média", num_br(r["Média"], 1))
            st.altair_chart(alt.Chart(rk).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("nome:N", sort="-x", title=None)),
                width="stretch")
    proxima("segmentos")


CAMPOS_PERSONA = {
    "Dados": [("nome", "Nome", "txt", 40), ("idade", "Idade (anos)", "int", 0), ("estado_civil", "Estado civil", "txt", 30),
              ("escolaridade", "Escolaridade", "txt", 40), ("profissao", "Profissão", "txt", 60),
              ("renda_mensal", "Renda mensal (R$)", "num", 0), ("cidade", "Cidade", "txt", 60),
              ("descricao", "Sobre a persona (máx. 380 caracteres)", "area", 380)],
    "Características": [("citacao", "Frase que representa a persona", "area", 250),
                        ("habitos", "Hábitos e comportamentos", "area", 500),
                        ("caracteristicas", "Características (#tags)", "txt", 200),
                        ("fluencia_digital", "Fluência digital (computador, smartphone, apps, compras on-line)", "txt", 200),
                        ("redes_sociais", "Redes sociais utilizadas", "txt", 150),
                        ("quem_influencia", "Quem influencia as decisões", "txt", 200)],
    "Rotina": [("rotina_manha", "Manhã", "area", 400), ("rotina_tarde", "Tarde", "area", 400), ("rotina_noite", "Noite", "area", 400)],
    "Insights": [("ganhos", "Ganhos esperados", "area", 500), ("desafios", "Dores e desafios", "area", 500),
                 ("duvidas_objecoes", "Dúvidas ou objeções", "area", 500),
                 ("aborrecimentos", "Aborrecimentos que devem ser evitados", "area", 500),
                 ("experiencia_ideal", "Experiência ideal e como o negócio ajuda", "area", 500)],
}


def pg_personas():
    cabecalho("personas")
    pers = ler_com_id("personas")
    c1, c2 = st.columns([4, 1.2])
    nomes = [f"{r['nome']}" for _, r in pers.iterrows()]
    if c2.button("➕ Nova persona", type="primary", width="stretch"):
        rid = novo_registro("personas", {"nome": f"Persona {len(pers) + 1}", "idade": 30})
        ss[wk("persona_sel")] = rid
        st.rerun()
    if pers.empty:
        st.info("Nenhuma persona cadastrada. Clique em **➕ Nova persona**.")
        proxima("personas")
        return
    ids = pers["id"].tolist()
    atual_id = ss.get(wk("persona_sel"), ids[0])
    atual_id = atual_id if atual_id in ids else ids[0]
    escolhido = c1.radio("Persona", ids, index=ids.index(atual_id), horizontal=True,
                         format_func=lambda i: "🧑 " + str(pers.loc[pers["id"] == i, "nome"].iloc[0]))
    ss[wk("persona_sel")] = escolhido
    row = pers[pers["id"] == escolhido].iloc[0].to_dict()

    st.markdown(f"<div style='text-align:center;font-size:4rem;line-height:1'>🧑</div>"
                f"<h3 style='text-align:center;margin-top:0'>{row['nome']}</h3>", unsafe_allow_html=True)
    novos = {}
    abas_p = st.tabs(list(CAMPOS_PERSONA))
    for aba, (grupo, campos) in zip(abas_p, CAMPOS_PERSONA.items()):
        with aba:
            cols = st.columns(3 if grupo == "Dados" else (3 if grupo == "Rotina" else 2))
            for i, (k, rot, tipo, mx) in enumerate(campos):
                alvo = cols[i % len(cols)] if tipo != "area" or grupo in ("Rotina", "Insights", "Características") else st
                val = row.get(k)
                chave_w = wk(f"pers_{escolhido}_{k}")
                if tipo == "int":
                    novos[k] = alvo.number_input(rot, 0, 120, int(val or 0), key=chave_w)
                elif tipo == "num":
                    novos[k] = alvo.number_input(rot, 0.0, 1e7, float(val or 0), 100.0, key=chave_w)
                elif tipo == "area":
                    novos[k] = alvo.text_area(rot, val or "", max_chars=mx, key=chave_w, height=110)
                else:
                    novos[k] = alvo.text_input(rot, val or "", max_chars=mx, key=chave_w)
    auto_registro("personas", escolhido, row, novos)
    a, b = st.columns([4, 1.2])
    with a:
        salvo_txt()
    if b.button("🗑️ Excluir esta persona", width="stretch"):
        apagar_registro("personas", escolhido)
        ss.pop(wk("persona_sel"), None)
        st.rerun()
    proxima("personas")


def pg_jornada():
    cabecalho("jornada")
    pers = d["personas"]["nome"].tolist()
    if not pers:
        st.info("Cadastre uma persona antes de montar a jornada.")
        if st.button("Ir para Personas"):
            ir("personas")
            st.rerun()
        return
    persona = st.radio("Persona", pers, horizontal=True, format_func=lambda n: "🧑 " + n, key=wk("jor_pers"))
    perguntas = {"Antes": "O que o cliente faz antes de entrar em contato com o seu produto ou serviço?",
                 "Durante": "Como é a compra, a entrega e o uso?",
                 "Depois": "O que acontece depois: avaliação, suporte e fidelização?"}
    fases = st.tabs(list(ETAPAS_FASE))
    for aba, fase in zip(fases, ETAPAS_FASE):
        with aba:
            st.markdown(f"**{perguntas[fase]}**")
            jc = d["jornada_cliente"]
            sub = jc[(jc["persona_nome"] == persona) & (jc["macro_fase"] == fase)]
            for etapa in ETAPAS_FASE[fase]:
                linhas = sub[sub["etapa"] == etapa]
                st.markdown(f"<div style='background:rgba(47,111,208,.08);padding:6px 12px;border-radius:6px;"
                            f"font-weight:600;margin-top:10px'>{etapa}</div>", unsafe_allow_html=True)
                if linhas.empty:
                    st.caption("— sem ações cadastradas nesta etapa —")
                for _, r in linhas.iterrows():
                    face = EMOCOES.get(r["estado_emocional"], "😐")
                    st.markdown(
                        f"<div class='pn-etapa'><span class='pn-face'>{face}</span> <b>{r['acoes_cliente']}</b><br>"
                        f"<small>💭 {r['pensamentos'] or '-'} · 📍 {r['pontos_contato'] or '-'}<br>"
                        f"💡 <i>Oportunidade:</i> {r['oportunidades'] or '-'}</small></div>", unsafe_allow_html=True)
            with st.expander(f"✏️ Editar as ações da fase “{fase}”", expanded=sub.empty):
                editor("jornada_cliente", {
                    "etapa": S("Etapa", options=ETAPAS_FASE[fase], required=True),
                    "acoes_cliente": T("Ação do cliente", width="medium", required=True),
                    "pensamentos": T("O que pensa"), "pontos_contato": T("Pontos de contato"),
                    "estado_emocional": S("Emoção", options=list(EMOCOES)),
                    "oportunidades": T("Oportunidade de melhoria", width="medium")},
                    "acoes_cliente", filtro={"persona_nome": persona, "macro_fase": fase})
    jc = d["jornada_cliente"]
    jp = jc[jc["persona_nome"] == persona]
    if not jp.empty:
        st.subheader("📉 Curva emocional da jornada")
        emo = {"Entusiasmado": 2, "Satisfeito": 1, "Neutro": 0, "Insatisfeito": -1, "Frustrado": -2}
        g = jp.assign(Emoção=jp["estado_emocional"].map(emo).fillna(0)).groupby("etapa", sort=False)["Emoção"].mean()
        g = g.reindex([e for e in ETAPAS if e in g.index]).reset_index()
        st.altair_chart(alt.Chart(g).mark_line(point=True, strokeWidth=3, color="#d6006e").encode(
            x=alt.X("etapa:N", sort=ETAPAS, title=None, axis=alt.Axis(labelAngle=0, labelLimit=120)),
            y=alt.Y("Emoção:Q", scale=alt.Scale(domain=[-2, 2]), title="😠  ←  emoção  →  🤩")).properties(height=260),
            width="stretch")
    proxima("jornada")


def pg_mercado():
    cabecalho("mercado")
    st.markdown("""
**Como estimar:** **TAM** é todo o mercado (todas as pessoas que poderiam comprar algo desse tipo na região).
**SAM** é a parte que o seu negócio consegue alcançar (pelo segmento escolhido, raio de entrega, canais).
**SOM** é a fatia que você realisticamente conquista (participação de mercado). Use dados do IBGE, Sebrae,
prefeitura, associações e pesquisa de campo, e anote as fontes.""")
    v = {}
    c1, c2 = st.columns([3, 1])
    v["tam_desc"] = c1.text_input("TAM - mercado total (quem é)", proj["tam_desc"], key=wk("tamd"))
    v["tam_clientes"] = c2.number_input("Nº de clientes potenciais (TAM)", 0.0, 1e9, float(proj["tam_clientes"]), 100.0, key=wk("tamc"))
    v["sam_desc"] = c1.text_input("SAM - mercado alcançável (quem é)", proj["sam_desc"], key=wk("samd"))
    v["sam_clientes"] = c2.number_input("Nº de clientes alcançáveis (SAM)", 0.0, 1e9, float(proj["sam_clientes"]), 100.0, key=wk("samc"))
    c3, c4 = st.columns(2)
    v["som_pct"] = c3.number_input("SOM - participação que pretende conquistar do SAM (%)", 0.0, 100.0, float(proj["som_pct"]), 0.5, key=wk("som"))
    v["gasto_anual_cliente"] = c4.number_input("Gasto médio anual de cada cliente com esse tipo de produto (R$)", 0.0, 1e7,
                                               float(proj["gasto_anual_cliente"]), 50.0, key=wk("gasto"))
    v["mercado_fontes"] = st.text_area("Fontes e premissas usadas", proj["mercado_fontes"], key=wk("mfont"))
    auto_unica("projeto", v)
    salvo_txt()
    mk = mercado(proj, R())
    if mk["preenchido"]:
        df = pd.DataFrame({"Nível": ["TAM", "SAM", "SOM"], "Clientes": [mk["tam_c"], mk["sam_c"], mk["som_c"]],
                           "Mercado (R$/ano)": [mk["tam"], mk["sam"], mk["som"]]})
        a, b = st.columns([1.3, 1])
        a.altair_chart(alt.Chart(df).mark_bar(color="#2f6fd0").encode(
            y=alt.Y("Nível:N", sort=["TAM", "SAM", "SOM"], title=None), x=alt.X("Mercado (R$/ano):Q", title="R$ por ano"),
            tooltip=["Nível", alt.Tooltip("Clientes:Q", format=",.0f"), alt.Tooltip("Mercado (R$/ano):Q", format=",.2f")]
        ).properties(height=200), width="stretch")
        b.metric("TAM", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
        b.metric("SAM", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
        b.metric("SOM", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
        st.subheader("A meta de vendas é realista?")
        p = mk["pct_som"]
        msg = (f"A receita prevista no Ano 1 ({brl(mk['receita_ano1'])}) corresponde a **{pct(p)} do SOM** "
               f"e a **{pct(mk['pct_sam'])} do SAM**.")
        if p > 1:
            st.error("🔴 " + msg + " A meta passa do que o negócio consegue conquistar: reveja as quantidades vendidas "
                     "(Produtos e materiais) ou a estimativa de mercado.")
        elif p > 0.8:
            st.warning("🟡 " + msg + " A meta é ambiciosa: depende de conquistar quase todo o mercado alcançável previsto.")
        else:
            st.success("🟢 " + msg + " A meta cabe no mercado estimado.")
    else:
        st.info("Preencha os quatro números para ver o tamanho do mercado e conferir a meta de vendas.")
    proxima("mercado")


# ==============================================================================
# PROBLEMA E SOLUÇÃO
# ==============================================================================
def pg_proposta():
    cabecalho("proposta")
    props = ler_com_id("propostas")
    pers = d["personas"]["nome"].tolist()
    c1, c2 = st.columns([4, 1.3])
    c1.caption("Uma proposta por persona: escolha os valores (etiquetas) e complete as frases.")
    if c2.button("➕ Nova proposta", type="primary", width="stretch"):
        novo_registro("propostas", {"persona_nome": pers[0] if pers else "", "valores": ""})
        st.rerun()
    if props.empty:
        st.info("Nenhuma proposta cadastrada. Clique em **➕ Nova proposta**.")
    for n, (_, r) in enumerate(props.iterrows(), start=1):
        row = r.to_dict()
        with st.expander(f"❤️ Proposta {n} — {row['persona_nome'] or 'sem persona'}:  {row['valores'][:90]}",
                         expanded=len(props) == 1):
            novos = {}
            opc_p = pers + ([row["persona_nome"]] if row["persona_nome"] and row["persona_nome"] not in pers else [])
            novos["persona_nome"] = st.selectbox("Persona", opc_p or [""], key=wk(f"pp_{row['id']}"),
                                                 index=opc_p.index(row["persona_nome"]) if row["persona_nome"] in opc_p else 0)
            atuais = [x.strip() for x in str(row["valores"]).split(";") if x.strip()]
            novos["valores"] = "; ".join(st.multiselect("Valores entregues (escolha ou digite novos)", opcoes("valor_proposta", atuais),
                                                        atuais, key=wk(f"pv_{row['id']}"), accept_new_options=True))
            a, b = st.columns(2)
            novos["que_desejam"] = a.text_area("Que desejam...", row["que_desejam"], key=wk(f"pq_{row['id']}"))
            novos["por_meio"] = b.text_area("Por meio de...", row["por_meio"], key=wk(f"pm_{row['id']}"))
            novos["diferentemente"] = a.text_area("Diferentemente de...", row["diferentemente"], key=wk(f"pd_{row['id']}"))
            novos["oferecem"] = b.text_area("Os nossos produtos ou serviços oferecem...", row["oferecem"], key=wk(f"po_{row['id']}"))
            auto_registro("propostas", row["id"], row, novos)
            st.info(f"**{proj['nome_negocio']}** ajuda **{novos['persona_nome']}**, que deseja **{novos['que_desejam']}**, "
                    f"por meio de **{novos['por_meio']}**. Diferentemente de **{novos['diferentemente']}**, "
                    f"oferecemos **{novos['oferecem']}**.")
            if st.button("🗑️ Excluir esta proposta", key=wk(f"pdel_{row['id']}")):
                apagar_registro("propostas", row["id"])
                st.rerun()
    salvo_txt()
    proxima("proposta")


def pivo_concorrencia():
    empresas = ["Nossa Empresa"] + [n for n in d["concorrentes"]["nome"] if str(n).strip()]
    notas = d["concorrencia_notas"]
    atributos = list(dict.fromkeys(notas["atributo"]))
    piv = pd.DataFrame({"Atributo": pd.Series(atributos, dtype=object)})
    for emp in empresas:
        mapa = notas[notas["empresa"] == emp].set_index("atributo")["nota"].to_dict()
        piv[emp] = pd.Series([mapa.get(a, None) for a in atributos], dtype=float)
    return piv, empresas


def pg_concorrencia():
    cabecalho("concorrencia")
    st.subheader("1. Concorrentes")
    editor("concorrentes", {"nome": T("Nome do concorrente", required=True),
                            "observacoes": T("Observações (onde fica, preço, pontos fortes)", width="large")}, "nome")
    piv, empresas = pivo_concorrencia()
    st.subheader("2. Notas por atributo (0 a 10)")

    def _norm(x):
        x = x.copy()
        x["Atributo"] = x["Atributo"].astype(object).where(x["Atributo"].notna(), "")
        for e in empresas:
            x[e] = pd.to_numeric(x[e], errors="coerce") if e in x else float("nan")
        return x[["Atributo"] + empresas]

    def _salvar(x):
        x = x[x["Atributo"].astype(str).str.strip() != ""]
        longo = x.melt(id_vars="Atributo", var_name="empresa", value_name="nota").rename(columns={"Atributo": "atributo"})
        salvar(conn, "concorrencia_notas", longo, "atributo")
        d["concorrencia_notas"] = ler(conn, "concorrencia_notas")

    cfg = {"Atributo": T("Atributo avaliado", required=True)}
    cfg.update({e: N(e, min_value=0, max_value=10, step=0.5) for e in empresas})
    pe = tabela_editavel(piv, cfg, "concorrencia_notas", "|".join(empresas), normalizar=_norm, salvar_fn=_salvar)
    pe = pe[pe["Atributo"].astype(str).str.strip() != ""]
    num = pe.copy()
    for e in empresas:
        num[e] = pd.to_numeric(num[e], errors="coerce").fillna(0)
    if not num.empty:
        st.subheader("3. Resultado")
        medias = num[empresas].mean()
        cols = st.columns(min(len(empresas), 5))
        for i, e in enumerate(empresas[:5]):
            cols[i].metric(e, num_br(medias[e], 1))
        longo = num.melt(id_vars="Atributo", var_name="Empresa", value_name="Nota")
        st.altair_chart(alt.Chart(longo).mark_line(point=True).encode(
            x=alt.X("Atributo:N", title=None, axis=alt.Axis(labelAngle=-35, labelLimit=160)),
            y=alt.Y("Nota:Q", scale=alt.Scale(domain=[0, 10])), color="Empresa:N").properties(height=340), width="stretch")
        if len(empresas) > 1:
            idx = num.set_index("Atributo")
            outros = idx[empresas[1:]].max(axis=1)
            st.success("**Forças frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] > outros]) or "nenhuma"))
            st.warning("**Fraquezas frente à concorrência:** " + (", ".join(idx.index[idx["Nossa Empresa"] < outros]) or "nenhuma"))
            dest = idx["Nossa Empresa"] - outros
            st.info(f"**Destaque do negócio:** {dest.idxmax()} (vantagem de {num_br(dest.max(), 1)} ponto(s) sobre o melhor concorrente).")
    proxima("concorrencia")


# ==============================================================================
# ESTRATÉGIA
# ==============================================================================
def _fatores(ambiente, naturezas, dims, rot_acao):
    editor("swot", {
        "dimensao": S("Dimensão", options=dims),
        "fator": T("Fator", required=True),
        "natureza": S("Classificação", options=naturezas, required=True),
        "impacto": S("Impacto", options=["Alto", "Baixo"], required=True),
        "diagnostico": T("Diagnóstico (por que é assim no seu negócio)", width="medium"),
        "plano_acao": T(rot_acao, width="medium"),
        "priorizado": CK("⭐ Priorizar")}, "fator", filtro={"ambiente": ambiente})


def pg_forcas():
    cabecalho("forcas")
    st.caption("Exemplos de fatores: qualificação da equipe, presença digital, controle de qualidade, controles financeiros, "
               "localização, relação com fornecedores, reputação da marca.")
    dims = opcoes("dimensao_interna", d["swot"].loc[d["swot"]["ambiente"] == "Interno", "dimensao"].tolist())
    _fatores("Interno", ["Força", "Fraqueza"], dims, "Como potencializar (força) ou corrigir (fraqueza)")
    proxima("forcas")


def pg_oportunidades():
    cabecalho("oportunidades")
    st.caption("Exemplos: preocupação com saúde, novas tecnologias de pagamento, inflação, renda dos consumidores, "
               "fiscalização, novos concorrentes, mudanças demográficas.")
    dims = opcoes("dimensao_externa", d["swot"].loc[d["swot"]["ambiente"] == "Externo", "dimensao"].tolist())
    _fatores("Externo", ["Oportunidade", "Ameaça"], dims, "Como aproveitar (oportunidade) ou neutralizar (ameaça)")
    proxima("oportunidades")


def html_swot(sw: pd.DataFrame) -> str:
    def lista(nat):
        itens = sw[sw["natureza"] == nat]
        if itens.empty:
            return "<i>nenhum fator</i>"
        return "<br>".join(("⭐ " if bool(r["priorizado"]) else "• ") + f"{r['fator']} <small>({r['impacto']})</small>"
                           for _, r in itens.iterrows())
    cel = "padding:14px 16px;border-radius:4px;min-height:200px;font-size:14px;line-height:1.55;"
    lat = ("writing-mode:vertical-rl;transform:rotate(180deg);font-weight:700;letter-spacing:2px;"
           "display:flex;align-items:center;justify-content:center;")
    return f"""
    <div style="display:grid;grid-template-columns:44px 1fr 1fr;gap:6px;font-family:sans-serif;">
      <div></div>
      <div style="background:#26548c;color:#fff;text-align:center;font-weight:700;padding:10px;">FATORES POSITIVOS</div>
      <div style="background:#d6006e;color:#fff;text-align:center;font-weight:700;padding:10px;">FATORES NEGATIVOS</div>
      <div style="background:#26548c;color:#fff;{lat}">INTERNOS</div>
      <div style="background:#26548c;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">💪 Forças</div>{lista("Força")}</div>
      <div style="background:#d6006e;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⛓️ Fraquezas</div>{lista("Fraqueza")}</div>
      <div style="background:#c8ccd2;color:#28323c;{lat}">EXTERNOS</div>
      <div style="background:#e2e6eb;color:#28323c;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⬆️ Oportunidades</div>{lista("Oportunidade")}</div>
      <div style="background:#5c6672;color:#fff;{cel}"><div style="font-size:20px;font-weight:700;margin-bottom:6px;">⛈️ Ameaças</div>{lista("Ameaça")}</div>
    </div>"""


def pg_swot():
    cabecalho("swot")
    sw = d["swot"]
    c1, c2 = st.columns(2)
    if c1.button("✏️ Editar forças e fraquezas", width="stretch"):
        ir("forcas")
        st.rerun()
    if c2.button("✏️ Editar oportunidades e ameaças", width="stretch"):
        ir("oportunidades")
        st.rerun()
    st.markdown(html_swot(sw), unsafe_allow_html=True)
    st.caption("⭐ fator priorizado · entre parênteses, o impacto")
    diag = pontuacao_swot(sw)
    st.subheader("Matriz de estratégias")
    q = diag["quadrantes"]
    g1, g2 = st.columns(2)
    for col, k in zip([g1, g2, g1, g2], ["FO", "FrO", "FA", "FrA"]):
        nome_q, comb, _ = QUADRANTES[k]
        destaque = "border:3px solid #2f6fd0;" if k == diag["predominante"] else "border:1px solid #ccc;"
        col.markdown(f"<div style='{destaque}padding:12px;border-radius:8px;margin-bottom:8px;text-align:center;'>"
                     f"<div style='font-size:28px;font-weight:700;'>{q[k]}</div><b>{comb}</b><br>{nome_q}</div>",
                     unsafe_allow_html=True)
    if diag["predominante"]:
        nome_q, comb, txt = QUADRANTES[diag["predominante"]]
        st.success(f"**Situação do negócio: {nome_q}.** {txt}")
    pri = sw[sw["priorizado"].astype(bool)]
    if not pri.empty:
        st.subheader("⭐ Ações estratégicas dos fatores priorizados")
        verbo = {"Força": "Como potencializar", "Fraqueza": "Como corrigir", "Oportunidade": "Como aproveitar", "Ameaça": "Como neutralizar"}
        for _, r in pri.iterrows():
            st.markdown(f"**{r['natureza']} · {r['fator']}** ({r['impacto']} impacto)  \n{r['diagnostico']}  \n"
                        f"➡️ *{verbo.get(r['natureza'], 'Ação')}:* {r['plano_acao']}")
    st.caption("Pontuação: impacto alto = 2 pontos; impacto baixo = 1 ponto.")
    proxima("swot")


# ==============================================================================
# MARKETING E OPERAÇÃO
# ==============================================================================
def pg_canais():
    cabecalho("canais")
    t1, t2 = st.tabs(["🛒 Canais de aquisição", "🔻 Funil de vendas"])
    with t1:
        st.caption("Dê notas de 0 a 10. Os três canais com maior média são os **canais de tração**.")
        cn = editor("canais", {"canal": S("Canal", options=opcoes("canal", d["canais"]["canal"].tolist()), required=True),
                               "atratividade": nota("Atratividade"), "alcance": nota("Alcance esperado"),
                               "cliente": nota("Custo por cliente (10 = barato)"), "relacionamento": nota("Relacionamento"),
                               "viabilidade": nota("Viabilidade")}, "canal")
        if not cn.empty:
            cn2 = cn.assign(Média=cn.iloc[:, 1:6].mean(axis=1)).sort_values("Média", ascending=False)
            st.success("**Canais de tração:** " + ", ".join(cn2["canal"].head(3)))
            st.altair_chart(alt.Chart(cn2).mark_bar(color="#2f6fd0").encode(
                x=alt.X("Média:Q", scale=alt.Scale(domain=[0, 10])), y=alt.Y("canal:N", sort="-x", title=None)), width="stretch")
    with t2:
        fu = editor("funil_vendas", {
            "etapa": S("Etapa", options=["Topo do Funil", "Meio do Funil", "Fundo do Funil"], required=True),
            "acao": T("Ação de venda", width="large"), "orcamento": moeda("Orçamento (R$)"),
            "alcancados": N("Pessoas alcançadas", min_value=0, format="%d"),
            "convertidos": N("Pessoas que avançaram / compraram", min_value=0, format="%d")}, "etapa")
        if not fu.empty:
            st.dataframe(tabela_funil(fu), hide_index=True, width="stretch")
            fundo = fu[fu["etapa"] == "Fundo do Funil"]["convertidos"].sum()
            m1, m2, m3 = st.columns(3)
            m1.metric("Orçamento total", brl(fu["orcamento"].sum()))
            m2.metric("Clientes conquistados", num_br(fundo, 0))
            m3.metric("CAC - custo de aquisição por cliente", brl(div(fu["orcamento"].sum(), fundo)))
            agg = fu.groupby("etapa", sort=False)[["alcancados", "convertidos"]].sum().reindex(
                ["Topo do Funil", "Meio do Funil", "Fundo do Funil"]).dropna().reset_index()
            st.altair_chart(alt.Chart(agg).mark_bar(color="#d6006e").encode(
                y=alt.Y("etapa:N", sort=None, title=None), x=alt.X("alcancados:Q", title="Pessoas alcançadas")), width="stretch")
    proxima("canais")


def tabela_funil(fu: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame({
        "Etapa": fu["etapa"], "Ação": fu["acao"], "Orçamento": fu["orcamento"].map(brl),
        "Alcançados": fu["alcancados"].map(lambda v: num_br(v, 0)),
        "Convertidos": fu["convertidos"].map(lambda v: num_br(v, 0)),
        "Conversão": [pct(div(c, a)) for c, a in zip(fu["convertidos"], fu["alcancados"])],
        "Custo por conversão": [brl(div(o, c)) for o, c in zip(fu["orcamento"], fu["convertidos"])]})


def dot_organograma() -> str:
    def q(s):
        return '"' + str(s).replace('"', "'").replace("\n", "\\n") + '"'
    linhas = ['digraph G {', 'rankdir=TB; nodesep=0.35; ranksep=0.5;',
              'node [shape=box, style="rounded,filled", fillcolor="#e8f0fb", fontname="Helvetica", fontsize=13, margin="0.25,0.12"];',
              'edge [color="#7a8aa0"];', f'raiz [label={q(proj["nome_negocio"] or "Empresa")}, fillcolor="#2f6fd0", fontcolor="white", fontsize=16];']
    soc = d["socios"]
    if not soc.empty:
        txt = "Sócios\n" + "\n".join(f"{r['nome']} ({r['funcao'] or 'sócio'})" for _, r in soc.iterrows())
        linhas.append(f'socios [label={q(txt)}, fillcolor="#cfe0f7"];')
        linhas.append("socios -> raiz [dir=none, style=dashed];")
        linhas.append("{rank=min; socios}")
    pes = d["pessoal"]
    for j, area in enumerate(dict.fromkeys(pes["area"].tolist())):
        if not area:
            continue
        sub = pes[pes["area"] == area]
        cargos = "\n".join(f"{r['cargo']} ({num_br(r['num_empregados'], 0)})" for _, r in sub.iterrows())
        total = num_br(sub["num_empregados"].sum(), 0)
        linhas.append(f'a{j} [label={q(area + " - " + total + " pessoa(s)")}, fillcolor="#fde3ef", fontsize=14];')
        linhas.append(f'c{j} [label={q(cargos)}, fillcolor="#ffffff"];')
        linhas.append(f'raiz -> a{j}; a{j} -> c{j};')
    linhas.append("}")
    return "\n".join(linhas)


def pg_operacional():
    cabecalho("operacional")
    t1, t2, t3, t4 = st.tabs(["📍 Localização", "🏗️ Layout e processos", "⚙️ Capacidade x demanda", "🗂️ Organograma"])
    v = {}
    with t1:
        c1, c2 = st.columns([3, 1])
        v["loc_endereco"] = c1.text_input("Endereço / região", proj["loc_endereco"], key=wk("le"))
        v["loc_area_m2"] = c2.number_input("Área (m²)", 0.0, 1e6, float(proj["loc_area_m2"]), 5.0, key=wk("la"))
        v["loc_justificativa"] = st.text_area("Por que este ponto? (fluxo, acesso, concorrentes, custo, segurança)",
                                              proj["loc_justificativa"], key=wk("lj"))
        v["loc_fluxo"] = st.text_area("Fluxo de pessoas e horários de pico", proj["loc_fluxo"], key=wk("lf"))
        v["op_horario"] = st.text_input("Horário de funcionamento", proj["op_horario"], key=wk("oh"))
    with t2:
        v["op_layout"] = st.text_area("Layout (arranjo físico): como os espaços serão organizados", proj["op_layout"], height=130, key=wk("ol"))
        v["op_processos"] = st.text_area("Processos operacionais: passo a passo da compra à entrega (use →)",
                                         proj["op_processos"], height=130, key=wk("op"))
        etapas = [e.strip() for e in re.split(r"->|→|\n", proj["op_processos"] or "") if e.strip()]
        if len(etapas) >= 2:
            dot = "digraph P { rankdir=LR; node [shape=box, style=\"rounded,filled\", fillcolor=\"#e8f0fb\", fontname=Helvetica, fontsize=10];\n"
            dot += " -> ".join(f'"{e[:40]}"' for e in etapas) + ";}"
            st.graphviz_chart(dot, width="stretch")
    with t3:
        v["op_capacidade_txt"] = st.text_area("Como a capacidade foi estimada", proj["op_capacidade_txt"], key=wk("oc"))
        alvos = ["Todos os produtos", "Fabricação Própria", "Serviço", "Revenda"] + d["catalogo"]["descricao"].tolist()
        editor("capacidade", {"recurso": T("Recurso (equipe, máquina, espaço)", required=True, width="medium"),
                              "produto_ou_grupo": S("Atende", options=alvos, required=True),
                              "capacidade_mensal": N("Capacidade por mês (unidades)", min_value=0.0),
                              "observacao": T("Observação", width="medium")}, "recurso")
        cd = capacidade_demanda(d, R())
        if not cd.empty:
            show = cd.copy()
            for c in ("Uso Mês 1", "Uso Mês 12"):
                show[c] = show[c].map(lambda x: ("🔴 " if x > 1 else "🟡 " if x > 0.85 else "🟢 ") + pct(x))
            st.dataframe(show, hide_index=True, width="stretch")
            g = cd.melt(id_vars="Recurso", value_vars=["Uso Mês 1", "Uso Mês 12"], var_name="Momento", value_name="Uso")
            st.altair_chart((alt.Chart(g).mark_bar().encode(x=alt.X("Uso:Q", axis=alt.Axis(format="%")), y=alt.Y("Recurso:N", title=None),
                                                           color="Momento:N", yOffset="Momento:N")
                             + alt.Chart(pd.DataFrame({"x": [1]})).mark_rule(color="red", strokeDash=[5, 3]).encode(x="x:Q")
                             ).properties(height=60 * len(cd) + 40), width="stretch")
            st.caption("Linha vermelha = 100% da capacidade. Acima dela, a operação não dá conta da demanda prevista.")
    with t4:
        st.caption("Gerado a partir dos sócios (aba Informações) e da mão de obra (aba Custos), agrupada por área.")
        try:
            st.graphviz_chart(dot_organograma(), width="stretch")
        except Exception as e:
            st.info(f"Não foi possível desenhar o organograma: {e}")
    auto_unica("projeto", v)
    salvo_txt()
    proxima("operacional")


def pg_fornecedores():
    cabecalho("fornecedores")
    fo = editor("fornecedores", {"nome": T("Fornecedor", required=True), "itens": T("O que fornece", width="medium"),
                                 "prazo_pagamento_dias": N("Prazo de pagamento (dias)", min_value=0, step=1),
                                 "prazo_entrega_dias": N("Prazo de entrega (dias)", min_value=0, step=1),
                                 "condicoes": T("Condições (pedido mínimo, frete, desconto)", width="medium"),
                                 "critico": CK("Crítico?")}, "nome")
    if not fo.empty:
        a, b, c = st.columns(3)
        a.metric("Fornecedores cadastrados", len(fo))
        b.metric("Prazo médio de pagamento (fornecedores)", f"{num_br(fo['prazo_pagamento_dias'].mean(), 0)} dias")
        c.metric("PMP usado no capital de giro", f"{num_br(R()['pmp'], 1)} dias")
        crit = fo[fo["critico"].astype(bool)]
        if not crit.empty:
            st.warning("**Fornecedores críticos** (ter alternativa homologada): " + ", ".join(crit["nome"]))
    proxima("fornecedores")

# ##############################################################################
# PARTE: pnbox_ui_paginas2.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: FERRAMENTAS FINANCEIRAS
# ==============================================================================

def tabela_comp(df):
    return df.assign(Valor=df["Valor"].map(brl), **{"%": df["%"].map(pct)})


def rosca(df, titulo, chave_g, tabela=True):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(
        theta=alt.Theta("Valor:Q", stack=True),
        color=alt.Color("Item:N", sort=list(df["Item"]),
                        legend=alt.Legend(title=None, orient="bottom", columns=2, labelLimit=260)),
        tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    arco = base.mark_arc(innerRadius=55, outerRadius=105)
    rot = base.transform_filter(alt.datum["%"] >= 0.04).mark_text(radius=127, fontSize=12).encode(
        text=alt.Text("%:Q", format=".1%"))
    st.altair_chart((arco + rot).properties(height=360), width="stretch", key=chave_g)
    if tabela:
        st.dataframe(tabela_comp(df), hide_index=True, width="stretch")


def barras(df, titulo, chave_g, cor="#2f6fd0"):
    st.markdown(f"**{titulo}**")
    if df.empty:
        st.info("Sem dados para este gráfico.")
        return
    base = alt.Chart(df).encode(y=alt.Y("Item:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260)),
                                x=alt.X("Valor:Q", title="R$"),
                                tooltip=["Item", alt.Tooltip("Valor:Q", format=",.2f"), alt.Tooltip("%:Q", format=".1%")])
    st.altair_chart((base.mark_bar(color=cor) + base.mark_text(align="left", dx=3).encode(text=alt.Text("%:Q", format=".1%")))
                    .properties(height=28 * len(df) + 40), width="stretch", key=chave_g)


# ==============================================================================
# INVESTIMENTOS E FONTES
# ==============================================================================
def pg_investimentos():
    cabecalho("investimentos")
    t1, t2, t3, t4 = st.tabs(["🏗️ Investimentos fixos", "📋 Pré-operacionais", "🔄 Capital de giro", "🏦 Fontes de recursos"])
    with t1:
        st.caption("Máquinas, equipamentos, móveis, veículos, instalações. Vida útil 0 = não deprecia (ex.: terreno).")
        invf = editor("inv_fixos", {
            "descricao": T("Descrição", required=True, width="large"), "quantidade": N("Qtd.", min_value=0.0),
            "valor_unitario": moeda("Valor unitário (R$)"),
            "vida_util_anos": N("Vida útil (anos)", min_value=0, max_value=50, step=1),
            "valor_residual_pct": N("Valor residual (%)", min_value=0.0, max_value=100.0)}, "descricao")
        r = R()
        a, b = st.columns(2)
        a.metric("Total dos investimentos fixos", brl(r["inv_fixo"]))
        b.metric("Depreciação mensal", brl(r["dep_m"]))
    with t2:
        st.caption("Gastos antes de abrir: reformas, legalização, marca, projeto, treinamento, marketing de inauguração.")
        editor("inv_pre", {"descricao": T("Descrição", required=True, width="large"), "valor": moeda("Valor (R$)")}, "descricao")
        st.metric("Total pré-operacional", brl(R()["inv_pre"]))
    with t3:
        st.caption("Recebimento: como os clientes pagam (0 dias = à vista). Pagamento: como a empresa paga os fornecedores. "
                   "Em cada tipo, os percentuais devem somar 100%.")
        editor("capital_giro_prazos", {
            "tipo": S("Tipo", options=["Recebimento", "Pagamento"], required=True),
            "prazo_dias": N("Prazo (dias)", min_value=0, max_value=360, step=1),
            "percentual": N("% das vendas / compras", min_value=0.0, max_value=100.0)}, "tipo")
        r = R()
        c = st.columns(4)
        c[0].metric("PMR (recebimento)", f"{num_br(r['pmr'], 1)} dias")
        c[1].metric("PMP (pagamento)", f"{num_br(r['pmp'], 1)} dias")
        c[2].metric("PME (estoques)", f"{num_br(r['pme'], 0)} dias")
        c[3].metric("Ciclo financeiro", f"{num_br(r['ciclo_fin'], 1)} dias")
        c = st.columns(3)
        c[0].metric("Necessidade de capital de giro", brl(r["ncg"]))
        c[1].metric("Reserva de segurança", brl(r["reserva"]), help="Altere em Informações e sócios → Premissas.")
        c[2].metric("Capital de giro total", brl(r["capital_giro"]))
        st.metric("Estoque inicial (calculado em Produtos e materiais)", brl(r["inv_estoque"]))
    with t4:
        st.caption("O capital próprio vem dos sócios (Informações e sócios). Aqui você cadastra empréstimos e financiamentos.")
        editor("emprestimos", {"descricao": T("Descrição (banco, linha de crédito)", required=True, width="large"),
                               "valor": moeda("Valor (R$)"), "taxa_anual_pct": N("Juros (% ao ano)", min_value=0.0),
                               "prazo_meses": N("Prazo (meses)", min_value=1, step=1),
                               "carencia_meses": N("Carência (meses)", min_value=0, step=1),
                               "sistema": S("Sistema", options=["Price", "SAC"])}, "descricao")
        if st.button("🏦 Ver parcelas e comparar propostas de crédito", key="inv_fin"):
            ir("financiamento")
            st.rerun()
        r = R()
        it = r["investimento_total"]
        a, b = st.columns(2)
        a.dataframe(pd.DataFrame({
            "Aplicação": ["Investimentos fixos", "Pré-operacionais", "Estoque inicial", "Capital de giro", "Total"],
            "Valor": [brl(x) for x in (r["inv_fixo"], r["inv_pre"], r["inv_estoque"], r["capital_giro"], it)],
            "%": [pct(div(x, it)) for x in (r["inv_fixo"], r["inv_pre"], r["inv_estoque"], r["capital_giro"], it)]}),
            hide_index=True, width="stretch")
        b.dataframe(pd.DataFrame({
            "Fonte": ["Capital próprio (sócios)", "Capital de terceiros", "Total das fontes"],
            "Valor": [brl(r["capital_proprio"]), brl(r["capital_terceiros"]), brl(r["fontes"])]}), hide_index=True, width="stretch")
        dif = r["fontes"] - it
        if abs(dif) <= 1:
            st.success("🟢 As fontes de recursos cobrem exatamente o investimento total.")
        else:
            falta = "faltam" if dif < 0 else "sobram"
            st.error(f"🔴 Diferença de **{brl(dif)}**: {falta} {brl(abs(dif))} para fechar o investimento total de {brl(it)}.")
            st.markdown("**Como corrigir:** aumente/reduza o capital dos sócios ou o valor dos empréstimos.")
            if st.button("Ajustar o capital dos sócios →"):
                ir("projeto", f"Ajuste o capital dos sócios em {brl(-dif)} no total para igualar o investimento de {brl(it)}.")
                st.rerun()
    proxima("investimentos")


# ==============================================================================
# PRODUTOS, MATERIAIS E FICHA TÉCNICA
# ==============================================================================
def pg_produtos():
    cabecalho("produtos")
    t1, t2, t3, t4 = st.tabs(["① Materiais e insumos", "② Produtos e serviços", "③ Ficha técnica", "④ Resumo de custos"])
    with t1:
        st.caption("Matérias-primas, embalagens e insumos. A quantidade inicial forma o estoque inicial. "
                   "Falta uma unidade ou categoria? Acrescente em ⚙️ Configurações e listas, ou digite direto.")
        mat = editor("materiais", {
            "nome": T("Material / insumo", required=True, width="large"),
            "categoria": S("Categoria", options=opcoes("categoria_material", d["materiais"]["categoria"].tolist()), required=True),
            "unidade": S("Unidade", options=opcoes("unidade", d["materiais"]["unidade"].tolist()), required=True),
            "valor_unitario": moeda("Valor por unidade (R$)"),
            "qtd_estoque_inicial": N("Qtd. no estoque inicial", min_value=0.0)}, "nome")
        _nova_lista_rapida(["unidade", "categoria_material"], "mat")
    with t2:
        st.caption("**Revenda:** informe o valor de aquisição. **Fabricação própria e serviços:** monte a ficha técnica "
                   "(aba ③); se não houver ficha, vale o custo informado.")
        editor("catalogo", {
            "categoria": S("Categoria", options=CATEGORIAS, required=True),
            "descricao": T("Produto / serviço", required=True, width="large"),
            "preco_venda": moeda("Preço de venda (R$)"),
            "custo_aquisicao": moeda("Valor de aquisição / custo informado (R$)"),
            "qtd_mensal": N("Qtd. vendida no Mês 1", min_value=0, step=1),
            "qtd_estoque_inicial": N("Estoque inicial (só revenda)", min_value=0, step=1),
            "horas_unit": N("Horas de trabalho por unidade (serviços)", min_value=0.0, step=0.5,
                            help="Quanto tempo da equipe cada serviço consome. Usado em 🧑‍💻 Serviços e custo da hora.")},
            "descricao")
        st.caption("Vende planos com pagamento recorrente (mensal, trimestral, anual)? Cadastre-os em "
                   "**💳 Assinaturas e receita recorrente**, não aqui: lá o programa considera renovações e cancelamentos.")
    with t3:
        cat, mat = d["catalogo"], d["materiais"]
        fabricaveis = cat[cat["categoria"] != "Revenda"]["descricao"].tolist()
        if not fabricaveis:
            st.info("Cadastre produtos de fabricação própria ou serviços na aba ②.")
        elif mat.empty:
            st.info("Cadastre os materiais na aba ①.")
        else:
            prod = st.selectbox("Produto / serviço", fabricaveis, key=wk("ft_prod"))
            ft_all = d["ficha_tecnica"]
            ft = ft_all[ft_all["produto"] == prod][["material", "quantidade"]].reset_index(drop=True)

            def _norm(x):
                x = x[["material", "quantidade"]].copy()
                x["material"] = x["material"].astype(object).where(x["material"].notna(), "")
                x["quantidade"] = pd.to_numeric(x["quantidade"], errors="coerce").fillna(0.0)
                return x

            def _salvar(x):
                x = x[x["material"].astype(str).str.strip() != ""]
                novo = pd.concat([d["ficha_tecnica"][d["ficha_tecnica"]["produto"] != prod], x.assign(produto=prod)],
                                 ignore_index=True)
                salvar(conn, "ficha_tecnica", novo, "material")
                d["ficha_tecnica"] = ler(conn, "ficha_tecnica")

            ft_ed = tabela_editavel(ft, {"material": S("Material", options=mat["nome"].tolist(), required=True),
                                         "quantidade": N("Quantidade (na unidade do material)", min_value=0.0, format="%.3f")},
                                    "ficha_tecnica", prod, normalizar=_norm, salvar_fn=_salvar)
            ft_ed = ft_ed[ft_ed["material"].astype(str).str.strip() != ""]
            precos = dict(zip(mat["nome"], mat["valor_unitario"]))
            unid = dict(zip(mat["nome"], mat["unidade"]))
            catm = dict(zip(mat["nome"], mat["categoria"]))
            calc = ft_ed.assign(Unidade=ft_ed["material"].map(unid), Categoria=ft_ed["material"].map(catm),
                                **{"Valor unit.": ft_ed["material"].map(precos).fillna(0)})
            calc["Custo"] = calc["quantidade"] * calc["Valor unit."]
            custo_p = float(calc["Custo"].sum())
            preco_p = float(cat.loc[cat["descricao"] == prod, "preco_venda"].iloc[0])
            k = st.columns(4)
            k[0].metric("Custo unitário", brl(custo_p))
            k[1].metric("Preço de venda", brl(preco_p))
            k[2].metric("Margem bruta", pct(div(preco_p - custo_p, preco_p)))
            k[3].metric("Markup", f"{num_br(div(preco_p, custo_p))}x" if custo_p else "n.a.")
            if not calc.empty:
                st.dataframe(calc.assign(**{"Valor unit.": calc["Valor unit."].map(brl), "Custo": calc["Custo"].map(brl)}),
                             hide_index=True, width="stretch")
                por_cat = calc.groupby("Categoria")["Custo"].sum()
                st.caption("Composição do custo: " + " · ".join(f"{c}: {pct(div(v, custo_p))}" for c, v in por_cat.items()))
    with t4:
        cu = R()["cat"]
        if not cu.empty:
            st.dataframe(pd.DataFrame({
                "Categoria": cu["categoria"], "Produto / serviço": cu["descricao"], "Preço": cu["preco_venda"].map(brl),
                "Custo unit.": cu["custo_unit"].map(brl), "Origem do custo": cu["origem_custo"],
                "Margem": cu["margem_pct"].map(pct), "Markup": cu["markup"].map(lambda v: f"{num_br(v)}x" if v == v else "n.a."),
                "Qtd./mês": cu["qtd_mensal"].map(lambda v: num_br(v, 0)), "Receita/mês": cu["receita_mensal"].map(brl),
                "Custo/mês": cu["custo_mensal"].map(brl)}), hide_index=True, width="stretch")
            r = R()
            a, b, c = st.columns(3)
            a.metric("Receita do Mês 1", brl(r["receita_m1"]))
            b.metric("Custo direto do Mês 1", brl(r["cdir_m1"]))
            c.metric("Estoque inicial (materiais + revenda)", brl(r["inv_estoque"]))
    proxima("produtos")


def _nova_lista_rapida(tipos, sufixo):
    with st.expander("➕ Acrescentar uma opção às listas (unidade, categoria...)"):
        a, b, c = st.columns([1.5, 2, 1])
        tipo = a.selectbox("Lista", tipos, format_func=lambda t: NOMES_LISTAS[t], key=wk(f"lt_{sufixo}"))
        valor = b.text_input("Nova opção", key=wk(f"lv_{sufixo}"))
        if c.button("Acrescentar", key=wk(f"lb_{sufixo}"), width="stretch") and valor.strip():
            novo_registro("listas", {"tipo": tipo, "valor": valor.strip()})
            d["listas"] = ler(conn, "listas")
            for k in list(ss.snap):
                if k.startswith(f"{PID}_materiais") or k.startswith(f"{PID}_catalogo"):
                    ss.snap.pop(k)
            st.toast(f"“{valor}” acrescentado.")
            st.rerun()


# ==============================================================================
# CUSTOS E DESPESAS
# ==============================================================================
def pg_custos():
    cabecalho("custos")
    with st.expander("📚 Entenda a classificação dos custos", expanded=False):
        st.markdown("""
| Tipo | O que é | Exemplos | Onde entra |
|---|---|---|---|
| **Custo direto (variável)** | Vai direto no produto e varia com a quantidade | matéria-prima, embalagem, mercadoria de revenda | Ficha técnica / valor de aquisição |
| **Mão de obra direta (MOD)** | Pessoas que produzem ou prestam o serviço | cozinheiros, auxiliares de produção | Mão de obra com área "Produção / operação" |
| **Custo indireto de produção (CIF)** | Gasto da produção que não dá para medir por unidade | energia da cozinha, manutenção de máquinas, aluguel da área produtiva | Custos fixos classificados como "Custo indireto" |
| **Custo de tecnologia / infraestrutura** | Base fixa para entregar um serviço digital | nuvem mínima, licenças de software, domínio, contas de desenvolvedor | Custos fixos classificados como "Custo de tecnologia" (gastos por assinante ficam em 💳 Assinaturas) |
| **Despesa** | Gasto para administrar e vender, fora da produção | pró-labore, contador, marketing, tarifas bancárias | Custos fixos "Despesa administrativa / comercial / financeira" |
| **Despesa variável de venda** | Proporcional à receita | impostos, comissões, taxas de cartão | Custos de comercialização (%) |
| **Investimento** | Compra de bens duráveis; entra no resultado só pela depreciação | máquinas, móveis, reforma | Investimentos e fontes |

No **custeio por absorção**, o produto recebe o custo direto **mais** uma parte (rateio) da MOD, dos CIF e da depreciação.
No **custeio variável**, o produto recebe só o custo direto; MOD, CIF e despesas ficam como custos do período.""")
    t1, t2, t3 = st.tabs(["% Custos de comercialização", "👷 Mão de obra", "🏢 Custos fixos"])
    with t1:
        st.caption(f"Impostos sobre vendas ({proj['enquadramento']}): **{num_br(proj['aliquota_imposto_pct'])}%** (altere em Informações e sócios).")
        cv = editor("custos_variaveis", {"descricao": T("Descrição", required=True, width="large"),
                                         "percentual": N("% sobre a receita", min_value=0.0, max_value=100.0)}, "descricao")
        st.metric("Total de custos variáveis de venda (impostos + comercialização)",
                  f"{num_br(proj['aliquota_imposto_pct'] + cv['percentual'].sum())}%")
    with t2:
        pe = editor("pessoal", {"cargo": T("Cargo", required=True, width="medium"),
                                "area": S("Área", options=opcoes("area_pessoal", d["pessoal"]["area"].tolist()), required=True,
                                          help="Produção / operação = mão de obra direta (entra no custo do produto por absorção)."),
                                "num_empregados": N("Nº", min_value=0, step=1), "salario_mensal": moeda("Salário (R$)"),
                                "encargos_pct": N("Encargos (%)", min_value=0.0), "beneficios_unit": moeda("Benefícios por pessoa (R$)"),
                                "atribuicoes": T("Atribuições", width="medium"),
                                "mes_inicio": N("Começa no mês", min_value=1, max_value=120, step=1,
                                                help="1 = desde a abertura. Use 13 para contratar no início do Ano 2, 25 no Ano 3...")},
                    "cargo")
        st.caption("📅 **Contratações futuras:** para planejar quem entra depois, crie uma linha com o cargo e informe o mês de "
                   "início. O custo entra na DRE e no fluxo a partir desse mês.")
        r = R()
        if (d["pessoal"]["mes_inicio"] > 1).any():
            ps = r["pes_serie"]
            st.altair_chart(alt.Chart(pd.DataFrame({"Mês": range(1, len(ps) + 1), "Custo com pessoal": ps})).mark_area(
                line={"color": "#2f6fd0"}, color="#2f6fd033").encode(x="Mês:Q", y=alt.Y("Custo com pessoal:Q", title="R$/mês (sem inflação)"))
                .properties(height=200, title="Custo com pessoal mês a mês, com as contratações planejadas"), width="stretch")
        c = st.columns(4)
        c[0].metric("Total com pessoal / mês (Mês 1)", brl(r["pessoal_m"]))
        c[1].metric("Mão de obra direta (produção)", brl(r["mod_m"]))
        c[2].metric("Administrativo", brl(r["pessoal_adm_m"]))
        c[3].metric("Comercial / vendas", brl(r["pessoal_com_m"]))
        st.caption("Referência: no Simples Nacional (anexos I a III), encargos de cerca de 30% a 40% (FGTS, férias, 13º, "
                   "INSS sobre 13º); fora do Simples, em geral acima de 60%. Confirme com o contador.")
    with t3:
        editor("custos_fixos", {"descricao": T("Descrição", required=True, width="large"),
                                "classificacao": S("Classificação", options=opcoes("classe_custo_fixo", d["custos_fixos"]["classificacao"].tolist()),
                                                   required=True),
                                "valor_mensal": moeda("Valor mensal (R$)")}, "descricao")
        r = R()
        c = st.columns(4)
        c[0].metric("Custos indiretos de produção", brl(r["cif_m"]))
        c[1].metric("Despesas administrativas", brl(r["desp_adm_m"]))
        c[2].metric("Despesas comerciais", brl(r["desp_com_m"]))
        c[3].metric("Despesas financeiras", brl(r["desp_fin_m"]))
    r = R()
    st.subheader("Estrutura de custos mensal (Mês 1)")
    est = pd.DataFrame({"Item": ["Custos diretos (materiais e mercadorias)", "Impostos e despesas variáveis de venda",
                                 "Mão de obra direta (produção)", "Custos indiretos de produção", "Depreciação",
                                 "Pessoal administrativo e comercial", "Despesas administrativas, comerciais e financeiras"],
                        "Valor": [r["cdir_m1"], r["receita_m1"] * (r["aliq"] + r["var_pct"]), r["mod_m"], r["cif_m"], r["dep_m"],
                                  r["pessoal_adm_m"] + r["pessoal_com_m"], r["desp_adm_m"] + r["desp_com_m"] + r["desp_fin_m"]]})
    est = est[est["Valor"] > 0]
    est["%"] = est["Valor"] / est["Valor"].sum()
    c1, c2 = st.columns(2)
    with c1:
        rosca(est, "Para onde vai o dinheiro no mês", "g_estcustos", tabela=False)
    with c2:
        st.dataframe(tabela_comp(est), hide_index=True, width="stretch")
        st.metric("Custos que vão no produto (absorção): MOD + CIF + depreciação", brl(r["overhead_m"]))
    proxima("custos")


# ==============================================================================
# FORMAÇÃO DE PREÇO
# ==============================================================================
def pg_preco():
    cabecalho("preco")
    v = {}
    c1, c2 = st.columns(2)
    v["margem_desejada_pct"] = c1.number_input("Margem de lucro desejada sobre o preço (%)", 0.0, 80.0,
                                               float(proj["margem_desejada_pct"]), 1.0, key=wk("mdes"))
    crits = ["Custo direto", "Receita", "Quantidade"]
    v["rateio_criterio"] = c2.selectbox("Critério de rateio dos custos indiretos (absorção)", crits, key=wk("rat"),
                                        index=crits.index(proj["rateio_criterio"]) if proj["rateio_criterio"] in crits else 0)
    auto_unica("projeto", v)
    r = R()
    st.markdown(f"""
**Markup divisor:** preço = custo ÷ [1 − (impostos + despesas variáveis + despesas fixas + margem desejada)]

= custo ÷ [1 − ({pct(r['aliq'])} + {pct(r['var_pct'])} + {pct(r['df_pct'])} + {pct(r['margem_desejada'])})]
= custo × **{num_br(r['markup_mult'], 3)}**

*Despesas fixas* = (custos fixos + pessoal + depreciação) ÷ receita do Mês 1 = {pct(r['df_pct'])}.""")
    if r["divisor_markup"] <= 0:
        st.error("A soma dos percentuais passou de 100%: não existe preço que cubra tudo. Reduza custos fixos ou a margem desejada.")
    cat = r["cat"]
    if cat.empty:
        st.info("Cadastre produtos em Produtos e materiais.")
        proxima("preco")
        return
    tab = pd.DataFrame({
        "Produto / serviço": cat["descricao"], "Categoria": cat["categoria"],
        "Custo direto (variável)": cat["custo_unit"].map(brl), "Custo por absorção": cat["custo_absorcao_unit"].map(brl),
        "Preço praticado": cat["preco_venda"].map(brl), "Preço sugerido (markup)": cat["preco_sugerido"].map(brl),
        "Diferença": cat["dif_preco_pct"].map(lambda v: ("🔻 " if v < -0.05 else "🔺 " if v > 0.05 else "✔️ ") + pct(v) if v == v else "n.a."),
        "Margem (variável)": cat["margem_pct"].map(pct), "Margem (absorção)": cat["margem_absorcao_pct"].map(pct),
        "PE (unid./mês)": cat["pe_unidades_mes"].map(lambda v: num_br(v, 0)), "Venda prevista (unid./mês)": cat["qtd_mensal"].map(lambda v: num_br(v, 0))})
    st.dataframe(tab, hide_index=True, width="stretch", height=480)
    st.caption("🔻 preço abaixo do sugerido (pode não cobrir todos os custos com a margem desejada) · 🔺 acima do sugerido "
               "(confira se o mercado aceita) · PE = quantidade mínima por mês para o negócio empatar, mantido o mix atual de vendas.")
    abaixo = cat[cat["dif_preco_pct"] < -0.05]
    if not abaixo.empty:
        st.warning(f"**{len(abaixo)} item(ns) com preço mais de 5% abaixo do sugerido:** " + ", ".join(abaixo["descricao"].head(8)))
    prej = cat[cat["margem_absorcao_pct"] < 0]
    if not prej.empty:
        st.error("**Itens que dão prejuízo no custeio por absorção:** " + ", ".join(prej["descricao"]))
    proxima("preco")


# ==============================================================================
# DRE, FLUXO DE CAIXA E INDICADORES
# ==============================================================================
def pg_dre():
    cabecalho("dre")
    r = R()
    st.subheader("Projeção mensal do Ano 1")
    mm = r["meses"]
    st.line_chart(mm.set_index("Mês")[["Receita", "Custos variáveis", "Custos fixos", "Resultado operacional"]])
    with st.expander("Ver tabela mensal"):
        st.dataframe(mm.assign(**{c: mm[c].map(brl) for c in mm.columns[1:]}), hide_index=True, width="stretch")
    st.subheader("Demonstração do resultado (Anos 1 a 3)")
    st.dataframe(tabela_dre(r), hide_index=True, width="stretch", height=560)
    mostrar_leitura("dre", leitura_dre(r))
    st.bar_chart(pd.DataFrame({"Receita": [r["anos"][a]["rec_bruta"] for a in r["ANOS"]],
                               "EBITDA": [r["anos"][a]["ebitda"] for a in r["ANOS"]],
                               "Resultado líquido": [r["anos"][a]["resultado"] for a in r["ANOS"]]},
                              index=[f"Ano {a:02d}" for a in r["ANOS"]]), stack=False)
    proxima("dre")


def pg_fluxo():
    cabecalho("fluxo")
    r = R()
    fm = r["fluxo_mensal"]
    if fm.empty:
        st.info("Sem dados.")
        return
    minimo = fm["Saldo final"].min()
    mes = fm.loc[fm["Saldo final"].idxmin(), "Mês"]
    c = st.columns(3)
    c[0].metric("Saldo inicial (após investimentos)", brl(fm["Saldo inicial"].iloc[0]))
    c[1].metric("Menor saldo do ano", brl(minimo), mes, delta_color="off")
    c[2].metric("Saldo no fim do Ano 1", brl(fm["Saldo final"].iloc[-1]))
    if minimo < 0:
        st.error(f"🔴 O caixa fica negativo em {mes}. Aumente a reserva/capital de giro, alongue prazos ou reduza custos.")
    g = fm[["Mês", "Recebimentos"]].assign(Pagamentos=fm[["Fornecedores", "Impostos e desp. variáveis", "Pessoal e custos fixos",
                                                          "Parcelas de empréstimos"]].sum(axis=1))
    longo = g.melt(id_vars="Mês", var_name="Tipo", value_name="R$")
    barras_ = alt.Chart(longo).mark_bar(opacity=.8).encode(x=alt.X("Mês:N", sort=list(fm["Mês"]), title=None), y="R$:Q",
                                                            color=alt.Color("Tipo:N", scale=alt.Scale(range=["#2f6fd0", "#d6006e"])),
                                                            xOffset="Tipo:N")
    linha = alt.Chart(fm).mark_line(point=True, color="#27ae60", strokeWidth=3).encode(
        x=alt.X("Mês:N", sort=list(fm["Mês"])), y=alt.Y("Saldo final:Q", title="R$"))
    st.altair_chart((barras_ + linha).properties(height=380), width="stretch")
    st.caption("Barras: recebimentos (azul) e pagamentos (rosa) do mês. Linha verde: saldo acumulado em caixa.")
    st.dataframe(fm.assign(**{c: fm[c].map(brl) for c in fm.columns[1:]}), hide_index=True, width="stretch")
    mostrar_leitura("fluxo", leitura_fluxo(r))
    proxima("fluxo")


def pg_indicadores():
    cabecalho("indicadores")
    r = R()
    sem = semaforo(r)
    st.subheader("🚦 Saúde do negócio")
    for i in range(0, len(sem), 3):
        cols = st.columns(3)
        for col, s_ in zip(cols, sem[i:i + 3]):
            with col.container(border=True):
                st.markdown(f"### {s_['Status']} {s_['Valor']}")
                st.markdown(f"**{s_['Indicador']}**")
                st.caption(s_["Leitura"])
                with st.popover("🔎 Como foi calculado"):
                    st.write(s_["Como foi calculado"])
    with st.expander("📖 O que significam estes números? (leitura em linguagem simples)"):
        for nome, valor, leitura in explicar_indicadores(r):
            st.markdown(f"**{nome}: {valor}**  \n{leitura}")
    st.subheader("Resumo financeiro (Ano 1, padrão Sebrae)")
    rs, a1 = r["resumo"], r["anos"][1]
    k = st.columns(4)
    k[0].metric("Investimento total", brl(r["investimento_total"]))
    k[1].metric("Receita anual", brl(rs["receita_anual"]))
    k[2].metric("Resultado operacional", brl(rs["resultado_operacional"]))
    k[3].metric("Lucratividade", pct(rs["lucratividade"]))
    k = st.columns(4)
    k[0].metric("Rentabilidade", pct(rs["rentabilidade"]))
    k[1].metric("Payback simples", f"{num_br(rs['payback_meses'], 1)} meses" if rs["payback_meses"] == rs["payback_meses"] else "não recupera")
    k[2].metric(f"VPL (TMA {num_br(proj['tma_anual_pct'], 1)}%)", brl(r["vpl"]))
    k[3].metric("TIR", pct(r["tir"]) if r["tir"] is not None else "n.a.")
    k = st.columns(4)
    k[0].metric("PE contábil (anual)", brl(a1["pe_contabil"]))
    k[1].metric("PE financeiro (anual)", brl(a1["pe_financeiro"]))
    k[2].metric("PE econômico (anual)", brl(a1["pe_economico"]))
    k[3].metric("Payback descontado", f"{num_br(r['payback_desc'], 1)} meses" if r["payback_desc"] == r["payback_desc"] else f"> {r['MESES']} meses")

    st.subheader("Matriz de indicadores (5 blocos)")
    ind = tabela_indicadores(r)
    bsel = st.selectbox("Bloco", ["Todos"] + list(ind["Bloco"].unique()), key=wk("bloco"))
    st.dataframe(ind if bsel == "Todos" else ind[ind["Bloco"] == bsel], hide_index=True, width="stretch")

    st.subheader("✔️ Checagens de consistência")
    chk = CHK()
    for i, c_ in chk.iterrows():
        ok = c_["Situação"] == "OK"
        a, b, c = st.columns([3.2, 2.2, 1.4])
        a.markdown(f"{'🟢' if ok else '🔴'} **{c_['Teste']}**  \n<small>{c_['Observação']}</small>", unsafe_allow_html=True)
        b.markdown(f"{c_['Valor apurado']}" + ("" if ok else f"  \n<small>📍 {c_['Onde corrigir']}</small>"), unsafe_allow_html=True)
        if not ok and c_["destino"]:
            if c.button("VERIFICAR → ir ao ajuste", key=f"chk_{i}", type="primary", width="stretch"):
                ir(c_["destino"], f"{c_['Teste']}: {c_['Valor apurado']}. {c_['Como corrigir'] or c_['Observação']}")
                st.rerun()
        elif ok:
            c.markdown("✅ OK")
    with st.expander("Notas metodológicas"):
        st.markdown(
            "- **Custos variáveis**: custo direto dos itens (ficha técnica ou valor de aquisição) + impostos + custos de comercialização.\n"
            "- **Custos fixos totais**: pessoal + custos fixos + depreciação (linear, com valor residual).\n"
            "- **Fluxo do projeto** (VPL/TIR): Ano 0 = investimento total; Anos 1 a 3 = EBITDA (no Simples, os tributos já estão "
            "na alíquota sobre vendas). Opcionalmente recupera o capital de giro no último ano do horizonte.\n"
            "- **Payback simples e rentabilidade** seguem o modelo do Sebrae: investimento ÷ resultado operacional mensal do Ano 1.\n"
            "- **Juros** pela Tabela Price, mês a mês. **Balanço simplificado**: caixa por diferença (Ativo = Passivo + PL), lucros retidos.")
    proxima("indicadores")


def pg_graficos():
    cabecalho("graficos")
    comp = composicoes(R(), d)
    st.subheader("Investimento inicial")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["investimento"], "Composição do investimento total", "g_inv")
    with g2:
        rosca(comp["fontes"], "Fontes de recursos", "g_fontes")
    barras(comp["inv_fixos"], "Maiores itens do ativo imobilizado", "g_fixos")
    st.subheader("Materiais e custos diretos")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["estoque"], "Estoque inicial por categoria", "g_est")
    with g2:
        rosca(comp["custo_direto"], "Custo direto mensal por categoria", "g_cd")
    st.subheader("Receitas")
    g1, g2 = st.columns(2)
    with g1:
        rosca(comp["receita_categoria"], "Receita por categoria", "g_rc")
    with g2:
        barras(comp["top_produtos"], "10 itens com maior receita mensal", "g_top")
    st.subheader("Para onde vai cada R$ 1,00 de receita (Ano 1)")
    barras(comp["estrutura_receita"], "Custos, despesas e lucro em % da receita bruta", "g_estr")
    st.subheader("Ponto de equilíbrio contábil (Ano 1)")
    eq = comp["equilibrio"].melt(id_vars="Receita", var_name="Série", value_name="R$")
    pe, rec = comp["pe_ponto"]
    linhas = alt.Chart(eq).mark_line(strokeWidth=3).encode(x=alt.X("Receita:Q", title="Receita anual (R$)"), y="R$:Q",
                                                            color=alt.Color("Série:N", legend=alt.Legend(orient="bottom", title=None)))
    marcas = alt.Chart(pd.DataFrame({"x": [pe, rec], "rot": ["Ponto de equilíbrio", "Receita projetada"]})).mark_rule(
        strokeDash=[6, 4]).encode(x="x:Q", tooltip=["rot", alt.Tooltip("x:Q", format=",.2f")])
    st.altair_chart((linhas + marcas).properties(height=380), width="stretch", key="g_pe")
    st.caption(f"PE contábil: {brl(pe)} por ano · receita projetada: {brl(rec)} · margem de segurança: {pct(div(rec - pe, rec))}")
    st.subheader("Recuperação do investimento (payback)")
    fa = comp["fluxo_acumulado"]
    st.bar_chart(fa.set_index("Ano")[["Acumulado", "Acumulado descontado (TMA)"]], stack=False)
    proxima("graficos")


# ==============================================================================
# SIMULADOR, EXPERIMENTOS, RISCOS, CRONOGRAMA, ORÇADO X REALIZADO
# ==============================================================================
def pg_simulador():
    cabecalho("simulador")
    r = R()
    t1, t2, t3 = st.tabs(["🎚️ E se...?", "🌪️ Sensibilidade", "📊 Cenários"])
    with t1:
        st.caption("Mova as barras e veja o resultado do Ano 1 mudar na hora (nada é gravado no projeto).")
        c = st.columns(5)
        vp = c[0].slider("Preço de venda", -30, 30, 0, 1, format="%d%%", key=wk("s1"))
        vv = c[1].slider("Volume de vendas", -50, 50, 0, 1, format="%d%%", key=wk("s2"))
        vm = c[2].slider("Custo de materiais", -30, 30, 0, 1, format="%d%%", key=wk("s3"))
        vs = c[3].slider("Salários e encargos", -30, 30, 0, 1, format="%d%%", key=wk("s4"))
        vf = c[4].slider("Demais custos fixos", -30, 30, 0, 1, format="%d%%", key=wk("s5"))
        base, sim = simular(r), simular(r, vp / 100, vv / 100, vm / 100, vs / 100, vf / 100)
        k = st.columns(4)
        k[0].metric("Receita (Ano 1)", brl(sim["receita"]), brl(sim["receita"] - base["receita"]))
        k[1].metric("Margem de contribuição", brl(sim["mc"]), brl(sim["mc"] - base["mc"]))
        k[2].metric("Resultado líquido", brl(sim["resultado"]), brl(sim["resultado"] - base["resultado"]))
        k[3].metric("Lucratividade", pct(sim["lucratividade"]),
                    f"{num_br((sim['lucratividade'] - base['lucratividade']) * 100)} p.p." if sim["lucratividade"] == sim["lucratividade"] else None)
        if sim["resultado"] < 0:
            st.error("🔴 Nesse cenário o negócio tem prejuízo no Ano 1.")
    with t2:
        sens = sensibilidade(r)
        st.caption("Efeito de variar cada item em 10% (para mais e para menos) no resultado do Ano 1. "
                   "As barras mais longas mostram onde o negócio é mais sensível.")
        longo = sens.melt(id_vars="Variável", value_vars=["+10%", "-10%"], var_name="Variação", value_name="Efeito no resultado")
        st.altair_chart(alt.Chart(longo).mark_bar().encode(
            y=alt.Y("Variável:N", sort=list(sens["Variável"]), title=None), x=alt.X("Efeito no resultado:Q", title="R$"),
            color=alt.Color("Variação:N", scale=alt.Scale(range=["#27ae60", "#d6006e"]))).properties(height=260), width="stretch")
        st.info(f"O item que mais mexe no lucro é **{sens.iloc[0]['Variável']}**: 10% de variação muda o resultado em "
                f"{brl(sens.iloc[0]['Impacto absoluto'])}.")
    with t3:
        v = {}
        c = st.columns(4)
        v["cen_otim_receita_pct"] = c[0].number_input("Otimista: receita maior em (%)", 0.0, 100.0, proj["cen_otim_receita_pct"], key=wk("c1"))
        v["cen_otim_custo_pct"] = c[1].number_input("Otimista: custos menores em (%)", 0.0, 100.0, proj["cen_otim_custo_pct"], key=wk("c2"))
        v["cen_pess_receita_pct"] = c[2].number_input("Pessimista: receita menor em (%)", 0.0, 100.0, proj["cen_pess_receita_pct"], key=wk("c3"))
        v["cen_pess_custo_pct"] = c[3].number_input("Pessimista: custos maiores em (%)", 0.0, 100.0, proj["cen_pess_custo_pct"], key=wk("c4"))
        auto_unica("projeto", v)
        st.dataframe(cenarios(R(), proj), hide_index=True, width="stretch")
    proxima("simulador")


def pg_experimentos():
    cabecalho("experimentos")
    st.caption("Antes de investir, teste as hipóteses mais arriscadas. Ex.: “Pessoas pagariam R$ 29,90 por uma marmita com macros?” "
               "→ teste: pré-venda para 30 pessoas → métrica: % que compra → sucesso: acima de 40%.")
    ex = editor("experimentos", {"hipotese": T("Hipótese (o que acreditamos)", required=True, width="medium"),
                                 "experimento": T("Experimento (como testar)", width="medium"), "metrica": T("Métrica"),
                                 "criterio_sucesso": T("Critério de sucesso"), "resultado": T("Resultado obtido", width="medium"),
                                 "status": S("Situação", options=["A testar", "Em teste", "Validada", "Invalidada", "Pivotar"])}, "hipotese")
    if not ex.empty:
        cont = ex["status"].replace("", "A testar").value_counts()
        st.caption(" · ".join(f"{k}: {v}" for k, v in cont.items()))
    proxima("experimentos")


def pg_riscos():
    cabecalho("riscos")
    rk = editor("riscos", {"risco": T("Risco", required=True, width="medium"),
                           "probabilidade": N("Probabilidade (1 a 5)", min_value=1, max_value=5, step=1),
                           "impacto": N("Impacto (1 a 5)", min_value=1, max_value=5, step=1),
                           "mitigacao": T("Plano de contingência / mitigação", width="large"), "responsavel": T("Responsável")}, "risco")
    if not rk.empty:
        g = rk.assign(Nível=rk["probabilidade"] * rk["impacto"])
        fundo = pd.DataFrame([{"p": p, "i": i, "n": p * i} for p in range(1, 6) for i in range(1, 6)])
        mapa = alt.Chart(fundo).mark_rect(opacity=.55).encode(
            x=alt.X("p:O", title="Probabilidade"), y=alt.Y("i:O", sort="descending", title="Impacto"),
            color=alt.Color("n:Q", scale=alt.Scale(domain=[1, 8, 25], range=["#27ae60", "#f1c40f", "#e74c3c"]), legend=None))
        pts = alt.Chart(g).mark_text(fontSize=12, fontWeight="bold").encode(
            x="probabilidade:O", y=alt.Y("impacto:O", sort="descending"), text="risco:N")
        st.altair_chart((mapa + pts).properties(height=360), width="stretch")
        alto = g[g["Nível"] >= 12]
        if not alto.empty:
            st.warning("**Riscos críticos (nível ≥ 12):** " + ", ".join(alto["risco"]))
    proxima("riscos")


def pg_cronograma():
    cabecalho("cronograma")
    cr = editor("cronograma", {"atividade": T("Atividade", required=True, width="medium"),
                               "mes_inicio": N("Mês de início (1 = primeiro mês)", min_value=1, step=1),
                               "duracao_meses": N("Duração (meses)", min_value=0.25, step=0.25),
                               "responsavel": T("Responsável"), "custo": moeda("Custo previsto (R$)")}, "atividade")
    if not cr.empty:
        g = cr.assign(fim=cr["mes_inicio"] + cr["duracao_meses"])
        st.altair_chart(alt.Chart(g).mark_bar(cornerRadius=4, color="#2f6fd0").encode(
            x=alt.X("mes_inicio:Q", title="Mês"), x2="fim:Q", y=alt.Y("atividade:N", sort=None, title=None),
            tooltip=["atividade", "responsavel", "mes_inicio", "duracao_meses"]).properties(height=34 * len(g) + 40), width="stretch")
        st.metric("Custo total previsto no cronograma", brl(cr["custo"].sum()))
    proxima("cronograma")


def pg_realizado():
    cabecalho("realizado")
    st.caption("Depois que o negócio abrir, lance mês a mês o que aconteceu de verdade. O programa compara com o planejado.")
    rz = editor("realizado", {"mes": N("Mês (1 a 12)", min_value=1, max_value=12, step=1, required=True),
                              "receita_real": moeda("Receita real (R$)"), "custos_variaveis_real": moeda("Custos variáveis reais (R$)"),
                              "custos_fixos_real": moeda("Custos fixos reais (R$)")}, "mes")
    rz = rz[rz["mes"] > 0]
    if not rz.empty:
        mm = R()["meses"].reset_index(drop=True)
        linhas = []
        for _, x in rz.sort_values("mes").iterrows():
            m = int(x["mes"]) - 1
            if 0 <= m < len(mm):
                prev = mm.iloc[m]
                res_real = x["receita_real"] - x["custos_variaveis_real"] - x["custos_fixos_real"]
                linhas.append({"Mês": f"Mês {m + 1}", "Receita prevista": prev["Receita"], "Receita real": x["receita_real"],
                               "Desvio receita": div(x["receita_real"] - prev["Receita"], prev["Receita"]),
                               "Resultado previsto": prev["Resultado operacional"], "Resultado real": res_real})
        cmp_ = pd.DataFrame(linhas)
        if not cmp_.empty:
            st.dataframe(cmp_.assign(**{c: cmp_[c].map(brl) for c in ["Receita prevista", "Receita real", "Resultado previsto", "Resultado real"]},
                                     **{"Desvio receita": cmp_["Desvio receita"].map(lambda v: ("🔴 " if v < -0.1 else "🟡 " if v < 0 else "🟢 ") + pct(v))}),
                         hide_index=True, width="stretch")
            longo = cmp_.melt(id_vars="Mês", value_vars=["Receita prevista", "Receita real"], var_name="Série", value_name="R$")
            st.altair_chart(alt.Chart(longo).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("Mês:N", sort=None), y="R$:Q", color="Série:N").properties(height=320), width="stretch")
    proxima("realizado")

# ##############################################################################
# PARTE: pnbox_ui_paginas4.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: SERVIÇOS (CUSTO DA HORA) E ASSINATURAS (RECEITA RECORRENTE)
# ==============================================================================


def _info(html: str):
    st.markdown(f'<div class="pn-card-info">{html}</div>', unsafe_allow_html=True)
    st.write("")


# ------------------------------------------------------------------------------
# SERVIÇOS E CUSTO DA HORA
# ------------------------------------------------------------------------------
def pg_servicos():
    cabecalho("servicos")
    _info("Em serviços, o principal custo é o <b>tempo da equipe</b>. O programa calcula quanto custa cada hora "
          "efetivamente cobrável e, a partir dela, o custo pleno e o preço sugerido de cada serviço. Cadastre a equipe da "
          "operação em <b>🧾 Custos e despesas</b> (área \"Produção / operação\").")
    t1, t2, t3 = st.tabs(["⏱️ Horas da equipe", "🧾 Custo e preço dos serviços", "📚 Como funciona"])
    v = {}
    with t1:
        c1, c2 = st.columns(2)
        v["horas_mes_pessoa"] = c1.number_input("Horas contratadas por pessoa no mês", 0.0, 400.0, float(proj["horas_mes_pessoa"] or 168),
                                                4.0, key=wk("hmes"),
                                                help="44 h semanais ≈ 220 h/mês com o descanso remunerado; 40 h ≈ 176 h; 6 h/dia em 20 dias = 120 h.")
        v["utilizacao_pct"] = c2.slider("Taxa de utilização (% das horas que vira trabalho cobrável)", 10, 100,
                                        int(proj["utilizacao_pct"] or 75), 5, key=wk("util"),
                                        help="Desconte reuniões, deslocamentos, orçamentos, treinamento, retrabalho e férias. "
                                             "Empresas de serviços costumam ficar entre 60% e 80%.")
        auto_unica("projeto", v)
        res = R()
        ch = custo_hora(res, d, proj)
        k = st.columns(4)
        k[0].metric("Pessoas na operação", num_br(ch["pessoas"], 0))
        k[1].metric("Horas contratadas/mês", num_br(ch["horas_contratadas"], 0))
        k[2].metric("Horas cobráveis/mês", num_br(ch["horas_cobraveis"], 0))
        k[3].metric("Custo da hora produtiva", brl(ch["custo_hora"]) if ch["custo_hora"] == ch["custo_hora"] else "n.a.",
                    help=f"(Mão de obra da operação {brl(res['mod_m'])} + custos indiretos e de tecnologia {brl(res['cif_m'])} "
                         f"+ depreciação {brl(res['dep_m'])}) ÷ {num_br(ch['horas_cobraveis'], 0)} horas cobráveis")
        if ch["pessoas"] == 0:
            st.warning("Não há ninguém na área \"Produção / operação\". Sem equipe não há horas para vender.")
            if st.button("Ir para Custos e despesas →"):
                ir("custos", "Cadastre na aba Mão de obra quem executa os serviços, com a área \"Produção / operação\".")
                st.rerun()
        if ch["horas_cobraveis"] > 0:
            st.markdown("**Ocupação da equipe (horas vendidas ÷ horas cobráveis)**")
            o1, o12 = ch["ocupacao_m1"], ch["ocupacao_m12"]
            st.progress(min(1.0, o1 if o1 == o1 else 0.0), text=f"Mês 1: {num_br(ch['horas_demandadas_m1'], 0)} h vendidas = {pct(o1)}")
            st.progress(min(1.0, o12 if o12 == o12 else 0.0), text=f"Mês 12: {num_br(ch['horas_demandadas_m12'], 0)} h vendidas = {pct(o12)}")
            if o12 == o12 and o12 > 1:
                st.error("No Mês 12 a equipe não dá conta das horas vendidas. Planeje contratações ou reduza a meta.")
            elif o12 == o12 and o12 < 0.5:
                st.info("Há muita hora ociosa: dá para vender mais serviços sem aumentar a equipe.")
    with t2:
        st.caption("Serviços cadastrados (os mesmos de Produtos e materiais, categoria Serviço). **Insumos** são materiais "
                   "gastos em cada execução; o tempo da equipe entra pelas **horas por serviço**.")
        editor("catalogo", {
            "descricao": T("Serviço", required=True, width="large"),
            "preco_venda": moeda("Preço cobrado (R$)"),
            "custo_aquisicao": moeda("Insumos por serviço (R$)"),
            "qtd_mensal": N("Qtd. no Mês 1", min_value=0, step=1),
            "horas_unit": N("Horas por serviço", min_value=0.0, step=0.5),
            "qtd_estoque_inicial": None}, "descricao", filtro={"categoria": "Serviço"})
        res = R()
        ch = custo_hora(res, d, proj)
        tab = ch["tabela"]
        if tab.empty:
            st.info("Informe as horas por serviço para ver o custo pleno e o preço sugerido.")
        else:
            st.markdown(f"**Custo da hora:** {brl(ch['custo_hora'])} · **Despesas de estrutura:** {pct(ch['desp_pct'])} da receita "
                        f"(média do Ano 1) · **Margem desejada:** {pct(res['margem_desejada'])}")
            vis = tab.copy()
            for c in ("Insumos", "Mão de obra (horas x custo hora)", "Custo pleno", "Preço atual", "Preço sugerido",
                      "Valor da hora cobrada"):
                vis[c] = vis[c].map(lambda x: brl(x) if x == x else "não calculável")
            vis["Margem após custos diretos"] = tab["Margem após custos diretos"].map(pct)
            for c in ("Horas por serviço", "Qtd/mês", "Horas/mês"):
                vis[c] = vis[c].map(lambda x: num_br(x, 1))
            st.dataframe(vis, hide_index=True, width="stretch")
            if ch["divisor"] <= 0.05:
                st.warning("As despesas de estrutura, os impostos e a margem desejada somam quase 100% da receita: nenhum preço "
                           "cobre tudo isso. Reduza a estrutura, revise a margem ou aumente o volume de vendas.")
            baixos = tab[(tab["Preço sugerido"] == tab["Preço sugerido"]) & (tab["Preço atual"] < tab["Preço sugerido"])]
            if not baixos.empty:
                st.warning("Abaixo do preço sugerido: " + ", ".join(baixos["Serviço"]) + ".")
            g = tab.melt("Serviço", ["Custo pleno", "Preço atual", "Preço sugerido"], "Valor", "R$").dropna()
            st.altair_chart(alt.Chart(g).mark_bar().encode(
                x=alt.X("Serviço:N", title=None, axis=alt.Axis(labelAngle=0, labelLimit=180)), xOffset="Valor:N",
                y=alt.Y("R$:Q", title="R$"),
                color=alt.Color("Valor:N", scale=alt.Scale(domain=["Custo pleno", "Preço atual", "Preço sugerido"],
                                                           range=["#c0392b", "#2f6fd0", "#18a058"]),
                                legend=alt.Legend(orient="bottom", title=None)),
                tooltip=["Serviço", "Valor", alt.Tooltip("R$:Q", format=",.2f")]).properties(height=320), width="stretch")
    with t3:
        st.markdown(r"""
#### Passo a passo do preço de um serviço
1. **Horas cobráveis** = pessoas da operação × horas contratadas no mês × taxa de utilização.
   Ninguém trabalha 100% do tempo em atividades que o cliente paga: há deslocamento, orçamento, reunião interna, estudo.
2. **Custo da hora produtiva** = (salários, encargos e benefícios da operação + custos indiretos e de tecnologia +
   depreciação dos equipamentos) ÷ horas cobráveis.
3. **Custo pleno do serviço** = insumos + horas do serviço × custo da hora.
4. **Preço sugerido** = custo pleno ÷ [1 − (impostos + despesas variáveis + despesas de estrutura + margem desejada)],
   o mesmo *markup* divisor usado para produtos.

#### Leituras úteis
- **Valor da hora cobrada** (preço ÷ horas) mostra se o serviço remunera o tempo da equipe.
- **Ocupação** acima de 100% indica que faltam pessoas; abaixo de 50% indica capacidade ociosa.
- Para profissional autônomo, o mesmo raciocínio vale: some os custos mensais e o pró-labore desejado, divida pelas horas
  produtivas e acrescente os impostos.
- Serviços vendidos por pacote (mensalidade de manutenção, plano de suporte) também podem ser tratados como
  **assinatura**, em 💳 Assinaturas e receita recorrente.
""")
    proxima("servicos")


# ------------------------------------------------------------------------------
# ASSINATURAS E RECEITA RECORRENTE
# ------------------------------------------------------------------------------
CICLOS = {1: "Mensal", 3: "Trimestral", 6: "Semestral", 12: "Anual"}


def pg_assinaturas():
    cabecalho("assinaturas")
    _info("Em negócios por assinatura o cliente não compra uma vez: ele <b>fica</b> e paga de novo a cada ciclo, até "
          "cancelar. Por isso o programa projeta <b>safras de assinantes</b> mês a mês, separando o dinheiro que entra "
          "(<b>cobranças</b>) da receita que pertence a cada mês (<b>receita reconhecida</b>). Um plano anual pago adiantado "
          "entra todo no caixa, mas vira receita 1/12 por mês.")
    t1, t2, t3, t4, t5 = st.tabs(["📋 Planos", "📈 Projeção mês a mês", "📊 MRR, churn, CAC e LTV", "☁️ Custos de tecnologia",
                                  "📚 Como funciona"])
    with t1:
        st.caption("Um plano por linha. **Preço por ciclo** é o valor cobrado de uma vez (ex.: R$ 3.000 no plano anual). "
                   "**Churn** é a % de assinantes que cancelam por mês. **Custo de entrega** é o gasto mensal por assinante "
                   "ativo (nuvem, APIs, SMS/e-mail, suporte). **Taxa de cobrança** é o % do meio de pagamento ou da loja de "
                   "aplicativos.")
        editor("assinaturas", {
            "plano": T("Plano", required=True, width="medium"),
            "publico": S("Público", options=["Empresas (B2B)", "Consumidores (B2C)"]),
            "preco_ciclo": moeda("Preço por ciclo (R$)"),
            "ciclo_meses": N("Ciclo (meses)", min_value=1, max_value=60, step=1,
                             help="1 = mensal, 3 = trimestral, 6 = semestral, 12 = anual"),
            "novos_mes1": N("Novos assinantes no Mês 1", min_value=0.0, step=1.0),
            "crescimento_novos_pct": N("Crescimento dos novos (% ao mês)", step=0.5,
                                       help="Quanto a captação de novos assinantes cresce por mês no Ano 1."),
            "churn_mensal_pct": N("Churn (% ao mês)", min_value=0.0, max_value=100.0, step=0.5),
            "custo_entrega_mensal": moeda("Custo de entrega por assinante/mês (R$)"),
            "taxa_cobranca_pct": N("Taxa de cobrança (%)", min_value=0.0, max_value=50.0, step=0.5)}, "plano")
        ass = d["assinaturas"]
        if not ass.empty:
            resumo = pd.DataFrame({
                "Plano": ass["plano"], "Ciclo": ass["ciclo_meses"].map(lambda c: CICLOS.get(int(c), f"{int(c)} meses")),
                "Equivale por mês": [brl(div(p, c)) for p, c in zip(ass["preco_ciclo"], ass["ciclo_meses"].clip(lower=1))],
                "Vida média do assinante": [f"{num_br(div(1, c / 100), 0)} meses" if c > 0 else "infinita (churn zero!)"
                                            for c in ass["churn_mensal_pct"]],
                "Lucro bruto por assinante/mês": [brl(div(p, c) * (1 - t / 100) - e) for p, c, t, e in
                                                  zip(ass["preco_ciclo"], ass["ciclo_meses"].clip(lower=1),
                                                      ass["taxa_cobranca_pct"], ass["custo_entrega_mensal"])]})
            st.dataframe(resumo, hide_index=True, width="stretch")
            if (ass["churn_mensal_pct"] <= 0).any():
                st.warning("Plano com churn zero: na prática, todo negócio por assinatura perde clientes. Informe uma estimativa.")
        st.markdown("**Custo de aquisição (CAC)**")
        verba = st.number_input("Quanto do gasto comercial mensal é dedicado a conquistar novos assinantes? (R$)", 0.0, 1e7,
                                float(proj["verba_aquisicao_mensal"] or 0), 100.0, key=wk("verba_cac"),
                                help="Anúncios, comissões de venda, eventos de captação... (valor já lançado em Custos e despesas). "
                                     "Deixe 0 para usar todo o gasto comercial fixo do Mês 1.")
        auto_unica("projeto", {"verba_aquisicao_mensal": verba})
    res = R()
    s, planos, rc = res["ass"], res["ass_planos"], res["recorrencia"]
    with t2:
        if s.empty:
            st.info("Cadastre ao menos um plano na aba 📋 Planos.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Assinantes ativos por plano**")
                st.altair_chart(alt.Chart(planos).mark_area(opacity=0.85).encode(
                    x=alt.X("Mês:Q", scale=alt.Scale(domain=[1, res["MESES"]])), y=alt.Y("Ativos:Q", stack=True, title="Assinantes"),
                    color=alt.Color("Plano:N", legend=alt.Legend(orient="bottom", columns=1, title=None)),
                    tooltip=["Mês", "Plano", alt.Tooltip("Ativos:Q", format=",.0f")]).properties(height=330), width="stretch")
            with c2:
                st.markdown("**Receita reconhecida (MRR) x dinheiro que entra (cobranças)**")
                lg = s.melt("Mês", ["Receita reconhecida", "Cobranças (caixa)"], "Série", "R$")
                st.altair_chart(alt.Chart(lg).mark_line(point=False, strokeWidth=2.5).encode(
                    x=alt.X("Mês:Q", scale=alt.Scale(domain=[1, res["MESES"]])), y="R$:Q",
                    color=alt.Color("Série:N", scale=alt.Scale(range=["#2f6fd0", "#e67e22"]),
                                    legend=alt.Legend(orient="bottom", title=None)),
                    tooltip=["Mês", "Série", alt.Tooltip("R$:Q", format=",.2f")]).properties(height=330), width="stretch")
            anos = []
            for k in res["ANOS"]:
                bloco = s.iloc[(k - 1) * 12:k * 12]
                r_, c_ = float(bloco["Receita reconhecida"].sum()), float(bloco["Cobranças (caixa)"].sum())
                ce, tx = float(bloco["Custo de entrega"].sum()), float(bloco["Taxas de cobrança"].sum())
                anos.append({"Ano": f"Ano {k}", "Novos": num_br(bloco["Novos"].sum(), 0),
                             "Cancelamentos": num_br(bloco["Cancelamentos"].sum(), 0),
                             "Ativos no fim": num_br(bloco["Ativos"].iloc[-1], 0), "Receita reconhecida": brl(r_),
                             "Cobranças (caixa)": brl(c_), "Receita antecipada acumulada": brl(res["anos"][k]["receita_diferida"]),
                             "Custo de entrega": brl(ce), "Taxas": brl(tx), "Margem bruta": pct(div(r_ - ce - tx, r_))})
            st.dataframe(pd.DataFrame(anos), hide_index=True, width="stretch")
            with st.expander("Ver a projeção mês a mês"):
                vis = s.copy()
                for c in ("Novos", "Cancelamentos", "Ativos"):
                    vis[c] = vis[c].map(lambda x: num_br(x, 1))
                for c in ("Receita reconhecida", "Cobranças (caixa)", "Custo de entrega", "Taxas de cobrança"):
                    vis[c] = vis[c].map(brl)
                st.dataframe(vis, hide_index=True, width="stretch", height=420)
            st.caption("Assinantes aparecem com casas decimais porque o churn é uma média: 2,5 assinantes perdidos significa "
                       "\"entre 2 e 3\". A receita das assinaturas entra automaticamente na DRE, no fluxo de caixa e nos indicadores.")
    with t3:
        if not rc.get("tem"):
            st.info("Cadastre planos com preço e novos assinantes para ver os indicadores.")
        else:
            k = st.columns(4)
            k[0].metric("MRR no Mês 12", brl(rc["mrr12"]), f"Mês 1: {brl(rc['mrr1'])}", delta_color="off",
                        help="Receita recorrente mensal: soma do valor mensal equivalente de todos os assinantes ativos.")
            k[1].metric("ARR (MRR x 12)", brl(rc["arr12"]), help="Receita recorrente anualizada no ritmo do Mês 12.")
            k[2].metric("Assinantes no Mês 12", num_br(rc["ativos12"], 0))
            k[3].metric("ARPU (receita média por assinante)", brl(rc["arpu"]), help="MRR ÷ assinantes ativos no Mês 12.")
            k = st.columns(4)
            k[0].metric("Churn mensal médio", pct(rc["churn"]), f"vida média ≈ {num_br(rc['vida_meses'], 0)} meses",
                        delta_color="off", help="Churn de cada plano ponderado pelos assinantes do Mês 12. Vida média = 1 ÷ churn.")
            k[1].metric("Margem bruta recorrente (Ano 1)", pct(rc["margem_bruta"]),
                        help="(Receita reconhecida − custo de entrega − taxas de cobrança) ÷ receita reconhecida.")
            k[2].metric("LTV (lucro bruto por assinante na vida)", brl(rc["ltv"]) if rc["ltv"] == rc["ltv"] else "n.a.",
                        help="ARPU × margem bruta ÷ churn.")
            k[3].metric("CAC (custo para conquistar 1 assinante)", brl(rc["cac"]) if rc["cac"] == rc["cac"] else "n.a.",
                        help=f"Gasto de aquisição do Mês 1 ({brl(rc['gasto_aquisicao_m1'])}, {rc.get('origem_cac', '')}) ÷ "
                             f"{num_br(rc['novos1'], 0)} novos assinantes no Mês 1.")
            lc, pb = rc["ltv_cac"], rc["payback_cac"]
            k = st.columns(3)
            k[0].metric("LTV/CAC", f"{num_br(lc, 1)}x" if lc == lc else "n.a.",
                        help="Referência: 3x ou mais. Abaixo de 1x, cada cliente novo dá prejuízo.")
            k[1].metric("Payback do CAC", f"{num_br(pb, 1)} meses" if pb == pb else "n.a.",
                        help="Meses de lucro bruto de um assinante para pagar o custo de conquistá-lo. Referência: até 12 meses.")
            k[2].metric("Receita antecipada no fim do Ano 1", brl(rc["receita_diferida_a1"]),
                        help="Dinheiro de planos já cobrados e ainda não prestados. É caixa, mas também é dívida com o cliente.")
            st.markdown("#### Leitura rápida")
            leit = []
            leit.append(("🟢" if lc == lc and lc >= 3 else ("🟡" if lc == lc and lc >= 1 else "🔴"),
                         f"LTV/CAC de {num_br(lc, 1)}x: " + ("a aquisição de clientes se paga com folga." if lc == lc and lc >= 3 else
                                                           "a conquista de clientes está cara para o que eles deixam de lucro.")))
            leit.append(("🟢" if rc["churn"] <= 0.03 else ("🟡" if rc["churn"] <= 0.07 else "🔴"),
                         f"Churn de {pct(rc['churn'])} ao mês: em um ano, cerca de {pct(1 - (1 - rc['churn']) ** 12, 0)} da base "
                         "seria substituída só para manter o tamanho."))
            leit.append(("🟢" if rc["margem_bruta"] >= 0.7 else ("🟡" if rc["margem_bruta"] >= 0.5 else "🔴"),
                         f"Margem bruta de {pct(rc['margem_bruta'])}: empresas de software costumam operar entre 70% e 85%."))
            for ic, tx in leit:
                st.markdown(f"{ic} {tx}")
    with t4:
        st.markdown("""
Classifique cada gasto de tecnologia pelo jeito como ele cresce:

| Tipo de gasto | Exemplos | Onde lançar no programa |
|---|---|---|
| **Cresce com cada assinante** | servidor/nuvem por uso, banco de dados, armazenamento, envio de SMS e e-mail, APIs pagas por chamada, uso de inteligência artificial, suporte ao cliente | Coluna **Custo de entrega por assinante** (aba Planos) |
| **% de cada cobrança** | meio de pagamento (cartão, Pix, boleto), comissão das lojas de aplicativos sobre assinaturas vendidas dentro do app | Coluna **Taxa de cobrança** (aba Planos) |
| **Fixo por mês** | infraestrutura mínima de nuvem, licenças de ferramentas (repositório de código, design, monitoramento), domínio, contas de desenvolvedor | 🧾 Custos e despesas → Custos fixos, classificação **Custo de tecnologia / infraestrutura** |
| **Equipe técnica** | desenvolvedores, suporte, DevOps | 🧾 Custos e despesas → Mão de obra, área **Produção / operação** |
| **Antes de abrir** | desenvolvimento da primeira versão (MVP), identidade visual, registro de marca | 🏗️ Investimentos → Pré-operacionais |
| **Equipamentos** | computadores e celulares para teste | 🏗️ Investimentos → Investimentos fixos (com depreciação) |

As lojas de aplicativos cobram comissão sobre compras e assinaturas feitas dentro do app, e essas regras vêm mudando
por decisões judiciais e regulatórias em vários países. Consulte as condições vigentes antes de fechar o plano. Vender para
empresas por boleto ou Pix, fora da loja, costuma reduzir a taxa de cobrança.
""")
        cf = d["custos_fixos"]
        tec = cf[cf["classificacao"].astype(str).str.startswith("Custo de tecnologia")]
        st.markdown(f"**Custos fixos de tecnologia cadastrados:** {brl(tec['valor_mensal'].sum())}/mês")
        if not tec.empty:
            st.dataframe(tec[["descricao", "valor_mensal"]].rename(columns={"descricao": "Item", "valor_mensal": "R$/mês"})
                         .assign(**{"R$/mês": lambda x: x["R$/mês"].map(brl)}), hide_index=True, width="stretch")
        if st.button("Lançar custos fixos de tecnologia →"):
            ir("custos", "Na aba Custos fixos, use a classificação \"Custo de tecnologia / infraestrutura\".")
            st.rerun()
    with t5:
        st.markdown(r"""
#### Faturamento não é receita
Um cliente que paga **R$ 3.000 por um plano anual** em janeiro:
- **Caixa:** entra R$ 3.000 em janeiro.
- **Receita (DRE):** R$ 250 por mês, de janeiro a dezembro, porque o serviço é prestado ao longo do ano.
- **Balanço:** em janeiro sobra R$ 2.750 de **receita antecipada**, uma obrigação com o cliente que diminui a cada mês.

#### Erro comum em planos de negócio
Lançar os planos trimestrais, semestrais e anuais como se fossem **vendidos de novo todo mês** (por exemplo, "8 planos
anuais de R$ 4.200 por mês" somando R$ 33.600 de receita mensal). Isso multiplica a receita: o plano anual é cobrado
uma vez por ano, e quem já assinou não paga de novo no mês seguinte. O certo é projetar **novos assinantes**,
**cancelamentos** e **renovações**, como esta ferramenta faz.

#### Os indicadores
- **MRR** (receita recorrente mensal) = soma dos valores mensais equivalentes dos assinantes ativos. **ARR** = MRR × 12.
- **Churn** = % de assinantes que cancelam por mês. Vida média do cliente ≈ 1 ÷ churn.
- **ARPU** = MRR ÷ assinantes ativos.
- **Margem bruta** = (receita − custo de entrega − taxas) ÷ receita.
- **LTV** (valor do cliente na vida) = ARPU × margem bruta ÷ churn.
- **CAC** (custo de aquisição) = gastos de marketing e vendas ÷ novos assinantes.
- **LTV/CAC**: referência de 3x ou mais. **Payback do CAC** = CAC ÷ (ARPU × margem bruta), idealmente até 12 meses.

#### Como o programa calcula
Cada mês de entrada forma uma **safra**. A cada renovação, a safra perde a parte que cancela. Quem paga por ciclo
(trimestral, anual) só sai na renovação: o churn mensal informado é convertido para o ciclo,
1 − (1 − churn)^meses do ciclo. A partir do Ano 2, a entrada de novos assinantes parte do Mês 12 e cresce à taxa anual de
crescimento das premissas.
""")
    proxima("assinaturas")

# ##############################################################################
# PARTE: pnbox_ui_paginas5.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: IDENTIDADE E OBJETIVOS SMART, ANÁLISE ESTRATÉGICA E REFORMA TRIBUTÁRIA
# ==============================================================================

VALORES_SUGERIDOS = ["Ética", "Transparência", "Qualidade", "Inovação", "Respeito ao cliente", "Responsabilidade social",
                     "Sustentabilidade", "Valorização das pessoas", "Compromisso com a região", "Simplicidade",
                     "Excelência no atendimento", "Colaboração", "Agilidade", "Confiança"]


def recarregar_tabela(tabela: str, df: pd.DataFrame, chave_: str):
    """Grava uma tabela alterada por um botão e força o editor a reabrir com os dados novos."""
    salvar(conn, tabela, df, chave_)
    d[tabela] = limpar(df[df[chave_].astype(str).str.strip() != ""].reset_index(drop=True), tabela)
    ss.buf.pop((PID, tabela, ""), None)
    ss.hash.pop((PID, tabela, ""), None)
    _nova_versao(tabela)
    invalidar()


def _selos_smart(chk: dict) -> str:
    return " ".join(f"{'✅' if ok else '❌'} {letra}" for letra, (ok, _) in chk.items())


# ------------------------------------------------------------------------------
# MISSÃO, VISÃO, VALORES E OBJETIVOS SMART
# ------------------------------------------------------------------------------
def pg_identidade():
    cabecalho("identidade")
    t1, t2, t3 = st.tabs(["🧭 Missão, visão e valores", "🎯 Objetivos SMART", "📚 Como escrever"])
    with t1:
        v = {}
        v["missao"] = st.text_area("Missão: por que a empresa existe?", proj["missao"], height=110, key=wk("missao"),
                                   placeholder="Oferecer [o quê] para [quem], por meio de [como], gerando [qual benefício].",
                                   help="Curta, no presente, centrada no cliente e no benefício que a empresa entrega.")
        v["visao"] = st.text_area("Visão: aonde a empresa quer chegar?", proj["visao"], height=90, key=wk("visao"),
                                  placeholder="Ser reconhecida como [posição desejada] em [mercado/região] até [ano].",
                                  help="Futuro desejado, inspirador e com horizonte de tempo.")
        atuais = [x.strip() for x in str(proj["valores"]).split(";") if x.strip()]
        escolhidos = st.multiselect("Valores: em que a empresa acredita e como se comporta?",
                                    list(dict.fromkeys(VALORES_SUGERIDOS + atuais)), atuais, key=wk("valores"),
                                    accept_new_options=True, help="Escolha de 3 a 6 e acrescente os seus.")
        v["valores"] = "; ".join(escolhidos)
        auto_unica("projeto", v)
        salvo_txt()
        dicas = []
        if v["missao"] and len(v["missao"].split()) > 45:
            dicas.append("A missão está longa: tente caber em até 2 linhas.")
        if v["visao"] and not re.search(r"\b(19|20)\d{2}\b|\banos?\b", v["visao"]):
            dicas.append("A visão não tem prazo: inclua um ano ou horizonte (ex.: \"até 2030\").")
        if escolhidos and not 3 <= len(escolhidos) <= 6:
            dicas.append("Recomenda-se de 3 a 6 valores, para que todos sejam lembrados e praticados.")
        for dc in dicas:
            st.warning(dc)
    with t2:
        st.caption("Cada objetivo precisa ser **S**pecífico, **M**ensurável, **A**tingível, **R**elevante e **T**emporal. "
                   "A **perspectiva** liga o objetivo ao Balanced Scorecard, e **contribui para** desenha o mapa estratégico.")
        nomes_obj = [x for x in d["objetivos"]["objetivo"].astype(str) if x.strip()]
        editor("objetivos", {
            "objetivo": T("Objetivo (verbo + o que muda)", required=True, width="large"),
            "perspectiva": S("Perspectiva (BSC)", options=PERSPECTIVAS),
            "indicador": T("Indicador"), "meta": N("Meta", step=1.0), "unidade": T("Unidade"),
            "prazo_meses": N("Prazo (mês do plano)", min_value=0, max_value=120, step=1),
            "responsavel": T("Responsável"),
            "por_que_atingivel": T("Por que é atingível?", width="medium"),
            "relevancia": T("Por que é relevante?", width="medium"),
            "iniciativa": T("Iniciativa / plano de ação", width="medium"),
            "contribui_para": S("Contribui para (objetivo)", options=nomes_obj)}, "objetivo")
        c1, c2 = st.columns([1, 2])
        if c1.button("✨ Sugerir objetivos a partir do plano", width="stretch"):
            sug = pd.DataFrame(objetivos_sugeridos(proj, d, R()))
            existentes = set(d["objetivos"]["objetivo"].astype(str))
            sug = sug[~sug["objetivo"].isin(existentes)]
            recarregar_tabela("objetivos", pd.concat([d["objetivos"], sug], ignore_index=True), "objetivo")
            st.rerun()
        c2.caption("Os rascunhos usam os números do plano (receita, ponto de equilíbrio, churn, horas da equipe). "
                   "Revise e reescreva com as palavras do grupo.")
        ob = d["objetivos"]
        if not ob.empty:
            ok_total = 0
            st.markdown("#### Verificação SMART")
            for _, o in ob.iterrows():
                chk = checar_smart(o)
                n_ok = sum(1 for x in chk.values() if x[0])
                ok_total += n_ok == 5
                with st.container(border=True):
                    st.markdown(f"**{o['objetivo']}** · {o['perspectiva'] or 'sem perspectiva'} · {_selos_smart(chk)}")
                    if n_ok == 5:
                        st.caption("✍️ " + frase_smart(o))
                    for letra, (ok, dica) in chk.items():
                        if not ok:
                            st.caption(f"❌ **{SMART[letra][0]}:** {dica}")
            st.metric("Objetivos que cumprem os 5 critérios", f"{ok_total} de {len(ob)}")
    with t3:
        st.markdown("""
**Missão** responde *por que existimos*: o que a empresa faz, para quem e qual benefício entrega. Evite frases genéricas
("ser a melhor empresa"). Exemplo: *"Oferecer refeições saudáveis e saborosas a quem tem pouco tempo, com entrega rápida e
preço justo."*

**Visão** responde *aonde queremos chegar*: um futuro desafiador, com prazo. Exemplo: *"Ser a principal referência em
alimentação saudável de Teófilo Otoni até 2030."*

**Valores** são os princípios que orientam as decisões e o comportamento de todos na empresa.

**Objetivos SMART** transformam a visão em metas acompanháveis:

| Letra | Critério | Pergunta de checagem |
|---|---|---|
| S | Específico | Está claro *o que* muda? Tem verbo de ação? |
| M | Mensurável | Qual indicador e qual número provam que foi alcançado? |
| A | Atingível | Os recursos do plano (equipe, verba, capacidade) permitem chegar lá? |
| R | Relevante | Ajuda a cumprir a estratégia e a viabilidade do negócio? |
| T | Temporal | Até quando? |

Exemplo ruim: *"Melhorar as vendas."* Exemplo SMART: *"Aumentar a receita mensal para R$ 45.000 até o mês 12, com a
abertura do canal de delivery, sob responsabilidade da sócia comercial."*
""")
    proxima("identidade")


# ------------------------------------------------------------------------------
# BUSINESS MODEL CANVAS
# ------------------------------------------------------------------------------
ORIGEM_CANVAS = {
    "Parcerias-chave": "campo abaixo ou fornecedores críticos", "Atividades-chave": "processos do Plano operacional",
    "Recursos-chave": "investimentos fixos e equipe", "Proposta de valor": "Propostas de valor",
    "Relacionamento com clientes": "campo abaixo ou oportunidades da Jornada (pós-venda)", "Canais": "Canais e funil",
    "Segmentos de clientes": "Segmentação (os mais bem avaliados)", "Estrutura de custos": "DRE do Ano 1",
    "Fontes de receita": "Produtos, serviços e assinaturas",
}


# ------------------------------------------------------------------------------
# DESENHOS DAS MATRIZES (HTML/SVG responsivos)
# ------------------------------------------------------------------------------

CORES_GE = {"Seguro e crescimento": ("#3fd0c0", "🛡️"), "Seletivo ou cauteloso": ("#ffd23f", "➤"),
            "Expansão limitada ou colheita": ("#ffa630", "↗"), "Perigo ou colheita": ("#ec4561", "⚠️")}
# linhas de cima para baixo (atratividade alta -> baixa); colunas da esquerda para a direita (força alta -> baixa)
GRADE_GE = [["Seguro e crescimento", "Seguro e crescimento", "Seletivo ou cauteloso"],
            ["Seguro e crescimento", "Seletivo ou cauteloso", "Expansão limitada ou colheita"],
            ["Seletivo ou cauteloso", "Expansão limitada ou colheita", "Perigo ou colheita"]]


def _esc(s) -> str:
    return _html.escape(str(s))


def html_ge(tab: pd.DataFrame) -> str:
    celulas = "".join(
        f"<div class='ge-c' style='background:{CORES_GE[n][0]}'><div class='ge-i'>{CORES_GE[n][1]}</div><div>{n.replace(' ou ', ' ou<br>')}</div></div>"
        for linha in GRADE_GE for n in linha)
    pts = ""
    for k, (_, r) in enumerate(tab.iterrows(), start=1):
        left = min(max((10 - float(r["Força competitiva"])) / 10 * 100, 3), 97)
        top = min(max((10 - float(r["Atratividade"])) / 10 * 100, 3), 97)
        pts += (f"<div class='ge-p' style='left:{left:.1f}%;top:{top:.1f}%' title='{_esc(r['Segmento'])}: atratividade "
                f"{num_br(r['Atratividade'], 1)}, força {num_br(r['Força competitiva'], 1)}'>{k}</div>")
    leg = "".join(f"<li><b>{k}.</b> {_esc(r['Segmento'])} <span>→ {_esc(r['Recomendação'])}</span></li>"
                  for k, (_, r) in enumerate(tab.iterrows(), start=1))
    css = """<style>
    .ge-w{display:flex;gap:10px;align-items:stretch;max-width:760px;margin:auto}
    .ge-y{writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;color:#7a8595;font-size:.85rem;border-left:2px solid #7a8595;padding-left:6px}
    .ge-m{flex:1;position:relative}
    .ge-g{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;aspect-ratio:1.15/1}
    .ge-c{border-radius:14px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;
          font-weight:600;font-size:clamp(.62rem,1.6vw,.98rem);color:#253041;padding:4px;line-height:1.2}
    .ge-i{font-size:1.5em;margin-bottom:4px;opacity:.75}
    .ge-p{position:absolute;transform:translate(-50%,-50%);width:34px;height:34px;border-radius:50%;background:#1f3f75;color:#fff;
          border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;font-weight:700}
    .ge-x{text-align:center;color:#7a8595;font-size:.85rem;border-top:2px solid #7a8595;margin-top:8px;padding-top:4px}
    .ge-l{max-width:760px;margin:10px auto 0;padding-left:20px} .ge-l li{margin:2px 0} .ge-l span{opacity:.7}
    </style>"""
    return (css + "<div class='ge-w'><div class='ge-y'>Baixa ← Atratividade do mercado → Alta</div><div class='ge-m'>"
            f"<div class='ge-g'>{celulas}</div>{pts}<div class='ge-x'>Alta ← Força competitiva da empresa no segmento → Baixa</div>"
            f"</div></div><ol class='ge-l' style='list-style:none'>{leg}</ol>")


QUAD_BCG_VIS = {"Estrela": ("#ede3fb", "#7b3fe4", "⭐", ["Mercado atrativo", "Gera bom faturamento", "Exige investimento"]),
                "Interrogação": ("#fff4c7", "#c99700", "❓", ["Ainda gera pouco retorno", "Expectativa de crescimento", "Decisão de risco"]),
                "Vaca leiteira": ("#eefbd0", "#5e8f00", "🐄", ["Gera muito caixa", "Exige pouco investimento", "Recursos realocáveis"]),
                "Abacaxi": ("#fde2e5", "#c8102e", "🍍", ["Gera pouco caixa", "Mercado sem crescimento", "Provável saída"])}


def html_bcg(tab: pd.DataFrame, lim: float) -> str:
    import math
    ymax = max(2 * lim, 1.0)
    quad = ""
    for q in ("Estrela", "Interrogação", "Vaca leiteira", "Abacaxi"):
        bg, cor, ic, itens = QUAD_BCG_VIS[q]
        al = "flex-end" if q in ("Vaca leiteira", "Abacaxi") else "flex-start"
        quad += (f"<div class='bc-q' style='background:{bg};justify-content:{al}'><div class='bc-t' style='color:{cor}'>{ic} {q}</div>"
                 + "".join(f"<div class='bc-i'>✓ {i}</div>" for i in itens) + "</div>")
    mx = max(float(tab["Receita anual"].max()), 1.0)
    pts = ""
    for _, r in tab.iterrows():
        p = min(max(float(r["Participação relativa"]), 0.1), 10)
        left = (1 - math.log10(p)) / 2 * 100
        top = (1 - min(max(float(r["Crescimento do mercado (%)"]), 0), ymax) / ymax) * 100
        tam = 26 + 50 * math.sqrt(max(float(r["Receita anual"]), 0) / mx)
        cor = QUAD_BCG_VIS[r["Quadrante"]][1]
        pts += (f"<div class='bc-p' style='left:{min(max(left, 4), 96):.1f}%;top:{min(max(top, 5), 95):.1f}%;width:{tam:.0f}px;"
                f"height:{tam:.0f}px;background:{cor}' title='{_esc(r['Produto'])}: crescimento {num_br(r['Crescimento do mercado (%)'], 1)}%, "
                f"participação {num_br(r['Participação relativa'], 2)}x, receita {brl(r['Receita anual'])}'></div>"
                f"<div class='bc-r' style='left:{min(max(left, 8), 92):.1f}%;top:calc({min(max(top, 5), 95):.1f}% + {tam / 2 + 2:.0f}px)'>"
                f"{_esc(str(r['Produto'])[:26])}</div>")
    yt = "".join(f"<span style='top:{(1 - v / ymax) * 100:.0f}%'>{num_br(v, 0)}%</span>" for v in (0, lim / 2, lim, 1.5 * lim, ymax))
    xt = "".join(f"<span style='left:{(1 - math.log10(v)) / 2 * 100:.0f}%'>{num_br(v, 1 if v < 1 else 0)}x</span>" for v in (10, 5, 1, 0.5, 0.1))
    css = """<style>
    .bc-w{display:flex;gap:6px;max-width:780px;margin:auto}
    .bc-ya{writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;font-weight:700;color:#334;font-size:.85rem}
    .bc-yt{position:relative;width:34px;font-size:.72rem;color:#556} .bc-yt span{position:absolute;right:2px;transform:translateY(-50%)}
    .bc-m{flex:1;position:relative}
    .bc-g{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;aspect-ratio:1.2/1;border:2px solid #2d4f8a;border-radius:6px;overflow:hidden}
    .bc-q{display:flex;flex-direction:column;padding:10px 14px;border:1px solid rgba(45,79,138,.35);font-size:clamp(.62rem,1.4vw,.85rem);color:#253041}
    .bc-t{font-weight:800;font-size:1.25em;margin-bottom:4px} .bc-i{opacity:.8;line-height:1.35}
    .bc-p{position:absolute;transform:translate(-50%,-50%);border-radius:50%;opacity:.82;border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.3)}
    .bc-r{position:absolute;transform:translateX(-50%);font-size:.72rem;font-weight:600;background:rgba(255,255,255,.85);
          padding:0 4px;border-radius:4px;white-space:nowrap;color:#223}
    .bc-xt{position:relative;height:18px;font-size:.72rem;color:#556} .bc-xt span{position:absolute;transform:translateX(-50%)}
    .bc-xa{text-align:center;font-weight:700;color:#334;font-size:.85rem}
    @media (max-width:640px){.bc-i{display:none}.bc-r{font-size:.62rem}.bc-ya,.bc-xa{font-size:.72rem}}
    </style>"""
    return (css + "<div class='bc-w'><div class='bc-ya'>Baixo ← Crescimento do mercado (% a.a.) → Alto</div>"
            f"<div class='bc-yt'>{yt}</div><div class='bc-m'><div class='bc-g'>{quad}</div>{pts}<div class='bc-xt'>{xt}</div>"
            "<div class='bc-xa'>Alta ← Participação relativa de mercado (x líder) → Baixa</div></div></div>")


CICLO_BSC = [("Aprendizado e crescimento", "#0d6fb8", "Conhecimento e habilidades da equipe sustentam inovações e melhorias."),
             ("Processos internos", "#123a5e", "Equipe capacitada melhora os processos de trabalho."),
             ("Clientes", "#d7262e", "Processos melhores conduzem a clientes mais satisfeitos."),
             ("Financeira", "#b5191f", "Clientes satisfeitos levam a melhores resultados financeiros.")]


def svg_bsc(ob: pd.DataFrame) -> str:
    """Ciclo do BSC: anel com quatro segmentos em forma de seta, no sentido horário."""
    import math
    cx, cy, R1, R0 = 300, 260, 245, 120
    Rm = (R1 + R0) / 2

    def pt(r, a):
        return cx + r * math.cos(math.radians(a)), cy + r * math.sin(math.radians(a))

    partes = []
    inicio = [-90, 0, 90, 180]
    for (nome, cor, desc), a0 in zip(CICLO_BSC, inicio):
        a1 = a0 + 88
        ponta = a1 + 7
        o0, o1 = pt(R1, a0), pt(R1, a1)
        i0, i1 = pt(R0, a0), pt(R0, a1)
        tip, notch = pt(Rm, ponta), pt(Rm, a0 + 7)
        d = (f"M{o0[0]:.1f},{o0[1]:.1f} A{R1},{R1} 0 0 1 {o1[0]:.1f},{o1[1]:.1f} L{tip[0]:.1f},{tip[1]:.1f} "
             f"L{i1[0]:.1f},{i1[1]:.1f} A{R0},{R0} 0 0 0 {i0[0]:.1f},{i0[1]:.1f} L{notch[0]:.1f},{notch[1]:.1f} Z")
        partes.append(f"<path d='{d}' fill='{cor}'/>")
        tx, ty = pt(Rm - 2, a0 + 51)
        n = int((ob["perspectiva"] == nome).sum()) if not ob.empty else 0
        palavras = desc.split()
        linhas, atual = [], ""
        for w in palavras:
            if len(atual) + len(w) > 19:
                linhas.append(atual)
                atual = w
            else:
                atual = (atual + " " + w).strip()
        linhas.append(atual)
        titulo = {"Aprendizado e crescimento": ["APRENDIZADO", "E CRESCIMENTO"], "Processos internos": ["PROCESSOS", "INTERNOS"]}.get(
            nome, [nome.upper()])
        y0 = ty - (len(titulo) * 15 + len(linhas) * 14 + 16) / 2 + 12
        txt = "".join(f"<tspan x='{tx:.0f}' y='{y0 + i * 15:.0f}' font-weight='800' font-size='11.5'>{_esc(t)}</tspan>" for i, t in enumerate(titulo))
        y1 = y0 + len(titulo) * 15 + 2
        txt += "".join(f"<tspan x='{tx:.0f}' y='{y1 + i * 14:.0f}' font-size='11'>{_esc(t)}</tspan>" for i, t in enumerate(linhas))
        txt += (f"<tspan x='{tx:.0f}' y='{y1 + len(linhas) * 14 + 4:.0f}' font-size='11.5' font-weight='700'>"
                f"{n} objetivo{'s' if n != 1 else ''}</tspan>")
        partes.append(f"<text fill='white' text-anchor='middle' font-family='Helvetica,Arial,sans-serif'>{txt}</text>")
    centro = (f"<circle cx='{cx}' cy='{cy}' r='{R0 - 8}' fill='white'/>"
              f"<text x='{cx}' y='{cy + 26}' text-anchor='middle' font-family='Helvetica,Arial,sans-serif' font-weight='900' "
              f"font-size='74' fill='#333'>BSC</text>")
    return (f"<svg viewBox='0 0 600 520' style='width:100%;max-width:600px;display:block;margin:auto' "
            f"xmlns='http://www.w3.org/2000/svg'>{''.join(partes)}{centro}</svg>")


def html_canvas(bl: dict) -> str:
    def caixa(nome, area, cor):
        itens = "".join(f"<li>{_esc(x)}</li>" for x in bl.get(nome, [])) or "<li><i>preencha a ferramenta de origem</i></li>"
        return (f"<div class='cv-b' style='grid-area:{area};background:{cor}'><div class='cv-t'>{nome.upper()}</div>"
                f"<ul>{itens}</ul><div class='cv-o'>vem de: {ORIGEM_CANVAS[nome]}</div></div>")
    css = """<style>
    .cv{display:grid;gap:8px;grid-template-columns:repeat(10,1fr);grid-template-areas:
      "pa pa ak ak pv pv rc rc sc sc" "pa pa kr kr pv pv ch ch sc sc" "cs cs cs cs cs rs rs rs rs rs";}
    .cv-b{border-radius:12px;padding:14px 14px;font-size:.84rem;min-height:150px;color:#fff}
    .cv-t{font-weight:800;margin-bottom:6px;letter-spacing:.02em;font-size:.95rem;line-height:1.15}
    .cv-b ul{margin:0;padding-left:18px} .cv-b li{margin:2px 0} .cv-o{font-size:.7rem;opacity:.75;margin-top:8px}
    @media (max-width:700px){.cv{grid-template-columns:1fr;grid-template-areas:"sc" "pv" "ch" "rc" "rs" "kr" "ak" "pa" "cs";}}
    </style>"""
    return css + "<div class='cv'>" + "".join([
        caixa("Parcerias-chave", "pa", "#6a3184"), caixa("Atividades-chave", "ak", "#f0901f"),
        caixa("Recursos-chave", "kr", "#e5653e"), caixa("Proposta de valor", "pv", "#44923a"),
        caixa("Relacionamento com clientes", "rc", "#8db833"), caixa("Canais", "ch", "#14a49a"),
        caixa("Segmentos de clientes", "sc", "#1592c4"), caixa("Estrutura de custos", "cs", "#4456a0"),
        caixa("Fontes de receita", "rs", "#22617a")]) + "</div>"


def pg_canvas():
    cabecalho("canvas")
    st.caption("O quadro é montado automaticamente com o que você já preencheu. Complete os dois blocos que não têm "
               "ferramenta própria (um item por linha).")
    c1, c2 = st.columns(2)
    v = {"canvas_relacionamento": c1.text_area("Relacionamento com clientes", proj["canvas_relacionamento"], height=110,
                                               key=wk("cv_rel"), placeholder="Atendimento pelo WhatsApp\nCartão fidelidade\nComunidade no Instagram"),
         "canvas_parcerias": c2.text_area("Parcerias-chave", proj["canvas_parcerias"], height=110, key=wk("cv_par"),
                                          placeholder="Produtores locais\nAplicativo de entregas\nSecretaria de Turismo")}
    auto_unica("projeto", v)
    st.markdown(html_canvas(canvas_blocos(proj, d, R())), unsafe_allow_html=True)
    proxima("canvas")


# ------------------------------------------------------------------------------
# MATRIZ BCG
# ------------------------------------------------------------------------------
def pg_bcg():
    cabecalho("bcg")
    res = R()
    st.info("A matriz BCG foi criada para analisar um **portfólio**. Em um negócio novo, use-a para as linhas de produto: "
            "informe quanto o mercado de cada uma cresce por ano e sua participação **relativa** (sua fatia ÷ fatia do maior "
            "concorrente; 1,0 = empatado com o líder, 0,5 = metade do líder).")
    produtos = [p for p in res["cat"]["descricao"].astype(str) if p.strip()]
    if res.get("recorrencia", {}).get("tem"):
        produtos.append("Assinaturas (todos os planos)")
    lim = st.slider("A partir de quanto o crescimento do mercado é considerado alto? (% ao ano)", 0.0, 30.0,
                    float(proj["bcg_limite_crescimento_pct"] or 10), 0.5, key=wk("bcg_lim"))
    auto_unica("projeto", {"bcg_limite_crescimento_pct": lim})
    editor("bcg", {"produto": S("Produto / linha", options=produtos, required=True),
                   "crescimento_mercado_pct": N("Crescimento do mercado (% ao ano)", step=0.5),
                   "participacao_relativa": N("Participação relativa (x líder)", min_value=0.0, step=0.1)}, "produto")
    if st.button("Incluir todos os produtos com valores iniciais"):
        faltam = [p for p in produtos if p not in set(d["bcg"]["produto"])]
        novos = pd.DataFrame({"produto": faltam, "crescimento_mercado_pct": proj["crescimento_anual_pct"],
                              "participacao_relativa": 0.5})
        recarregar_tabela("bcg", pd.concat([d["bcg"], novos], ignore_index=True), "produto")
        st.rerun()
    tab = bcg_tabela(d, res, lim)
    if tab.empty:
        return proxima("bcg")
    st.markdown(html_bcg(tab, lim), unsafe_allow_html=True)
    st.caption(f"O tamanho da bolha é a receita do Ano 1. A linha horizontal fica em {num_br(lim, 1)}% a.a. (limite escolhido acima) "
               "e a vertical em 1,0x (empatado com o líder). Passe o mouse sobre as bolhas para ver os valores.")
    vis = tab.copy()
    vis["Receita anual"] = vis["Receita anual"].map(brl)
    vis["Margem"] = vis["Margem"].map(pct)
    vis["Quadrante"] = vis["Quadrante"].map(lambda q: f"{QUADRANTES_BCG[q][0]} {q}")
    st.dataframe(vis, hide_index=True, width="stretch")
    for q, (ic, txt_) in QUADRANTES_BCG.items():
        itens = tab.loc[tab["Quadrante"] == q, "Produto"].tolist()
        if itens:
            st.markdown(f"{ic} **{q}** ({', '.join(itens)}): {txt_}")
    proxima("bcg")


# ------------------------------------------------------------------------------
# MATRIZ GE / McKINSEY
# ------------------------------------------------------------------------------
def pg_ge():
    cabecalho("ge")
    st.caption("Cada segmento de clientes é posicionado pela **atratividade do mercado** (média das notas de tamanho, lucro, "
               "longevidade e concorrência da Segmentação) e pela **força competitiva** da empresa nele (automática: média entre "
               "a aderência do produto ao segmento e a nota da empresa frente aos concorrentes; você pode informar a sua).")
    tab = ge_tabela(d)
    if tab.empty:
        st.warning("Cadastre os segmentos em Segmentação de mercado.")
        if st.button("Ir para Segmentação →"):
            ir("segmentos")
            st.rerun()
        return proxima("ge")
    with st.expander("✏️ Informar a força competitiva por segmento (opcional)"):
        editor("ge", {"unidade": S("Segmento", options=d["segmentos"]["nome"].tolist(), required=True),
                      "forca": N("Força competitiva (0 a 10)", min_value=0.0, max_value=10.0, step=0.5),
                      "justificativa": T("Justificativa", width="large")}, "unidade")
        tab = ge_tabela(d)
    st.markdown(html_ge(tab), unsafe_allow_html=True)
    st.write("")
    for n, (cor, ic) in CORES_GE.items():
        if (tab["Recomendação"] == n).any():
            st.markdown(f"{ic} **{n}:** {ACOES_GE[n]}")
    st.dataframe(tab, hide_index=True, width="stretch")
    proxima("ge")


# ------------------------------------------------------------------------------
# CINCO FORÇAS DE PORTER
# ------------------------------------------------------------------------------
def radar_porter(tab: pd.DataFrame, tamanho=(5.2, 5.2)):
    import numpy as np
    vals = tab["Intensidade"].tolist()
    rot = ["Rivalidade", "Novos\nentrantes", "Substitutos", "Clientes", "Fornecedores"]
    ang = np.linspace(0, 2 * np.pi, len(vals), endpoint=False).tolist()
    fig, ax = plt.subplots(figsize=tamanho, subplot_kw={"polar": True})
    ax.plot(ang + ang[:1], vals + vals[:1], color="#2f6fd0", linewidth=2)
    ax.fill(ang + ang[:1], vals + vals[:1], color="#2f6fd0", alpha=0.25)
    ax.set_xticks(ang)
    ax.set_xticklabels(rot, fontsize=10)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=8, color="#777")
    fig.tight_layout()
    return fig


def pg_porter():
    cabecalho("porter")
    sug = porter_sugestoes(proj, d, R())
    if d["porter"].empty:
        ini = pd.DataFrame({"forca": [f for f, _ in FORCAS_PORTER], "intensidade": [sug[f][0] for f, _ in FORCAS_PORTER],
                            "justificativa": [f"(sugestão) {sug[f][1]}" for f, _ in FORCAS_PORTER], "estrategia": ""})
        recarregar_tabela("porter", ini, "forca")
    st.caption("Nota de **1 (fraca)** a **5 (muito forte)**: quanto mais forte a força, mais ela pressiona preços e margens. "
               "O programa sugeriu notas iniciais a partir do plano; revise e justifique com as suas pesquisas.")
    with st.expander("❓ Perguntas que ajudam a avaliar cada força"):
        for f, q in FORCAS_PORTER:
            st.markdown(f"- **{f}:** {q}")
    editor("porter", {"forca": T("Força", disabled=True, width="medium"),
                      "intensidade": N("Intensidade (1 a 5)", min_value=0, max_value=5, step=1),
                      "justificativa": T("Justificativa", width="large"),
                      "estrategia": T("Como o negócio vai se proteger", width="large")}, "forca")
    tab = porter_tabela(d)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.pyplot(radar_porter(tab), width="stretch")
    with c2:
        st.markdown("#### Leitura")
        st.write(porter_leitura(tab))
        st.markdown("#### Sugestões do programa")
        for f, _ in FORCAS_PORTER:
            st.caption(f"**{f}:** nota {sug[f][0]}. {sug[f][1]}")
    proxima("porter")


# ------------------------------------------------------------------------------
# BALANCED SCORECARD E MAPA ESTRATÉGICO
# ------------------------------------------------------------------------------
def dot_mapa(ob: pd.DataFrame) -> str:
    cores = {"Financeira": "#d4efdf", "Clientes": "#fdebd0", "Processos internos": "#d6eaf8", "Aprendizado e crescimento": "#e8daef"}
    def q(s):
        return '"' + str(s).replace('"', "'") + '"'
    linhas = ['digraph G {', 'rankdir=BT; nodesep=0.3; ranksep=0.45;',
              'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11, width=2.2];']
    nomes = {}
    for i, (_, o) in enumerate(ob.iterrows()):
        nomes[o["objetivo"]] = f"o{i}"
        rot = "\\n".join(re.findall(r".{1,28}(?:\s|$)", str(o["objetivo"]))[:4]) or str(o["objetivo"])
        linhas.append(f'o{i} [label={q(rot)}, fillcolor="{cores.get(o["perspectiva"], "#eeeeee")}"];')
    for p in PERSPECTIVAS:
        ids = [nomes[o] for o in ob.loc[ob["perspectiva"] == p, "objetivo"]]
        if ids:
            linhas.append("{rank=same; " + "; ".join(ids) + "}")
    tem_lig = False
    for _, o in ob.iterrows():
        alvo = str(o["contribui_para"] or "").strip()
        if alvo in nomes and alvo != o["objetivo"]:
            linhas.append(f'{nomes[o["objetivo"]]} -> {nomes[alvo]};')
            tem_lig = True
    if not tem_lig:  # ligações automáticas: cada perspectiva sustenta a de cima
        for a, b in zip(PERSPECTIVAS[::-1][:-1], PERSPECTIVAS[::-1][1:]):
            de = [nomes[o] for o in ob.loc[ob["perspectiva"] == a, "objetivo"]]
            para = [nomes[o] for o in ob.loc[ob["perspectiva"] == b, "objetivo"]]
            for x in de:
                for y in para[:1]:
                    linhas.append(f"{x} -> {y} [style=dashed];")
    linhas.append("}")
    return "\n".join(linhas)


def pg_bsc():
    cabecalho("bsc")
    res = R()
    ob = d["objetivos"]
    ind = bsc_indicadores_plano(proj, d, res)
    st.caption("O BSC organiza os objetivos SMART em quatro perspectivas e mostra como o aprendizado sustenta os processos, "
               "que sustentam os clientes, que geram o resultado financeiro. Os objetivos vêm de 🧭 Missão, visão, valores e objetivos.")
    if ob.empty:
        st.warning("Ainda não há objetivos cadastrados.")
        if st.button("Cadastrar objetivos SMART →"):
            ir("identidade", "Na aba Objetivos SMART, use o botão de sugestões ou cadastre os seus, escolhendo a perspectiva de cada um.")
            st.rerun()
    st.markdown(svg_bsc(ob), unsafe_allow_html=True)
    cols = st.columns(2)
    for i, p in enumerate(PERSPECTIVAS[::-1][::-1]):
        with cols[i % 2].container(border=True):
            st.markdown(f"#### {['💰', '🤝', '⚙️', '🌱'][i]} {p}")
            sub = ob[ob["perspectiva"] == p]
            if sub.empty:
                st.caption("Nenhum objetivo nesta perspectiva.")
            for _, o in sub.iterrows():
                chk = checar_smart(o)
                st.markdown(f"**{o['objetivo']}**  \n{o['indicador']}: meta {num_br(o['meta'], 1)} {o['unidade']} até o mês "
                            f"{int(o['prazo_meses'] or 0)} · {_selos_smart(chk)}  \n*Iniciativa:* {o['iniciativa'] or '-'}")
            if ind.get(p):
                st.caption("Indicadores do plano: " + " · ".join(f"{n} = {v}" for n, v in ind[p]))
    if not ob.empty:
        st.markdown("#### 🗺️ Mapa estratégico")
        st.graphviz_chart(dot_mapa(ob), width="stretch")
        if not ob["contribui_para"].astype(str).str.strip().ne("").any():
            st.caption("Ligações tracejadas são automáticas. Para desenhar as suas, preencha \"Contribui para\" nos objetivos.")
    faltam = [p for p in PERSPECTIVAS if p not in set(ob["perspectiva"])]
    if faltam and not ob.empty:
        st.info("Perspectivas sem objetivo: " + ", ".join(faltam) + ".")
    proxima("bsc")


# ------------------------------------------------------------------------------
# REFORMA TRIBUTÁRIA
# ------------------------------------------------------------------------------
def aliquotas_padrao_reforma(ano: int) -> tuple:
    """CBS: referência de 11,3% calculada pelo TCU para 2027 (sujeita ao Senado). IBS: 0,1% em 2027-2028; de 2029 a 2032
    cresce 10% ao ano da alíquota cheia estimada (17,7%) e fica integral em 2033. Valores editáveis."""
    if ano <= 2028:
        return 11.3, 0.1
    if ano >= 2033:
        return 11.3, 17.7
    return 11.3, round(17.7 * 0.1 * (ano - 2028), 2)


def pg_reforma():
    cabecalho("reforma")
    res = R()
    a1 = res["anos"][1]
    st.markdown("""
<div class="pn-card-info">A partir de 2027, a empresa do Simples pode escolher entre continuar pagando tudo no <b>DAS</b>
("Simples puro") ou recolher <b>IBS e CBS por fora</b>, no regime regular ("Simples híbrido"). No híbrido ela paga IVA com
crédito sobre as compras, e os <b>clientes empresas</b> aproveitam crédito integral, o que torna o preço mais competitivo
em vendas B2B. Esta simulação é didática e simplificada: confirme a decisão com um contador.</div>""", unsafe_allow_html=True)
    st.write("")
    c1, c2, c3 = st.columns(3)
    ano = int(c1.select_slider("Ano simulado", list(range(2027, 2034)), int(proj["reforma_ano"] or 2027), key=wk("rf_ano")))
    cbs0, ibs0 = aliquotas_padrao_reforma(ano)
    anexos = list(TABELAS_SIMPLES)
    an0 = proj["anexo_simples"] if proj["anexo_simples"] in anexos else (
        "Anexo III - Serviços" if str(proj["tipo_negocio"]).startswith(("Serviço", "Assinatura")) else "Anexo I - Comércio")
    anexo = c2.selectbox("Anexo do Simples", anexos, index=anexos.index(an0), key=wk("rf_anexo"))
    b2b = c3.slider("Vendas para empresas que aproveitam crédito (%)", 0, 100, int(proj["reforma_b2b_pct"] or 0), 5, key=wk("rf_b2b"))
    c1, c2, c3 = st.columns(3)
    cbs = c1.number_input("CBS (%)", 0.0, 30.0, cbs0, 0.1, key=wk(f"rf_cbs_{ano}"))
    ibs = c2.number_input("IBS (%)", 0.0, 30.0, ibs0, 0.1, key=wk(f"rf_ibs_{ano}"))
    compras_pad = a1["custo_dir"] + a1["fixos"]
    cc = float(proj["reforma_compras_credito"])
    compras = c3.number_input("Compras com direito a crédito no ano (R$)", 0.0, 1e10, compras_pad if cc < 0 else cc, 1000.0,
                              key=wk("rf_compras"), help="Sugestão: custos diretos + custos fixos não salariais do Ano 1. "
                                                         "Salários e pró-labore não geram crédito.")
    auto_unica("projeto", {"reforma_ano": ano, "reforma_b2b_pct": b2b, "reforma_cbs_pct": cbs, "reforma_ibs_pct": ibs,
                           "reforma_compras_credito": compras})
    r = simular_reforma(a1["rec_bruta"], res["aliq"], anexo, ano, cbs / 100, ibs / 100, b2b / 100, compras)
    tab = pd.DataFrame({
        "Item": ["DAS", "IBS + CBS sobre as vendas", "(−) Crédito de IBS + CBS das compras", "Total de tributos (DAS + IBS/CBS)",
                 "Crédito que os clientes empresas aproveitam"],
        "Simples puro": [brl(r["das_puro"]), "incluso no DAS", "não aproveita", brl(r["puro_total"]), brl(r["puro_credito_cliente"])],
        "Simples híbrido": [brl(r["das_hibrido"]), brl(r["iva_debito"]), brl(-r["iva_credito"]), brl(r["hib_total"]),
                            brl(r["hib_credito_cliente"])]})
    st.dataframe(tab, hide_index=True, width="stretch")
    k = st.columns(3)
    k[0].metric("Diferença de imposto (híbrido − puro)", brl(r["diferenca"]))
    k[1].metric("Crédito a mais para clientes B2B no híbrido", brl(r["hib_credito_cliente"] - r["puro_credito_cliente"]))
    k[2].metric("Parcela do DAS substituída pela reforma", pct(r["parcela_substituida"]))
    if b2b == 0:
        st.info("Vendendo só para consumidor final, o crédito não beneficia o cliente: o Simples puro tende a ser mais vantajoso.")
    elif r["diferenca"] > 0 and r["hib_credito_cliente"] - r["puro_credito_cliente"] > r["diferenca"]:
        st.success("O híbrido custa mais imposto para a empresa, mas gera aos clientes empresas um crédito maior que essa diferença: "
                   "pode valer a pena se isso permitir vender mais ou cobrar um preço melhor no B2B.")
    elif r["diferenca"] <= 0:
        st.success("Neste cenário o híbrido sai mais barato e ainda gera mais crédito aos clientes.")
    else:
        st.warning("Neste cenário o híbrido aumenta o imposto mais do que o crédito que gera aos clientes.")
    st.caption("Fontes: LC 214/2025 (art. 41, opção pelo regime regular do IBS e da CBS); partilha do DAS da 1ª faixa da LC 123/2006 "
               "com redação da LC 155/2016; CBS de referência de 11,3% para 2027 calculada pelo TCU, sujeita à definição do Senado; "
               "estimativas de 26,5% a 28,6% para IBS + CBS em 2033. Todos os percentuais são editáveis.")
    proxima("reforma")

# ##############################################################################
# PARTE: pnbox_ui_paginas6.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: VALIDAÇÃO, MARKETING, INCERTEZA E DEFESA (v7.1)
# ==============================================================================


def MC() -> dict:
    """Simulação de Monte Carlo com os dados atuais (recalcula quando algo muda)."""
    if "mc" not in _CACHE:
        res = R()
        _CACHE["mc"] = monte_carlo(res, faixas_mc(d)) if res["anos"][1]["rec_bruta"] > 0 else None
    return _CACHE["mc"]


def _semear(tabela: str, linhas: list, chave_: str):
    recarregar_tabela(tabela, pd.concat([d[tabela], pd.DataFrame(linhas)], ignore_index=True), chave_)
    st.rerun()


def _ir_botao(rotulo: str, destino: str, chave_: str):
    if st.button(rotulo, key=chave_):
        ir(destino)
        st.rerun()


# ------------------------------------------------------------------------------
# PLACAR DE EVIDÊNCIAS
# ------------------------------------------------------------------------------
def pg_evidencias():
    cabecalho("evidencias")
    res = R()
    st.markdown("""<div class="pn-card-info">Um plano de negócio é um conjunto de <b>hipóteses</b> até que os números sejam testados
com clientes reais (Ries, 2012; Blank, 2013). Marque a origem de cada premissa. O <b>peso</b> vem da análise de sensibilidade:
quanto mais a premissa mexe no resultado, mais ela conta no placar.</div>""", unsafe_allow_html=True)
    st.write("")
    pr = premissas_plano(proj, d, res)
    faltam = [p for p in pr["premissa"] if p not in set(d["evidencias"]["premissa"])]
    if faltam:
        recarregar_tabela("evidencias", pd.concat([d["evidencias"], pd.DataFrame(
            {"premissa": faltam, "status": "Estimada (sem fonte)", "fonte": "", "observacao": ""})], ignore_index=True), "premissa")
        st.rerun()
    pl = placar_evidencias(proj, d, res)
    k = st.columns(3)
    k[0].metric("Placar de evidências", pct(pl["score"], 0), help="Média das notas (estimada = 0; pesquisada = 0,5; testada = 1) "
                                                                  "ponderada pelo impacto de cada premissa no resultado.")
    k[1].metric("Premissas testadas com clientes", f"{pl['testadas']} de {pl['n']}")
    k[2].metric("Valide primeiro", pl["proxima"] or "tudo testado")
    st.progress(min(max(pl["score"], 0.0), 1.0))
    vis = pl["tabela"][["premissa", "valor_plano", "peso", "status"]].rename(columns={
        "premissa": "Premissa", "valor_plano": "Valor no plano", "peso": "Peso", "status": "Origem"})
    vis["Peso"] = vis["Peso"].map(lambda v: pct(v, 0))
    vis["Origem"] = vis["Origem"].map(STATUS_CURTO)
    st.dataframe(vis, hide_index=True, width="stretch")
    st.subheader("Informe a origem de cada premissa")
    ev = d["evidencias"][d["evidencias"]["premissa"].isin(set(pr["premissa"]))].reset_index(drop=True)

    def _salvar(x):
        outras = d["evidencias"][~d["evidencias"]["premissa"].isin(set(x["premissa"]))]
        completo = pd.concat([outras, x], ignore_index=True)
        salvar(conn, "evidencias", completo, "premissa")
        d["evidencias"] = limpar(completo.reset_index(drop=True), "evidencias")

    tabela_editavel(ev, {"premissa": T("Premissa", disabled=True, width="medium"),
                         "status": S("Origem do número", options=list(STATUS_EVIDENCIA), required=True, width="large"),
                         "fonte": T("Fonte (quem, onde, quando)", width="large"),
                         "observacao": T("O que o teste mostrou", width="large")},
                    "evidencias", "placar", normalizar=lambda x: limpar(x, "evidencias"), salvar_fn=_salvar, adicionar=False)
    st.caption("Exemplos de evidência: entrevistas registradas, respostas do teste de preço, pré-vendas, lista de espera, "
               "orçamentos de fornecedores, dados do IBGE ou do Sebrae, preços dos concorrentes.")
    c = st.columns(3)
    with c[0]:
        _ir_botao("🎙️ Registrar entrevistas", "entrevistas", "ev_ent")
    with c[1]:
        _ir_botao("🙋 Testar o preço", "teste_preco", "ev_vw")
    with c[2]:
        _ir_botao("🧪 Planejar experimentos", "experimentos", "ev_exp")
    proxima("evidencias")


# ------------------------------------------------------------------------------
# ENTREVISTAS COM CLIENTES
# ------------------------------------------------------------------------------
def pg_entrevistas():
    cabecalho("entrevistas")
    res = R()
    with st.expander("📋 Roteiro de entrevista de descoberta", expanded=d["entrevistas"].empty):
        st.markdown(ROTEIRO_ENTREVISTA)
    st.caption("Registre uma linha por pessoa. Em **dores**, separe as dores por ponto e vírgula e use sempre as mesmas palavras "
               "para a mesma dor: assim o programa consegue contá-las.")
    editor("entrevistas", {"data": T("Data"), "entrevistado": T("Entrevistado (nome ou código)", required=True),
                           "perfil": T("Perfil / segmento"), "dores": T("Dores citadas (separe com ;)", width="large"),
                           "solucao_atual": T("O que usa hoje", width="medium"),
                           "pagaria": S("Pagaria?", options=PAGARIA), "valor_pagaria": moeda("Quanto pagaria"),
                           "aprendizado": T("Principal aprendizado", width="large")}, "entrevistado")
    preco = preco_referencia(proj, res)
    r = resumo_entrevistas(d, preco)
    if r["n"] == 0:
        return proxima("entrevistas")
    k = st.columns(4)
    k[0].metric("Entrevistas", r["n"], help="Recomenda-se ao menos 10 a 15 por segmento antes de confiar nos padrões.")
    k[1].metric("Pagariam", pct(r["pct_sim"], 0), f"+ {pct(r['pct_talvez'], 0)} talvez", delta_color="off")
    k[2].metric("Valor mediano citado", brl(r["mediana_valor"]) if r["mediana_valor"] == r["mediana_valor"] else "n.a.")
    k[3].metric("Preço do plano", brl(preco) if preco == preco else "n.a.",
                f"{pct(r['valor_vs_preco'] - 1, 0)} vs. citado" if r["valor_vs_preco"] == r["valor_vs_preco"] else None,
                delta_color="off")
    if not r["dores"].empty:
        g = r["dores"].assign(Pct=r["dores"]["% das entrevistas"] * 100)
        st.altair_chart(alt.Chart(g).mark_bar(color="#2f6fd0").encode(
            x=alt.X("Pct:Q", title="% das entrevistas", scale=alt.Scale(domain=[0, 100])),
            y=alt.Y("Dor citada:N", sort="-x", title=None), tooltip=["Dor citada", "Entrevistas"]).properties(height=34 * len(g) + 40),
            width="stretch")
        top, qt = r["dor_top"]
        if qt >= r["n"] / 2:
            st.success(f"A dor **\"{top}\"** apareceu em {qt} de {r['n']} entrevistas: é um problema recorrente, bom sinal para a proposta de valor.")
        else:
            st.warning("Nenhuma dor apareceu em metade das entrevistas. Talvez o segmento esteja amplo demais ou o problema não seja tão forte.")
    if r["n"] < 10:
        st.info(f"Faltam {10 - r['n']} entrevista(s) para chegar ao mínimo de 10 usado na régua de prontidão.")
    proxima("entrevistas")


# ------------------------------------------------------------------------------
# TESTE DE PREÇO (VAN WESTENDORP) E INTENÇÃO DE COMPRA
# ------------------------------------------------------------------------------
def _grafico_vw(vw: dict, preco: float):
    cv = vw["curvas"].melt(id_vars="Preço", value_vars=["Muito barato", "Barato", "Caro", "Muito caro"],
                           var_name="Curva", value_name="Proporção")
    linhas = alt.Chart(cv).mark_line(strokeWidth=2.5).encode(
        x=alt.X("Preço:Q", title="Preço (R$)"), y=alt.Y("Proporção:Q", axis=alt.Axis(format="%"), title="% dos respondentes"),
        color=alt.Color("Curva:N", scale=alt.Scale(domain=["Muito barato", "Barato", "Caro", "Muito caro"],
                                                   range=["#7fb3d5", "#2471a3", "#e59866", "#c0392b"]),
                        legend=alt.Legend(orient="bottom", title=None)))
    faixa = alt.Chart(pd.DataFrame({"a": [vw["pmc"]], "b": [vw["pme"]]})).mark_rect(opacity=0.12, color="#27ae60").encode(x="a:Q", x2="b:Q")
    camadas = faixa + linhas
    if preco == preco:
        camadas += alt.Chart(pd.DataFrame({"p": [preco], "t": ["preço do plano"]})).mark_rule(strokeDash=[5, 4], color="black").encode(
            x="p:Q", tooltip=["t"])
    st.altair_chart(camadas.properties(height=380), width="stretch")


def pg_teste_preco():
    cabecalho("teste_preco")
    res = R()
    t1, t2 = st.tabs(["🙋 Teste de preço (Van Westendorp)", "📊 Intenção de compra ajustada"])
    with t1:
        st.caption("Pergunte a pessoas do público-alvo as quatro questões abaixo, sobre o mesmo produto, e lance as respostas. "
                   "O método (Van Westendorp, 1976) mostra a faixa de preço que os clientes consideram aceitável.")
        produtos = [p for p in res["cat"]["descricao"].astype(str) if p.strip()]
        if res.get("recorrencia", {}).get("tem"):
            produtos.append("Assinatura (ticket médio)")
        opcoes = ["(preço médio do plano)"] + produtos
        atual = proj["vw_produto"] if proj["vw_produto"] in opcoes else "(preço médio do plano)"
        esc = st.selectbox("Produto testado", opcoes, index=opcoes.index(atual), key=wk("vw_prod"))
        auto_unica("projeto", {"vw_produto": "" if esc.startswith("(") else esc})
        nome = esc if not esc.startswith("(") else "o produto"
        with st.expander("📝 Perguntas prontas para o Google Forms"):
            st.code(f"1. A que preço você acharia {nome} tão barato que duvidaria da qualidade?\n"
                    f"2. A que preço você acharia {nome} barato, uma boa compra?\n"
                    f"3. A que preço {nome} começaria a parecer caro, mas ainda pensaria em comprar?\n"
                    f"4. A que preço {nome} ficaria caro demais, e você não compraria de jeito nenhum?", language=None)
            st.caption("Use respostas numéricas (R$). Exporte as respostas para planilha e cole nas colunas abaixo.")
        editor("precos_vw", {"respondente": T("Respondente", required=True), "muito_barato": moeda("1. Barato demais"),
                             "barato": moeda("2. Barato"), "caro": moeda("3. Caro"), "muito_caro": moeda("4. Caro demais")},
               "respondente")
        preco = preco_referencia(proj, res)
        vw = van_westendorp(d["precos_vw"])
        if vw is None:
            st.info("Lance pelo menos 5 respostas para ver as curvas.")
        elif not vw["ok"]:
            st.info(f"{vw['n']} resposta(s) válida(s). São necessárias pelo menos 5 (de preferência 30 ou mais). "
                    "Respostas fora de ordem (barato maior que caro, por exemplo) são descartadas.")
        else:
            _grafico_vw(vw, preco)
            k = st.columns(4)
            k[0].metric("Limite inferior (PMC)", brl(vw["pmc"]), help="Abaixo disso o cliente desconfia da qualidade.")
            k[1].metric("Preço ótimo (OPP)", brl(vw["opp"]), help="Menos pessoas acham barato demais ou caro demais.")
            k[2].metric("Preço de indiferença (IPP)", brl(vw["ipp"]), help="Tantos acham caro quanto barato.")
            k[3].metric("Limite superior (PME)", brl(vw["pme"]), help="Acima disso a maioria considera caro demais.")
            if preco == preco:
                if vw["pmc"] <= preco <= vw["pme"]:
                    st.success(f"O preço do plano ({brl(preco)}) está dentro da faixa aceitável. Com {vw['n']} respostas, "
                               "esta premissa pode ser marcada como **testada com clientes** no placar de evidências.")
                elif preco > vw["pme"]:
                    st.error(f"O preço do plano ({brl(preco)}) está **acima** do que os clientes aceitam ({brl(vw['pme'])}). "
                             "Reveja custos, a proposta de valor ou o público.")
                else:
                    st.warning(f"O preço do plano ({brl(preco)}) está **abaixo** da faixa ({brl(vw['pmc'])}): há espaço para "
                               "cobrar mais, e preço baixo demais pode passar impressão de baixa qualidade.")
            if vw["n"] < 30:
                st.caption(f"Com {vw['n']} respostas o resultado é exploratório. Amostras de 30 ou mais tornam as curvas mais estáveis.")
    with t2:
        st.caption("Pesquisas de intenção superestimam a demanda: quem diz que \"certamente compraria\" nem sempre compra. "
                   "Aplique fatores de conversão à intenção declarada e compare com o número de clientes do plano.")
        c = st.columns(4)
        ic = c[0].number_input("Certamente compraria (%)", 0.0, 100.0, float(proj["intencao_certamente_pct"] or 0), 1.0, key=wk("int_c"))
        ip = c[1].number_input("Provavelmente compraria (%)", 0.0, 100.0, float(proj["intencao_provavel_pct"] or 0), 1.0, key=wk("int_p"))
        fc = c[2].number_input("Dos que \"certamente\", quantos compram (%)", 0.0, 100.0, float(proj["fator_certamente_pct"] or 70),
                               5.0, key=wk("fat_c"))
        fp = c[3].number_input("Dos que \"provavelmente\", quantos compram (%)", 0.0, 100.0, float(proj["fator_provavel_pct"] or 30),
                               5.0, key=wk("fat_p"))
        auto_unica("projeto", {"intencao_certamente_pct": ic, "intencao_provavel_pct": ip, "fator_certamente_pct": fc,
                               "fator_provavel_pct": fp})
        ia = intencao_ajustada(proj, res)
        if not ia["ok"]:
            st.info("Preencha o SAM (clientes alcançáveis) em Tamanho do mercado e os percentuais da pesquisa.")
        else:
            k = st.columns(3)
            k[0].metric("Clientes pela intenção declarada", num_br(ia["declarados"], 0))
            k[1].metric("Clientes após o ajuste", num_br(ia["ajustados"], 0))
            k[2].metric("Clientes que o plano supõe no Ano 1", num_br(ia["plano"], 0) if ia["plano"] == ia["plano"] else "n.a.",
                        help="Receita do Ano 1 ÷ gasto anual por cliente (Tamanho do mercado).")
            if ia["cobertura"] == ia["cobertura"]:
                if ia["cobertura"] >= 1:
                    st.success(f"A demanda ajustada cobre {pct(ia['cobertura'], 0)} dos clientes do plano.")
                else:
                    st.warning(f"A demanda ajustada cobre só {pct(ia['cobertura'], 0)} dos clientes do plano: a meta de vendas parece otimista.")
        st.caption("Fatores sugeridos (70% e 30%) são uma regra prática de pesquisa de mercado; ajuste conforme o setor e o preço.")
    proxima("teste_preco")


# ------------------------------------------------------------------------------
# CANVAS DA PROPOSTA DE VALOR
# ------------------------------------------------------------------------------
def pg_vpc():
    cabecalho("vpc")
    st.caption("Lado direito: o **perfil do cliente**. Lado esquerdo: o **mapa de valor** da empresa (Osterwalder et al., 2014). "
               "Escreva um item por linha. O encaixe acontece quando cada dor e ganho importante tem uma resposta da oferta.")
    job = st.text_input("Que trabalho (job) o cliente contrata o seu produto para fazer?", proj["vpc_job"], key=wk("vpc_job"),
                        placeholder="Quando estou sem tempo no almoço, quero comer algo saudável, para manter a dieta sem perder o horário.",
                        help="Formato de Christensen et al. (2016): quando [situação], quero [motivação], para [resultado].")
    v = {"vpc_job": job}
    ce, cd = st.columns(2)
    with cd:
        st.markdown("##### 🙂 Perfil do cliente")
        for k_ in ("vpc_tarefas", "vpc_dores", "vpc_ganhos"):
            _, rot, ajuda = next(b for b in BLOCOS_VPC if b[0] == k_)
            v[k_] = st.text_area(rot, proj[k_], height=130, key=wk(k_), help=ajuda)
    with ce:
        st.markdown("##### 🎁 Mapa de valor")
        for k_ in ("vpc_produtos", "vpc_aliviadores", "vpc_criadores"):
            _, rot, ajuda = next(b for b in BLOCOS_VPC if b[0] == k_)
            v[k_] = st.text_area(rot, proj[k_], height=130, key=wk(k_), help=ajuda)
    auto_unica("projeto", v)
    salvo_txt()
    enc = encaixe_vpc(proj, d)
    if enc["total"] == 0:
        return proxima("vpc")
    st.subheader("Encaixe: como a oferta responde a cada dor e ganho")
    opcoes = [""] + linhas_txt(proj["vpc_aliviadores"]) + linhas_txt(proj["vpc_criadores"])

    def _salvar(x):
        salvar(conn, "vpc_encaixe", x, "item")
        d["vpc_encaixe"] = limpar(x.reset_index(drop=True), "vpc_encaixe")

    tabela_editavel(enc["tabela"][["item", "tipo", "resposta"]], {
        "item": T("Dor ou ganho do cliente", disabled=True, width="large"), "tipo": T("Tipo", disabled=True, width="small"),
        "resposta": S("Respondido por", options=opcoes, width="large")}, "vpc_encaixe",
        "|".join(enc["tabela"]["item"])[:200], normalizar=lambda x: limpar(x, "vpc_encaixe"), salvar_fn=_salvar, adicionar=False)
    enc = encaixe_vpc(proj, d)
    st.metric("Encaixe produto-cliente", pct(enc["encaixe"], 0), f"{enc['cobertos']} de {enc['total']} itens respondidos",
              delta_color="off")
    if enc["encaixe"] < 0.6:
        st.warning("Muitas dores e ganhos ainda não têm resposta na oferta. Ou a oferta precisa mudar, ou esses itens não são prioridade.")
    proxima("vpc")


# ------------------------------------------------------------------------------
# CURVA DE VALOR E ERRC (OCEANO AZUL)
# ------------------------------------------------------------------------------
def pg_oceano():
    cabecalho("oceano")
    st.caption("A **curva de valor** mostra como cada empresa se sai nos atributos que o setor oferece (notas da Análise da "
               "concorrência). A matriz **ERRC** propõe uma curva nova: o que **E**liminar, **R**eduzir, **E**levar e **C**riar "
               "para se diferenciar sem aumentar custos (Kim; Mauborgne, 2005).")
    atributos = list(dict.fromkeys(d["concorrencia_notas"]["atributo"].astype(str)))
    if not atributos:
        st.warning("Primeiro dê notas aos atributos na Análise da concorrência.")
        _ir_botao("Ir para Análise da concorrência →", "concorrencia", "oc_conc")
        return proxima("oceano")
    editor("errc", {"atributo": S("Atributo (escolha ou digite um novo para Criar)", options=atributos, required=True, width="medium"),
                    "acao": S("Ação", options=ACOES_ERRC, required=True),
                    "nova_nota": N("Nova nota (0 a 10)", min_value=0.0, max_value=10.0, step=0.5),
                    "justificativa": T("Por quê?", width="large")}, "atributo")
    st.caption("Para **criar** um atributo novo, que nenhum concorrente oferece, escreva-o primeiro na Análise da concorrência "
               "com nota 0 para todos e depois marque aqui a ação Criar com a nota pretendida.")
    cv = curva_valor(d)
    ordem = list(dict.fromkeys(cv["Atributo"]))
    st.altair_chart(alt.Chart(cv).mark_line(point=True).encode(
        x=alt.X("Atributo:N", sort=ordem, title=None, axis=alt.Axis(labelAngle=-30)), y=alt.Y("Nota:Q", scale=alt.Scale(domain=[0, 10])),
        color=alt.Color("Empresa:N", legend=alt.Legend(orient="bottom", title=None)),
        strokeDash=alt.condition(alt.datum.Empresa == "Nova curva (ERRC)", alt.value([1, 0]), alt.value([4, 3])),
        strokeWidth=alt.condition(alt.datum.Empresa == "Nova curva (ERRC)", alt.value(4), alt.value(1.6)),
        tooltip=["Empresa", "Atributo", "Nota"]).properties(height=380), width="stretch")
    errc = d["errc"]
    if not errc.empty:
        cols = st.columns(4)
        for c_, (ac, ic) in zip(cols, [("Eliminar", "✂️"), ("Reduzir", "⬇️"), ("Elevar", "⬆️"), ("Criar", "✨")]):
            with c_.container(border=True):
                st.markdown(f"**{ic} {ac}**")
                for a in errc.loc[errc["acao"] == ac, "atributo"]:
                    st.markdown(f"- {a}")
        st.info(leitura_errc(d))
    proxima("oceano")


# ------------------------------------------------------------------------------
# POSICIONAMENTO E MAPA PERCEPTUAL
# ------------------------------------------------------------------------------
def pg_posicionamento():
    cabecalho("posicionamento")
    st.caption("A declaração de posicionamento diz, em uma frase, **para quem** é o negócio, **o que** entrega e **por que** é "
               "diferente (modelo de Moore, 1991). O mapa perceptual mostra onde a marca fica na cabeça do cliente.")
    c = st.columns(2)
    v = {}
    v["posic_segmento"] = c[0].text_input("Para (público-alvo)", proj["posic_segmento"], key=wk("ps_seg"))
    v["posic_necessidade"] = c[1].text_input("que (necessidade ou problema)", proj["posic_necessidade"], key=wk("ps_nec"))
    v["posic_marca"] = c[0].text_input("a marca", proj["posic_marca"] or proj["nome_negocio"], key=wk("ps_marca"))
    v["posic_categoria"] = c[1].text_input("é (categoria)", proj["posic_categoria"], key=wk("ps_cat"))
    v["posic_beneficio"] = c[0].text_input("que (benefício principal)", proj["posic_beneficio"], key=wk("ps_ben"))
    v["posic_concorrente"] = c[1].text_input("Diferentemente de (principal alternativa)", proj["posic_concorrente"], key=wk("ps_conc"))
    v["posic_diferencial"] = st.text_input("o nosso produto (diferencial que prova o benefício)", proj["posic_diferencial"], key=wk("ps_dif"))
    auto_unica("projeto", v)
    frase = frase_posicionamento(proj)
    if frase:
        st.success("**Declaração de posicionamento:** " + frase)
    st.subheader("Mapa perceptual")
    c = st.columns(2)
    ex = c[0].text_input("Eixo horizontal", proj["mapa_eixo_x"], key=wk("mp_x"))
    ey = c[1].text_input("Eixo vertical", proj["mapa_eixo_y"], key=wk("mp_y"))
    auto_unica("projeto", {"mapa_eixo_x": ex, "mapa_eixo_y": ey})
    if d["mapa_perceptual"].empty:
        nomes = [proj["nome_negocio"] or "Nossa Empresa"] + [n for n in d["concorrentes"]["nome"] if str(n).strip()]
        if st.button("Incluir a empresa e os concorrentes cadastrados"):
            _semear("mapa_perceptual", [{"marca": n, "eixo_x": 5.0, "eixo_y": 5.0} for n in nomes], "marca")
    editor("mapa_perceptual", {"marca": T("Marca", required=True), "eixo_x": N("Nota no eixo horizontal (0 a 10)", min_value=0.0,
                                                                                max_value=10.0, step=0.5),
                               "eixo_y": N("Nota no eixo vertical (0 a 10)", min_value=0.0, max_value=10.0, step=0.5)}, "marca")
    mp = d["mapa_perceptual"]
    if not mp.empty:
        nossa = proj["nome_negocio"] or "Nossa Empresa"
        g = mp.assign(Tipo=["Nossa empresa" if m in (nossa, "Nossa Empresa") else "Concorrente" for m in mp["marca"]])
        base = alt.Chart(g).encode(x=alt.X("eixo_x:Q", scale=alt.Scale(domain=[0, 10]), title=ex),
                                   y=alt.Y("eixo_y:Q", scale=alt.Scale(domain=[0, 10]), title=ey))
        pts = base.mark_circle(size=380, opacity=0.85).encode(
            color=alt.Color("Tipo:N", scale=alt.Scale(domain=["Nossa empresa", "Concorrente"], range=["#e67e22", "#2f6fd0"]),
                            legend=alt.Legend(orient="bottom", title=None)), tooltip=["marca", "eixo_x", "eixo_y"])
        txt = base.mark_text(dy=-18, fontSize=11).encode(text="marca:N")
        meio = alt.Chart(pd.DataFrame({"v": [5]})).mark_rule(color="#999").encode(x="v:Q") + \
            alt.Chart(pd.DataFrame({"v": [5]})).mark_rule(color="#999").encode(y="v:Q")
        st.altair_chart((meio + pts + txt).properties(height=420), width="stretch")
        st.caption("Procure um espaço vazio no mapa que seja valorizado pelo cliente. Se a sua marca está colada em um concorrente, "
                   "o cliente não tem motivo para trocar.")
    proxima("posicionamento")


# ------------------------------------------------------------------------------
# ORÇAMENTO DE MARKETING, FUNIL AARRR, CALENDÁRIO E NPS
# ------------------------------------------------------------------------------
def pg_orcamento_mkt():
    cabecalho("orcamento_mkt")
    res = R()
    t1, t2, t3, t4 = st.tabs(["💸 Orçamento e CAC por canal", "🪜 Funil AARRR", "🗓️ Calendário de ações", "⭐ Meta de NPS"])
    with t1:
        st.caption("Quanto será gasto por mês em cada canal e quantos clientes novos cada um deve trazer. O programa calcula o custo "
                   "de aquisição (CAC) e o retorno de cada canal.")
        canais = [c for c in d["canais"]["canal"].astype(str) if c.strip()]
        editor("orcamento_canais", {"canal": S("Canal", options=canais, required=True) if canais else T("Canal", required=True),
                                    "gasto_mensal": moeda("Gasto mensal"), "clientes_mes": N("Clientes novos por mês", min_value=0.0, step=1.0),
                                    "observacao": T("Como será medido", width="large")}, "canal")
        oc = orcamento_canais(proj, d, res)
        if not oc.empty:
            vis = oc.copy()
            for c_ in ("Gasto mensal", "CAC", "Valor do cliente"):
                vis[c_] = vis[c_].map(lambda x: brl(x) if x == x else "n.a.")
            vis["Retorno por R$ 1"] = vis["Retorno por R$ 1"].map(lambda x: f"R$ {num_br(x, 2)}" if x == x else "n.a.")
            vis = vis.rename(columns={"Valor do cliente": f"Valor do cliente ({oc.attrs.get('rotulo_valor', '')})"})
            st.dataframe(vis, hide_index=True, width="stretch")
            tot_g, tot_c = float(oc["Gasto mensal"].sum()), float(oc["Clientes novos/mês"].sum())
            k = st.columns(3)
            k[0].metric("Orçamento mensal de aquisição", brl(tot_g))
            k[1].metric("CAC médio", brl(div(tot_g, tot_c)) if tot_c else "n.a.")
            verba = float(proj["verba_aquisicao_mensal"] or 0)
            com = float(res["desp_com_m"] + res["pessoal_com_m"])
            k[2].metric("Gasto comercial no plano (mês)", brl(verba if verba > 0 else com),
                        help="Verba de aquisição informada em Assinaturas ou despesas comerciais fixas.")
            if tot_g > (verba if verba > 0 else com) * 1.05:
                st.warning("O orçamento por canal é maior que o gasto comercial previsto nos custos: inclua a diferença em Custos e "
                           "despesas, senão a DRE fica otimista.")
            pior = oc.sort_values("Retorno por R$ 1").iloc[0]
            if pior["Retorno por R$ 1"] == pior["Retorno por R$ 1"] and pior["Retorno por R$ 1"] < 1:
                st.error(f"O canal **{pior['Canal']}** devolve menos do que custa: cada R$ 1 investido traz R$ {num_br(pior['Retorno por R$ 1'], 2)} "
                         "em valor de cliente.")
    with t2:
        st.caption("Funil de crescimento em cinco etapas (Ellis; Brown, 2017): para cada uma, a métrica que mostra se está "
                   "funcionando, a meta e a ação.")
        if d["aarrr"].empty:
            if st.button("Criar as cinco etapas"):
                _semear("aarrr", [{"etapa": e, "metrica": "", "meta": "", "acao": ""} for e, _ in ETAPAS_AARRR], "etapa")
            for e, perg in ETAPAS_AARRR:
                st.markdown(f"- **{e}:** {perg}")
        editor("aarrr", {"etapa": S("Etapa", options=[e for e, _ in ETAPAS_AARRR], required=True),
                         "metrica": T("Métrica", width="medium"), "meta": T("Meta"), "acao": T("Ação principal", width="large")}, "etapa")
    with t3:
        editor("calendario_mkt", {"mes": N("Mês do plano", min_value=1, max_value=120, step=1), "acao": T("Ação", required=True, width="large"),
                                  "canal": T("Canal"), "custo": moeda("Custo")}, "acao")
        cal = d["calendario_mkt"]
        if not cal.empty:
            g = cal.groupby("mes", as_index=False)["custo"].sum()
            st.altair_chart(alt.Chart(g).mark_bar(color="#e67e22").encode(
                x=alt.X("mes:O", title="Mês"), y=alt.Y("custo:Q", title="Custo (R$)")).properties(height=240), width="stretch")
            st.caption(f"Total do calendário: {brl(cal['custo'].sum())}")
    with t4:
        st.markdown("O **Net Promoter Score** pergunta: *de 0 a 10, quanto você recomendaria a empresa a um amigo?* NPS = % de "
                    "promotores (9 e 10) menos % de detratores (0 a 6), variando de -100 a +100 (Reichheld, 2003).")
        nps = st.slider("Meta de NPS para o primeiro ano", -100, 100, int(proj["nps_meta"] or 50), 5, key=wk("nps"))
        auto_unica("projeto", {"nps_meta": nps})
        st.caption("Referência comum: acima de 50 é considerado muito bom e acima de 75, excelente. Ligue esta meta a um objetivo da "
                   "perspectiva de clientes no Balanced Scorecard.")
    proxima("orcamento_mkt")


# ------------------------------------------------------------------------------
# SIMULAÇÃO DE MONTE CARLO E PERDA ACEITÁVEL
# ------------------------------------------------------------------------------
def pg_montecarlo():
    cabecalho("montecarlo")
    res = R()
    if res["anos"][1]["rec_bruta"] <= 0:
        st.warning("Cadastre produtos ou assinaturas para simular.")
        return proxima("montecarlo")
    st.caption("Em vez de um único cenário, o programa sorteia 3.000 combinações de preço, volume, custos e crescimento dentro das "
               "faixas abaixo (distribuição triangular: mínimo, valor do plano, máximo) e mostra a **chance** de o projeto dar certo.")
    fx = faixas_mc(d)

    def _salvar(x):
        salvar(conn, "mc_faixas", x, "variavel")
        d["mc_faixas"] = limpar(x.reset_index(drop=True), "mc_faixas")

    tabela_editavel(fx, {"variavel": T("Variável", disabled=True, width="large"),
                         "minimo_pct": N("Pior caso (%)", step=1.0, max_value=0.0),
                         "maximo_pct": N("Melhor caso (%)", step=1.0, min_value=0.0)},
                    "mc_faixas", "faixas", normalizar=lambda x: limpar(x, "mc_faixas"), salvar_fn=_salvar, adicionar=False)
    st.caption("Para custos, o \"melhor caso\" positivo significa custo maior (é o limite de alta). Use pesquisas, orçamentos e o "
               "histórico do setor para definir as faixas.")
    mc = MC()
    k = st.columns(4)
    k[0].metric("Chance de VPL positivo", pct(mc["p_vpl_pos"], 0))
    k[1].metric("VPL mediano", brl(mc["p50"]))
    k[2].metric("VPL nos 5% piores casos", brl(mc["p5"]))
    k[3].metric("Chance de caixa negativo no Ano 1", pct(mc["p_caixa_neg"], 0))
    hist = pd.DataFrame({"VPL": mc["vpl"]})
    st.altair_chart(alt.Chart(hist).mark_bar(opacity=0.85).encode(
        x=alt.X("VPL:Q", bin=alt.Bin(maxbins=45), title="VPL (R$)"), y=alt.Y("count():Q", title="Cenários"),
        color=alt.condition(alt.datum.VPL > 0, alt.value("#27ae60"), alt.value("#c0392b"))).properties(height=300) +
        alt.Chart(pd.DataFrame({"x": [0]})).mark_rule(color="black").encode(x="x:Q"), width="stretch")
    st.info(leitura_mc(mc))
    st.subheader("Perda aceitável dos sócios")
    st.caption("Em vez de perguntar só \"quanto posso ganhar?\", o empreendedor experiente pergunta \"quanto aceito perder?\" e "
               "desenha o negócio para que o pior caso caiba nesse limite (Sarasvathy, 2001). Informe a perda aceitável de cada "
               "sócio em Informações e sócios.")
    rv = st.slider("Quanto do investimento fixo seria recuperado vendendo os bens, se o negócio fechar (%)", 0, 100,
                   int(proj["revenda_ativos_pct"] if proj["revenda_ativos_pct"] is not None else 40), 5, key=wk("revenda"))
    auto_unica("projeto", {"revenda_ativos_pct": rv})
    pa = perda_aceitavel(proj, d, res)
    k = st.columns(3)
    k[0].metric("Perda potencial se fechar", brl(pa["potencial"]),
                help="Capital dos sócios + empréstimos (normalmente garantidos pelos sócios) − venda dos bens.")
    k[1].metric("Perda aceitável somada", brl(pa["aceitavel"]))
    k[2].metric("Situação", ("cabe no limite" if pa["cabe"] else "acima do limite") if pa["ok"] else "não informada")
    if pa["ok"] and not pa["cabe"]:
        st.warning("O pior caso passa do que os sócios aceitam perder. Opções: começar menor (MVP), alugar em vez de comprar "
                   "equipamentos, buscar parceiros ou pré-vendas antes de investir.")
    _ir_botao("👥 Informar a perda aceitável dos sócios", "projeto", "mc_socios")
    proxima("montecarlo")


# ------------------------------------------------------------------------------
# CURVA DE ADOÇÃO (BASS)
# ------------------------------------------------------------------------------
def pg_adocao():
    cabecalho("adocao")
    res = R()
    st.caption("O modelo de Bass (1969) descreve como um produto novo se espalha: alguns clientes compram por iniciativa própria "
               "(**inovação**, p) e a maioria compra porque viu outros usando (**imitação**, q). A curva serve para testar se o "
               "crescimento previsto no plano é plausível.")
    c = st.columns(3)
    m0 = float(proj["bass_mercado"] or 0) or float(proj["sam_clientes"] or 0)
    m = c[0].number_input("Mercado potencial (clientes)", 0.0, 1e9, m0, 100.0, key=wk("bass_m"),
                          help="Padrão: SAM (clientes alcançáveis) de Tamanho do mercado.")
    p = c[1].number_input("Coeficiente de inovação (p)", 0.001, 0.5, float(proj["bass_p"] or 0.03), 0.005, format="%.3f", key=wk("bass_p"))
    q = c[2].number_input("Coeficiente de imitação (q)", 0.0, 1.5, float(proj["bass_q"] or 0.38), 0.02, format="%.2f", key=wk("bass_q"))
    auto_unica("projeto", {"bass_mercado": m if m != float(proj["sam_clientes"] or 0) else 0.0, "bass_p": p, "bass_q": q})
    st.caption("Valores médios de p = 0,03 e q = 0,38 vêm de uma meta-análise de estudos de difusão (Sultan; Farley; Lehmann, 1990). "
               "Produtos que dependem muito de indicação (aplicativos, serviços locais) costumam ter q maior.")
    bc = bass_comparacao(proj, res)
    if not bc["ok"]:
        st.info("Informe o mercado potencial ou preencha o SAM em Tamanho do mercado.")
        return proxima("adocao")
    tb = bc["tabela"]
    series = ["Adotantes acumulados"] + (["Clientes no plano"] if bc["tem_plano"] else [])
    g = tb.melt(id_vars="Ano", value_vars=series, var_name="Série", value_name="Clientes")
    st.altair_chart(alt.Chart(g).mark_line(point=True, strokeWidth=3).encode(
        x=alt.X("Ano:O"), y=alt.Y("Clientes:Q"), color=alt.Color("Série:N", legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["Ano", "Série", alt.Tooltip("Clientes:Q", format=",.0f")]).properties(height=320), width="stretch")
    vis = tb.copy()
    for c_ in ("Adotantes acumulados", "Novos adotantes no ano", "Clientes no plano"):
        if c_ in vis:
            vis[c_] = vis[c_].map(lambda x: num_br(x, 0))
    vis["% do mercado potencial"] = vis["% do mercado potencial"].map(lambda x: pct(x, 1))
    if "Plano ÷ Bass" in vis:
        vis["Plano ÷ Bass"] = vis["Plano ÷ Bass"].map(lambda x: f"{num_br(x, 2)}x")
    st.dataframe(vis, hide_index=True, width="stretch")
    if bc["tem_plano"]:
        acima = tb[tb["Plano ÷ Bass"] > 1.2]
        if acima.empty:
            st.success("O número de clientes do plano está dentro do que a difusão do produto permite.")
        else:
            st.warning(f"Nos anos {', '.join(str(a) for a in acima['Ano'])} o plano supõe mais clientes do que a curva de adoção "
                       "sugere. Ou o crescimento está otimista, ou o grupo precisa explicar o que acelera a adoção (parcerias, "
                       "indicação, mídia).")
    else:
        st.caption("Preencha o gasto anual por cliente em Tamanho do mercado para comparar com o plano.")
    proxima("adocao")


# ------------------------------------------------------------------------------
# RÉGUA DE PRONTIDÃO
# ------------------------------------------------------------------------------
def pg_prontidao():
    cabecalho("prontidao")
    res = R()
    pr = avaliar_prontidao(proj, d, res, MC(), preco_referencia(proj, res))
    st.metric("Nível de prontidão do negócio", f"{pr['nivel']} de 9")
    st.progress(pr["nivel"] / 9)
    st.caption("Inspirada nos níveis de prontidão usados por aceleradoras e programas de inovação: cada degrau só conta se os "
               "anteriores estiverem cumpridos.")
    for n_, nome, crit in NIVEIS_PRONTIDAO:
        ok = pr["cond"][n_]
        atingido = n_ <= pr["nivel"]
        ic = "✅" if atingido else ("🟡" if ok else "⬜")
        with st.container(border=True):
            st.markdown(f"{ic} **{n_}. {nome}**: {crit}")
            if n_ == pr["nivel"] + 1:
                st.caption("👉 Próximo degrau: é aqui que o grupo deve concentrar o esforço agora.")
            elif ok and not atingido:
                st.caption("Critério cumprido, mas um degrau anterior ainda está pendente.")
    proxima("prontidao")


# ------------------------------------------------------------------------------
# BANCA SIMULADA
# ------------------------------------------------------------------------------
def pg_banca():
    cabecalho("banca")
    res = R()
    perg = perguntas_banca(proj, d, res, MC(), preco_referencia(proj, res))
    st.caption("Perguntas que uma banca faria, geradas a partir dos pontos fracos do próprio plano. Escreva a resposta do grupo: "
               "o treino vale mais do que decorar o pitch. As respostas vão para o relatório e para o prompt de IA.")
    resp = {str(r["pergunta"]): str(r["resposta"]) for _, r in d["banca_respostas"].iterrows()}
    novas = {}
    for i, q in enumerate(perg):
        with st.container(border=True):
            st.markdown(f"**{i + 1}. [{q['tema']}]** {q['pergunta']}")
            st.caption(f"Por que a banca perguntaria: {q['motivo']}")
            novas[q["pergunta"]] = st.text_area("Resposta do grupo", resp.get(q["pergunta"], ""), key=wk(f"banca_{i}_{abs(hash(q['pergunta'])) % 10**6}"),
                                                label_visibility="collapsed", placeholder="Resposta com dados do plano...")
    mudou = any(novas[k] != resp.get(k, "") for k in novas)
    if mudou:
        todas = {**resp, **novas}
        df = pd.DataFrame({"pergunta": list(todas), "resposta": list(todas.values())})
        df = df[df["resposta"].astype(str).str.strip() != ""]
        salvar(conn, "banca_respostas", df, "pergunta")
        d["banca_respostas"] = limpar(df.reset_index(drop=True), "banca_respostas")
        marcar_salvo()
        invalidar()
    n_ok = sum(1 for k in novas if len(novas[k].strip()) > 20)
    st.metric("Perguntas respondidas", f"{n_ok} de {len(perg)}")
    proxima("banca")


# ------------------------------------------------------------------------------
# FORMALIZAÇÃO
# ------------------------------------------------------------------------------
def pg_formalizacao():
    cabecalho("formalizacao")
    res = R()
    ps = porte_sugerido(res, d)
    k = st.columns(2)
    k[0].metric("Receita prevista no Ano 1", brl(ps["receita"]))
    k[1].metric("Enquadramento sugerido", ps["porte"])
    for o in ps["obs"]:
        st.warning(o)
    st.caption(f"Limites usados: MEI até {brl(LIMITE_MEI)} por ano (com um sócio, no máximo um empregado e atividade permitida), "
               f"ME até {brl(LIMITE_ME)} e EPP até {brl(LIMITE_EPP)}. O teto do MEI está em discussão no Congresso (proposta aprovada "
               "na Câmara em 2026 prevê R$ 110 mil em 2027 e R$ 140 mil em 2028): confira a regra vigente no Portal do Empreendedor.")
    itens = itens_formalizacao(proj, d)
    faltam = [i for i in itens if i not in set(d["formalizacao"]["item"])]
    if faltam and st.button(f"Incluir {len(faltam)} item(ns) sugerido(s) para este negócio"):
        _semear("formalizacao", [{"item": i, "status": "Pendente", "observacao": ""} for i in faltam], "item")
    editor("formalizacao", {"item": T("Item", required=True, width="large"), "status": S("Situação", options=STATUS_FORMAL),
                            "observacao": T("Observação (custo, prazo, órgão)", width="large")}, "item")
    fo = d["formalizacao"]
    if not fo.empty:
        feito = float(fo["status"].isin(["Concluído", "Em andamento", "Não se aplica"]).mean())
        st.progress(feito, text=f"{pct(feito, 0)} encaminhado")
    proxima("formalizacao")

# ##############################################################################
# PARTE: pnbox_ui_paginas7.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB 8 - INTERFACE DAS NOVAS FERRAMENTAS
# ==============================================================================


# ------------------------------------------------------------------------------
# TRILHA GUIADA E "O QUE FALTA"
# ------------------------------------------------------------------------------
def proximo_passo() -> str | None:
    """Primeira ferramenta essencial que ainda não está concluída."""
    for p in ESSENCIAIS:
        if p in PAG and situacao(p) != "✅":
            return p
    for p in [x[0] for x in PAGINAS]:
        if situacao(p) == "◐":
            return p
    return None


def pg_trilha():
    cabecalho("trilha")
    st.caption("A trilha mostra a ordem recomendada de preenchimento. As ferramentas essenciais formam o plano mínimo; "
               "as demais aprofundam a análise.")
    prox = proximo_passo()
    if prox:
        with st.container(border=True):
            st.markdown(f"### 👉 Próximo passo: {PAG[prox][2]} {PAG[prox][3]}")
            st.caption(PAG[prox][4])
            if st.button("Ir para esta ferramenta", type="primary", key="trilha_ir"):
                ir(prox)
                st.rerun()
    else:
        st.success("Todas as ferramentas essenciais estão concluídas. Agora vale aprofundar a validação e preparar a defesa.")
    ess = [p for p in ESSENCIAIS if p in PAG]
    feitos = sum(1 for p in ess if situacao(p) == "✅")
    st.progress(feitos / len(ess), text=f"Plano mínimo: {feitos} de {len(ess)} ferramentas essenciais concluídas")
    st.subheader("O que falta, por ordem de importância")
    chk = CHK()
    pend = chk[chk["Situação"] != "OK"]
    n = 0
    for i, r in pend.iterrows():
        n += 1
        a, b = st.columns([4, 1.2])
        a.markdown(f"⚠️ **{r['Teste']}**: {r['Valor apurado']}  \n<small>{r['Como corrigir'] or r['Observação']}</small>",
                   unsafe_allow_html=True)
        if r["destino"] and b.button("Corrigir →", key=f"tri_chk_{i}", width="stretch"):
            ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir'] or r['Observação']}")
            st.rerun()
    for p in ess:
        s_ = situacao(p)
        if s_ != "✅":
            n += 1
            a, b = st.columns([4, 1.2])
            a.markdown(f"{'◐' if s_ == '◐' else '○'} **{PAG[p][3]}** ({'em andamento' if s_ == '◐' else 'não iniciada'})")
            if b.button("Abrir →", key=f"tri_pg_{p}", width="stretch"):
                ir(p)
                st.rerun()
    if n == 0:
        st.success("Nada pendente nas ferramentas essenciais e nas checagens.")
    st.subheader("Mapa completo")
    for g in GRUPOS:
        itens = [p for p in PAGINAS if p[1] == g]
        st.markdown(f"**{g}**: " + " · ".join(f"{situacao(p[0]) or '·'} {p[3]}" for p in itens))


# ------------------------------------------------------------------------------
# GLOSSÁRIO
# ------------------------------------------------------------------------------
def pg_glossario():
    cabecalho("glossario")
    try:
        st.subheader("📖 O que significam os números deste projeto")
        for nome, valor, leitura in explicar_indicadores(R()):
            with st.container(border=True):
                st.markdown(f"**{nome}: {valor}**  \n{leitura}")
    except Exception:
        st.info("A leitura dos números aparece quando houver produtos, custos e investimentos.")
    st.subheader("🔤 Termos")
    busca = st.text_input("Procurar termo", key="glos_busca", placeholder="Ex.: VPL, markup, churn")
    for termo, defi in sorted(GLOSSARIO.items(), key=lambda x: x[0].lower()):
        if busca and busca.lower() not in (termo + defi).lower():
            continue
        st.markdown(f"**{termo}**: {defi}")
    proxima("glossario")


# ------------------------------------------------------------------------------
# IMPORTAR E EXPORTAR PLANILHAS
# ------------------------------------------------------------------------------
TABELAS_IMPORTAVEIS = {
    "catalogo": ("Produtos e serviços", {"categoria": "Categoria", "descricao": "Descrição", "preco_venda": "Preço de venda",
                                         "custo_aquisicao": "Custo unitário", "qtd_mensal": "Quantidade por mês",
                                         "qtd_estoque_inicial": "Estoque inicial", "horas_unit": "Horas por unidade"}),
    "custos_fixos": ("Custos fixos", {"descricao": "Descrição", "classificacao": "Classificação", "valor_mensal": "Valor mensal"}),
    "inv_fixos": ("Investimentos fixos", {"descricao": "Descrição", "quantidade": "Quantidade", "valor_unitario": "Valor unitário",
                                          "vida_util_anos": "Vida útil (anos)", "valor_residual_pct": "Valor residual (%)"}),
    "materiais": ("Materiais", {"nome": "Material", "categoria": "Categoria", "unidade": "Unidade", "valor_unitario": "Valor unitário"}),
    "pessoal": ("Equipe", {"cargo": "Cargo", "area": "Área", "num_empregados": "Pessoas", "salario_mensal": "Salário",
                           "encargos_pct": "Encargos (%)", "beneficios_unit": "Benefícios", "mes_inicio": "Mês de início"}),
    "entrevistas": ("Entrevistas", {"data": "Data", "entrevistado": "Entrevistado", "perfil": "Perfil", "dores": "Dores",
                                    "solucao_atual": "Solução atual", "pagaria": "Pagaria", "valor_pagaria": "Valor que pagaria",
                                    "aprendizado": "Aprendizado"}),
    "precos_vw": ("Teste de preço (Van Westendorp)", {"respondente": "Respondente", "muito_barato": "Barato demais",
                                                      "barato": "Barato", "caro": "Caro", "muito_caro": "Caro demais"}),
    "concorrentes": ("Concorrentes", {"nome": "Nome"}),
    "riscos": ("Riscos", {"risco": "Risco", "probabilidade": "Probabilidade", "impacto": "Impacto", "mitigacao": "Mitigação",
                          "responsavel": "Responsável"}),
}


def pg_planilha():
    cabecalho("planilha")
    t1, t2 = st.tabs(["⬆️ Importar de planilha", "⬇️ Exportar para Excel"])
    with t1:
        st.caption("Traga listas grandes de uma planilha em vez de digitar linha a linha. A primeira linha da planilha precisa ter "
                   "os nomes das colunas; o programa reconhece os nomes do modelo, com ou sem acentos.")
        tab = st.selectbox("Para qual ferramenta?", list(TABELAS_IMPORTAVEIS), format_func=lambda k: TABELAS_IMPORTAVEIS[k][0],
                           key="imp_tab")
        nome_t, rot = TABELAS_IMPORTAVEIS[tab]
        modelo = pd.DataFrame(columns=list(rot.values()))
        b = io.BytesIO()
        with pd.ExcelWriter(b, engine="openpyxl") as w:
            modelo.to_excel(w, index=False, sheet_name=nome_t[:31])
        st.download_button(f"📄 Baixar planilha-modelo de {nome_t}", b.getvalue(), f"modelo_{tab}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        arq = st.file_uploader("Planilha preenchida (.xlsx ou .csv)", type=["xlsx", "csv"], key=wk(f"imp_up_{tab}"))
        if arq:
            try:
                bruto = ler_planilha(arq.name, arq.getvalue())
            except Exception as e:
                st.error(f"Não foi possível ler a planilha: {e}")
                return proxima("planilha")
            mapa = mapear_colunas(bruto, colunas_de(tab), rot)
            if not mapa:
                st.error("Nenhuma coluna reconhecida. Use os nomes da planilha-modelo.")
                return proxima("planilha")
            novo = bruto[list(mapa)].rename(columns=mapa)
            for c in colunas_de(tab):
                if c not in novo:
                    novo[c] = None
            tipos = dict(SCHEMA[tab])
            for c in novo.columns:
                if tipos.get(c) in ("REAL", "INTEGER"):
                    novo[c] = novo[c].map(converter_numero)
            novo = limpar(novo[colunas_de(tab)].dropna(how="all"), tab)
            st.write(f"**{len(novo)} linha(s)** reconhecida(s). Colunas usadas: " + ", ".join(f"{a} → {rot.get(b_, b_)}"
                                                                                              for a, b_ in mapa.items()))
            st.dataframe(novo.head(20), hide_index=True, width="stretch")
            modo = st.radio("Como incluir", ["Acrescentar às linhas atuais", "Substituir as linhas atuais"], horizontal=True,
                            key="imp_modo")
            if st.button("✅ Importar", type="primary", key="imp_ok"):
                base = novo if modo.startswith("Substituir") else pd.concat([d[tab], novo], ignore_index=True)
                recarregar_tabela(tab, base, colunas_de(tab)[0])
                st.success(f"{len(novo)} linha(s) importada(s) em {nome_t}.")
    with t2:
        st.caption("Uma pasta de trabalho com todas as tabelas do projeto, a DRE, o fluxo de caixa e os indicadores.")
        if st.button("Gerar planilha", key="exp_xlsx"):
            ss.xlsx_v8 = planilha_excel(R())
        if ss.get("xlsx_v8"):
            st.download_button("⬇️ Baixar planilha (Excel)", ss.xlsx_v8, f"{proj['nome_negocio'] or 'projeto'}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    proxima("planilha")


# ------------------------------------------------------------------------------
# HISTÓRICO DE VERSÕES DO PROJETO (com comparação de indicadores)
# ------------------------------------------------------------------------------
def indicadores_de_versao(vid: int) -> dict | None:
    linha = conn.execute("SELECT dados FROM versoes WHERE id=?", (vid,)).fetchone()
    if not linha:
        return None
    dados = json.loads(linha[0])
    p_ = dados.get("projeto", [{}])[0] if dados.get("projeto") else {}
    p_ = {**PROJETO_PADRAO, **{k: v for k, v in p_.items() if v is not None}}
    d_ = {}
    for t in SCHEMA:
        if t in ("projeto", "versoes"):
            continue
        cols = colunas_de(t)
        df = pd.DataFrame(dados.get(t, []))
        for c in cols:
            if c not in df:
                df[c] = None
        d_[t] = limpar(df[cols] if len(df) else pd.DataFrame(columns=cols), t)
    pv = dados.get("proposta_valor")
    d_["proposta_valor"] = pd.DataFrame(pv) if pv else d_.get("proposta_valor", pd.DataFrame())
    r_ = executar_motor(p_, d_)
    return {"Receita Ano 1": r_["anos"][1]["rec_bruta"], "Lucro Ano 1": r_["anos"][1]["resultado"],
            "Investimento": r_["investimento_total"], "VPL": r_["vpl"], "TIR": r_["tir"]}


def pg_historico():
    cabecalho("historico")
    st.caption("Cada versão guarda uma foto completa do projeto. Compare os indicadores para ver o efeito de cada mudança e "
               "restaure uma versão anterior se precisar. O programa também cria cópias automáticas a cada 10 minutos.")
    a, b = st.columns([3, 1])
    nome_v = a.text_input("Nome da versão", placeholder="Ex.: Depois de testar o preço", label_visibility="collapsed", key="hist_nome")
    if b.button("💾 Salvar versão agora", width="stretch", key="hist_salvar"):
        salvar_versao(nome_v.strip() or "Versão manual")
        st.rerun()
    vs = pd.read_sql("SELECT id, nome, data FROM versoes ORDER BY id DESC", conn)
    if vs.empty:
        st.info("Nenhuma versão salva ainda.")
        return proxima("historico")
    if st.button("📊 Comparar os indicadores de todas as versões", key="hist_comp"):
        linhas_ = []
        for _, v in vs.iterrows():
            try:
                ind = indicadores_de_versao(int(v["id"]))
            except Exception:
                ind = None
            if ind:
                linhas_.append({"Versão": v["nome"], "Data": v["data"], **ind})
        r_ = R()
        linhas_.insert(0, {"Versão": "▶ Projeto atual", "Data": "agora", "Receita Ano 1": r_["anos"][1]["rec_bruta"],
                           "Lucro Ano 1": r_["anos"][1]["resultado"], "Investimento": r_["investimento_total"],
                           "VPL": r_["vpl"], "TIR": r_["tir"]})
        ss.hist_comp = pd.DataFrame(linhas_)
    if ss.get("hist_comp") is not None:
        st.dataframe(ss.hist_comp, hide_index=True, width="stretch", column_config={
            c: N(format="R$ %.0f") for c in ("Receita Ano 1", "Lucro Ano 1", "Investimento", "VPL")} | {"TIR": N(format="percent")})
    for _, v in vs.iterrows():
        a, b, c = st.columns([4, 1, 1])
        a.markdown(f"**{v['nome']}** · {v['data']}")
        if b.button("Restaurar", key=f"hrest_{v['id']}", width="stretch"):
            restaurar_versao(int(v["id"]))
            ss.hist_comp = None
            st.toast("Versão restaurada. Uma cópia do estado anterior foi guardada.")
            st.rerun()
        if c.button("Apagar", key=f"hapv_{v['id']}", width="stretch"):
            conn.execute("DELETE FROM versoes WHERE id=?", (int(v["id"]),))
            conn.commit()
            ss.hist_comp = None
            st.rerun()
    proxima("historico")


# ------------------------------------------------------------------------------
# FINANCIAMENTO: SAC, PRICE, CARÊNCIA E COMPARAÇÃO DE PROPOSTAS
# ------------------------------------------------------------------------------
def pg_financiamento():
    cabecalho("financiamento")
    t1, t2 = st.tabs(["🏦 Empréstimos do plano", "⚖️ Comparar propostas de crédito"])
    with t1:
        st.caption("Estes empréstimos entram no investimento, no fluxo de caixa e na DRE. Escolha o sistema de amortização e a "
                   "carência (meses em que se paga só os juros).")
        editor("emprestimos", {"descricao": T("Linha de crédito", required=True, width="large"), "valor": moeda("Valor (R$)"),
                               "taxa_anual_pct": N("Juros (% ao ano)", min_value=0.0), "prazo_meses": N("Prazo total (meses)", min_value=1, step=1),
                               "carencia_meses": N("Carência (meses)", min_value=0, step=1),
                               "sistema": S("Sistema", options=["Price", "SAC"])}, "descricao")
        emp = d["emprestimos"]
        for _, r in emp.iterrows():
            if float(r["valor"] or 0) <= 0:
                continue
            rs = resumo_emprestimo(r["valor"], r["taxa_anual_pct"], r["prazo_meses"], r["carencia_meses"], r["sistema"])
            with st.expander(f"📄 {r['descricao']}: parcela de {brl(rs['primeira'])} a {brl(rs['menor'])} · juros totais {brl(rs['juros'])}"):
                k = st.columns(4)
                k[0].metric("Total pago", brl(rs["total_pago"]))
                k[1].metric("Juros totais", brl(rs["juros"]))
                k[2].metric("Maior parcela", brl(rs["maior"]))
                k[3].metric("CET (a.a.)", pct(rs["cet"]))
                g = rs["tabela"].melt(id_vars="Mês", value_vars=["Juros", "Amortização"], var_name="Parte", value_name="R$")
                st.altair_chart(alt.Chart(g).mark_bar().encode(x="Mês:O", y="R$:Q", color=alt.Color("Parte:N", legend=alt.Legend(orient="bottom")))
                                .properties(height=220), width="stretch")
                st.dataframe(rs["tabela"], hide_index=True, width="stretch", height=240,
                             column_config={c: N(format="R$ %.2f") for c in ("Juros", "Amortização", "Parcela", "Saldo devedor")})
    with t2:
        st.caption("Compare propostas antes de decidir. As taxas mudam com frequência: use as condições informadas pelo banco "
                   "ou pela cooperativa na data da simulação.")
        base = ss.setdefault("propostas_credito", pd.DataFrame([
            {"Proposta": "Proposta A (exemplo)", "Valor": 30000.0, "Juros (% a.a.)": 18.0, "Prazo (meses)": 36, "Carência (meses)": 6, "Sistema": "Price"},
            {"Proposta": "Proposta B (exemplo)", "Valor": 30000.0, "Juros (% a.a.)": 22.0, "Prazo (meses)": 24, "Carência (meses)": 0, "Sistema": "SAC"}]))
        ed = st.data_editor(base, num_rows="dynamic", hide_index=True, width="stretch", key="ed_propostas",
                            column_config={"Sistema": S(options=["Price", "SAC"]), "Valor": N(format="R$ %.2f")})
        ss.propostas_credito = ed
        linhas_ = []
        for _, r in ed.dropna(subset=["Valor"]).iterrows():
            try:
                rs = resumo_emprestimo(r["Valor"], r["Juros (% a.a.)"], r["Prazo (meses)"], r["Carência (meses)"], r["Sistema"])
            except Exception:
                continue
            linhas_.append({"Proposta": r["Proposta"], "Primeira parcela": rs["primeira"], "Maior parcela": rs["maior"],
                            "Total pago": rs["total_pago"], "Juros totais": rs["juros"], "CET (a.a.)": rs["cet"]})
        if linhas_:
            cmp_ = pd.DataFrame(linhas_)
            st.dataframe(cmp_, hide_index=True, width="stretch", column_config={
                c: N(format="R$ %.2f") for c in ("Primeira parcela", "Maior parcela", "Total pago", "Juros totais")} | {"CET (a.a.)": N(format="percent")})
            mel = cmp_.sort_values("CET (a.a.)").iloc[0]
            st.info(f"Menor custo efetivo: **{mel['Proposta']}** ({pct(mel['CET (a.a.)'])} ao ano). Confira também se a parcela cabe "
                    "no fluxo de caixa dos primeiros meses.")
    proxima("financiamento")


# ------------------------------------------------------------------------------
# DADOS PÚBLICOS DO MUNICÍPIO (IBGE) E POF
# ------------------------------------------------------------------------------
def _gravar_dado(chave, rotulo, valor, periodo, fonte):
    dp = d["dados_publicos"]
    dp = dp[dp["chave"] != chave]
    dp = pd.concat([dp, pd.DataFrame([{"chave": chave, "rotulo": rotulo, "valor": valor, "periodo": periodo, "fonte": fonte}])],
                   ignore_index=True)
    salvar(conn, "dados_publicos", dp, "chave")
    d["dados_publicos"] = limpar(dp.reset_index(drop=True), "dados_publicos")


def pg_dados_publicos():
    cabecalho("dados_publicos")
    st.caption("Busca dados oficiais do IBGE pela internet (gratuito) para dimensionar o mercado com números verificáveis. "
               "Os valores ficam guardados no projeto e entram no relatório com a fonte.")
    t1, t2 = st.tabs(["🏙️ Município (Censo 2022, PIB, empresas)", "🛒 Gastos das famílias (POF 2017-2018)"])
    with t1:
        c = st.columns([1, 3])
        ufs = list(UFS)
        uf = c[0].selectbox("UF", ufs, index=ufs.index(proj["uf_sigla"]) if proj["uf_sigla"] in ufs else ufs.index("MG"), key=wk("dp_uf"))
        chave_m = f"mun_{uf}"
        if chave_m not in ss:
            try:
                ss[chave_m] = municipios_da_uf(uf)
            except Exception as e:
                ss[chave_m] = []
                st.warning(f"Não foi possível consultar a lista de municípios do IBGE agora ({e}). Verifique a internet.")
        lista = ss[chave_m]
        nomes = [n for _, n in lista]
        idx = nomes.index(proj["municipio_nome"]) if proj["municipio_nome"] in nomes else 0
        nome = c[1].selectbox("Município", nomes, index=idx, key=wk("dp_mun")) if nomes else None
        if nome and st.button("🔎 Buscar dados do IBGE", type="primary", key="dp_buscar"):
            cod = dict((n, i) for i, n in lista)[nome]
            with st.spinner("Consultando a API do IBGE..."):
                dados = dados_municipio(cod)
            if not dados:
                st.error("O IBGE não respondeu. Tente de novo em alguns minutos.")
            else:
                for ch, rot, v, per, fonte in dados:
                    _gravar_dado(ch, rot, v, per, fonte)
                auto_unica("projeto", {"uf_sigla": uf, "municipio_nome": nome, "municipio_ibge": cod})
                st.rerun()
        dp = d["dados_publicos"]
        mun = dp[~dp["chave"].astype(str).str.startswith("pof")]
        if not mun.empty:
            st.markdown(f"**{proj['municipio_nome'] or ''} ({proj['uf_sigla'] or ''})**")
            cols = st.columns(3)
            for i, (_, r) in enumerate(mun.iterrows()):
                v = r["valor"]
                txt = brl(v) if "R$" in r["rotulo"] and "mil" not in r["rotulo"] else num_br(v, 1 if v < 100 else 0)
                cols[i % 3].metric(f"{r['rotulo']} ({r['periodo']})", txt, help=r["fonte"])
            m = {r["chave"]: r["valor"] for _, r in mun.iterrows()}
            st.markdown("**Usar no Tamanho do mercado (TAM)**")
            b = st.columns(3)
            if "populacao" in m and b[0].button(f"TAM = população ({num_br(m['populacao'], 0)})", key="dp_tam_pop"):
                auto_unica("projeto", {"tam_clientes": m["populacao"], "tam_desc": f"População de {proj['municipio_nome']} (Censo 2022, IBGE)"})
                st.toast("TAM atualizado.")
            if "domicilios" in m and b[1].button(f"TAM = domicílios ({num_br(m['domicilios'], 0)})", key="dp_tam_dom"):
                auto_unica("projeto", {"tam_clientes": m["domicilios"], "tam_desc": f"Domicílios de {proj['municipio_nome']} (Censo 2022, IBGE)"})
                st.toast("TAM atualizado.")
            if "unidades_locais" in m and b[2].button(f"TAM = empresas ({num_br(m['unidades_locais'], 0)})", key="dp_tam_emp"):
                auto_unica("projeto", {"tam_clientes": m["unidades_locais"],
                                       "tam_desc": f"Unidades locais de empresas em {proj['municipio_nome']} (Cempre, IBGE)"})
                st.toast("TAM atualizado.")
    with t2:
        uf2 = proj["uf_sigla"] or "MG"
        st.caption(f"Despesa média mensal das famílias por tipo de gasto em {UFS.get(uf2, (0, 'Brasil'))[1]} (POF 2017-2018). "
                   "Útil para estimar quanto um domicílio gasta por ano na categoria do seu negócio.")
        if st.button("Carregar gastos das famílias", key="dp_pof"):
            with st.spinner("Consultando a POF no IBGE..."):
                ss.pof_df, ss.pof_origem = pof_uf(uf2)
        if ss.get("pof_df") is not None:
            pof = ss.pof_df
            st.caption(f"Origem: {ss.pof_origem}. Valores de 2018, sem correção pela inflação.")
            esc = st.selectbox("Categoria de gasto do seu negócio", list(pof["Tipo de despesa"]), key="dp_pof_cat")
            v = float(pof.loc[pof["Tipo de despesa"] == esc, "R$ por mês por família"].iloc[0])
            k = st.columns(2)
            k[0].metric("Por mês, por família", brl(v))
            k[1].metric("Por ano, por família", brl(v * 12))
            if st.button("Guardar e usar como gasto anual por cliente", key="dp_pof_usar"):
                _gravar_dado("pof_" + _norm(esc)[:30], esc.split(" ", 1)[-1], v, "2017-2018",
                             f"IBGE. Pesquisa de Orçamentos Familiares 2017-2018, tabela 6715 ({UFS.get(uf2, (0, 'Brasil'))[1]}).")
                auto_unica("projeto", {"gasto_anual_cliente": round(v * 12, 2)})
                st.toast("Gasto anual por cliente atualizado.")
            st.dataframe(pof, hide_index=True, width="stretch", height=300, column_config={"R$ por mês por família": N(format="R$ %.2f")})
    proxima("dados_publicos")


# ------------------------------------------------------------------------------
# REFERÊNCIAS DO SETOR E SOBREVIVÊNCIA
# ------------------------------------------------------------------------------
def pg_referencias():
    cabecalho("referencias")
    res = R()
    st.caption("Compare os números do plano com a média das empresas brasileiras da mesma atividade, segundo o IBGE. Diferenças "
               "grandes não são erro por si só, mas precisam de explicação.")
    ops = atividades_referencia()
    atual = proj["setor_referencia"] if proj["setor_referencia"] in ops else ops[0]
    esc = st.selectbox("Atividade mais parecida com o seu negócio", ops, index=ops.index(atual), key=wk("ref_set"))
    auto_unica("projeto", {"setor_referencia": esc})
    if res["anos"][1]["rec_bruta"] > 0:
        cs = comparar_setor(esc, proj, d, res)
        vis = cs.copy()
        for c in ("Plano", "Setor (Brasil)"):
            vis[c] = [pct(v) if ("÷" in i or "margem" in i.lower()) else brl(v) for i, v in zip(cs["Indicador"], cs[c])]
        vis["Plano ÷ setor"] = cs["Plano ÷ setor"].map(lambda v: f"{num_br(v, 2)}x" if v == v else "n.a.")
        st.dataframe(vis, hide_index=True, width="stretch")
        for a in leitura_setor(cs):
            st.warning(a)
        st.caption(FONTE_PAC if esc.startswith("Comércio") else FONTE_PAS)
    st.subheader("Quantas empresas sobrevivem")
    k = st.columns(3)
    for col, (rot, v) in zip(k, SOBREVIVENCIA_5A.items()):
        col.metric(rot, pct(v, 1), help="Percentual de empresas ainda ativas cinco anos depois de abertas.")
    st.caption("Fonte: Sebrae, seminário DataSebrae \"Aplicações de análise de sobrevivência\" (jul. 2026), com dados da Receita Federal.")
    st.markdown("**Causas mais citadas por quem fechou** (pesquisa Sobrevivência de Empresas, Sebrae): " +
                "; ".join(f"{c} ({pct(v, 0)})" for c, v in CAUSAS_FECHAMENTO) + ".")
    a1 = res["anos"][1]
    minimo = float(res["fluxo_mensal"]["Saldo final"].min())
    if minimo < res["reserva"] * 0.5:
        st.warning(f"Atenção ao capital de giro: o menor saldo de caixa do Ano 1 é {brl(minimo)}. É a principal causa de fechamento.")
    if a1["rec_bruta"] and a1["rec_bruta"] < a1["pe_contabil"] * 1.2:
        st.warning("A receita prevista está a menos de 20% do ponto de equilíbrio: pouca folga se as vendas vierem abaixo do esperado.")
    proxima("referencias")


# ------------------------------------------------------------------------------
# SWOT CRUZADA
# ------------------------------------------------------------------------------
def pg_swot_cruzada():
    cabecalho("swot_cruzada")
    st.caption("A SWOT só vale se virar estratégia (Wright; Kroll; Parnell, 2000). Cruze os fatores priorizados e escreva o que a "
               "empresa vai fazer em cada combinação.")
    sug = sugestoes_cruzada(d)
    cols = st.columns(2)
    for i, (q, s_) in enumerate(sug.items()):
        with cols[i % 2].container(border=True):
            st.markdown(f"**{q}**  \n<small>{s_['dica']}</small>", unsafe_allow_html=True)
            st.caption(("Fatores: " + "; ".join(s_["a"][:3]) + "  ×  " + "; ".join(s_["b"][:3])) if s_["a"] and s_["b"]
                       else "Cadastre os fatores na análise SWOT.")
            n = int((d["swot_estrategias"]["quadrante"] == q).sum())
            st.caption(f"{n} estratégia(s) registrada(s)")
    if d["swot_estrategias"].empty and any(s_["pares"] for s_ in sug.values()):
        if st.button("Criar uma linha para cada combinação de fatores priorizados", key="swc_seed"):
            linhas_ = [{"quadrante": q, "estrategia": f"[{p}] ", "acao": "", "prazo_meses": 6.0}
                       for q, s_ in sug.items() for p in s_["pares"][:2]]
            recarregar_tabela("swot_estrategias", pd.DataFrame(linhas_), "estrategia")
            st.rerun()
    editor("swot_estrategias", {"quadrante": S("Tipo de estratégia", options=list(QUADRANTES_CRUZADA), required=True, width="medium"),
                                "estrategia": T("Estratégia", required=True, width="large"), "acao": T("Ação concreta", width="large"),
                                "prazo_meses": N("Prazo (meses)", min_value=1, step=1)}, "estrategia")
    proxima("swot_cruzada")


# ------------------------------------------------------------------------------
# LEAN CANVAS
# ------------------------------------------------------------------------------
def html_lean(bl: dict) -> str:
    cores = {"Problema": "#c0392b", "Alternativas existentes": "#c0392b", "Solução": "#2c7fb8", "Métricas-chave": "#2c7fb8",
             "Proposta única de valor": "#44923a", "Vantagem injusta": "#7b3fe4", "Canais": "#14a49a",
             "Segmentos de clientes": "#1592c4", "Primeiros clientes (early adopters)": "#1592c4",
             "Estrutura de custos": "#4456a0", "Fontes de receita": "#22617a"}
    areas = {"Problema": "pr", "Alternativas existentes": "al", "Solução": "so", "Métricas-chave": "me",
             "Proposta única de valor": "pv", "Vantagem injusta": "va", "Canais": "ca", "Segmentos de clientes": "se",
             "Primeiros clientes (early adopters)": "ea", "Estrutura de custos": "cu", "Fontes de receita": "re"}

    def caixa(nome):
        itens = "".join(f"<li>{_esc(x)}</li>" for x in bl.get(nome, [])) or "<li><i>a preencher</i></li>"
        return (f"<div class='lc-b' style='grid-area:{areas[nome]};background:{cores[nome]}'><div class='lc-t'>{nome.upper()}</div>"
                f"<ul>{itens}</ul></div>")
    css = """<style>
    .lc{display:grid;gap:8px;grid-template-columns:repeat(10,1fr);grid-template-areas:
      "pr pr so so pv pv va va se se" "pr pr so so pv pv va va se se" "al al me me pv pv ca ca ea ea" "cu cu cu cu cu re re re re re";}
    .lc-b{border-radius:12px;padding:12px;color:#fff;font-size:.83rem;min-height:110px}
    .lc-t{font-weight:800;font-size:.9rem;margin-bottom:6px} .lc-b ul{margin:0;padding-left:18px}
    @media (max-width:700px){.lc{grid-template-columns:1fr;grid-template-areas:"pr" "al" "se" "ea" "pv" "so" "ca" "re" "cu" "me" "va";}}
    </style>"""
    return css + "<div class='lc'>" + "".join(caixa(n) for n in areas) + "</div>"


def pg_lean():
    cabecalho("lean")
    st.caption("O Lean Canvas (Maurya, 2012) adapta o Business Model Canvas para negócios novos e incertos: começa pelo problema e "
               "pelas métricas que mostram se a solução funciona. Ele é montado com o que você já preencheu; complete os três campos abaixo.")
    c = st.columns(3)
    v = {"lean_alternativas": c[0].text_area("Alternativas existentes (uma por linha)", proj["lean_alternativas"], height=110,
                                             key=wk("lean_alt"), help="Como os clientes resolvem o problema hoje."),
         "lean_metricas": c[1].text_area("Métricas-chave (uma por linha)", proj["lean_metricas"], height=110, key=wk("lean_met"),
                                         help="Os poucos números que mostram se o negócio está funcionando."),
         "lean_vantagem": c[2].text_area("Vantagem injusta (uma por linha)", proj["lean_vantagem"], height=110, key=wk("lean_van"),
                                         help="Algo que não pode ser copiado ou comprado com facilidade.")}
    auto_unica("projeto", v)
    st.markdown(html_lean(lean_canvas(proj, d, R())), unsafe_allow_html=True)
    proxima("lean")


# ------------------------------------------------------------------------------
# VALUATION E RODADA DE INVESTIMENTO
# ------------------------------------------------------------------------------
def pg_valuation():
    cabecalho("valuation")
    res = R()
    st.caption("Quanto vale o negócio? Três métodos dão uma faixa de valor, não um número exato. A rodada mostra quanto da empresa "
               "um investidor levaria em troca do dinheiro.")
    c = st.columns(3)
    g = c[0].number_input("Crescimento na perpetuidade (% a.a.)", -5.0, 15.0, float(proj["val_crescimento_perp_pct"] or 3), 0.5, key=wk("val_g"),
                          help="Quanto o fluxo cresce por ano depois do horizonte. Deve ser menor que a TMA.")
    meb = c[1].number_input("Múltiplo de EBITDA", 0.5, 20.0, float(proj["val_multiplo_ebitda"] or 4), 0.5, key=wk("val_meb"))
    mrec = c[2].number_input("Múltiplo de receita", 0.1, 10.0, float(proj["val_multiplo_receita"] or 1), 0.1, key=wk("val_mrec"))
    auto_unica("projeto", {"val_crescimento_perp_pct": g, "val_multiplo_ebitda": meb, "val_multiplo_receita": mrec})
    vl = valuation(res, proj, d)
    st.dataframe(vl["metodos"], hide_index=True, width="stretch", column_config={"Valor do patrimônio dos sócios": N(format="R$ %.0f")})
    k = st.columns(3)
    k[0].metric("Valor presente dos fluxos do horizonte", brl(vl["vp_fluxos"]))
    k[1].metric("Valor presente da perpetuidade", brl(vl["vp_terminal"]))
    k[2].metric("Dívida (empréstimos)", brl(vl["divida"]))
    if vl["vp_terminal"] > vl["vp_fluxos"]:
        st.info("Mais da metade do valor vem da perpetuidade, ou seja, de depois do horizonte do plano. Isso é comum, mas torna o "
                "valor muito sensível à taxa de crescimento escolhida.")
    st.caption("Múltiplos usados em pequenos negócios costumam ser baixos e variam muito por setor; trate-os como referência didática.")
    st.subheader("Rodada de investimento")
    c = st.columns(2)
    pre = c[0].number_input("Valor da empresa antes do aporte (pre-money, R$)", 0.0, 1e10,
                            float(proj["rodada_pre_money"] or max(vl["equity_dcf"], 0.0)), 1000.0, key=wk("rod_pre"))
    inv = c[1].number_input("Valor do aporte do investidor (R$)", 0.0, 1e10, float(proj["rodada_investimento"] or 0), 1000.0, key=wk("rod_inv"))
    auto_unica("projeto", {"rodada_pre_money": pre, "rodada_investimento": inv})
    ct = cap_table(d, pre, inv)
    st.dataframe(ct, hide_index=True, width="stretch", column_config={
        "Participação antes": N(format="percent"), "Participação depois": N(format="percent"), "Valor da participação": N(format="R$ %.0f")})
    if inv > 0 and pre > 0:
        st.markdown(f"Com post-money de **{brl(pre + inv)}**, o investidor fica com **{pct(inv / (pre + inv))}** da empresa e cada 1% "
                    f"vale **{brl((pre + inv) / 100)}**.")
    proxima("valuation")


# ------------------------------------------------------------------------------
# IMPACTO SOCIAL E ODS
# ------------------------------------------------------------------------------
def pg_impacto():
    cabecalho("impacto")
    st.caption("Para negócios de impacto, editais e projetos de extensão: a teoria da mudança liga o que a empresa faz às "
               "transformações que pretende causar, e os indicadores mostram se isso está acontecendo.")
    v = {}
    cols = st.columns(3)
    for i, (k_, rot, ajuda) in enumerate(ETAPAS_MUDANCA):
        v[k_] = cols[i % 3].text_area(f"{i + 1}. {rot}", proj[k_], height=110, key=wk(k_), help=ajuda)
    ods_atual = [o for o in str(proj["ods"] or "").split(";") if o in ODS]
    sel = st.multiselect("Objetivos de Desenvolvimento Sustentável (ODS) relacionados", ODS, default=ods_atual, key=wk("ods"))
    v["ods"] = ";".join(sel)
    auto_unica("projeto", v)
    if any(proj[k_] for k_, _, _ in ETAPAS_MUDANCA):
        st.markdown(" ➜ ".join(f"**{rot}**: {proj[k_]}" for k_, rot, _ in ETAPAS_MUDANCA if proj[k_]))
    st.subheader("Indicadores de impacto")
    editor("impacto_indicadores", {"indicador": T("Indicador", required=True, width="large"),
                                   "tipo": S("Tipo", options=TIPOS_INDICADOR_IMPACTO), "linha_base": T("Situação atual"),
                                   "meta": T("Meta"), "ods": S("ODS", options=sel or ODS)}, "indicador")
    proxima("impacto")


# ------------------------------------------------------------------------------
# EDITAIS DE FOMENTO
# ------------------------------------------------------------------------------
def pg_editais():
    cabecalho("editais")
    st.warning("Estrutura baseada nas seções usuais de editais de inovação, como o Programa Centelha. Cada edital tem campos e "
               "limites próprios: ajuste os limites de caracteres conforme o edital vigente antes de enviar.")
    res = R()
    trl = st.select_slider("Nível de maturidade tecnológica (TRL)", options=list(TRL), value=int(proj["edital_trl"] or 3),
                           format_func=lambda n: f"{n}: {TRL[n]}", key=wk("trl"))
    auto_unica("projeto", {"edital_trl": trl})
    ed = d["edital_textos"]
    if ed.empty or st.button("🔄 Refazer o rascunho a partir do plano (substitui os textos)", key="edt_refazer"):
        ras = rascunho_edital(proj, d, res)
        base = pd.DataFrame([{"secao": s_, "texto": ras.get(s_, ""), "limite": float(lim)} for s_, lim in SECOES_EDITAL])
        recarregar_tabela("edital_textos", base, "secao")
        st.rerun()
    novos = []
    mudou = False
    for i, r in ed.iterrows():
        lim = int(r["limite"] or 2000)
        with st.container(border=True):
            c1, c2 = st.columns([4, 1])
            c1.markdown(f"**{r['secao']}**")
            nl = c2.number_input("Limite", 50, 20000, lim, 50, key=wk(f"edl_{i}"), label_visibility="collapsed")
            txt = st.text_area(r["secao"], r["texto"], height=120, key=wk(f"edt_{i}"), label_visibility="collapsed")
            n = len(txt)
            st.caption(("🔴 " if n > nl else "🟢 ") + f"{n} de {nl} caracteres")
            novos.append({"secao": r["secao"], "texto": txt, "limite": float(nl)})
            mudou = mudou or txt != r["texto"] or nl != lim
    if mudou:
        df = pd.DataFrame(novos)
        salvar(conn, "edital_textos", df, "secao")
        d["edital_textos"] = limpar(df, "edital_textos")
        marcar_salvo()
    st.subheader("Orçamento do projeto")
    editor("edital_orcamento", {"rubrica": S("Rubrica", options=RUBRICAS_EDITAL, required=True), "item": T("Item", width="large"),
                                "valor": moeda("Valor (R$)")}, "item")
    orc = d["edital_orcamento"]
    if not orc.empty:
        st.dataframe(orc.groupby("rubrica", as_index=False)["valor"].sum().rename(columns={"rubrica": "Rubrica", "valor": "Total"}),
                     hide_index=True, column_config={"Total": N(format="R$ %.2f")})
        st.metric("Total solicitado", brl(orc["valor"].sum()))
    texto = "\n\n".join(f"{r['secao'].upper()}\n{r['texto']}" for _, r in d["edital_textos"].iterrows())
    if not orc.empty:
        texto += "\n\nORÇAMENTO\n" + "\n".join(f"{r['rubrica']} - {r['item']}: {brl(r['valor'])}" for _, r in orc.iterrows())
    st.download_button("⬇️ Baixar texto da proposta (.txt)", texto.encode("utf-8"), "proposta_edital.txt", "text/plain")
    proxima("editais")


# ------------------------------------------------------------------------------
# NOVIDADES DO PROGRAMA (histórico de versões do ViabilizaLab)
# ------------------------------------------------------------------------------
def pg_novidades():
    cabecalho("novidades")
    st.caption(f"Você está usando a versão **{VERSAO_PROGRAMA}**. O histórico completo também está no arquivo CHANGELOG.md do "
               "repositório no GitHub.")
    st.markdown(NOVIDADES_MD)


# ------------------------------------------------------------------------------
# COMENTÁRIOS DO PROFESSOR (aparecem no topo de cada ferramenta)
# ------------------------------------------------------------------------------
def comentarios_da_pagina(p: str):
    cm = d.get("comentarios")
    if cm is None:
        return
    da_pag = cm[(cm["pagina"] == p) & (cm["resolvido"].fillna(0).astype(int) == 0)]
    for i, r in da_pag.iterrows():
        with st.container(border=True):
            a, b = st.columns([5, 1])
            a.markdown(f"💬 **Comentário de {r['autor'] or 'professor(a)'}** ({r['data']}): {r['comentario']}")
            if b.button("Resolvido ✓", key=f"cm_ok_{p}_{i}"):
                cm.loc[i, "resolvido"] = 1
                salvar(conn, "comentarios", cm, "comentario")
                d["comentarios"] = limpar(cm.reset_index(drop=True), "comentarios")
                st.rerun()
    if ss.get("modo_correcao"):
        with st.expander("✍️ Escrever comentário nesta ferramenta (modo de correção)"):
            txt = st.text_area("Comentário", key=f"cm_txt_{p}")
            if st.button("Adicionar comentário", key=f"cm_add_{p}") and txt.strip():
                novo = pd.DataFrame([{"pagina": p, "comentario": txt.strip(), "autor": ss.get("prof_nome") or "Professor(a)",
                                      "data": datetime.now().strftime("%d/%m/%Y"), "resolvido": 0}])
                cm2 = pd.concat([cm, novo], ignore_index=True)
                salvar(conn, "comentarios", cm2, "comentario")
                d["comentarios"] = limpar(cm2, "comentarios")
                st.rerun()


# ------------------------------------------------------------------------------
# MODO PROFESSOR: SEMELHANÇA ENTRE PLANOS E RODADA DE INVESTIMENTO DA TURMA
# ------------------------------------------------------------------------------
def professor_extras(projetos: list):
    """projetos: lista de (rotulo, proj, d) já carregados no modo professor."""
    st.subheader("🧭 Prontidão e evidências da turma")
    linhas_ = []
    for rot, p_, d_ in projetos:
        try:
            r_ = executar_motor(p_, d_)
            pl = placar_evidencias(p_, d_, r_)
            mc = monte_carlo(r_, faixas_mc(d_), n=800) if r_["anos"][1]["rec_bruta"] > 0 else None
            pr = avaliar_prontidao(p_, d_, r_, mc, preco_referencia(p_, r_))
            linhas_.append({"Projeto": rot, "Prontidão (1-9)": pr["nivel"], "Placar de evidências": pl["score"],
                            "Entrevistas": len(d_["entrevistas"]), "Chance de VPL > 0": mc["p_vpl_pos"] if mc else float("nan"),
                            "Comentários abertos": int((d_["comentarios"]["resolvido"].fillna(0).astype(int) == 0).sum())})
        except Exception:
            continue
    if linhas_:
        st.dataframe(pd.DataFrame(linhas_), hide_index=True, width="stretch", column_config={
            "Placar de evidências": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "Chance de VPL > 0": N(format="percent")})
    if len(projetos) >= 2:
        st.subheader("🔁 Planos parecidos entre si")
        st.caption("Compara trechos de texto (sequências de três palavras) e valores numéricos. Semelhança alta pede uma conversa "
                   "com os grupos; não é prova de cópia.")
        sm = semelhanca_planos(projetos)
        st.dataframe(sm, hide_index=True, width="stretch", column_config={
            "Textos em comum": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "Números em comum": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1)})
        alto = sm[(sm["Textos em comum"] > 0.35) | (sm["Números em comum"] > 0.5)]
        if not alto.empty:
            st.warning(f"{len(alto)} par(es) com semelhança alta.")
        st.subheader("💰 Rodada de investimento da turma")
        st.caption("Cada grupo recebe um capital fictício e distribui entre os planos dos colegas (o próprio plano não conta). "
                   "Lance os valores combinados em sala.")
        cap = st.number_input("Capital fictício por grupo (R$)", 1000.0, 1e9, 100000.0, 10000.0, key="rod_cap")
        nomes = [r for r, _, _ in projetos]
        if ss.get("rodada_nomes") != nomes:
            ss.rodada_df = pd.DataFrame(0.0, index=nomes, columns=nomes)
            ss.rodada_nomes = nomes
        ed = st.data_editor(ss.rodada_df, width="stretch", key="ed_rodada",
                            column_config={c: N(format="R$ %.0f", min_value=0.0) for c in nomes})
        ss.rodada_df = ed
        excedeu = [i for i in ed.index if float(ed.loc[i].drop(i, errors="ignore").sum()) > cap]
        if excedeu:
            st.error("Passaram do capital disponível: " + ", ".join(excedeu))
        st.dataframe(resultado_rodada(ed), hide_index=True, width="stretch", column_config={"Capital recebido": N(format="R$ %.0f")})

# ##############################################################################
# PARTE: pnbox_ui_paginas8.py
# ##############################################################################




ICONE_NIVEL = {"ok": "✅", "atencao": "⚠️", "critico": "⛔"}


def mostrar_alertas(alertas):
    for nivel, msg in alertas:
        {"ok": st.success, "atencao": st.warning, "critico": st.error}[nivel](msg, icon=ICONE_NIVEL[nivel])


def _select(rotulo, campo, opcoes_, ajuda=None):
    atual = proj.get(campo) or ""
    lista = [""] + list(opcoes_)
    return st.selectbox(rotulo, lista, index=lista.index(atual) if atual in lista else 0,
                        format_func=lambda x: x or "(escolha)", key=wk(campo), help=ajuda)


def _aba_produto():
    st.caption("Kotler e Keller (2012) descrevem o produto em níveis: o benefício que o cliente compra, "
               "o produto em si e o que vem junto com ele.")
    v = {}
    c1, c2, c3 = st.columns(3)
    v["mix_nucleo"] = c1.text_area("🎯 Benefício central: o que o cliente realmente compra",
                                   proj.get("mix_nucleo") or "", height=120, key=wk("mx_n"),
                                   placeholder="Ex.: não é um bolo, é a comemoração resolvida sem trabalho.")
    v["mix_real"] = c2.text_area("📦 Produto em si: marca, qualidade, embalagem, linhas",
                                 proj.get("mix_real") or "", height=120, key=wk("mx_r"))
    v["mix_ampliado"] = c3.text_area("🎁 Produto ampliado: garantia, entrega, suporte, pós-venda",
                                     proj.get("mix_ampliado") or "", height=120, key=wk("mx_a"))
    v["mix_ciclo"] = _select("Fase do ciclo de vida do produto no mercado", "mix_ciclo", CICLO_VIDA)
    if v["mix_ciclo"]:
        st.info(CICLO_VIDA[v["mix_ciclo"]], icon="📈")
    v["mkt_produto"] = st.text_area("📝 Justificativa e detalhes do produto", proj["mkt_produto"], height=110,
                                    key=wk("m1"))
    auto_unica("projeto", v)
    cat = R()["cat"]
    if not cat.empty:
        st.markdown("**Portfólio (do catálogo de produtos)**")
        pt = pareto_produtos(cat)
        a, b, c = st.columns(3)
        a.metric("Itens no catálogo", num_br(pt["total"], 0))
        b.metric("Itens que fazem 80% da receita", num_br(pt["n"], 0))
        c.metric("Peso do item principal", pct(pt.get("share_top", float("nan"))))
        if pt["n"]:
            st.caption("Itens que sustentam a receita: " + ", ".join(pt["itens"]))
    dif = [proj.get("vpc_criadores"), proj.get("posic_diferencial")]
    if not any(str(x or "").strip() for x in dif):
        st.warning("Os diferenciais do produto ainda não aparecem no canvas da proposta de valor nem no "
                   "posicionamento. Um diferencial que só está no texto do 4Ps é difícil de defender na banca.",
                   icon="⚠️")


def _aba_preco():
    res = R()
    cat = res["cat"]
    c1, c2 = st.columns(2)
    with c1:
        est = _select("Estratégia de preço", "mix_estrategia_preco", ESTRATEGIAS_PRECO)
        if est:
            st.caption(ESTRATEGIAS_PRECO[est][0])
    with c2:
        pos = _select("Posição de preço em relação ao mercado", "mix_posicao_preco", POSICOES_PRECO)
    auto_unica("projeto", {"mix_estrategia_preco": est, "mix_posicao_preco": pos})

    st.markdown("#### Preços dos concorrentes")
    st.caption("Anote o preço que cada concorrente cobra pelo produto equivalente ao seu. "
               "Vale preço de vitrine, cardápio, site ou orçamento pedido.")
    prods = [x for x in cat["descricao"].astype(str) if x.strip()] if not cat.empty else []
    concs = [n for n in d["concorrentes"]["nome"].astype(str) if n.strip()]
    editor("precos_concorrentes", {
        "concorrente": S("Concorrente", options=concs, required=True) if concs else T("Concorrente", required=True),
        "produto": S("Produto equivalente ao seu", options=prods) if prods else T("Produto equivalente ao seu"),
        "preco": moeda("Preço cobrado"), "fonte": T("Onde viu o preço", width="medium")}, "concorrente")
    comp = comparar_concorrentes(cat, d["precos_concorrentes"])
    if not comp.empty:
        mostra = comp.copy()
        for col in ["Seu preço", "Média dos concorrentes", "Menor", "Maior"]:
            mostra[col] = mostra[col].map(brl)
        mostra["Índice"] = mostra["Índice"].map(lambda x: f"{x:.2f}".replace(".", ","))
        st.dataframe(mostra, hide_index=True, width="stretch")
        st.caption("Índice = seu preço ÷ média dos concorrentes. Acima de 1, você é mais caro.")

    st.markdown("#### Margem de cada item")
    mc = mc_produtos(res)
    if not mc.empty:
        t = mc.copy()
        t["Preço"], t["Custo direto"] = t["Preço"].map(brl), t["Custo direto"].map(brl)
        t["MC %"], t["Peso na receita"] = t["MC %"].map(pct), t["Peso na receita"].map(pct)
        st.dataframe(t, hide_index=True, width="stretch")
        st.caption("MC % = preço menos custo direto, impostos sobre a venda e outros custos variáveis, sobre o preço.")

    st.markdown("#### Desconto ou aumento: quanto o volume precisa mudar")
    mm = mc_media(res)
    if mm == mm and mm > 0:
        a, b = st.columns(2)
        dsc = a.slider("Desconto em estudo (%)", 0, 50, 10, key=wk("mx_desc")) / 100
        aum = b.slider("Aumento em estudo (%)", 0, 50, 10, key=wk("mx_aum")) / 100
        vol = volume_para_desconto(mm, dsc)
        queda = queda_tolerada_aumento(mm, aum)
        a.metric("Volume a mais para empatar", "impossível" if math.isinf(vol) else pct(vol))
        b.metric("Queda de volume tolerada", pct(queda))
        st.caption(f"Com margem de contribuição média de {pct(mm)}, um desconto de {dsc * 100:.0f}% só mantém "
                   f"o lucro se as vendas subirem {'mais do que é possível' if math.isinf(vol) else pct(vol)}. "
                   f"Um aumento de {aum * 100:.0f}% mantém o lucro mesmo que as vendas caiam até {pct(queda)}. "
                   "É uma conta de equilíbrio, não uma previsão de como os clientes vão reagir.")
    else:
        st.caption("Cadastre produtos com preço e custo para ver esta conta.")

    st.markdown("#### Condições de pagamento")
    st.caption("Como os clientes vão pagar. As taxas e prazos de exemplo são apenas referência: "
               "confira os da sua maquininha ou banco.")
    if d["meios_pagamento"].empty and st.button("Preencher com um exemplo", key=wk("mx_pag_ex")):
        recarregar_tabela("meios_pagamento", PAGAMENTO_EXEMPLO.copy(), "forma")
        st.rerun()
    editor("meios_pagamento", {"forma": T("Forma de pagamento", required=True),
                               "pct_vendas": N("% das vendas", min_value=0.0, max_value=100.0, format="%.1f"),
                               "taxa_pct": N("Taxa cobrada (%)", min_value=0.0, max_value=30.0, format="%.2f"),
                               "prazo_dias": N("Dias até o dinheiro entrar", min_value=0.0, format="%.0f")}, "forma")
    pg = d["meios_pagamento"]
    if not pg.empty:
        r = resumo_pagamento(pg)
        a, b, c = st.columns(3)
        a.metric("Soma das formas", f"{r['soma']:.0f}%".replace(".", ","))
        b.metric("Taxa média sobre a receita", f"{r['taxa']:.2f}%".replace(".", ","))
        c.metric("Prazo médio de recebimento", f"{r['prazo']:.0f} dias")
        if abs(r["soma"] - 100) > 0.5:
            st.warning("As formas de pagamento precisam somar 100% das vendas.", icon="⚠️")
        else:
            st.caption("Ao aplicar, a taxa média substitui a linha \"Taxas de cartões\" nos custos variáveis "
                       "e os prazos substituem os prazos de recebimento do capital de giro.")
            if st.button("✅ Aplicar ao plano financeiro", key=wk("mx_pag_ap")):
                cv, pz = aplicar_pagamento(pg, d["custos_variaveis"], d["capital_giro_prazos"])
                recarregar_tabela("custos_variaveis", cv, "descricao")
                recarregar_tabela("capital_giro_prazos", pz, "tipo")
                st.toast("Taxas e prazos aplicados ao plano.", icon="✅")
                st.rerun()

    v = {"mkt_preco": st.text_area("📝 Justificativa e detalhes do preço", proj["mkt_preco"], height=110, key=wk("m2"))}
    auto_unica("projeto", v)


def _aba_praca():
    dist = _select("Intensidade da distribuição", "mix_distribuicao", DISTRIBUICAO)
    if dist:
        st.caption(DISTRIBUICAO[dist])
    auto_unica("projeto", {"mix_distribuicao": dist})
    st.markdown("#### Canais de venda")
    st.caption("Onde o cliente compra, que parte das vendas passa por cada canal e quanto o canal cobra "
               "(comissão do aplicativo ou do marketplace, margem do revendedor). Confira as taxas atuais de cada canal.")
    editor("canais_venda", {"canal": S("Canal", options=opcoes("canal_venda", CANAIS_VENDA), required=True),
                            "pct_vendas": N("% das vendas", min_value=0.0, max_value=100.0, format="%.1f"),
                            "comissao_pct": N("Comissão ou taxa do canal (%)", min_value=0.0, max_value=60.0, format="%.1f"),
                            "observacao": T("Observação", width="medium")}, "canal")
    cv = d["canais_venda"]
    if not cv.empty:
        r = resumo_canais(cv)
        a, b, c = st.columns(3)
        a.metric("Soma dos canais", f"{num_br(r['soma'], 0)}%")
        b.metric("Comissão média sobre a receita", f"{num_br(r['comissao'], 2)}%")
        c.metric("Canal principal", r["maior"], f"{num_br(r['maior_pct'], 0)}% das vendas", delta_color="off")
        if abs(r["soma"] - 100) <= 0.5:
            st.caption(f"Ao aplicar, a comissão média entra nos custos variáveis como \"{_LINHA_COMISSAO}\" e a linha "
                       "\"Comissões de aplicativos de delivery\" é zerada, para a comissão não ser contada duas vezes.")
            if st.button("✅ Aplicar comissões ao plano financeiro", key=wk("mx_can_ap")):
                recarregar_tabela("custos_variaveis", aplicar_canais(cv, d["custos_variaveis"]), "descricao")
                st.toast("Comissões aplicadas.", icon="✅")
                st.rerun()
        else:
            st.warning("Os canais precisam somar 100% das vendas.", icon="⚠️")
    st.markdown("#### Localização")
    st.caption((f"Endereço informado: {proj['loc_endereco']}. " if str(proj.get("loc_endereco") or "").strip() else
                "A localização ainda não foi informada. ") + "Endereço, área e fluxo de pessoas ficam no plano operacional; "
               "o aluguel entra nos custos fixos.")
    if st.button("🏭 Abrir o plano operacional", key=wk("mx_op")):
        ir("operacional")
        st.rerun()
    auto_unica("projeto", {"mkt_praca": st.text_area("📝 Justificativa e detalhes da praça", proj["mkt_praca"], height=110,
                                                     key=wk("m3"))})


def _aba_promocao():
    res = R()
    obj = _select("Objetivo principal da comunicação nesta fase", "mix_objetivo_com", OBJETIVOS_COMUNICACAO)
    if obj:
        st.caption("Ações que costumam servir a esse objetivo: " + OBJETIVOS_COMUNICACAO[obj])
    st.markdown("#### O volume do plano e a promoção")
    st.caption("O volume de vendas do plano precisa de clientes. Aqui o programa estima quantos clientes ativos o "
               "primeiro mês exige e compara com os clientes novos que o orçamento de promoção promete trazer.")
    a, b = st.columns(2)
    tk = a.number_input("Gasto médio por compra (R$); deixe 0 para usar o preço médio do catálogo", min_value=0.0,
                        value=float(proj.get("mix_ticket") or 0), step=1.0, key=wk("mx_tk"))
    fr = b.number_input("Quantas vezes por mês o cliente típico compra", min_value=0.0,
                        value=float(proj.get("mix_freq") or 0), step=0.5, key=wk("mx_fr"))
    auto_unica("projeto", {"mix_objetivo_com": obj, "mix_ticket": tk, "mix_freq": fr})
    pv = promocao_vs_volume(proj, d, res)
    k = st.columns(4)
    k[0].metric("Clientes ativos que o plano exige", num_br(pv["cn"]["clientes"], 0) if pv["cn"].get("ok") else "n.a.")
    k[1].metric("Clientes novos por mês (promoção)", num_br(pv["novos"], 0) if pv["novos"] else "n.a.")
    k[2].metric("Custo por cliente novo (CAC)", brl(pv["cac"]) if pv["cac"] == pv["cac"] else "n.a.")
    k[3].metric("Verba de promoção ÷ receita", pct(pv["verba_receita"]) if pv["verba_receita"] == pv["verba_receita"] else "n.a.")
    if pv["novos"] <= 0:
        st.info("Os clientes novos por mês vêm do orçamento de marketing por canal.", icon="💡")
    if st.button("💸 Abrir orçamento de marketing e calendário", key=wk("mx_orc")):
        ir("orcamento_mkt")
        st.rerun()
    auto_unica("projeto", {"mkt_promocao": st.text_area("📝 Justificativa e detalhes da promoção", proj["mkt_promocao"],
                                                        height=110, key=wk("m4"))})


def _aba_servico():
    res = R()
    st.caption("Em serviços o cliente compra algo que não pode ver antes. Booms e Bitner (1981) acrescentam três Ps "
               "ao mix: as pessoas que atendem, o processo pelo qual o serviço acontece e a evidência física que o torna visível.")
    if not e_servico(proj, res):
        st.info("O projeto não está marcado como serviço, mas estes três Ps também ajudam no varejo e na alimentação, "
                "onde o atendimento faz parte do que se vende.", icon="💡")
    st.markdown("#### 👥 Pessoas")
    pes = d["pessoal"]
    n = float(pes["num_empregados"].sum()) if not pes.empty else 0.0
    st.caption(f"Equipe no plano: {num_br(n, 0)} pessoa(s), cadastradas em Custos e despesas.")
    v = {"mix_pessoas": st.text_area("Quem atende o cliente, que preparo precisa ter e como o atendimento será padronizado",
                                     proj.get("mix_pessoas") or "", height=100, key=wk("mx_pe"))}
    st.markdown("#### 🔁 Processos: mapa do serviço")
    st.caption("Liste as etapas do serviço na ordem em que acontecem, separando o que o cliente vê do que acontece nos "
               "bastidores (Shostack, 1984). Marque onde algo costuma dar errado.")
    editor("blueprint", {"etapa": T("Etapa", required=True, width="medium"),
                         "linha": S("Onde acontece", options=LINHAS_BLUEPRINT),
                         "tempo_min": N("Tempo (minutos)", min_value=0.0, step=1.0),
                         "ponto_falha": T("O que pode dar errado", width="medium")}, "etapa")
    bp = resumo_blueprint(d["blueprint"])
    if bp["etapas"]:
        a, b, c = st.columns(3)
        a.metric("Etapas", num_br(bp["etapas"], 0))
        b.metric("Tempo total", f"{num_br(bp['tempo'], 0)} min")
        c.metric("Pontos de falha", num_br(len(bp["falhas"]), 0))
        if bp["falhas"]:
            st.caption("Etapas com risco de falha: " + ", ".join(bp["falhas"]) +
                       ". Cada uma merece um padrão de atendimento ou uma checagem.")
    v["mix_processos"] = st.text_area("Observações sobre o processo", proj.get("mix_processos") or "", height=80,
                                      key=wk("mx_pr"))
    st.markdown("#### 🏠 Evidência física")
    v["mix_evidencia"] = st.text_area("O que o cliente vê e toca: ambiente, uniforme, embalagem, site, contrato, "
                                      "certificados, avaliações", proj.get("mix_evidencia") or "", height=100, key=wk("mx_ev"))
    auto_unica("projeto", v)


def _aba_diagnostico():
    vw = None
    try:
        vw = van_westendorp(d["precos_vw"])
    except Exception:
        pass
    res = R()
    st.markdown("#### Coerência entre os Ps")
    st.caption("Os Ps precisam contar a mesma história. As combinações abaixo aparecem conforme as escolhas feitas nas outras abas.")
    co = coerencia_mix(proj, d, res)
    if co.empty:
        st.info("Escolha a posição e a estratégia de preço, a distribuição e o objetivo da comunicação para ver a coerência do mix.",
                icon="💡")
    else:
        st.dataframe(co, hide_index=True, width="stretch")
        ok = int(co["Situação"].str.startswith("✅").sum())
        st.caption(f"{ok} de {len(co)} combinações coerentes.")
    st.markdown("#### Produto e preço")
    mostrar_alertas(checar_produto_preco(proj, d, res, vw))
    st.markdown("#### Praça e promoção")
    mostrar_alertas(checar_praca_promocao(proj, d, res))
    st.markdown("#### Autoavaliação pelos olhos do cliente (4Cs)")
    st.caption("Lauterborn (1990) propõe olhar cada P do ponto de vista de quem compra. Dê uma nota de 1 a 5 e escreva a evidência.")
    v, notas = {}, []
    for campo_n, campo_t, titulo, pergunta in QUATRO_CS:
        c1, c2 = st.columns([1, 3])
        v[campo_n] = c1.slider(titulo, 0, 5, int(proj.get(campo_n) or 0), key=wk(campo_n), help="0 = ainda não avaliado")
        v[campo_t] = c2.text_area(pergunta, proj.get(campo_t) or "", height=80, key=wk(campo_t))
        notas.append((titulo.split(" (")[0], v[campo_n]))
    auto_unica("projeto", v)
    avaliadas = [(t, n) for t, n in notas if n > 0]
    if len(avaliadas) == 4:
        pior = min(avaliadas, key=lambda x: x[1])
        st.altair_chart(alt.Chart(pd.DataFrame(avaliadas, columns=["C", "Nota"])).mark_bar(color="#2f6fd0").encode(
            x=alt.X("Nota:Q", scale=alt.Scale(domain=[0, 5])), y=alt.Y("C:N", sort=None, title=None)), width="stretch")
        st.caption(f"O ponto mais fraco pelo olhar do cliente é **{pior[0].lower()}** (nota {pior[1]}). "
                   "É por ele que vale começar a ajustar o mix.")


def pg_marketing():
    cabecalho("marketing")
    st.caption("Cada P tem escolhas que alimentam o restante do plano. Os campos de texto continuam "
               "valendo como justificativa.")
    abas = st.tabs(["📦 Produto", "🏷️ Preço", "📍 Praça", "📣 Promoção", "👥 Pessoas, processos e evidência física",
                    "🔎 Diagnóstico do mix"])
    with abas[0]:
        _aba_produto()
    with abas[1]:
        _aba_preco()
    with abas[2]:
        _aba_praca()
    with abas[3]:
        _aba_promocao()
    with abas[4]:
        _aba_servico()
    with abas[5]:
        _aba_diagnostico()
    proxima("marketing")


# ------------------------------------------------------------------------------
# PLANO DE MARKETING (v8.4)
# ------------------------------------------------------------------------------
SELO_ETAPA = {2: "✅", 1: "◐", 0: "○"}


def _botao_ir(rotulo, pagina, chave_):
    if st.button(rotulo, key=wk(chave_)):
        ir(pagina)
        st.rerun()


def pg_plano_mkt():
    cabecalho("plano_mkt")
    res = R()
    st.caption("O plano segue a sequência proposta na literatura: análise da situação, objetivos, estratégia, táticas, "
               "ações e controle (Smith, 2011; McDonald, 2007; Las Casas, 2011). Quase tudo vem das outras ferramentas; "
               "aqui o grupo amarra as partes e completa o que falta.")
    et = etapas_plano_mkt(proj, d)
    cols = st.columns(6)
    for c, (nome, s, falta) in zip(cols, et):
        c.metric(nome, SELO_ETAPA[s], help=f"Completa quando houver {falta}.")
    cont = conteudo_plano_mkt(proj, d, res)
    abas = st.tabs([f"{i + 1}. {n}" for i, (n, _) in enumerate(ETAPAS_PLANO)])
    with abas[0]:
        st.caption(ETAPAS_PLANO[0][1])
        st.markdown(cont[0]["texto"] if not _tem_txt(proj.get("pm_situacao")) else
                    cont[0]["texto"].replace(str(proj["pm_situacao"]), "").strip() or "")
        c = st.columns(4)
        for col, (rot, pg) in zip(c, [("🎯 Mercado", "mercado"), ("🥊 Concorrência", "concorrencia"), ("🧩 SWOT", "swot"),
                                      ("🔀 SWOT cruzada", "swot_cruzada")]):
            with col:
                _botao_ir(rot, pg, "pm_ir_" + pg)
        auto_unica("projeto", {"pm_situacao": st.text_area(
            "Síntese do diagnóstico: as três conclusões mais importantes para o marketing", proj.get("pm_situacao") or "",
            height=120, key=wk("pm_sit"))})
    with abas[1]:
        st.caption(ETAPAS_PLANO[1][1] + " Entram aqui os objetivos da perspectiva Clientes cadastrados em Missão, visão, valores e objetivos.")
        if cont[1]["tabela"] is not None:
            st.dataframe(cont[1]["tabela"], hide_index=True, width="stretch")
        else:
            st.info("Nenhum objetivo da perspectiva Clientes ainda. Exemplos: conquistar 300 clientes ativos até o mês 6; "
                    "chegar a 40% de recompra em 30 dias até o mês 9.", icon="💡")
        _botao_ir("🧭 Abrir os objetivos SMART", "identidade", "pm_ir_obj")
    with abas[2]:
        st.caption(ETAPAS_PLANO[2][1])
        fp = frase_posicionamento(proj)
        st.markdown("**Declaração de posicionamento:** " + (fp or "_ainda não preenchida_"))
        _botao_ir("📌 Abrir o posicionamento", "posicionamento", "pm_ir_pos")
        a, b = st.columns(2)
        with a:
            ans = _select("Direção de crescimento (matriz de Ansoff)", "pm_ansoff", ANSOFF,
                          "Para um negócio novo, pense no principal movimento dos próximos anos depois da abertura.")
            if ans:
                st.caption(f"**{ANSOFF[ans][0]}.** {ANSOFF[ans][1]}")
        with b:
            gen = _select("Estratégia competitiva (Porter)", "pm_generica", GENERICAS)
            if gen:
                st.caption(GENERICAS[gen])
        auto_unica("projeto", {"pm_ansoff": ans, "pm_generica": gen})
        if gen.startswith("Liderança") and proj.get("mix_posicao_preco") == "Premium":
            st.warning("Liderança em custo com preço premium: as duas escolhas puxam para lados opostos.", icon="⚠️")
        if gen and "diferenciação" in gen.lower() and proj.get("mix_estrategia_preco") == "Penetração":
            st.warning("Diferenciação com preço de penetração: se o diferencial existe, ele pode ser cobrado.", icon="⚠️")
    with abas[3]:
        st.caption(ETAPAS_PLANO[3][1])
        if cont[3]["tabela"] is not None:
            st.dataframe(cont[3]["tabela"], hide_index=True, width="stretch")
        co = coerencia_mix(proj, d, res)
        if not co.empty:
            ok = int(co["Situação"].str.startswith("✅").sum())
            (st.success if ok == len(co) else st.warning)(f"Coerência do mix: {ok} de {len(co)} combinações coerentes.")
        _botao_ir("📣 Abrir o mix de marketing", "marketing", "pm_ir_mix")
    with abas[4]:
        st.caption(ETAPAS_PLANO[4][1])
        if cont[4]["tabela"] is not None:
            st.markdown(cont[4]["texto"])
            st.dataframe(cont[4]["tabela"], hide_index=True, width="stretch")
        else:
            st.info("As ações ficam no calendário do orçamento de marketing.", icon="💡")
        _botao_ir("🗓️ Abrir o calendário de marketing", "orcamento_mkt", "pm_ir_cal")
    with abas[5]:
        st.caption(ETAPAS_PLANO[5][1] + " McDonald (2007) lembra que o plano só fecha o ciclo quando os resultados "
                   "voltam para a análise da situação seguinte.")
        if d["kpis_mkt"].empty and st.button("✨ Sugerir indicadores", key=wk("pm_kpi_sug")):
            recarregar_tabela("kpis_mkt", pd.DataFrame([{"indicador": i, "meta": "", "frequencia": f, "fonte": fo, "responsavel": ""}
                                                        for i, f, fo in KPIS_SUGERIDOS]), "indicador")
            st.rerun()
        editor("kpis_mkt", {"indicador": T("Indicador", required=True, width="medium"), "meta": T("Meta"),
                            "frequencia": S("Frequência", options=FREQUENCIAS), "fonte": T("De onde vem o dado", width="medium"),
                            "responsavel": T("Responsável")}, "indicador")
        sem_meta = [i for i, m in zip(d["kpis_mkt"]["indicador"], d["kpis_mkt"]["meta"]) if not str(m).strip()]
        if sem_meta:
            st.warning("Indicadores sem meta não servem para controle: " + ", ".join(sem_meta) + ".", icon="⚠️")
        _botao_ir("📊 Abrir o acompanhamento do realizado", "realizado", "pm_ir_real")
    with st.expander("📄 Ver o plano de marketing montado"):
        for i, sec in enumerate(cont, 1):
            st.markdown(f"**{i}. {sec['titulo']}**")
            st.markdown(sec["texto"])
            if sec["tabela"] is not None:
                st.dataframe(sec["tabela"], hide_index=True, width="stretch")
    proxima("plano_mkt")


# ------------------------------------------------------------------------------
# SAZONALIDADE (v8.6)
# ------------------------------------------------------------------------------
def pg_sazonalidade():
    cabecalho("sazonalidade")
    st.caption("Poucos negócios vendem igual o ano todo. Aqui o grupo informa o peso de cada mês (100 = mês médio). "
               "O total do ano continua o mesmo; muda a distribuição dos meses e, com ela, o fluxo de caixa.")
    a, b = st.columns(2)
    ini = a.selectbox("Mês do calendário em que o negócio abre", list(range(1, 13)),
                      index=int(proj.get("saz_mes_inicio") or 1) - 1, format_func=lambda m: MESES_NOME[m - 1], key=wk("saz_ini"))
    perfil = b.selectbox("Perfil de partida (ilustrativo)", list(PERFIS_SAZONAIS), key=wk("saz_perfil"),
                         index=list(PERFIS_SAZONAIS).index(proj["saz_perfil"]) if proj.get("saz_perfil") in PERFIS_SAZONAIS else 0)
    auto_unica("projeto", {"saz_mes_inicio": float(ini)})
    if st.button("Usar este perfil", key=wk("saz_usar")):
        auto_unica("projeto", {"saz_perfil": perfil})
        recarregar_tabela("sazonalidade", tabela_perfil(perfil), "mes")
        st.rerun()
    st.caption("Os perfis são ilustrativos. Ajuste os índices com o que o grupo sabe do negócio: vendas de concorrentes, "
               "conversa com fornecedores, datas comemorativas e o clima da cidade.")
    editor("sazonalidade", {"mes": N("Mês do calendário", min_value=1, max_value=12, step=1, required=True),
                            "indice": N("Índice (100 = mês médio)", min_value=0.0, step=5.0)}, "mes")
    res = R()
    mm = res["meses"]
    graf = pd.DataFrame({"Mês do plano": [f"{i + 1:02d}" for i in range(12)], "Receita": mm["Receita"],
                         "Resultado operacional": mm["Resultado operacional"]})
    st.altair_chart(alt.Chart(graf.melt("Mês do plano", var_name="Série", value_name="R$")).mark_line(point=True).encode(
        x="Mês do plano:N", y="R$:Q", color="Série:N"), width="stretch")
    mostrar_leitura("sazonalidade_resultado", leitura_sazonalidade(proj, res))
    if st.button("💵 Ver o fluxo de caixa", key=wk("saz_fluxo")):
        ir("fluxo")
        st.rerun()
    proxima("sazonalidade")

# ##############################################################################
# PARTE: pnbox_ui_paginas3.py
# ##############################################################################
# ==============================================================================
# VIABILIZALAB - INTERFACE: CONCLUSÃO, ENTREGA, PROFESSOR E CONFIGURAÇÕES
# ==============================================================================


# ------------------------------------------------------------------------------
# PONTOS DE RESTAURAÇÃO (versões salvas dentro do próprio projeto)
# ------------------------------------------------------------------------------
def foto_projeto() -> str:
    dados = {}
    for t in SCHEMA:
        if t == "versoes":
            continue
        dados[t] = ler(conn, t).to_dict("records")
    return json.dumps(dados, ensure_ascii=False, default=str)


def salvar_versao(nome: str, automatica=False):
    conn.execute("INSERT INTO versoes (nome, data, dados) VALUES (?,?,?)",
                 (("[auto] " if automatica else "") + nome, datetime.now().strftime("%d/%m/%Y %H:%M"), foto_projeto()))
    if automatica:   # guarda só as 5 cópias automáticas mais recentes
        ids = [r[0] for r in conn.execute("SELECT id FROM versoes WHERE nome LIKE '[auto]%' ORDER BY id DESC").fetchall()]
        for i in ids[5:]:
            conn.execute("DELETE FROM versoes WHERE id=?", (i,))
    conn.commit()


def restaurar_versao(vid: int):
    linha = conn.execute("SELECT dados FROM versoes WHERE id=?", (vid,)).fetchone()
    if not linha:
        return
    salvar_versao("Antes de restaurar", automatica=True)
    dados = json.loads(linha[0])
    for t, regs in dados.items():
        if t not in SCHEMA:
            continue
        cols = colunas_de(t)
        if t in ("projeto", "proposta_valor"):
            if regs:
                salvar_unica(conn, t, {c: regs[0][c] for c in cols if c in regs[0]})
            continue
        conn.execute(f"DELETE FROM {t}")
        for r in regs:
            cc = [c for c in cols if c in r]
            if cc:
                conn.execute(f"INSERT INTO {t} ({', '.join(cc)}) VALUES ({', '.join('?' * len(cc))})", [r[c] for c in cc])
    conn.commit()
    for k in ("buf", "snap", "hash"):
        ss[k] = {}
    ss.gen += 1
    invalidar()


def copia_automatica():
    """Cria um ponto de restauração a cada 10 minutos de uso."""
    agora = time.time()
    ultimo = ss.get("ultima_copia_auto", {}).get(ss.db)
    if ultimo is None:
        ss.setdefault("ultima_copia_auto", {})[ss.db] = agora
    elif agora - ultimo >= 600:
        salvar_versao("Cópia automática (10 min)", automatica=True)
        ss.ultima_copia_auto[ss.db] = agora


# ------------------------------------------------------------------------------
# ANÁLISE COM IA
# ------------------------------------------------------------------------------
def pg_ia():
    cabecalho("ia")
    res = R()
    st.markdown("""
<div class="pn-card-info">
<b>Como funciona:</b> o programa monta um texto com todos os números do seu plano e as instruções para a IA.
Você copia e cola em qualquer assistente gratuito (ChatGPT, Gemini, Copilot, Claude, DeepSeek...).
Nenhum dado sai do seu computador sem que você mesmo cole. <b>A IA comenta; quem decide é o grupo.</b>
Confira os números que ela citar com os do programa.
</div>""", unsafe_allow_html=True)
    st.write("")
    so_viab = st.toggle("🎯 Só o teste de viabilidade real (veredito, cenário realista e ajustes para viabilizar)", key=wk("ia_so_viab"),
                        help="Monta um texto focado: a IA compara as premissas com o mercado, refaz as contas num cenário realista, "
                             "dá o veredito (viável, viável com ajustes ou inviável) e diz o que mudar.")
    c1, c2 = st.columns([2, 1])
    with c1:
        escolha = ["viabilidade_real", "validacao"] if so_viab else st.multiselect("O que você quer que a IA faça?", list(TAREFAS_IA), format_func=lambda k: TAREFAS_IA[k],
                                 default=[k for k in TAREFAS_IA if k != "recorrencia" or R()["recorrencia"].get("tem")],
                                 key=wk("ia_tarefas"))
    with c2:
        publico = st.selectbox("Quem está pedindo", ["estudante de Administração", "estudante de Ciências Contábeis",
                                                      "estudante de Economia", "empreendedor iniciante", "professor"],
                               key=wk("ia_publico"))
    piv, _ = pivo_concorrencia()
    texto = gerar_prompt_ia(proj, d, res, escolha, publico, piv)
    st.caption(f"{len(texto):,} caracteres · cerca de {len(texto) // 4:,} palavras-token".replace(",", "."))
    t1, t2 = st.tabs(["📋 Copiar (botão no canto da caixa)", "✏️ Ver e editar antes de copiar"])
    with t1:
        st.code(texto, language="markdown", wrap_lines=True, height=420)
    with t2:
        editado = st.text_area("Texto", texto, height=420, label_visibility="collapsed", key=wk("ia_edit"))
        texto = editado
    st.download_button("⬇️ Baixar como arquivo .txt", texto.encode("utf-8"), f"prompt_IA_{PID}.txt", "text/plain",
                       width="stretch", type="primary")
    with st.expander("💡 Dicas para tirar mais proveito da IA"):
        st.markdown("""
- Se a IA disser que o texto é longo demais, cole em duas partes e escreva "aguarde a segunda parte".
- Peça depois: *"Agora explique como se eu nunca tivesse estudado finanças"* ou *"Resuma em 5 tópicos para o slide"*.
- Assistentes que executam código (ChatGPT com análise de dados, Gemini, Claude) conseguem desenhar os gráficos.
- Nunca copie a resposta da IA direto para o trabalho: use como ponto de partida e cite que a usou.
""")
    proxima("ia")


# ------------------------------------------------------------------------------
# SUMÁRIO EXECUTIVO E CONCLUSÃO
# ------------------------------------------------------------------------------
def pg_sumario():
    cabecalho("sumario")
    res = R()
    st.info("O sumário executivo é a primeira coisa que o leitor vê, mas a última que se escreve. "
            "Use o rascunho automático como ponto de partida e reescreva com as palavras do grupo.")
    if st.button("✨ Gerar rascunho automático a partir dos dados", type="primary"):
        novo = gerar_sumario(proj, d, res)
        proj["sumario_executivo"] = novo
        salvar_unica(conn, "projeto", {"sumario_executivo": novo})
        # chave nova = caixa de texto nova (o navegador não reaproveita o valor antigo do componente)
        ss["sumario_versao"] = ss.get("sumario_versao", 0) + 1
        marcar_salvo()
        st.rerun()
    txt = st.text_area("Sumário executivo", proj["sumario_executivo"], height=360,
                       key=wk(f"sumario_txt_{ss.get('sumario_versao', 0)}"),
                       placeholder="Clique no botão acima para gerar um rascunho, ou escreva aqui.")
    st.caption(f"{len(txt.split())} palavras. Recomenda-se entre 250 e 500.")
    st.divider()
    st.subheader("Conclusão e parecer do grupo")
    st.caption("O negócio é viável? Em que condições? O que precisa ser validado antes de investir? "
               "Use os indicadores e o semáforo para justificar.")
    conc = st.text_area("Conclusão", proj["conclusao"], height=240, key=wk("conclusao_txt"), label_visibility="collapsed")
    auto_unica("projeto", {"sumario_executivo": txt, "conclusao": conc})
    salvo_txt()
    proxima("sumario")


# ------------------------------------------------------------------------------
# MODO APRESENTAÇÃO
# ------------------------------------------------------------------------------
def pg_apresentacao():
    cabecalho("apresentacao")
    res = R()
    a1 = res["anos"][1]
    sem = semaforo(res)
    mk = mercado(proj, res)
    slides = ["capa", "mercado", "proposta", "numeros", "semaforo", "investimento", "dre", "caixa", "fim"]
    ss.setdefault("slide", 0)
    ss.slide = min(ss.slide, len(slides) - 1)
    n1, n2, n3 = st.columns([1, 3, 1])
    if n1.button("◀ Anterior", width="stretch", disabled=ss.slide == 0):
        ss.slide -= 1
        st.rerun()
    n2.progress((ss.slide + 1) / len(slides), text=f"Slide {ss.slide + 1} de {len(slides)}")
    if n3.button("Próximo ▶", width="stretch", disabled=ss.slide == len(slides) - 1, type="primary"):
        ss.slide += 1
        st.rerun()
    s = slides[ss.slide]
    st.markdown('<div class="pn-slide">', unsafe_allow_html=True)
    if s == "capa":
        st.markdown(f"<h1 style='text-align:center;font-size:3.4rem;margin-top:2rem'>{proj['nome_negocio']}</h1>", unsafe_allow_html=True)
        props = d["propostas"]
        if not props.empty:
            st.markdown(f"<h3 style='text-align:center;font-weight:400'>{props['oferecem'].iloc[0]}</h3>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center'>{proj['cidade_uf']} · {str(proj['autores']).replace(chr(10), ', ')}</p>",
                    unsafe_allow_html=True)
    elif s == "mercado":
        st.header("🌍 O mercado")
        if mk["preenchido"]:
            c = st.columns(3)
            c[0].metric("TAM (mercado total)", brl(mk["tam"]), f"{num_br(mk['tam_c'], 0)} clientes", delta_color="off")
            c[1].metric("SAM (alcançável)", brl(mk["sam"]), f"{num_br(mk['sam_c'], 0)} clientes", delta_color="off")
            c[2].metric("SOM (nossa meta)", brl(mk["som"]), f"{num_br(mk['som_c'], 0)} clientes", delta_color="off")
            st.markdown(f"### A receita do Ano 1 representa **{pct(mk['pct_som'])}** do SOM.")
        else:
            st.info("Preencha a ferramenta Tamanho do mercado.")
    elif s == "proposta":
        st.header("💎 Proposta de valor")
        for _, p_ in d["propostas"].iterrows():
            st.markdown(f"#### Para **{p_['persona_nome']}**, que deseja {p_['que_desejam']}")
            st.markdown(f"Oferecemos **{p_['oferecem']}**, por meio de {p_['por_meio']}, diferente de {p_['diferentemente']}.")
    elif s == "numeros":
        st.header("💰 Os números do primeiro ano")
        c = st.columns(4)
        c[0].metric("Investimento total", brl(res["investimento_total"]))
        c[1].metric("Receita bruta", brl(a1["rec_bruta"]))
        c[2].metric("Lucro líquido", brl(a1["resultado"]))
        c[3].metric("Lucratividade", pct(div(a1["resultado"], a1["rec_bruta"])))
        c = st.columns(4)
        c[0].metric(f"VPL ({res['N']} anos)", brl(res["vpl"]))
        c[1].metric("TIR", pct(res["tir"]) if res["tir"] is not None else "n.a.")
        c[2].metric("Payback descontado", f"{num_br(res['payback_desc'], 1)} meses" if res["payback_desc"] == res["payback_desc"] else f"> {res['MESES']} meses")
        c[3].metric("Ponto de equilíbrio", brl(a1["pe_contabil"]))
    elif s == "semaforo":
        st.header("🚦 Saúde do negócio")
        for i in range(0, len(sem), 3):
            cols = st.columns(3)
            for c, it in zip(cols, sem[i:i + 3]):
                with c.container(border=True):
                    st.markdown(f"<div style='font-size:2.2rem'>{it['Status']}</div><b>{it['Indicador']}</b><br>"
                                f"<span style='font-size:1.6rem'>{it['Valor']}</span><br><small>{it['Leitura']}</small>",
                                unsafe_allow_html=True)
    elif s == "investimento":
        st.header("🏗️ Investimento e fontes")
        comp = composicoes(res, d)
        c1, c2 = st.columns(2)
        with c1:
            rosca(comp["investimento"], "Composição do investimento", "ap_inv", tabela=False)
        with c2:
            fontes = pd.DataFrame({"Item": ["Capital próprio", "Capital de terceiros"],
                                   "Valor": [res["capital_proprio"], res["capital_terceiros"]]})
            fontes["%"] = fontes["Valor"] / max(fontes["Valor"].sum(), 1)
            rosca(fontes, "Fontes de recursos", "ap_fontes", tabela=False)
    elif s == "dre":
        st.header(f"📈 Receita e resultado em {res['N']} ano{'s' if res['N'] > 1 else ''}")
        anos = pd.DataFrame([{"Ano": f"Ano {k}", "Receita": v["rec_bruta"], "Resultado": v["resultado"]} for k, v in res["anos"].items()])
        longo = anos.melt("Ano", var_name="Série", value_name="R$")
        st.altair_chart(alt.Chart(longo).mark_bar().encode(x=alt.X("Ano:N", title=None), xOffset="Série:N", y="R$:Q",
                                                           color=alt.Color("Série:N", scale=alt.Scale(range=["#2f6fd0", "#18a058"])))
                        .properties(height=420), width="stretch")
    elif s == "caixa":
        st.header("🏦 Saldo de caixa mês a mês (Ano 1)")
        fm = res["fluxo_mensal"]
        if not fm.empty:
            st.altair_chart(alt.Chart(fm.assign(n=range(1, len(fm) + 1))).mark_area(
                line={"color": "#2f6fd0"}, color="#2f6fd033").encode(x=alt.X("n:Q", title="Mês"), y=alt.Y("Saldo final:Q", title="R$"))
                .properties(height=420), width="stretch")
    else:
        st.markdown("<h1 style='text-align:center;font-size:3.4rem;margin-top:3rem'>Obrigado! 🙌</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center;font-weight:400'>Perguntas?</h3>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.caption("Dica: aperte F11 para tela cheia e recolha o menu lateral (>> no topo) durante a apresentação.")


# ------------------------------------------------------------------------------
# RELATÓRIO E EXPORTAÇÃO
# ------------------------------------------------------------------------------
def planilha_excel(res) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([proj]).T.reset_index().set_axis(["Campo", "Valor"], axis=1).to_excel(w, sheet_name="Projeto", index=False)
        for t, df_t in d.items():
            if isinstance(df_t, pd.DataFrame) and t != "versoes":
                df_t.to_excel(w, sheet_name=t[:31], index=False)
        res["cat"].to_excel(w, sheet_name="Custos_e_precos", index=False)
        res["inv"].to_excel(w, sheet_name="Depreciacao", index=False)
        res["meses"].to_excel(w, sheet_name="Projecao_mensal", index=False)
        res["fluxo_mensal"].to_excel(w, sheet_name="Fluxo_caixa_mensal", index=False)
        tabela_dre(res).to_excel(w, sheet_name="DRE", index=False)
        tabela_indicadores(res).to_excel(w, sheet_name="Indicadores", index=False)
        pd.DataFrame(semaforo(res)).to_excel(w, sheet_name="Semaforo", index=False)
        cenarios(res, proj).to_excel(w, sheet_name="Cenarios", index=False)
        CHK().drop(columns=["destino"], errors="ignore").to_excel(w, sheet_name="Checagens", index=False)
    return buf.getvalue()


def pg_relatorio():
    cabecalho("relatorio")
    res = R()
    chk = CHK()
    pend = chk[chk["Situação"] != "OK"]
    comp = completude()
    c1, c2, c3 = st.columns(3)
    c1.metric("Plano preenchido", pct(comp, 0))
    c2.metric("Checagens com problema", len(pend))
    c3.metric("Páginas estimadas do PDF", "25 a 35")
    if not pend.empty:
        with st.expander(f"⚠️ {len(pend)} ponto(s) para revisar antes de entregar", expanded=True):
            for i, r in pend.iterrows():
                a, b = st.columns([4, 1])
                a.markdown(f"**{r['Teste']}**: {r['Valor apurado']}. {r['Como corrigir']}")
                if b.button("Ir corrigir →", key=f"rel_chk_{i}", width="stretch"):
                    ir(r["destino"], f"{r['Teste']}: {r['Valor apurado']}. {r['Como corrigir']}")
                    st.rerun()

    st.subheader("📦 Baixar")
    piv, _ = pivo_concorrencia()
    piv_i = piv
    k = st.columns(3)
    with k[0].container(border=True):
        st.markdown("#### 📄 Relatório completo (PDF)")
        st.caption("Todas as ferramentas, tabelas, gráficos e checagens, na ordem do plano de negócio do Sebrae.")
        if st.button("Gerar PDF", type="primary", width="stretch", key="b_pdf"):
            with st.spinner("Montando o relatório... (pode levar alguns segundos)"):
                extras = {"comp": composicoes(res, d), "dre": tabela_dre(res), "indicadores": tabela_indicadores(res),
                          "cenarios": cenarios(res, proj), "checagens": tabela_checagens(res, d),
                          "swot_diag": pontuacao_swot(d["swot"]), "funil": tabela_funil(d["funil_vendas"]),
                          "concorrencia_pivot": piv_i}
                ss[f"pdf_{PID}"] = gerar_pdf(proj, d, res, extras)
        if ss.get(f"pdf_{PID}"):
            st.download_button("⬇️ Baixar PDF", ss[f"pdf_{PID}"], f"Plano_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🎤 Pitch deck (PDF 16:9)")
        st.caption("12 slides prontos para a apresentação oral.")
        if st.button("Gerar pitch", width="stretch", key="b_pitch"):
            with st.spinner("Montando os slides..."):
                ss[f"pitch_{PID}"] = gerar_pitch(proj, d, res)
        if ss.get(f"pitch_{PID}"):
            st.download_button("⬇️ Baixar pitch", ss[f"pitch_{PID}"], f"Pitch_{Path(ss.db).stem}.pdf", "application/pdf", width="stretch")
    with k[2].container(border=True):
        st.markdown("#### 📝 Documento editável (Word)")
        st.caption("Para o grupo revisar o texto e formatar nas normas da ABNT.")
        if st.button("Gerar Word", width="stretch", key="b_docx"):
            with st.spinner("Montando o documento..."):
                ss[f"docx_{PID}"] = gerar_docx(proj, d, res, piv_i) or b""
        if ss.get(f"docx_{PID}") is not None and f"docx_{PID}" in ss:
            if ss[f"docx_{PID}"]:
                st.download_button("⬇️ Baixar .docx", ss[f"docx_{PID}"], f"Plano_{Path(ss.db).stem}.docx",
                                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document", width="stretch")
            else:
                st.warning("A biblioteca python-docx não está instalada neste computador.")
    with st.container(border=True):
        st.markdown("#### 🎓 Relatório acadêmico no padrão ABNT (PDF)")
        st.caption("Capa, folha de rosto, sumário com páginas, seções numeradas, tabelas com título e fonte, citações e referências. "
                   "Margens de 3 e 2 cm, fonte 12 e espaçamento 1,5.")
        c = st.columns(3)
        v = {"abnt_instituicao": c[0].text_input("Instituição", proj["abnt_instituicao"], key=wk("abnt_inst")),
             "abnt_curso": c[1].text_input("Curso", proj["abnt_curso"], key=wk("abnt_curso")),
             "abnt_disciplina": c[2].text_input("Disciplina", proj["abnt_disciplina"], key=wk("abnt_disc"))}
        c = st.columns(3)
        v["abnt_orientador"] = c[0].text_input("Professor(a) / orientador(a)", proj["abnt_orientador"], key=wk("abnt_ori"))
        v["abnt_cidade"] = c[1].text_input("Cidade", proj["abnt_cidade"] or proj["cidade_uf"], key=wk("abnt_cid"))
        v["abnt_ano"] = c[2].text_input("Ano", proj["abnt_ano"] or str(datetime.now().year), key=wk("abnt_ano"))
        auto_unica("projeto", v)
        st.caption("Os autores vêm do campo Autores (Informações e sócios), separados por ponto e vírgula.")
        with st.expander("📚 Referências consultadas pelo grupo (NBR 6023)"):
            st.caption("Uma referência por linha, já no formato ABNT. Use **negrito** com dois asteriscos, como no exemplo. As bases de dados "
                       "usadas pelo programa (IBGE, Sebrae) entram sozinhas.")
            editor("referencias", {"referencia": T("Referência", required=True, width="large")}, "referencia")
            st.code("SOBRENOME, Nome. **Título do livro**: subtítulo. 2. ed. Cidade: Editora, 2020.", language=None)
        if st.button("Gerar relatório ABNT", type="primary", key="b_abnt"):
            with st.spinner("Montando o relatório acadêmico..."):
                ss[f"abnt_{PID}"] = gerar_relatorio_abnt(proj, d, res)
        if ss.get(f"abnt_{PID}"):
            st.download_button("⬇️ Baixar relatório ABNT", ss[f"abnt_{PID}"], f"Relatorio_ABNT_{Path(ss.db).stem}.pdf",
                               "application/pdf")
    k = st.columns(2)
    with k[0].container(border=True):
        st.markdown("#### 📊 Planilha (Excel)")
        st.caption("Todos os dados e cálculos, uma aba por tabela.")
        st.download_button("⬇️ Baixar planilha", planilha_excel(res), f"Planilha_{Path(ss.db).stem}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with k[1].container(border=True):
        st.markdown("#### 🗄️ Arquivo do projeto (.pnbox)")
        st.caption("Guarde no pen drive ou na nuvem. É o que o professor recebe para avaliar.")
        conn.commit()
        st.download_button("⬇️ Baixar .pnbox", Path(ss.db).read_bytes(), f"{Path(ss.db).stem}.pnbox",
                           "application/octet-stream", width="stretch", type="primary")
    st.divider()
    st.subheader("🕑 Pontos de restauração")
    st.caption("Salve uma versão antes de grandes mudanças. O programa também cria uma cópia automática a cada 10 minutos "
               "(as 5 mais recentes ficam guardadas).")
    a, b = st.columns([3, 1])
    nome_v = a.text_input("Nome da versão", placeholder="Ex.: Antes de mudar os preços", label_visibility="collapsed")
    if b.button("💾 Salvar versão", width="stretch"):
        salvar_versao(nome_v.strip() or "Versão manual")
        st.toast("Versão salva.")
        st.rerun()
    vs = pd.read_sql("SELECT id, nome, data FROM versoes ORDER BY id DESC", conn)
    if vs.empty:
        st.caption("Nenhuma versão salva ainda.")
    for _, v in vs.iterrows():
        a, b, c = st.columns([4, 1, 1])
        a.markdown(f"**{v['nome']}** · {v['data']}")
        if b.button("Restaurar", key=f"rest_{v['id']}", width="stretch"):
            restaurar_versao(int(v["id"]))
            st.toast("Versão restaurada. Uma cópia do estado anterior foi guardada.")
            st.rerun()
        if c.button("Apagar", key=f"apv_{v['id']}", width="stretch"):
            conn.execute("DELETE FROM versoes WHERE id=?", (int(v["id"]),))
            conn.commit()
            st.rerun()


# ------------------------------------------------------------------------------
# MODO PROFESSOR
# ------------------------------------------------------------------------------
def pg_professor():
    cabecalho("professor")
    st.markdown("""
<div class="pn-card-info">
Envie os arquivos <b>.pnbox</b> dos grupos. O programa recalcula tudo, confere as checagens, mede o preenchimento e
sugere uma nota pela rubrica. A nota sugerida é só um apoio: a banca registra a nota final.
Os arquivos não são guardados; ficam só na memória desta sessão.
</div>""", unsafe_allow_html=True)
    st.write("")
    with st.expander("📐 Rubrica de avaliação (critérios e pesos)", expanded=False):
        ss.setdefault("rubrica", pd.DataFrame(RUBRICA_PADRAO, columns=["Critério", "Peso"]))
        rub = st.data_editor(ss.rubrica, num_rows="fixed", hide_index=True, width="stretch", key="ed_rubrica",
                             column_config={"Critério": T(disabled=True, width="large"), "Peso": N(min_value=0, max_value=100, step=1)})
        st.caption(f"Total: {num_br(rub['Peso'].sum(), 0)} pontos. Os critérios são calculados automaticamente; ajuste apenas os pesos.")
        rubrica = list(zip(rub["Critério"], rub["Peso"]))
    disc, profn = st.columns(2)
    disciplina = disc.text_input("Disciplina", key="prof_disc")
    professor = profn.text_input("Professor(a)", key="prof_nome")
    ss.modo_correcao = st.toggle("✍️ Modo de correção: escrever comentários dentro das ferramentas do projeto aberto",
                                 value=bool(ss.get("modo_correcao")), key="tg_correcao",
                                 help="Abra o .pnbox do grupo (menu lateral), ative este modo e navegue pelas ferramentas. Os "
                                      "comentários ficam salvos no arquivo e aparecem para o grupo no topo de cada ferramenta.")
    arqs = st.file_uploader("Arquivos .pnbox dos grupos", type=["pnbox"], accept_multiple_files=True, key="prof_up")
    if not arqs:
        st.info("Nenhum arquivo enviado ainda.")
        return
    avs = []
    with st.spinner("Avaliando os projetos..."):
        tmp = Path(tempfile.mkdtemp())
        for a in arqs:
            dados = a.getvalue()
            if not dados.startswith(b"SQLite format 3"):
                st.warning(f"{a.name}: arquivo inválido, ignorado.")
                continue
            cam = tmp / a.name
            cam.write_bytes(dados)
            try:
                p_, d_ = carregar_projeto_arquivo(cam)
                av = avaliar(p_, d_, rubrica)
                av["Arquivo"] = a.name
                avs.append(av)
            except Exception as e:      # arquivo corrompido ou muito incompleto
                st.warning(f"{a.name}: não foi possível avaliar ({e}).")
    if not avs:
        return
    tab = pd.DataFrame([{k: v for k, v in av.items() if k not in ("res", "notas")} for av in avs])
    st.subheader("📋 Painel da turma")
    mostrar = tab[["Arquivo", "Projeto", "Integrantes", "Preenchimento", "Checagens OK", "Raio-X", "Investimento",
                   "Lucro Ano 1", "VPL", "TIR", "Nota sugerida"]]
    st.dataframe(mostrar, hide_index=True, width="stretch", column_config={
        "Preenchimento": st.column_config.ProgressColumn("Preenchido", format="percent", min_value=0, max_value=1),
        "Checagens OK": st.column_config.ProgressColumn("Checagens OK", format="percent", min_value=0, max_value=1),
        "Investimento": N(format="R$ %.0f"), "Lucro Ano 1": N(format="R$ %.0f"), "VPL": N(format="R$ %.0f"),
        "TIR": N(format="percent"), "Nota sugerida": N(format="%.1f")})
    c = st.columns(3)
    c[0].metric("Projetos", len(tab))
    c[1].metric("Nota média sugerida", num_br(tab["Nota sugerida"].mean(), 1))
    c[2].metric("Projetos com VPL positivo", f"{int((tab['VPL'] > 0).sum())} de {len(tab)}")

    st.subheader("✍️ Notas da banca")
    st.caption("A tabela vem com as notas sugeridas. Altere o que quiser; as fichas em PDF usam estas notas.")
    base = pd.DataFrame([[av["Projeto"]] + av["notas"] for av in avs], columns=["Projeto"] + [c_ for c_, _ in rubrica])
    ed = st.data_editor(base, hide_index=True, width="stretch", key="ed_notas_banca",
                        column_config={"Projeto": T(disabled=True)} |
                        {c_: N(c_[:28] + "…" if len(c_) > 28 else c_, min_value=0, max_value=float(p), step=0.5)
                         for c_, p in rubrica})
    ed["Total"] = ed[[c_ for c_, _ in rubrica]].sum(axis=1)
    st.dataframe(ed[["Projeto", "Total"]], hide_index=True)
    for av, (_, linha) in zip(avs, ed.iterrows()):
        av["notas_finais"] = [float(linha[c_]) for c_, _ in rubrica]

    st.subheader("🔍 Detalhes por grupo")
    for av in avs:
        with st.expander(f"{av['Projeto']} · nota sugerida {num_br(av['Nota sugerida'], 1)}"):
            st.write(f"**Integrantes:** {av['Integrantes'] or '-'}")
            st.write(f"**Pendências:** {av['Pendências'] or 'nenhuma'}")
            st.dataframe(pd.DataFrame({"Critério": [c_ for c_, _ in rubrica], "Peso": [p for _, p in rubrica],
                                       "Sugerida": av["notas"]}), hide_index=True, width="stretch")

    carregados = []
    for a_ in arqs:
        try:
            cam = tmp / a_.name
            p_, d_ = carregar_projeto_arquivo(cam)
            carregados.append((p_.get("nome_negocio") or a_.name, p_, d_))
        except Exception:
            pass
    professor_extras(carregados)

    st.subheader("📦 Exportar")
    a, b = st.columns(2)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name="Turma", index=False)
        ed.to_excel(w, sheet_name="Notas_banca", index=False)
    a.download_button("📊 Planilha da turma (Excel)", buf.getvalue(), "avaliacao_turma.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    if b.button("📄 Gerar fichas de avaliação (PDF)", width="stretch", type="primary"):
        ss.fichas = ficha_avaliacao_pdf(avs, rubrica, professor, disciplina)
    if ss.get("fichas"):
        b.download_button("⬇️ Baixar fichas", ss.fichas, "fichas_avaliacao.pdf", "application/pdf", width="stretch")


# ------------------------------------------------------------------------------
# CONFIGURAÇÕES
# ------------------------------------------------------------------------------
def pg_config():
    cabecalho("config")
    t1, t2, t3 = st.tabs(["📚 Listas personalizadas", "💾 Onde ficam meus projetos", "🧹 Limpar este computador"])
    with t1:
        st.caption("As opções padrão continuam disponíveis. Aqui você acrescenta categorias, unidades, áreas e "
                   "classificações próprias do seu negócio; elas aparecem nas listas de escolha das ferramentas.")
        tipo = st.selectbox("Lista", list(NOMES_LISTAS), format_func=lambda k: NOMES_LISTAS[k], key="cfg_lista")
        st.markdown("**Padrão:** " + ", ".join(LISTAS_PADRAO.get(tipo, [])))
        editor("listas", {"valor": T("Novo item", required=True)}, "valor", filtro={"tipo": tipo})
    with t2:
        if NAVEGADOR:
            st.markdown("""
Você está usando a **versão on-line**. Os projetos ficam guardados **no armazenamento deste navegador, neste computador**.

- Atualizações do programa **não apagam** seus projetos: o arquivo do programa é trocado, mas a pasta de projetos do navegador continua.
- Eles **somem** se você limpar os dados de navegação do site, usar janela anônima ou trocar de navegador/computador.
- Por segurança, **baixe o .pnbox** (página Relatório e exportação) ao fim de cada sessão de trabalho.
""")
        elif MODO_WEB:
            st.warning("Nesta versão on-line temporária os projetos duram só enquanto a página estiver aberta. Baixe o .pnbox.")
        else:
            st.markdown(f"Os projetos ficam nesta pasta do computador:\n\n`{PASTA}`\n\n"
                        "Atualizar ou reinstalar o programa não apaga essa pasta.")
        st.write(f"Versão do programa: **{VERSAO_PROGRAMA}** · projetos encontrados: **{len(listar_projetos())}**")
    with t3:
        st.warning("Use em computadores de laboratório compartilhados, ao terminar. Todos os projetos guardados "
                   "**neste computador/navegador** serão apagados. Baixe antes o .pnbox do que quiser guardar.")
        ok = st.checkbox("Entendi e já baixei o que precisava", key="cfg_limpar_ok")
        if st.button("🧹 Apagar todos os projetos deste computador", disabled=not ok, type="primary"):
            conn.close()
            for p_ in listar_projetos():
                try:
                    p_.unlink()
                except OSError:
                    pass
            for k in ("db", "buf", "snap", "hash", "ultima_copia_auto"):
                ss.pop(k, None)
            ss.bemvindo_ok = False
            st.success("Pronto. Um projeto novo será criado ao recarregar.")
            st.rerun()


# ==============================================================================
# ROTEADOR
# ==============================================================================
ROTAS = {"inicio": pg_inicio, "projeto": pg_projeto, "identidade": pg_identidade, "canvas": pg_canvas, "porter": pg_porter,
         "ge": pg_ge, "bcg": pg_bcg, "bsc": pg_bsc, "reforma": pg_reforma, "segmentos": pg_segmentos, "personas": pg_personas,
         "jornada": pg_jornada, "mercado": pg_mercado, "proposta": pg_proposta, "concorrencia": pg_concorrencia,
         "forcas": pg_forcas, "oportunidades": pg_oportunidades, "swot": pg_swot, "marketing": pg_marketing, "plano_mkt": pg_plano_mkt, "sazonalidade": pg_sazonalidade,
         "canais": pg_canais, "operacional": pg_operacional, "fornecedores": pg_fornecedores,
         "investimentos": pg_investimentos, "produtos": pg_produtos, "servicos": pg_servicos, "assinaturas": pg_assinaturas, "custos": pg_custos, "preco": pg_preco,
         "dre": pg_dre, "fluxo": pg_fluxo, "indicadores": pg_indicadores, "graficos": pg_graficos,
         "simulador": pg_simulador, "experimentos": pg_experimentos, "riscos": pg_riscos, "cronograma": pg_cronograma,
         "realizado": pg_realizado, "ia": pg_ia, "sumario": pg_sumario, "apresentacao": pg_apresentacao,
         "relatorio": pg_relatorio, "professor": pg_professor, "config": pg_config,
         "evidencias": pg_evidencias, "entrevistas": pg_entrevistas, "teste_preco": pg_teste_preco, "vpc": pg_vpc,
         "oceano": pg_oceano, "posicionamento": pg_posicionamento, "orcamento_mkt": pg_orcamento_mkt,
         "montecarlo": pg_montecarlo, "adocao": pg_adocao, "prontidao": pg_prontidao, "banca": pg_banca,
         "formalizacao": pg_formalizacao, "trilha": pg_trilha, "glossario": pg_glossario, "planilha": pg_planilha,
         "historico": pg_historico, "financiamento": pg_financiamento, "dados_publicos": pg_dados_publicos,
         "referencias": pg_referencias, "swot_cruzada": pg_swot_cruzada, "lean": pg_lean, "valuation": pg_valuation,
         "impacto": pg_impacto, "editais": pg_editais, "novidades": pg_novidades}
try:
    ROTAS.get(ss.pagina, pg_inicio)()
    copia_automatica()
finally:
    try:
        conn.commit()
        conn.close()
    except Exception:
        pass
