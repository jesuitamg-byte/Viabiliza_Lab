# 📊 PNBOX Acadêmico

Programa gratuito para elaborar **plano de negócio e análise de viabilidade econômico-financeira**, com estrutura inspirada no PNBOX do Sebrae e voltado ao ensino de Administração.

**Versão 6.** São 35 ferramentas em cartões grandes, organizadas como no PNBOX, para negócios de **produtos, serviços e assinaturas**:

- **Cliente e mercado:** segmentos, personas, jornada do cliente, tamanho do mercado (TAM/SAM/SOM).
- **Estratégia:** propostas de valor, concorrência, forças e fraquezas, oportunidades e ameaças, matriz SWOT.
- **Marketing e operação:** 4Ps, canais e funil, localização, layout, capacidade x demanda, organograma, fornecedores.
- **Finanças:** investimentos e fontes, produtos e fichas técnicas, custos classificados (MOD, CIF, despesas), formação de preço por *markup* e custeio por absorção, DRE de 3 anos, fluxo de caixa mensal, indicadores com semáforo, gráficos, simulador "e se...?", sensibilidade e cenários.
- **Serviços:** custo da hora produtiva (equipe, horas contratadas e taxa de utilização), custo pleno e preço sugerido de cada serviço, ocupação da equipe.
- **Assinaturas e receita recorrente:** planos mensais, trimestrais ou anuais projetados por safras de assinantes, com churn, MRR, ARR, ARPU, CAC, LTV, LTV/CAC, receita antecipada no balanço e custos de tecnologia (nuvem, APIs, lojas de aplicativos, meio de pagamento).
- **Tributos:** calculadora do Simples Nacional (anexos I a V e fator R).
- **Complementares:** quadro de experimentação, riscos, cronograma de implantação, orçado x realizado.
- **Entrega:** texto pronto para análise com IA gratuita, sumário executivo automático, modo apresentação, relatório PDF, *pitch deck*, Word, Excel e modo professor.

Também tem: salvamento automático a cada alteração, ponto de restauração automático a cada 10 minutos, tema claro/escuro, listas personalizáveis (categorias, unidades, áreas) três casos de exemplo (FitBox, Café Mucuri e Conecta Vale, um aplicativo por assinatura) e layout adaptado para celular.

---

## Três formas de usar

### 1. On-line, sem instalar nada (recomendado para alunos)

| Versão | Link |
|---|---|
| Com o exemplo FitBox preenchido | **https://jesuitamg-byte.github.io/gabriel_jesuita/** |
| Em branco, para começar do zero | **https://jesuitamg-byte.github.io/gabriel_jesuita/?versao=branco** |
| Com o exemplo de assinatura (Conecta Vale) | **https://jesuitamg-byte.github.io/gabriel_jesuita/?versao=assinatura** |

- O programa roda **dentro do navegador** (Chrome, Edge ou Firefox). Nada é instalado e **os dados não saem do seu computador**.
- A primeira abertura leva de 30 segundos a 1 minuto.
- Os projetos ficam guardados **no próprio navegador**. Atualizações do programa não os apagam; eles só somem se você limpar os dados de navegação, usar janela anônima ou trocar de computador.
- Por segurança, ao terminar vá a **Relatório e exportação** e baixe o **.pnbox**. Para continuar em outro computador, abra esse arquivo em **📂 Abrir arquivo .pnbox**, no menu lateral.
- Em laboratório compartilhado, use **Configurações → 🧹 Limpar este computador** ao sair.

### 2. Instalador para Windows (.exe)

Na página **[Releases](https://github.com/jesuitamg-byte/gabriel_jesuita/releases/latest)** há dois instaladores:

- `PNBOX-Setup-Modelo-FitBox.exe`: abre com o exemplo preenchido;
- `PNBOX-Setup-Em-Branco.exe`: abre com um projeto vazio.

Não é preciso ter Python. O instalador cria um atalho na área de trabalho. Se o Windows exibir "O Windows protegeu o computador", clique em **Mais informações → Executar assim mesmo**. Os projetos ficam em `Documentos\PNBOX`. Ao fechar a aba do navegador, o programa se encerra sozinho.

### 3. Com Python já instalado

Baixe `app_pnbox.py` e `Abrir PNBOX.bat`, coloque os dois na mesma pasta e dê dois cliques no `.bat`. Na primeira vez ele instala as bibliotecas necessárias.

---

## Para o professor

- Cada projeto é um arquivo `.pnbox` (banco SQLite), que o aluno pode entregar junto com o PDF.
- As checagens automáticas de consistência têm o botão **VERIFICAR → ir ao ajuste**, que leva o aluno direto à ferramenta onde está o problema.
- O cartão **👩‍🏫 Modo professor** recebe vários `.pnbox` de uma vez e monta o painel da turma (preenchimento, checagens, VPL, TIR), sugere nota por rubrica com pesos editáveis, permite registrar a nota da banca e gera as fichas de avaliação em PDF e a planilha da turma.

## Arquivos deste repositório

| Arquivo | Função |
|---|---|
| `app_pnbox.py` | O programa (arquivo único, em Python/Streamlit) |
| `index.html` | Versão on-line que roda no navegador (GitHub Pages + stlite) |
| `Abrir PNBOX.bat` | Atalho para abrir no Windows com Python instalado |
| `requirements.txt` | Bibliotecas usadas |
| `src/` | Código-fonte em módulos; `python build.py` gera o `app_pnbox.py` |
| `installer/` e `.github/workflows/` | Geração automática dos instaladores `.exe` |
