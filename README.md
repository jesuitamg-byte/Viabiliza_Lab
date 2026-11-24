# 📊 PNBOX Acadêmico

Programa gratuito para elaborar **plano de negócio e análise de viabilidade econômico-financeira**, com estrutura inspirada no PNBOX do Sebrae e voltado ao ensino de Administração.

Ele reúne, num só lugar: segmentação de mercado, proposta de valor, personas, jornada do cliente, concorrência, matriz SWOT, canais e funil de vendas, investimentos, materiais e fichas técnicas, custos, DRE de 3 anos, indicadores (VPL, TIR, payback, pontos de equilíbrio), cenários, gráficos e relatório completo em PDF e Excel.

---

## Três formas de usar

### 1. On-line, sem instalar nada (recomendado para alunos)

| Versão | Link |
|---|---|
| Com o exemplo FitBox preenchido | **https://jesuitamg-byte.github.io/gabriel_jesuita/** |
| Em branco, para começar do zero | **https://jesuitamg-byte.github.io/gabriel_jesuita/?versao=branco** |

- O programa roda **dentro do navegador** (Chrome, Edge ou Firefox). Nada é instalado e **os dados não saem do seu computador**.
- A primeira abertura leva de 30 segundos a 1 minuto.
- Ao terminar, vá à aba **Relatório** e clique em **🗄️ Baixar arquivo do projeto (.pnbox)**. Para continuar outro dia, abra esse arquivo em **📂 Abrir arquivo .pnbox**, no menu lateral.

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
- A aba **Resultados** traz checagens automáticas de consistência (fontes x investimento, margens negativas, produtos sem custo, prazos que não somam 100%, SWOT incoerente, caixa negativo).

## Arquivos deste repositório

| Arquivo | Função |
|---|---|
| `app_pnbox.py` | O programa (arquivo único, em Python/Streamlit) |
| `index.html` | Versão on-line que roda no navegador (GitHub Pages + stlite) |
| `Abrir PNBOX.bat` | Atalho para abrir no Windows com Python instalado |
| `requirements.txt` | Bibliotecas usadas |
| `installer/` e `.github/workflows/` | Geração automática dos instaladores `.exe` |
