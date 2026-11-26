import re
partes=["pnbox_db.py","pnbox_modelo.py","pnbox_motor.py","pnbox_estrategia.py","pnbox_validacao.py","pnbox_v8.py","pnbox_mix.py","pnbox_leituras.py","pnbox_v9.py","pnbox_pdf.py","pnbox_extras.py","pnbox_abnt.py","pnbox_ui_base.py","pnbox_ui_paginas1.py","pnbox_ui_paginas2.py","pnbox_ui_paginas4.py","pnbox_ui_paginas5.py","pnbox_ui_paginas6.py","pnbox_ui_paginas7.py","pnbox_ui_paginas8.py","pnbox_ui_paginas3.py"]
cab='''"""
================================================================================
VIABILIZALAB v8.9 (antigo PNBOX Acadêmico) - PLANO DE NEGÓCIO E VIABILIDADE ECONÔMICO-FINANCEIRA
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
'''
imports=set(); corpo=[]
for p in partes:
    txt=open(p,encoding='utf-8').read()
    txt=re.sub(r'^"""[\s\S]*?"""\n','',txt,count=1)
    linhas=[]
    for l in txt.splitlines():
        if re.match(r'\s*from pnbox_\w+ import',l):
            # multi-line import: skip until closing paren
            linhas.append('#SKIP'+l); continue
        linhas.append(l)
    txt="\n".join(linhas)
    # remove multi-line pnbox imports
    txt=re.sub(r'#SKIP[^\n]*\([^)]*\)','',txt)
    txt=re.sub(r'#SKIP[^\n]*','',txt)
    out=[]
    for l in txt.splitlines():
        if re.match(r'^(import \w|from \w[\w.]* import )',l) and not l.startswith('from pnbox'):
            imports.add(l); continue
        out.append(l)
    corpo.append(f"\n\n# {'#'*78}\n# PARTE: {p}\n# {'#'*78}\n"+"\n".join(out))
BOOT=open('boot.py',encoding='utf-8').read()
try:
    _chg = open('../CHANGELOG.md', encoding='utf-8').read().replace("'" * 3, '"' * 3)
except OSError:
    _chg = "Historico de versoes indisponivel."
corpo.insert(0, "\n\n# Historico de versoes (tela Novidades), gerado de CHANGELOG.md\nNOVIDADES_MD = r" + "'" * 3 + _chg + "'" * 3 + "\n")
open('app_pnbox.py','w',encoding='utf-8').write(cab+BOOT+"\n".join(sorted(imports))+"\n"+"".join(corpo)+"\n")

import shutil; shutil.copy('app_pnbox.py','../app_pnbox.py'); print('ok')
