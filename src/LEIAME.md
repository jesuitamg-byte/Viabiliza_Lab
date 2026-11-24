# Código-fonte (desenvolvimento)

O programa é escrito em módulos e depois juntado num arquivo único (`app_pnbox.py`) pelo `build.py`:

| Módulo | Conteúdo |
|---|---|
| `boot.py` | inicialização automática (porta livre, instala bibliotecas, encerra ao fechar a aba) |
| `pnbox_db.py` | estrutura do banco `.pnbox` e migração de versões antigas |
| `pnbox_modelo.py` | estudo de caso FitBox |
| `pnbox_motor.py` | todos os cálculos econômico-financeiros |
| `pnbox_pdf.py` | relatório em PDF e gráficos |
| `pnbox_app.py` | interface (Streamlit) |

Para gerar o arquivo único: `cd src && python build.py && copy app_pnbox_v4.py ..\app_pnbox.py`
