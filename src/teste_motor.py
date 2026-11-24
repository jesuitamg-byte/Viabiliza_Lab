import sys, os, tempfile
from pathlib import Path
exec(open('pnbox_db.py',encoding='utf-8').read())
exec(open('pnbox_modelo.py',encoding='utf-8').read())
exec(open('pnbox_motor.py',encoding='utf-8').read())
p=Path(tempfile.mkdtemp())/'t.pnbox'
criar_projeto_modelo(p)
conn=conectar(p)
proj=ler_unica(conn,'projeto')
d={t: ler(conn,t) for t in SCHEMA if t not in ('projeto','proposta_valor')}
d['proposta_valor']=ler_unica(conn,'proposta_valor')
res=executar_motor(proj,d)
c=res['cat']
print("qtd por categoria:", c.groupby('categoria')['qtd_mensal'].sum().to_dict())
print("overhead", round(res['overhead_m']), "mod", round(res['mod_m']), "cif", round(res['cif_m']), "markup", round(res['markup_mult'],3), "df%", round(res['df_pct'],3))
print(c[['descricao','custo_unit','custo_absorcao_unit','preco_venda','preco_sugerido','pe_unidades_mes']].head(8).round(2).to_string())
print(res['fluxo_mensal'][['Mês','Saldo final']].round(0).T.to_string())
print(sensibilidade(res).round(0).to_string())
print(mercado(proj,res))
print(capacidade_demanda(d,res).round(2).to_string())
for s_ in semaforo(res): print(s_['Status'], s_['Indicador'], s_['Valor'])
chk=tabela_checagens_v5(res,d,proj); print(chk[['Teste','Situação','destino']].to_string())
print(raio_x(res,chk,0.8,semaforo(res)))
