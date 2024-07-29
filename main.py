import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from fpdf import FPDF
from matplotlib import pyplot as plt

# Escolhendo ano e mês do relatório

while True:
    escolha_usuario = input('Digite o mês e ano do relatório [formato mm/aaaa]: ')
    try:
        data_rel = datetime.strptime(escolha_usuario, '%m/%Y')
        break
    except ValueError:
        print('Favor digitar mes e ano válidos no formato MM/AAAA')
        print('Exemplo: agosto de 2021 -> 08/2021\n')

meses = ('', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
         'Novembro', 'Dezembro', 'Janeiro', 'Fevereiro')

if 12 > data_rel.month >= 3:
    ano_letivo = f'{data_rel.year}/{data_rel.year + 1}'
    inicio = datetime.strptime(f'01/03/{data_rel.year}', '%d/%m/%Y')
    fim = datetime.strptime(f'01/{data_rel.month+1}/{data_rel.year}', '%d/%m/%Y') - relativedelta(days=1)
elif data_rel.month == 12:
    ano_letivo = f'{data_rel.year}/{data_rel.year + 1}'
    inicio = datetime.strptime(f'01/03/{data_rel.year}', '%d/%m/%Y')
    fim = datetime.strptime(f'01/01/{data_rel.year + 1}', '%d/%m/%Y') - relativedelta(days=1)
else:
    ano_letivo = f'{data_rel.year - 1}/{data_rel.year}'
    inicio = datetime.strptime(f'01/03/{data_rel.year - 1}', '%d/%m/%Y')
    fim = datetime.strptime(f'01/{data_rel.month+1}/{data_rel.year}', '%d/%m/%Y') - relativedelta(days=1)
periodo_rel = (inicio, fim)

# Determinando residentes ativos no período

sheet_res_url = 'residentes.csv'
db = pd.read_csv(f'{sheet_res_url}', parse_dates=[1], dayfirst=True)
residentes_r1 = db.loc[(db['ENTRADA'] >= periodo_rel[0]) & (db['ENTRADA'] <= periodo_rel[1])]
residentes_r2 = db.loc[(db['ENTRADA'] >= periodo_rel[0] - relativedelta(months=12)) & (db['ENTRADA'] < periodo_rel[0])]
residentes_r3 = db.loc[(db['ENTRADA'] >= periodo_rel[0] - relativedelta(months=24)) &
                       (db['ENTRADA'] < periodo_rel[0] - relativedelta(months=12))]

# Criando DATAFRAME CIRURGIAS

sheet_url = 'caderneta.csv'
df = pd.read_csv(f'{sheet_url}', usecols=[2, 5, 6, 9], parse_dates=[1], dayfirst=True)
df = df.loc[(df['DATA DA CIRURGIA'] >= periodo_rel[0]) & (df['DATA DA CIRURGIA'] <= periodo_rel[1])]
df.rename({'RESIDENTE ': 'RESIDENTE', 'DATA DA CIRURGIA': 'DATA', 'NÍVEL DE PARTICIPAÇÃO': 'PARTICIPAÇÃO'}, axis=1,
          inplace=True)

# Incorporando residentes ativos no DATAFRAME

count = 1
for i in (residentes_r1, residentes_r2, residentes_r3):
    for item in i['NOME'].to_list():
        df.loc[df['RESIDENTE'] == item, 'ANO'] = f'r{count}'
        if item in i.loc[i['DESISTIU'] == 'SIM', 'NOME'].to_list():
            df.loc[df['RESIDENTE'] == item, 'DESISTIU'] = 'SIM'
        else:
            df.loc[df['RESIDENTE'] == item, 'DESISTIU'] = 'NÃO'
    count += 1

# Criado DataFrame (df) com as seguintes colunas:
# ['RESIDENTE', 'DATA', 'ESPECIALIDADE', 'PARTICIPAÇÃO', 'ANO', 'DESISTIU']
# Em seguida será gerado o PDF do relatório.

# Gerando o PDF final
pdf = FPDF(orientation='L', unit='mm', format='A4')
epw = pdf.w - 2 * pdf.l_margin
cell_width = epw / 100

# Página inicial
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt='', ln=1, align='C')
pdf.cell(287, 10, txt='Hospital Universitário', ln=1, align='C')
pdf.cell(287, 10, txt='Serviço de Ortopedia e Traumatologia', ln=1, align='C')
pdf.image('logo.png', x=105, y=40, w=100, h=100)
pdf.cell(287, 105, txt='', ln=1, align='C')
pdf.cell(287, 10, txt='Relatório Residentes - Participação em Cirurgias', ln=1, align='C')
pdf.cell(287, 10, txt=f'Ano Letivo {ano_letivo}', ln=1, align='C')
pdf.cell(287, 10, txt=f'Mês de {meses[data_rel.month]}', ln=1, align='C')

# Relatório acumulado ano
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Cirurgias Acumuladas - Ano Letivo {ano_letivo}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')
inicio_x = pdf.get_x()
inicio_y = pdf.get_y()
cell_height = pdf.font_size * 1.5

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 35, cell_height, 'Residente', border=1)
pdf.cell(cell_width * 15, cell_height, 'Acumulado', align='R', border=1)
pdf.ln(cell_height)

pdf.set_font('Arial', size=12)
resultado = df.loc[df['DESISTIU'] != 'SIM']['RESIDENTE'].value_counts()
desistentes = df.loc[df['DESISTIU'] == 'SIM']['RESIDENTE'].value_counts()
for item in resultado.index:
    pdf.cell(cell_width * 35, cell_height, f'{item}', border=1)
    pdf.cell(cell_width * 15, cell_height, f'{resultado[item]}', align='R', border=1)
    pdf.ln(cell_height)

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 35, cell_height, 'Total de cirurgias', border=1)
pdf.cell(cell_width * 15, cell_height, f'{resultado.sum() + desistentes.sum()}', align='R', border=1)
pdf.ln(cell_height)

pdf.set_x(inicio_x)
pdf.set_y(inicio_y)

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 55, cell_height)
pdf.cell(cell_width * 9, cell_height, 'Médias', border=1)
pdf.cell(cell_width * 9, cell_height, 'R1', align='C', border=1)
pdf.cell(cell_width * 9, cell_height, 'R2', align='C', border=1)
pdf.cell(cell_width * 9, cell_height, 'R3', align='C', border=1)
pdf.cell(cell_width * 9, cell_height, 'Total', align='C', border=1)
pdf.ln(cell_height)

pdf.set_font('Arial', size=12)
r1_num = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].nunique()
r2_num = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].nunique()
r3_num = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].nunique()
num_meses = (periodo_rel[1].year - periodo_rel[0].year) * 12 + (periodo_rel[1].month - periodo_rel[0].month) + 1
for i in range(0, num_meses):
    r1_cx = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') & (df['DATA'] <= periodo_rel[1] -
                                                                      relativedelta(months=i))].shape[0]
    r2_cx = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') & (df['DATA'] <= periodo_rel[1] -
                                                                      relativedelta(months=i))].shape[0]
    r3_cx = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') & (df['DATA'] <= periodo_rel[1] -
                                                                      relativedelta(months=i))].shape[0]
    pdf.cell(cell_width * 55, cell_height)
    pdf.cell(cell_width * 9, cell_height, f'{meses[(periodo_rel[1] - relativedelta(months=i)).month]}', border=1)
    pdf.cell(cell_width * 9, cell_height, f'{round(r1_cx/r1_num, 1)}', align='C', border=1)
    pdf.cell(cell_width * 9, cell_height, f'{round(r2_cx/r2_num, 1)}', align='C', border=1)
    pdf.cell(cell_width * 9, cell_height, f'{round(r3_cx/r3_num, 1)}', align='C', border=1)
    pdf.cell(cell_width * 9, cell_height, f'{round((r1_cx + r2_cx + r3_cx)/(r1_num + r2_num + r3_num), 1)}', align='C',
             border=1)
    pdf.ln(cell_height)

# Gráficos acumuladas ano
pdf.add_page()
acumuladas = dict()
nomes = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].unique().tolist()
acumuladas['Residentes'] = nomes
for i in range(num_meses-1, -1, -1):
    mes = periodo_rel[1] - relativedelta(months=i)
    acumuladas[f'{meses[mes.month]}'] = list()
    for i2 in nomes:
        numero = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') &
                        (df['DATA'] <= mes) & (df['RESIDENTE'] == i2)].shape[0]
        acumuladas[f'{meses[(periodo_rel[1] - relativedelta(months=i)).month]}'].append(numero)
grafico = pd.DataFrame(data=acumuladas)
ax = grafico.set_index('Residentes').T.plot(kind='line', grid=True, style='o-', figsize=(9, 5),
                                            color=['red', 'black', '#ffad33', '#009933', '#0033cc', 'grey', '#990099'],
                                            title='Cirurgias Acumuladas R1')
plt.savefig('acumulada_r1.png', bbox_inches='tight', dpi=100)
pdf.image('acumulada_r1.png', x=10, y=10)

pdf.add_page()
acumuladas = dict()
nomes = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].unique().tolist()
acumuladas['Residentes'] = nomes
for i in range(num_meses-1, -1, -1):
    mes = periodo_rel[1] - relativedelta(months=i)
    acumuladas[f'{meses[mes.month]}'] = list()
    for i2 in nomes:
        numero = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') &
                        (df['DATA'] <= mes) & (df['RESIDENTE'] == i2)].shape[0]
        acumuladas[f'{meses[(periodo_rel[1] - relativedelta(months=i)).month]}'].append(numero)
grafico = pd.DataFrame(data=acumuladas)
ax = grafico.set_index('Residentes').T.plot(kind='line', grid=True, style='o-', figsize=(9, 5),
                                            color=['red', 'black', '#ffad33', '#009933', '#0033cc', 'grey', '#990099'],
                                            title='Cirurgias Acumuladas R2')
plt.savefig('acumulada_r2.png', bbox_inches='tight', dpi=100)
pdf.image('acumulada_r2.png', x=10, y=10)

pdf.add_page()
acumuladas = dict()
nomes = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].unique().tolist()
acumuladas['Residentes'] = nomes
for i in range(num_meses-1, -1, -1):
    mes = periodo_rel[1] - relativedelta(months=i)
    acumuladas[f'{meses[mes.month]}'] = list()
    for i2 in nomes:
        numero = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') &
                        (df['DATA'] <= mes) & (df['RESIDENTE'] == i2)].shape[0]
        acumuladas[f'{meses[(periodo_rel[1] - relativedelta(months=i)).month]}'].append(numero)
grafico = pd.DataFrame(data=acumuladas)
ax = grafico.set_index('Residentes').T.plot(kind='line', grid=True, style='o-', figsize=(9, 5),
                                            color=['red', 'black', '#ffad33', '#009933', '#0033cc', 'grey', '#990099'],
                                            title='Cirurgias Acumuladas R3')
plt.savefig('acumulada_r3.png', bbox_inches='tight', dpi=100)
pdf.image('acumulada_r3.png', x=10, y=10)
plt.clf()

# Relatório, cirurgias no mês
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Cirurgias Mensais - {meses[data_rel.month]}/{data_rel.year}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 35, cell_height, 'Residente', border=1)
pdf.cell(cell_width * 15, cell_height, 'Cirurgias', align='R', border=1)
pdf.ln(cell_height)

pdf.set_font('Arial', size=12)
resultado = df.loc[df['DATA'] > periodo_rel[1] - relativedelta(months=1)]['RESIDENTE'].value_counts()
for x in resultado.index.tolist():
    pdf.cell(cell_width * 35, cell_height, f'{x}', border=1)
    pdf.cell(cell_width * 15, cell_height, f'{resultado[x]}', align='R', border=1)
    pdf.ln(cell_height)

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 35, cell_height, 'Total de cirurgias', border=1)
pdf.cell(cell_width * 15, cell_height, f'{resultado.sum()}', align='R', border=1)
pdf.ln(cell_height)

pdf.set_x(inicio_x)
pdf.set_y(inicio_y)

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 55, cell_height)
pdf.cell(cell_width * 10, cell_height, 'Ano', border=1)
pdf.cell(cell_width * 10, cell_height, 'Média', align='R', border=1)
pdf.ln(cell_height)

pdf.set_font('Arial', size=12)
resultado = df.loc[df['DATA'] > periodo_rel[1] - relativedelta(months=1)]['ANO'].value_counts()
for x in sorted(resultado.index.tolist()):
    num_res = df.loc[(df['ANO'] == x) & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].nunique()
    pdf.cell(cell_width * 55, cell_height)
    pdf.cell(cell_width * 10, cell_height, x.upper(), border=1)
    pdf.cell(cell_width * 10, cell_height, f'{round(resultado[x]/num_res, 1)}', align='R', border=1)
    pdf.ln(cell_height)
pdf.ln(cell_height)

# Gráfico de cirurgias no mês
pdf.add_page()
grafico = df.loc[df['DATA'] > periodo_rel[1] - relativedelta(months=1)]['RESIDENTE'].value_counts().\
    reindex(df.loc[df['DESISTIU'] != 'SIM']['RESIDENTE'].unique(), fill_value=0)
ax = grafico.plot(kind='bar', rot=0, width=0.6, figsize=(9, 5), edgecolor='black', title='Cirurgias do Mês',
                  color=['#301437', '#511a75', '#5e4dab', '#6683bd', '#8db0c5', '#ced3d9', '#dccecb', '#ca9c80',
                         '#b96456', '#953250', '#5d1749'])
ax.set_xticklabels([x.replace(' ', '\n') for x in [x.replace(' de', '') for x in grafico.index.tolist()]])
ax.bar_label(ax.containers[0])
plt.savefig('cirurgias_mes.png', bbox_inches='tight', dpi=100)
pdf.image('cirurgias_mes.png', x=10, y=10)
plt.clf()

# Relatório grau de participação no mês
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Grau de Participação - {meses[data_rel.month]}/{data_rel.year}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=12)
pdf.cell(cell_width * 30, cell_height, 'Residente', border=1)
pdf.set_font('Arial', 'B', size=10)
pdf.cell(cell_width * 14, cell_height, 'Observei', align='C', border=1)
pdf.cell(cell_width * 14, cell_height, 'Auxiliei', align='C', border=1)
pdf.cell(cell_width * 14, cell_height, 'Realizei <50%', align='C', border=1)
pdf.cell(cell_width * 14, cell_height, 'Realizei >50%', align='C', border=1)
pdf.cell(cell_width * 14, cell_height, 'Total sob supervisão', align='C', border=1)
pdf.ln(cell_height)

pdf.set_font('Arial', size=12)
resultado = df.loc[(df['DATA'] > periodo_rel[1] - relativedelta(months=1))][['RESIDENTE', 'PARTICIPAÇÃO']].value_counts(
    )
for x in df.loc[df['DESISTIU'] != 'SIM']['RESIDENTE'].unique().tolist():
    pdf.cell(cell_width * 30, cell_height, f'{x}', border=1)
    try:
        pdf.cell(cell_width * 14, cell_height, str(resultado[x]['OBSERVEI']), align='C', border=1)
    except KeyError:
        pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
    try:
        pdf.cell(cell_width * 14, cell_height, str(resultado[x]['AUXILIEI']), align='C', border=1)
    except KeyError:
        pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
    try:
        pdf.cell(cell_width * 14, cell_height, str(resultado[x]['REALIZEI MENOS DE 50%']), align='C', border=1)
    except KeyError:
        pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
    try:
        pdf.cell(cell_width * 14, cell_height, str(resultado[x]['REALIZEI MAIS DE 50%']), align='C', border=1)
    except KeyError:
        pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
    try:
        pdf.cell(cell_width * 14, cell_height, str(resultado[x]['REALIZADO TOTAL SOB SUPERVISÃO']), align='C', border=1)
    except KeyError:
        pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
    pdf.ln(cell_height)

pdf.set_font('Arial', 'B', size=12)
resultado = df.loc[(df['DATA'] > periodo_rel[1] - relativedelta(months=1))][['PARTICIPAÇÃO']].value_counts()
pdf.cell(cell_width * 30, cell_height, 'Total Geral', border=1)
try:
    pdf.cell(cell_width * 14, cell_height, str(resultado['OBSERVEI']), align='C', border=1)
except KeyError:
    pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
try:
    pdf.cell(cell_width * 14, cell_height, str(resultado['AUXILIEI']), align='C', border=1)
except KeyError:
    pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
try:
    pdf.cell(cell_width * 14, cell_height, str(resultado['REALIZEI MENOS DE 50%']), align='C', border=1)
except KeyError:
    pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
try:
    pdf.cell(cell_width * 14, cell_height, str(resultado['REALIZEI MAIS DE 50%']), align='C', border=1)
except KeyError:
    pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
try:
    pdf.cell(cell_width * 14, cell_height, str(resultado['REALIZADO TOTAL SOB SUPERVISÃO']), align='C', border=1)
except KeyError:
    pdf.cell(cell_width * 14, cell_height, '0', align='C', border=1)
pdf.ln(cell_height)

# Relatório proporção de participação no ano
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Proporção de Participação - Ano Letivo {ano_letivo}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=10)
cell_height = pdf.font_size * 1.5
pdf.cell(cell_width * 8, cell_height, 'R1', border=1)
pdf.cell(cell_width * 8, cell_height, 'Observei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Auxiliei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Menos 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Mais 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Total', align='C', border=1)
pdf.ln(cell_height)
for i in range(num_meses, 0, -1):
    r1_obs = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'OBSERVEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r1_aux = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'AUXILIEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r1_menos = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MENOS DE 50%')
                      & (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                      (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r1_mais = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MAIS DE 50%') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r1_tudo = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') &
                     (df['PARTICIPAÇÃO'] == 'REALIZADO TOTAL SOB SUPERVISÃO') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r1_cx = df.loc[(df['ANO'] == 'r1') & (df['DESISTIU'] != 'SIM') &
                   (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                   (df['DATA'] <= periodo_rel[0] + relativedelta(months=i + 1) - relativedelta(days=1))].shape[0]
    pdf.set_font('Arial', 'B', size=10)
    pdf.cell(cell_width * 8, cell_height, meses[i+2], border=1)
    pdf.set_font('Arial', size=10)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r1_obs/r1_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r1_aux/r1_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r1_menos/r1_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r1_mais/r1_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r1_tudo/r1_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    pdf.ln(cell_height)

pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=10)
cell_height = pdf.font_size * 1.5
pdf.cell(cell_width * 8, cell_height, 'R2', border=1)
pdf.cell(cell_width * 8, cell_height, 'Observei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Auxiliei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Menos 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Mais 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Total', align='C', border=1)
pdf.ln(cell_height)
for i in range(num_meses, 0, -1):
    r2_obs = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'OBSERVEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r2_aux = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'AUXILIEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r2_menos = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MENOS DE 50%')
                      & (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                      (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r2_mais = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MAIS DE 50%') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r2_tudo = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') &
                     (df['PARTICIPAÇÃO'] == 'REALIZADO TOTAL SOB SUPERVISÃO') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r2_cx = df.loc[(df['ANO'] == 'r2') & (df['DESISTIU'] != 'SIM') &
                   (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                   (df['DATA'] <= periodo_rel[0] + relativedelta(months=i + 1) - relativedelta(days=1))].shape[0]
    pdf.set_font('Arial', 'B', size=10)
    pdf.cell(cell_width * 8, cell_height, meses[i+2], border=1)
    pdf.set_font('Arial', size=10)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r2_obs/r2_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r2_aux/r2_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r2_menos/r2_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r2_mais/r2_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r2_tudo/r2_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    pdf.ln(cell_height)

pdf.set_x(inicio_x)
pdf.set_y(inicio_y)

pdf.set_font('Arial', 'B', size=10)
cell_height = pdf.font_size * 1.5
pdf.cell(cell_width * 52, cell_height)
pdf.cell(cell_width * 8, cell_height, 'R3', border=1)
pdf.cell(cell_width * 8, cell_height, 'Observei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Auxiliei', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Menos 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Mais 50%', align='C', border=1)
pdf.cell(cell_width * 8, cell_height, 'Total', align='C', border=1)
pdf.ln(cell_height)
for i in range(num_meses, 0, -1):
    r3_obs = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'OBSERVEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r3_aux = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'AUXILIEI') &
                    (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                    (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r3_menos = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MENOS DE 50%')
                      & (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                      (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r3_mais = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') & (df['PARTICIPAÇÃO'] == 'REALIZEI MAIS DE 50%') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r3_tudo = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') &
                     (df['PARTICIPAÇÃO'] == 'REALIZADO TOTAL SOB SUPERVISÃO') &
                     (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                     (df['DATA'] <= periodo_rel[0] + relativedelta(months=i+1) - relativedelta(days=1))].shape[0]
    r3_cx = df.loc[(df['ANO'] == 'r3') & (df['DESISTIU'] != 'SIM') &
                   (df['DATA'] >= periodo_rel[0] + relativedelta(months=i-1)) &
                   (df['DATA'] <= periodo_rel[0] + relativedelta(months=i + 1) - relativedelta(days=1))].shape[0]
    pdf.cell(cell_width * 52, cell_height)
    pdf.set_font('Arial', 'B', size=10)
    pdf.cell(cell_width * 8, cell_height, meses[i+2], border=1)
    pdf.set_font('Arial', size=10)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r3_obs/r3_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r3_aux/r3_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r3_menos/r3_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r3_mais/r3_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    try:
        pdf.cell(cell_width * 8, cell_height, f'{round(r3_tudo/r3_cx*100, 1)}%', align='C', border=1)
    except ZeroDivisionError:
        pdf.cell(cell_width * 8, cell_height, f'0%', align='C', border=1)
    pdf.ln(cell_height)

# Relatório cirurgias por especialidade
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Cirurgias por Especialidade - {meses[data_rel.month]}/{data_rel.year}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=9)
cell_height = pdf.font_size * 1.5
top_y = pdf.get_y()
pdf.multi_cell(cell_width * 21.35, cell_height * 2, 'Residentes', border=1)
pdf.set_font('Arial', 'B', size=8)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 21.35)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Pediátrica', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 28.5)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Coluna', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 35.65)
pdf.multi_cell(cell_width * 7.15, cell_height, 'Ombro e\nCotovelo', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 42.8)
pdf.multi_cell(cell_width * 7, cell_height * 2, 'Mão', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 49.8)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Quadril', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 56.95)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Joelho', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 64.1)
pdf.multi_cell(cell_width * 7.15, cell_height, 'Pé e\nTornozelo', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 71.25)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Trauma', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 78.4)
pdf.multi_cell(cell_width * 7.6, cell_height, 'Reconstrução\nAlongamento', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 86)
pdf.multi_cell(cell_width * 7, cell_height*2, 'Tumor', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 93)
pdf.multi_cell(cell_width * 7, cell_height*2, 'Plantão', align='C', border=1)

top_y = pdf.get_y()
pdf.set_font('Arial', size=9)
cell_height = pdf.font_size * 2
resultado = df.loc[df['DATA'] > periodo_rel[1] - relativedelta(months=1)][['RESIDENTE', 'ESPECIALIDADE']].value_counts()
for x in df.loc[df['DESISTIU'] != 'SIM']['RESIDENTE'].unique().tolist():
    pdf.multi_cell(cell_width * 21.35, cell_height, f'{x}', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 21.35)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['PEDIÁTRICA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 28.5)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['COLUNA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 35.65)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['OMBRO E COTOVELO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 42.8)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['MÃO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 49.8)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['QUADRIL']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 56.95)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['JOELHO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 64.1)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['PÉ E TORNOZELO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 71.25)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['TRAUMA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 78.4)
    try:
        pdf.multi_cell(cell_width * 7.6, cell_height, str(resultado[x]['RECONSTRUÇÃO E ALONGAMENTO']), align='C',
                       border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.6, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 86)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['TUMOR']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 93)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['PLANTÃO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    top_y = pdf.get_y()

# Gráfico cirurgias por especialidade
pdf.add_page()
especialidades = dict()
nomes = df.loc[(df['DATA'] > periodo_rel[1] - relativedelta(months=1)) & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].\
    value_counts().index.tolist()
esp = df['ESPECIALIDADE'].unique().tolist()
especialidades['Residentes'] = nomes
for i in esp:
    especialidades[i] = list()
    for i2 in nomes:
        numero = df.loc[(df['DESISTIU'] != 'SIM') & (df['RESIDENTE'] == i2) &
                        (df['DATA'] > periodo_rel[1] - relativedelta(months=1)) & (df['ESPECIALIDADE'] == i)].shape[0]
        especialidades[i].append(numero)
grafico = pd.DataFrame(data=especialidades)
try:
    ax = grafico.plot(kind='bar', stacked=True, x='Residentes', xlabel='', rot=0, width=0.8, figsize=(9, 5),
                      edgecolor='black', title='Cirurgias por Especialidade', colormap='Paired')
    ax.set_xticklabels([x.replace(' ', '\n') for x in [x.replace(' de', '') for x in grafico['Residentes'].tolist()]])
    plt.savefig('cx_especialidades.png', bbox_inches='tight', dpi=100)
    pdf.image('cx_especialidades.png', x=10, y=10)
except IndexError:
    pdf.cell(epw, cell_height, 'GRÁFICO: sem registro de cirurgias no mês.', align='C', border=0)
plt.clf()

# Relatório cirurgias totais por especialidade
pdf.add_page()
pdf.set_font('Arial', 'B', size=14)
pdf.cell(287, 10, txt=f'Cirurgias Realizadas Total - {meses[data_rel.month]}/{data_rel.year}', ln=1, align='C')
pdf.cell(287, 10, txt='', ln=1, align='C')

pdf.set_font('Arial', 'B', size=9)
cell_height = pdf.font_size * 1.5
top_y = pdf.get_y()
pdf.multi_cell(cell_width * 21.35, cell_height * 2, 'Residentes', border=1)
pdf.set_font('Arial', 'B', size=8)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 21.35)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Pediátrica', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 28.5)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Coluna', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 35.65)
pdf.multi_cell(cell_width * 7.15, cell_height, 'Ombro e\nCotovelo', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 42.8)
pdf.multi_cell(cell_width * 7, cell_height * 2, 'Mão', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 49.8)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Quadril', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 56.95)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Joelho', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 64.1)
pdf.multi_cell(cell_width * 7.15, cell_height, 'Pé e\nTornozelo', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 71.25)
pdf.multi_cell(cell_width * 7.15, cell_height * 2, 'Trauma', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 78.4)
pdf.multi_cell(cell_width * 7.6, cell_height, 'Reconstrução\nAlongamento', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 86)
pdf.multi_cell(cell_width * 7, cell_height*2, 'Tumor', align='C', border=1)
pdf.set_y(top_y)
pdf.set_x(pdf.l_margin + cell_width * 93)
pdf.multi_cell(cell_width * 7, cell_height*2, 'Plantão', align='C', border=1)

top_y = pdf.get_y()
pdf.set_font('Arial', size=9)
cell_height = pdf.font_size * 2
resultado = df.loc[(df['DATA'] > periodo_rel[1] - relativedelta(months=1)) &
                   (df['PARTICIPAÇÃO'] == 'REALIZADO TOTAL SOB SUPERVISÃO')][['RESIDENTE',
                                                                              'ESPECIALIDADE']].value_counts()
for x in df.loc[df['DESISTIU'] != 'SIM']['RESIDENTE'].unique().tolist():
    pdf.multi_cell(cell_width * 21.35, cell_height, f'{x}', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 21.35)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['PEDIÁTRICA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 28.5)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['COLUNA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 35.65)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['OMBRO E COTOVELO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 42.8)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['MÃO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 49.8)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['QUADRIL']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 56.95)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['JOELHO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 64.1)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['PÉ E TORNOZELO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 71.25)
    try:
        pdf.multi_cell(cell_width * 7.15, cell_height, str(resultado[x]['TRAUMA']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.15, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 78.4)
    try:
        pdf.multi_cell(cell_width * 7.6, cell_height, str(resultado[x]['RECONSTRUÇÃO E ALONGAMENTO']), align='C',
                       border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7.6, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 86)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['TUMOR']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    pdf.set_y(top_y)
    pdf.set_x(pdf.l_margin + cell_width * 93)
    try:
        pdf.multi_cell(cell_width * 7, cell_height, str(resultado[x]['PLANTÃO']), align='C', border=1)
    except KeyError:
        pdf.multi_cell(cell_width * 7, cell_height, '0', align='C', border=1)
    top_y = pdf.get_y()

# Gráfico cirurgias totais por especialidade
pdf.add_page()
especialidades = dict()
nomes = df.loc[(df['DATA'] > periodo_rel[1] - relativedelta(months=1)) & (df['DESISTIU'] != 'SIM')]['RESIDENTE'].\
    value_counts().index.tolist()
esp = df['ESPECIALIDADE'].unique().tolist()
especialidades['Residentes'] = nomes
for i in esp:
    especialidades[i] = list()
    for i2 in nomes:
        numero = df.loc[(df['DESISTIU'] != 'SIM') & (df['RESIDENTE'] == i2) &
                        (df['PARTICIPAÇÃO'] == 'REALIZADO TOTAL SOB SUPERVISÃO') &
                        (df['DATA'] > periodo_rel[1] - relativedelta(months=1)) & (df['ESPECIALIDADE'] == i)].shape[0]
        especialidades[i].append(numero)
grafico = pd.DataFrame(data=especialidades)
try:
    ax = grafico.plot(kind='bar', stacked=True, x='Residentes', xlabel='', rot=0, width=0.8, figsize=(9, 5),
                      edgecolor='black', title='Cirurgias Realizadas Total por Especialidade', colormap='Paired')
    ax.set_xticklabels([x.replace(' ', '\n') for x in [x.replace(' de', '') for x in grafico['Residentes'].tolist()]])
    plt.savefig('total_especialidades.png', bbox_inches='tight', dpi=100)
    pdf.image('total_especialidades.png', x=10, y=10)
except IndexError:
    pdf.cell(epw, cell_height, 'GRÁFICO: sem registro de cirurgias no mês.', align='C', border=0)
plt.clf()

# Encerrando
pdf.output(f'relatorio_{data_rel.month}_{data_rel.year}.pdf')
print(f'\nRelatório de {meses[data_rel.month]}/{data_rel.year} gerado com sucesso!')
