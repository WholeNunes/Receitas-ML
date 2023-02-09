import pandas as pd
import numpy as np

###RECEITAS DO MERCADO LIVRE!

df = pd.read_excel('CAMINHO DO ARQUIVO', skiprows=[0,1])
writer = pd.ExcelWriter('CAMINHO DE DESTINO', engine='xlsxwriter')

def limpeza(self):
    global dfLimpeza
    global dadosnulos
    global dfKIT
    global dadosdoPACOTE
    global dadosdosKITs

    dfLimpeza = self
    dfLimpeza.update(dfLimpeza['Receita por envio (BRL)'].fillna(0))
    dfLimpeza.update(dfLimpeza['Receita por produtos (BRL)'].fillna(0))
    dfLimpeza.update(dfLimpeza['Custo de envio'].fillna(0))
    dfLimpeza.update(dfLimpeza['Tarifa de venda e impostos'].fillna(0))
    dfLimpeza.update(dfLimpeza['Total (BRL)'].fillna(0))
    dfLimpeza['SKU'] = dfLimpeza['SKU'].replace(" ", np.nan)
    dfLimpeza['NF-e em anexo'] = dfLimpeza['NF-e em anexo'].replace(" ", np.nan)
    dfLimpeza['Variação'] = dfLimpeza['Variação'].replace(" ", "Padrão")
    dfLimpeza['SKU'] = dfLimpeza['SKU'].str.replace("OP", "PA")
    dfLimpeza.rename(columns={'Receita por produtos (BRL)': 'Receita Total'}, inplace=True)
    dfLimpeza.rename(columns={'Total (BRL)': 'Receita sem impostos'}, inplace=True)
    dfLimpeza.rename(columns={'Tarifa de venda e impostos': 'Tarifas'}, inplace=True)
    dfLimpeza['Receita + Envio'] = dfLimpeza['Receita Total'] + dfLimpeza['Receita por envio (BRL)']
    dfLimpeza['Tarifa Total'] = dfLimpeza['Tarifas'] + dfLimpeza['Custo de envio']


    dadosnulos = dfLimpeza[dfLimpeza['SKU'].str.contains('Autorizado').isnull()] #CONCAT
    dfLimpeza.dropna(subset=['SKU'], inplace=True)

    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Devolução a caminho']
    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Devolverão o seu produto']
    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Devolução em preparação']
    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Devolução para revisar']
    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Devolução finalizada com reembolso para o comprador']
    dfLimpeza = dfLimpeza[dfLimpeza['Status'] != 'Cancelada pelo comprador']

    dfKIT = dfLimpeza[dfLimpeza['SKU'].str.contains('KIT', na=False)] #CONCAT
    dfKITdrop = dfKIT.index
    dfLimpeza = dfLimpeza.drop(dfKITdrop)

    dadosdoPACOTE = dfLimpeza[dfLimpeza['NF-e em anexo'].isnull()] #CONCAT
    dadosdoPACOTEdrop = dadosdoPACOTE.index
    dfLimpeza = dfLimpeza.drop(dadosdoPACOTEdrop)

    dadosdosKITs = dfLimpeza[dfLimpeza['Título do anúncio'].str.contains('Kit')] #CONCAT
    dadosdosKITsdrop = dadosdosKITs.index
    dfLimpeza = dfLimpeza.drop(dadosdosKITsdrop)

    dfLimpeza['SKU'] = dfLimpeza['SKU'].str[:6]

def receitas(self):
    self = self[['SKU', 'Título do anúncio', 'Variação', 'Unidades', 'Receita sem impostos', 'Receita + Envio',
            'Receita Total']].groupby(['SKU', 'Variação', 'Título do anúncio']).sum()
    self.to_excel(writer, sheet_name='Receitas')

def despesas(self):
    self = self[['SKU', 'Título do anúncio', 'Variação', 'Unidades','Custo de envio', 'Tarifas', 'Tarifa Total']].groupby(
        ['SKU', 'Variação', 'Título do anúncio']).sum()
    self.to_excel(writer, sheet_name='Despesas')

def ml(self):
    self = self[['SKU', 'Título do anúncio', 'Variação', 'Unidades','Preço unitário de venda do anúncio (BRL)', 'Receita Total']].groupby(
        ['SKU']).agg({'Unidades':'sum', 'Preço unitário de venda do anúncio (BRL)':'mean', 'Receita Total':'sum'})
    self.loc['Total'] = self.sum()
    self.to_excel(writer, sheet_name='ML')

def funcaoKIT(self):
    self = pd.concat([dfKIT, dadosdoPACOTE, dadosdosKITs, dadosnulos])
    self = self[['SKU', 'Título do anúncio', 'Variação', 'Unidades','Preço unitário de venda do anúncio (BRL)', 'Receita Total']].sort_index()
    self.loc['Total'] = self.sum()
    self.to_excel(writer, sheet_name='KITs e Pacotes')


limpeza(df)
receitas(dfLimpeza)
despesas(dfLimpeza)
ml(dfLimpeza)
funcaoKIT(dfLimpeza)

writer.save()
