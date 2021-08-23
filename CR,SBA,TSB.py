#Nome: Tássia Bolotari Affonso
#Croston, SBA e TSB

from sba import *

from random import *
import statistics
import random as ty
import numpy
import statistics
import xlrd
import xlwt
import pandas as pd
import sys

if __name__ == "__main__":

    for i in range (12723):
        #print(i)
        produto=i
        planilha = xlrd.open_workbook("AAA.xlsx")
        aba = planilha.sheet_by_index(0)
        leadtime=int(aba.cell_value(rowx=produto, colx=0))
        p=pd.DataFrame()
        

        resultado1,resultado2,resultado3,resultado4 = create_model(produto,leadtime)


        data = {'REAL': [resultado1],
        'CROSTON' : [resultado2],
        'SBA' : [resultado3],
        'TSB':[resultado4]}

        #print(data)
        if i==10000:
            tt=pd.DataFrame(data)
        else:
            p=pd.DataFrame(data)
            tt=tt.append(p)
            tt.to_excel('outputsba10000-12000.xls')

   

def create_model(produto,leadtime):

    ylt=[]
    zt=[]
    pt=[]
    sba=[]
    tsb=[]
    zlt=[]
    dlt=[]


    alfa=0.05
    beta=0.05

    somatsb=0
    somacroston=0
    somasba=0

    somatotal=0
    #Vetor com demandas históricas
    R1=[]
    planilha = xlrd.open_workbook("AAA.xlsx")
    aba = planilha.sheet_by_index(0)
    for u in range(132):
        R1.append(aba.cell_value(rowx=produto, colx=u+1))


    #3 períodos de aquecimento

    qq=0
    y=len(R1)
    x=y-19
    l=range(0,36)
    #lnbt1=range(24,x-1)
    #lnbt2=range(1,36)
    l1=range(36,x+leadtime)

    for i in range(0,x+leadtime):
        zt.append(0.0)
        pt.append(0.0)
        zlt.append(0.0)
        dlt.append(0.0)
        ylt.append(0.0)
        sba.append(0.0)
        tsb.append(0.0)

    pt[0]= 0
    #print(pt)

    #Período de aquecimento de 3 anos 
    for i in l:
        if R1[i]==0:
            zt[i]=zt[i-1]
            pt[i]=pt[i-1]
            qq=qq+1
            print(i)
            print(zt[i])
            print(pt[i])
            print(qq)
        else:
            zt[i]=zt[i-1]+alfa*(R1[i]-zt[i-1])
            pt[i]=pt[i-1]+alfa*(qq-pt[i-1])
            qq=1
            print(i)
            print(zt[i])
            print(pt[i])
            print(qq)
            #print(pt[i])
    #print(pt)


    previsaocroston=[]
    previsaosba=[]
    previsaotsb=[]

    #CROSTON
    #print("Previsões pelo método de Croston")
    for i in l1:
        if R1[i]==0:
            zt[i]=zt[i-1]
            pt[i]=pt[i-1]
            qq=qq+1
            #print(zt)
            #print(pt)
        else:
            zt[i]=zt[i-1]+alfa*(R1[i]-zt[i-1])
            pt[i]=pt[i-1]+alfa*(qq-pt[i-1])
            qq=1
        if pt[i]==0:
            ylt[i]=0
        else:
            ylt[i]=zt[i]/pt[i]
        previsaocroston.append(ylt[i])
        #print(ylt[i])
    #print(len(previsaocroston))
    #sys.exit()

    #SBA
    #print("\n")
    #print("Previsões pelo método SBA")
    for i in l1:
        if R1[i]==0:
            zt[i]=zt[i-1]
            pt[i]=pt[i-1]
            qq=qq+1
            #print(zt)
            #print(pt)
        else:
            zt[i]=zt[i-1]+alfa*(R1[i]-zt[i-1])
            pt[i]=pt[i-1]+alfa*(qq-pt[i-1])
            qq=1
        if pt[i]==0:
            sba[i]=0
        else:
            sba[i]=(1-(alfa/2))*zt[i]/pt[i]
        previsaosba.append(sba[i])
        #print(sba[i])


    #TSB
    #Período de aquecimento de três anos
    for i in l:
        if R1[i]==0:
            zlt[i]=zlt[i-1]
            dlt[i]=dlt[i-1]+beta*(0-dlt[i-1])
        else:
            zlt[i]=zlt[i-1]+alfa*(R1[i]-zlt[i-1])
            dlt[i]=dlt[i-1]+beta*(1-dlt[i-1])

    #print("\n")
    #print("Previsões pelo método TSB")
    for i in l1:
        if R1[i]==0:
            zlt[i]=zlt[i-1]
            dlt[i]=dlt[i-1]+beta*(0-dlt[i-1])
        else:
            zlt[i]=zlt[i-1]+alfa*(R1[i]-zlt[i-1])
            dlt[i]=dlt[i-1]+beta*(1-dlt[i-1])
        tsb[i]=dlt[i]*zlt[i]
        previsaotsb.append(tsb[i])
        #print(tsb[i])

    #Calculando valores reais no leadtime
    for i in range(leadtime):
        somatotal=somatotal+R1[i+x]
    #print("Demanda real:")
    #print(somatotal)

    for i in range(leadtime):
        somacroston=somacroston+previsaocroston[len(previsaocroston)-leadtime]
        somasba=somasba+previsaosba[len(previsaosba)-leadtime]
        somatsb=somatsb+previsaotsb[len(previsaotsb)-leadtime]
    #print(somatotal,somacroston,somasba,somatsb)
    #sys.exit()
    return somatotal,somacroston,somasba,somatsb