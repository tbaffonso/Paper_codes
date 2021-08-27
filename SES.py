#Nome: Tássia Bolotari Affonso
#Suavização Expoencial 

from suavizacao import *

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

    for i in range (1244,12723):
        #print(i)
        produto=i
        planilha = xlrd.open_workbook("AAA.xlsx")
        aba = planilha.sheet_by_index(0)
        leadtime=int(aba.cell_value(rowx=produto, colx=0))
        p=pd.DataFrame()
        

        resultado1,resultado2,resultado3,resultado4,r5,r6,r7,r8,r9,r10,r11 = create_model(produto,leadtime)


        data = {'REAL': [resultado1],
        'SA 0.1' : [resultado2],
        'SA 0.2' : [resultado3],
        'SA 0.3':[resultado4],
        'SA 0.4' : [r5],
        'SA 0.5' : [r6],
        'SA 0.6':[r7],
        'SA 0.7' : [r8],
        'SA 0.8' : [r9],
        'SA 0.9':[r10],
        'SA 1.0' : [r11]}

        #print(data)
        if i==1244:
            tt=pd.DataFrame(data)
        else:
            p=pd.DataFrame(data)
            tt=tt.append(p)
            tt.to_excel('SA_REAL_2.xls')

   

def create_model(produto,leadtime):
    #Vetor com demandas históricas
    R1=[]    
    d1=[]
    d2=[]
    d3=[]
    d4=[]
    d5=[]
    d6=[]
    d7=[]
    d8=[]
    d9=[]
    d10=[]

    somatotal=0
    sa01=0
    sa02=0
    sa03=0
    sa04=0
    sa05=0
    sa06=0
    sa07=0
    sa08=0
    sa09=0
    sa10=0

    planilha = xlrd.open_workbook("AAA.xlsx")
    aba = planilha.sheet_by_index(0)
    for u in range(132):
        R1.append(aba.cell_value(rowx=produto, colx=u+1))

    qq=0
    y=len(R1)
    x=y-19

    for i in range(132):
        d1.append(0.0)
        d2.append(0.0)
        d3.append(0.0)
        d4.append(0.0)
        d5.append(0.0)
        d6.append(0.0)
        d7.append(0.0)
        d8.append(0.0)
        d9.append(0.0)
        d10.append(0.0)
        
#----
    alfa=0.1

    for i in range(131):
        if R1[i]==0:
            d1[i]=R1[i]
        else:
            d1[i+1]=alfa*R1[i]+(1-alfa)*d1[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        somatotal=somatotal+R1[i+x]
        sa01=sa01+d1[i+x]
        #print(d1[i+x])
    #sys.exit()
#----
    alfa1=0.2
    
    for i in range(131):
        if R1[i]==0:
            d2[i]=R1[i]
        else:
            d2[i+1]=alfa1*R1[i]+(1-alfa1)*d2[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa02=sa02+d2[i+x]
#----
    alfa2=0.3
    
    for i in range(131):
        if R1[i]==0:
            d3[i]=R1[i]
        else:
            d3[i+1]=alfa2*R1[i]+(1-alfa2)*d3[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa03=sa03+d3[i+x]

#------
    alfa4=0.4
    
    for i in range(131):
        if R1[i]==0:
            d4[i]=R1[i]
        else:
            d4[i+1]=alfa4*R1[i]+(1-alfa4)*d4[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa04=sa04+d4[i+x]

#------
#------
    alfa5=0.5
    
    for i in range(131):
        if R1[i]==0:
            d5[i]=R1[i]
        else:
            d5[i+1]=alfa5*R1[i]+(1-alfa5)*d5[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa05=sa05+d5[i+x]

#------------
    alfa6=0.6
    
    for i in range(131):
        if R1[i]==0:
            d6[i]=R1[i]
        else:
            d6[i+1]=alfa6*R1[i]+(1-alfa6)*d6[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa06=sa06+d6[i+x]

#------------
    alfa7=0.7
    
    for i in range(131):
        if R1[i]==0:
            d7[i]=R1[i]
        else:
            d7[i+1]=alfa7*R1[i]+(1-alfa7)*d7[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa07=sa07+d7[i+x]

#------------
    alfa8=0.8
    
    for i in range(131):
        if R1[i]==0:
            d8[i]=R1[i]
        else:
            d8[i+1]=alfa8*R1[i]+(1-alfa8)*d8[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa08=sa08+d8[i+x]

#------------
    alfa9=0.9
    
    for i in range(131):
        if R1[i]==0:
            d9[i]=R1[i]
        else:
            d9[i+1]=alfa9*R1[i]+(1-alfa9)*d9[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa09=sa09+d9[i+x]

#------------
    alfa10=1.0
    
    for i in range(131):
        if R1[i]==0:
            d10[i]=R1[i]
        else:
            d10[i+1]=alfa10*R1[i]+(1-alfa10)*d10[i]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        #somatotal=somatotal+R1[i+x]
        sa10=sa10+d10[i+x]

#------


    return somatotal,sa01, sa02, sa03,sa04,sa05,sa06,sa07,sa08,sa09,sa10