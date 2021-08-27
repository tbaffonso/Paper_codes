#Nome: Tássia Bolotari Affonso
#Bootstrapping

from percentil2 import *

import pandas as pd
import random as ty
from random import *
import numpy
import statistics
import xlrd
import xlwt
import sys


if __name__ == "__main__":
    

    for i in range (2000,12000):
        #print(i)
        produto=i
        planilha = xlrd.open_workbook("AAA.xlsx")
        aba = planilha.sheet_by_index(0)
        leadtime=int(aba.cell_value(rowx=produto, colx=0))
        p=pd.DataFrame()

        resultado0, resultado111,resultado112,resultado113,resultado114,resultado115,resultado116,resultado117,resultado118,resultado119,resultado1110,resultado1111,resultado1112,resultado1,resultado2,resultado3,resultado4,resultado5,resultado6,resultado7,resultado8,resultado9,resultado10,resultado11,resultado12,resultado13,resultado14, resultado15, resultado16,resultado17, resultado18 = create_model(produto,leadtime)


        data = {'REAL': [resultado0,resultado0,resultado0],
        'Percentil 10': [resultado111, resultado112, resultado113],
        'Percentil 20' : [resultado114, resultado115, resultado116],
        'Percentil 30' : [resultado117, resultado118, resultado119],
        'Percentil 40':[resultado1110, resultado1111, resultado1112],
        'Percentil 50': [resultado1, resultado2, resultado3],
        'Percentil 60' : [resultado4, resultado5, resultado6],
        'Percentil 70' : [resultado7, resultado8, resultado9],
        'Percentil 80':[resultado10, resultado11, resultado12],
        'Percentil 90':[resultado13, resultado14, resultado15],
        'Percentil 100':[resultado16, resultado17, resultado18]}

        #print(data)
        if i==2000:
            tt=pd.DataFrame(data)
        else:
            p=pd.DataFrame(data)
            tt=tt.append(p)
            tt.to_excel('outputbt_todos3.xls')
            

def create_model(produto,leadtime):

    n=1000  #número de interações

    #Vetor com demandas históricas
    R1=[]
    planilha = xlrd.open_workbook("AAA.xlsx")
    aba = planilha.sheet_by_index(0)
    for u in range(132):
        R1.append(aba.cell_value(rowx=produto, colx=u+1))
    #print(R1)
    #sys.exit()

    R2=[]
    for i in range(113):
        R2.append(R1[i])

    #Outros parâmetros do problema
    bt=[]  #usado para gerar sequência de 0 e 1 no bootstrapping
    btp=[] #usado para gerar NOVA sequência de 0 e 1 no bootstrapping
    btpp=[] #vetor com jittering
    btp1=[]
    btpp1=[]
    demandas=[]
    prev=[]
    prev1=[]
    prev2=[]
    btptwo=[]
    btptwo1=[]
    btpptwo=[]
    btpu=[]


    q=1
    y=len(R1)
    x=y-19
    l=range(0,x)
    lnbt1=range(0,x-1)

    k=0
    nzeros=0
    ndemanda=0

    perc6=0
    perc66=0
    perc6two=0

    cont00=0
    cont01=0
    cont10=0
    cont11=0
    somatotal=0


    #BOOTSTRAPPING

    #Substituir demandas por 0 e 1
    for i in l:
        if R1[i]==0:
            bt.append(0)
            nzeros=nzeros+1
        else:
            bt.append(1)
            demandas.append(R1[i])
            ndemanda=ndemanda+1
            
    if ndemanda==0:
        ndemanda=1

    if nzeros==0:
        nzeros=1

    #Calculando as probabilidades
    for i in lnbt1:
        if bt[i]==0:
            if bt[i+1]==0:
                cont00=cont00+1
            else:
                cont01=cont01+1
        else:
            if bt[i+1]==0:
                cont10=cont10+1
            else:
                cont11=cont11+1

    p00= cont00/nzeros      
    p01= cont01/nzeros      
    p10= cont10/ndemanda    
    p11= cont11/ndemanda    

    somaprob=p00+p01+p10+p11


    #Boostrapping Normal - Markov

    for t in range(n):
    
    #Calculando nova sequência de 0 e 1
        for i in l:
            if i%2==0:
                o=uniform(0,somaprob)
                if o<p00:
                    btp1.append(0)
                    btp1.append(0)
                if p00<o<p00+p01:
                    btp1.append(0)
                    btp1.append(1)
                if p00+p01<o<p00+p01+p10:
                    btp1.append(1)
                    btp1.append(0)
                if p00+p01+p10<o<p00+p01+p10+p11:         
                    btp1.append(1)
                    btp1.append(1)
                else:
                    pp=0


    #Escolhendo valores + jittering
        for i in lnbt1:
            if btp1[i]==1:
                t=choice(demandas)
                media=statistics.mean(R2)
                desviopadrao=statistics.stdev(R2)
                kl= (t-media)/desviopadrao
                z=uniform(-kl,kl) 
                jit=1+int(t+z*(t**(1/2)))
                if jit<=0:
                    btpp1.append(t)
                else:
                    btpp1.append(jit)
            else:
                btpp1.append(0)

        for i in range(leadtime):
            perc66=perc66+btpp1[i]
            #print(perc66)

        btpp1.clear()
        btp1.clear()
        prev1.append(perc66)
        perc66=0


    previsaoreal1=0
    prev1.sort()
    #print(prev1)
    previsaoreal000=prev1[0]
    previsaoreal2=prev1[99]
    previsaoreal3=prev1[199]
    previsaoreal4=prev1[299]
    previsaoreal5=prev1[399]
    previsaoreal6=prev1[499]
    previsaoreal7=prev1[599]
    previsaoreal8=prev1[699]
    previsaoreal9=prev1[799]
    previsaoreal10=prev1[899]
    previsaoreal11=prev1[999]


    #Calculando valores reais no leadtime
    for i in range(leadtime):
        somatotal=somatotal+R1[i+x]
    #print(somatotal)


    #Movimento SWAP 


    #Calculando nova sequência de 0 e 1
    for i in l:
        if i%2==0:
            o=uniform(0,somaprob)
            if o<p00:
                btp.append(0)
                btp.append(0)
            if p00<o<p00+p01:
                btp.append(0)
                btp.append(1)
            if p00+p01<o<p00+p01+p10:
                btp.append(1)
                btp.append(0)
            if p00+p01+p10<o<p00+p01+p10+p11:         
                btp.append(1)
                btp.append(1)
            else:
                pp=0
                
    for t in range(n):
        #ty.seed()
        for i in l:
            btpu.append(btp[i])

        rm1=randint(0,x-1)
        rm2=randint(0,x-1)

        #print(rm1,rm2)
        valor1=btpu[rm1]
        valor2=btpu[rm2]

        btpu[rm1]=valor2
        btpu[rm2]=valor1

        #btp.append(btp1[i])
        #btpu.append(btp1[i])
        #btptwo.append(btp1[i])
        #btptwo1.append(btp1[i])

        #escolhendo valores + jittering
        for i in l:
            if btpu[i]==1:
                t=choice(demandas)
                #print(t)
                media=statistics.mean(R2)
                desviopadrao=statistics.stdev(R2)
                kl= (t-media)/desviopadrao
                z=uniform(-kl,kl)
                jit=1+int(t+z*(t**(1/2)))
                if jit<=0:
                    btpp.append(t)
                else:
                    btpp.append(jit)
            else:
                btpp.append(0)
        #print (btpp)
    

        for i in range(leadtime):
            perc6=perc6+btpp[i]
        #print(perc6)

        btpu.clear()
        btpp.clear()
        prev.append(perc6)
        perc6=0
        #print(prev)
            


    previsaoreal=0
    prev.sort()
    previsaoreal22000=prev[0]
    previsaoreal222=prev[99]
    previsaoreal223=prev[199]
    previsaoreal224=prev[299]
    previsaoreal225=prev[399]
    previsaoreal226=prev[499]
    previsaoreal227=prev[599]
    previsaoreal228=prev[699]
    previsaoreal229=prev[799]
    previsaoreal2210=prev[899]
    previsaoreal2211=prev[999]


    #Movimento 2-OPT


    #Calculando nova sequência de 0 e 1
    for i in l:
        if i%2==0:
            o=uniform(0,somaprob)
            if o<p00:
                btptwo.append(0)
                btptwo.append(0)
            if p00<o<p00+p01:
                btptwo.append(0)
                btptwo.append(1)
            if p00+p01<o<p00+p01+p10:
                btptwo.append(1)
                btptwo.append(0)
            if p00+p01+p10<o<p00+p01+p10+p11:         
                btptwo.append(1)
                btptwo.append(1)
            else:
                pp=0

    for t in range(n):
        #ty.seed()

        for i in l:
            btptwo1.append(btptwo[i])

        rr=randint(0,x-1)
        rr2=randint(0,x-1)

        if rr>rr2:
            maior=rr
            menor=rr2
        else:
            maior=rr2
            menor=rr
        #print(maior,menor)

        btptwo1[menor]=btptwo[maior]
        btptwo1[maior]=btptwo[menor]

        for i in range(menor, maior):
            maior=maior-1
            menor=menor+1
            if maior!=menor:
                btptwo1[maior]=btptwo[menor]
                btptwo1[menor]=btptwo[maior]

    
        #btptwo.append(btp1[i])
        #btptwo1.append(btp1[i])

        #escolhendo valores + jittering
        for i in l:
            if btptwo1[i]==1:
                t=choice(demandas)
                #print(t)
                media=statistics.mean(R2)
                desviopadrao=statistics.stdev(R2)
                kl= (t-media)/desviopadrao
                z=uniform(-kl,kl)
                jit=1+int(t+z*(t**(1/2)))
                if jit<=0:
                    btpptwo.append(t)
                else:
                    btpptwo.append(jit)
            else:
                btpptwo.append(0)
        #print (btpp)
    

        for i in range(leadtime):
            perc6two=perc6two+btpptwo[i]
        #print(perc6)

        btpptwo.clear()
        btptwo1.clear()
        prev2.append(perc6two)
        perc6two=0
        #print(prev)
            


    previsaoreal2=0
    prev2.sort()
    previsaoreal33000=prev2[0]
    previsaoreal332=prev2[99]
    previsaoreal333=prev2[199]
    previsaoreal334=prev2[299]
    previsaoreal335=prev2[399]
    previsaoreal336=prev2[499]
    previsaoreal337=prev2[599]
    previsaoreal338=prev2[699]
    previsaoreal339=prev2[799]
    previsaoreal3310=prev2[899]
    previsaoreal3311=prev2[999]


    #Markov, SWAP, 2-OPT
    return somatotal,previsaoreal2, previsaoreal222, previsaoreal332,previsaoreal3, previsaoreal223, previsaoreal333,previsaoreal4, previsaoreal224, previsaoreal334,previsaoreal5, previsaoreal225, previsaoreal335,previsaoreal6, previsaoreal226, previsaoreal336,previsaoreal7, previsaoreal227, previsaoreal337,previsaoreal8, previsaoreal228, previsaoreal338,previsaoreal9, previsaoreal229, previsaoreal339,previsaoreal10, previsaoreal2210, previsaoreal3310,previsaoreal11, previsaoreal2211, previsaoreal3311
