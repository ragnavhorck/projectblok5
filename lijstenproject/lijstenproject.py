#-------------------------------------------------------------------------------
# Name:        
# Purpose:     Inlezen van alle xml bestanden
#              spectra eruit halen
#               in lijsten zetten              
#
# Author:      Ernst Hamer
#
# Created:     02-10-2017
# Copyright:   (c) Ernst Hamer 2016
#              Hogeschool van Arnhem en Nijmegen
#-------------------------------------------------------------------------------
import xlrd
import os

def main():
    
    #print(os.listdir()

    
    paden = [ 'SUPPTAB_S03 Differentially regulated proteins in control (KR) 0-24 h after DEX treatment.xlsx',
             'SUPPTAB_S04 Differentially regulated proteins in wild type DD 0-24 h after DEX treatment.xlsx',
             'SUPPTAB_S05 Differentially regulated proteins mpk3 0-24 h after DEX treatment.xlsx',
             'SUPPTAB_S06 Differentially regulated proteins in mpk6 0-24 h after DEX treatment.xlsx']
    
    lijstprotein = []
    
    for path in paden:
        file = openfile(path)
        bestand = readfile(file)
        lijstprotein.append(bestand)
    
#   snijding = lijstprotein[0].intersection(lijstprotein[1])
#   print(len(lijstprotein[1]))
    
    for t in range(1,len(lijstprotein)):
        lijstprotein[t] = lijstprotein[t] - lijstprotein[0] #het wegfilteren van de negatieve controle
    
    #0:    KR
    #1:    DD
    #2:    MPK3
    #3:    MPK6
    
    print(len(lijstprotein[1])) #lengte lijst DD
    print(len(lijstprotein[2])) #lengte lijst MPK3
    print(len(lijstprotein[3])) #lengte lijst MPK6
 #   print (float(ugh+pls+doe))
    
    print("lijst DD " + str(lijstprotein[1])) #lijst met alle niet weggefilterde accessiecodes van DD
    print() #zorgt voor een lege regel in de console
    print("lijst MPK3 " +str(lijstprotein[2])) #lijst met alle niet weggefilterde accessiecodes van MPK3
    print() #zorgt voor een lege regel in de console
    print("lijst MPK6 " + str(lijstprotein[3]))#lijst met alle niet weggefilterde accessiecodes van MPK6
    print() #zorgt voor een lege regel in de console
    overlapje = lijstprotein[1].intersection(lijstprotein[2]).intersection(lijstprotein[3])
    print("de overlapping van de DD & MKP3 & MPK6" + str(overlapje))
 #  print("check: "+ str(overlapje.intersection(snijding)))
 #  print("check: "+str(snijding))
 
    newfile1= open("lijstgefilterdDD.txt","w") #maakt een txt bestand aan waar de lijst met DD accessiecode gescheiden is met een enter
    
    for accessie in lijstprotein[1]:
        newfile1.write(accessie)
        newfile1.write("\n")
        
    newfile2= open("lijstgefilterdMPK3.txt","w")#maakt een txt bestand aan waar de lijst met MKP3 accessiecode gescheiden is met een enter
    
    for accessie in lijstprotein[2]:
        newfile2.write(accessie)
        newfile2.write("\n")
        
    newfile3= open("lijstgefilterdMPK6.txt","w")#maakt een txt bestand aan waar de lijst met MPK6 accessiecode gescheiden is met een enter
    
    for accessie in lijstprotein[3]:
        newfile3.write(accessie)
        newfile3.write("\n")
    
        
        
    
    
def openfile(path):
    try:
        file = xlrd.open_workbook(path) #leest het bestand in 
        return file 
    except IOError:
        print("bestand niet gevonden") #nette exeption error als het bestand niet goed staat weergeven
        quit()
        
def readfile(file):
    print (file.nsheets) #print het aantal sheets
    print (file.sheet_names()) #print de namen van de sheets
    first_sheet = file.sheet_by_index(3)#gaat naar de sheet met alle proteins 
    print()#zorgt voor een lege regel in de console
    print(first_sheet.nrows) #leest een row van het exel bestand
    
    stommelijst = []#lege lijst
    
    for t in range(3,first_sheet.nrows):
        
        rij = first_sheet.row_values(t)
        if(float(rij[3])>2 and float(rij[2]) < 0.05): #filtert op het feit dat max fold > 2 en anova <0,05
            stommelijst.append(rij[0])
        
        #print(first_sheet.row_values(t)[0])
    return set(stommelijst)
    """
    print (first_sheet.row_values(3)[0])
    print()

    cell = first_sheet.cell(0,0)
    print (cell)
    print (cell.value)
    print (first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=5))
    """
    '''
    def nieuwefile(lijstprotein):
    
    newfile= open("lijstgefilterd.txt","w")
    
    for accessie in lijstprotein[1]:
        newfile.write(accessie)
        newfile.write("\n")
'''
main()
