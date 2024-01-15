from csv_ical import Convert
import csv
import xlwt
import matplotlib.pyplot as plt

### Conversion du fichier ics au format csv ###
convert = Convert()
convert.CSV_FILE_LOCATION = "ADECal.csv"
convert.SAVE_LOCATION = "ADECal.ics"
convert.read_ical(convert.SAVE_LOCATION)
convert.make_csv()
convert.save_csv(convert.CSV_FILE_LOCATION)

###Initialisation des dictionnaires avec les matières utiliser et les salles associer au type de cours ###

##la matière est par rapport au prof
##La matière projet tuteuré n'est pas prise en compte par absence de salle, notre programme étant basé sur ces dernières il nous est donc impossible de comtabiliser les quelsques heures de cette "matière"
dico_1={"math" :["Mathématiques des transmissions","Mathématiques du signal","Mathématiques des systèmes numériques","Mathématiques du signal DS","Matrices et graphes","DS Matrices et graphes"],"réseau":["CISCO","EXAMEN CISCO","initiation réseau","Principes et architecture des réseaux","Technologie de l'internet","DS Initiation réseaux","SAE11: Cybersécurité","SAE12: Réseaux","DS Principes et architecture des réseaux","Réseaux locaux et équipements actifs","DS Réseaux opérateurs","DS Services réseaux avancés","DS TP Initiation réseaux","Réseaux locaux et équipements actifs PKT","Réseaux opérateurs MR","Réseaux WLAN","SAE12: Réseaux","DS Réseaux locaux et équipements actifs","Services réseaux avancés"],"télécommunication":["Electronique","DS TP blanc Electronique","Fibres optiques","Transmissions guidées","Supports de transmission pour les réseaux locaux","Electromagnétisme","Signaux et systèmes pour les transmissions","SAE13: Telecommunication","DS Fibres optiques","DS Transmissions guidées","DS Electromagnétisme","Technologies d'accès","Infrastructure sans fil","Infrastructure sans fi","DS Technologies d'accès","Transmission large bande","DS Infrastructure sans fil","DS Transmission large bande","DS Electronique","DS TP Infrastructure sans fil","Réseaux cellulaires","DS TP Electronique","DS Réseaux opérateurs SM","Réseaux opérateurs SM"],"info":["Architecture des systèmes numériques et informatiques","Analyse et traitement de données structurées","Administration système","Scripts avancés","Bases des systèmes d'exploitation","Applications informatiques","Gestion d'annuaires unifiés","Fondamentaux de la programmation","Introduction aux technologies Web","SAE15: Traiter des données","SAE14: Se présenter sur Internet","Sources de données","Fondamentaux de la programmation DS TP","DS TP Fondamentaux de la programmation","DS Architecture des systèmes numériques et informatiques","DS TP Gestion d'annuaires unifiés","Scripts avancés DS TP","DS TP Applications informatiques","DS Gestion d'annuaires unifiés","DS TP Scripts avancés","DS Administration système"],"autres":["Anglais DS","DS Communication","Communication","Anglais","Anglais Technique","Anglais général","PPP","Expression Culture Comm pro - Outils info","Expression Culture Communications professionnelles","Gestion de projet","PortFolio - CISCO","Projet Personnel et Professionnel","DS Anglais Dijkstra","DS Anglais","DS Expression Culture Communications professionnelles"]}
dico_2={"cm":["RT-Amphi"],"tp":["TEAMS","RT-Labo Electronique 1","RT-Labo Informatique 1","RT-Labo Informatique 2","RT-Labo Informatique 3","RT-Labo Telecoms 1","RT-Labo Telecoms 2","RT-Labo reseaux 1","RT-Labo reseaux 2",""],"td":["RT-Salle-TD1","RT-Salle-TD2","RT-Salle-TD3","RT-Salle-TD4"]}
L=[]
with open("ADECal.csv",newline='',encoding = 'utf-8') as csvfile:
    reader=csv.reader(csvfile,delimiter=",")
    for row in reader:
        if row != [] :            ###Cette ligne permet d'enlever les listes vides qui sont les retours a la ligne dans le fichier csv
            L.append(row)
            
####Fonction####

def annee (a):
    p = []
    for i in L:
        if i[3][3] != 'A':
            if str(a) == '1' :
                if i[3][2] == '1' or i[3][2] == 'B' or i[3][2] == 'F':
                    p.append(i)
            elif str(a) == '2':
                if i[3][2] == '2':
                    p.append(i)

    return p
            
            
def matiere(liste,mat,type):   ###Ici liste vaut la liste azerty, mat correspond a la matière
    td=[]
    tp=[]
    cm=[]
    a = []
    for values in mat:
        for v in type :
            for f in liste:
                if f[0] in values:
                    if f[-1] in type["cm"] or f[-1][3] == 'A':
                        a = [f[1],f[2],f[3],f[-1]]
                        cm.append(a)
                    elif f[-1] in type["tp"] :
                        a = [f[1],f[2],f[3],f[-1]]
                        tp.append(a)
                    elif f[-1] in type["td"] :
                        a = [f[1],f[2],f[3],f[-1]]
                        td.append(a)
    return [cm,td,tp]
 
      
    
def heure(liste):
    h = [0,0,0]
    x = 0
    y = 0
    for i in range(3):
        for d in liste[i] :
            if d[0][14] + d[0][15] == d[1][14] + d[1][15]:
                x = int(d[0][11] + d[0][12])
                y = int(d[1][11] + d[1][12])
                if x < y :
                    h[i] += y - x
                    del liste[i][0]
    return h

def personne(liste,annee):
    liste2 = []
    if annee == '1' :
        liste2.append(liste[0])
        liste2.append(liste[1]//4)
        liste2.append(liste[2]//6)
    elif annee == '2':
        liste2.append(liste[0])
        liste2.append(liste[1]//4)
        liste2.append(liste[2]//4)
    return liste2

annee_BUT=str(input("quelle annee ? "))
azerty=annee(annee_BUT)

#print(azerty)

#ma = str(input("Choisissez la matière que vous souhaiter : "))
#print(matiere(azerty,"Projets tuteurés - Présentation des sujets",dico_2)) 

m=[]        # contiendra les heures des matières
for cle in dico_1.keys():
    mat = matiere(azerty,dico_1[cle],dico_2)
    heure_m = heure(mat)
    m.append(personne(heure_m,annee_BUT))


#print(m)      Permet d'afficher la liste qui contient toute les matières de l'année voulu.



cm=[]
td=[]
tp=[]

for i in range(len(m)):
    cm.append(m[i][0])
    td.append(m[i][1])
    tp.append(m[i][2])

print(cm)
print(td)
print(tp)



# au départ nous allions utiliser un histogramme ace matplotlib mais par souci de 
#simplicté nous avond décidé d'utiliser une autre solution avec les diagrammes barpour obtenir le même rendu

x = ["Maths","reseau","Telecom","Info","Autres"]

plt.bar(x, cm, color='r')
plt.bar(x, td, bottom=cm, color='g')
plt.bar(x, tp, bottom=[td[i]+cm[i] for i in range(len(td))], color='b')
plt.show()




###Ne pas oublier de diviser les td par deux et par trois les tp pour avoir le nombre d'heure d'un étudiant en rt1 et rt2 faut diviser td et tp par deux car deux groupes.

#print(personne(heure_m,annee_BUT))








#Création du tableur
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('feuille1')
if annee_BUT == '1':
    sheet.write(0, 3, "RT1")  ##(numéro ligne commence à 0, numéro colonne commence à 0,'str à écrire'ou int)
elif annee_BUT == '2' :
    sheet.write(0, 3, "RT2")
sheet.write(2, 0, "CM")
sheet.write(3, 0, "TD")
sheet.write(4, 0, "TP")
sheet.write(1, 1, "Mathématiques")
sheet.write(1, 2, "Réseau")
sheet.write(1, 3, "Télécommunication")
sheet.write(1, 4, "Informatique")
sheet.write(1, 5, "Autres")
for i in range(5):
    sheet.write(2, i+1, cm[i])
for i in range(5):
    sheet.write(3, i+1, td[i])
for i in range(5):
    sheet.write(4, i+1, tp[i])
workbook.save('myFile.xls')

#Faire 1 histogrammes parramètrable pour les deux années
