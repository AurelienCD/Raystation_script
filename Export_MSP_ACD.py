# -*- coding: utf-8 -*-
from connect import *
#import pandas as pd
import numpy as np
import os
#import pandas as pd
#from Indices_modulation import fonction_calculate_MCSv_LT
#from openpyxl import Workbook
#from openpyxl import load_workbook


from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter import ttk

import datetime

###############################################################################################################################
###############################################################################################################################
###############################################################################################################################
###############################################################################################################################

def fonction_calculate_MCSv_LT(patient,case,plan):

 the_plan = plan
 #the_plan = get_current('Plan')
 #List_Name_Plans.append(the_plan.Name)
 the_beamset = get_current('BeamSet')
 
 ##################################
 ## Calcul du nombre d'arcs, de CP et de MU dans le plan
 ##
 ###############################

 # Nombre d'arcs
 Nombre_Arcs = the_beamset.Beams.Count

 # Nombre de CP
 List_Nombre_CP=[]
 for i_arc in range(Nombre_Arcs):
  List_Nombre_CP.append(the_beamset.Beams[i_arc].Segments.Count)

 # Nombre de MU
 List_Nombre_MU=[]
 for i_arc in range(Nombre_Arcs):
  List_Nombre_MU.append(the_beamset.Beams[i_arc].BeamMU)

 Total_MU = sum(List_Nombre_MU)
 #List_MU_Plans.append(Total_MU)
 
 print(" \n ------------------------------------------------------------------------------- \n ")
 print("  Nom du Plan :", the_plan.Name) 
 #print("  N° du Plan :", num_plan)
 print("  Le nombre d'arcs:", Nombre_Arcs)
 print("  Le nombre de CP pour chaque arc:", List_Nombre_CP)
 print("  Le nombre de MU pour chaque arc:", List_Nombre_MU)
 print("  Le nombre total de MU:", Total_MU)
 print(" \n  ")


 ##################################
 ## Structure des données en Array 4D: 
 ## Array = [arc][CP][Gauche ou Droite][indice de la lame]
 ##################################
 
 list_Arc_Array=[]                       # liste pour stocker les positions des lames de chaque CP et de chaque arc
 list_Arc_Weight=[]                       # liste pour stocker les poids de chaque CP dans chaque arc
 list_indice_jaw_Y=[]

 for i_arc in the_beamset.Beams:
  list_CP_Array=[]                      # liste pour stocker les positions des lames de chaque CP
  list_CP_Weight=[]                      # liste pour stocker les poids de chaque CP (le poids = MU_CP/MU_Arc)
  list_indice_jaw_Y_CP=[]
  
  for i_cp in i_arc.Segments:
  
   Temp_Array_G_D = np.copy(i_cp.LeafPositions)            # ici on recupere de Raystation une liste de Array avec les positions des lames du cote gauche et droit
   Temp_Array_G_D = np.array(Temp_Array_G_D)               # ici on transforme la liste des positions des lames sous la forme d'une array 2D [Gauche/Droite][indice_de_lame]
   list_CP_Array.append(Temp_Array_G_D)
   
   Temp_Weight = i_cp.RelativeWeight
   list_CP_Weight.append(Temp_Weight)
   
   jaw_position_Y1 = i_cp.JawPositions[2]     # Position des machoires. Les indice [2] et [3] correspondent respectivement aux directions Y1 et Y2 (bas et et haut) 
   jaw_position_Y2 = i_cp.JawPositions[3]     # Remarque: La position des machoires peut changer d'un CP à un autre
   indice_jaw_position_Y1=int((jaw_position_Y1+19.75)/0.5) + 1                 # Indice de la lame dans le champ de traitement (limite du bas du "open field"). Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y1 + 1    (car l'indexation dans python commence à 0 et le TPS on commence à 1)
   indice_jaw_position_Y2=int((jaw_position_Y2+19.75)/0.5) + 1                 # Indice de la première lame en dehors du champ. Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y2      (car dans l'indexation start:end le end n'est pas inclus)  
   list_indice_jaw_Y_CP.append([indice_jaw_position_Y1,indice_jaw_position_Y2])
      
   
   
  list_CP_Array = np.array(list_CP_Array)                 # ici on transforme la liste des CP en array 3D [CP][G/D][indice_de_lame]
  list_Arc_Array.append(list_CP_Array)
  
  list_CP_Weight=np.array(list_CP_Weight)
  list_Arc_Weight.append(list_CP_Weight)
  
  list_indice_jaw_Y_CP = np.array(list_indice_jaw_Y_CP)
  list_indice_jaw_Y.append(list_indice_jaw_Y_CP)
  
 list_Arc_Array = np.array(list_Arc_Array)                 # ici on transforme la liste des Arcs en array à 4D [Arc][CP][G/D][indice_de_lame]
 list_Arc_Weight = np.array(list_Arc_Weight)
 list_indice_jaw_Y = np.array(list_indice_jaw_Y)
 #print("\n indice y \n",list_indice_jaw_Y)
 #print("\n indice y shape \n",list_indice_jaw_Y.shape)
 
 ##################################
 ## Calcul du LSV
 ##
 ##################################
 
 list_Arc_Array_LSV=[]                      # liste pour stocker les LSV de chaque CP pour chaque arc

 for i_arc in range(Nombre_Arcs):
  list_CP_Array_LSV=[]
  
  for i_cp_ in range(List_Nombre_CP[i_arc]):
   Array_LSV_CP_G = np.copy(list_Arc_Array[i_arc,i_cp_,0])            # array avec les positions des lames gauches de l'arc "i_arc" et du CP "i_cp_"   (on boucle sur tout les CP et tout les arcs) 
   Array_LSV_CP_D = np.copy(list_Arc_Array[i_arc,i_cp_,1])            # array avec les positions des lames droites de l'arc "i_arc" et du CP "i_cp_"   (on boucle sur tout les CP et tout les arcs)

   indice_jaw_position_Y1=list_indice_jaw_Y[i_arc,i_cp_,0]             # Indice de la lame dans le champ de traitement (limite du bas du "open field"). Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y1 + 1    (car l'indexation dans python commence à 0 et le TPS on commence à 1)
   indice_jaw_position_Y2=list_indice_jaw_Y[i_arc,i_cp_,1]                # Indice de la première lame en dehors du champ. Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y2      (car dans l'indexation start:end le end n'est pas inclus)  
   
   #print("\n numero du CP dans Raystation est : ",i_cp_+1)            # On a effectué des tests pour verifier que l'indice des lames est correct  
   #print("\n indice de la premiere lame dans le champ =  ",indice_jaw_position_Y1+1)    # On print l'indice calculé par le code et on le compare à ce que l'on voit sur le TPS
   #print("\n indice de la derniere lame dans le champ =  ",indice_jaw_position_Y2)
   
   
   
   Array_LSV_CP_G = np.copy(Array_LSV_CP_G[indice_jaw_position_Y1:indice_jaw_position_Y2])
   Array_LSV_CP_D = np.copy(Array_LSV_CP_D[indice_jaw_position_Y1:indice_jaw_position_Y2])
   N_open_field = Array_LSV_CP_G.size                        ###***!!!***### Calcul du N pour la formule du LSV. N= nombre de lame de gauche le champ de traitement. On a le meme nombre à droite.
   # Calcul de Pos_Max du papier qui est obtenu en faisant max(Pos)-min(Pos)

   Max_Pos_LSV_G_CP=np.amax(Array_LSV_CP_G)               # calcul du max(pos) pour les lames de gauche et dans le champ de traitement
   Max_Pos_LSV_D_CP=np.amax(Array_LSV_CP_D)               # idem à droite

   Min_Pos_LSV_G_CP=np.amin(Array_LSV_CP_G)               # calcul du min(pos) pour les lames de gauche et dans le champ de traitement
   Min_Pos_LSV_D_CP=np.amin(Array_LSV_CP_D)               # idem à droite

   Pos_Max_LSV_CP_G = np.copy(Max_Pos_LSV_G_CP - Min_Pos_LSV_G_CP)           ###***!!!***### Calcul du Pos_Max de gauche pour la formule du LSV. Pos_max est l'écart maximum entre la position maximale et minimale occupée par les lames dans un CP
   Pos_Max_LSV_CP_D = np.copy(Max_Pos_LSV_D_CP - Min_Pos_LSV_D_CP)           ###***!!!***### Calcul du Pos_Max de droite pour la formule du LSV.
   
    
   ######Pos_Max_LSV_CP_G_D = np.copy(Max_Pos_LSV_D_CP-Min_Pos_LSV_G_CP)
   ######Pos_Max_LSV_CP_G = Pos_Max_LSV_CP_G_D
   ######Pos_Max_LSV_CP_D = Pos_Max_LSV_CP_G_D   
   
   
   # Reshape des Array pour le calcul du LSV afin d'avoir l'element n et l'element (n+1)
   
   Array_LSV_CP_G_calc_1 = np.copy(Array_LSV_CP_G[:-1])            # Cela correspond à la position "n" de gauche. Le dernier element est supprimé. On a une array avec (N-1) element
   Array_LSV_CP_D_calc_1 = np.copy(Array_LSV_CP_D[:-1])            

   Array_LSV_CP_G_calc_2 = np.copy(Array_LSV_CP_G[1:])             # Cela correspond à la position "n+1" de gauche. Le premier element est supprimé. On a une array avec (N-1) element
   Array_LSV_CP_D_calc_2 = np.copy(Array_LSV_CP_D[1:])


   LSV_CP_G_diff_Array = np.copy(Array_LSV_CP_G_calc_1 - Array_LSV_CP_G_calc_2)        # On fait la difference entre la position "n" et "n+1"
   LSV_CP_G_diff_Array = np.absolute(LSV_CP_G_diff_Array)            # On prend la valeur absolue de la difference 

   LSV_CP_G_up = (N_open_field-1)* Pos_Max_LSV_CP_G - np.sum(LSV_CP_G_diff_Array)       # Calcul du numerateur du LSV gauche 
   LSV_CP_G_down = (N_open_field-1)* Pos_Max_LSV_CP_G             # Calcul du dénominateur du LSV gauche
   LSV_CP_G = LSV_CP_G_up / LSV_CP_G_down                # Calcul du LSV gauche


   LSV_CP_D_diff_Array = np.copy(Array_LSV_CP_D_calc_1 - Array_LSV_CP_D_calc_2)        # idem à droite
   LSV_CP_D_diff_Array = np.absolute(LSV_CP_D_diff_Array)

   LSV_CP_D_up = (N_open_field-1)* Pos_Max_LSV_CP_D - np.sum(LSV_CP_D_diff_Array) 
   LSV_CP_D_down = (N_open_field-1)* Pos_Max_LSV_CP_D
   LSV_CP_D = LSV_CP_D_up / LSV_CP_D_down

   LSV_CP = LSV_CP_G * LSV_CP_D                  ###***!!!***###   Calcul du LSV du CP   ###***!!!***### 
   list_CP_Array_LSV.append(LSV_CP)
   

  list_CP_Array_LSV = np.array(list_CP_Array_LSV)
  list_Arc_Array_LSV.append(list_CP_Array_LSV)
  
 list_Arc_Array_LSV = np.array(list_Arc_Array_LSV)

 #print(" \n ---------------------- \n ")
 #print(" list_Arc_Array_LSV = \n ",list_Arc_Array_LSV)
 #print(" Max_Pos_LSV_G_CP = \n ",Max_Pos_LSV_G_CP)
 #print(" Max_Pos_LSV_D_CP = \n ",Max_Pos_LSV_D_CP)
 #print("N_open_field=",N_open_field)
 #print("LSV_CP_D_diff_Array shape =",LSV_CP_D_diff_Array.shape)
 #print(" \n ---------------------- \n ")
 #print(" shape = ",list_Arc_Array_LSV.shape)
 #print(" dim = ",list_Arc_Array_LSV.ndim)


 #######################################
 ## Calcul du AAV
 ##
 ######################################

 ## Calcul du Maximum aperture AAV_max

 list_Arc_Array_AAV_max=[]

 for i_arc in range(Nombre_Arcs):
  for i_cp_ in range(List_Nombre_CP[i_arc]):
   temp_Array_AAV_CP_G = np.copy(list_Arc_Array[i_arc,i_cp_,0])
   temp_Array_AAV_CP_D = np.copy(list_Arc_Array[i_arc,i_cp_,1]) 
   #temp_Array_AAV_CP_diff = np.copy(Array_AAV_CP_D - Array_AAV_CP_G)         # La position de droite est toujours superieure à celle de gauche Donc les elements de temp_Array_AAV_CP_diff sont positifs. 

   indice_jaw_position_Y1=list_indice_jaw_Y[i_arc,i_cp_,0]                 # Indice de la lame dans le champ de traitement (limite du bas du "open field"). Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y1 + 1    (car l'indexation dans python commence à 0 et le TPS on commence à 1)
   indice_jaw_position_Y2=list_indice_jaw_Y[i_arc,i_cp_,1]                 # Indice de la première lame en dehors du champ. Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y2      (car dans l'indexation start:end le end n'est pas inclus)  

   temp_Array_AAV_CP_G[:indice_jaw_position_Y1] = 0
   temp_Array_AAV_CP_G[indice_jaw_position_Y2:] = 0
  
   temp_Array_AAV_CP_D[:indice_jaw_position_Y1] = 0
   temp_Array_AAV_CP_D[indice_jaw_position_Y2:] = 0   
  
   if i_cp_ == 0:                      # pour initialiser le Array_AAV_CP_diff
    Array_AAV_min_CP_G = np.copy(list_Arc_Array[i_arc,i_cp_,0])
    Array_AAV_max_CP_D = np.copy(list_Arc_Array[i_arc,i_cp_,1])
    indice_jaw_position_Y1_min_AAV = indice_jaw_position_Y1
    indice_jaw_position_Y2_max_AAV = indice_jaw_position_Y2
    
   for i_indice in range(Array_AAV_min_CP_G.shape[0]):                 # nombre de lames: 80
    if Array_AAV_min_CP_G[i_indice]>temp_Array_AAV_CP_G[i_indice]:       #  
     Array_AAV_min_CP_G[i_indice]=temp_Array_AAV_CP_G[i_indice]       # On cherche la plus petite position de la lame à gauche (lame la plus à gauche donc min_pos_gauche )
     
    if Array_AAV_max_CP_D[i_indice]<temp_Array_AAV_CP_D[i_indice]:       # 
     Array_AAV_max_CP_D[i_indice]=temp_Array_AAV_CP_D[i_indice]       # On cherche la plus grande position de la lame à droite (lame la plus à droite donc max_pos_droite )   
   
   if indice_jaw_position_Y1_min_AAV>indice_jaw_position_Y2:
    indice_jaw_position_Y1_min_AAV = indice_jaw_position_Y2

   if indice_jaw_position_Y2_max_AAV<indice_jaw_position_Y1:
    indice_jaw_position_Y2_max_AAV = indice_jaw_position_Y1   

  
  Array_AAV_min_CP_G[:indice_jaw_position_Y1_min_AAV] = 0
  Array_AAV_min_CP_G[indice_jaw_position_Y2_max_AAV:] = 0
  
  Array_AAV_max_CP_D[:indice_jaw_position_Y1_min_AAV] = 0
  Array_AAV_max_CP_D[indice_jaw_position_Y2_max_AAV:] = 0
 
  Array_AAV_CP_diff = np.copy(Array_AAV_max_CP_D - Array_AAV_min_CP_G)
  AAV_max_arc = np.sum(Array_AAV_CP_diff)   
  list_Arc_Array_AAV_max.append(AAV_max_arc)
  
 list_Arc_Array_AAV_max = np.array(list_Arc_Array_AAV_max)            ###***!!!***###   Calcul du AAV_max de l'arc   ###***!!!***### 

 #print("---------------------------------------------------------------------")
 #print("\n list_Arc_Array_AAV_max \n", list_Arc_Array_AAV_max)
 #print("---------------------------------------------------------------------")

 ## Calcul du   AAV

 list_Arc_Array_AAV=[]
 for i_arc in range(Nombre_Arcs):
  list_CP_Array_AAV=[]
  for i_cp_ in range(List_Nombre_CP[i_arc]):
   Array_AAV_CP_G_calc1 = np.copy(list_Arc_Array[i_arc,i_cp_,0])
   Array_AAV_CP_D_calc2 = np.copy(list_Arc_Array[i_arc,i_cp_,1]) 
   Array_AAV_CP_diff_calc = np.copy(Array_AAV_CP_D_calc2 - Array_AAV_CP_G_calc1)     # Calcul de la surface entre lame de droite et lame de gauche pour un CP 

   indice_jaw_position_Y1=list_indice_jaw_Y[i_arc,i_cp_,0]                 # Indice de la lame dans le champ de traitement (limite du bas du "open field"). Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y1 + 1    (car l'indexation dans python commence à 0 et le TPS on commence à 1)
   indice_jaw_position_Y2=list_indice_jaw_Y[i_arc,i_cp_,1]                # Indice de la première lame en dehors du champ. Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y2      (car dans l'indexation start:end le end n'est pas inclus)  
   
   Array_AAV_CP_diff_calc[:indice_jaw_position_Y1] = 0
   Array_AAV_CP_diff_calc[indice_jaw_position_Y2:] = 0
   
   AAV_CP_up = np.sum(Array_AAV_CP_diff_calc)              # Calcul de la surface totale délimitée par les lames

   AAV_CP = AAV_CP_up/list_Arc_Array_AAV_max[i_arc]                 ###***!!!***###   Calcul du AAV du CP   ###***!!!***###        
   
   list_CP_Array_AAV.append(AAV_CP)
   
  list_CP_Array_AAV=np.array(list_CP_Array_AAV)
  list_Arc_Array_AAV.append(list_CP_Array_AAV)

 list_Arc_Array_AAV=np.array(list_Arc_Array_AAV)
 #print("List_arc_aav",list_Arc_Array_AAV)


 #######################################
 ## Calcul du MCSv
 ##
 ######################################


 list_Arc_MCSv=[]
 for i_arc in range(Nombre_Arcs):
  indice_MCSv=0
  for i_cp_ in range( (List_Nombre_CP[i_arc]-1) ):              # on ne prend pas le dernier element car on fait la somme de l'element "n" et "n+1"
   
   temp_MCSv_AAV = ( list_Arc_Array_AAV[i_arc,i_cp_]+ list_Arc_Array_AAV[i_arc,i_cp_+1] )/2   # Calcul du AAV entre le CP i et le CP i+1
   temp_MCSv_LSV = ( list_Arc_Array_LSV[i_arc,i_cp_]+ list_Arc_Array_LSV[i_arc,i_cp_+1] )/2    # idem por le LSV
   
   temp_MCSv = temp_MCSv_AAV * temp_MCSv_LSV * list_Arc_Weight[i_arc,i_cp_]       # Calcul du MSCv du CP. il faut prendre en compte le poids de chaque CP
   
   indice_MCSv = indice_MCSv + temp_MCSv                ###***!!!***###   Calcul du MCSv de l'arc   ###***!!!***### 
   
  list_Arc_MCSv.append(indice_MCSv)


 Final_MCSv = 0
 for i_arc in range(Nombre_Arcs):
  Final_MCSv = Final_MCSv + list_Arc_MCSv[i_arc] * List_Nombre_MU[i_arc]/Total_MU       ###***!!!***###   Calcul du MCSv du plan   ###***!!!***### 
 
 #List_MCSv_Plans.append(Final_MCSv)

 print(" \n ")
 print("  Indice de modulation  ")
 #list_Arc_MCSv_print = [round(i_1,4) for i_1 in list_Arc_MCSv]
 print("  list_Arc_MCSv:  ", list_Arc_MCSv)
 #print("\n  ")
 print("  MCSv = ", round(Final_MCSv,4))
 #print(" ")
 
 
 
 
 
 ##################################
 ## Calcul du LT 
 ## On prend en compte que le champ de traitement (LT2)
 ##################################
 
 list_Arc_Array_LT=[]

 for i_arc in range(Nombre_Arcs):
 
  #Array_LT_CP_Diff_G = np.zeros(80)
  #Array_LT_CP_Diff_D = np.zeros(80)
  
  Array_LT_CP_Diff_G = np.zeros(60)      #modification CB : nombre de lames 60 pour un varian
  Array_LT_CP_Diff_D = np.zeros(60)


  for i_cp_ in range(List_Nombre_CP[i_arc]-1):

   Array_LT_CP_G_0 = np.copy(list_Arc_Array[i_arc,i_cp_,0])            # array avec les positions des lames gauches de l'arc "i_arc" et du CP "i_cp_"   (on boucle sur tout les CP et tout les arcs) 
   Array_LT_CP_G_1 = np.copy(list_Arc_Array[i_arc,i_cp_+1,0])

   Array_LT_CP_D_0 = np.copy(list_Arc_Array[i_arc,i_cp_,1])            # array avec les positions des lames droites de l'arc "i_arc" et du CP "i_cp_"   (on boucle sur tout les CP et tout les arcs)
   Array_LT_CP_D_1 = np.copy(list_Arc_Array[i_arc,i_cp_+1,1])


   indice_jaw_position_Y1_0 = list_indice_jaw_Y[i_arc,i_cp_,0]
   indice_jaw_position_Y2_0 = list_indice_jaw_Y[i_arc,i_cp_,1]
   indice_jaw_position_Y1_1 = list_indice_jaw_Y[i_arc,i_cp_+1,0]
   indice_jaw_position_Y2_1 = list_indice_jaw_Y[i_arc,i_cp_+1,1]

   indice_jaw_position_Y1 = max(indice_jaw_position_Y1_0,indice_jaw_position_Y1_1)
   indice_jaw_position_Y2 = min(indice_jaw_position_Y2_0,indice_jaw_position_Y2_1) 

   Array_LT_CP_G_0[:indice_jaw_position_Y1] = 0 ; Array_LT_CP_D_0[:indice_jaw_position_Y1] = 0
   Array_LT_CP_G_0[indice_jaw_position_Y2:] = 0 ; Array_LT_CP_D_0[indice_jaw_position_Y2:] = 0

   Array_LT_CP_G_1[:indice_jaw_position_Y1] = 0 ; Array_LT_CP_D_1[:indice_jaw_position_Y1] = 0 
   Array_LT_CP_G_1[indice_jaw_position_Y2:] = 0 ; Array_LT_CP_D_1[indice_jaw_position_Y2:] = 0




   Array_LT_CP_Diff_G_temp = np.copy(Array_LT_CP_G_1-Array_LT_CP_G_0)
   Array_LT_CP_Diff_G_temp = np.absolute(Array_LT_CP_Diff_G_temp)

   Array_LT_CP_Diff_D_temp = np.copy(Array_LT_CP_D_1-Array_LT_CP_D_0)
   Array_LT_CP_Diff_D_temp = np.absolute(Array_LT_CP_Diff_D_temp)

   Array_LT_CP_Diff_G = Array_LT_CP_Diff_G + Array_LT_CP_Diff_G_temp
   Array_LT_CP_Diff_D = Array_LT_CP_Diff_D + Array_LT_CP_Diff_D_temp
    
  #print("\n Array_LT_CP_Diff_G \n",Array_LT_CP_Diff_G)
  #print("\n Array_LT_CP_Diff_D \n",Array_LT_CP_Diff_D)
  
  Array_LT_CP_Diff_G =np.where(Array_LT_CP_Diff_G != 0 , Array_LT_CP_Diff_G , np.nan ) 
  Array_LT_CP_Diff_D =np.where(Array_LT_CP_Diff_D != 0 , Array_LT_CP_Diff_D , np.nan ) 

  #print("\n Array_LT_CP_Diff_G \n",Array_LT_CP_Diff_G)
  #print("\n Array_LT_CP_Diff_D \n",Array_LT_CP_Diff_D)

  LT_mean_G = np.nanmean(Array_LT_CP_Diff_G) 
  LT_mean_D = np.nanmean(Array_LT_CP_Diff_D) 
  
  LT_mean_G_D = (LT_mean_G + LT_mean_D)/2
  
  list_Arc_Array_LT.append(LT_mean_G_D)
  
  #print("LT_mean_G = ",LT_mean_G)
  #print("LT_mean_D = ",LT_mean_D)
  #print("LT_mean_G_D = ",LT_mean_G_D)
  
 
 
 LT_mean_Plan = 0
 for i_arc in range(Nombre_Arcs):
  LT_mean_Plan = LT_mean_Plan + list_Arc_Array_LT[i_arc]/Nombre_Arcs
 
 #List_LT_Plans.append(LT_mean_Plan)
 
 print("  LT du champ de traitement")
 print("  LT List (cm)  = ",list_Arc_Array_LT)
 print("  LT Moyen du plan(cm) = ", round(LT_mean_Plan,4))
 
 
 
 ##################################
 ## Calcul du LTMCS 
 ## On prend en compte que le champ de traitement (LTMCS 2)
 ##################################
 
 the_LTMCS = Final_MCSv * (200-LT_mean_Plan)/200
 #List_LTMCS_Plans.append(the_LTMCS)
 
 print("  LTMCS = ", round(the_LTMCS,4))
 print(" \n ") 
 
 ##################################
 ## Calcul du SAS(1cm)
 ## Ratio Of Small Segments (<1cm)
 ##################################
 list_sum_1_weight=[]
 list_sum_2_weight=[]
 
 for i_arc in range(Nombre_Arcs):
  Sum_Number_of_seg_SAS_weight = 0
  Sum_Number_of_all_seg_weight = 0
  for i_cp_ in range(List_Nombre_CP[i_arc]):
  
   Array_SAS_CP_G = np.copy(list_Arc_Array[i_arc,i_cp_,0])
   Array_SAS_CP_D = np.copy(list_Arc_Array[i_arc,i_cp_,1])
   
   indice_jaw_position_Y1=list_indice_jaw_Y[i_arc,i_cp_,0]               # Indice de la lame dans le champ de traitement (limite du bas du "open field"). Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y1 + 1    (car l'indexation dans python commence à 0 et le TPS on commence à 1)
   indice_jaw_position_Y2=list_indice_jaw_Y[i_arc,i_cp_,1]               # Indice de la première lame en dehors du champ. Le numéro de lame correspondant dans le TPS est = indice_jaw_position_Y2      (car dans l'indexation start:end le end n'est pas inclus)  
   
   Array_SAS_CP_G = Array_SAS_CP_G[indice_jaw_position_Y1:indice_jaw_position_Y2]
   Array_SAS_CP_D = Array_SAS_CP_D[indice_jaw_position_Y1:indice_jaw_position_Y2]
   
   Array_SAS_diff = np.copy(Array_SAS_CP_D - Array_SAS_CP_G)
   Array_SAS_ = np.where(Array_SAS_diff<=1,1,0)
   Number_of_seg_SAS = np.sum(Array_SAS_)
   Number_of_all_seg = Array_SAS_.size
     
   Sum_Number_of_seg_SAS_weight = Sum_Number_of_seg_SAS_weight + Number_of_seg_SAS * list_Arc_Weight[i_arc,i_cp_]
   Sum_Number_of_all_seg_weight = Sum_Number_of_all_seg_weight + Number_of_all_seg * list_Arc_Weight[i_arc,i_cp_]
      
  list_sum_1_weight.append(Sum_Number_of_seg_SAS_weight)
  list_sum_2_weight.append(Sum_Number_of_all_seg_weight)
  
 total_Seg_SAS_weight = 0
 total_Seg_weight = 0
 for i_arc in range(Nombre_Arcs):
  total_Seg_SAS_weight = total_Seg_SAS_weight + list_sum_1_weight[i_arc] * List_Nombre_MU[i_arc]/Total_MU
  total_Seg_weight = total_Seg_weight + list_sum_2_weight[i_arc] * List_Nombre_MU[i_arc]/Total_MU
  
 The_SAS_final_weight = total_Seg_SAS_weight/total_Seg_weight
 #List_SAS_weight_Plans.append(The_SAS_final_weight) 

 print("  Ratio du Nombre de Segments inf à 1cm SAS(1cm) avec pondération en MU \n")
 #print("\n  Nombre de petits segments pondéré en UM =  ", total_Seg_SAS_weight)
 #print("\n  Nombre total de segments pondéré en UM =  ", total_Seg_weight)
 print("\n  SAS(1cm) =  ", round(The_SAS_final_weight,4))
 print(" \n ") 
 return [round(The_SAS_final_weight,4),round(Final_MCSv,4),round(LT_mean_Plan,4),round(the_LTMCS,4)]
 



################## Modif ACD ################

### Set case and plan variable ###
patient = get_current('Patient')
case = get_current("Case")
plan = get_current('Plan')

date = datetime.datetime.now()

exportDate = str(date.day) + "/" + str(date.month) + "/" + str(date.year)

### Get patient's informations ###
namePatient = patient.Name
nameToSplit = namePatient.split("^")
firstNamePatient = nameToSplit[1]
namePatient = nameToSplit[0]

### Fill ListToExport with patient's information ###
patientInfo = [namePatient, firstNamePatient]
patientID = patient.PatientID


indices = fonction_calculate_MCSv_LT(patient,case,plan)

ListToExport = str(exportDate) + "\t" + str(patientInfo) + "\t" + str(patientID) + "\t" + str(indices[0]) + "\t" + str(indices[1]) + "\t" + str(indices[2]) + "\t" + str(indices[3])+ "\n"



################## UI ACD ################

root = Tk()

root.title("Copy and past the values below:")

v = StringVar()
textbox1 = Entry(root, textvariable=v)
textbox1.grid(column=0, row=4)
textbox1.config(width=100)
textbox1.insert(END, ListToExport)

def Quit():
 root.destroy()

butt1 = Button(root, text = 'Quit', command = Quit)
butt1.grid(column=0, row=6)
butt1.config(width=100)

root.mainloop()

################## UI ACD ################
