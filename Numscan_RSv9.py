# -*- coding: utf-8 -*-

# Date: 25/02/2020
# The aim of this script is to find the ID and the number of the serie of CT or MRI, to show it, and to rename the imaging with these informations
# This script was based on "Numscan" script from Raysearchlab
# For more information : Aurélien Corroyer-Dulmont 5768 or a.corroyer-dulmont@baclesse.unicancer.fr


### UPDATE:
# 11/03/2020: acquisition date was added to the name of the CT
# 04/05/2020: information was added about initial and new CT

import wpf, os, sys, System, clr, random
from System import Windows
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from math import *
from connect import *



case=get_current('Case')

try:
 StudyID_CT1=case.Examinations._CT_1.GetStoredDicomTagValueForVerification(Group=32,Element=16)
 dateTime_CT1=str(case.Examinations._CT_1.GetExaminationDateTime())
 #acquisitionDate_CT1=dateTime_CT1[:10] #to select only the date not the time; for RS version 9A, needs to be dateTime_CT1[:10] and not 9
 Serie_CT1=case.Examinations._CT_1.GetStoredDicomTagValueForVerification(Group=32,Element=17)
 NumStudyID_CT1=StudyID_CT1.TryGetValue('Study ID')[1]
 NumSerie_CT1=Serie_CT1.TryGetValue('Series Number')[1]
 case.Examinations._CT_1.Name = 'CT_ID: ' + NumStudyID_CT1 + '_' + NumSerie_CT1 + '_Initial_Date: ' + dateTime_CT1
 CT1_Exist = True
except:
 print("There is no CT named ""CT_1"" with that patient ")
 CT1_Exist = False

try:
 StudyID_CT2=case.Examinations._CT_2_reprise.GetStoredDicomTagValueForVerification(Group=32,Element=16)
 dateTime_CT2_reprise=str(case.Examinations._CT_2_reprise.GetExaminationDateTime())
 #acquisitionDate_CT2_reprise=dateTime_CT2_reprise[:10] #to select only the date not the time
 Serie_CT2=case.Examinations._CT_2_reprise.GetStoredDicomTagValueForVerification(Group=32,Element=17)
 NumStudyID_CT2=StudyID_CT2.TryGetValue('Study ID')[1]
 NumSerie_CT2=Serie_CT2.TryGetValue('Series Number')[1]
 case.Examinations._CT_2_reprise.Name = 'CT_ID: ' + NumStudyID_CT2 + '_' + NumSerie_CT2 + '_New_Date: ' + dateTime_CT2_reprise
 CT2_reprise_Exist = True
except:
 print("There is no CT named ""CT_2_reprise"" with that patient ")
 CT2_reprise_Exist = False

try:
 StudyID_CT2=case.Examinations._CT_2.GetStoredDicomTagValueForVerification(Group=32,Element=16)
 dateTime_CT2=str(case.Examinations._CT_2.GetExaminationDateTime())
 #acquisitionDate_CT2=dateTime_CT2[:10] #to select only the date not the time
 Serie_CT2=case.Examinations._CT_2.GetStoredDicomTagValueForVerification(Group=32,Element=17)
 NumStudyID_CT2=StudyID_CT2.TryGetValue('Study ID')[1]
 NumSerie_CT2=Serie_CT2.TryGetValue('Series Number')[1]
 case.Examinations._CT_2.Name = 'CT_ID: ' + NumStudyID_CT2 + '_' + NumSerie_CT2 + '_New_Date: ' + dateTime_CT2
 CT2_Exist = True
except:
 print("There is no CT named ""CT_2"" with that patient ")
 CT2_Exist = False


try:
 StudyID_CT3=case.Examinations._CT_3_reprise.GetStoredDicomTagValueForVerification(Group=32,Element=16)
 dateTime_CT3=str(case.Examinations._CT_3.GetExaminationDateTime())
 #acquisitionDate_CT3=dateTime_CT3[:10] #to select only the date not the time
 Serie_CT3=case.Examinations._CT_3_reprise.GetStoredDicomTagValueForVerification(Group=32,Element=17)
 NumStudyID_CT3=StudyID_CT3.TryGetValue('Study ID')[1]
 NumSerie_CT3=Serie_CT3.TryGetValue('Series Number')[1]
 case.Examinations._CT_3_reprise.Name = 'CT_ID: ' + NumStudyID_CT3 + '_' + NumSerie_CT3 + '_New_Date: ' + dateTime_CT3
 CT1_Exist = True
except:
 print("There is no CT named ""CT_3"" with that patient ")
 CT3_Exist = False



if CT1_Exist == True:
 MessageBox.Show("CT1, scanner ID:\n" + NumStudyID_CT1 + "\n\nCT1, serie number:\n" + NumSerie_CT1 + "\n\nCT1, acquisition date:\n" + dateTime_CT1 + "\n\nImage name was changed adding these informations")

if CT2_reprise_Exist == True:
 MessageBox.Show("CT2, scanner ID:\n" + NumStudyID_CT2 + "\n\nCT2, serie number:\n" + NumSerie_CT2 + "\n\nCT2_reprise, acquisition date:\n" + dateTime_CT2_reprise + "\n\nImage name was changed adding these informations")

if CT2_Exist == True:
 MessageBox.Show("CT2, scanner ID:\n" + NumStudyID_CT2 + "\n\nCT2, serie number:\n" + NumSerie_CT2 + "\n\nCT2, acquisition date:\n" + dateTime_CT2 + "\n\nImage name was changed adding these informations")

if CT3_Exist == True:
 MessageBox.Show("CT3, scanner ID:\n" + NumStudyID_CT3 + "\n\nCT3, serie number:\n" + NumSerie_CT3 + "\n\nCT3, acquisition date:\n" + dateTime_CT3 + "\n\nImage name was changed adding these informations")

""" this part was decided to be removed on march 2020
### To get information about which scan was used to performed dosimetry
try:
 case=get_current('Case')
 plan=get_current('Plan')
 BS=get_current('BeamSet')
 CT=BS.GetPlanningExamination()
 PlanName = str(plan.Name)
 dateTime_CT=str(CT.GetExaminationDateTime())
 #acquisitionDate_CT=dateTime_CT[:10] #to select only the date not the time
 StudyID=CT.GetStoredDicomTagValueForVerification(Group=32,Element=16)
 Serie=CT.GetStoredDicomTagValueForVerification(Group=32,Element=17)
 NumStudyID=StudyID.TryGetValue('Study ID')[1]
 NumSerie=Serie.TryGetValue('Series Number')[1]
 dosimetry_Exist = True
except:
 print("There is no plan with that patient ")
 dosimetry_Exist = False

if dosimetry_Exist == True:
 MessageBox.Show("Plan: " + str(PlanName) + " was performed on CT ID:\n\n" + NumStudyID + "\n\nserie: " + NumSerie + "\n\nand acquisition date: " + dateTime_CT)"""