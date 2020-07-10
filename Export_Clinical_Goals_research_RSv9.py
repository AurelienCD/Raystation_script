# -*- coding: utf-8 -*-

# Export the clinical goals to excel or csv files
# Clinical goals reading is based on "print goal" function from Mark Geurts; Github : https://github.com/wrssc/ray_scripts/blob/master/library/Goals.py
# Author : Aurélien Corroyer-Dulmont
# Version : 24 june 2020

# Update xx/xx/2020 : 



########################### TO DO #############################################################################################################################################
### xxx

import os
import codecs
import csv
from tkinter import *
from tkinter import ttk
from System.Windows.Forms import MessageBox
from connect import *
import datetime


def Export_Clinical_Goals(goal):
	clinicalGoal = []

	clinicalGoal.append(goal.ForRegionOfInterest.Name)

	if goal.PlanningGoal.Type == 'VolumeAtDose':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()*100,2))
		clinicalGoal.append('%')

	elif goal.PlanningGoal.Type == 'DoseAtVolume':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()/100,2))
		clinicalGoal.append('Gy')
	else:
		clinicalGoal.append('invalide parameters')
	
	#value = round(goal.GetClinicalGoalValue()/100,2)

	#clinicalGoal.append(round(goal.GetClinicalGoalValue()/100,2))
	
	clinicalGoal.append(goal.EvaluateClinicalGoal())

	return clinicalGoal


def ExportToCSV(exportDate, patientInfo, patientID, planName, listToExport):
	savepath = "Q:/Solène/0_SEIN_RAYSTATION/Raystation_Clinical_Goals_Research.csv"
	filesave = open(savepath, 'a', encoding='Latin-1')
	for elm in listToExport:
		filesave.write(str(exportDate) + ";" + str(patientInfo) + ";" + str(patientID) + ";" + str(planName) + ";" + str(elm[0]) + ";" + str(elm[1]) + ";" + str(elm[2]) + ";" + str(elm[3]) + ";" + str(elm[4]) + ";" + "\n")
	filesave.write("\n")
	filesave.close()
	MessageBox.Show('Clinical Goals Export completed')



### Set case and plan variable ###
patient = get_current('Patient')
case = get_current("Case")
plan = get_current('Plan')
planName = plan.Name

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

### Get the clinical goals ###
Goals = plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions ### goals sera ainsi une liste avec autant d'entrée que de clinical goals (par ROI ou autre) ###


listToExport = []
for elm in Goals:
 listToExport.append(Export_Clinical_Goals(elm))

ExportToCSV(exportDate, patientInfo, patientID, planName, listToExport)
