# -*- coding: utf-8 -*-

# Export the clinical goals to a box to copy to an excel or csv files
# Clinical goals reading is based on "print goal" function from Mark Geurts; Github : https://github.com/wrssc/ray_scripts/blob/master/library/Goals.py
# Author : Aurélien Corroyer-Dulmont
# Version : 24 june 2020

# Update xx/xx/2020 : 


import os, csv, codecs, sys, System, clr, random, wpf
from math import *
from connect import *
clr.AddReferenceByName("PresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReferenceByName("PresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReference("System.Windows.Forms")
from System.Windows import *
from System.Windows.Forms import MessageBox
import datetime


def Export_Clinical_and_Evaluation_Doses(goal):

	clinicalGoal = []

	clinicalGoal.append(plan.Name)
	clinicalGoal.append(goal.ForRegionOfInterest.Name)

	if goal.PlanningGoal.Type == 'VolumeAtDose':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()*100,2))
		clinicalGoal.append('%')

	elif goal.PlanningGoal.Type == 'AbsoluteVolumeAtDose':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}cc at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2), round(goal.PlanningGoal.ParameterValue,2) / 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}cc at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2), round(goal.PlanningGoal.ParameterValue,2) / 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue(),2))
		clinicalGoal.append('cc')

	elif goal.PlanningGoal.Type == 'DoseAtVolume':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()/100,2))
		clinicalGoal.append('Gy')

	elif goal.PlanningGoal.Type == 'DoseAtAbsoluteVolume':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}Gy at {}cc'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2)))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}Gy at {}cc'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2)))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()/100,2))
		clinicalGoal.append('Gy')

	elif goal.PlanningGoal.Type == 'AverageDose':
		if goal.PlanningGoal.GoalCriteria == 'AtMost':
			clinicalGoal.append('At most {}Gy average dose'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100))
		elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
			clinicalGoal.append('At least {}Gy average dose'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100))
		clinicalGoal.append(round(goal.GetClinicalGoalValue()/100,2))
		clinicalGoal.append('Gy')
	else:
		clinicalGoal.append('invalide parameters')
	
	clinicalGoal.append(goal.EvaluateClinicalGoal())


	### Analyse if Evaluation Dose exist ###
	for elm in case.Examinations:
		if len(elm.Name) < 20:
			evaluationDoseExist = False
		else:
			evaluationDoseExist = True
			CTmodifiedName = elm.Name
	

	### Get the clinical goal for the evaluation dose ###
	if evaluationDoseExist == True:

		clinicalGoal.append(CTmodifiedName)
		clinicalGoal.append(goal.ForRegionOfInterest.Name)

		for elm in case.TreatmentDelivery.FractionEvaluations[0].DoseOnExaminations:
			if elm.OnExamination.Name == CTmodifiedName:
				evaluationDose = elm.DoseEvaluations[0]

		beam_set = get_current("BeamSet")
		fractions = beam_set.FractionationPattern.NumberOfFractions

		if goal.PlanningGoal.Type == 'VolumeAtDose':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				clinicalGoal.append('At most {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				clinicalGoal.append('At least {}% at {}Gy'.format(round(goal.PlanningGoal.AcceptanceLevel,2) * 100, round(goal.PlanningGoal.ParameterValue,2) / 100))
			clinicalGoal.append(round(evaluationDose.GetRelativeVolumeAtDoseValues(RoiName = goal.ForRegionOfInterest.Name, DoseValues = [goal.PlanningGoal.ParameterValue/fractions])[0]*100,2))
			clinicalGoal.append('%')

		elif goal.PlanningGoal.Type == 'AbsoluteVolumeAtDose':
			clinicalGoal.append("Not Available for Evaluation Dose")
			clinicalGoal.append("NaN")
			clinicalGoal.append("NaN")


		elif goal.PlanningGoal.Type == 'DoseAtVolume':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				clinicalGoal.append('At most {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				clinicalGoal.append('At least {}Gy at {}%'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100, round(goal.PlanningGoal.ParameterValue,2) * 100))
			clinicalGoal.append(round(evaluationDose.GetDoseAtRelativeVolumes(RoiName = goal.ForRegionOfInterest.Name, RelativeVolumes=[goal.PlanningGoal.ParameterValue])[0]*fractions/100,2))
			clinicalGoal.append('Gy')


		elif goal.PlanningGoal.Type == 'DoseAtAbsoluteVolume':
			clinicalGoal.append("Not Available for Evaluation Dose")
			clinicalGoal.append("NaN")
			clinicalGoal.append("NaN")


		elif goal.PlanningGoal.Type == 'AverageDose':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				clinicalGoal.append('At most {}Gy average dose'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100))
			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				clinicalGoal.append('At least {}Gy average dose'.format(round(goal.PlanningGoal.AcceptanceLevel,2) / 100))
			clinicalGoal.append(round(evaluationDose.GetDoseStatistic(RoiName = goal.ForRegionOfInterest.Name, DoseType = "Average")*fractions/100,2))
			clinicalGoal.append('Gy')

		else:
			clinicalGoal.append('invalide parameters')
		
		### Analyse the result of the evaluation dose ###
		if goal.PlanningGoal.Type == 'VolumeAtDose':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				if evaluationDose.GetRelativeVolumeAtDoseValues(RoiName = goal.ForRegionOfInterest.Name, DoseValues = [goal.PlanningGoal.ParameterValue/fractions])[0] < goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')
					
			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				if evaluationDose.GetRelativeVolumeAtDoseValues(RoiName = goal.ForRegionOfInterest.Name, DoseValues = [goal.PlanningGoal.ParameterValue/fractions])[0] > goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')

		elif goal.PlanningGoal.Type == 'AbsoluteVolumeAtDose':
			clinicalGoal.append("Not Available for Evaluation Dose")


		elif goal.PlanningGoal.Type == 'DoseAtVolume':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				if evaluationDose.GetDoseAtRelativeVolumes(RoiName = goal.ForRegionOfInterest.Name, RelativeVolumes=[goal.PlanningGoal.ParameterValue])[0]*fractions < goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')
					
			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				if evaluationDose.GetDoseAtRelativeVolumes(RoiName = goal.ForRegionOfInterest.Name, RelativeVolumes=[goal.PlanningGoal.ParameterValue])[0]*fractions > goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')


		elif goal.PlanningGoal.Type == 'DoseAtAbsoluteVolume':
			clinicalGoal.append("Not Available for Evaluation Dose")


		elif goal.PlanningGoal.Type == 'AverageDose':
			if goal.PlanningGoal.GoalCriteria == 'AtMost':
				if evaluationDose.GetDoseStatistic(RoiName = goal.ForRegionOfInterest.Name, DoseType = "Average")*fractions < goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')

			elif goal.PlanningGoal.GoalCriteria == 'AtLeast':
				if evaluationDose.GetDoseStatistic(RoiName = goal.ForRegionOfInterest.Name, DoseType = "Average")*fractions > goal.PlanningGoal.AcceptanceLevel:
					clinicalGoal.append('True')
				else:
					clinicalGoal.append('False')

	return clinicalGoal


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

### Get the clinical goals ###
Goals = plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions ### goals sera ainsi une liste avec autant d'entrée que de clinical goals (par ROI ou autre) ###


listToExport = []
for elm in Goals:
 listToExport.append(Export_Clinical_and_Evaluation_Doses(elm))

FinalListToExport = ""


### Analyse if Evaluation Dose exist ###
for elm in case.Examinations:
	if len(elm.Name) < 20:
		evaluationDoseExist = False
	else:
		evaluationDoseExist = True

### Get the clinical goal for the evaluation dose ###
if evaluationDoseExist == False:
	for elm in listToExport:
		FinalListToExport += str(exportDate) + "\t" + str(patientInfo) + "\t" + str(patientID) + "\t" + str(elm[0]) + "\t" + str(elm[1]) + "\t" + str(elm[2]) + "\t" + str(elm[3]) + "\t" + str(elm[4]) + "\t" + str(elm[5]) + "\t" + "\n"
else:
	for elm in listToExport:
		FinalListToExport += str(exportDate) + "\t" + str(patientInfo) + "\t" + str(patientID) + "\t" + str(elm[0]) + "\t" + str(elm[1]) + "\t" + str(elm[2]) + "\t" + str(elm[3]) + "\t" + str(elm[4]) + "\t" + str(elm[5]) + "\t" + "\n" + str(exportDate) + "\t" + str(patientInfo) + "\t" + str(patientID) + "\t" + str(elm[6]) + "\t" + str(elm[7]) + "\t" + str(elm[8]) + "\t" + str(elm[9]) + "\t" + str(elm[10]) + "\t" + str(elm[11]) + "\t" + "\n"


# Initialization Constants
Window = System.Windows.Window
Application = System.Windows.Application
Button = System.Windows.Controls.Button
StackPanel = System.Windows.Controls.StackPanel
Label = System.Windows.Controls.Label
Thickness = System.Windows.Thickness
DropShadowBitmapEffect = System.Windows.Media.Effects.DropShadowBitmapEffect
TextBox = System.Windows.Controls.TextBox

# Create window
my_window = Window()
my_window.Title = 'Copy and past the clinical goals below:'
my_window.Width = 900
my_window.Height = 350

# Create StackPanel to Layout UI elements 
my_stack = StackPanel()
my_stack.Margin = Thickness(15)
my_window.Content = my_stack

my_textbox = TextBox()
my_textbox.Text = str(FinalListToExport)
my_stack.Children.Add (my_textbox)

my_app = Application()
my_app.Run (my_window)












