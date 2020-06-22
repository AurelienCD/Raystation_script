# -*- coding: utf-8 -*-

# Export the clinical goals to excel or csv files
# Clinical goals reading is based on "print goal" function from Mark Geurts; Github : https://github.com/wrssc/ray_scripts/blob/master/library/Goals.py
# Author : Aurélien Corroyer-Dulmont
# Version : 17 june 2020

# Update xx/xx/2020 : 



########################### TO DO #############################################################################################################################################
### xxx




import pandas as pad 
from openpyxl import load_workbook
import win32com.client
from path import Path
import shutil
import os
import codecs
import csv

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from connect import *



### Get information about which type of localisation is it, obtained from user ###
root = Tk()
root.title('Script_Export_Clinical_Goal')
lbl1 = Label(root, text='Please chose tumour localisation:')
lbl1.grid(column=0,row=0)
lbl1.config(width=50)

combo1_variable = StringVar()
combo1_values = ['Gynéco', 'Thorax 2Gy', 'Thorax 2,4Gy', 'Thorax 3Gy', 'Sein D 2Gy', 'Sein D 2,4Gy', 'Sein G 2Gy', 'Sein G 2,4Gy', \
	'Sein D hypoxG', 'Sein G hypoxG', 'Sein D START', 'Sein G START', 'Canal anal', 'Pelvis + prostate 2Gy', 'Pelvis + prostate 2,4Gy', \
	'Pelvis + prostate 3Gy', 'Crane', 'Crane 40,05Gy', 'Oesophage', 'ORL 2Gy', 'ORL 2,4Gy', 'ORL 3Gy']

combo1 = ttk.Combobox(root, values= combo1_values, textvariable=combo1_variable)
combo1.grid(column=0,row=1)
combo1.current(0)
butt1 = Button(root, text="Quit", command=root.destroy)
butt1 = Button(root, text="Save", command=root.destroy)
butt1.grid(column=0,row=4)
butt1.config(width=50)
root.mainloop()

tumourLocalisation = combo1_variable.get()


### Get information about the starting date of treatment, obtained from user ###
root = Tk()
v = StringVar()
root.title('Script_Export_Clinical_Goal')
lbl1 = Label(root, text='Please add the starting date of treatment (dd/mm/aaaa):')
lbl1.grid(column=0,row=1)
lbl1.config(width=50)
textbox1 = Entry(root, textvariable=v)
textbox1.grid(column=0,row=2)
textbox1.config(width=50)
butt1 = Button(root, text="Save", command=root.destroy)
butt1.grid(column=0,row=3)
butt1.config(width=50)
root.mainloop()
startingTreatmentDay = v.get()






### Get patient's informations ###
namePatient = patient.Name
nameToSplit = namePatient.split("^")
firstNamePatient = nameToSplit[1]
namePatient = nameToSplit[0]


### Set case and plan variable ###
case = get_current("Case")
plan = get_current('Plan') ## à voir en fonction de comment est nommée le plan dans leur utlisation en routine


### Fill ListToExport with patient's information ###
listToExport = [namePatient, firstNamePatient]
listToExport.append(patient.PatientID)



### Get treatment information ###
prescriptionDose = 0
for elm in plan.BeamSets:
	prescriptionDose = prescriptionDose + int(elm.DicomPlanLabel[:2])

approbationDate = plan.BeamSets[0].Review.ReviewTime
approbationDate = str(approbationDate.Day) + "/" + str(approbationDate.Month) + "/" + str(approbationDate.Year)

machineName = plan.BeamSets[0].MachineReference.MachineName

energy = plan.BeamSets[0].Beams[0].BeamQualityId

listToExport.extend((prescriptionDose, approbationDate, startingTreatmentDay, machineName, energy))

###################### Get clinical goals informations ######################

### Get the clinical goals ###
goals = plan.TreatmentCourse.EvaluationSetup.EvaluationFunctions ### goals sera ainsi une liste avec autant d'entrée que de clinical goals (par ROI ou autre) ###

if tumourLocalisation =="Gynéco":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "ParoiRectale":
			if elm.PlanningGoal.ParameterValue == 4000:
				paroiRectaleValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				paroiRectaleValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6000:
				paroiRectaleValuesV60 = (elm.GetClinicalGoalValue())*100
			else:
				paroiRectaleValuesV40 = ""
				paroiRectaleValuesV50 = ""
				paroiRectaleValuesV60 = ""

		elif elm.ForRegionOfInterest.Name == "ParoiVesicale":
			if elm.PlanningGoal.ParameterValue == 4000:
				paroiVesicaleValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6000:
				paroiVesicaleValuesV60 = (elm.GetClinicalGoalValue())*100
			else:
				paroiVesicaleValuesV40 = ""
				paroiVesicaleValuesV60 = ""

		elif elm.ForRegionOfInterest.Name == "CanalAnal":
			if elm.PlanningGoal.ParameterValue == 5600:
				canalAnalValuesV56 = (elm.GetClinicalGoalValue())*100
			else:
				canalAnalValuesV56 = ""

		elif elm.ForRegionOfInterest.Name == "Grêle (cavité péritonéale)":                ################# ATTENTION POUR CELUI CI IL FAUT DEUX PARAMETRES DU V40
			if elm.PlanningGoal.ParameterValue == 4000:									################### un en % l'autre en cc, regarder sur patient type Cindy
				greleValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				greleValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				greleValuesV40 = ""
				greleValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Colon" or elm.ForRegionOfInterest.Name == "Sigmoïde" or elm.ForRegionOfInterest.Name == "colon sigmoïde":
			if elm.PlanningGoal.ParameterValue == 4500:
				colonValuesV45 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				colonValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				colonValuesV45 = ""
				colonValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Moelle":
			if elm.PlanningGoal.ParameterValue == 1000:
				moelleValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				moelleValuesV20 = (elm.GetClinicalGoalValue())*100
			else:
				moelleValuesV10 = ""
				moelleValuesV20 = ""
		
		elif elm.ForRegionOfInterest.Name == "Reins":
			if elm.PlanningGoal.ParameterValue == 1200:
				ReinsValuesV12 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				ReinsValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				ReinsValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				ReinsValuesV12 = ""
				ReinsValuesV20 = ""
				ReinsValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 3000:
				FoieValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV30 = ""

	### Get dose statistics informations ###
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 78Gy", RelativeVolumes = [0.02])
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 78Gy", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 78Gy", DoseType = "Average")

	paroiRectaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiRectale", RelativeVolumes = [0.02])
	paroiRectaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiRectale", DoseType = "Average")

	paroiVesicaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	paroiVesicaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	canalAnalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "CanalAnal", DoseType = "Average")
	canalAnalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "CanalAnal", RelativeVolumes = [0.02])

	greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Grêle (cavité péritonéale)", DoseType = "Average")
	greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Grêle (cavité péritonéale)", RelativeVolumes = [0.02])

	colonD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.05])
	colonAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Colon", DoseType = "Average")
	colonD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.02])

	TFD10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.10])
	TFD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.05])
	TFDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFD", DoseType = "Average")
	TFD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.02])

	TFG10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.10])
	TFG5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.05])
	TFGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFG", DoseType = "Average")
	TFG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.02])

	plexusSacreAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PlexusSacree", DoseType = "Average")
	plexusSacreD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PlexusSacree", RelativeVolumes = [0.02])

	moelleD75 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.75])
	moelleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Moelle", DoseType = "Average")
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	QueueDeChevalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "QueueDeCheval", DoseType = "Average")
	QueueDeChevalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "QueueDeCheval", RelativeVolumes = [0.02])

	ReinsAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	ReinsD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, paroiRectaleValuesV40, paroiRectaleValuesV50, paroiRectaleValuesV60, paroiRectaleAverage, paroiRectaleD2, paroiVesicaleValuesV40, \
		paroiVesicaleValuesV60, paroiVesicaleAverage, paroiVesicaleD2, canalAnalAverage, canalAnalD2, greleValuesV40, greleValuesV50, greleAverage,greleD2, \
		colonValuesV45, colonValuesV50, colonD5, colonAverage, colonD2, TFD10, TFD5, TFDAverage, TFD2, TFG10, TFG5, TFGAverage, TFG2, plexusSacreAverage, \
		plexusSacreD2, moelleValuesV10, moelleValuesV20, moelleD75, moelleAverage, moelleD2, QueueDeChevalAverage, QueueDeChevalD2, ReinsValuesV12, \
		ReinsValuesV20, ReinsValuesV30, ReinsAverage, ReinsD2, FoieValuesV30, FoieAverage, FoieD2))

########################## ATTENTION POUR LA VALEUR DU V40 X2 pour le grêle à remplacer ####################################################################################


elif tumourLocalisation == "Thorax 2Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 3500:
				coeurValuesV35 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4000:
				coeurValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				coeurValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV35 = ""
				coeurValuesV40 = ""
				coeurValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Poumons-PTV":
			if elm.PlanningGoal.ParameterValue == 500:
				poumons_PTVValuesV5= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1300:
				poumons_PTVValuesV13 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				poumons_PTVValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				poumons_PTVValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_PTVValuesV5 = ""
				poumons_PTVValuesV13 = ""
				poumons_PTVValuesV20 = ""
				poumons_PTVValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Oesophage":
			if elm.PlanningGoal.ParameterValue == 4500:
				OesophageValuesV45 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				OesophageValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5500:
				OesophageValuesV55 = (elm.GetClinicalGoalValue())*100
			else:
				OesophageValuesV45 = ""
				OesophageValuesV50 = ""
				OesophageValuesV55 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				ThyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV50 = ""
		
		elif elm.ForRegionOfInterest.Name == "Reins":
			if elm.PlanningGoal.ParameterValue == 1200:
				ReinsValuesV12 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				ReinsValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				ReinsValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				ReinsValuesV12 = ""
				ReinsValuesV20 = ""
				ReinsValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2600:
				FoieValuesV26 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				FoieValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV26 = ""
				FoieValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Estomac":
			if elm.PlanningGoal.ParameterValue == 5400:
				EstomacValuesV54 = (elm.GetClinicalGoalValue())*100
			else:
				EstomacValuesV54 = ""

	### Get dose statistics informations ###
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumons-PTV", DoseType = "Average")
	poumons_PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumons-PTV", RelativeVolumes = [0.02])
	
	OesophageAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oesophage", DoseType = "Average")
	OesophageD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oesophage", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	Plexus_brachial_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_D", DoseType = "Average")
	Plexus_brachial_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_D", RelativeVolumes = [0.02])

	Plexus_brachial_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_G", DoseType = "Average")
	Plexus_brachial_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_G", RelativeVolumes = [0.02])
	
	ReinsAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	ReinsD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	EstomacAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	EstomacD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, coeurValuesV35, coeurValuesV40, coeurValuesV50, coeurAverage, coeurD2, poumons_PTVValuesV5, poumons_PTVValuesV13, poumons_PTVValuesV20, poumons_PTVValuesV30, \
		poumons_PTVAverage, poumons_PTVD2, OesophageValuesV45, OesophageValuesV50, OesophageValuesV55,  OesophageAverage, OesophageD2, ThyroideValuesV50, ThyroideAverage, ThyroideD2, \
		ReinsValuesV12, ReinsValuesV20, ReinsValuesV30, ReinsAverage, ReinsD2, FoieValuesV26, FoieValuesV30, FoieAverage, FoieD2, EstomacValuesV54, EstomacAverage, EstomacD2))


elif tumourLocalisation == "Thorax 2,4Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 2920:
				coeurValuesV292 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3330:
				coeurValuesV333 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4170:
				coeurValuesV417 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV292 = ""
				coeurValuesV333 = ""
				coeurValuesV417 = ""

		elif elm.ForRegionOfInterest.Name == "Poumons-PTV":
			if elm.PlanningGoal.ParameterValue == 2000:
				poumons_PTVValuesV20= (elm.GetClinicalGoalValue())*100
			else:
				poumons_PTVValuesV20 = ""

		elif elm.ForRegionOfInterest.Name == "Oesophage":
			if elm.PlanningGoal.ParameterValue == 4500:
				OesophageValuesV45 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				OesophageValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5500:
				OesophageValuesV55 = (elm.GetClinicalGoalValue())*100
			else:
				OesophageValuesV45 = ""
				OesophageValuesV50 = ""
				OesophageValuesV55 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 4170:
				ThyroideValuesV417 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV417 = ""
		
		elif elm.ForRegionOfInterest.Name == "Reins":
			if elm.PlanningGoal.ParameterValue == 1000:
				ReinsValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1670:
				ReinsValuesV167 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				ReinsValuesV25 = (elm.GetClinicalGoalValue())*100
			else:
				ReinsValuesV10 = ""
				ReinsValuesV167 = ""
				ReinsValuesV25 = ""

		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 1950:
				FoieValuesV195 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2250:
				FoieValuesV225 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV195 = ""
				FoieValuesV225 = ""

		elif elm.ForRegionOfInterest.Name == "Estomac":
			if elm.PlanningGoal.ParameterValue == 5400:
				EstomacValuesV54 = (elm.GetClinicalGoalValue())*100
			else:
				EstomacValuesV54 = ""

	### Get dose statistics informations ###
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumons-PTV", DoseType = "Average")
	poumons_PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumons-PTV", RelativeVolumes = [0.02])
	
	OesophageAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oesophage", DoseType = "Average")
	OesophageD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oesophage", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	Plexus_brachial_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_D", DoseType = "Average")
	Plexus_brachial_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_D", RelativeVolumes = [0.02])

	Plexus_brachial_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_G", DoseType = "Average")
	Plexus_brachial_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_G", RelativeVolumes = [0.02])
	
	ReinsAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	ReinsD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	EstomacAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	EstomacD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, coeurValuesV292, coeurValuesV333, coeurValuesV417, coeurAverage, coeurD2, poumons_PTVValuesV20, \
		poumons_PTVAverage, poumons_PTVD2, OesophageValuesV45, OesophageValuesV50, OesophageValuesV55,  OesophageAverage, OesophageD2, ThyroideValuesV417, ThyroideAverage, ThyroideD2, \
		ReinsValuesV10, ReinsValuesV167, ReinsValuesV25, ReinsAverage, ReinsD2, FoieValuesV195, FoieValuesV225, FoieAverage, FoieD2, EstomacValuesV54, EstomacAverage,EstomacD2))


elif tumourLocalisation == "Thorax 3Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 3360:
				coeurValuesV336 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3840:
				coeurValuesV384 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 480:
				coeurValuesV48 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV292 = ""
				coeurValuesV333 = ""
				coeurValuesV417 = ""

		elif elm.ForRegionOfInterest.Name == "Poumons-PTV":
			if elm.PlanningGoal.ParameterValue == 480:
				poumons_PTVValuesV48= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1250:
				poumons_PTVValuesV125 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1950:
				poumons_PTVValuesV195 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				poumons_PTVValuesV288 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_PTVValuesV48 = ""
				poumons_PTVValuesV125 = ""
				poumons_PTVValuesV195 = ""
				poumons_PTVValuesV288 = ""

		elif elm.ForRegionOfInterest.Name == "Oesophage":
			if elm.PlanningGoal.ParameterValue == 4320:
				OesophageValuesV432 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4800:
				OesophageValuesV48 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5280:
				OesophageValuesV528 = (elm.GetClinicalGoalValue())*100
			else:
				OesophageValuesV432 = ""
				OesophageValuesV48 = ""
				OesophageValuesV528 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 4800:
				ThyroideValuesV48 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV48 = ""
		
		elif elm.ForRegionOfInterest.Name == "Reins":
			if elm.PlanningGoal.ParameterValue == 1150:
				ReinsValuesV115 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				ReinsValuesV192 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				ReinsValuesV288 = (elm.GetClinicalGoalValue())*100
			else:
				ReinsValuesV115 = ""
				ReinsValuesV192 = ""
				ReinsValuesV288 = ""

		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2500:
				FoieValuesV25 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				FoieValuesV288 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV25 = ""
				FoieValuesV288 = ""

		elif elm.ForRegionOfInterest.Name == "Estomac":
			if elm.PlanningGoal.ParameterValue == 5180:
				EstomacValuesV518 = (elm.GetClinicalGoalValue())*100
			else:
				EstomacValuesV518 = ""

	### Get dose statistics informations ###
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumons-PTV", DoseType = "Average")
	poumons_PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumons-PTV", RelativeVolumes = [0.02])
	
	OesophageAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oesophage", DoseType = "Average")
	OesophageD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oesophage", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	Plexus_brachial_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_D", DoseType = "Average")
	Plexus_brachial_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_D", RelativeVolumes = [0.02])

	Plexus_brachial_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus_brachial_G", DoseType = "Average")
	Plexus_brachial_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus_brachial_G", RelativeVolumes = [0.02])
	
	ReinsAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	ReinsD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	EstomacAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	EstomacD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, coeurValuesV336, coeurValuesV384, coeurValuesV48, coeurAverage, coeurD2, poumons_PTVValuesV48, poumons_PTVValuesV125, poumons_PTVValuesV195, poumons_PTVValuesV288, \
		poumons_PTVAverage, poumons_PTVD2, OesophageValuesV432, OesophageValuesV48, OesophageValuesV528,  OesophageAverage, OesophageD2, ThyroideValuesV48, ThyroideAverage, ThyroideD2, \
		ReinsValuesV115, ReinsValuesV192, ReinsValuesV288, ReinsAverage, ReinsD2, FoieValuesV25, FoieValuesV288, FoieAverage, FoieD2, EstomacValuesV518, EstomacAverage,EstomacD2))



elif tumourLocalisation == "Sein D 2Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1500:
				coeurValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				coeurValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				coeurValuesV25 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV15 = ""
				coeurValuesV20 = ""
				coeurValuesV25 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 1500:
				poumons_droitValuesV15= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				poumons_droitValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				poumons_droitValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				poumons_droitValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV15 = ""
				poumons_droitValuesV20 = ""
				poumons_droitValuesV30 = ""
				poumons_droitValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 1000:
				poumons_gaucheValuesV10= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1200:
				poumons_gaucheValuesV12 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1500:
				poumons_gaucheValuesV15 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV10 = ""
				poumons_gaucheValuesV12 = ""
				poumons_gaucheValuesV15 = ""


		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 500:
				sein_controValuesV5 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 700:
				sein_controValuesV7 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1000:
				sein_controValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				sein_controValuesV20 = (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV5 = ""
				sein_controValuesV7 = ""
				sein_controValuesV10 = ""
				sein_controValuesV20 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				ThyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV50 = ""
		
		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2600:
				FoieValuesV26 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				FoieValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV26 = ""
				FoieValuesV30 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_N_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV N DOSI", RelativeVolumes = [0.02])
	PTV_N_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV N DOSI", RelativeVolumes = [0.95])
	PTV_N_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV N DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_N_DOSID95, PTV_N_DOSIAverage, PTV_N_DOSID2, moelleD2, coeurValuesV15, \
		coeurValuesV20, coeurValuesV25, coeurAverage, coeurD2, poumons_droitValuesV15, poumons_droitValuesV20, poumons_droitValuesV30, poumons_droitValuesV35, \
		poumons_droitAverage, poumons_droitD2, poumons_gaucheValuesV10, poumons_gaucheValuesV12, poumons_gaucheValuesV15, poumons_gaucheAverage, poumons_gaucheD2, \
		sein_controValuesV5, sein_controValuesV7, sein_controValuesV10, sein_controValuesV20, sein_controAverage,sein_controD2, ThyroideValuesV50, ThyroideAverage, ThyroideD2, \
		FoieValuesV26, FoieValuesV30, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


elif tumourLocalisation == "Sein G 2Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1500:
				coeurValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				coeurValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				coeurValuesV25 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV15 = ""
				coeurValuesV20 = ""
				coeurValuesV25 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 1500:
				poumons_gaucheValuesV15= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				poumons_gaucheValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				poumons_gaucheValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				poumons_gaucheValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV15 = ""
				poumons_droitValuesV20 = ""
				poumons_droitValuesV30 = ""
				poumons_droitValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 1000:
				poumons_droitValuesV10= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1200:
				poumons_droitValuesV12 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1500:
				poumons_droitValuesV15 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV10 = ""
				poumons_droitValuesV12 = ""
				poumons_droitValuesV15 = ""


		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 500:
				sein_controValuesV5 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 700:
				sein_controValuesV7 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1000:
				sein_controValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				sein_controValuesV20 = (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV5 = ""
				sein_controValuesV7 = ""
				sein_controValuesV10 = ""
				sein_controValuesV20 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				ThyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV50 = ""
		
		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2600:
				FoieValuesV26 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				FoieValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV26 = ""
				FoieValuesV30 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_N_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV N DOSI", RelativeVolumes = [0.02])
	PTV_N_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV N DOSI", RelativeVolumes = [0.95])
	PTV_N_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV N DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_N_DOSID95, PTV_N_DOSIAverage, PTV_N_DOSID2, moelleD2, coeurValuesV15, \
		coeurValuesV20, coeurValuesV25, coeurAverage, coeurD2, poumons_gaucheValuesV15, poumons_gaucheValuesV20, poumons_gaucheValuesV30, poumons_gaucheValuesV35, \
		poumons_gaucheAverage, poumons_gaucheD2, poumons_droitValuesV10, poumons_droitValuesV12, poumons_droitValuesV15, poumons_droitAverage, poumons_droitD2,  \
		sein_controValuesV5, sein_controValuesV7, sein_controValuesV10, sein_controValuesV20, sein_controAverage,sein_controD2, ThyroideValuesV50, ThyroideAverage, ThyroideD2, \
		FoieValuesV26, FoieValuesV30, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


elif tumourLocalisation == "Sein D 2,4Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1440:
				coeurValuesV144 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				coeurValuesV192 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2400:
				coeurValuesV24 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV144 = ""
				coeurValuesV192 = ""
				coeurValuesV24 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 1440:
				poumons_droitValuesV144= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				poumons_droitValuesV192 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				poumons_droitValuesV288 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3360:
				poumons_droitValuesV336 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV144 = ""
				poumons_droitValuesV192 = ""
				poumons_droitValuesV288 = ""
				poumons_droitValuesV336 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 960:
				poumons_gaucheValuesV96= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1150:
				poumons_gaucheValuesV115 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1440:
				poumons_gaucheValuesV144 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV96 = ""
				poumons_gaucheValuesV115 = ""
				poumons_gaucheValuesV144 = ""


		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 480:
				sein_controValuesV48 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 670:
				sein_controValuesV67 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 960:
				sein_controValuesV96 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				sein_controValuesV192 = (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV48 = ""
				sein_controValuesV67 = ""
				sein_controValuesV96 = ""
				sein_controValuesV192 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 4800:
				ThyroideValuesV48 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV48 = ""
		
		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2500:
				FoieValuesV25 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				FoieValuesV288 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV25 = ""
				FoieValuesV288 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV144, \
		coeurValuesV192, coeurValuesV24, coeurAverage, coeurD2, poumons_droitValuesV144, poumons_droitValuesV192, poumons_droitValuesV288, poumons_droitValuesV336, poumons_droitAverage, \
		poumons_droitD2, poumons_gaucheValuesV96, poumons_gaucheValuesV115, poumons_gaucheValuesV144, poumons_gaucheAverage, poumons_gaucheD2,  \
		sein_controValuesV48, sein_controValuesV67, sein_controValuesV96, sein_controValuesV192, sein_controAverage,sein_controD2, ThyroideValuesV48, ThyroideAverage, ThyroideD2, \
		FoieValuesV25, FoieValuesV288, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


elif tumourLocalisation == "Sein G 2,4Gy":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1440:
				coeurValuesV144 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				coeurValuesV192 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2400:
				coeurValuesV24 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV144 = ""
				coeurValuesV192 = ""
				coeurValuesV24 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 1440:
				poumons_gaucheValuesV144= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				poumons_gaucheValuesV192 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				poumons_gaucheValuesV288 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3360:
				poumons_gaucheValuesV336 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV144 = ""
				poumons_gaucheValuesV192 = ""
				poumons_gaucheValuesV288 = ""
				poumons_gaucheValuesV336 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 960:
				poumons_droitValuesV96= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1150:
				poumons_droitValuesV115 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1440:
				poumons_droitValuesV144 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV96 = ""
				poumons_droitValuesV115 = ""
				poumons_droitValuesV144 = ""


		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 480:
				sein_controValuesV48 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 670:
				sein_controValuesV67 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 960:
				sein_controValuesV96 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1920:
				sein_controValuesV192 = (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV48 = ""
				sein_controValuesV67 = ""
				sein_controValuesV96 = ""
				sein_controValuesV192 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 4800:
				ThyroideValuesV48 = (elm.GetClinicalGoalValue())*100
			else:
				ThyroideValuesV48 = ""
		
		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2500:
				FoieValuesV25 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2880:
				FoieValuesV288 = (elm.GetClinicalGoalValue())*100
			else:
				FoieValuesV25 = ""
				FoieValuesV288 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV144, \
		coeurValuesV192, coeurValuesV24, coeurAverage, coeurD2, poumons_gaucheValuesV144, poumons_gaucheValuesV192, poumons_gaucheValuesV288, poumons_gaucheValuesV336, poumons_gaucheAverage, \
		poumons_gaucheD2, poumons_droitValuesV96, poumons_droitValuesV115, poumons_droitValuesV144, poumons_droitAverage, poumons_droitD2,  \
		sein_controValuesV48, sein_controValuesV67, sein_controValuesV96, sein_controValuesV192, sein_controAverage,sein_controD2, ThyroideValuesV48, ThyroideAverage, ThyroideD2, \
		FoieValuesV25, FoieValuesV288, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


elif tumourLocalisation == "Sein D hypoG":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1700:
				coeurValuesV17 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				coeurValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV17 = ""
				coeurValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 1700:
				poumons_droitValuesV17= (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV17 = ""

	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	LADCoronaryAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "LADCoronary", DoseType = "Average")
	LADCoronaryD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "LADCoronary", RelativeVolumes = [0.02])

	Plexus_brachial_homolatéralAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial homolatéral", DoseType = "Average")
	Plexus_brachial_homolatéralD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial homolatéral", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV17, \
		coeurValuesV35, coeurAverage, coeurD2, poumons_droitValuesV17, poumons_droitAverage, poumons_droitD2, poumons_gaucheAverage, poumons_gaucheD2, sein_controAverage,sein_controD2, \
		ThyroideAverage, ThyroideD2, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2, LADCoronaryAverage, LADCoronaryD2, Plexus_brachial_homolatéralAverage, Plexus_brachial_homolatéralD2))


elif tumourLocalisation == "Sein G hypoG":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 1700:
				coeurValuesV17 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				coeurValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV17 = ""
				coeurValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 1700:
				poumons_gaucheValuesV17= (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV17 = ""

	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	LADCoronaryAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "LADCoronary", DoseType = "Average")
	LADCoronaryD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "LADCoronary", RelativeVolumes = [0.02])

	Plexus_brachial_homolatéralAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial homolatéral", DoseType = "Average")
	Plexus_brachial_homolatéralD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial homolatéral", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV17, \
		coeurValuesV35, coeurAverage, coeurD2, poumons_gaucheValuesV17, poumons_gaucheAverage, poumons_gaucheD2, poumons_droitAverage, poumons_droitD2, sein_controAverage,sein_controD2, \
		ThyroideAverage, ThyroideD2, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2, LADCoronaryAverage, LADCoronaryD2, Plexus_brachial_homolatéralAverage, Plexus_brachial_homolatéralD2))


elif tumourLocalisation == "Sein D START":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 800:
				coeurValuesV8 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1800:
				coeurValuesV18 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV8 = ""
				coeurValuesV18 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 400:
				poumons_droitValuesV4= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 800:
				poumons_droitValuesV8 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1600:
				poumons_droitValuesV16 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV4 = ""
				poumons_droitValuesV8 = ""
				poumons_droitValuesV16 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 400:
				poumons_gaucheValuesV4= (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV4 = ""

		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 200:
				sein_controValuesV2= (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV2 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	LADCoronaryAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "LADCoronary", DoseType = "Average")
	LADCoronaryD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "LADCoronary", RelativeVolumes = [0.02])

	Plexus_brachial_homolatéralAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial homolatéral", DoseType = "Average")
	Plexus_brachial_homolatéralD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial homolatéral", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV8, \
		coeurValuesV18, coeurAverage, coeurD2, poumons_gaucheValuesV4, poumons_gaucheValuesV8, poumons_gaucheValuesV16, poumons_gaucheAverage, poumons_gaucheD2, poumons_droitValuesV4, \
		poumons_droitAverage, poumons_droitD2, sein_controValuesV2, sein_controAverage,sein_controD2, ThyroideAverage, ThyroideD2, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))

elif tumourLocalisation == "Sein G START":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 800:
				coeurValuesV8 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1800:
				coeurValuesV18 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV8 = ""
				coeurValuesV18 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon droit":
			if elm.PlanningGoal.ParameterValue == 400:
				poumons_droitValuesV4= (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 800:
				poumons_droitValuesV8 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1600:
				poumons_droitValuesV16 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_droitValuesV4 = ""
				poumons_droitValuesV8 = ""
				poumons_droitValuesV16 = ""

		elif elm.ForRegionOfInterest.Name == "Poumon gauche":
			if elm.PlanningGoal.ParameterValue == 400:
				poumons_gaucheValuesV4= (elm.GetClinicalGoalValue())*100
			else:
				poumons_gaucheValuesV4 = ""

		elif elm.ForRegionOfInterest.Name == "Sein controlatéral":
			if elm.PlanningGoal.ParameterValue == 200:
				sein_controValuesV2= (elm.GetClinicalGoalValue())*100
			else:
				sein_controValuesV2 = ""


	### Get dose statistics informations ###
	PTV_sein_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.02])
	PTV_sein_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV sein DOSI", RelativeVolumes = [0.95])
	PTV_sein_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV sein DOSI", DoseType = "Average")

	PTV_CMID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.02])
	PTV_CMID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV CMI", RelativeVolumes = [0.95])
	PTV_CMIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV CMI", DoseType = "Average")

	PTV_susclav_DOSID2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.02])
	PTV_susclav_DOSID95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV susclav DOSI", RelativeVolumes = [0.95])
	PTV_susclav_DOSIAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV susclav DOSI", DoseType = "Average")

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])
	
	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_droitAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon droit", DoseType = "Average")
	poumons_droitD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon droit", RelativeVolumes = [0.02])

	poumons_gaucheAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumon gauche", DoseType = "Average")
	poumons_gaucheD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumon gauche", RelativeVolumes = [0.02])

	sein_controAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sein controlatéral", DoseType = "Average")
	sein_controD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sein controlatéral", RelativeVolumes = [0.02])

	ThyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Thyroide", DoseType = "Average")
	ThyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Thyroide", RelativeVolumes = [0.02])

	FoieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	FoieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	Tete_huméraleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	Tete_huméraleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])

	LADCoronaryAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "LADCoronary", DoseType = "Average")
	LADCoronaryD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "LADCoronary", RelativeVolumes = [0.02])

	Plexus_brachial_homolatéralAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial homolatéral", DoseType = "Average")
	Plexus_brachial_homolatéralD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial homolatéral", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_sein_DOSID95, PTV_sein_DOSIAverage, PTV_sein_DOSID2, PTV_CMID95, PTV_CMIAverage, PTV_CMID2, PTV_susclav_DOSID95, PTV_susclav_DOSIAverage, PTV_susclav_DOSID2, moelleD2, coeurValuesV8, \
		coeurValuesV18, coeurAverage, coeurD2, poumons_droitValuesV4, poumons_droitValuesV8, poumons_droitValuesV16, poumons_droitAverage, poumons_droitD2, poumons_gaucheValuesV4, \
		poumons_gaucheAverage, poumons_gaucheD2, sein_controValuesV2, sein_controAverage, sein_controD2, ThyroideAverage, ThyroideD2, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


if tumourLocalisation =="Canal anal":

	for elm in goals:
		
		if elm.ForRegionOfInterest.Name == "ParoiVesicale":
			if elm.PlanningGoal.ParameterValue == 6500:
				paroiVesicaleValuesV65 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				paroiVesicaleValuesV70 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 8000:
				paroiVesicaleValuesV80 = (elm.GetClinicalGoalValue())*100
			else:
				paroiVesicaleValuesV65 = ""
				paroiVesicaleValuesV70 = ""
				paroiVesicaleValuesV80 = ""

		elif elm.ForRegionOfInterest.Name == "Grêle (cavité péritonéale)":
			if elm.PlanningGoal.ParameterValue == 4000:
				greleValuesV40 = (elm.GetClinicalGoalValue())*100
			else:
				greleValuesV40 = ""

		elif elm.ForRegionOfInterest.Name == "TFD":
			if elm.PlanningGoal.ParameterValue == 5000:
				TFDValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5500:
				TFDValuesV55 = (elm.GetClinicalGoalValue())*100
			else:
				TFDValuesV50 = ""
				TFDValuesV55 = ""

		elif elm.ForRegionOfInterest.Name == "TFG":
			if elm.PlanningGoal.ParameterValue == 5000:
				TFGValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5500:
				TFGValuesV55 = (elm.GetClinicalGoalValue())*100
			else:
				TFGValuesV50 = ""
				TFGValuesV55 = ""
		
		elif elm.ForRegionOfInterest.Name == "Moelle osseuse / ailes iliaques":
			if elm.PlanningGoal.ParameterValue == 1000:
				moelleValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				moelleValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				moelleValuesV25 = (elm.GetClinicalGoalValue())*100
			else:
				moelleValuesV10 = ""
				moelleValuesV20 = ""
				moelleValuesV25 = ""

		elif elm.ForRegionOfInterest.Name == "Organes génitaux externes":
			if elm.PlanningGoal.ParameterValue == 2000:
				Org_genValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				Org_genValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4000:
				Org_genValuesV40 = (elm.GetClinicalGoalValue())*100
			else:
				Org_genValuesV20 = ""
				Org_genValuesV30 = ""
				Org_genValuesV40 = ""

	### Get dose statistics informations ###
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")

	paroiVesicaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	paroiVesicaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Grêle (cavité péritonéale)", DoseType = "Average")
	greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Grêle (cavité péritonéale)", RelativeVolumes = [0.02])

	colonAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Colon", DoseType = "Average")
	colonD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.02])

	TFDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFD", DoseType = "Average")
	TFD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.02])

	TFGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFG", DoseType = "Average")
	TFG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.02])

	plexusSacreAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PlexusSacree", DoseType = "Average")
	plexusSacreD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PlexusSacree", RelativeVolumes = [0.02])

	moelleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Moelle osseuse / ailes iliaques", DoseType = "Average")
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.02])

	QueueDeChevalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "QueueDeCheval", DoseType = "Average")
	QueueDeChevalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "QueueDeCheval", RelativeVolumes = [0.02])

	Org_genAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Organes génitaux externes", DoseType = "Average")
	Org_genD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Organes génitaux externes", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, paroiVesicaleValuesV65, paroiVesicaleValuesV70, paroiVesicaleValuesV80, paroiVesicaleAverage, paroiVesicaleD2, \
		greleValuesV40, greleAverage,greleD2, colonAverage, colonD2, TFDValuesV50, TFDValuesV55, TFDAverage, \
		TFD2, TFGValuesV50, TFGValuesV55, TFGAverage, TFG2, plexusSacreAverage, plexusSacreD2, moelleValuesV10, moelleValuesV20, moelleValuesV25, moelleAverage, moelleD2, \
		QueueDeChevalAverage, QueueDeChevalD2, Org_genValuesV20, Org_genValuesV30, Org_genValuesV40, Org_genAverage, Org_genD2))


if tumourLocalisation =="Pelvis + prostate 2Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "ParoiRectale":
			if elm.PlanningGoal.ParameterValue == 5000:
				paroiRectaleValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6000:
				paroiRectaleValuesV60 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6500:
				paroiRectaleValuesV65 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				paroiRectaleValuesV70 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7500:
				paroiRectaleValuesV75 = (elm.GetClinicalGoalValue())*100
			else:
				paroiRectaleValuesV50 = ""
				paroiRectaleValuesV60 = ""
				paroiRectaleValuesV65 = ""
				paroiRectaleValuesV70 = ""
				paroiRectaleValuesV75 = ""

		elif elm.ForRegionOfInterest.Name == "ParoiVesicale":
			if elm.PlanningGoal.ParameterValue == 6000:
				paroiVesicaleValuesV60 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				paroiVesicaleValuesV70 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 8000:
				paroiVesicaleValuesV80 = (elm.GetClinicalGoalValue())*100
			else:
				paroiVesicaleValuesV60 = ""
				paroiVesicaleValuesV70 = ""
				paroiVesicaleValuesV80 = ""

		elif elm.ForRegionOfInterest.Name == "CanalAnal":
			if elm.PlanningGoal.ParameterValue == 5600:
				canalAnalValuesV56 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				canalAnalValuesV70 = (elm.GetClinicalGoalValue())*100
			else:
				canalAnalValuesV56 = ""
				canalAnalValuesV70 = ""

		elif elm.ForRegionOfInterest.Name == "Grêle (cavité péritonéale)":
			if elm.PlanningGoal.ParameterValue == 4000:
				greleValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				greleValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				greleValuesV40 = ""
				greleValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Colon" or elm.ForRegionOfInterest.Name == "Sigmoïde" or elm.ForRegionOfInterest.Name == "colon sigmoïde":
			if elm.PlanningGoal.ParameterValue == 4500:
				colonValuesV45 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				colonValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				colonValuesV45 = ""
				colonValuesV50 = ""
		
		elif elm.ForRegionOfInterest.Name == "Moelle osseuse / ailes iliaques":
			if elm.PlanningGoal.ParameterValue == 1000:
				moelleValuesV10 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				moelleValuesV20 = (elm.GetClinicalGoalValue())*100
			else:
				moelleValuesV10 = ""
				moelleValuesV20 = ""

		elif elm.ForRegionOfInterest.Name == "Bulbe pénien":
			if elm.PlanningGoal.ParameterValue == 5000:
				Bulb_penienValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				Bulb_penienValuesV70 = (elm.GetClinicalGoalValue())*100
			else:
				Bulb_penienValuesV50 = ""
				Bulb_penienValuesV70 = ""

	### Get dose statistics informations ###
	PTV_pelD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.02])
	PTV_pelD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.95])
	PTV_pelAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV pelvis", DoseType = "Average")

	PTV_proD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.02])
	PTV_proD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.95])
	PTV_proAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV prostate", DoseType = "Average")

	paroiRectaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Paroi recatale", DoseType = "Average")
	paroiRectaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Paroi recatale", RelativeVolumes = [0.02])

	paroiVesicaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	paroiVesicaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	canalAnalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	canalAnalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Grêle (cavité péritonéale)", DoseType = "Average")
	greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Grêle (cavité péritonéale)", RelativeVolumes = [0.02])
	
	colonD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.05])
	colonAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Colon", DoseType = "Average")
	colonD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.02])

	TFD10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.10])
	TFD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.05])
	TFDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFD", DoseType = "Average")
	TFD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.02])

	TFG10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.10])
	TFG5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.05])
	TFGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFG", DoseType = "Average")
	TFG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.02])

	plexusSacreAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PlexusSacree", DoseType = "Average")
	plexusSacreD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PlexusSacree", RelativeVolumes = [0.02])
	
	moelleD75 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.75])
	moelleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Moelle osseuse / ailes iliaques", DoseType = "Average")
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.02])

	QueueDeChevalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "QueueDeCheval", DoseType = "Average")
	QueueDeChevalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "QueueDeCheval", RelativeVolumes = [0.02])

	Bulb_penienAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Bulbe pénien", DoseType = "Average")
	Bulb_penienD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Bulbe pénien", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_pelD95, PTV_pelAverage, PTV_pelD2, PTV_proD95, PTV_proAverage, PTV_proD2, paroiRectaleValuesV50, paroiRectaleValuesV55, paroiRectaleValuesV60, paroiRectaleValuesV70, paroiRectaleValuesV75, \
		paroiRectaleAverage, paroiRectaleD2, paroiVesicaleValuesV60, paroiVesicaleValuesV70, paroiVesicaleValuesV80, paroiVesicaleAverage, paroiVesicaleD2, canalAnalValuesV56,canalAnalValuesV70,  \
		canalAnalAverage, canalAnalD2, greleValuesV40, greleValuesV50, greleAverage,greleD2, colonValuesV45, colonValuesV50, colonD5, colonAverage, colonD2, TFD10, TFD5, \
		TFDAverage, TFD2,  TFG10, TFG5, TFGAverage, TFG2, plexusSacreAverage, plexusSacreD2, moelleValuesV10, moelleValuesV20, moelleD75, moelleAverage, moelleD2, \
		QueueDeChevalAverage, QueueDeChevalD2, Bulb_penienValuesV50, Bulb_penienValuesV70, Bulb_penienAverage, Bulb_penienD2))


if tumourLocalisation =="Pelvis + prostate 2,4Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "ParoiRectale":
			if elm.PlanningGoal.ParameterValue == 4370:
				paroiRectaleValuesV437 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5670:
				paroiRectaleValuesV567 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6150:
				paroiRectaleValuesV615 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6620:
				paroiRectaleValuesV662 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7100:
				paroiRectaleValuesV71 = (elm.GetClinicalGoalValue())*100
			else:
				paroiRectaleValuesV437 = ""
				paroiRectaleValuesV567 = ""
				paroiRectaleValuesV615 = ""
				paroiRectaleValuesV662 = ""
				paroiRectaleValuesV71 = ""
		
		elif elm.ForRegionOfInterest.Name == "ParoiVesicale":
			if elm.PlanningGoal.ParameterValue == 5440:
				paroiVesicaleValuesV544 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6360:
				paroiVesicaleValuesV636 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7300:
				paroiVesicaleValuesV73 = (elm.GetClinicalGoalValue())*100
			else:
				paroiVesicaleValuesV544 = ""
				paroiVesicaleValuesV636 = ""
				paroiVesicaleValuesV73 = ""

		elif elm.ForRegionOfInterest.Name == "CanalAnal":
			if elm.PlanningGoal.ParameterValue == 5300:
				canalAnalValuesV53 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6620:
				canalAnalValuesV662 = (elm.GetClinicalGoalValue())*100
			else:
				canalAnalValuesV53 = ""
				canalAnalValuesV662 = ""

		elif elm.ForRegionOfInterest.Name == "Grêle (cavité péritonéale)":
			if elm.PlanningGoal.ParameterValue == 3640:
				greleValuesV364 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4550:
				greleValuesV455 = (elm.GetClinicalGoalValue())*100
			else:
				greleValuesV364 = ""
				greleValuesV455 = ""
	
		elif elm.ForRegionOfInterest.Name == "Moelle osseuse / ailes iliaques":
			if elm.PlanningGoal.ParameterValue == 930:
				moelleValuesV93 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1850:
				moelleValuesV185 = (elm.GetClinicalGoalValue())*100
			else:
				moelleValuesV93 = ""
				moelleValuesV185 = ""

		elif elm.ForRegionOfInterest.Name == "Bulbe pénien":
			if elm.PlanningGoal.ParameterValue == 5000:
				Bulb_penienValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				Bulb_penienValuesV70 = (elm.GetClinicalGoalValue())*100
			else:
				Bulb_penienValuesV50 = ""
				Bulb_penienValuesV70 = ""

	### Get dose statistics informations ###
	PTV_pelD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.95])
	PTV_pelAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV pelvis", DoseType = "Average")
	PTV_pelD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.02])
	
	PTV_proD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.95])
	PTV_proAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV prostate", DoseType = "Average")
	PTV_proD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.02])

	paroiRectaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Paroi recatale", DoseType = "Average")
	paroiRectaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Paroi recatale", RelativeVolumes = [0.02])

	paroiVesicaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	paroiVesicaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	canalAnalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	canalAnalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Grêle (cavité péritonéale)", DoseType = "Average")
	greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Grêle (cavité péritonéale)", RelativeVolumes = [0.02])
	
	colonD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.05])
	colonAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Colon", DoseType = "Average")
	colonD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.02])

	TFD10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.10])
	TFD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.05])
	TFDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFD", DoseType = "Average")
	TFD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.02])

	TFG10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.10])
	TFG5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.05])
	TFGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFG", DoseType = "Average")
	TFG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.02])

	plexusSacreAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PlexusSacree", DoseType = "Average")
	plexusSacreD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PlexusSacree", RelativeVolumes = [0.02])
	
	moelleD75 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.75])
	moelleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Moelle osseuse / ailes iliaques", DoseType = "Average")
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.02])

	QueueDeChevalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "QueueDeCheval", DoseType = "Average")
	QueueDeChevalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "QueueDeCheval", RelativeVolumes = [0.02])

	Bulb_penienAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Bulbe pénien", DoseType = "Average")
	Bulb_penienD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Bulbe pénien", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_pelD95, PTV_pelAverage, PTV_pelD2, PTV_proD95, PTV_proAverage, PTV_proD2, paroiRectaleValuesV437, paroiRectaleValuesV568, paroiRectaleValuesV615, paroiRectaleValuesV662, paroiRectaleValuesV71, \
		paroiRectaleAverage, paroiRectaleD2, paroiVesicaleValuesV544, paroiVesicaleValuesV636, paroiVesicaleValuesV73, paroiVesicaleAverage, paroiVesicaleD2, canalAnalValuesV53,canalAnalValuesV662,  \
		canalAnalAverage, canalAnalD2, greleValuesV364, greleValuesV455, greleAverage, greleD2, colonD5, colonAverage, colonD2, TFD10, TFD5, \
		TFDAverage, TFD2,  TFG10, TFG5, TFGAverage, TFG2, plexusSacreAverage, plexusSacreD2, moelleValuesV93, moelleValuesV185, moelleD75, moelleAverage, moelleD2, \
		QueueDeChevalAverage, QueueDeChevalD2, Bulb_penienValuesV50, Bulb_penienValuesV70, Bulb_penienAverage, Bulb_penienD2))


if tumourLocalisation =="Pelvis + prostate 3Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "ParoiRectale":
			if elm.PlanningGoal.ParameterValue == 4400:
				paroiRectaleValuesV44 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5200:
				paroiRectaleValuesV52 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5700:
				paroiRectaleValuesV57 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6100:
				paroiRectaleValuesV61 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6600:
				paroiRectaleValuesV66 = (elm.GetClinicalGoalValue())*100
			else:
				paroiRectaleValuesV44 = ""
				paroiRectaleValuesV52 = ""
				paroiRectaleValuesV57 = ""
				paroiRectaleValuesV61 = ""
				paroiRectaleValuesV66 = ""
		
		elif elm.ForRegionOfInterest.Name == "ParoiVesicale":
			if elm.PlanningGoal.ParameterValue == 4800:
				paroiVesicaleValuesV48 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5600:
				paroiVesicaleValuesV56 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6400:
				paroiVesicaleValuesV64 = (elm.GetClinicalGoalValue())*100
			else:
				paroiVesicaleValuesV48 = ""
				paroiVesicaleValuesV56 = ""
				paroiVesicaleValuesV64 = ""

		elif elm.ForRegionOfInterest.Name == "CanalAnal":
			if elm.PlanningGoal.ParameterValue == 4900:
				canalAnalValuesV49 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6400:
				canalAnalValuesV64 = (elm.GetClinicalGoalValue())*100
			else:
				canalAnalValuesV49 = ""
				canalAnalValuesV64 = ""

		elif elm.ForRegionOfInterest.Name == "Grêle (cavité péritonéale)":
			if elm.PlanningGoal.ParameterValue == 3200:
				greleValuesV32 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4000:
				greleValuesV40 = (elm.GetClinicalGoalValue())*100
			else:
				greleValuesV32 = ""
				greleValuesV40 = ""

		elif elm.ForRegionOfInterest.Name == "Colon" or elm.ForRegionOfInterest.Name == "Sigmoïde" or elm.ForRegionOfInterest.Name == "colon sigmoïde":
			if elm.PlanningGoal.ParameterValue == 3900:
				colonValuesV39 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4400:
				colonValuesV44 = (elm.GetClinicalGoalValue())*100
			else:
				colonValuesV39 = ""
				colonValuesV44 = ""
	
		elif elm.ForRegionOfInterest.Name == "Moelle osseuse / ailes iliaques":
			if elm.PlanningGoal.ParameterValue == 800:
				moelleValuesV8 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 1600:
				moelleValuesV16 = (elm.GetClinicalGoalValue())*100
			else:
				moelleValuesV8 = ""
				moelleValuesV16 = ""

		elif elm.ForRegionOfInterest.Name == "Bulbe pénien":
			if elm.PlanningGoal.ParameterValue == 5000:
				Bulb_penienValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 7000:
				Bulb_penienValuesV70 = (elm.GetClinicalGoalValue())*100
			else:
				Bulb_penienValuesV50 = ""
				Bulb_penienValuesV70 = ""

	### Get dose statistics informations ###
	PTV_pelD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.95])
	PTV_pelAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV pelvis", DoseType = "Average")
	PTV_pelD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV pelvis", RelativeVolumes = [0.02])
	
	PTV_proD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.95])
	PTV_proAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV prostate", DoseType = "Average")
	PTV_proD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV prostate", RelativeVolumes = [0.02])

	paroiRectaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Paroi recatale", DoseType = "Average")
	paroiRectaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Paroi recatale", RelativeVolumes = [0.02])

	paroiVesicaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	paroiVesicaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	canalAnalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ParoiVesicale", DoseType = "Average")
	canalAnalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ParoiVesicale", RelativeVolumes = [0.02])

	greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Grêle (cavité péritonéale)", DoseType = "Average")
	greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Grêle (cavité péritonéale)", RelativeVolumes = [0.02])
	
	colonD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.05])
	colonAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Colon", DoseType = "Average")
	colonD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Colon", RelativeVolumes = [0.02])

	TFD10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.10])
	TFD5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.05])
	TFDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFD", DoseType = "Average")
	TFD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFD", RelativeVolumes = [0.02])

	TFG10 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.10])
	TFG5 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.05])
	TFGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TFG", DoseType = "Average")
	TFG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TFG", RelativeVolumes = [0.02])

	plexusSacreAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PlexusSacree", DoseType = "Average")
	plexusSacreD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PlexusSacree", RelativeVolumes = [0.02])
	
	moelleD75 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.75])
	moelleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Moelle osseuse / ailes iliaques", DoseType = "Average")
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle osseuse / ailes iliaques", RelativeVolumes = [0.02])

	QueueDeChevalAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "QueueDeCheval", DoseType = "Average")
	QueueDeChevalD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "QueueDeCheval", RelativeVolumes = [0.02])

	Bulb_penienAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Bulbe pénien", DoseType = "Average")
	Bulb_penienD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Bulbe pénien", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTV_pelD95, PTV_pelAverage, PTV_pelD2, PTV_proD95, PTV_proAverage, PTV_proD2, paroiRectaleValuesV44, paroiRectaleValuesV52, paroiRectaleValuesV57, paroiRectaleValuesV61, paroiRectaleValuesV66, \
		paroiRectaleAverage, paroiRectaleD2, paroiVesicaleValuesV48, paroiVesicaleValuesV56, paroiVesicaleValuesV64, paroiVesicaleAverage, paroiVesicaleD2, canalAnalValuesV49,canalAnalValuesV54,  \
		canalAnalAverage, canalAnalD2, greleValuesV32, greleValuesV40, greleAverage, greleD2, colonValuesV39, colonValuesV44, colonD5, colonAverage, colonD2, TFD10, TFD5, \
		TFDAverage, TFD2,  TFG10, TFG5, TFGAverage, TFG2, plexusSacreAverage, plexusSacreD2, moelleValuesV8, moelleValuesV16, moelleD75, moelleAverage, moelleD2, \
		QueueDeChevalAverage, QueueDeChevalD2, Bulb_penienValuesV50, Bulb_penienValuesV70, Bulb_penienAverage, Bulb_penienD2))


if tumourLocalisation =="Crane":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Tronc cérébral":
			if elm.PlanningGoal.ParameterValue == 5400:
				tronc_cerValuesV54 = (elm.GetClinicalGoalValue())*100
			else:
				tronc_cerValuesV54 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil D":
			if elm.PlanningGoal.ParameterValue == 3500:
				oeil_DValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_DValuesV35 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil G":
			if elm.PlanningGoal.ParameterValue == 3500:
				oeil_GValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_GValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne D":
			if elm.PlanningGoal.ParameterValue == 4500:
				oreille_intDValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intDValuesV45 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne G":
			if elm.PlanningGoal.ParameterValue == 4500:
				oreille_intGValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intGValuesV45 = ""

		elif elm.ForRegionOfInterest.Name == "Encéphale":
			if elm.PlanningGoal.ParameterValue == 1800:
				encephaleValuesV18 = (elm.GetClinicalGoalValue())*100
			else:
				encephaleValuesV18 = ""


		elif elm.ForRegionOfInterest.Name == "Glande lacrymale D":
			if elm.PlanningGoal.ParameterValue == 2100:
				glande_lacrDValuesV21 = (elm.GetClinicalGoalValue())*100
			else:
				glande_lacrDValuesV21 = ""

		elif elm.ForRegionOfInterest.Name == "Glande lacrymale G":
			if elm.PlanningGoal.ParameterValue == 2100:
				glande_lacrGValuesV21 = (elm.GetClinicalGoalValue())*100
			else:
				glande_lacrGValuesV21 = ""
	

	### Get dose statistics informations ###
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])

	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	tronc_cerD05 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.005])
	tronc_cerAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TroncCerebral", DoseType = "Average")
	tronc_cerD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.02])

	chiasmaAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Chiasma", DoseType = "Average")
	chiasmaD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Chiasma", RelativeVolumes = [0.02])

	oeil_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil D", DoseType = "Average")
	oeil_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil D", RelativeVolumes = [0.02])

	oeil_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil G", DoseType = "Average")
	oeil_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil G", RelativeVolumes = [0.02])

	cristallin_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cristallin D", DoseType = "Average")
	cristallin_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cristallin D", RelativeVolumes = [0.02])

	cristallin_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cristallin G", DoseType = "Average")
	cristallin_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cristallin G", RelativeVolumes = [0.02])

	nerf_opt_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique D", DoseType = "Average")
	nerf_opt_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique D", RelativeVolumes = [0.02])

	nerf_opt_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique G", DoseType = "Average")
	nerf_opt_GG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique G", RelativeVolumes = [0.02])

	lobe_tempDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal D", DoseType = "Average")
	lobe_tempDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal D", RelativeVolumes = [0.02])

	lobe_tempGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal G", DoseType = "Average")
	lobe_tempG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal G", RelativeVolumes = [0.02])

	oreille_intDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne D", DoseType = "Average")
	oreille_intDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne D", RelativeVolumes = [0.02])

	oreille_intGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne G", DoseType = "Average")
	oreille_intG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne G", RelativeVolumes = [0.02])

	hypophyseAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hypophyse", DoseType = "Average")
	hypophyseD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hypophyse", RelativeVolumes = [0.02])

	encephaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Encéphale", DoseType = "Average")
	encephaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Encéphale", RelativeVolumes = [0.02])

	peauAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Peau", DoseType = "Average")
	peauD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Peau", RelativeVolumes = [0.02])

	glande_lacrDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Glande lacrymale D", DoseType = "Average")
	glande_lacrDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Glande lacrymale D", RelativeVolumes = [0.02])
	
	glande_lacrGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Glande lacrymale G", DoseType = "Average")
	glande_lacrGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Glande lacrymale G", RelativeVolumes = [0.02])

	hippocampeAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hippocampe", DoseType = "Average")
	hippocampeGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hippocampe", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, tronc_cerValuesV54, tronc_cerD05, tronc_cerAverage, tronc_cerD2, chiasmaAverage, chiasmaD2, oeil_DValuesV35, oeil_DAverage, oeil_DD2, \
		oeil_GValuesV35, oeil_GAverage, oeil_GD2, cristallin_DAverage, cristallin_D2, cristallin_GAverage, cristallin_G2, nerf_opt_DAverage, nerf_opt_D2, nerf_opt_GAverage, nerf_opt_G2, \
		lobe_tempDAverage, lobe_tempDD2, lobe_tempGAverage, lobe_tempGD2, oreille_intDValuesV45, oreille_intDAverage, oreille_intDD2, oreille_intGValuesV45, oreille_intGAverage, oreille_intGD2, \
		hypophyseAverage, hypophyseD2, encephaleValuesV20, encephaleValuesV50, encephaleValuesV60, encephaleValuesV45, encephaleAverage, encephaleD2, peauAverage, peauD2, \
		glande_lacrDValuesV26, glande_lacrDAverage, glande_lacrD2, glande_lacrGValuesV26, glande_lacrGAverage, glande_lacrGD2, hippocampeAverage, hippocampeD2))



if tumourLocalisation =="Crane 40,05Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Tronc cérébral":
			if elm.PlanningGoal.ParameterValue == 4760:
				tronc_cerValuesV476 = (elm.GetClinicalGoalValue())*100
			else:
				tronc_cerValuesV476 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil D":
			if elm.PlanningGoal.ParameterValue == 2900:
				oeil_DValuesV29 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_DValuesV29 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil G":
			if elm.PlanningGoal.ParameterValue == 2900:
				oeil_GValuesV29 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_GValuesV29 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne D":
			if elm.PlanningGoal.ParameterValue == 3700:
				oreille_intDValuesV37 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intDValuesV37 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne G":
			if elm.PlanningGoal.ParameterValue == 3700:
				oreille_intGValuesV37 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intGValuesV37 = ""

		elif elm.ForRegionOfInterest.Name == "Encéphale":
			if elm.PlanningGoal.ParameterValue == 2000:
				encephaleValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				encephaleValuesV50 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 6000:
				encephaleValuesV60 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4500:
				encephaleValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				encephaleValuesV20 = ""
				encephaleValuesV50 = ""
				encephaleValuesV60 = ""
				encephaleValuesV45 = ""

		elif elm.ForRegionOfInterest.Name == "Glande lacrymale D":
			if elm.PlanningGoal.ParameterValue == 2600:
				glande_lacrDValuesV26 = (elm.GetClinicalGoalValue())*100
			else:
				glande_lacrDValuesV26 = ""

		elif elm.ForRegionOfInterest.Name == "Glande lacrymale G":
			if elm.PlanningGoal.ParameterValue == 2600:
				glande_lacrGValuesV26 = (elm.GetClinicalGoalValue())*100
			else:
				glande_lacrGValuesV26 = ""
	

	### Get dose statistics informations ###
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	tronc_cerD05 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.005])
	tronc_cerAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TroncCerebral", DoseType = "Average")
	tronc_cerD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.02])

	chiasmaAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Chiasma", DoseType = "Average")
	chiasmaD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Chiasma", RelativeVolumes = [0.02])

	oeil_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil D", DoseType = "Average")
	oeil_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil D", RelativeVolumes = [0.02])

	oeil_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil G", DoseType = "Average")
	oeil_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil G", RelativeVolumes = [0.02])

	cristallin_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cristallin D", DoseType = "Average")
	cristallin_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cristallin D", RelativeVolumes = [0.02])

	cristallin_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cristallin G", DoseType = "Average")
	cristallin_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cristallin G", RelativeVolumes = [0.02])

	nerf_opt_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique D", DoseType = "Average")
	nerf_opt_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique D", RelativeVolumes = [0.02])

	nerf_opt_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique G", DoseType = "Average")
	nerf_opt_GG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique G", RelativeVolumes = [0.02])

	lobe_tempDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal D", DoseType = "Average")
	lobe_tempDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal D", RelativeVolumes = [0.02])

	lobe_tempGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal G", DoseType = "Average")
	lobe_tempG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal G", RelativeVolumes = [0.02])

	oreille_intDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne D", DoseType = "Average")
	oreille_intDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne D", RelativeVolumes = [0.02])

	oreille_intGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne G", DoseType = "Average")
	oreille_intG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne G", RelativeVolumes = [0.02])

	hypophyseAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hypophyse", DoseType = "Average")
	hypophyseD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hypophyse", RelativeVolumes = [0.02])

	encephaleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Encéphale", DoseType = "Average")
	encephaleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Encéphale", RelativeVolumes = [0.02])

	peauAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Peau", DoseType = "Average")
	peauD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Peau", RelativeVolumes = [0.02])

	glande_lacrDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Glande lacrymale D", DoseType = "Average")
	glande_lacrDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Glande lacrymale D", RelativeVolumes = [0.02])
	
	glande_lacrGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Glande lacrymale G", DoseType = "Average")
	glande_lacrGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Glande lacrymale G", RelativeVolumes = [0.02])

	hippocampeAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hippocampe", DoseType = "Average")
	hippocampeGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hippocampe", RelativeVolumes = [0.02])


	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, tronc_cerValuesV476, tronc_cerD05, tronc_cerAverage, tronc_cerD2, chiasmaAverage, chiasmaD2, oeil_DValuesV29, oeil_DAverage, oeil_DD2, \
		oeil_GValuesV29, oeil_GAverage, oeil_GD2, cristallin_DAverage, cristallin_D2, cristallin_GAverage, cristallin_G2, nerf_opt_DAverage, nerf_opt_D2, nerf_opt_GAverage, nerf_opt_G2, \
		lobe_tempDAverage, lobe_tempDD2, lobe_tempGAverage, lobe_tempGD2, oreille_intDValuesV37, oreille_intDAverage, oreille_intDD2, oreille_intGValuesV37, oreille_intGAverage, oreille_intGD2, \
		hypophyseAverage, hypophyseD2, encephaleValuesV18, encephaleAverage, encephaleD2, peauAverage, peauD2, glande_lacrDValuesV21, glande_lacrDAverage, glande_lacrD2, glande_lacrGValuesV21, \
		glande_lacrGAverage, glande_lacrGD2, hippocampeAverage, hippocampeD2))



if tumourLocalisation =="Oesophage":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Coeur":
			if elm.PlanningGoal.ParameterValue == 3500:
				coeurValuesV35 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4000:
				coeurValuesV40 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 5000:
				coeurValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				coeurValuesV35 = ""
				coeurValuesV40 = ""
				coeurValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Poumons-PTV":
			if elm.PlanningGoal.ParameterValue == 2000:
				poumons_PTVValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				poumons_PTVValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				poumons_PTVValuesV20 = ""
				poumons_PTVValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Foie":
			if elm.PlanningGoal.ParameterValue == 2000:
				foieValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				foieValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				foieValuesV20 = ""
				foieValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Estomac":
			if elm.PlanningGoal.ParameterValue == 5400:
				estomacValuesV54 = (elm.GetClinicalGoalValue())*100
			else:
				estomacValuesV54 = ""


		elif elm.ForRegionOfInterest.Name == "colone grêle":		                ################# ATTENTION POUR CELUI CI IL FAUT DEUX PARAMETRES DU V30
			if elm.PlanningGoal.ParameterValue == 3000:									################### un en % l'autre en cc, regarder sur patient type Cindy
				colone_greleValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:									################### un en % l'autre en cc, regarder sur patient type Cindy
				colone_greleValuesV35 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4500:									################### un en % l'autre en cc, regarder sur patient type Cindy
				colone_greleValuesV45 = (elm.GetClinicalGoalValue())*100

		elif elm.ForRegionOfInterest.Name == "Reins":
			if elm.PlanningGoal.ParameterValue == 1200:
				ReinsValuesV12 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2000:
				ReinsValuesV20 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				ReinsValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				ReinsValuesV12 = ""
				ReinsValuesV20 = ""
				ReinsValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				thyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				thyroideValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Duodénum":
			if elm.PlanningGoal.ParameterValue == 4500:
				duodenumValuesV45 = (elm.GetClinicalGoalValue())*100
			if elm.PlanningGoal.ParameterValue == 5000:
				duodenumValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				duodenumValuesV45 = ""
				duodenumValuesV50 = ""
	
		elif elm.ForRegionOfInterest.Name == "Œsophage sain":
			if elm.PlanningGoal.ParameterValue == 4500:
				oesophage_sainValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				oesophage_sainValuesV45 = ""

	### Get dose statistics informations ###
	PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.95])
	PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV", DoseType = "Average")
	PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV", RelativeVolumes = [0.02])
	
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	coeurAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Coeur", DoseType = "Average")
	coeurD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Coeur", RelativeVolumes = [0.02])

	poumons_PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Poumons-PTV", DoseType = "Average")
	poumons_PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Poumons-PTV", RelativeVolumes = [0.02])

	foieAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Foie", DoseType = "Average")
	foieD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Foie", RelativeVolumes = [0.02])
	
	estomacAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Estomac", DoseType = "Average")
	estomacD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Estomac", RelativeVolumes = [0.02])

	colone_greleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "colon grêle", DoseType = "Average")
	colone_greleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "colon grêle", RelativeVolumes = [0.02])

	reinsAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	reinsD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	thyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	thyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	plexus_brachAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus bracchiaux", DoseType = "Average")
	plexus_brachD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus bracchiaux", RelativeVolumes = [0.02])

	duodenumAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Duodénum", DoseType = "Average")
	duodenumD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Duodénum", RelativeVolumes = [0.02])

	oesophage_sainAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Œsophage sain", DoseType = "Average")
	oesophage_sainD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Œsophage sain", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTVD95, PTVAverage, PTVD2, moelleD2, coeurValuesV35, coeurValuesV40, coeurValuesV50, coeurAverage, coeurD2, poumons_PTVValuesV20,poumons_PTVValuesV30, foieValuesV20, foieValuesV30, \
		foieAverage, foieD2, estomacValuesV54, estomacAverage, estomacD2, colone_greleValuesV30, ####ATTENTION X2 #### \
		colone_greleValuesV35, colone_greleValuesV45, colone_greleAverage, colone_greleD2, ReinsValuesV12, ReinsValuesV20, ReinsValuesV30, reinsAverage, reinsD2, thyroideValuesV50, thyroideAverage, thyroideD2, \
		plexus_brachAverage, plexus_brachD2, duodenumValuesV45, duodenumValuesV50, duodenumAverage, duodenumD2, oesophage_sainValuesV45, oesophage_sainAverage, oesophage_sainD2))


if tumourLocalisation =="ORL 2Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Parotide D":
			if elm.PlanningGoal.ParameterValue == 1500:
				parotide_DValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				parotide_DValuesV25 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				parotide_DValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_DValuesV15 = ""
				parotide_DValuesV25 = ""
				parotide_DValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide G":
			if elm.PlanningGoal.ParameterValue == 1500:
				parotide_GValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 2500:
				parotide_GValuesV25 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:
				parotide_GValuesV30 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_GValuesV15 = ""
				parotide_GValuesV25 = ""
				parotide_GValuesV30 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide unique":
			if elm.PlanningGoal.ParameterValue == 2000:
				parotide_UValuesV20 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_UValuesV20 = ""

		elif elm.ForRegionOfInterest.Name == "Larynx":
			if elm.PlanningGoal.ParameterValue == 3000:
				larynxValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				larynxValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				larynxValuesV30 = ""
				larynxValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				thyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				thyroideValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne D":
			if elm.PlanningGoal.ParameterValue == 4500:
				oreille_intDValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intDValuesV45 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne G":
			if elm.PlanningGoal.ParameterValue == 4500:
				oreille_intGValuesV45 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intGValuesV45 = ""

		elif elm.ForRegionOfInterest.Name == "Cavité buccale":		                
			if elm.PlanningGoal.ParameterValue == 1500:								
				cav_bucValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:									
				colone_greleValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4500:									
				colone_greleValuesV45 = (elm.GetClinicalGoalValue())*100

		elif elm.ForRegionOfInterest.Name == "Sous max D":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxDValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxDValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Sous max G":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxGValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxGValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Oeil D":
			if elm.PlanningGoal.ParameterValue == 3500:
				oeil_DValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_DValuesV35 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil G":
			if elm.PlanningGoal.ParameterValue == 3500:
				oeil_GValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_GValuesV35 = ""

	### Get dose statistics informations ###
	PTV_70D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 70", RelativeVolumes = [0.95])
	PTV_70Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 70", DoseType = "Average")
	PTV_70D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 70", RelativeVolumes = [0.02])

	PTV_63D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 63", RelativeVolumes = [0.95])
	PTV_63Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 63", DoseType = "Average")
	PTV_63D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 63", RelativeVolumes = [0.02])

	PTV_56D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 56", RelativeVolumes = [0.95])
	PTV_56Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 56", DoseType = "Average")
	PTV_56D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 56", RelativeVolumes = [0.02])
	
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	tronc_cerAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TroncCerebral", DoseType = "Average")
	tronc_cerD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.02])

	parotide_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide D", DoseType = "Average")
	parotide_D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide D", RelativeVolumes = [0.02])

	parotide_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide G", DoseType = "Average")
	parotide_G2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide G", RelativeVolumes = [0.02])

	parotide_UAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide unique", DoseType = "Average")
	parotide_U2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide unique", RelativeVolumes = [0.02])

	larynxAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Larynx", DoseType = "Average")
	larynxD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Larynx", RelativeVolumes = [0.02])

	thyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	thyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	plex_brachDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial D", DoseType = "Average")
	plex_brachDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial D", RelativeVolumes = [0.02])

	plex_brachGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial G", DoseType = "Average")
	plex_brachGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial G", RelativeVolumes = [0.02])

	oreille_intDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne D", DoseType = "Average")
	oreille_intDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne D", RelativeVolumes = [0.02])

	oreille_intGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne G", DoseType = "Average")
	oreille_intG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne G", RelativeVolumes = [0.02])

	oreille_moyDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne D", DoseType = "Average")
	oreille_moyDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne D", RelativeVolumes = [0.02])

	oreille_moyGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne G", DoseType = "Average")
	oreille_moyG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne G", RelativeVolumes = [0.02])

	ATMDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM D", DoseType = "Average")
	ATMDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM D", RelativeVolumes = [0.02])

	ATMGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM G", DoseType = "Average")
	ATMGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM G", RelativeVolumes = [0.02])

	chiasmaAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Chiasma", DoseType = "Average")
	chiasmaD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Chiasma", RelativeVolumes = [0.02])

	cerveletAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cervelet", DoseType = "Average")
	cerveletD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cervelet", RelativeVolumes = [0.02])

	mandibuleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Mandibule", DoseType = "Average")
	mandibuleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Mandibule", RelativeVolumes = [0.02])
	
	nerf_opt_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique D", DoseType = "Average")
	nerf_opt_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique D", RelativeVolumes = [0.02])

	nerf_opt_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique G", DoseType = "Average")
	nerf_opt_GG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique G", RelativeVolumes = [0.02])

	cav_bucAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cavité buccale", DoseType = "Average")
	cav_bucD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cavité buccale", RelativeVolumes = [0.02])

	lobe_tempDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal D", DoseType = "Average")
	lobe_tempDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal D", RelativeVolumes = [0.02])

	lobe_tempGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal G", DoseType = "Average")
	lobe_tempG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal G", RelativeVolumes = [0.02])

	sous_maxDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max D", DoseType = "Average")
	sous_maxDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max D", RelativeVolumes = [0.02])	

	sous_maxGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max G", DoseType = "Average")
	sous_maxGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max G", RelativeVolumes = [0.02])

	hypophyseAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hypophyse", DoseType = "Average")
	hypophyseD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hypophyse", RelativeVolumes = [0.02])

	oeil_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil D", DoseType = "Average")
	oeil_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil D", RelativeVolumes = [0.02])

	oeil_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil G", DoseType = "Average")
	oeil_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil G", RelativeVolumes = [0.02])

	cons_pharAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Constricteurs du pharynx", DoseType = "Average")
	cons_pharD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Constricteurs du pharynx", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_70D95, PTV_70Average, PTV_70D2, PTV_63D95, PTV_63Average, PTV_63D2, PTV_56D95, PTV_56Average, PTV_56D2, moelleD2, tronc_cerAverage, tronc_cerD2, \
		parotide_DValuesV15, parotide_DValuesV25, parotide_DValuesV30, parotide_DAverage, parotide_DD2, parotide_GValuesV15, parotide_GValuesV25, parotide_GValuesV30, parotide_GAverage, parotide_GD2, \
		parotide_UValuesV20, parotide_UAverage, parotide_UD2, larynxValuesV30, larynxValuesV35, thyroideValuesV50, thyroideAverage, thyroideD2, plex_brachDAverage, plex_brachD2, \
		plex_brachGAverage, plex_brachG2, oreille_intDValuesV45, oreille_intDAverage, oreille_intDD2, oreille_intGValuesV45, oreille_intGAverage, oreille_intGD2, \
		oreille_moyDAverage, oreille_moyDD2, oreille_moyGAverage, oreille_moyGD2, ATMDAverage, ATMDD2, ATMGAverage, ATMGD2, chiasmaAverage, chiasmaD2, \
		cerveletAverage, cerveletD2, mandibuleAverage, mandibuleD2, nerf_opt_DAverage, nerf_opt_DG2, nerf_opt_GAverage, nerf_opt_GG2, cav_bucValuesV15, cav_bucValuesV30, \
		cav_bucValuesV45, cav_bucAverage, cav_bucD2, lobe_tempDAverage, lobe_tempDD2, lobe_tempGAverage, lobe_tempG2, sous_maxDValuesV35, \
		sous_maxDAverage, sous_maxD2, sous_maxGValuesV35, sous_maxGAverage, sous_maxG2, hypophyseAverage, hypophyseD2, oeil_DValuesV35, oeil_DAverage, oeil_DD2, \
		oeil_GValuesV35, oeil_GAverage, oeil_GD2, cons_pharAverage, cons_pharD2))


if tumourLocalisation =="ORL 2,4Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Parotide D":
			if elm.PlanningGoal.ParameterValue == 2300:
				parotide_DValuesV23 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3300:
				parotide_DValuesV33 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3700:
				parotide_DValuesV37 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_DValuesV23 = ""
				parotide_DValuesV33 = ""
				parotide_DValuesV37 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide G":
			if elm.PlanningGoal.ParameterValue == 2300:
				parotide_GValuesV23 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3300:
				parotide_GValuesV33 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3700:
				parotide_GValuesV37 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_GValuesV23 = ""
				parotide_GValuesV33 = ""
				parotide_GValuesV37 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide unique":
			if elm.PlanningGoal.ParameterValue == 2800:
				parotide_UValuesV28 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_UValuesV28 = ""

		elif elm.ForRegionOfInterest.Name == "Larynx":
			if elm.PlanningGoal.ParameterValue == 3600:
				larynxValuesV36 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4800:
				larynxValuesV48 = (elm.GetClinicalGoalValue())*100
			else:
				larynxValuesV36 = ""
				larynxValuesV48 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5300:
				thyroideValuesV53 = (elm.GetClinicalGoalValue())*100
			else:
				thyroideValuesV53 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne D":
			if elm.PlanningGoal.ParameterValue == 5000:
				oreille_intDValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intDValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne G":
			if elm.PlanningGoal.ParameterValue == 5000:
				oreille_intGValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intGValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Cavité buccale":		                
			if elm.PlanningGoal.ParameterValue == 1500:								
				cav_bucValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:									
				colone_greleValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4500:									
				colone_greleValuesV35 = (elm.GetClinicalGoalValue())*100

		elif elm.ForRegionOfInterest.Name == "Sous max D":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxDValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxDValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Sous max G":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxGValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxGValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Oeil D":
			if elm.PlanningGoal.ParameterValue == 4300:
				oeil_DValuesV43 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_DValuesV43 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil G":
			if elm.PlanningGoal.ParameterValue == 4300:
				oeil_GValuesV43 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_GValuesV43 = ""

	### Get dose statistics informations ###
	PTV_696D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 69,6", RelativeVolumes = [0.95])
	PTV_696Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 69,6", DoseType = "Average")
	PTV_696D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 69,6", RelativeVolumes = [0.02])

	PTV_623595 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 62,35", RelativeVolumes = [0.95])
	PTV_6235Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 62,35", DoseType = "Average")
	PTV_6235D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 62,35", RelativeVolumes = [0.02])

	PTV_5365D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 53,65", RelativeVolumes = [0.95])
	PTV_5365Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 53,65", DoseType = "Average")
	PTV_5365D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 53,65", RelativeVolumes = [0.02])
	
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	tronc_cerAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TroncCerebral", DoseType = "Average")
	tronc_cerD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.02])

	parotide_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide D", DoseType = "Average")
	parotide_D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide D", RelativeVolumes = [0.02])

	parotide_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide G", DoseType = "Average")
	parotide_G2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide G", RelativeVolumes = [0.02])

	parotide_UAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide unique", DoseType = "Average")
	parotide_U2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide unique", RelativeVolumes = [0.02])

	larynxAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Larynx", DoseType = "Average")
	larynxD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Larynx", RelativeVolumes = [0.02])

	thyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	thyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	plex_brachDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial D", DoseType = "Average")
	plex_brachDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial D", RelativeVolumes = [0.02])

	plex_brachGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial G", DoseType = "Average")
	plex_brachGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial G", RelativeVolumes = [0.02])

	oreille_intDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne D", DoseType = "Average")
	oreille_intDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne D", RelativeVolumes = [0.02])

	oreille_intGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne G", DoseType = "Average")
	oreille_intG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne G", RelativeVolumes = [0.02])

	oreille_moyDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne D", DoseType = "Average")
	oreille_moyDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne D", RelativeVolumes = [0.02])

	oreille_moyGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne G", DoseType = "Average")
	oreille_moyG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne G", RelativeVolumes = [0.02])

	ATMDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM D", DoseType = "Average")
	ATMDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM D", RelativeVolumes = [0.02])

	ATMGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM G", DoseType = "Average")
	ATMGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM G", RelativeVolumes = [0.02])

	chiasmaAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Chiasma", DoseType = "Average")
	chiasmaD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Chiasma", RelativeVolumes = [0.02])

	cerveletAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cervelet", DoseType = "Average")
	cerveletD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cervelet", RelativeVolumes = [0.02])

	mandibuleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Mandibule", DoseType = "Average")
	mandibuleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Mandibule", RelativeVolumes = [0.02])
	
	nerf_opt_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique D", DoseType = "Average")
	nerf_opt_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique D", RelativeVolumes = [0.02])

	nerf_opt_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique G", DoseType = "Average")
	nerf_opt_GG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique G", RelativeVolumes = [0.02])

	cav_bucAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cavité buccale", DoseType = "Average")
	cav_bucD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cavité buccale", RelativeVolumes = [0.02])

	lobe_tempDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal D", DoseType = "Average")
	lobe_tempDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal D", RelativeVolumes = [0.02])

	lobe_tempGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal G", DoseType = "Average")
	lobe_tempG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal G", RelativeVolumes = [0.02])

	sous_maxDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max D", DoseType = "Average")
	sous_maxDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max D", RelativeVolumes = [0.02])	

	sous_maxGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max G", DoseType = "Average")
	sous_maxGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max G", RelativeVolumes = [0.02])

	hypophyseAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hypophyse", DoseType = "Average")
	hypophyseD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hypophyse", RelativeVolumes = [0.02])

	oeil_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil D", DoseType = "Average")
	oeil_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil D", RelativeVolumes = [0.02])

	oeil_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil G", DoseType = "Average")
	oeil_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil G", RelativeVolumes = [0.02])

	cons_pharAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Constricteurs du pharynx", DoseType = "Average")
	cons_pharD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Constricteurs du pharynx", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_696D95, PTV_696Average, PTV_696D2, PTV_623595, PTV_6235Average, PTV_6235D2, PTV_5365D95, PTV_5365verage, PTV_5365D2, moelleD2, tronc_cerAverage, tronc_cerD2, \
		parotide_DValuesV23, parotide_DValuesV33, parotide_DValuesV37, parotide_DAverage, parotide_DD2, parotide_GValuesV23, parotide_GValuesV33, parotide_GValuesV37, parotide_GAverage, parotide_GD2, \
		parotide_UValuesV28, parotide_UAverage, parotide_UD2, larynxValuesV36, larynxValuesV48, thyroideValuesV53, thyroideAverage, thyroideD2, plex_brachDAverage, plex_brachD2, \
		plex_brachGAverage, plex_brachG2, oreille_intDValuesV50, oreille_intDAverage, oreille_intDD2, oreille_intGValuesV50, oreille_intGAverage, oreille_intGD2, \
		oreille_moyDAverage, oreille_moyDD2, oreille_moyGAverage, oreille_moyGD2, ATMDAverage, ATMDD2, ATMGAverage, ATMGD2, chiasmaAverage, chiasmaD2, cerveletAverage, cerveletD2, \
		mandibuleAverage, mandibuleD2, nerf_opt_DAverage, nerf_opt_DG2, nerf_opt_GAverage, nerf_opt_GG2, cav_bucValuesV15, cav_bucValuesV30, cav_bucValuesV35, cav_bucAverage, cav_bucD2, \
		lobe_tempDAverage, lobe_tempDD2, lobe_tempGAverage, lobe_tempG2, sous_maxDValuesV35, sous_maxDAverage, sous_maxD2, sous_maxGValuesV35, sous_maxGAverage, sous_maxG2, \
		hypophyseAverage, hypophyseD2, oeil_DValuesV43, oeil_DAverage, oeil_DD2, oeil_GValuesV43, oeil_GAverage, oeil_GD2, cons_pharAverage, cons_pharD2))


if tumourLocalisation =="ORL 3Gy":

	for elm in goals:

		if elm.ForRegionOfInterest.Name == "Parotide D":
			if elm.PlanningGoal.ParameterValue == 2200:
				parotide_DValuesV22 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3100:
				parotide_DValuesV31 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				parotide_DValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_DValuesV22 = ""
				parotide_DValuesV31 = ""
				parotide_DValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide G":
			if elm.PlanningGoal.ParameterValue == 2200:
				parotide_GValuesV22 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3100:
				parotide_GValuesV31 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3500:
				parotide_GValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_GValuesV22 = ""
				parotide_GValuesV31 = ""
				parotide_GValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Parotide unique":
			if elm.PlanningGoal.ParameterValue == 2600:
				parotide_UValuesV26 = (elm.GetClinicalGoalValue())*100
			else:
				parotide_UValuesV26 = ""

		elif elm.ForRegionOfInterest.Name == "Larynx":
			if elm.PlanningGoal.ParameterValue == 3500:
				larynxValuesV35 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4600:
				larynxValuesV46 = (elm.GetClinicalGoalValue())*100
			else:
				larynxValuesV35 = ""
				larynxValuesV46 = ""

		elif elm.ForRegionOfInterest.Name == "Thyroide":
			if elm.PlanningGoal.ParameterValue == 5000:
				thyroideValuesV50 = (elm.GetClinicalGoalValue())*100
			else:
				thyroideValuesV50 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne D":
			if elm.PlanningGoal.ParameterValue == 4600:
				oreille_intDValuesV46 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intDValuesV46 = ""

		elif elm.ForRegionOfInterest.Name == "Oreille interne G":
			if elm.PlanningGoal.ParameterValue == 4600:
				oreille_intGValuesV46 = (elm.GetClinicalGoalValue())*100
			else:
				oreille_intGValuesV46 = ""

		elif elm.ForRegionOfInterest.Name == "Cavité buccale":		                
			if elm.PlanningGoal.ParameterValue == 1500:								
				cav_bucValuesV15 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 3000:									
				colone_greleValuesV30 = (elm.GetClinicalGoalValue())*100
			elif elm.PlanningGoal.ParameterValue == 4500:									
				colone_greleValuesV35 = (elm.GetClinicalGoalValue())*100

		elif elm.ForRegionOfInterest.Name == "Sous max D":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxDValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxDValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Sous max G":
			if elm.PlanningGoal.ParameterValue == 3500:
				sous_maxGValuesV35 = (elm.GetClinicalGoalValue())*100
			else:
				sous_maxGValuesV35 = ""

		elif elm.ForRegionOfInterest.Name == "Oeil D":
			if elm.PlanningGoal.ParameterValue == 4000:
				oeil_DValuesV40 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_DValuesV40 = ""
		
		elif elm.ForRegionOfInterest.Name == "Oeil G":
			if elm.PlanningGoal.ParameterValue == 4000:
				oeil_GValuesV40 = (elm.GetClinicalGoalValue())*100
			else:
				oeil_GValuesV40 = ""

	### Get dose statistics informations ###
	PTV_69D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 69", RelativeVolumes = [0.95])
	PTV_69Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 69", DoseType = "Average")
	PTV_69D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 69", RelativeVolumes = [0.02])

	PTV_62195 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 62,1", RelativeVolumes = [0.95])
	PTV_621Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 62,1", DoseType = "Average")
	PTV_621D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 62,1", RelativeVolumes = [0.02])

	PTV_5403D95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 54,03", RelativeVolumes = [0.95])
	PTV_5403Average = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 554,03", DoseType = "Average")
	PTV_5403D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 54,03", RelativeVolumes = [0.02])
	
	moelleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Moelle", RelativeVolumes = [0.02])

	tronc_cerAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "TroncCerebral", DoseType = "Average")
	tronc_cerD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "TroncCerebral", RelativeVolumes = [0.02])

	parotide_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide D", DoseType = "Average")
	parotide_D2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide D", RelativeVolumes = [0.02])

	parotide_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide G", DoseType = "Average")
	parotide_G2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide G", RelativeVolumes = [0.02])

	parotide_UAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Parotide unique", DoseType = "Average")
	parotide_U2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Parotide unique", RelativeVolumes = [0.02])

	larynxAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Larynx", DoseType = "Average")
	larynxD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Larynx", RelativeVolumes = [0.02])

	thyroideAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Reins", DoseType = "Average")
	thyroideD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Reins", RelativeVolumes = [0.02])

	plex_brachDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial D", DoseType = "Average")
	plex_brachDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial D", RelativeVolumes = [0.02])

	plex_brachGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Plexus brachial G", DoseType = "Average")
	plex_brachGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Plexus brachial G", RelativeVolumes = [0.02])

	oreille_intDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne D", DoseType = "Average")
	oreille_intDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne D", RelativeVolumes = [0.02])

	oreille_intGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille interne G", DoseType = "Average")
	oreille_intG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille interne G", RelativeVolumes = [0.02])

	oreille_moyDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne D", DoseType = "Average")
	oreille_moyDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne D", RelativeVolumes = [0.02])

	oreille_moyGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oreille moyenne G", DoseType = "Average")
	oreille_moyG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oreille moyenne G", RelativeVolumes = [0.02])

	ATMDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM D", DoseType = "Average")
	ATMDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM D", RelativeVolumes = [0.02])

	ATMGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "ATM G", DoseType = "Average")
	ATMGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "ATM G", RelativeVolumes = [0.02])

	chiasmaAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Chiasma", DoseType = "Average")
	chiasmaD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Chiasma", RelativeVolumes = [0.02])

	cerveletAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cervelet", DoseType = "Average")
	cerveletD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cervelet", RelativeVolumes = [0.02])

	mandibuleAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Mandibule", DoseType = "Average")
	mandibuleD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Mandibule", RelativeVolumes = [0.02])
	
	nerf_opt_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique D", DoseType = "Average")
	nerf_opt_DG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique D", RelativeVolumes = [0.02])

	nerf_opt_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Nerf optique G", DoseType = "Average")
	nerf_opt_GG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Nerf optique G", RelativeVolumes = [0.02])

	cav_bucAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Cavité buccale", DoseType = "Average")
	cav_bucD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Cavité buccale", RelativeVolumes = [0.02])

	lobe_tempDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal D", DoseType = "Average")
	lobe_tempDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal D", RelativeVolumes = [0.02])

	lobe_tempGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Lobe temporal G", DoseType = "Average")
	lobe_tempG2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Lobe temporal G", RelativeVolumes = [0.02])

	sous_maxDAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max D", DoseType = "Average")
	sous_maxDD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max D", RelativeVolumes = [0.02])	

	sous_maxGAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Sous max G", DoseType = "Average")
	sous_maxGD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Sous max G", RelativeVolumes = [0.02])

	hypophyseAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Hypophyse", DoseType = "Average")
	hypophyseD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Hypophyse", RelativeVolumes = [0.02])

	oeil_DAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil D", DoseType = "Average")
	oeil_DD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil D", RelativeVolumes = [0.02])

	oeil_GAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Oeil G", DoseType = "Average")
	oeil_GD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Oeil G", RelativeVolumes = [0.02])

	cons_pharAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "Constricteurs du pharynx", DoseType = "Average")
	cons_pharD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "Constricteurs du pharynx", RelativeVolumes = [0.02])

	### Complete the listToExport ###
	listToExport.extend((\
		PTV_69D95, PTV_69Average, PTV_69D2, PTV_62195, PTV_621Average, PTV_621D2, PTV_5403D95, PTV_5403verage, PTV_5403D2, moelleD2, tronc_cerAverage, tronc_cerD2, \
		parotide_DValuesV22, parotide_DValuesV31, parotide_DValuesV35, parotide_DAverage, parotide_DD2, parotide_GValuesV22, parotide_GValuesV31, parotide_GValuesV35, parotide_GAverage, parotide_GD2, \
		parotide_UValuesV26, parotide_UAverage, parotide_UD2, larynxValuesV35, larynxValuesV46, thyroideValuesV50, thyroideAverage, thyroideD2, plex_brachDAverage, plex_brachD2, \
		plex_brachGAverage, plex_brachG2, oreille_intDValuesV46, oreille_intDAverage, oreille_intDD2, oreille_intGValuesV46, oreille_intGAverage, oreille_intGD2, \
		oreille_moyDAverage, oreille_moyDD2, oreille_moyGAverage, oreille_moyGD2, ATMDAverage, ATMDD2, ATMGAverage, ATMGD2, chiasmaAverage, chiasmaD2, cerveletAverage, cerveletD2, \
		mandibuleAverage, mandibuleD2, nerf_opt_DAverage, nerf_opt_DG2, nerf_opt_GAverage, nerf_opt_GG2, cav_bucValuesV15, cav_bucValuesV30, cav_bucValuesV35, cav_bucAverage, cav_bucD2, \
		lobe_tempDAverage, lobe_tempDD2, lobe_tempGAverage, lobe_tempG2, sous_maxDValuesV35, sous_maxDAverage, sous_maxD2, sous_maxGValuesV35, sous_maxGAverage, sous_maxG2, \
		hypophyseAverage, hypophyseD2, oeil_DValuesV40, oeil_DAverage, oeil_DD2, oeil_GValuesV40, oeil_GAverage, oeil_GD2, cons_pharAverage, cons_pharD2))


import codecs
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from connect import *

case = get_current("Case")
plan = get_current('Plan')
PTVD2 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 78Gy", RelativeVolumes = [0.02])
PTVD95 = plan.TreatmentCourse.TotalDose.GetDoseAtRelativeVolumes(RoiName = "PTV 78Gy", RelativeVolumes = [0.95])
PTVAverage = plan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName = "PTV 78Gy", DoseType = "Average")

PTVD2 = 0.5
PTVD95 = 5
PTVAverage = 444

listToExport = [12, 5, 78, 97, 11, 1]

savepath = "Q:/Aurelien_Dynalogs/Raystation_scripting/Raystation_Clinical_Goals.csv"
filesave = csv.writer(open(savepath, 'w', encoding='Latin-1'))
for elm in listToExport:
	filesave.writerow(str(elm))
filesave.close()








def ExportToCSV(listToExport, tumourLocalisation):
	savepath = "Q:/Aurelien_Dynalogs/Raystation_scripting/Raystation_Clinical_Goals.csv"
	filesave = csv.writer(open(savepath, 'w', encoding='Latin-1'))
	filesave.writerow(str(PTVD2) + "\n" + str(PTVD95) + "\n" + str(PTVAverage))
	filesave.close()
	messagebox.showinfo('Clinical Goals Export completed')