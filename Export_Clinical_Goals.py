# -*- coding: utf-8 -*-

# Export the clinical goals to excel or csv files
# Clinical goals reading is based on "print goal" function from Mark Geurts; Github : https://github.com/wrssc/ray_scripts/blob/master/library/Goals.py
# Author : Aurélien Corroyer-Dulmont
# Version : 17 june 2020

# Update xx/xx/2020 : 



########################### TO DO #############################################################################################################################################
### faire une liste déroulante où l'utilisateur choisit le type de dosimétrie/localisation tumoral, reprendre cette info pour savoir dans quel onglet stocker les résultats ###





import pandas as pad 
from openpyxl import load_workbook
import win32com.client
from path import Path
import shutil
import os
import codecs

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
		poumons_gaucheAverage, poumons_gaucheD2, sein_controValuesV2, sein_controAverage,sein_controD2, ThyroideAverage, ThyroideD2, FoieAverage, FoieD2, Tete_huméraleAverage, Tete_huméraleD2))


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
		PTVD95, PTVAverage, PTVD2, paroiVesicaleValuesV65, paroiVesicaleValuesV70, paroiVesicaleValuesV80, paroiVesicaleD2, \
		greleValuesV40, greleAverage,greleD2, colonAverage, colonD2, TFDValuesV50, TFDValuesV55, TFDAverage, \
		TFD2, TFGValuesV50, TFGValuesV55, TFGAverage, TFG2, plexusSacreAverage, plexusSacreD2, moelleValuesV10, moelleValuesV20, moelleD25, moelleAverage, moelleD2, \
		QueueDeChevalAverage, QueueDeChevalD2, Org_genValuesV20, Org_genValuesV30, Org_genValuesV40, Org_genAverage, Org_genD2))











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



def ExportToCSV(listToExport, tumourLocalisation)
	savepath = "Q:/Aurelien_Dynalogs/Raystation_scripting/Raystation_scripting.csv"
	filesave = codecs.open(savepath, 'w', encoding='Latin-1')
	filesave.write(str(PTVD2) + "\n" + str(PTVD95) + "\n" + str(PTVAverage))
	filesave.close()
	messagebox.showinfo('Clinical Goals Export completed')