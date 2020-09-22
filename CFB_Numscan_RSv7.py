# -*- coding: utf-8 -*-

# Date: 25/02/2020
# The aim of this script is to find the ID and the number of the serie of CT or MRI, to show it, and to rename the imaging with these informations
# This script was based on "Numscan" script from Raysearchlab
# For more information : Aurélien Corroyer-Dulmont 5768 or a.corroyer-dulmont@baclesse.unicancer.fr


### UPDATE:
# 11/03/2020: acquisition date was added to the name of the CT
# 04/05/2020: information was added about initial and new CT
# 23/06/2020: four choices with WPF buttons: rename the CTs, rename the CTs and add "initial" and "reprise" information, add the date or check the CT number and ID which was used for the dosimetry
# 22/09/2020: add function (in numscan_and_serie function) to assign the density curve to the CTs based on informations from the dicom tags

import os, sys, System, clr, random, wpf
from math import *
from connect import *
clr.AddReferenceByName("PresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReferenceByName("PresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReference("System.Windows.Forms")
from System.Windows import *
from System.Windows.Forms import MessageBox


def Numscan_and_series(sender, args):
	""" Renomme tous les scanners présents sous la forme : CT:25103-3 """
	
	try:
		case=get_current('Case')
		date = []
		for elm in case.Examinations:
			if elm.EquipmentInfo.Modality == "CT":
				### To assign the density curve based on informations from dicom tag ###
				data = elm.GetAcquisitionDataFromDicom()
				CTmanufacturer = data['EquipmentModule']['Manufacturer']
				if CTmanufacturer == "Philips":
					elm.EquipmentInfo.SetImagingSystemReference(ImagingSystemName = 'bigboreCFB')
				if CTmanufacturer == "SIEMENS":
					elm.EquipmentInfo.SetImagingSystemReference(ImagingSystemName = 'Confidence')
				### To assign the density curve based on informations from dicom tag ###

				StudyID=elm.GetStoredDicomTagValueForVerification(Group=32,Element=16)
				Serie=elm.GetStoredDicomTagValueForVerification(Group=32,Element=17)
				NumStudyID=StudyID['Study ID']
				NumSerie=Serie['Series Number']
				elm.Name = 'CT:' + NumStudyID + '-' + NumSerie

			elif elm.EquipmentInfo.Modality == "MR":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				if date.Minute < 10:
					dateMinute = "0" + str(date.Minute)
				else:
					dateMinute = date.Minute
				if date.Hour < 10:
					dateHour = "0" + str(date.Hour)
				else:
					dateHour = date.Hour
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(dateHour) + ":" + str(dateMinute)
				elm.Name = 'MRI:' + dateTime_CT

			elif elm.EquipmentInfo.Modality == "PET":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(date.Hour) + ":" + str(date.Minute)
				elm.Name = 'PET:' + dateTime_CT

	except:
		print("There is no CT")



def Reprise_sim(sender, args):
	""" Renomme le scanner initial (avec la date la plus ancienne) et le nouveau scanner sous les formes respectives : CT:25103-3_INITIAL_20052020 et CT:25121-2_REPRISE_20062020 """

	try:
		case=get_current('Case')
		CTdate = []
		for elm in case.Examinations:
			if elm.EquipmentInfo.Modality == "CT":
				CTdate.append(elm.GetExaminationDateTime())
		for elm in case.Examinations:
			StudyID=elm.GetStoredDicomTagValueForVerification(Group=32,Element=16)
			Serie=elm.GetStoredDicomTagValueForVerification(Group=32,Element=17)
			date=elm.GetExaminationDateTime()
			if date.Month < 10:
				dateMonth = "0" + str(date.Month)
			else:
				dateMonth = date.Month
			if date.Day < 10:
				dateDay = "0" + str(date.Day)
			else:
				dateDay = date.Day
			dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year)
			NumStudyID=StudyID['Study ID']
			NumSerie=Serie['Series Number']
			if elm.GetExaminationDateTime() == min(CTdate) and elm.EquipmentInfo.Modality == "CT":
				elm.Name = 'CT:' + NumStudyID + '-' + NumSerie + '_INITIAL_' + dateTime_CT
			elif elm.GetExaminationDateTime() == max(CTdate) and elm.EquipmentInfo.Modality == "CT":
				elm.Name = 'CT:' + NumStudyID + '-' + NumSerie + '_REPRISE_' + dateTime_CT
			
			elif elm.EquipmentInfo.Modality == "MR":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				if date.Minute < 10:
					dateMinute = "0" + str(date.Minute)
				else:
					dateMinute = date.Minute
				if date.Hour < 10:
					dateHour = "0" + str(date.Hour)
				else:
					dateHour = date.Hour
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(dateHour) + ":" + str(dateMinute)
				elm.Name = 'MRI:' + dateTime_CT

			elif elm.EquipmentInfo.Modality == "PET":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(date.Hour) + ":" + str(date.Minute)
				elm.Name = 'PET:' + dateTime_CT

			else:
				StudyID=elm.GetStoredDicomTagValueForVerification(Group=32,Element=16)
				Serie=elm.GetStoredDicomTagValueForVerification(Group=32,Element=17)
				NumStudyID=StudyID['Study ID']
				NumSerie=Serie['Series Number']
				elm.Name = 'CT:' + NumStudyID + '-' + NumSerie

	except:
		print("There is no CT")


def Fusion(sender, args):
	""" Renomme tous les scanners sous la forme : CT:25103-3_20052020 """

	try:
		case=get_current('Case')
		for elm in case.Examinations:
			if elm.EquipmentInfo.Modality == "CT":
				StudyID=elm.GetStoredDicomTagValueForVerification(Group=32,Element=16)
				Serie=elm.GetStoredDicomTagValueForVerification(Group=32,Element=17)
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year)
				NumStudyID=StudyID['Study ID']
				NumSerie=Serie['Series Number']
				elm.Name = 'CT:' + NumStudyID + '-' + NumSerie + '_' + dateTime_CT
			
			elif elm.EquipmentInfo.Modality == "MR":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				if date.Minute < 10:
					dateMinute = "0" + str(date.Minute)
				else:
					dateMinute = date.Minute
				if date.Hour < 10:
					dateHour = "0" + str(date.Hour)
				else:
					dateHour = date.Hour
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(dateHour) + ":" + str(dateMinute)
				elm.Name = 'MRI:' + dateTime_CT

			elif elm.EquipmentInfo.Modality == "PET":
				date=elm.GetExaminationDateTime()
				if date.Month < 10:
					dateMonth = "0" + str(date.Month)
				else:
					dateMonth = date.Month
				if date.Day < 10:
					dateDay = "0" + str(date.Day)
				else:
					dateDay = date.Day
				dateTime_CT = str(dateDay) + str(dateMonth) + str(date.Year) + "_" + str(date.Hour) + ":" + str(date.Minute)
				elm.Name = 'PET:' + dateTime_CT
	except:
		print("There is no CT")


def Physic_validation(sender, args):
	""" To get information about which scan was used to performed dosimetry """

	try:
		case=get_current('Case')
		plan=get_current('Plan')
		BS=get_current('BeamSet')
		CT=BS.GetPlanningExamination()
		PlanName = str(plan.Name)
		StudyID=CT.GetStoredDicomTagValueForVerification(Group=32,Element=16)
		Serie=CT.GetStoredDicomTagValueForVerification(Group=32,Element=17)
		NumStudyID=StudyID['Study ID']
		NumSerie=Serie['Series Number']
		dosimetry_Exist = True
	except:
		dosimetry_Exist = False
		MessageBox.Show("There is no plan with that patient ")

	if dosimetry_Exist == True:
		MessageBox.Show("Plan: '" + str(PlanName) + "' was performed on CT:\n\nID: " + str(NumStudyID) + "\n\nSerie: " + str(NumSerie))


# Initialization Constants
Window = System.Windows.Window
Application = System.Windows.Application
Button = System.Windows.Controls.Button
StackPanel = System.Windows.Controls.StackPanel
Label = System.Windows.Controls.Label
Thickness = System.Windows.Thickness
DropShadowBitmapEffect = System.Windows.Media.Effects.DropShadowBitmapEffect


# Create window
my_window = Window()
my_window.Title = 'Please choose a protocole:'
my_window.Width = 450
my_window.Height = 155

# Create StackPanel to Layout UI elements 
my_stack = StackPanel()
my_stack.Margin = Thickness(15)
my_window.Content = my_stack

# Create Button and add a Button Click event handler
my_button1 = Button()
my_button1.Content = 'Numscan and series'
my_button1.FontSize = 12
my_button1.BitmapEffect = DropShadowBitmapEffect()
my_button1.Click += Numscan_and_series
my_stack.Children.Add (my_button1)

my_button2 = Button()
my_button2.Content = 'Reprise sim'
my_button2.FontSize = 12
my_button2.BitmapEffect = DropShadowBitmapEffect()
my_button2.Click += Reprise_sim
my_stack.Children.Add (my_button2)

my_button3 = Button()
my_button3.Content = 'Fusion'
my_button3.FontSize = 12
my_button3.BitmapEffect = DropShadowBitmapEffect()
my_button3.Click += Fusion
my_stack.Children.Add (my_button3)

my_button4 = Button()
my_button4.Content = 'Physic validation'
my_button4.FontSize = 12
my_button4.BitmapEffect = DropShadowBitmapEffect()
my_button4.Click += Physic_validation
my_stack.Children.Add (my_button4)

# Run application
my_app = Application()
my_app.Run (my_window)







