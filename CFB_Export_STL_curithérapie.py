# -*- coding: utf-8 -*-

# Script to export STL file from raystation ROI
# script based on publication: https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6414136/
# Author : Aurélien Corroyer-Dulmont
# Version : 13 july 2020

# Update xx/xx/2020 : 

from connect import *
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application, Form, Label, ComboBox, Button
from System.Drawing import Point, Size



#Define Forms class
class SelectROIForm(Form):

	def __init__(self):
		# Set the size of the form
		self.Size = Size(500, 200)
		# Set title of the form
		self.Text = 'Export to STL format'

		# Add a label
		label = Label()
		label.Text = 'Chose the ROI to export in STL format:'
		label.Location = Point(15, 15)
		label.AutoSize = True
		self.Controls.Add(label)

		#Add Combobox
		case = get_current("Case")
		roi_names = [r.Name for r in case.PatientModel.RegionsOfInterest]
		self.Combobox = ComboBox()
		self.Combobox.DataSource = roi_names
		self.Combobox.Location = Point(15, 55)
		self.Combobox.Size = Size(300, 20)
		self.Controls.Add(self.Combobox)

		# Add button to press OK and close the form
		button = Button()
		button.Text = 'OK'
		button.AutoSize = True
		button.Location = Point(15, 95)
		button.Click += self.ok_button_clicked
		self.Controls.Add(button)

	def ok_button_clicked(self, sender, event):
		# Method invoked when the button is clicked
		patient = get_current('Patient')
		case = get_current("Case")
		examination = get_current("Examination")
		case.PatientModel.StructureSets[examination.Name].RoiGeometries[str(self.Combobox.SelectedValue)].ExportRoiGeometryAsSTL(\
			DestinationDirectory="xxx", OutputUnit ='Millimeter')

		# Close the form
		self.Close()
		
# Create an instance of the form and run it
form = SelectROIForm()
Application.Run(form)