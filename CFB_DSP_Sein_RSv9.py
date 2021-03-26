# -*- coding: utf-8 -*-

# Script to obtain DSP of 2 differents beams (one at the zero position of the CT)
# Author : Aurélien Corroyer-Dulmont
# Version : 08 march 2021

# Update xx/xx/2021 : 


from connect import *

examination = get_current("Examination")
case = get_current("Case")
Originalplan = get_current("Plan")
beam_set = get_current("BeamSet")


### To obtain the zero position of the CT
locPosition = case.PatientModel.StructureSets[examination.Name].LocalizationPoiGeometry.Point
locX = locPosition.x
locY = locPosition.y
locZ = locPosition.z


### Creation of a plan/beamset/beam with zero CT coordinate
case.CopyPlan(PlanName=Originalplan.Name, NewPlanName=r"DSP0")
plan = case.TreatmentPlans["DSP0"]

### To get the selected beamset
for elm in plan.BeamSets:
	if elm.Comment == beam_set.Comment:
		beamSetOfChoiceDSP0 = elm
	else:
		print("not that BeamSet")

beamSetOfChoiceDSP0.Beams[0].Name = r"Beam1_DSP0"
beamSetOfChoiceDSP0.Beams[1].Name = r"Beam2_DSP0"
beamSetOfChoiceDSP0.Beams[0].GantryAngle = 0
beamSetOfChoiceDSP0.Beams[1].GantryAngle = 0

## To change the isocenter position
beamSetOfChoiceDSP0.Beams[0].Isocenter.EditIsocenter(Name=r"Beam1_DSP0", Color="98, 184, 234", Position={ 'x': locX, 'y': locY, 'z': locZ })

## To get the DSP of the DSP0 beam
dsp_beamDSP0 = beamSetOfChoiceDSP0.Beams[0].GetSSD()


### Creation of a plan/beamset/beam with zero CT coordinate (only for x)
case.CopyPlan(PlanName=Originalplan.Name, NewPlanName=r"DSPttt")
plan = case.TreatmentPlans["DSPttt"]

### To get the selected beamset
for elm in plan.BeamSets:
	if elm.Comment == beam_set.Comment:
		beamSetOfChoiceDSPttt = elm
	else:
		print("not that BeamSet")


beamSetOfChoiceDSPttt.Beams[0].Name = r"Beam1_DSPttt"
beamSetOfChoiceDSPttt.Beams[1].Name = r"Beam2_DSPttt"
beamSetOfChoiceDSPttt.Beams[0].GantryAngle = 0
beamSetOfChoiceDSPttt.Beams[1].GantryAngle = 0

## To keep the y and z position which will not change
LocY_beamDSPttt = beamSetOfChoiceDSPttt.Beams[1].Isocenter.Position.y
LocZ_beamDSPttt = beamSetOfChoiceDSPttt.Beams[1].Isocenter.Position.z

## To change the isocenter position
beamSetOfChoiceDSPttt.Beams[1].Isocenter.EditIsocenter(Name=r"Beam2_DSPttt", Color="98, 184, 234", Position={ 'x': locX, 'y': LocY_beamDSPttt, 'z': LocZ_beamDSPttt })

## To get the DSP of the DSPttt beam
dsp_beamDSPttt = beamSetOfChoiceDSPttt.Beams[1].GetSSD()



## To show the difference
import System, clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
MessageBox.Show("DSP0 : " + str(round(dsp_beamDSP0,2)) + " cm\n\nDSP traitement : " + str(round(dsp_beamDSPttt,2)) + " cm")