# -*- coding: utf-8 -*-

# Script to obtain DSP of 2 differents beams (one at the zero position of the CT)
# Author : Aurélien Corroyer-Dulmont
# Version : 08 march 2021

# Update xx/xx/2020 : 


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
plan.BeamSets[0].Beams[0].Name = r"DSP0"
plan.BeamSets[0].Beams[1].Name = r"DSPttt"
plan.BeamSets[0].Beams[0].GantryAngle = 0
plan.BeamSets[0].Beams[1].GantryAngle = 0

## To change the isocenter position
plan.BeamSets[0].Beams[0].Isocenter.EditIsocenter(Name=r"DSP0", Color="98, 184, 234", Position={ 'x': locX, 'y': locY, 'z': locZ })

## To get the DSP of the DSP0 beam
dsp_beamDSP0 = plan.BeamSets[0].Beams[0].GetSSD()



### Creation of a plan/beamset/beam with zero CT coordinate (only for x)
case.CopyPlan(PlanName=Originalplan.Name, NewPlanName=r"DSPttt")
plan = case.TreatmentPlans["DSPttt"]
plan.BeamSets[0].Beams[0].Name = r"DSP0"
plan.BeamSets[0].Beams[1].Name = r"DSPttt"
plan.BeamSets[0].Beams[0].GantryAngle = 0
plan.BeamSets[0].Beams[1].GantryAngle = 0

## To keep the y and z position which will not change
LocY_beamDSPttt = plan.BeamSets[0].Beams[1].Isocenter.Position.y
LocZ_beamDSPttt = plan.BeamSets[0].Beams[1].Isocenter.Position.z

## To change the isocenter position
plan.BeamSets[0].Beams[1].Isocenter.EditIsocenter(Name=r"DSPttt", Color="98, 184, 234", Position={ 'x': locX, 'y': LocY_beamDSPttt, 'z': LocZ_beamDSPttt })

## To get the DSP of the DSPttt beam
dsp_beamDSPttt = plan.BeamSets[0].Beams[1].GetSSD()



## To show the difference
import System, clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
MessageBox.Show("DSP0 : " + str(round(dsp_beamDSP0,2)) + " cm\n\nDSP traitement : " + str(round(dsp_beamDSPttt,2)) + " cm")