Sub Main()
Dim studio As Object
 
'Starts CST MICROWAVE STUDIO
Set studio = CreateObject("CSTStudio.Application")
 
Dim proj As Object
Set proj = studio.Active3D

Dim Text As String, textline As String, posLat As Integer, posLong As Integer

myFile = "C:\Users\konstantin.grotov\Documents\CST_python\input.txt"

Open myFile For Input As #1
Dim t As String
Dim route As String, output As String
Do Until EOF(1)
    Line Input #1, textline
    route = Split(textline)(0)
    output = Split(textline)(1)
    N = CInt(Split(textline)(2))
    'route = data[0]
Loop
Close #1

'route  = "C:\Users\konstantin.grotov\Documents\MOM_Antenna-main\MOM_Antenna-main\data\cubic_geometry\lengths_for_cst\TO_CST_CUBIC_Lfrom10to24_N3_tau6_iter100_runsmax10_seed52.txt"
'N = 3
is_cubic_geometry = True

StoreDoubleParameter "r", 0.5
SetParameterDescription  ( "r",  "wire radius"  )


Component.New "Geometry"
Dim tau As Double
Dim i,j,l As Integer

Dim lengths(100,100) As Double
Dim x(100,100) As Double
Dim y(100,100) As Double

Dim length_x_y() As String

myFile = route
Open myFile For Input As #1
'Dim t As String
If is_cubic_geometry Then
Do Until EOF(1)
    Line Input #1, textline
    length_x_y = Split(textline)
	lengths((i \ N),j Mod N) = CDbl(length_x_y(0))
	x((i \ N),j Mod N) = CDbl(length_x_y(1))
	y((i \ N),j Mod N) = CDbl(length_x_y(2))
	i += 1
	j += 1
Loop
Close #1

AddToHistory("make planewave", "With PlaneWave" & vbCrLf & _
     ".Reset" & vbCrLf & _
     ".Normal ""-1"", ""0"", ""0""" & vbCrLf & _
     ".EVector ""0"", ""1"", ""0""" & vbCrLf & _
     ".Polarization ""Linear""" & vbCrLf & _
     ".ReferenceFrequency ""0""" & vbCrLf & _
     ".PhaseDifference ""-90.0""" & vbCrLf & _
     ".CircularDirection ""Left""" & vbCrLf & _
     ".AxialRatio ""0.0""" & vbCrLf & _
     ".SetUserDecouplingPlane ""False""" & vbCrLf & _
     ".Store" & vbCrLf & _
"End With")

AddToHistory("make monitors", "With Monitor" & vbCrLf & _
          ".Reset" & vbCrLf & _
          ".Domain ""Frequency""" & vbCrLf & _
          ".FieldType ""Farfield""" & vbCrLf & _
          ".ExportFarfieldSource ""False""" & vbCrLf & _
          ".UseSubvolume ""False""" & vbCrLf & _
          ".Coordinates ""Structure""" & vbCrLf & _
          ".SetSubvolume ""-30.25"", ""30.25"", ""-70.5"", ""70.5"", ""-30.25"", ""30.25""" & vbCrLf & _
          ".SetSubvolumeOffset ""10"", ""10"", ""10"", ""10"", ""10"", ""10""" & vbCrLf & _
          ".SetSubvolumeInflateWithOffset ""False""" & vbCrLf & _
          ".SetSubvolumeOffsetType ""FractionOfWavelength""" & vbCrLf & _
          ".EnableNearfieldCalculation ""True""" & vbCrLf & _
          ".CreateUsingLinearSamples ""0.8"", ""1.2"", ""20""" & vbCrLf & _
"End With")

AddToHistory("make probe","With Probe" & vbCrLf & _
     ".Reset"  & vbCrLf & _
     ".ID 0" & vbCrLf & _
     ".AutoLabel 1"  & vbCrLf & _
     ".Field ""RCS"""  & vbCrLf & _
     ".Orientation ""All"""  & vbCrLf & _
     ".SetPosition1 ""-10"""  & vbCrLf & _
     ".SetPosition2 ""0.0"""  & vbCrLf & _
     ".SetPosition3 ""0.0"""   & vbCrLf & _
     ".SetCoordinateSystemType ""Cartesian"""  & vbCrLf & _
     ".Create"  & vbCrLf & _
"End With")

t = ""
For i=0 To N - 1
	For j=0 To N - 1
	t = t & "With Cylinder" & vbCrLf
	     t = t & ".Reset" & vbCrLf
	     t = t & ".Name ""c" & CStr(i) & CStr(j) & """" & vbCrLf
	     t = t & ".Component ""Geometry""" & vbCrLf
	     t = t & ".Material ""PEC""" & vbCrLf
	     t = t & ".OuterRadius ""r""" & vbCrLf
	     t = t & ".InnerRadius ""0""" & vbCrLf
	     t = t & ".Axis ""y""" & vbCrLf
	     t = t & ".Yrange " & CStr(-lengths(i,j)/2) & ", " &  CStr(lengths(i,j)/2) & vbCrLf
	     t = t & ".Xcenter " & CStr(y(i, j)) & vbCrLf
	     t = t & ".Zcenter " & CStr(x(i, j)) & vbCrLf
	     t = t & ".Segments ""0""" & vbCrLf
	     t = t & ".Create" & vbCrLf
	t = t & "End With" & vbCrLf
	Next
Next
	AddToHistory("make cubic geometry", t)
Else
Do Until EOF(1)
    Line Input #1, textline
    length_x_y = Split(textline)
	lengths((i,0) = CDbl(length_x_y(0))
	x(i,0) = CDbl(length_x_y(1))
	y(i,0) = CDbl(length_x_y(2))
	i += 1
Loop
Close #1

t = ""
For i=0 To N - 1
	t = t & "With Cylinder" & vbCrLf & _
	     ".Reset" & vbCrLf & _
	     ".Name ""c" & CStr(i) & """" & vbCrLf & _
	     ".Component ""Geometry""" & vbCrLf & _
	     ".Material ""PEC""" & vbCrLf & _
	     ".OuterRadius ""r""" & vbCrLf & _
	     ".InnerRadius ""0""" & vbCrLf & _
	     ".Axis ""y""" & vbCrLf & _
	     ".Yrange " & CStr(-lengths(i,0)/2) & ", " &  CStr(lengths(i,0)/2) & vbCrLf & _
	     ".Xcenter " & CStr(y(i, 0)) & vbCrLf & _
	     ".Zcenter " & CStr(x(i, 0)) & vbCrLf & _
	     ".Segments ""0""" & vbCrLf & _
	     ".Create" & vbCrLf & _
	"End With" & vbCrLf
Next
AddToHistory("make circular geometry", t)
End If

	Dim sCommand As String

	'@ define units
	sCommand = ""
	sCommand = sCommand + "With Units " + vbLf
	sCommand = sCommand + ".Geometry ""mm""" + vbLf
	sCommand = sCommand + ".Frequency ""GHz""" + vbLf
	sCommand = sCommand + ".Time ""ns""" + vbLf
	sCommand = sCommand + "End With"
	AddToHistory "define units", sCommand

AddToHistory("Set frequency Solver", "ChangeSolverType ""HF Frequency Domain""")
AddToHistory("Set frequency range", "Solver.FrequencyRange ""2"", ""12""")
Mesh.SetCreator "High Frequency" 

With FDSolver
     .Reset 
     .SetMethod "Tetrahedral", "General purpose" 
     .OrderTet "Second" 
     .OrderSrf "First" 
     .Stimulation "Plane Wave", "1" 
     .ResetExcitationList 
     .AutoNormImpedance "False" 
     .NormingImpedance "50" 
     .ModesOnly "False" 
     .ConsiderPortLossesTet "True" 
     .SetShieldAllPorts "False" 
     .AccuracyHex "1e-6" 
     .AccuracyTet "1e-4" 
     .AccuracySrf "1e-3" 
     .LimitIterations "False" 
     .MaxIterations "0" 
     .SetCalcBlockExcitationsInParallel "True", "True", "" 
     .StoreAllResults "False" 
     .StoreResultsInCache "False" 
     .UseHelmholtzEquation "True" 
     .LowFrequencyStabilization "True" 
     .Type "Auto" 
     .MeshAdaptionHex "False" 
     .MeshAdaptionTet "False" 
     .AcceleratedRestart "True" 
     .FreqDistAdaptMode "Distributed" 
     .NewIterativeSolver "True" 
     .TDCompatibleMaterials "False" 
     .ExtrudeOpenBC "False" 
     .SetOpenBCTypeHex "Default" 
     .SetOpenBCTypeTet "Default" 
     .AddMonitorSamples "True" 
     .CalcPowerLoss "True" 
     .CalcPowerLossPerComponent "False" 
     .StoreSolutionCoefficients "True" 
     .UseDoublePrecision "False" 
     .UseDoublePrecision_ML "True" 
     .MixedOrderSrf "False" 
     .MixedOrderTet "False" 
     .PreconditionerAccuracyIntEq "0.15" 
     .MLFMMAccuracy "Default" 
     .MinMLFMMBoxSize "0.3" 
     .UseCFIEForCPECIntEq "True" 
     .UseFastRCSSweepIntEq "false" 
     .UseSensitivityAnalysis "False" 
     .RemoveAllStopCriteria "Hex"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
     .RemoveAllStopCriteria "Tet"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
     .RemoveAllStopCriteria "Srf"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
     .SweepMinimumSamples "3" 
     .SetNumberOfResultDataSamples "1001" 
     .SetResultDataSamplingMode "Automatic" 
     .SweepWeightEvanescent "1.0" 
     .AccuracyROM "1e-4" 
     .AddSampleInterval "", "", "4", "Automatic", "True" 
     .AddSampleInterval "", "", "", "Automatic", "False" 
     .MPIParallelization "False"
     .UseDistributedComputing "False"
     .NetworkComputingStrategy "RunRemote"
     .NetworkComputingJobCount "3"
     .UseParallelization "True"
     .MaxCPUs "1024"
     .MaximumNumberOfCPUDevices "2"
End With

With IESolver
     .Reset 
     .UseFastFrequencySweep "True" 
     .UseIEGroundPlane "False" 
     .SetRealGroundMaterialName "" 
     .CalcFarFieldInRealGround "False" 
     .RealGroundModelType "Auto" 
     .PreconditionerType "Auto" 
     .ExtendThinWireModelByWireNubs "False" 
     .ExtraPreconditioning "False" 
End With

With IESolver
     .SetFMMFFCalcStopLevel "0" 
     .SetFMMFFCalcNumInterpPoints "6" 
     .UseFMMFarfieldCalc "True" 
     .SetCFIEAlpha "0.500000" 
     .LowFrequencyStabilization "False" 
     .LowFrequencyStabilizationML "True" 
     .Multilayer "False" 
     .SetiMoMACC_I "0.0001" 
     .SetiMoMACC_M "0.0001" 
     .DeembedExternalPorts "True" 
     .SetOpenBC_XY "True" 
     .OldRCSSweepDefintion "False" 
     .SetRCSOptimizationProperties "True", "100", "0.00001" 
     .SetAccuracySetting "Custom" 
     .CalculateSParaforFieldsources "True" 
     .ModeTrackingCMA "True" 
     .NumberOfModesCMA "3" 
     .StartFrequencyCMA "-1.0" 
     .SetAccuracySettingCMA "Default" 
     .FrequencySamplesCMA "0" 
     .SetMemSettingCMA "Auto" 
     .CalculateModalWeightingCoefficientsCMA "True" 
End With


Debug.Print "End"

proj.FDSolver.Start

SelectTreeItem "1D Results\Probes\RCS\RCS (Cartesian) (-10 0 0)(Abs) [pw]"
'SelectTreeItem("1D Results\Probes\RCS\")
With Plot1D
.Plotview("magnitude")
End With
ExportPlotData output


proj.SaveAs "C:\Users\konstantin.grotov\Documents\tratra.cst", True
proj.Quit
 
Set studio = Nothing
End Sub
