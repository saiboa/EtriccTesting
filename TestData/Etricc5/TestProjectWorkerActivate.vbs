'-------------------------------
'Project Eurobaltic
'-------------------------------
Dim project as Egemin.EPIA.WCS.Core.Project = DirectCast(Root, Egemin.EPIA.WCS.Core.Project)


'//////// ====  ADD APPLICATION ====  ////////////////////
'---------------------------------------------------------
'Declare application ID and application variables
'---------------------------------------------------------
Dim applicationID As String = "APPLICATION1"
Dim application As Egemin.EPIA.Core.Application

'---------------------------------------------------------
'Set file path, mode and parameters for Worker Application
'---------------------------------------------------------
'application = project.Applications(applicationID)
'If Not application Is Nothing Then
'    project.Applications.Remove(application)
'End If
'application = New Egemin.EPIA.Core.Application()
'application.ID = applicationID
'application.Mode = application.MODE.LOCAL
'application.Parameters("WorkerFilePath").Value = "C:\EpiaTestCenter2\AutoTestCenter\Main\Source\TestRuns\bin\FX1_1'\Debug\TestRuns.dll"
'application.Parameters("WorkerAssembly").Value = "Worker"
'application.Parameters("WorkerType").Value = "TestRuns.TestWorker"
'project.Applications.Add(application)



'///  ====  ADD FACILITIES ====  ///
'-----------------------------------
'Initialize Facilities
'-----------------------------------
'Dim TestsFacility as New Egemin.EPIA.Core.Definitions.Facility()

'Dim CurrentTestIDParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'CurrentTestIDParameter.Type = CurrentTestIDParameter.TYPE.INT

'Dim LogMessageParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'LogMessageParameter.Type = LogMessageParameter.TYPE.STRING

'Dim RunStatusParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'RunStatusParameter.Type = RunStatusParameter.TYPE.STRING

'Dim TestCenterRootParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'TestCenterRootParameter.Type = TestCenterRootParameter.TYPE.STRING

'Dim TestIDParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'TestIDParameter.Type = TestIDParameter.TYPE.INT

'Dim TestTitleParameter as New Egemin.EPIA.Core.Definitions.Parameter()
'TestTitleParameter.Type = TestTitleParameter.TYPE.STRING

'TestsFacility.ID = "Tests"

'Dim facility As Egemin.EPIA.Core.Definitions.Facility = Project.Facilities( TestsFacility.ID)

'If facility Is Nothing Then
'    Project.Facilities.Add(TestsFacility)
'
'    CurrentTestIDParameter.ID = "CurrentTestID"
'   CurrentTestIDParameter.ValueAsInt = 0
'
'    LogMessageParameter.ID = "LogMessage"
'    LogMessageParameter.ValueAsString = "msg"
'
'    RunStatusParameter.ID = "RunStatus"
'    RunStatusParameter.ValueAsString = "StartRunning"
'
'    TestCenterRootParameter.ID = "TestCenterRoot"
'    TestCenterRootParameter.ValueAsString = "C:\EpiaTestCenter2\AutoTestCenter\Main"

'    TestIDParameter.ID = "TestID"
'    TestIDParameter.ValueAsInt = 0

'    TestTitleParameter.ID = "TestTitleArray"
'    TestTitleParameter.ValueAsString = "All,"

'    TestsFacility.Parameters.Clear()
'    TestsFacility.Parameters.Insert(CurrentTestIDParameter)
'    TestsFacility.Parameters.Insert(LogMessageParameter)
'    TestsFacility.Parameters.Insert(RunStatusParameter)
'    TestsFacility.Parameters.Insert(TestCenterRootParameter)
'    TestsFacility.Parameters.Insert(TestIDParameter)
'    TestsFacility.Parameters.Insert(TestTitleParameter)

'End If

'///  ====  ADD Schedules ====  ///
'-------------------------------
'Add Schedules and Rules
'-------------------------------
 Dim schedule As New Egemin.EPIA.WCS.Scheduling.Schedule()
 Dim rule As New Egemin.EPIA.WCS.Scheduling.Rule()

'The following code provides an example!
'It shows how to create pick solving schedules

'-------------------------------
'0040-01-05 for test rule PICK delay
'-------------------------------
'schedule = project.Schedules("0040-01-05.PICK")
'If Not schedule Is Nothing Then
' project.Schedules.Remove( schedule )
'End If
'schedule = New Egemin.EPIA.WCS.Scheduling.Schedule( "0040-01-05", schedule.TYPE.PICK )
'rule = New Egemin.EPIA.WCS.Scheduling.Rule()
'rule.Type = rule.TYPE.DELAY
'rule.Arguments.Add( "70" )
'schedule.Rules.Add( rule )

'rule = New Egemin.EPIA.WCS.Scheduling.Rule()
'rule.Type = rule.TYPE.CLOSEST_HIGHEST
'rule.Arguments.Add( "* 0 0" )
'schedule.Rules.Add( rule )

'project.Schedules.Add( schedule )


'-------------------------------
'AREA30 for test rule DEADLOCK VIA
'-------------------------------
'schedule = project.Schedules("AREA_LAYOUT_FLV.DEADLOCK")
'If Not schedule Is Nothing Then
' project.Schedules.Remove( schedule )
'End If
'schedule = New Egemin.EPIA.WCS.Scheduling.Schedule( "AREA_LAYOUT_FLV", schedule.TYPE.DEADLOCK )
'rule = New Egemin.EPIA.WCS.Scheduling.Rule()
'rule.Type = rule.TYPE.VIA
'rule.Arguments.Add( "X030" )
'schedule.Rules.Add( rule )

'project.Schedules.Add( schedule )

'///  ====  ACTIVATE ====  ///
'-------------------------------
'Activate Project
'-------------------------------
project.Activate()