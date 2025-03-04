Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.cosworks
Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Timers

Module Program

    Sub Main()
        ' define shape def excel, set up excel and timer
        Dim strFileName As String = "Z:\StudentFolders\Haider\Projects\Optimization\SolidWorks Interfacing Optimizer\tempfolder\Shape Definition.xlsx"
        Dim xlApp As Excel.Application = New Excel.Application
        Dim timer As Timer = New Timer(180000)
        AddHandler timer.Elapsed, New ElapsedEventHandler(AddressOf TimerElapsed)
        timer.Start()

        ' dimension errors and SW managers
        Dim model As ModelDoc2
        Dim sketchMgr As SketchManager
        Dim featureMgr As FeatureManager
        Dim errCode As Integer
        Dim warnCode As Integer
        Dim errorCode As Integer

        ' launch SW and load cosmos
        app = LaunchSW()
        Const sAddinName As String = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\Simulation\cosworks.dll"
        Dim status As Integer = app.LoadAddIn(sAddinName)

        'make new part, close and reopen in silent
        model = app.NewPart()
        Dim savestatus As Boolean = model.SaveAs3("C:\Users\HuRo Lab.BMEG-DHBX74Z2\Desktop\SolidWorks Temp Sims\current.SLDPRT", 0, 1)
        app.CloseAllDocuments(True) 'Closing all documents without save
        model = app.OpenDoc6("C:\Users\HuRo Lab.BMEG-DHBX74Z2\Desktop\SolidWorks Temp Sims\current.SLDPRT", 1, 1, "", errCode, warnCode)

        sketchMgr = model.SketchManager
        featureMgr = model.FeatureManager

        'enable cosmos in active doc
        Dim ActDoc As CWModelDoc = SetupCosmos(app)

        'import points from excel 
        Dim TOPpoints As Double(,) = PointsFromXL(xlApp, strFileName, "TOP")
        Dim BOTpoints As Double(,) = PointsFromXL(xlApp, strFileName, "BOT")
        Dim Ends As Double(,) = PointsFromXL(xlApp, strFileName, "Ends")
        Dim Base(0, 2) As Double
        Dim Edge(0, 2) As Double
        Base(0, 0) = Ends(0, 0)
        Base(0, 1) = Ends(0, 1)
        Base(0, 2) = Ends(0, 2)
        Edge(0, 0) = Ends(1, 0)
        Edge(0, 1) = Ends(1, 1)
        Edge(0, 2) = Ends(1, 2)
        Dim extrude As Double = ParameterFromXL(xlApp, strFileName, "Parameters", "A1")

        ' sketch contour
        model.Insert3DSketch2(False)
        Dim nPointsTOP As Integer = CreateSurface(featureMgr, model, TOPpoints)
        Dim nPointsBOT As Integer = CreateSurface(featureMgr, model, BOTpoints)
        CreateEnds(model, TOPpoints, BOTpoints, nPointsTOP, nPointsBOT)
        model.Insert3DSketch2(True)

        ' extrude sketch and add axis of rotation
        ExtrudeSketch(model, featureMgr, TOPpoints, extrude)
        CreateAxisofRotation(model)

        Dim skPoint As SketchPoint
        model.Insert3DSketch2(False)
        skPoint = sketchMgr.CreatePoint(Edge(0, 0), Edge(0, 1), extrude / 2) ' 
        model.InsertSketch2(True)

        ' create study
        Dim Study As CWStudy = CreateNonLinearStudy(ActDoc, errCode, errorCode)

        ' identify entities of interest
        Dim base_entity(0) As Object
        Dim rot_entity(0) As Object
        Dim inplane1_entity(0) As Object
        Dim inplane2_entity(0) As Object

        base_entity(0) = IdentifyEntity(model, "FACE", Base(0, 0), Base(0, 1), extrude / 2, errCode)
        Dim TOPEdge As Object = IdentifyEntity(model, "EDGE", TOPpoints(2, 0), TOPpoints(2, 1), extrude, errCode)
        Dim BOTEdge As Object = IdentifyEntity(model, "EDGE", BOTpoints(2, 0), BOTpoints(2, 1), extrude, errCode)
        Dim Axis As Object = IdentifyEntity(model, "AXIS", 0, 0, 0, errCode)
        Dim boolstatus As Boolean = model.Extension.SelectByID2("", "FACE", Edge(0, 0), Edge(0, 1), extrude / 2, False, 0, Nothing, 0)
        rot_entity(0) = IdentifyEntity(model, "FACE", Edge(0, 0), Edge(0, 1), extrude / 2, errCode)
        inplane1_entity(0) = IdentifyEntity(model, "FACE", (TOPpoints(1, 0) + BOTpoints(1, 0)) / 2, (TOPpoints(1, 1) + BOTpoints(1, 1)) / 2, 0, errCode)
        inplane2_entity(0) = IdentifyEntity(model, "FACE", (TOPpoints(1, 0) + BOTpoints(1, 0)) / 2, (TOPpoints(1, 1) + BOTpoints(1, 1)) / 2, extrude, errCode)

        Dim LBCMgr As CWLoadsAndRestraintsManager = Study.LoadsAndRestraintsManager
        If LBCMgr Is Nothing Then ErrorMsg(app, "Failed to get loads and restraints manager", True)

        'Add a restraint
        Dim CWFeatObj3 As CWRestraint
        CWFeatObj3 = LBCMgr.AddRestraint(0, (base_entity), Axis, errCode)
        'If errCode <> 0 Then ErrorMsg(app, "Failed to create restraint", True)

        Dim CWFeatObj4 As CWRestraint
        Dim displacement(5) As Object

        displacement(0) = 0.0#
        displacement(1) = 180
        displacement(2) = 0.0#
        displacement(3) = 1
        displacement(4) = 1
        displacement(5) = 0
        CWFeatObj4 = LBCMgr.AddPrescribedDisplacement(displacement, 3, (rot_entity), Axis, errCode)
        'If errCode <> 0 Then ErrorMsg(app, "Failed to create restraint", True)

        Dim CWFeatObj5 As CWRestraint
        Dim displacement_inplane1(5) As Object

        displacement_inplane1(0) = 0.0#
        displacement_inplane1(1) = 0.0#
        displacement_inplane1(2) = 0.0#
        displacement_inplane1(3) = 0
        displacement_inplane1(4) = 0
        displacement_inplane1(5) = 1
        CWFeatObj5 = LBCMgr.AddPrescribedDisplacement(displacement_inplane1, 3, (inplane1_entity), Axis, errCode)

        Dim CWFeatObj6 As CWRestraint
        Dim displacement_inplane2(5) As Object

        displacement_inplane2(0) = 0.0#
        displacement_inplane2(1) = 0.0#
        displacement_inplane2(2) = 0.0#
        displacement_inplane2(3) = 0
        displacement_inplane2(4) = 0
        displacement_inplane2(5) = 1
        CWFeatObj6 = LBCMgr.AddPrescribedDisplacement(displacement_inplane2, 3, (inplane2_entity), Axis, errCode)

        'Create mesh
        Dim Mesh As CWMesh
        Mesh = Study.Mesh
        If Mesh Is Nothing Then ErrorMsg(app, "Failed to create mesh object", True)
        Mesh.MesherType = 0
        Mesh.Quality = 1


        Const MeshEleSize As Double = 15 'mm

        errCode = Study.CreateMesh(0, MeshEleSize, Nothing)
        If errCode <> 0 Then ErrorMsg(app, "Failed to create mesh", True)

        Dim num As Integer
        Dim idList As Object
        Dim normalNum As Object = Nothing
        Dim normalVec As Object = Nothing
        Dim ncount As Integer
        Dim node_base As Object = Nothing
        Dim node_TOP As Object = Nothing
        Dim node_BOT As Object = Nothing
        Dim nodes As Object

        num = Mesh.GetSurfaceNodesAndNormals(idList, normalNum, normalVec)
        node_base = Mesh.GetNodeDataFromEntity(base_entity(0), ncount)
        node_TOP = Mesh.GetNodeDataFromEntity(TOPEdge, ncount)
        node_BOT = Mesh.GetNodeDataFromEntity(BOTEdge, ncount)
        nodes = Mesh.GetNodes()

        'Run analysis
        Debug.Print("Running the analysis")
        Debug.Print("")
        errCode = Study.RunAnalysis
        'If errCode <> 0 Then ErrorMsg(app, "Analysis failed with error code as defined in swsRunAnalysisError_e: " & errCode, True)

        Dim Results As CWResults
        Dim nStep As Integer
        Results = Study.Results
        'If CWFeatObj6 Is Nothing Then ErrorMsg(app, "Failed to get results object", True)
        Debug.Print("Study results...")
        nStep = Results.GetMaximumAvailableSteps

        Dim selectedAndModelReactionFM As Object = Nothing
        Dim selectedOnlyReactionFM As Object = Nothing

        Dim times As Object
        times = Results.GetTimeOrFrequencyAtEachStep(0, errCode)
        Dim timesString As String(,) = var_string(times)

        'xlApp.Visible = True
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheetTime As Excel.Worksheet = xlWorkBook.ActiveSheet
        write2excel(times, xlWorkSheetTime, timesString, "Time")

        Dim node_base_string As String(,) = var_string(node_base)
        Dim xlWorkSheetNodeBase As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        write2excel(node_base, xlWorkSheetNodeBase, node_base_string, "Node Base")

        Dim node_TOP_string As String(,) = var_string(node_TOP)
        Dim xlWorkSheetNodeTOP As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        write2excel(node_TOP, xlWorkSheetNodeTOP, node_TOP_string, "Node TOP")

        Dim node_BOT_string As String(,) = var_string(node_BOT)
        Dim xlWorkSheetNodeBOT As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        write2excel(node_BOT, xlWorkSheetNodeBOT, node_BOT_string, "Node BOT")

        Dim nodes_string As String(,) = var_string(nodes)
        Dim xlWorkSheetNodes As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        write2excel(nodes, xlWorkSheetNodes, nodes_string, "Nodes")

        Dim forces2(nStep - 1) As Object
        Dim forcesstring(nodes.length / 4 * 9 - 1, nStep - 1) As String
        Dim xlWorkSheetForce As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        For j = 0 To (nStep - 1)
            forces2(j) = Results.GetReactionForcesAndMomentsWithSelections(j + 1, Nothing, 0, (rot_entity), selectedAndModelReactionFM, selectedOnlyReactionFM, errCode)
            If errCode <> 0 Then ErrorMsg(app, "Failed to get reaction forces and moments", True)
            For jj = 0 To (forces2(j).length - 1)
                forcesstring(jj, j) = forces2(j)(jj).ToString
            Next jj

        Next j
        xlWorkSheetForce.Name = "Loads"
        Dim rangeforce As Excel.Range = xlWorkSheetForce.Range(xlWorkSheetForce.Cells(1, 1), xlWorkSheetForce.Cells(nodes.length / 4 * 9, nStep))
        rangeforce.Value2 = forcesstring

        Dim deform(nStep - 1) As Object
        Dim deform_string(nodes.length / 4 * 5 - 1, nStep - 1) As String
        Dim xlWorkSheetDeform As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        For k = 0 To (nStep - 1)
            deform(k) = Results.GetTranslationalDisplacement(k + 1, Nothing, 2, errCode)
            For kk = 0 To (deform(k).length - 1)
                deform_string(kk, k) = deform(k)(kk).ToString
            Next kk
        Next k
        xlWorkSheetDeform.Name = "Deform"
        Dim rangedeform As Excel.Range = xlWorkSheetDeform.Range(xlWorkSheetDeform.Cells(1, 1), xlWorkSheetDeform.Cells(nodes.length / 4 * 5, nStep))
        rangedeform.Value2 = deform_string

        Dim VM_MinMax(nStep - 1) As Object
        Dim VM_MinMax_string(4, nStep - 1) As String
        Dim xlWorkSheetVM_MinMax As Excel.Worksheet = xlWorkBook.Sheets.Add(Count:=1)
        For k = 0 To (nStep - 1)
            VM_MinMax(k) = Results.GetMinMaxStress(9, 0, k + 1, Nothing, 0, errCode)
            For kk = 0 To (VM_MinMax(k).length - 1)
                VM_MinMax_string(kk, k) = VM_MinMax(k)(kk).ToString
            Next kk
        Next k
        xlWorkSheetVM_MinMax.Name = "VM_MinMax"
        Dim rangeVM_MinMax As Excel.Range = xlWorkSheetVM_MinMax.Range(xlWorkSheetVM_MinMax.Cells(1, 1), xlWorkSheetVM_MinMax.Cells(4, nStep))
        rangeVM_MinMax.Value2 = VM_MinMax_string

        xlWorkBook.SaveAs("Z:\StudentFolders\Haider\Projects\Optimization\SolidWorks Interfacing Optimizer\tempfolder\Results.xlsx")
        xlApp.Workbooks.Close()

        xlApp.Quit()
        Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        app.CloseAllDocuments(True) 'Closing all documents without save
        status = app.UnloadAddIn(sAddinName)
        app.ExitApp()
        Runtime.InteropServices.Marshal.ReleaseComObject(app)
    End Sub
    Function PointsFromXL(xlApp As Excel.Application, strFilename As String, sheet As String) As Double(,)
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(strFilename)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(sheet)
        Dim xlRange As Excel.Range = xlWorkSheet.UsedRange

        Dim points(xlRange.Rows.Count - 1, xlRange.Columns.Count - 1) As Double
        Array.Copy(xlRange.Value, points, xlRange.Rows.Count * xlRange.Columns.Count)
        xlApp.Workbooks.Close()
        PointsFromXL = points
    End Function
    Function ParameterFromXL(xlApp As Excel.Application, strFilename As String, sheet As String, cell As String) As Double
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(strFilename)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets(sheet)
        Dim para As Double = xlWorkSheet.Range(cell).Value

        xlApp.Workbooks.Close()
        ParameterFromXL = para
    End Function
    Function PointsFromTxt(file As String) As Double(,)
        Dim TxtArray As String() = System.IO.File.ReadAllLines(file)
        Dim points(TxtArray.Length - 1, 2) As Double
        Dim i As Integer
        For i = 0 To TxtArray.Length - 1
            Dim tempString As String() = Split(TxtArray(i), ",")
            points(i, 0) = CDbl(tempString(0))
            points(i, 1) = CDbl(tempString(1))
            points(i, 2) = CDbl(tempString(2))
        Next
        PointsFromTxt = points
    End Function
    Function var_string(var As Object) As String(,)
        Dim temp(var.length - 1, 0) As String

        For i = 0 To (var.length - 1)
            temp(i, 0) = var(i).ToString
        Next i
        var_string = temp
    End Function
    Function LaunchSW() As SldWorks
        Dim app As SldWorks = CreateObject("SldWorks.application")
        app.Visible = True
        app.CloseAllDocuments(True) 'Closing all documents without save
        LaunchSW = app
    End Function

    Sub write2excel(var As Object, sheet As Excel.Worksheet, var_string As String(,), sheet_name As String)
        Dim end_range As String = "A" & var.length
        Dim rangetime As Excel.Range = sheet.Range("A1", end_range)
        rangetime.Value2 = var_string
        sheet.Name = sheet_name
    End Sub
    Function CreateSurface(featureMgr As FeatureManager, model As ModelDoc2, points As Double(,)) As Integer
        Dim nPoints As Integer = points.Length / 3
        Dim nLinesTOP As Integer = nPoints
        Dim lines(nLinesTOP - 1) As SketchLine
        Dim pt1(2) As Double
        Dim pt2(2) As Double
        For i = 0 To nPoints - 2
            pt1(0) = points(i, 0)
            pt1(1) = points(i, 1)
            pt1(2) = points(i, 2)

            pt2(0) = points(i + 1, 0)
            pt2(1) = points(i + 1, 1)
            pt2(2) = points(i + 1, 2)
            lines(i) = model.CreateLine2(pt1(0), pt1(1), pt1(2), pt2(0), pt2(1), pt2(2))
        Next

        Dim line As SketchLine
        For i = 0 To lines.Length - 2
            line = lines(i)
            If i = 0 Then
                line.Select(False)
            Else
                line.Select(True)
            End If
        Next
        featureMgr.MakeStyledCurves2(0.001, 1)
        CreateSurface = nPoints
    End Function
    Sub CreateEnds(model As ModelDoc2, TOPpoints As Double(,), BOTpoints As Double(,), nPointsTOP As Integer, nPointsBOT As Integer)
        model.CreateLine2(TOPpoints(0, 0), TOPpoints(0, 1), TOPpoints(0, 2), BOTpoints(0, 0), BOTpoints(0, 1), BOTpoints(0, 2))
        model.CreateLine2(TOPpoints(nPointsTOP - 1, 0), TOPpoints(nPointsTOP - 1, 1), TOPpoints(nPointsTOP - 1, 2),
                          BOTpoints(nPointsBOT - 1, 0), BOTpoints(nPointsBOT - 1, 1), BOTpoints(nPointsBOT - 1, 2))
    End Sub
    Sub ExtrudeSketch(model As ModelDoc2, featureMgr As FeatureManager, points As Double(,), extrusion As Double)
        Dim swModelDocExtension As ModelDocExtension = model.Extension
        swModelDocExtension.SelectByID2("Sketch1", "SKETCH", points(0, 0), points(0, 1), points(0, 2), False, 0, Nothing, 0)
        featureMgr.FeatureExtrusion3(True, False, False, 0, 0, extrusion, 0, False, False, False, False, 0, 0, False, False, False, False, False, False, False, 0, 0, False)
    End Sub
    Sub ErrorMsg(ByVal SwApp As SldWorks, ByVal Message As String, ByVal EndTest As Boolean)
        SwApp.SendMsgToUser2(Message, 0, 0)
        SwApp.RecordLine("'*** WARNING - General")
        SwApp.RecordLine("'*** " & Message)
        SwApp.RecordLine("")
        If EndTest Then
        End If
    End Sub
    Sub CreateAxisofRotation(model As ModelDoc2)
        Dim swModelDocExtension As ModelDocExtension = model.Extension
        swModelDocExtension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
        swModelDocExtension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
        model.InsertAxis2(True)
    End Sub
    Function SetupCosmos(app As SldWorks) As CWModelDoc
        Dim COSMOSObject As Object = app.GetAddInObject("SldWorks.Simulation")
        'If COSMOSObject Is Nothing Then ErrorMsg(app, "COSMOSObject object not found", True)
        Dim COSMOSWORKS As Object = COSMOSObject.CosmosWorks
        'If COSMOSWORKS Is Nothing Then ErrorMsg(app, "COSMOSWORKS object not found", True)
        Dim ActDoc As CWModelDoc = COSMOSWORKS.ActiveDoc()
        'If ActDoc Is Nothing Then ErrorMsg(app, "No active document", True)
        SetupCosmos = ActDoc
    End Function
    Function CreateNonLinearStudy(ActDoc As CWModelDoc, errCode As Integer, errorCode As Integer) As CWStudy
        Dim StudyMngr As CWStudyManager = ActDoc.StudyManager()
        Dim Study As CWStudy = StudyMngr.CreateNewStudy3("Nonlinear", 5, 1, errCode)

        Dim SolidMgr As CWSolidManager = Study.SolidManager
        If SolidMgr Is Nothing Then ErrorMsg(app, "CWSolidManager object not created", True)
        Dim CompCount As Integer = SolidMgr.ComponentCount

        Dim SolidComponent As CWSolidComponent = SolidMgr.GetComponentAt(0, errorCode)
        Dim SolidBody As CWSolidBody = SolidComponent.GetSolidBodyAt(0, errCode)
        If errCode <> 0 Then ErrorMsg(app, "No solid body", True)
        Dim bApp As Boolean = SolidBody.SetLibraryMaterial("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2021\Custom Materials\Custom Materials.sldmat", "TPU 95A Substitute")
        If bApp = False Then ErrorMsg(app, "No material applied", True)
        CreateNonLinearStudy = Study
    End Function
    Function IdentifyEntity(model As ModelDoc2, type As String, locx As Double, locy As Double, locz As Double, errCode As Integer) As Object
        Dim boolstatus As Boolean = model.Extension.SelectByID2("", type, locx, locy, locz, False, 0, Nothing, 0)
        Dim obj As Object = model.SelectionManager.GetSelectedObject6(1, -1)
        Dim PID As Object = model.Extension.GetPersistReference3(obj)
        Dim SelObj As Object = model.Extension.GetObjectByPersistReference3((PID), errCode)
        IdentifyEntity = SelObj 'Face Fixed
    End Function
    Sub TimerElapsed(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        app.CloseAllDocuments(True) 'Closing all documents without save
        app.ExitApp()
        End
    End Sub
    Public app As SldWorks
End Module

