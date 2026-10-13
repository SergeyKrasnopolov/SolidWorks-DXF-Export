Attribute VB_Name = "ExportBatchDXF"
Option Explicit

' Batch DXF export. One file + referenced configuration = one unique item.
' Export geometry is adapted from the supplied ExportLayoutDXF module.
' No source document is saved. Test in SOLIDWORKS before production use.
Private Const DRAWING_TEMPLATE As String = ""
Private Const SHEET_METAL_OPTIONS As Long = 1
Private swApp As SldWorks.SldWorks
Private fso As Object
Private parts As Object
Private reservedNames As Object
Private issues As Collection
Private itemWarnings As String
Private totalInstances As Long
Private suppressedNodes As Long

Private Sub AddPart(ByVal path As String, ByVal config As String, ByVal virtualPart As Boolean)
    Dim key As String, row As Variant
    If Len(path) = 0 Or Len(config) = 0 Then
        issues.Add "Incomplete part identity: " & path & " [" & config & "]"
        Exit Sub
    End If
    key = LCase$(path) & vbLf & LCase$(config)
    If parts.Exists(key) Then
        row = parts(key)
        row(2) = CLng(row(2)) + 1
        parts(key) = row
    Else
        parts.Add key, Array(path, config, CLng(1), virtualPart)
    End If
    totalInstances = totalInstances + 1
End Sub

Private Sub CollectChildren(ByVal parent As SldWorks.Component2)
    On Error GoTo Failed
    Dim children As Variant, i As Long
    Dim child As SldWorks.Component2
    children = parent.GetChildren
    If IsEmpty(children) Then Exit Sub
    For i = LBound(children) To UBound(children)
        Set child = children(i)
        CollectComponent child
    Next i
    Exit Sub
Failed:
    issues.Add "Cannot fully traverse " & parent.Name2 & ": " & Err.Description
End Sub

Private Sub CollectComponent(ByVal comp As SldWorks.Component2)
    On Error GoTo Failed
    Dim model As SldWorks.ModelDoc2, path As String, kind As Long
    If comp.IsSuppressed Then
        suppressedNodes = suppressedNodes + 1
        Exit Sub
    End If
    path = comp.GetPathName
    Set model = comp.GetModelDoc2
    If Not model Is Nothing Then
        kind = model.GetType
    Else
        Select Case LCase$(fso.GetExtensionName(path))
            Case "sldprt": kind = swDocPART
            Case "sldasm": kind = swDocASSEMBLY
        End Select
    End If
    If kind = swDocPART Then
        AddPart path, comp.ReferencedConfiguration, comp.IsVirtual
    ElseIf kind = swDocASSEMBLY Then
        If model Is Nothing Then
            issues.Add "Unloaded subassembly; quantity is incomplete: " & comp.Name2
        Else
            ' Traverse every occurrence, including repeated subassemblies.
            CollectChildren comp
        End If
    Else
        issues.Add "Unknown or unloaded component: " & comp.Name2 & " | " & path
    End If
    Exit Sub
Failed:
    issues.Add "Cannot inspect component: " & Err.Description
End Sub

Private Function SafeStem(ByVal value As String) As String
    Dim bad As Variant, ch As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        value = Replace(value, CStr(ch), "_")
    Next ch
    For i = 0 To 31
        value = Replace(value, Chr$(i), "_")
    Next i
    value = Left$(Trim$(value), 70)
    Do While Len(value) > 0
        If Right$(value, 1) <> "." And Right$(value, 1) <> " " Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) = 0 Then value = "unnamed"
    SafeStem = value
End Function

Private Function FreePath(ByVal folder As String, ByVal stem As String, ByVal extension As String) As String
    Dim candidate As String, suffix As String, n As Long
    n = 1
    Do
        suffix = ""
        If n > 1 Then suffix = "__" & CStr(n)
        candidate = fso.BuildPath(folder, stem & suffix & extension)
        If Len(candidate) > 240 Then Err.Raise vbObjectError + 30, , "Output path is too long. Choose a shorter folder."
        If Not fso.FileExists(candidate) And Not fso.FolderExists(candidate) And Not reservedNames.Exists(LCase$(candidate)) Then Exit Do
        n = n + 1
    Loop
    reservedNames.Add LCase$(candidate), True
    FreePath = candidate
End Function

Private Function Csv(ByVal value As String) As String
    Csv = """" & Replace(value, """", """""") & """"
End Function

Private Sub SaveUtf8(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile path, 1 ' Never replace another report.
    stream.Close
End Sub

Private Function ActivateModel(ByVal model As SldWorks.ModelDoc2) As SldWorks.ModelDoc2
    Dim active As SldWorks.ModelDoc2, errs As Long, name As String
    name = model.GetPathName
    If Len(name) > 0 Then Set active = swApp.ActivateDoc3(name, False, swDontRebuildActiveDoc, errs)
    If active Is Nothing Then Set active = swApp.ActivateDoc3(model.GetTitle, False, swDontRebuildActiveDoc, errs)
    If active Is Nothing Then Exit Function
    If Len(name) > 0 Then
        If LCase$(active.GetPathName) <> LCase$(name) Then Exit Function
    Else
        If active.GetTitle <> model.GetTitle Then Exit Function
    End If
    Set ActivateModel = active
End Function

Private Function ExportOne(ByVal path As String, ByVal config As String, ByVal target As String, ByRef detail As String) As Boolean
    On Error GoTo Failed
    Dim model As SldWorks.ModelDoc2, active As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim wasLoaded As Boolean, wasVisible As Boolean, oldConfig As String
    Dim errs As Long, warns As Long, result As Boolean, cleanupError As String
    Dim alignment(11) As Double, vAlignment As Variant
    itemWarnings = ""
    If Not fso.FileExists(path) Then Err.Raise vbObjectError + 31, , "Part file does not exist."
    Set model = swApp.GetOpenDocumentByName(path)
    wasLoaded = Not (model Is Nothing)
    If wasLoaded Then
        oldConfig = model.ConfigurationManager.ActiveConfiguration.Name
        wasVisible = model.Visible
    Else
        Set model = swApp.OpenDoc6(path, swDocPART, swOpenDocOptions_Silent, config, errs, warns)
    End If
    If model Is Nothing Then Err.Raise vbObjectError + 32, , "Cannot open part. Error code: " & errs
    If model.GetType <> swDocPART Then Err.Raise vbObjectError + 33, , "Document is not a part."
    If model.GetConfigurationByName(config) Is Nothing Then Err.Raise vbObjectError + 34, , "Configuration not found: " & config
    model.Visible = True
    ' Verify identity after activation; never export a different same-named document.
    Set active = ActivateModel(model)
    If active Is Nothing Then Err.Raise vbObjectError + 35, , "Cannot activate part. Error code: " & errs
    If LCase$(active.GetPathName) <> LCase$(path) Then Err.Raise vbObjectError + 36, , "Wrong document activated."
    If Not model.ShowConfiguration2(config) Then Err.Raise vbObjectError + 37, , "Cannot activate configuration: " & config
    model.ForceRebuild3 False
    model.ClearSelection2 True
    If HasFlatPattern(model.FirstFeature) Then
        Set part = model
        vAlignment = alignment
        If Not part.ExportToDWG2(target, path, swExportToDWG_ExportSheetMetal, True, _
                                vAlignment, False, False, SHEET_METAL_OPTIONS, Empty) Then
            Err.Raise vbObjectError + 38, , "Sheet-metal export failed."
        End If
    Else
        ExportProjection model, target
    End If
    If Not fso.FileExists(target) Then Err.Raise vbObjectError + 39, , "DXF was not created."
    If fso.GetFile(target).Size = 0 Then Err.Raise vbObjectError + 40, , "DXF is empty."
    result = True
    detail = itemWarnings
    If warns <> 0 Then detail = detail & " OpenDoc warning: " & warns
    GoTo Cleanup
Failed:
    detail = Err.Description
Cleanup:
    On Error Resume Next
    Err.Clear
    If Not model Is Nothing Then
        If wasLoaded Then
            If Len(oldConfig) > 0 Then
                If model.ConfigurationManager.ActiveConfiguration.Name <> oldConfig Then
                    If Not model.ShowConfiguration2(oldConfig) Then cleanupError = "Could not restore configuration. "
                    model.ForceRebuild3 False
                End If
            End If
            model.Visible = wasVisible
        Else
            swApp.CloseDoc model.GetTitle
        End If
    End If
    If Err.Number <> 0 Then cleanupError = cleanupError & Err.Description
    On Error GoTo 0
    If Len(cleanupError) > 0 Then
        detail = detail & " RESTORE WARNING: " & cleanupError
        issues.Add path & " [" & config & "]: " & cleanupError
    End If
    ExportOne = result
End Function

Private Sub RunBatch(ByVal source As SldWorks.ModelDoc2)
    On Error GoTo Failed
    Dim assembly As SldWorks.AssemblyDoc, root As SldWorks.Component2
    Dim picker As Object, selected As Object, folder As String
    Dim key As Variant, row As Variant, issue As Variant
    Dim report As String, reportPath As String, summaryPath As String
    Dim target As String, detail As String, status As String, summary As String
    Dim exported As Long, failedCount As Long, processed As Long, resolveStatus As Long
    Dim active As SldWorks.ModelDoc2, errs As Long, sourceName As String, sourceConfig As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set parts = CreateObject("Scripting.Dictionary")
    Set reservedNames = CreateObject("Scripting.Dictionary")
    Set issues = New Collection
    totalInstances = 0
    suppressedNodes = 0
    sourceName = source.GetPathName
    If Len(sourceName) = 0 Then sourceName = source.GetTitle
    sourceConfig = source.ConfigurationManager.ActiveConfiguration.Name
    Set picker = CreateObject("Shell.Application")
    Set selected = picker.BrowseForFolder(0, "Choose the folder for DXF files and the quantity report", &H41, 0)
    If selected Is Nothing Then Exit Sub
    folder = selected.Self.path
    If Not fso.FolderExists(folder) Then Err.Raise vbObjectError + 41, , "Select an existing local or network folder."
    reportPath = FreePath(folder, "DXF_report_" & Format$(Now, "yyyymmdd_hhnnss"), ".csv")
    summaryPath = FreePath(folder, "DXF_summary_" & Format$(Now, "yyyymmdd_hhnnss"), ".txt")
    ' Verify output access before opening any part.
    SaveUtf8 summaryPath, "Batch started: " & sourceName & " [" & sourceConfig & "]" & vbCrLf
    If source.GetType = swDocASSEMBLY Then
        Set assembly = source
        resolveStatus = assembly.ResolveAllLightWeightComponents(False)
        If assembly.GetLightWeightComponentCount > 0 Then issues.Add "Some components remain lightweight; counts may be incomplete. Resolve result: " & resolveStatus
        Set root = source.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
        If root Is Nothing Then Err.Raise vbObjectError + 42, , "Cannot read assembly tree. Open the assembly fully resolved."
        CollectChildren root
        Set root = Nothing
    Else
        AddPart source.GetPathName, sourceConfig, False
    End If
    ' Collection is complete before configurations/documents are changed.
    report = "File;Configuration;Quantity;DXF;Status;Details" & vbCrLf
    For Each key In parts.Keys
        row = parts(key)
        processed = processed + 1
        target = FreePath(folder, "DXF_" & SafeStem(fso.GetBaseName(CStr(row(0)))) & "__" & SafeStem(CStr(row(1))), ".dxf")
        detail = ""
        If CBool(row(3)) Then
            status = "SKIPPED"
            detail = "Virtual part: save it as an external SLDPRT first."
            failedCount = failedCount + 1
        ElseIf ExportOne(CStr(row(0)), CStr(row(1)), target, detail) Then
            status = "OK"
            exported = exported + 1
            If Len(detail) > 0 Then status = "WARNING"
        Else
            status = "ERROR"
            failedCount = failedCount + 1
        End If
        report = report & Csv(CStr(row(0))) & ";" & Csv(CStr(row(1))) & ";" & CStr(row(2)) & ";" & _
                 Csv(target) & ";" & Csv(status) & ";" & Csv(detail) & vbCrLf
    Next key
    SaveUtf8 reportPath, report
    summary = "Assembly / part: " & sourceName & vbCrLf & "Configuration: " & sourceConfig & vbCrLf & _
              "Unique parts (file + configuration): " & parts.Count & vbCrLf & _
              "Total part instances: " & totalInstances & vbCrLf & _
              "DXF exported: " & exported & vbCrLf & _
              "Failed / skipped unique parts: " & failedCount & vbCrLf & _
              "Suppressed component nodes skipped: " & suppressedNodes & vbCrLf & _
              "Traversal / restore issues: " & issues.Count & vbCrLf & _
              "Report: " & reportPath & vbCrLf
    If issues.Count > 0 Then summary = summary & "Counts may be incomplete. See issues below." & vbCrLf
    For Each issue In issues
        summary = summary & vbCrLf & CStr(issue)
    Next issue
    ' Replace only the start marker that this run created, never an older user's file.
    Set active = ActivateModel(source)
    If active Is Nothing Then summary = summary & vbCrLf & "Could not reactivate the source document."
    ReplaceOwnSummary summaryPath, summary
    MsgBox summary, vbInformation, "DXF batch export"
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Len(reportPath) > 0 And Len(report) > 0 Then
        If Not fso.FileExists(reportPath) Then SaveUtf8 reportPath, report
    End If
    If Len(summaryPath) > 0 Then ReplaceOwnSummary summaryPath, "BATCH INCOMPLETE: " & detail & vbCrLf & _
        "Processed: " & processed & " / " & parts.Count & vbCrLf & "Partial report: " & reportPath
    Set active = ActivateModel(source)
    On Error GoTo 0
    MsgBox "Batch stopped: " & detail & vbCrLf & "Partial report: " & reportPath, vbExclamation
End Sub

Private Sub ReplaceOwnSummary(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile path, 2
    stream.Close
End Sub

Private Function HasFlatPattern(ByVal feat As SldWorks.Feature) As Boolean
    Dim child As SldWorks.Feature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "FlatPattern" Then
            HasFlatPattern = True
            Exit Function
        End If
        Set child = feat.GetFirstSubFeature
        Do While Not child Is Nothing
            If child.GetTypeName2 = "FlatPattern" Then
                HasFlatPattern = True
                Exit Function
            End If
            Set child = child.GetNextSubFeature
        Loop
        Set feat = feat.GetNextFeature
    Loop
End Function

Private Sub ExportProjection(ByVal model As SldWorks.ModelDoc2, ByVal target As String)
    On Error GoTo Failed
    Dim doc As SldWorks.ModelDoc2, draw As SldWorks.DrawingDoc
    Dim view As SldWorks.view, Sheet As SldWorks.Sheet
    Dim activeModel As SldWorks.ModelDoc2
    Dim template As String, tempView As String, tempSheet As String
    Dim config As String, docTitle As String
    Dim bounds As Variant, pos As Variant
    Dim width As Double, height As Double
    Dim errs As Long, warns As Long, activateErr As Long
    Dim oldMulti As Long, oldNoScale As Long
    Dim oldMap As Boolean, oldMapDialog As Boolean, oldPaper As Boolean
    Dim oldBlocks As Boolean, oldAutoScale As Boolean
    Dim settingsRead As Boolean, named As Boolean
    Dim failure As String, failureNumber As Long

    template = DRAWING_TEMPLATE
    If Len(template) = 0 Then template = swApp.GetUserPreferenceStringValue(swDefaultTemplateDrawing)
    If Len(template) = 0 Then
        Err.Raise vbObjectError + 10, , "Set the default drawing template, or DRAWING_TEMPLATE in this module."
    End If
    If Not CreateObject("Scripting.FileSystemObject").FileExists(template) Then
        Err.Raise vbObjectError + 11, , "Drawing template not found: " & template
    End If

    config = model.ConfigurationManager.ActiveConfiguration.Name
    ' User explicitly sets Normal To before running. No unreliable selection-based rotation.
    tempView = "DXF_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100))
    model.NameView tempView
    named = True
    tempSheet = tempView

    oldMulti = swApp.GetUserPreferenceIntegerValue(swDxfMultiSheetOption)
    oldNoScale = swApp.GetUserPreferenceIntegerValue(swDxfOutputNoScale)
    oldMap = swApp.GetUserPreferenceToggle(swDxfMapping)
    oldMapDialog = swApp.GetUserPreferenceToggle(swDXFDontShowMap)
    oldPaper = swApp.GetUserPreferenceToggle(swDxfExportAllSheetsToPaperSpace)
    oldBlocks = swApp.GetUserPreferenceToggle(swDxfExportViewAsBlock)
    oldAutoScale = swApp.GetUserPreferenceToggle(swAutomaticScaling3ViewDrawings)
    settingsRead = True
    swApp.SetUserPreferenceIntegerValue swDxfMultiSheetOption, swDxfActiveSheetOnly
    ' Both sheet and view are explicitly 1:1; disable additional scale conversion.
    swApp.SetUserPreferenceIntegerValue swDxfOutputNoScale, 0
    swApp.SetUserPreferenceToggle swDxfMapping, False
    swApp.SetUserPreferenceToggle swDXFDontShowMap, True
    swApp.SetUserPreferenceToggle swDxfExportAllSheetsToPaperSpace, False
    swApp.SetUserPreferenceToggle swDxfExportViewAsBlock, False
    swApp.SetUserPreferenceToggle swAutomaticScaling3ViewDrawings, False

    Set doc = swApp.NewDocument(template, swDwgPaperA3size, 0, 0)
    If doc Is Nothing Then Err.Raise vbObjectError + 12, , "Could not create temporary drawing."
    docTitle = doc.GetTitle
    Set draw = doc
    doc.SetUserPreferenceIntegerValue swUnitSystem, swUnitSystem_MMGS
    If Not draw.NewSheet3(tempSheet, swDwgPapersUserDefined, swDwgTemplateNone, _
                         1, 1, True, "", 1, 1, "") Then
        Err.Raise vbObjectError + 13, , "Could not create an empty drawing sheet."
    End If
    If Not draw.ActivateSheet(tempSheet) Then Err.Raise vbObjectError + 14, , "Could not activate export sheet."
    Set view = draw.CreateDrawViewFromModelView3(model.GetPathName, tempView, 0.5, 0.5, 0)
    If view Is Nothing Then Err.Raise vbObjectError + 15, , "Could not create projection view."
    view.ReferencedConfiguration = config
    view.UseSheetScale = False
    view.ScaleDecimal = 1#
    If Not view.SetDisplayMode3(False, swHIDDEN, False, False) Then
        Err.Raise vbObjectError + 16, , "Could not set hidden-lines-removed display."
    End If
    view.SetDisplayTangentEdges2 swTangentEdgesHidden
    doc.ForceRebuild3 False

    ' Remove automatic annotations (centre marks, labels, dimensions) on the export sheet.
    RemoveAnnotations doc, draw
    bounds = view.GetOutline
    width = CDbl(bounds(2)) - CDbl(bounds(0))
    height = CDbl(bounds(3)) - CDbl(bounds(1))
    If width <= 0 Or height <= 0 Then Err.Raise vbObjectError + 17, , "The projection is empty."
    ' Sheet and view sizes are in metres. Leave a 10 mm margin.
    Set Sheet = draw.GetCurrentSheet
    If Not Sheet.SetSize(swDwgPapersUserDefined, width + 0.02, height + 0.02) Then
        Err.Raise vbObjectError + 20, , "Could not resize export sheet to fit the layout."
    End If
    pos = view.Position
    pos(0) = CDbl(pos(0)) + 0.01 - CDbl(bounds(0))
    pos(1) = CDbl(pos(1)) + 0.01 - CDbl(bounds(1))
    view.Position = pos
    doc.ForceRebuild3 False
    doc.ClearSelection2 True
    If Not doc.Extension.SaveAs(target, swSaveAsCurrentVersion, swSaveAsOptions_Silent, _
                                Nothing, errs, warns) Then
        Err.Raise vbObjectError + 18, , "DXF SaveAs failed. Error code: " & errs & "; warnings: " & warns
    End If
    If warns <> 0 Then itemWarnings = itemWarnings & " DXF SaveAs warning: " & warns
    GoTo Cleanup
Failed:
    failure = Err.Description
    failureNumber = Err.Number
Cleanup:
    On Error Resume Next
    If Not doc Is Nothing Then swApp.CloseDoc doc.GetTitle
    If named Then model.DeleteNamedView tempView
    If settingsRead Then
        swApp.SetUserPreferenceIntegerValue swDxfMultiSheetOption, oldMulti
        swApp.SetUserPreferenceIntegerValue swDxfOutputNoScale, oldNoScale
        swApp.SetUserPreferenceToggle swDxfMapping, oldMap
        swApp.SetUserPreferenceToggle swDXFDontShowMap, oldMapDialog
        swApp.SetUserPreferenceToggle swDxfExportAllSheetsToPaperSpace, oldPaper
        swApp.SetUserPreferenceToggle swDxfExportViewAsBlock, oldBlocks
        swApp.SetUserPreferenceToggle swAutomaticScaling3ViewDrawings, oldAutoScale
    End If
    Set activeModel = ActivateModel(model)
    If Err.Number <> 0 Then itemWarnings = itemWarnings & " Projection cleanup: " & Err.Description
    If activeModel Is Nothing Then itemWarnings = itemWarnings & " Could not reactivate part after projection."
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "ExportProjection", failure
End Sub

Private Sub RemoveAnnotations(ByVal doc As SldWorks.ModelDoc2, ByVal draw As SldWorks.DrawingDoc)
    Dim view As SldWorks.view
    Dim ann As SldWorks.Annotation
    Dim selected As Boolean
    doc.ClearSelection2 True
    Set view = draw.GetFirstView
    Do While Not view Is Nothing
        Set ann = view.GetFirstAnnotation3
        Do While Not ann Is Nothing
            If ann.Select3(True, Nothing) Then selected = True
            Set ann = ann.GetNext3
        Loop
        Set view = view.GetNextView
    Loop
    If selected Then
        If Not doc.Extension.DeleteSelection2(swDelete_Absorbed) Then
            Err.Raise vbObjectError + 19, , "Could not remove drawing annotations."
        End If
    End If
    doc.ClearSelection2 True
End Sub




Public Sub main()
    On Error GoTo Failed
    Dim source As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set source = swApp.ActiveDoc
    If source Is Nothing Then Err.Raise vbObjectError + 50, , "Open an assembly or part first."
    If source.GetType <> swDocASSEMBLY And source.GetType <> swDocPART Then
        Err.Raise vbObjectError + 51, , "Only assemblies and parts are supported."
    End If
    RunBatch source
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "DXF batch export"
End Sub
