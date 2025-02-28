Attribute VB_Name = "ReportPageShapes"
Public Sub ReportPageShapes()
    Dim vPag As Visio.Page
    Set vPag = ActivePage
    Dim FilePath As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Get the page name and split based on the "." delimiter
    Dim pageNameParts As Variant
    pageNameParts = Split(vPag.name, "_")
    
'
'    ' Process pageNameParts(0) based on its first character
'    Dim folderPrefix As String
'    If Left(pageNameParts(0), 1) = "k" Then
'        folderPrefix = "KU_" & UCase(Mid(pageNameParts(0), 2, Len(pageNameParts(0)) - 3)) & "_" & UCase(Right(pageNameParts(0), 2))
'    ElseIf Left(pageNameParts(0), 1) = "c" Then
'        folderPrefix = "CB_" & UCase(Mid(pageNameParts(0), 2, Len(pageNameParts(0)) - 3)) & "_" & UCase(Right(pageNameParts(0), 2))
'    ElseIf Left(pageNameParts(0), 1) = "t" Then
'        folderPrefix = "TCR_" & "BEACONS"
'    Else
'        folderPrefix = UCase(Mid(pageNameParts(0), 1, Len(pageNameParts(0)) - 2)) & "_" & UCase(Right(pageNameParts(0), 2)) ' Default handling if not "k" or "c"
'    End If

    ' Create folder structure based on pageNameParts(1)
    Dim targetFolder As String
    targetFolder = "C:\Users\SHUTEW\Desktop\VToR\untitled\" & pageNameParts(0)

    ' Check if folder exists, if not, create it
    If Dir(targetFolder, vbDirectory) = "" Then
        MkDir targetFolder
    End If

    ' Create "json" folder inside targetFolder if it doesn't exist
    If Dir(targetFolder & "\json", vbDirectory) = "" Then
        MkDir targetFolder & "\json"
    End If

    ' Set the FilePath to save JSON file
    FilePath = targetFolder & "\json\" & vPag.name & ".txt"

    ' Write initial JSON structure
    Open FilePath For Output As #1
    Write #1, '"{""spacecraftId"": """ & pageNameParts(1) & "_" & folderPrefix & """,""shapes"": ["
    Close #1

    ' Write content using the WriteContent subroutine
    WriteContent "{""spacecraftId"": """ & vPag.name & """,""shapes"": [", FilePath

    Dim shp As Visio.shape

    ' Loop through shapes and process them
    For Each shp In vPag.Shapes
        If (shp.name Like "coax*") Or (shp.name Like "logical*") Or (shp.name Like "Dynamic connector*") Or (shp.name Like "waveguide*") Or (shp.name Like "ConnLine*") Then
            GetLineTraceInfo shp, dict, FilePath
        ElseIf (shp.name Like "MaxarTWTA*") Then
            If shp.Shapes.count > 0 Then
                Dim s As Visio.shape
                For Each s In shp.Shapes
                    If (s.name Like "coax*") Or (s.name Like "logical*") Or (s.name Like "Dynamic connector*") Or (s.name Like "waveguide*") Or (s.name Like "ConnLine*") Then
                        GetLineTraceInfo s, dict, FilePath
                    End If
                Next
            End If
        End If
    Next

    WriteContent "]}", FilePath
    CleanFile2 FilePath
    CleanFile FilePath
    RenameFile FilePath, vPag.name
    Debug.Print "Done!"
End Sub

Public Sub GetLineTraceInfo(ByRef shp As Visio.shape, dict As Object, FilePath As String)
    Dim vsoConnects As Visio.Connects
    Dim vsoConnect As Visio.Connect
    Dim intCounter As Integer
    Dim vsoConnectFromCell As Visio.Cell
    Dim vsoConnectToCell As Visio.Cell

    Dim iPropSect As Integer
    iPropSect = Visio.VisSectionIndices.visSectionProp

    Dim shapeExists As Boolean
    shapeExists = False

    If dict.Exists(shp.name) Then
        shapeExists = True
    End If

    If shapeExists = False Then
        dict.Add shp.name, shp.name
        Dim FileContent As String

        ' Check if the shape has a property section with at least 5 properties
        If shp.SectionExists(iPropSect, Visio.VisExistsFlags.visExistsAnywhere) <> 0 Then
            Dim i As Integer
            If (shp.Section(iPropSect).count >= 5) Then
                Dim vcell As Visio.Cell
                Dim vSrcRd As Visio.Cell
                Dim vSrcPort As Visio.Cell
                Dim vTargetRd As Visio.Cell
                Dim vTargetPort As Visio.Cell

                ' Loop through the properties and assign appropriate cells based on labels
                For i = 0 To shp.Section(iPropSect).count - 1 Step 1
                    Set vCellLabel = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsLabel)
                    Set vcell = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsValue)

                    Select Case vCellLabel.formula
                    Case """x-source-parent-rd"""
                        Set vSrcRd = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsValue)
                    Case """x-source-value"""
                        Set vSrcPort = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsValue)
                    Case """x-target-parent-rd"""
                        Set vTargetRd = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsValue)
                    Case """x-target-value"""
                        Set vTargetPort = shp.CellsSRC(iPropSect, i, Visio.VisCellIndices.visCustPropsValue)
                    End Select
                Next i

                ' Create a JSON-like string for the shape
                FileContent = "{""Name"":""" & shp.name & """,""From"":""" & vSrcRd.ResultStr("") & "-" & vSrcPort.ResultStr("") & """,""To"":""" & vTargetRd.ResultStr("") & "-" & vTargetPort.ResultStr("") & """},"
                
                ' Write the content to the file
                WriteContent FileContent, FilePath
            End If
        End If
    End If
End Sub

' Subroutine to write content to the file
Private Sub WriteContent(FileContent As String, FilePath As String)
    Const ForAppending = 8
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(FilePath, ForAppending, True)
    f.Write FileContent
    f.Close
End Sub

' Subroutine to clean the file content (removes trailing ",]}")
Public Sub CleanFile(FilePath As String)
    Dim objFSO
    Dim objTS 'define a TextStream object
    Dim strContents As String
    Dim strContents1 As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTS = objFSO.OpenTextFile(FilePath, 1)
    strContents = objTS.ReadAll
    strContents1 = Replace(strContents, ",]}", "]}") ' Remove trailing comma
    objTS.Close

    Set objTS = objFSO.OpenTextFile(FilePath, 2)
    objTS.Write strContents1
    objTS.Close
End Sub

' Subroutine to further clean the file content (removes double commas ",,")
Public Sub CleanFile2(FilePath As String)
    Dim objFSO
    Dim objTS 'define a TextStream object
    Dim strContents As String
    Dim strContents1 As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTS = objFSO.OpenTextFile(FilePath, 1)
    strContents = objTS.ReadAll
    strContents1 = Replace(strContents, ",,", ",") ' Remove double commas
    objTS.Close

    Set objTS = objFSO.OpenTextFile(FilePath, 2)
    objTS.Write strContents1
    objTS.Close
End Sub

' Subroutine to rename the file after cleaning up
Public Sub RenameFile(FilePath As String, NewFileName As String)
    Dim NewFilePath As String
    NewFilePath = Left(FilePath, InStrRev(FilePath, "\")) & NewFileName & ".json"
    
    ' Debugging: Print the original and new file paths
    Debug.Print "Original File Path: " & FilePath
    Debug.Print "New File Path: " & NewFilePath
    
    ' Check if the original file exists before renaming
    If Dir(FilePath) <> "" Then
        On Error Resume Next
        ' Check if the new file path already exists, and delete it if it does
        If Dir(NewFilePath) <> "" Then
            Kill NewFilePath ' Delete if already exists
        End If
        On Error GoTo 0

        ' Attempt to rename the file
        On Error GoTo RenameError
        Name FilePath As NewFilePath
        Debug.Print "File renamed successfully to " & NewFilePath
        Exit Sub

RenameError:
        MsgBox "Error renaming file: " & Err.Description
    Else
        MsgBox "File not found: " & FilePath
    End If
End Sub


