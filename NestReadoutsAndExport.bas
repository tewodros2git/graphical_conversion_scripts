Attribute VB_Name = "NestReadoutsAndExport"
Public Sub NestReadoutsAndExport()
    RearrangeShapes
    ExportVisioToSVGAndVisio
End Sub

Sub RearrangeShapes()
    Dim shape As shape
    Dim parentShape As shape
    Dim childShape As shape
    Dim uniqueID As Variant ' For dictionary keys
    Dim title As String
    Dim pattern As Variant
    Dim patterns1 As Variant
    Dim patterns2 As Variant
    Dim shapesToNest1 As Object ' Dictionary for first group of parent shapes
    Dim shapesToNest2 As Object ' Dictionary for second group of parent shapes
    Dim childShapes As Collection
    Dim groupShapeIDs As Collection
    Dim i As Long

    ' Initialize dictionaries for parent shapes in each pattern group
    Set shapesToNest1 = CreateObject("Scripting.Dictionary")
    Set shapesToNest2 = CreateObject("Scripting.Dictionary")

    ' Define patterns for each group
    patterns1 = Array() '"_TWTA_", "_EPC_", "_LCAMP_", "_CAMP_")
    patterns2 = Array("_DOWN_CONVERTER_", "RECEIVER", "_CMDRX_", "_BBE_", "_XMTR_", "BEACONS", "TWTA_", "TCR_UNIT_", "_TWTA_", "_EPC_", "_LCAMP_", "_LCHAMP_", "_CAMP_")

    ' Loop through each shape to identify parents for both pattern groups
    For Each shape In ActivePage.Shapes
        title = shape.name

        ' Identify parents for patterns1
        For Each pattern In patterns1
            If InStr(1, title, pattern, vbTextCompare) > 0 Then
                uniqueID = Right(title, 3) ' Last 3 characters of title for unique ID
                If Not shapesToNest1.Exists(uniqueID) Then
                    shapesToNest1.Add uniqueID, shape
                End If
                Exit For
            End If
        Next pattern

        ' Identify parents for patterns2
        For Each pattern In patterns2
            If InStr(1, title, pattern, vbTextCompare) > 0 Then
                uniqueID = title ' Use full title as unique ID
                If Not shapesToNest2.Exists(uniqueID) Then
                    shapesToNest2.Add uniqueID, shape
                End If
                Exit For
            End If
        Next pattern
    Next shape

    ' Loop through each uniqueID in patterns1 to find and nest child shapes under each parent
    For Each uniqueID In shapesToNest1.Keys
        Set parentShape = shapesToNest1(uniqueID)

        ' Initialize collections for child shapes and group shape IDs
        Set childShapes = New Collection
        Set groupShapeIDs = New Collection
        groupShapeIDs.Add parentShape.ID ' Add parent ID

        ' Identify children based on uniqueID match in their titles
        For Each shape In ActivePage.Shapes
            title = shape.name
            If InStr(1, title, "Readout.", vbTextCompare) > 0 And InStr(1, title, uniqueID, vbTextCompare) > 0 Then
                childShapes.Add shape
                groupShapeIDs.Add shape.ID
            End If
        Next shape

        ' Group and rename child shapes if any are found
        GroupAndRenameShapes parentShape, childShapes, groupShapeIDs
    Next uniqueID

    ' Loop through each uniqueID in patterns2 to find and nest child shapes under each parent
    For Each uniqueID In shapesToNest2.Keys
        Set parentShape = shapesToNest2(uniqueID)

        ' Initialize collections for child shapes and group shape IDs
        Set childShapes = New Collection
        Set groupShapeIDs = New Collection
        groupShapeIDs.Add parentShape.ID ' Add parent ID

        ' Identify children based on Prop.parent value match
        For Each shape In ActivePage.Shapes
            If shape.CellExistsU("Prop.parent", False) Then
                Dim parentProp As String
                parentProp = shape.Cells("Prop.parent").ResultStr("")
                If parentProp = uniqueID Then
                    childShapes.Add shape
                    groupShapeIDs.Add shape.ID
                End If
            End If
        Next shape

        ' Group and rename child shapes if any are found
        GroupAndRenameShapes parentShape, childShapes, groupShapeIDs
    Next uniqueID

    'MsgBox "Shapes have been rearranged, grouped, and renamed successfully."
End Sub

Sub GroupAndRenameShapes(parentShape As shape, childShapes As Collection, groupShapeIDs As Collection)
    Dim i As Long
    Dim childShape As shape
    Dim title As String
    Dim newGroup As shape

    If childShapes.count > 0 Then
        ' Clear selection in ActiveWindow
        ActiveWindow.DeselectAll

        ' Add each shape to selection using ID
        For Each ShapeID In groupShapeIDs
            ActiveWindow.Select ActivePage.Shapes.ItemFromID(ShapeID), visSelect
        Next ShapeID

        ' Group the selected shapes
        Set newGroup = ActiveWindow.Selection.Group
        newGroup.name = parentShape.name '& "_Group"

        ' Rename child shapes based on specific patterns within the group
        For i = 1 To newGroup.Shapes.count
            Set childShape = newGroup.Shapes(i)
            title = childShape.name

            ' Apply renaming rules based on patterns
            Dim newName As String
            Select Case True
                Case InStr(1, title, "LIN-STEP", vbTextCompare) > 0 Or _
                     InStr(1, title, "OPA", vbTextCompare) > 0
                    newName = "Readout.OPA"
                Case InStr(1, title, "LCHAMP-GAIN", vbTextCompare) > 0
                    newName = "Readout.Gain"
                Case InStr(1, title, "ON-OFF-STATUS", vbTextCompare) > 0 And Not InStr(1, title, "POWER-ON-OFF-STATUS", vbTextCompare) > 0
                    newName = "Readout.Status"
                Case InStr(1, title, "LOAD-CURRENT", vbTextCompare) > 0 Or _
                     InStr(1, title, "BUS-CURRENT", vbTextCompare) > 0 Or InStr(1, title, "LDI", vbTextCompare) > 0
                    newName = "Readout.LDI"
                Case InStr(1, title, "ARU-STS", vbTextCompare) > 0 Or _
                      InStr(1, title, "HELIX-PROT", vbTextCompare) > 0
                    newName = "Readout.ARU"
                Case InStr(1, title, "HI-VOLT-STS", vbTextCompare) > 0
                    newName = "Readout.IBO"
                Case InStr(1, title, "LCHAMP-OUTLEVEL", vbTextCompare) > 0 Or _
                     InStr(1, title, "OP-POWER", vbTextCompare) > 0
                    newName = "Readout.OutputPwr"
                Case InStr(1, title, "LCHAMP-RFBLANK", vbTextCompare) > 0 Or _
                     InStr(1, title, "Mute", vbTextCompare) > 0
                    newName = "Readout.Mute"
                Case InStr(1, title, "HELIX-CURRENT", vbTextCompare) > 0
                    newName = "Readout.Helix"
                Case InStr(1, title, "POWER", vbTextCompare) > 0
                    newName = "Readout.High-Voltage"
                Case InStr(1, title, "ANODE-VOLTAGE", vbTextCompare) > 0
                    newName = "Readout.Anode"
                Case InStr(1, title, "ARU-ENABLE", vbTextCompare) > 0
                    newName = "Readout.ARU-ENABLE"
                Case InStr(1, title, "FIL-VOLTAGE", vbTextCompare) > 0 Or _
                     InStr(1, title, "ALC-IP-ATTR-VOLT", vbTextCompare) > 0
                    newName = "Readout.InputPwr"
                Case InStr(1, title, "ATT-", vbTextCompare) > 0 Or _
                     InStr(1, title, "ALC-SETTING", vbTextCompare) > 0 Or InStr(1, title, "ATTEN", vbTextCompare) > 0
                    newName = "Readout.Attenuation"
                Case InStr(1, title, "BAND", vbTextCompare) > 0 Or _
                     InStr(1, title, "MODE", vbTextCompare) > 0
                    newName = "Readout.Mode"
                Case Else
                    Debug.Print "Edge Case: " & title
                    newName = "" ' Default case
            End Select

            ' Update the name if a new name was assigned
            If newName <> "" Then
                childShape.name = newName
                childShape.NameU = newName

                ' Update User-defined cells, if necessary
                Dim rowIndex As Integer
                For rowIndex = 0 To childShape.rowCount(visSectionUser) - 1
                    If InStr(1, childShape.CellsSRC(visSectionUser, rowIndex, visUserValue).FormulaU, title) > 0 Then
                        childShape.CellsSRC(visSectionUser, rowIndex, visUserValue).FormulaU = Chr(34) & newName & Chr(34)
                    End If
                Next rowIndex

                ' Update Prop.Label and Prop.Name if they exist
                If childShape.CellExists("Prop.Label", False) Then
                    childShape.Cells("Prop.Label").FormulaU = Chr(34) & newName & Chr(34)
                End If
                If childShape.CellExists("Prop.Name", False) Then
                    childShape.Cells("Prop.Name").FormulaU = Chr(34) & newName & Chr(34)
                End If

                ' Set dummy value to force refresh
                childShape.Cells("Width").FormulaU = childShape.Cells("Width").Result("in") & " in"
            End If
        Next i
    End If
End Sub

Sub ExportVisioToSVGAndVisio()
    Dim svgFileName As String
    Dim visioFileName As String
    Dim vPag As Visio.Page
    Set vPag = ActivePage
    Dim pageNameParts As Variant
    Dim baseFolder As String
    Dim svgFolder As String
    Dim visioFolder As String

    ' Get the page name and split based on the "_" delimiter
    pageNameParts = Split(vPag.name, "_")

    ' Define the base folder structure
    baseFolder = "C:\Users\SHUTEW\Desktop\VToR\untitled\" & pageNameParts(0) & "\"

    ' Ensure the base folder exists
    If Dir(baseFolder, vbDirectory) = "" Then
        MkDir baseFolder
    End If

    ' Create the svg folder
    svgFolder = baseFolder & "svg\"
    If Dir(svgFolder, vbDirectory) = "" Then
        MkDir svgFolder
    End If

    ' Create the visio folder
    visioFolder = baseFolder & "visio\"
    If Dir(visioFolder, vbDirectory) = "" Then
        MkDir visioFolder
    End If

    ' Define the SVG file name
    svgFileName = svgFolder & vPag.name & ".svg"

    ' Export the current page to SVG
    ActivePage.Export svgFileName

    ' Define the Visio file name
    visioFileName = visioFolder & vPag.name & ".vsd"

    ' Save the current page as a Visio file
    Application.ActiveDocument.SaveAs visioFileName

    ' Notify the user of success
    MsgBox "Export successful!" & vbCrLf & _
           "SVG: " & svgFileName & vbCrLf & _
           "Visio: " & visioFileName, vbInformation
End Sub


