Attribute VB_Name = "NestReadoutsAndExport"
Public Sub NestReadoutsAndExport()
    RearrangeShapes
    AddRevisionTimeStamp
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
    patterns2 = Array("_BPS_", "_POWER_", "_MLO_", "_LNA_", "_DOWNCONVERTER_", "_DOWN_CONVERTER_", "RECEIVER", "_CMDRX_", "_BBE_", "_XMTR_", "BEACONS", "_CHAMP_", "TWTA_", "TCR_UNIT_", "_TWTA_", "_EPC_", "_LCAMP_", "_LCHAMP_", "_CAMP_")

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
                     InStr(1, title, "ALC-SETTING", vbTextCompare) > 0 Or InStr(1, title, "ATTEN", vbTextCompare) > 0 _
                     And Not InStr(1, title, "OPA-ATTEN", vbTextCompare) > 0
                    newName = "Readout.Attenuation"
                Case InStr(1, title, "BAND", vbTextCompare) > 0 Or _
                     InStr(1, title, "MODE", vbTextCompare) > 0
                    newName = "Readout.Mode"
                 Case InStr(1, title, "FCA", vbTextCompare) > 0
                    newName = "Readout.FCA"
                Case InStr(1, title, "GCA", vbTextCompare) > 0
                    newName = "Readout.GCA"
                Case InStr(1, title, "HI-POWER", vbTextCompare) > 0
                    newName = "Readout.HPWR"
                Case InStr(1, title, "LO-POWER", vbTextCompare) > 0
                    newName = "Readout.LPWR"
                Case InStr(1, title, "BIAS", vbTextCompare) > 0
                    newName = "Readout.BIAS"
                Case InStr(1, title, "VOLTAGE", vbTextCompare) > 0
                    newName = "Readout.VOLTAGE"
                Case InStr(1, title, "TEMP", vbTextCompare) > 0
                    newName = "Readout.TEMP"
                Case InStr(1, title, "RF-STATUS", vbTextCompare) > 0
                    newName = "Readout.RF-STATUS"
                Case InStr(1, title, "FREQ", vbTextCompare) > 0
                    newName = "Readout.FREQ"
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
Function GetDateStr(lbl As String)
    Dim ix As Integer
    Dim revstr As String
    Dim nRevision As Integer
    Dim dstring As String
    Dim tdstring As String
    Dim finalStr As String
    Dim cr As String
    Dim userName As String

    cr = Chr(13) + Chr(10)

    ' Allow for user to create a custom name that is not his/her Windows username. Edit local environment variables (for non-
    ' admin users) using the command: rundll32 sysdm.cpl,,EditEnvironmentVariables.
    ' Default to the Windows user if not found.
    userName = Environ("CustomUserName")
    If userName = "" Then
        userName = Environ("username")
    End If

    ' Get revision.
    ix = InStr(lbl, "Revision")
    If ix > 0 Then
        ix = ix + 10
        revstr = Trim(Mid(lbl, ix))
        nRevision = CInt(revstr)
    End If

    ' Get previous date modified.
    ix = InStr(lbl, "Date Modified")
    If ix > 0 Then
        dstring = Mid(lbl, ix + 15, 11)
    End If

    ' Get today's date. If different, increment version number. If same, set version to 0.
    tdstring = Format(Date, "dd.mmm.yyyy")
    If dstring = tdstring Then
        nRevision = nRevision + 1
    Else
        nRevision = 0
    End If

    ' Format the new version/date string.
    finalStr = "GDN_2.0" + cr + "Date Modified: " + tdstring + cr + "Modified By: " + userName + cr + "Revision: " + Str(nRevision)
    GetDateStr = finalStr
End Function

' ---------------------------------------------------------------------
'   Change date label on the currently open display.
'
Public Sub AddRevisionTimeStamp()
    ' Search for the date-stamp string.
    Dim shp As shape
    Dim shpF As shape
    Dim found As Boolean
    found = False
    
    ' Get the page height
    Dim pageHeight As Double
    pageHeight = ActivePage.PageSheet.Cells("PageHeight").Result(visUnitInches)
    PageWidth = ActivePage.PageSheet.Cells("PageWidth").Result(visUnitInches)

    ' Loop through all shapes to find the shape with the "date modified" text
    For Each shp In ActivePage.Shapes
        If InStr(LCase(shp.Text), "date modified") > 0 Then
            Set shpF = shp
            found = True
            Exit For
        End If
    Next

    ' If not found, create the shape
    If Not found Then
        Set shpF = ActivePage.DrawRectangle(PageWidth, pageHeight - 0.5833, PageWidth - 1.7275, pageHeight)  ' Example dimensions, adjust as needed
        shpF.Text = "Date Modified: "
    End If

    ' Ensure that the shape is valid before applying color and updating text
    If Not shpF Is Nothing Then
        On Error Resume Next ' Suppress errors temporarily
        
        ' Apply the fill color using THEMEGUARD, same as in your manual code
        shpF.CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "9 pt"
        shpF.CellsSRC(visSectionCharacter, 0, visCharacterOverline).FormulaU = "FALSE"
        shpF.CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "0"
        shpF.CellsSRC(visSectionObject, visRowGradientProperties, visLineGradientEnabled).FormulaU = "FALSE"
        shpF.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(29,17,174))"
        shpF.CellsSRC(visSectionCharacter, 0, visCharacterDblUnderline).FormulaU = "FALSE"
        shpF.CellsSRC(visSectionObject, visRowFill, visFillForegnd).FormulaU = "THEMEGUARD(RGB(245,222,179))"
        shpF.CellsSRC(visSectionObject, visRowFill, visFillBkgnd).FormulaU = "THEMEGUARD(SHADE(FillForegnd,LUMDIFF(THEMEVAL(""FillColor""),THEMEVAL(""FillColor2""))))"
        shpF.CellsSRC(visSectionObject, visRowGradientProperties, visFillGradientEnabled).FormulaU = "FALSE"

        ' Handle any errors in case the Fill property is not accessible
        If Err.Number <> 0 Then
            MsgBox "Error setting fill color: " & Err.Description
            Err.Clear
        End If
        
        ' Update the text of the shape
        Dim lbstr As String
        lbstr = GetDateStr(shpF.Text)
        shpF.Text = lbstr
        shpF.name = "TimeStamp"
                
        On Error GoTo 0 ' Reset error handling
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


