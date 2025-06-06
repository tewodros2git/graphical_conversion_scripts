Attribute VB_Name = "ReplaceSSwitchonTCRs"
Sub ReplaceShapesWithStencilObjects()
    Dim visioApp As Object
    Dim visioDoc As Object
    Dim pg As Object
    Dim shp As Object
    Dim stencilPath As String
    Dim stencil As Object
    Dim masterNames As String
    Dim replacementMasterName As String
    Dim additionalMasterName As String
    Dim replacementMaster As Object
    Dim additionalMaster As Object
    Dim replacementShape As Object
    Dim additionalShape As Object
    Dim shapesToRemove As New Collection ' Collection to store shapes to be removed
    
    On Error GoTo ErrorHandler
    
    ' Define the path to the stencil containing the replacement objects
    stencilPath = "C:\Users\SHUTEW\Desktop\VToR\process_adhoc2\Stencil11.vssx"
    
    ' Set Visio application and active document
    Set visioApp = Application
    Set visioDoc = visioApp.ActiveDocument
    
    If visioDoc Is Nothing Then
        Debug.Print "No active document!"
        GoTo Cleanup
    End If
    
    ' Set the active page
    Debug.Print "Using active page..."
    Set pg = visioApp.ActivePage
    If pg Is Nothing Then
        Debug.Print "No active page found!"
        GoTo Cleanup
    End If
    
    ' Open the stencil document
    Debug.Print "Opening stencil document..."
    Set stencil = visioApp.Documents.OpenEx(stencilPath, visOpenReadOnly)
    If stencil Is Nothing Then
        Debug.Print "Failed to open stencil document!"
        GoTo Cleanup
    End If
    
    ' Loop through each shape on the page
    For Each shp In pg.Shapes
        If InStr(shp.name, "SWITCH") > 0 Or InStr(shp.name, "TCR_TOGGLE") > 0 Then
            ' Get the names of the replacement and additional masters dynamically based on shp.Name
            masterNames = GetReplacementMasterNames(shp.name, shp)
            replacementMasterName = Split(masterNames, ",")(0)
            additionalMasterName = Split(masterNames, ",")(1)
        
            ' Find the replacement master in the stencil
            Set replacementMaster = Nothing
            For Each Master In stencil.Masters
                If Master.name = replacementMasterName Then
                    Set replacementMaster = Master
                    Exit For
                End If
            Next Master
            
            If Not replacementMaster Is Nothing Then
                Debug.Print "Replacement master found for shape " & shp.name & ": " & replacementMasterName
                
                ' Add a new shape based on the replacement master
                Set replacementShape = pg.Drop(replacementMaster, shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX), shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY))
                If Not replacementShape Is Nothing Then
                    ' Copy shape name
                    replacementShape.name = shp.name
                    ' Append a unique suffix to the original shape name to ensure uniqueness
                    shp.name = shp.name & "_" & "old"
                    ' Remove the trailing number Visio adds for uniqueness
                    replacementShape.name = Left(replacementShape.name, InStr(replacementShape.name, ".") - 1)
                    
                    ' Call CopyShapeData subroutine to copy shape data
                    CopyShapeData shp, replacementShape

                    ' Find the additional master in the stencil
                    Set additionalMaster = Nothing
                    For Each Master In stencil.Masters
                        If Master.name = additionalMasterName Then
                            Set additionalMaster = Master
                            Exit For
                        End If
                    Next Master

                    If Not additionalMaster Is Nothing Then
                        ' Position the additional shape relative to the replacement shape
                        Dim offsetX As Double
                        Dim offsetY As Double
                        offsetX = 0 ' Adjust the X offset as needed
                        offsetY = 0 ' Adjust the Y offset as needed
                        Set additionalShape = pg.Drop(additionalMaster, replacementShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX) + offsetX, replacementShape.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY) + offsetY)
                        If Not additionalShape Is Nothing Then
                            Debug.Print "Additional shape added on top of replacement shape."
                        Else
                            Debug.Print "Failed to add additional shape."
                        End If
                    Else
                        Debug.Print "Failed to find additional master: " & additionalMasterName
                    End If
                    
                    ' Add replaced shape to the list for removal
                    shapesToRemove.Add shp
                Else
                    Debug.Print "Failed to add replacement shape for shape: " & shp.name
                End If
            End If
        End If
    Next shp
    
    ' Remove replaced shapes
    For Each shp In shapesToRemove
        RemoveReplacedShape shp
    Next shp
    
Cleanup:
    ' Clean up
    If Not stencil Is Nothing Then
        stencil.Close
    End If
    Set stencil = Nothing
    Set visioDoc = Nothing
    Set visioApp = Nothing
    Set pg = Nothing
    Set shp = Nothing
    Set replacementMaster = Nothing
    Set replacementShape = Nothing
    Set additionalMaster = Nothing
    Set additionalShape = Nothing
    Debug.Print "DONE"
    Exit Sub
    
ErrorHandler:
    Debug.Print "An error occurred: " & Err.Description & " (Error number: " & Err.Number & ")"
    Resume Cleanup
End Sub

Function GetReplacementMasterNames(shapeName As String, shape) As String
    Dim replacementMaster As String
    Dim additionalMaster As String
    Debug.Print shapeName
    If shapeName Like "*TCR_V_SWITCH*" Then
        replacementMaster = "VSwitch1"
        additionalMaster = "JPortsV1"
    ElseIf shapeName Like "*TCR_TOGGLE*" Or shapeName Like "*S_SWITCH*" Then
        Dim i As Integer
        Dim propValue As String
        Dim labelText As String
        Dim found As Boolean
        found = False
        
        If shape.SectionExists(visSectionProp, visExistsLocally) Then
            For i = 0 To shape.rowCount(visSectionProp) - 1
                labelText = LCase(Trim(shape.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr("")))
                If labelText = "j-ports" Or labelText = "j_ports" Then
                    propValue = shape.CellsSRC(visSectionProp, i, visCustPropsValue).ResultStr("")
                   
                    If propValue = "0,270,180,null" Or propValue = "180,0,270,null" Then
                        replacementMaster = "S_SWITCH"
                        additionalMaster = "JPortsS_Switch"
                    ElseIf propValue = "180,270,0,null" Then
                        replacementMaster = "S_SWITCH.Rev"
                        additionalMaster = "JPortsSRev"
                    ElseIf propValue = "180,90,0,null" Then
                        replacementMaster = "S_SWITCH.Rev2"
                        additionalMaster = "JPortsSRev2"
                    ElseIf propValue = "0,90,180,null" Then
                        replacementMaster = "S_SWITCH.Rev2"
                        additionalMaster = "JPortsSRev1"
                    ElseIf propValue = "270,180,0,null" Or propValue = "0,180,270,null" Then
                        replacementMaster = "S_SWITCH.Rev3"
                        additionalMaster = "JPortsSRev3"
                    End If
                    
                    found = True
                    Exit For
                End If
            Next i
        End If

    Else
        ' Default master name if no match is found
        replacementMaster = "DefaultMasterName"
        additionalMaster = "DefaultAdditionalMasterName"
    End If
    
    ' Return the master names as a comma-separated string
    GetReplacementMasterNames = replacementMaster & "," & additionalMaster
End Function
Sub RemoveReplacedShape(replacedShape As Object)
    ' Remove the replaced shape from the page
    replacedShape.Delete
End Sub
Public Sub CopyShapeData(ByVal sourceShape As Object, ByVal targetShape As Object)
    ' Transfer all data rows from the source shape to the target shape
    
    Const allCells As Boolean = True
    Const forceAdd As Boolean = True
    Const matchByName As Boolean = True
    Const matchByLabel As Boolean = True
    
    Dim vSource As Variant
    vSource = GetSourceData(sourceShape, allCells)
    
    If SetTargetData(targetShape, allCells, vSource) Then
        Debug.Print "Data successfully transferred from " & sourceShape.name & " to " & targetShape.name
    Else
        Debug.Print "Data failed to transfer from " & sourceShape.name & " to " & targetShape.name
    End If
End Sub
Private Sub CopyShapeProperties(ByVal sourceShape As Object, ByVal targetShape As Object, ByVal pg As Object)
    Dim prop As Object
    
    ' Iterate through each property of the source shape
    For Each prop In sourceShape.Section(visSectionProp).Rows
        ' Check if the property exists in the target shape
        If targetShape.CellExistsU("Prop." & prop.name, Visio.visExistsAnywhere) <> 0 Then
            ' Copy the property value from source to target shape
            targetShape.CellsU("Prop." & prop.name).FormulaU = prop.Cells("Value").Result("")
        End If
    Next prop
End Sub
Public Function GetSourceData(ByVal shp As Visio.shape, ByVal allCells As Boolean) As Variant
    On Error GoTo errHandler
    
     If shp Is Nothing Then
        GetSourceData = Nothing
        Exit Function
    End If
    
    Dim iRows As Integer
    Dim irow As Integer
    Dim cellName As String
    Dim cellValue As String
    
    ' Check if the shape has any rows in the properties section
    iRows = shp.rowCount(Visio.VisSectionIndices.visSectionProp)
    
    If iRows = 0 Then
        GetSourceData = Nothing
        Exit Function
    End If
    
    Dim avarFormulaArray() As Variant
    ReDim avarFormulaArray(1 To iRows, 1 To 2) As Variant ' 2 columns: name and value
    
    ' Iterate over each row in the properties section
    For irow = 0 To iRows - 1
        cellName = shp.CellsSRC(Visio.VisSectionIndices.visSectionProp, irow, Visio.VisCellIndices.visCustPropsLabel).RowNameU
        cellValue = shp.CellsSRC(Visio.VisSectionIndices.visSectionProp, irow, Visio.VisCellIndices.visCustPropsValue).ResultStr("")
        
        ' Store cell name and value in the array
        avarFormulaArray(irow + 1, 1) = cellName
        avarFormulaArray(irow + 1, 2) = cellValue
    Next irow

    GetSourceData = avarFormulaArray

    Exit Function

errHandler:
    MsgBox "An unexpected error occurred in GetSourceData: " & Err.Description, vbCritical, "Error"
    Resume exitHere
    
exitHere:
    Exit Function
End Function
Public Function UnwrapTextForLabelShapes(shp)
 ' Loop through each shape on the page
    For Each shp In pg.Shapes
        If shp.name Like "label*" Then
            ' Unwrap the text if it is wrapped
            shp.Cells("TxtWidth").FormulaU = "0"
            'Debug.Print "Unwrapped text for shape: " & shp.name
        End If
    Next shp

End Function

Public Function SetTargetData(ByVal shp As Visio.shape, ByVal allCells As Boolean, ByVal aryData As Variant) As Boolean
    ' Set the data on the target shape
    On Error GoTo errHandler

    'Debug.Print "Data type of aryData: " & VarType(aryData)

    Dim totalRows As Integer
    totalRows = UBound(aryData, 1)

    If totalRows = 0 Then
        SetTargetData = False
        Exit Function
    End If

    Dim irow As Integer
    Dim rowName As String
    Dim rowLabel As String
    Dim rowValue As String

    ' Loop through each row in the array
    For irow = LBound(aryData, 1) To totalRows
        ' Ensure the array contains enough elements for a row
        If UBound(aryData, 2) >= 2 Then
            rowName = ReplaceInvalidCharacters(aryData(irow, 1))
            'rowLabel = "" ' There's no label in the array, so it's empty
            rowValue = aryData(irow, 2)
            'Debug.Print rowName & " : " & rowValue

            ' Add a named row to the shape's properties section
            Dim row As Integer
            row = shp.AddNamedRow(Visio.VisSectionIndices.visSectionProp, rowName, 0)
            shp.CellsSRC(Visio.VisSectionIndices.visSectionProp, row, Visio.VisCellIndices.visCustPropsValue).FormulaU = """" & rowValue & """"
            shp.CellsSRC(Visio.VisSectionIndices.visSectionProp, row, Visio.VisCellIndices.visCustPropsLabel).FormulaU = """" & rowName & """"
    End If
    Next irow

    SetTargetData = True
    Exit Function

errHandler:
    MsgBox Err.Description, vbCritical, "SetTargetData"
    SetTargetData = False
End Function
Sub FindConnectedConnectors(shapeName As String)
    Dim visShape As Visio.shape
    Dim connectedConnectors As Visio.Connects
    Dim connector As Visio.shape
    Dim connectedShape As Visio.shape
    Dim connectedShapeName As String
    Dim foundShape As Boolean
    
    foundShape = False
    
    ' Search for the shape by name
    For Each visShape In ActivePage.Shapes
        If visShape.name = shapeName Then
            Set connectedConnectors = visShape.Connects
            foundShape = True
            Exit For
        End If
    Next visShape
    
    If foundShape Then
        If connectedConnectors.count > 0 Then
            For Each connector In connectedConnectors
                Set connectedShape = connector.ToSheet
                connectedShapeName = connectedShape.name
                Debug.Print "Connector connected to shape: " & connectedShapeName
            Next connector
        Else
            Debug.Print "No connectors connected to the specified shape."
        End If
    Else
        Debug.Print "Shape with name '" & shapeName & "' not found."
    End If
End Sub
Private Sub SetCellFormula(ByVal shp As Visio.shape, _
    ByVal iSect As Integer, ByVal irow As Integer, ByVal iCell As Integer, _
    ByVal formula As String)
    ' Transfer cell formula if different
    If Not shp.CellsSRC(iSect, irow, iCell).FormulaU = formula Then
        shp.CellsSRC(iSect, irow, iCell).FormulaForceU = """" & formula & """"

    End If
End Sub
Function ReplaceInvalidCharacters(ByVal name As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.pattern = "[^a-zA-Z0-9_]"
    regex.Global = True
    ReplaceInvalidCharacters = regex.Replace(name, "_")
End Function




