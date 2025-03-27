Attribute VB_Name = "AutoFixArrows"
Option Explicit

Sub SortShapesByName(ByRef shapeArray() As Visio.shape)
    Dim shapeCount As Long
    Dim i As Long, j As Long
    Dim temp As Visio.shape
    
    ' Get the number of shapes in the array
    shapeCount = UBound(shapeArray) - LBound(shapeArray) + 1
    
    ' Exit if the array is empty or has only one shape
    If shapeCount <= 1 Then Exit Sub
    
    ' Bubble sort shapes by name length, then alphabetically
    For i = LBound(shapeArray) To UBound(shapeArray) - 1
        For j = i + 1 To UBound(shapeArray)
            ' First compare by string length
            If Len(shapeArray(i).name) > Len(shapeArray(j).name) Or _
               (Len(shapeArray(i).name) = Len(shapeArray(j).name) And _
               StrComp(shapeArray(i).name, shapeArray(j).name, vbTextCompare) > 0) Then
               
                ' Swap shapes
                Set temp = shapeArray(i)
                Set shapeArray(i) = shapeArray(j)
                Set shapeArray(j) = temp
            End If
        Next j
    Next i
End Sub


Public Sub FixArrows()

    Dim doc As Document
    For Each doc In Visio.Documents
    
        Dim scid As String
        scid = doc.Pages(1).name
        Debug.Print scid
        scid = Left(scid, InStr(scid, "_") - 1)
        
        Dim username As String
        username = Environ("USERNAME")
        
        ' get arrows on current page
        Dim s As shape
        Dim arrows As Integer
        Dim arrs() As shape
        arrows = 0
        For Each s In doc.Pages(1).Shapes
            If InStr(s.name, "HOT") <> 0 Then
                ReDim Preserve arrs(arrows)
                Set arrs(arrows) = s
                arrows = arrows + 1
            End If
        Next s
        
        SortShapesByName arrs
        
        Dim fileName As String
        Dim fileNumber As Integer
        Dim lineContent As String
        
        ' get arrow data from 1.0 for current page
        fileName = "C:\Users\SHUTEW\Desktop\VToR\untitled\temp\output\" & scid & "\" & doc.Pages(1).name + "_arrows"
        Debug.Print fileName
        
        Dim lines() As String
        Dim fileContent As String
        Dim fso As Object
        Dim file As Object
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(fileName, 1)
        If Not file.AtEndOfStream Then
        
            fileContent = file.ReadAll
            file.Close
            
            lines = Split(fileContent, vbLf)
            
            Dim l As Variant
            Dim i As Integer
            i = 0
            Dim errors As String
            
            ' each arrow in current page in order
            For Each l In lines
                'Debug.Print l
            
                Dim tokes() As String
                tokes = Split(l, ",")
                
                Dim act As shape
                Set act = arrs(i)
                Dim link As String
                link = Replace(act.CellsSRC(visSectionProp, 0, visCustPropsValue).FormulaU, Chr(34), "") & ".vsd"
                link = Replace(link, "ACTIVE_", "")
                'Debug.Print "-> " & link
                i = i + 1
                
                Dim found As Boolean
                found = False
                
                ' check all other pages
                Dim d As Document
                For Each d In Visio.Documents
                
                    ' found the linked page
                    If InStr(d.name, link) <> 0 Then
                        'Debug.Print "     " & d.name
                        
                        Dim so As shape
                        Dim arrs2() As shape
                        Dim arrows2 As Integer
                        arrows2 = 0
                        ' gather arrows on linked page
                        For Each so In d.Pages(1).Shapes
                            If InStr(so.name, "HOT") <> 0 Then
                                ReDim Preserve arrs2(arrows2)
                                Set arrs2(arrows2) = so
                                arrows2 = arrows2 + 1
                            End If
                        Next so
                        
                        SortShapesByName arrs2
                        
                        'Debug.Print "     need to link to index " & tokes(9)
                        'ebug.Print "     len " & UBound(arrs2)
                        
                        If CInt(tokes(9)) < UBound(arrs2) Then
                            
                            Dim need As String
                            ' name of linked arrow
                            need = Chr(34) & arrs2(Abs(CInt(tokes(9)))).name & Chr(34)
                            
                            'Debug.Print "     [" & act.name & "] -> [" & need + "]"
                            
                            ' find dynamic connector on linked page
                            For Each so In d.Pages(1).Shapes
                                If InStr(so.name, "Dynamic conn") <> 0 Then
                                    Dim li1 As String
                                    Dim li2 As String
                                    li1 = so.CellsSRC(visSectionProp, 2, visCustPropsValue).FormulaU
                                    li2 = so.CellsSRC(visSectionProp, 4, visCustPropsValue).FormulaU
                                    If StrComp(li1, need) = 0 Or StrComp(li2, need) = 0 Then
                                        'Debug.Print "        will go to  " & d.Pages(1) & so.name
                                        Dim rowIndex As Integer
                                        rowIndex = act.CellsRowIndex("Prop.Trace")
                                        act.CellsSRC(visSectionProp, rowIndex, visCustPropsValue).FormulaU = """" & so.name & """"
                                        found = True
                                    End If
                                End If
                            Next so
                        
                        End If
                        
                    End If
                
                Next d
                
                If Not found Then
                    errors = errors & doc.Pages(1).name & "->" & link & "::" & act.name & vbNewLine
                End If
                
            Next l
            
            
        
        End If
        
    Next doc
    
    Debug.Print "===================================="
    Debug.Print ""
    Debug.Print "Errors:"
    Debug.Print errors
    Debug.Print ""
    Debug.Print "===================================="
    Debug.Print "End Script"
    
End Sub
