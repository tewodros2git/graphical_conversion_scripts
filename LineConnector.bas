Attribute VB_Name = "LineConnector"

Public Sub RunLines()
    AdjustSwitchShapes
    AddConnectionPoints
    GetConnectionsFromBackground
    AddTextField
    
End Sub
 Sub AdjustSwitchShapes()
    Dim shp As Visio.shape
    For Each shp In Visio.ActivePage.Shapes
        ' Check if the shape's name contains "switch" (case-insensitive)
        If InStr(1, LCase(shp.name), "switch") <> 0 Then
            ' Set the height and width to 0.25 inches
            shp.CellsU("Width").FormulaForceU = "0.25 in"
            shp.CellsU("Height").FormulaForceU = "0.25 in"
        End If
    Next
End Sub
Sub AddConnectionPoints()
    ' Add connection points to switch shapes.
    Dim NewRow As Integer
    Dim nwidth As Integer
    Dim nheight As Integer
    Dim x As Integer
    Dim y As Integer
    Dim vPage As Visio.Page
    Dim shp As Visio.shape
    Dim missingObjs As String
    Dim shpsAddded As String
    Dim npts As Integer
    shpsAdded = ""
    '  For each shape on the active page, identify the shapes that are circulators, loads, switches and
    '  ATS shapes, and add connection points to the top and bottom, left and right of the shape.
    For Each shp In Visio.ActivePage.Shapes
       strName = LCase(shp.name)
        If (strName Like "*tcr_v_switch*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)

            ' Add 2 top connection points.
           NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - 5) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."

            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + 5) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
            ' Add 2 bottom connection points.
            ' Add connections to bottom.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - 5) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."

            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + 5) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."

            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

        ElseIf (strName Like "*tcr_toggle*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)

            ' Add 1 connection to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."

            ' Add 1 connection to the top.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."

            ' Add 1 connection to the bottom.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."

            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

         ElseIf (strName Like "*beacon*") Or (strName Like "*xmitrs*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
            ' Add 1 connection to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
   
         ElseIf (strName Like "*dual_bfn_splitter*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
            ' Add 1 connection to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
            ' Add 2 connections to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
        
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
        
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
            
        ElseIf (strName Like "*triple_bfn_splitter*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
            ' Add 1 connection to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            ' Add 3 connections to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
            
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
            
        ElseIf (strName Like "*quad_bfn_splitter*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
            ' Add 1 connection to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            ' Add 4 connections to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
            
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                     
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
            
         ElseIf (strName Like "*five_bfn_splitter*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
            ' Add 1 connection to the middle left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            ' Add 5 equidistant connections to the right.
            Dim i As Integer
            For i = 1 To 5
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * ((i - 3) / 4))) + " pt."
            Next i
        
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

        ElseIf (strName Like "*eight_bfn_splitter*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
            ' Add 1 connection to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
            ' Add 8 connections to the right.
            For i = 1 To 4
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.375) + (i - 1) * (nheight * 0.25)) + " pt."
            Next i
            
            For i = 5 To 8
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.125) + (i - 5) * (nheight * 0.25)) + " pt."
            Next i
            
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

            
            ElseIf (strName Like "*dual_bfn_coupler*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
            ' Add 2 connections to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
        
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
        
            ' Add 1 connection to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

        ' Add connection points to the four sides of all switches, circulators, loads and satpad shapes.
       ElseIf (strName Like "*circulator*") Or (strName Like "*load*") Or (strName Like "*switch*") Or (strName Like "*iot_coupler*") _
       Or (strName Like "*ats*") Or (strName Like "*satpad*") Or (strName Like "*hot-arrow*") Or (strName Like "*hot_arrow*") _
       Or (strName Like "*CHANNEL_POST*") Or (strName Like "*channel_post*") Or (strName Like "*diplexer_combiner*") Or (strName Like "*junction_block_spli*") Or (strName Like "*gnd_*") Or (strName Like "*hi_power_mode_splitter*") _
       And Not (strName Like "*tcr_toggle_switch*") Or (strName Like "*hi_power_mode_splitter*") Or (strName Like "*junction_block_coup*") Or (strName Like "*channel_post*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)

            ' Add the new points to left and right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."

            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."

            ' Add connections to top and bottom.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."

            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
            
            shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
        ElseIf (strName Like "*dc_converter*") Or strName Like "*DC_CONVERTER*" Then
        
               ' Add connection to top middle.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
                        
            shpsAddded = shpsAddded + vbCrLf + strName + vbCrLf

         ElseIf (strName Like "*receiver*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
                
                 NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
        ElseIf (strName Like "*epc*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                'INA, top left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + 15) + " pt."
                'INB bottom left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - 15) + " pt."
                'OUTA top right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + 15) + " pt."
                'OUTB bottom right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - 15) + " pt."
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
            ElseIf (strName Like "*camp*") Or (strName Like "*lchamp*") Or (strName Like "*twta*") Or (strName Like "*down_converter*") Or (strName Like "*lna*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
           ElseIf (strName Like "*camp*") Or (strName Like "*lchamp*") Or (strName Like "*twta*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
          
                ElseIf (strName Like "*imux*") Then
                    Dim channelsValue As String
                    Dim originalLength As Integer
                    Dim modifiedLength As Integer
                    
                    channelsValue = shp.CellsU("Prop.Channel_Name").ResultStr("")
                    originalLength = Len(channelsValue)
                    modifiedLength = Len(Replace(channelsValue, ";", ""))
                    npts = originalLength - modifiedLength
                    If npts < 1 Then
                        npts = 12
                    End If
                    
                'Debug.Print "IMUX"
                ' 1 channel left, 12 right
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add the new point to the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                offsets = Round(nheight / (npts + 1), 0)
                stPt = offsets
                
                Dim jj As Long
                For jj = 1 To npts
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(stPt) + " pt."
                    stPt = stPt + offsets
                Next jj
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

                ElseIf (strName Like "*t8_cb_even_omux*") Or (strName Like "*t8_cb_odd_omux*") Then
                    channelsValue = shp.CellsU("Prop.Channel_Name").ResultStr("")
                    originalLength = Len(channelsValue)
                    modifiedLength = Len(Replace(channelsValue, ";", ""))
                    npts = originalLength - modifiedLength
                    If npts < 1 Then
                        npts = 12
                    End If
                    
                    ' 1 channel right, 6 on top, 6 on bottom
                    ' Get dimensions of where the points should go.
                    nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                    nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                    x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                    y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
                    
                    ' Add the new point to the right (1 point)
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
                    
                    ' Add points to the top (6 points)
                    offsets = Round(nwidth / 7, 0) ' Dividing by 7 to space 6 points evenly
                    stPt = x - (nwidth * 0.5)
                    
                    For jj = 1 To 6
                        NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(stPt + 6) + " pt."
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
                        stPt = stPt + offsets
                    Next jj
                    
                    ' Add points to the bottom (6 points)
                    stPt = x - (nwidth * 0.5)
                    
                    For jj = 1 To 6
                        NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(stPt + 6) + " pt."
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                        stPt = stPt + offsets
                    Next jj
                    
                    shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf


               ElseIf ((strName Like "*omux*") Or (strName Like "*cmux*")) And (Not strName Like "*t8_cb_even_omux*") And (Not strName Like "*t8_cb_odd_omux*") Then
                    channelsValue = shp.CellsU("Prop.Channel_Name").ResultStr("")
                    originalLength = Len(channelsValue)
                    modifiedLength = Len(Replace(channelsValue, ";", ""))
                    npts = originalLength - modifiedLength
                    If npts < 1 Then
                        npts = 12
                    End If
                    
                    ' 1 channel right, 12 left
                    ' Get dimensions of where the points should go.
                    nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                    nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                    x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                    y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
                
                    ' Add the new point to the right
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
                
                    ' Add to the left
                    offsets = Round(nheight / (npts + 1), 0)
                    stPt = offsets
                    
                    For jj = 1 To npts
                        NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                        shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(stPt) + " pt."
                        stPt = stPt + offsets
                    Next jj
                
                    shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

           ElseIf (strName Like "*six_bfn_splitter*") Then
                'Debug.Print "Splitter"
                
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add the new point to the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Add the new points to the right
                stPt = y + (Round(nheight / 2, 0))
                offsets = Round((stPt) / 6, 0)
            
                Dim k As Long
                For k = 1 To 6
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(stPt) + " pt."
                    stPt = stPt - offsets
                Next k
                
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf strName Like "*comm_dual*" Or strName Like "*COMM_DUAL*" Or strName Like "*comm_dual_in_pol*" Or strName Like "*dual_connect*" Or strName Like "*dual_in*" Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
            ElseIf (strName Like "*comm_quad*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                 NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + 10) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + 10) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
            ElseIf (strName Like "*omni_antenna*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
                
                ' Add the first connection point to the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + 10) + " pt."
            
                ' Add the second connection point to the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - 10) + " pt."
            
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf (strName Like "*comm_tx*") Or (strName Like "*comm_rx*") Or (strName Like "*rx_antenna*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
                ' Add the new points to left and right.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf (strName Like "*xmtr*") Then
                
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add connection points at the top
                ' Top Middle
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
            
                ' Midpoint between top middle and right end
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.25)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
            
                ' Add connection points to the left side
                ' Midpoint on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint on the left in the lower half
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                ' Add connection points to the right side
                ' Midpoint on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint on the right in the lower half
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                        
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf (strName Like "*hybrid*") Then
            Debug.Print "HYBRID"
                '2 left, 2 right
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
                ' Add the new points to Right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + 10) + " pt."
                
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - 10) + " pt."

                'left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + 10) + " pt."
        
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - 10) + " pt."
        
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
                
            ElseIf (strName Like "*tcr_cmdrx*") Then
            ' Get shape dimensions
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add connection points to the left side
                ' Midpoint on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Add connection points to the right side
                ' Midpoint on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint of the lower half on the right side
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                ' Add connection points to the bottom
                ' Bottom middle
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
            
                ' Bottom between middle and start on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.25)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
                        
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf (strName Like "*tcr_unit*") Then
                ' Debug.Print "BBE"
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add connection points to the left side
                ' Midpoint on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Add connection points to the right side
                ' Midpoint on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint of the lower half on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
                
             ElseIf (strName Like "*bbe*") Then
                ' Debug.Print "BBE"
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add connection points to the left side
                ' Midpoint on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint of the upper half on the left
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
            
                ' Add connection points to the right side
                ' Midpoint on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Midpoint of the lower half on the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

        Else
        'skip if background, readout or label is in shape name
        If InStr(strName, "background") = 0 And InStr(strName, "readout") = 0 And InStr(strName, "label") = 0 Then
            missingObjs = missingObjs + strName + vbCrLf
        End If
        
        
        End If
    Next
    'MsgBox shpsAddded + " have been added"
    MsgBox "Objects not handled" + vbCrLf + missingObjs
End Sub


Sub GetConnectionsFromBackground()
        Dim shp As Visio.shape
        
        For Each shp In Visio.ActivePage.Shapes

        
            If InStr(1, shp.name, "background") <> 0 Then
            
                Dim rowCount As Integer
                rowCount = shp.rowCount(Visio.visSectionProp)
                Dim i As Long
                Dim cxns As String
                Dim splitCxns() As String

                
                For i = 0 To rowCount - 1
                    'Debug.Print shp.CellsSRC(visSectionProp, i, visCustPropsValue).ResultStr(visNone)
                    cxns = shp.CellsSRC(visSectionProp, i, visCustPropsValue).ResultStr(visNone)
                    splitCxns = Split(cxns, ",")
                    Dim shp1 As String
                    Dim shp2 As String
                    Dim dir1 As String
                    Dim dir2 As String
                    Dim prt1 As String
                    Dim prt2 As String
                    'IS32-KU-T-SWITCH-43_9, IS32-KU-IMUX-12CH-O_162, P2,LEFT,P1,RIGHT
                    shp1 = Trim(splitCxns(0))
                    shp2 = Trim(splitCxns(1))
                    
                    prt1 = Trim(splitCxns(2))
                    dir1 = Trim(splitCxns(3))
                    prt2 = Trim(splitCxns(4))
                    dir2 = Trim(splitCxns(5))

                    If (shp1 <> "" And shp2 <> "") _
                    Then
                        GlueToShapes shp1, shp2, dir1, dir2, cxns, prt1, prt2
                    Else
                        'Debug.Print shp1 + " " + shp2 + " - skipping"
                    End If
                Next i
                Exit For
            End If
            
        Next


End Sub


Public Sub AddTextField()
        Dim shp As Visio.shape
        For Each shp In Visio.ActivePage.Shapes

        
            If InStr(1, LCase(shp.name), "readout") <> 0 Then
                shp.Text = "-"
                shp.Cells("Char.Size").FormulaForceU = "7pt"
                
                ' Align text to the left
            shp.CellsU("Para.HorzAlign").FormulaForceU = "0"
            End If
        Next

End Sub
Function ConvertPort(prt, shpType)
        Dim cnvPrt As String
        Dim parsePrt As String


    If shpType Like "*SWITCH*" _
    Then

        If prt = "P1" Then
            cnvPrt = "1"
        ElseIf prt = "P2" Then
            cnvPrt = "2"
        ElseIf prt = "P3" Then
            cnvPrt = "3"
        ElseIf prt = "P4" Then
            cnvPrt = "4"
        Else
            cnvPrt = "-1"
        End If
    'EPC has 2 on left 2 on right
    ElseIf shpType Like "*EPC*" Then
        
        If prt = "INA" Or prt = "IN1" Then
            cnvPrt = "2"
        ElseIf prt = "INB" Or prt = "IN2" Then
            cnvPrt = "4"
        ElseIf prt = "OUTA" Or prt = "OUT1" Then
            cnvPrt = "1"
        ElseIf prt = "OUTB" Or prt = "OUT2" Then
            cnvPrt = "3"
        End If

    ElseIf shpType Like "*TWTA*" Or shpType Like "*LCHAMP*" Or shpType Like "*CAMP*" Or shpType Like "*DOWN_CONVERTER*" _
    Or shpType Like "*receiver*" Or shpType Like "*DUAL_OUTPUT_RECEIVER*" Or shpType Like "*T_RECEIVER*" Or shpType Like "*LNA*" Then
        
        If prt = "IN1" Then
            cnvPrt = "2"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        ElseIf prt = "OUT2" Then
            cnvPrt = "3"
        End If
'    ElseIf shpType Like "*DUAL_BFN_SPLITTER*" Or shpType Like "*spli*" Or shpType Like "*JUNCTION_BLOCK_SPLI*" Then
'
'        If prt = "IN1" Then
'            cnvPrt = "2"
'        ElseIf prt = "OUT1" Then
'            cnvPrt = "3"
'        ElseIf prt = "OUT2" Then
'            cnvPrt = "1"
'        End If
'     ElseIf shpType Like "*TRIPLE_BFN_SPLITTER*" Then
'
'        If prt = "IN1" Then
'            cnvPrt = "2"
'        ElseIf prt = "P1" Then
'            cnvPrt = "3"
'        ElseIf prt = "P2" Then
'            cnvPrt = "1"
'         ElseIf prt = "P3" Then
'            cnvPrt = "4"
'        End If

    ElseIf shpType Like "*COMM_DUAL*" Or shpType Like "*COMM_DUAL_IN_POL*" Or shpType Like "*DUAL_IN*" Then
        
        If prt = "H-TX" Then
            cnvPrt = "1"
        ElseIf prt = "RHCP" Then
            cnvPrt = "1"
        ElseIf prt = "V-TX" Then
            cnvPrt = "2"
        ElseIf prt = "LHCP" Then
            cnvPrt = "2"
        End If
     ElseIf shpType Like "*COMM_QUAD*" Then
        
        If prt = "RHCP" Then
            cnvPrt = "1"
        ElseIf prt = "LHCP" Then
            cnvPrt = "2"
        ElseIf prt = "V-TX" Then
            cnvPrt = "2"
        ElseIf prt = "V-RX" Then
            cnvPrt = "3"
        ElseIf prt = "H-RX" Then
            cnvPrt = "1"
        End If
        
      'splitter/6 ports on input side
       ElseIf shpType Like "*SPLITTER*" Or shpType Like "*SPLI*" Then
        If prt = "IN1" Then
            cnvPrt = 1 ' Directly assign the numeric value
        ElseIf InStr(1, prt, "OUT") > 0 Then
            ' Extract the numeric part of the OUT port
            parsePrt = Mid(prt, 4)
        ElseIf InStr(1, prt, "P") Then
            ' Extract the numeric part of the P port
            parsePrt = Mid(prt, 2)
            
            ' Ensure parsePrt is a valid numeric value
            If IsNumeric(parsePrt) Then
                cnvPrt = CInt(parsePrt) + 1
            Else
                cnvPrt = -1 ' Use a default error value or handle appropriately
            End If
        Else
            cnvPrt = -1 ' Handle unexpected port names
        End If
    
        ' Convert to string after calculation
        cnvPrt = CStr(cnvPrt)

    'imux/12 ports on input side
   ElseIf shpType Like "*IMUX*" Then
    If shpType Like "*MINI*" Then
        ' Special case for "MINI" IMUX shapes
        If prt = "IN1" Then
            cnvPrt = 2
        ElseIf prt = "OUT1" Then
            cnvPrt = 1
        End If
    Else
    Dim portCounter As Integer
    portCounter = 1 ' Start numbering from 1
    
    If prt = "IN1" Or prt = "OUT1" Then
        cnvPrt = "2" ' Reserve port 2 for special case
    Else
        cnvPrt = CStr(portCounter) ' Assign the current port number
        
        ' Increment portCounter while skipping 2
        If portCounter = 1 Then
            portCounter = 3 ' Skip 2 directly
        Else
            portCounter = portCounter + 1
            End If
         End If
        End If
    
'    ElseIf shpType Like "*CMUX*" Or shpType Like "*OMUX*" Then
'
'    'Dim portCounter As Integer
'    portCounter = 1 ' Start numbering from 1
'
'    If prt = "IN1" Or prt = "OUT1" Then
'        cnvPrt = "1" ' Reserve port 2 for special case
'    Else
'        cnvPrt = CStr(portCounter) ' Assign the current port number
'
'        ' Increment portCounter while skipping 2
'        If portCounter = 1 Then
'            portCounter = 3 ' Skip 2 directly
'        Else
'            portCounter = portCounter + 1
'        End If
'    End If

'
'    'omux/12 ports on output side
   ElseIf shpType Like "*OMUX*" Or shpType Like "*CMUX*" Then
   Dim portexist As Boolean
   
    ' Check if it has an OUT1 or "P" channels (P1, P3, P5, ...)
    If prt = "OUT1" Then
        portexist = True
        ' If it's OUT1, convert it to "1"
        cnvPrt = "1"
    ElseIf Left(prt, 1) = "P" And IsNumeric(Mid(prt, 2)) Then
        ' If it's a "P" channel, extract the number after "P"
        parsePrt = Mid(prt, 2)
        If CInt(parsePrt) = 1 And portexist = True Then
            cnvPrt = "2"
        ElseIf CInt(parsePrt) Mod 2 = 0 Then
            ' If the number is even, divide by 2
            cnvPrt = CStr(CInt(parsePrt) / 2)
        Else
            ' If the number is odd, add 1 and then divide by 2
            cnvPrt = CStr((CInt(parsePrt) + 1) / 2)
        End If
    End If
    
    ' Only assign if parsePrt was used
    If parsePrt <> "" Then
        cnvPrt = CStr(parsePrt)
    End If
'End If

        'cnvPrt = CStr(parsePrt)
    'End If

    ElseIf shpType Like "*ARROW*" Or shpType Like "*GND_*" Or shpType Like "*channel_post*" Then
        cnvPrt = "1"
    ElseIf shpType Like "*COUPLER*" Or shpType Like "*COUP*" Or shpType Like "* JUNCTION_BLOCK_COUP*" Or shpType Like "*dc_converter*" Or shpType Like "*DC_CONVERTER*" Then
        
        If prt = "IN-P1" Then
            cnvPrt = "2"
        ElseIf prt = "IN-P3" Then
            cnvPrt = "3"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        End If
    ElseIf shpType Like "*diplexer_combiner*" Then
        
        If prt = "IN-P1" Then
            cnvPrt = "2"
        ElseIf prt = "IN-P3" Then
            cnvPrt = "3"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        End If
    ElseIf shpType Like "*tcr_unit*" Then
        
        If prt = "IN1" Then
            cnvPrt = "2"
        ElseIf prt = "OUT2" Then
            cnvPrt = "3"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        End If
    ElseIf shpType Like "*XMTR*" Or shpType Like "*tcr_unit*" Then
        
        If prt = "RANGE0" Then
            cnvPrt = "3"
        ElseIf prt = "RANGE1" Then
            cnvPrt = "4"
        ElseIf prt = "HP" Then
            cnvPrt = "2"
        ElseIf prt = "LP" Then
            cnvPrt = "5"
        ElseIf prt = "BBE1" Then
            cnvPrt = "6"
        ElseIf prt = "BBE2" Then
            cnvPrt = "1"
        End If
        
    ElseIf shpType Like "*HYBRID*" Then
        
        If prt = "P1" Then
            cnvPrt = "1"
        ElseIf prt = "P2" Then
            cnvPrt = "2"
        ElseIf prt = "P3" Then
            cnvPrt = "3"
        ElseIf prt = "P4" Then
            cnvPrt = "4"
        Else
            cnvPrt = "-1"
        End If
    ElseIf shpType Like "*TCR_CMDRX*" Then
        
        If prt = "IN1" Then
            cnvPrt = "2"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        Else
            While InStr(1, prt, "NONE")
            cnvPrt = cnvPrt + 1
        Wend
        End If
    ElseIf shpType Like "*BBE*" Then
        
        If prt = "IN1" Then
            cnvPrt = "2"
        ElseIf prt = "IN2" Then
            cnvPrt = "3"
        ElseIf prt = "OUT1" Then
            cnvPrt = "1"
        ElseIf prt = "OUT2" Then
            cnvPrt = "4"
        Else
            While InStr(1, prt, "NONE")
            cnvPrt = cnvPrt + 1
        Wend
        End If
    ElseIf shpType Like "*ANTENNA*" Then
        
        If prt = "IN1" Then
            cnvPrt = "1"
        ElseIf prt = "IN2" Then
            cnvPrt = "2"
        ElseIf prt = "NONE" Then
            cnvPrt = "1"
        Else
            While InStr(1, prt, "NONE")
            cnvPrt = cnvPrt + 1
        Wend
     End If
     End If
    ConvertPort = cnvPrt

End Function


Public Sub GlueToShapes(shpStr1 As String, shpStr2 As String, dir1 As String, dir2 As String, rawStr As String, prt1 As String, prt2 As String)
    Dim myConnector As Visio.shape

    ' drop it somewhere
    Set myConnector = ActiveWindow.Page.Drop(Application.ConnectorToolDataObject, 1, 10)

    Dim shp As Visio.shape
    Dim subShp As Visio.shape

    
    Dim shp1 As Visio.shape
    Dim shp2 As Visio.shape
    Dim row1 As Integer
    Dim row2 As Integer
    Dim shpsFound As Boolean
    For Each shp In ActivePage.Shapes

            Dim isSet1 As Boolean
            Dim isSet2 As Boolean
            
            If shp.name = shpStr1 Then
                Set shp1 = shp
                isSet1 = True
                'Debug.Print "TRUE1 " + shp.Name + " " + shpStr1
             ElseIf shp.name = shpStr2 Then
                Set shp2 = shp
                isSet2 = True
                'Debug.Print "TRUE2 " + shp.Name + " " + shpStr2
            End If

            If isSet1 And isSet2 Then
            'Debug.Print "**************************** bothSet"
            shpsFound = True
                Exit For
            End If
    Next
    
    If shpsFound Then
        row1 = GetConnectionRowNum(dir1, prt1, shp1.name, shp1)
        row2 = GetConnectionRowNum(dir2, prt2, shp2.name, shp2)
        
        'Debug.Print "***** row info " _
        '; " Row 1 " + CStr(row1) + " Row 2 " + CStr(row2) + " shp1 " + shpStr1 + " shp2 " + shpStr2 + _
        'CStr(shp1.CellsSRC(visSectionConnectionPts, row1, 0)) + " " + CStr(shp2.CellsSRC(visSectionConnectionPts, row2, 0))
        Dim shpExist1 As Boolean
        Dim shpExist2 As Boolean
        If shp1.CellsSRCExists(visSectionConnectionPts, row1, 0, False) Then
            shpExists1 = True
            'Debug.Print shp1.CellsSRC(visSectionConnectionPts, row1, 0)
        End If
        If shp2.CellsSRCExists(visSectionConnectionPts, row2, 0, False) Then
            shpExist2 = True
            'Debug.Print shp2.CellsSRC(visSectionConnectionPts, row2, 0)
        End If
        
        If shp1.CellsSRCExists(visSectionConnectionPts, row1, 0, False) And shp2.CellsSRCExists(visSectionConnectionPts, row2, 0, False) Then

            
            'Debug.Print "GLUETO " + " Row 1 " + CStr(row1) + " Row 2 " + CStr(row2) + " shp1 " + shpStr1 + " shp2 " + shpStr2
            myConnector.Cells("BeginX").GlueTo shp1.CellsSRC(visSectionConnectionPts, row1, 0)
            myConnector.Cells("EndX").GlueTo shp2.CellsSRC(visSectionConnectionPts, row2, 0)
            Dim irow As Integer

            Dim rowCount As Integer
            rowCount = 1
            'x-rd
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-rd"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            vsoCell.Formula = Chr$(34) & myConnector.name & Chr$(34)
            'x-edge-type
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-edge-type"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            vsoCell.Formula = """Dynamic connector"""
            'x-parent
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-source-parent-rd"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            vsoCell.Formula = Chr$(34) & shp1.name & Chr$(34)
            'x-source-value
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-source-value"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            Dim sPrt As String
            sPrt = ConvertPort(prt1, shp1.name)
            vsoCell.Formula = Chr$(34) & sPrt & Chr$(34)
            'x-target
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-target-parent-rd"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            vsoCell.Formula = Chr$(34) & shp2.name & Chr$(34)
            'x-target-value
            irow = myConnector.AddNamedRow(visSectionProp, rowName(rowCount), 0)
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsLabel)
            vsoCell.Formula = """x-target-value"""
            Set vsoCell = myConnector.CellsSRC(visSectionProp, irow, visCustPropsValue)
            Dim tPrt As String
            tPrt = ConvertPort(prt2, shp2.name)
            vsoCell.Formula = Chr$(34) & tPrt & Chr$(34)
            
                  If shpStr1 Like "*HOT_ARROW*" Then
            ' Check if "Trace" row exists before adding it
            If Not shp1.CellExistsU("Prop.Trace", visExistsLocally) Then
                Dim ro1 As Integer
                Dim ro1cell As Visio.Cell
                ro1 = shp1.AddNamedRow(visSectionProp, "Trace", visTagDefault)
                Set ro1cell = shp1.CellsSRC(visSectionProp, ro1, visCustPropsValue)
                ro1cell.Formula = Chr$(34) & myConnector.name & Chr$(34)
            End If
             End If
        
        If shpStr2 Like "*HOT_ARROW*" Then
            ' Check if "Trace" row exists before adding it
            If Not shp2.CellExistsU("Prop.Trace", visExistsLocally) Then
                Dim ro2 As Integer
                Dim ro2cell As Visio.Cell
                ro2 = shp2.AddNamedRow(visSectionProp, "Trace", visTagDefault)
                Set ro2cell = shp2.CellsSRC(visSectionProp, ro2, visCustPropsValue)
                ro2cell.Formula = Chr$(34) & myConnector.name & Chr$(34)
            End If
        End If

        Else
            Debug.Print CStr(row1) + "  " + CStr(row2)
            Debug.Print "** ErRR **** " + shp1.name + " " + shp2.name + " " + CStr(shpExist1) + CStr(shpExist2) + "::" + CStr(shp1.CellsSRCExists(visSectionConnectionPts, row1, 0, False)) + CStr(shp2.CellsSRCExists(visSectionConnectionPts, row2, 0, False))
        End If


    Else
        'MsgBox "Could Not Find Shape " + rawString + CStr(isSet1) + " " + CStr(isSet2)
    End If
    


End Sub



Function GetConnectionRowNum(dir1 As String, prt As String, shpType As String, shp As shape)
    Dim row As Integer
    Dim parsePrt As String
    'Debug.Print "shape " + shpType
    'standard 2 or 4 port object
    If shpType Like "*tcr_V*" Then
        
        If prt = "P1" Then
            row = 0
        ElseIf prt = "P2" Then
            row = 1
        ElseIf prt = "P3" Then
            row = 2
        ElseIf prt = "P4" Then
            row = 3
        End If
    ElseIf shpType Like "*COMM_DUAL*" Then
        If InStr(1, dir1, "BOTTOM") <> 0 Then
            row = 0
        ElseIf InStr(1, dir1, "TOP") <> 0 Then
            row = 1
        Else
            row = -1
        End If
    ElseIf shpType Like "*SWITCH*" _
        Or shpType Like "*CAMP*" _
        Or shpType Like "*CHAMP*" _
        Or shpType Like "*DUAL_IN*" _
        Or shpType Like "*COMM_QUAD*" _
        Or shpType Like "*GND_*" _
        Or shpType Like "*CHANNEL_POST*" _
        Or shpType Like "*COUPLER*" _
        Or shpType Like "COUP*" _
        Or shpType Like "*diplexer_combiner*" _
    Then
    'Debug.Print "HERE  " + dir1
        If InStr(1, dir1, "BOTTOM") <> 0 Then
            row = 2
        ElseIf InStr(1, dir1, "TOP") <> 0 Then
            row = 3
        ElseIf InStr(1, dir1, "LEFT") <> 0 Then
            row = 0
        ElseIf InStr(1, dir1, "RIGHT") <> 0 Then
            row = 1
        Else
            row = -1
        End If
    'EPC has 2 on left 2 on right
    ElseIf shpType Like "*EPC*" Then
        
        If prt = "INA" Or prt = "IN1" Then
            row = 0
        ElseIf prt = "INB" Or prt = "IN2" Then
            row = 1
        ElseIf prt = "OUTA" Or prt = "OUT1" Then
            row = 2
        ElseIf prt = "OUTB" Or prt = "OUT2" Then
            row = 3
        End If
     'DOWN_CONVERTER has 1 on left 1 on right
    ElseIf shpType Like "*DOWN_CONVERTER*" Or shpType Like "*LNA*" Then
        
        If prt = "IN1" Then
            row = 0
        ElseIf prt = "OUT1" Then
            row = 1
        End If
      ElseIf shpType Like "*TCR_UNIT*" Then
        
        If prt = "IN1" Then
            row = 0
        ElseIf prt = "OUT1" Then
            row = 1
        ElseIf prt = "OUT2" Then
            row = 2
        End If

    'EPC has 1 on left 1 on right
   ElseIf shpType Like "*TWTA*" Or shpType Like "*LCHAMP*" Or shpType Like "*DUAL_OUTPUT_RECEIVER*" Or shpType Like "*SPLITTER*" _
    Or shpType Like "*T_RECEIVER*" Or shpType Like "*spli*" Or shpType Like "*JUNCTION_BLOCK_SPLI*" Then
        
    If prt = "IN1" Then
        row = 0
    ElseIf InStr(1, prt, "OUT") > 0 Then
        ' Extract the numeric part from "OUTx"
        Dim portNumber As String
        portNumber = Mid(prt, 4)
          
        ' Ensure the extracted part is numeric and calculate the row index
        If IsNumeric(portNumber) Then
            row = CInt(portNumber)
        Else
            row = -1 ' Handle unexpected or invalid port names
        End If
    ElseIf InStr(1, prt, "P") > 0 Then
        ' Handle cases for ports like "P1", "P2", etc.
        Dim pNumber As String
        pNumber = Mid(prt, 2)
        
        ' Ensure the extracted part is numeric and calculate the row index
        If IsNumeric(pNumber) Then
            row = CInt(pNumber)
        Else
            row = -1 ' Handle unexpected or invalid port names
        End If
    Else
        row = -1 ' Default for unknown ports
    End If

        'antenna has 1 on bottom 1 on top
    ElseIf shpType Like "*COMM_DUAL*" Or shpType Like "*COMM_DUAL_IN_POL*" Or shpType Like "*DUAL_CONNECT*" Or shpType Like "*DUAL_IN*" Then
        
        If prt = "RHCP" Then
            row = 0
        ElseIf prt = "LHCP" Then
            row = 1
        End If
    ElseIf shpType Like "*COMM_QUAD*" Then
        
        If prt = "RHCP" Then
            row = 0
        ElseIf prt = "LHCP" Then
            row = 1
        ElseIf prt = "V-RX" Then
            row = 1
        ElseIf prt = "H-RX" Then
            row = 2
        ElseIf prt = "V-TX" Then
            row = 3
        
        End If
    ElseIf shpType Like "*ANTENNA*" Then
        
        If prt = "NONE" Then
            row = 0
        ElseIf prt = "IN1" Then
            row = 0
        ElseIf prt = "IN2" Then
            row = 1
        Else
            While InStr(1, prt, "NONE")
            row = row + 1
        Wend
        End If
     'splitter/6 ports on input side
'    ElseIf shpType Like "*SIX_BFN_SPLITTER*" Or shpType Like "*SIX_BFN_SPLITTER*" Then
'        'has an IN1 and 12 "P" channels (p1,p3,p5...)
'        If prt = "IN1" Then
'            row = 0
'        ElseIf InStr(1, prt, "P") Then
'            parsePrt = Mid(prt, 2)
'            If CInt(parsePrt) Mod 2 = 0 Then
'                row = CInt(parsePrt) / 2
'            Else
'                row = (CInt(parsePrt) + 1) / 2
'            End If
'
'        End If
    'imux/12 ports on input side
    ElseIf shpType Like "*IMUX*" Then
    ' Handle shapes with IN1 and 12 "P" channels (e.g., P1, P3, P5, etc.)
    If shpType Like "*MINI*" Then
        ' Special case for "MINI" IMUX shapes
        If prt = "IN1" Then
            row = 0
        ElseIf prt = "OUT1" Then
            row = 1
        End If
    Else
        ' General IMUX shape handling
        If prt = "IN1" Or prt = "OUT1" Then
            row = 0
        ElseIf InStr(1, prt, "P") > 0 Then
            ' Parse the port number and calculate the row based on channel semicolons
            Dim channelsValue As String
            Dim originalLength As Integer
            Dim modifiedLength As Integer
            Dim index As Integer
            Dim smc As Integer
            Dim ports As Integer
            'Dim parsePrt As String
            
            ' Extract the numeric part of the port name
            parsePrt = Mid(prt, 2)
            
            ' Get channel names from shape properties
            channelsValue = shp.CellsU("Prop.Channel_Name").ResultStr("")
            
            ' Find the position of the current port in the channel list
            index = InStr(channelsValue, parsePrt)
            
            ' Calculate the semicolon count up to the current index
            originalLength = Len(Left(channelsValue, index))
            modifiedLength = Len(Replace(Left(channelsValue, index), ";", ""))
            smc = originalLength - modifiedLength
            
            ' Calculate the total number of semicolons in the entire channel list
            originalLength = Len(channelsValue)
            modifiedLength = Len(Replace(channelsValue, ";", ""))
            ports = originalLength - modifiedLength
            
            ' Determine the row index
            row = ports - smc + 1
            Else
                ' Default case for unhandled ports
                row = -1 ' Or any appropriate value indicating an error or invalid input
            End If
        End If

    'omux/12 ports on output side
    ElseIf shpType Like "*OMUX*" Or shptyp Like "*CMUX*" Then
        'has an OUT1 and 12 "P" channels (p1,p3,p5...)
        If prt = "OUT1" Then
            row = 0
        ElseIf InStr(1, prt, "P") Then
            
            Dim channelsValue2 As String
            Dim originalLength2 As Integer
            Dim modifiedLength2 As Integer
            Dim index2 As Integer
            Dim smc2 As Integer
            Dim ports2 As Integer
           
            parsePrt = Mid(prt, 2)
            channelsValue2 = shp.CellsU("Prop.Channel_Name").ResultStr("")
            index2 = InStr(channelsValue2, parsePrt)
            originalLength2 = Len(Left(channelsValue2, index2))
            modifiedLength2 = Len(Replace(Left(channelsValue2, index2), ";", ""))
            smc2 = originalLength2 - modifiedLength2
            
            originalLength2 = Len(channelsValue2)
            modifiedLength2 = Len(Replace(channelsValue2, ";", ""))
            ports2 = originalLength2 - modifiedLength2
            
            row = ports2 - smc2 + 1
            
        End If
    ElseIf shpType Like "*ARROW*" Then
        'row = 0
        Dim connToValue As String
        connToValue = shp.CellsU("Prop.ConnTo").ResultStr("")
        If connToValue = "top" Then
            row = 3
        ElseIf connToValue = "bottom" Then
            row = 2
        ElseIf connToValue = "right" Then
            row = 1
        Else
            row = 0
        End If

    ElseIf shpType Like "*XMTR*" Then
        
        If prt = "RANGE0" Then
            row = 0
        ElseIf prt = "RANGE1" Then
            row = 1
        ElseIf prt = "LP" Then
            row = 2
        ElseIf prt = "HP" Then
            row = 3
        ElseIf prt = "BBE1" Then
            row = 4
        ElseIf prt = "BBE2" Then
            row = 5
        End If
        
    ElseIf shpType Like "*HYBRID*" Then
        
        If prt = "P1" Then
            row = 0
        ElseIf prt = "P4" Then
            row = 1
        ElseIf prt = "P2" Then
            row = 2
        ElseIf prt = "P3" Then
            row = 3
        Else
            row = -1
        End If
    ElseIf shpType Like "*TCR_CMDRX*" Then
        
        If prt = "IN1" Then
            row = 0
        ElseIf prt = "OUT1" Then
            row = 1
        Else
            While InStr(1, prt, "NONE")
            row = row + 1
        Wend
        End If
    ElseIf shpType Like "*BBE*" Then
        
        If prt = "IN1" Then
            row = 0
        ElseIf prt = "IN2" Then
            row = 1
        ElseIf prt = "OUT1" Then
            row = 2
        ElseIf prt = "OUT2" Then
            row = 3
        Else
            row = -1
        End If
    End If


    'Debug.Print "The row is **** " + CStr(row) + " " + dir1
    GetConnectionRowNum = row

End Function

Function rowName(ByRef count As Integer)
    rname = "Row_" & count
    count = count + 1
    rowName = rname
End Function
