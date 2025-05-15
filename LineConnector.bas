Attribute VB_Name = "LineConnector"

Public Sub RunLines()
    AdjustSwitchShapes
    AddConnectionPoints
    GetConnectionsFromBackground
    AddTextField
    
End Sub
 Sub AdjustSwitchShapes()
    Dim shp As Visio.Shape
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
    Dim shp As Visio.Shape
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

         ElseIf (strName Like "*beacon*") Or (strName Like "*power_monitor*") Or (strName Like "*uhf_bps*") Then
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
   
         ElseIf (strName Like "*dual_bfn_splitter*") Or (strName Like "*epic_filter_splitter*") Then
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
            
        ElseIf (strName Like "*quad_bfn_splitter*") Or (strName Like "*quad_bfn_coupler*") Then
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
            
         ElseIf (strName Like "*epic_omt_quad_plexer*") Then
             ' Get dimensions of where the points should go.
             nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
             nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
             x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
             y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
      
             
             ' Add 2 equidistant connections to the right (1/3 and 2/3 height).
             NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
             shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
             shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight / 3)) + " pt."
             
             NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
             shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
             shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight / 3)) + " pt."
             
             ' Add 1 connection at the top-center.
             NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
             shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
             shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
             
             ' Add 1 connection at the bottom-center.
             NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
             shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
             shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
             
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

            
            ElseIf (strName Like "*dual_bfn_coupler*") Or (strName Like "*dual_coupler*") Or (strName Like "*epic_ul_filter_coupler*") Then
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
            
            ElseIf (strName Like "*triple_bfn_coupler*") Then
            ' Get dimensions of where the points should go.
            nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
            nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
            x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
            y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
        
            ' Add 3 connections to the left.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.3)) + " pt."
        
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.3)) + " pt."
        
            ' Add 1 connection to the right.
            NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
            shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
            shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
        
            shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf


        ' Add connection points to the four sides of all switches, circulators, loads and satpad shapes.
       ElseIf (strName Like "*circulator*") Or (strName Like "*load*") Or (strName Like "*switch*") Or (strName Like "*iot_coupler*") _
       Or (strName Like "*ats*") Or (strName Like "*satpad*") Or (strName Like "*hot-arrow*") Or (strName Like "*hot_arrow*") _
       Or (strName Like "*CHANNEL_POST*") Or (strName Like "*channel_post*") Or (strName Like "*diplexer*") Or (strName Like "*epic_junction_block*") Or (strName Like "*junction_block_spli*") Or (strName Like "*gnd_*") Or (strName Like "*hi_power_mode_splitter*") _
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

         ElseIf (strName Like "*receiver*") Or (strName Like "*xmitrs*") Or (strName Like "*epic_tpam*") Then
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
            ElseIf (strName Like "*camp*") Or (strName Like "*epic_dl_filter*") Or (strName Like "*uhf_mlo*") Or (strName Like "*champ*") Or (strName Like "*twta*") Or (strName Like "*ka_downconverter*") Or (strName Like "*down_converter*") Or (strName Like "*lna*") Then
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
           ElseIf (strName Like "*camp*") Or (strName Like "*champ*") Or (strName Like "*twta*") Or (strName Like "*boeing_epic_rtn_fil*") Then
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
            
            ElseIf (strName Like "*boeing_tlm_xmtr*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' === Add 2 connections to the left ===
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.25)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.25)) + " pt."
            
                ' === Add 5 connections to the right ===
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.4)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.2)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.2)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.4)) + " pt."
            
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
                
           ElseIf (strName Like "*boeing_cmd_rx*") Then
                ' Get dimensions of where the points should go.
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' === Add 1 connection to the left ===
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' === Add 4 connections to the right ===
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.3)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.1)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.1)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.3)) + " pt."
            
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf

            ElseIf (strName Like "*xmtr*") And Not (strName Like "*boeing_tlm_xmtr*") Then
                
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

            ElseIf (strName Like "*hybrid*") Or (strName Like "*epic_mpa_2x2*") Or (strName Like "*epic_omt*") Then
            'Debug.Print "HYBRID"
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
                
            ElseIf (strName Like "*epic_8x8_output_mux*") Or (strName Like "*epic_mpa_8x8*") Then

                ' Get dimensions of where the points should go
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Calculate spacing for 8 points
                Spacing = nheight / 9 ' 8 spaces => 9 gaps
                For i = 1 To 8
                    OffsetY = y + (nheight / 2) - (i * Spacing)
            
                    ' Right side
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth / 2)) & " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(OffsetY) & " pt."
            
                    ' Left side
                    NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x - (nwidth / 2)) & " pt."
                    shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(OffsetY) & " pt."
                Next i
            
                shpsAddded = shpsAdded + vbCrLf + strName + vbCrLf
                
                
            ElseIf (strName Like "*dual_epic_rtn*") Then
            
                ' Get shape dimensions
                nwidth = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).Result(visPoints)
                nheight = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).Result(visPoints)
                x = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).Result(visPoints)
                y = shp.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).Result(visPoints)
            
                ' Add mid-right connection point
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y) + " pt."
            
                ' Add two equidistant points to the right
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.33)) + " pt."
            
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x + (nwidth * 0.5)) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.33)) + " pt."
            
                ' Add left connection points (unchanged)
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
                
            ElseIf (strName Like "*epic_omt_quad_plexer*") Then
            
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

                ' Add connections to top and bottom.
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y - (nheight * 0.5)) + " pt."
    
                NewRow = shp.AddRow(visSectionConnectionPts, visRowLast, visTagDefault)
                shp.CellsSRC(visSectionConnectionPts, NewRow, visX).Formula = CStr(x) + " pt."
                shp.CellsSRC(visSectionConnectionPts, NewRow, visY).Formula = CStr(y + (nheight * 0.5)) + " pt."
                
                shpsAdded = shpsAdded + vbCrLf + strName + vbCrLf
                
            ElseIf (strName Like "*epic_omt_rh_lh*") Then
            
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
        Dim shp As Visio.Shape
        
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
        Dim shp As Visio.Shape
        For Each shp In Visio.ActivePage.Shapes

        
            If InStr(1, LCase(shp.name), "readout") <> 0 Then
                shp.Text = "-"
                shp.Cells("Char.Size").FormulaForceU = "7pt"
                
                ' Align text to the left
            shp.CellsU("Para.HorzAlign").FormulaForceU = "0"
            End If
        Next

End Sub


Function ConvertPort(prt As String, shpType As String) As String
    Dim cnvPrt As String
    Dim parsePrt As String
    Dim i As Integer

    cnvPrt = "-1" ' Default fallback

    Select Case True
        Case shpType Like "*SWITCH*", shpType Like "*tcr_toggle_inp*"
            Select Case prt
                Case "P1": cnvPrt = "1"
                Case "P2": cnvPrt = "2"
                Case "P3": cnvPrt = "3"
                Case "P4": cnvPrt = "4"
            End Select

        Case shpType Like "*EPC*"
            Select Case prt
                Case "INA", "IN1": cnvPrt = "2"
                Case "INB", "IN2": cnvPrt = "4"
                Case "OUTA", "OUT1": cnvPrt = "1"
                Case "OUTB", "OUT2": cnvPrt = "3"
            End Select

        Case shpType Like "*CMD_RX*"
            Select Case prt
                Case "INA", "IN1": cnvPrt = "2"
                Case "RNG-OUT1": cnvPrt = "4"
                Case "RNG-OUT2": cnvPrt = "5"
                Case "CMD-OUT1", "OUT1": cnvPrt = "1"
                Case "CMD-OUT2", "OUT2": cnvPrt = "3"
            End Select

        Case shpType Like "*TLM_XMTR*"
            Select Case prt
                Case "OUT1": cnvPrt = "1"
                Case "OUT2": cnvPrt = "2"
                Case "CMD-OUT1": cnvPrt = "3"
                Case "CMD-OUT2": cnvPrt = "4"
                Case "CMD-OUT3": cnvPrt = "5"
                Case "RNG-OUT1": cnvPrt = "6"
                Case "RNG-OUT2": cnvPrt = "7"
            End Select

        Case shpType Like "*TWTA*" Or shpType Like "*EPIC_DL_FILTER*" Or shpType Like "*UHF_MLO*" Or shpType Like "*CHAMP*" Or shpType Like "*CAMP*" Or _
             shpType Like "*DOWN_CONVERTER*" Or shpType Like "*receiver*" Or shpType Like "*BOEING_EPIC_RTN_FIL*" Or shpType Like "*RECEIVER*" Or _
             shpType Like "*DUAL_OUTPUT_RECEIVER*" Or shpType Like "*T_RECEIVER*" Or shpType Like "*LNA*"
            Select Case prt
                Case "IN1": cnvPrt = "2"
                Case "OUT1": cnvPrt = "1"
                Case "OUT2": cnvPrt = "3"
            End Select

        Case shpType Like "*XMITRS*"
            Select Case prt
                Case "IN1", "OUT1": cnvPrt = "1"
                Case "IN2", "OUT2", "NONE": cnvPrt = "2"
            End Select
            
         Case shpType Like "*EPIC_TPAM*"
            Select Case prt
                Case "IN1": cnvPrt = 2
                Case "OUT1": cnvPrt = 1
            End Select

        Case shpType Like "*DUAL_EPIC_RTN*"
            Select Case prt
                Case "IN1": cnvPrt = "2"
                Case "OUT1": cnvPrt = "1"
                Case "IN2": cnvPrt = "3"
            End Select

        Case shpType Like "*COMM_DUAL*" Or shpType Like "*COMM_DUAL_IN_POL*" Or shpType Like "*DUAL_IN*"
            Select Case prt
                Case "H-TX", "RHCP": cnvPrt = "1"
                Case "V-TX", "LHCP": cnvPrt = "2"
            End Select

        Case shpType Like "*EPIC_OMT_QUAD_PLEXER*"
            Select Case prt
                Case "TX-RHCP", "V-TX": cnvPrt = "1"
                Case "TX-LHCP", "V-RX": cnvPrt = "3"
                Case "RX-RHCP": cnvPrt = "4"
                Case "RX-LHCP": cnvPrt = "2"
            End Select

        Case shpType Like "*COMM_QUAD*"
            Select Case prt
                Case "RHCP", "V-TX": cnvPrt = "1"
                Case "LHCP", "V-RX": cnvPrt = "3"
                Case "H-TX": cnvPrt = "4"
                Case "H-RX": cnvPrt = "2"
            End Select
         Case shpType Like "*EPIC_OMT_QUAD_PLEXER*"
            Select Case prt
                Case "RHCP", "V-TX": cnvPrt = "1"
                Case "LHCP", "V-RX": cnvPrt = "3"
                Case "H-TX": cnvPrt = "4"
                Case "H-RX": cnvPrt = "2"
            End Select
         Case shpType Like "*EPIC_OMT_RH_LH*"
            Select Case prt
                Case "RHCP", "V-TX": cnvPrt = 1
                Case "LHCP", "H-TX": cnvPrt = 2
            End Select

        Case shpType Like "*COUPLER*" Or shpType Like "*COUP*" Or shpType Like "*EPIC_JUNCTION_BLOCK*" Or shpType Like "*JUNCTION_BLOCK_COUP*" Or shpType Like "*EPIC_UL_FILTER_COUPLER*"
            Select Case prt
                Case "OUT1": cnvPrt = "1"
                Case "IN-P1", "P1", "IN1": cnvPrt = "2"
                Case "IN-P3", "P2", "IN2": cnvPrt = "3"
                Case "P3": cnvPrt = "4"
                Case "P4": cnvPrt = "5"
            End Select

        Case shpType Like "*SPLITTER*" Or shpType Like "*SPLI*" Or shpType Like "*block_splitter*"
            If prt = "IN1" Then
                cnvPrt = "1"
            ElseIf prt Like "P#" Then
                cnvPrt = CStr(Val(Mid(prt, 2)) + 1)
            ElseIf prt Like "OUT#" Then
                cnvPrt = CStr(Val(Mid(prt, 4)) + 1)
            End If

        Case shpType Like "*IMUX*" Or shpType Like "*EPIC_8X8_OUTPUT_MUX*"
            If shpType Like "*MIN*" Then
                If prt = "OUT1" Or prt = "IN1" Then
                    cnvPrt = "2"
                Else
                    cnvPrt = "1"
                End If
            Else
                If prt = "OUT1" Or prt = "IN1" Then
                    cnvPrt = "1"
                ElseIf prt Like "P#" Then
                    cnvPrt = Mid(prt, 2)
                End If
            End If

        Case shpType Like "*OMUX*", shpType Like "*CMUX*"
            If prt = "OUT1" Then
                cnvPrt = "1"
            ElseIf prt Like "P#" Then
                cnvPrt = Mid(prt, 2)
            End If

        Case shpType Like "*dc_converter*" Or shpType Like "*DC_CONVERTER*"
            Select Case prt
                Case "IN-P1": cnvPrt = "2"
                Case "IN-P3": cnvPrt = "3"
                Case "OUT1": cnvPrt = "1"
            End Select

        Case shpType Like "*beacon*", shpType Like "*BEACONS*", shpType Like "*POWER_MONITOR*", shpType Like "*UHF_BPS*"
            If prt = "NONE" Or prt = "OUT1" Then
                cnvPrt = "1"
            ElseIf prt Like "P*" Then
                cnvPrt = "2"
            End If

        Case shpType Like "*ANTENNA*"
            Select Case prt
                Case "IN1", "NONE": cnvPrt = "1"
                Case "IN2": cnvPrt = "2"
            End Select

        Case shpType Like "*ARROW*", shpType Like "*CHANNEL_POST*", shpType Like "*GND*", shpType Like "*LOAD_PASS_THRU*"
            If prt = "NONE" Then cnvPrt = "1"

        Case shpType Like "*HYBRID*", shpType Like "*HYBRID_MUX*", shpType Like "*EPIC_MPA_2x2*", shpType Like "*EPIC_OMT_H_V*"
            Select Case prt
                Case "P1", "H-POL": cnvPrt = "1"
                Case "P2", "V-POL": cnvPrt = "2"
                Case "P3": cnvPrt = "3"
                Case "P4": cnvPrt = "4"
            End Select

        Case shpType Like "*diplexer*"
            Select Case prt
                Case "IN-P1": cnvPrt = "2"
                Case "IN-P3": cnvPrt = "3"
                Case "OUT1": cnvPrt = "1"
            End Select

        Case shpType Like "*tcr_unit*", shpType Like "*TCR_UNIT*"
            Select Case prt
                Case "IN1": cnvPrt = "2"
                Case "OUT1": cnvPrt = "1"
                Case "OUT2": cnvPrt = "3"
            End Select

        Case shpType Like "*XMTR*" And Not shpType Like "*TLM_XMTR*"
            Select Case prt
                Case "RANGE0": cnvPrt = "3"
                Case "RANGE1": cnvPrt = "4"
                Case "HP": cnvPrt = "2"
                Case "LP": cnvPrt = "5"
                Case "BBE1": cnvPrt = "6"
                Case "BBE2": cnvPrt = "1"
            End Select

        Case shpType Like "*TCR_CMDRX*"
            Select Case prt
                Case "IN1": cnvPrt = "2"
                Case "OUT1": cnvPrt = "1"
                Case "NONE": cnvPrt = "1"
            End Select

        Case shpType Like "*BBE*"
            Select Case prt
                Case "IN1": cnvPrt = "2"
                Case "IN2": cnvPrt = "3"
                Case "OUT1": cnvPrt = "1"
                Case "OUT2": cnvPrt = "4"
            End Select

    End Select

    ConvertPort = cnvPrt
End Function

Public Sub GlueToShapes(shpStr1 As String, shpStr2 As String, dir1 As String, dir2 As String, rawStr As String, prt1 As String, prt2 As String)
    Dim myConnector As Visio.Shape

    ' drop it somewhere
    Set myConnector = ActiveWindow.Page.Drop(Application.ConnectorToolDataObject, 1, 10)

    Dim shp As Visio.Shape
    Dim subShp As Visio.Shape

    
    Dim shp1 As Visio.Shape
    Dim shp2 As Visio.Shape
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
Function GetConnectionRowNum(dir1 As String, prt As String, shpType As String, shp As Shape) As Integer
    Dim row As Integer
    Dim portNumber As String
    Dim parsePrt As String
    Dim connToValue As String
    Dim channelsValue As String
    Dim index As Integer
    Dim smc As Integer
    Dim ports As Integer

    row = -1 ' Default row value

    Select Case True
        ' Standard 2 or 4 port object
        Case shpType Like "*tcr_V*"
            Select Case prt
                Case "P1": row = 0
                Case "P2": row = 1
                Case "P3": row = 2
                Case "P4": row = 3
            End Select

        ' COMM_DUAL types
        Case shpType Like "*COMM_DUAL*"
            If InStr(1, dir1, "BOTTOM") > 0 Then row = 0
            If InStr(1, dir1, "TOP") > 0 Then row = 1

        ' General SWITCH, CAMP, etc.
        Case shpType Like "*SWITCH*" Or shpType Like "*CAMP*" Or shpType Like "*CHAMP*" _
            Or shpType Like "*DUAL_IN*" Or shpType Like "*GND_*" Or shpType Like "*LOAD_PASS_THRU*" Or shpType Like "*CHANNEL_POST*" _
            Or shpType Like "*diplexer*" Or shpType Like "*tcr_toggle_inp*"
            Select Case True
                Case InStr(1, dir1, "BOTTOM") > 0: row = 2
                Case InStr(1, dir1, "TOP") > 0: row = 3
                Case InStr(1, dir1, "LEFT") > 0: row = 0
                Case InStr(1, dir1, "RIGHT") > 0: row = 1
            End Select

        ' EPC type
        Case shpType Like "*EPC*"
            Select Case prt
                Case "INA", "IN1": row = 1
                Case "INB", "IN2": row = 0
                Case "OUTA", "OUT1": row = 3
                Case "OUTB", "OUT2": row = 2
            End Select

        ' CMD_RX
        Case shpType Like "*CMD_RX*"
            Select Case prt
                Case "INA", "IN1": row = 1
                Case "RNG-OUT1": row = 3
                Case "RNG-OUT2": row = 4
                Case "CMD-OUT1": row = 0
                Case "CMD-OUT2": row = 2
            End Select

        ' TLM_XMTR
        Case shpType Like "*TLM_XMTR*"
            Select Case prt
                Case "OUT1": row = 0
                Case "OUT2": row = 1
                Case "CMD-OUT1": row = 2
                Case "CMD-OUT2": row = 3
                Case "CMD-OUT3": row = 4
                Case "RNG-OUT1": row = 5
                Case "RNG-OUT2": row = 6
            End Select

        ' DOWN_CONVERTER or LNA
        Case shpType Like "*DOWN_CONVERTER*" Or shpType Like "*LNA*"
            If prt = "IN1" Then row = 0
            If prt = "OUT1" Then row = 1

        ' XMITR (not TLM_XMITR)
        Case shpType Like "*XMITR*" And Not shpType Like "*TLM_XMITR*"
            Select Case prt
                Case "IN1", "OUT1": row = 1
                Case "OUT2", "NONE": row = 0
            End Select
        Case shpType Like "*EPIC_TPAM*"
            Select Case prt
                Case "IN1": row = 1
                Case "OUT1": row = 0
            End Select

        ' DUAL_EPIC_RTN
        Case shpType Like "*DUAL_EPIC_RTN*"
            Select Case prt
                Case "IN1": row = 1
                Case "OUT1": row = 0
                Case "IN2": row = 2
            End Select

        ' TCR_UNIT
        Case shpType Like "*TCR_UNIT*"
            Select Case prt
                Case "IN1": row = 0
                Case "OUT1": row = 1
                Case "OUT2": row = 2
            End Select

        ' TWTA, CHAMP, etc.
        Case shpType Like "*TWTA*" Or shpType Like "*EPIC_DL_FILTER*" Or shpType Like "*UHF_MLO*" _
            Or shpType Like "*CHAMP*" Or shpType Like "*DUAL_OUTPUT_RECEIVER*" _
            Or shpType Like "*T_RECEIVER*" Or shpType Like "*RECEIVER*" Or shpType Like "*BOEING_EPIC_RTN*"
            If prt = "IN1" Then
                row = 0
            ElseIf InStr(prt, "OUT") > 0 Then
                portNumber = Mid(prt, 4)
                If IsNumeric(portNumber) Then row = CInt(portNumber)
            ElseIf InStr(prt, "P") > 0 Then
                portNumber = Mid(prt, 2)
                If IsNumeric(portNumber) Then row = CInt(portNumber)
            End If

        ' COMM_QUAD
        Case shpType Like "*COMM_QUAD*"
            Select Case prt
                Case "V-RX": row = 3
                Case "LHCP", "V-TX": row = 1
                Case "H-TX": row = 2
                Case "RHCP", "H-RX": row = 0
            End Select

        ' EPIC_OMT_QUAD_PLEXER
        Case shpType Like "*EPIC_OMT_QUAD_PLEXER*"
            Select Case prt
                Case "TX-RHCP", "V-TX": row = 0
                Case "TX-LHCP", "H-TX": row = 3
                Case "RX-RHCP", "V-RX": row = 1
                Case "RX-LHCP", "H-RX": row = 2
            End Select
        Case shpType Like "*EPIC_OMT_RH_LH*"
            Select Case prt
                Case "RHCP", "V-TX": row = 0
                Case "LHCP", "H-TX": row = 1
            End Select
        ' CHANNEL_POST
        Case shpType Like "*CHANNEL_POST*"
            If prt = "NONE" Then row = 0

        ' BEACON, POWER_MONITOR, UHF_BPS
        Case shpType Like "*BEACON*" Or shpType Like "*POWER_MONITOR*" Or shpType Like "*UHF_BPS*"
            If prt = "NONE" Or prt = "P*" Or prt = "OUT1" Then row = 0

        ' ANTENNA
        Case shpType Like "*ANTENNA*"
            Select Case prt
                Case "NONE", "IN1": row = 0
                Case "IN2": row = 1
            End Select

        ' COUPLER types
        Case shpType Like "*COUPLER*" Or shpType Like "*COUP*" Or shpType Like "*EPIC_JUNCTION_BLOCK*" Or shpType Like "*JUNCTION_BLOCK_COUP*" _
            Or shpType Like "*EPIC_UL_FILTER_COUPLER*"
            Select Case prt
                Case "OUT1": row = 1
                Case "IN-P1", "P1", "IN1": row = 0
                Case "IN-P3", "P2", "IN2": row = 2
                Case "P3": row = 3
                Case "P4": row = 4
            End Select

        ' SPLITTER
        Case shpType Like "*SPLITTER*" Or shpType Like "*SPLI*" Or shpType Like "*block_splitter*"
            If prt = "IN1" Then
                row = 0
            ElseIf Left(prt, 1) = "P" And IsNumeric(Mid(prt, 2)) Then
                row = CInt(Mid(prt, 2))
            ElseIf Left(prt, 3) = "OUT" And IsNumeric(Mid(prt, 4)) Then
                row = CInt(Mid(prt, 4))
            End If

        ' IMUX, OMUX, CMUX
        Case shpType Like "*IMUX*" Or shpType Like "*OMUX*" Or shpType Like "*CMUX*" Or shpType Like "*EPIC_8X8_OUTPUT_MUX*"
            If prt = "OUT1" Or prt = "IN1" Then
                row = 0
            ElseIf InStr(prt, "P") > 0 Then
                parsePrt = Mid(prt, 2)
                channelsValue = shp.CellsU("Prop.Channel_Name").ResultStr("")
                index = InStr(channelsValue, parsePrt)
                smc = Len(Left(channelsValue, index)) - Len(Replace(Left(channelsValue, index), ";", ""))
                ports = Len(channelsValue) - Len(Replace(channelsValue, ";", ""))
                row = ports - smc + 1
            End If

        ' ARROW
        Case shpType Like "*ARROW*"
            connToValue = shp.CellsU("Prop.ConnTo").ResultStr("")
            Select Case connToValue
                Case "top": row = 3
                Case "bottom": row = 2
                Case "right": row = 1
                Case Else: row = 0
            End Select

        ' XMTR
        Case shpType Like "*XMTR*"
            Select Case prt
                Case "RANGE0": row = 0
                Case "RANGE1": row = 1
                Case "LP": row = 2
                Case "HP": row = 3
                Case "BBE1": row = 4
                Case "BBE2": row = 5
            End Select

        ' HYBRID types
        Case shpType Like "*HYBRID*" Or shpType Like "*HYBRID_MUX*" Or shpType Like "*EPIC_OMT_H_V*" Or shpType Like "*EPIC_MPA_2x2*"
            Select Case prt
                Case "P1", "H-POL": row = 0
                Case "P4", "V-POL": row = 1
                Case "P2": row = 2
                Case "P3": row = 3
            End Select

        ' TCR_CMDRX
        Case shpType Like "*TCR_CMDRX*"
            If prt = "IN1" Then row = 0
            If prt = "OUT1" Then row = 1

        ' BBE
        Case shpType Like "*BBE*"
            Select Case prt
                Case "IN1": row = 0
                Case "IN2": row = 1
                Case "OUT1": row = 2
                Case "OUT2": row = 3
            End Select
    End Select

    GetConnectionRowNum = row
End Function
Function rowName(ByRef count As Integer)
    rname = "Row_" & count
    count = count + 1
    rowName = rname
End Function
