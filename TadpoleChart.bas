Attribute VB_Name = "TadpoleChart"
Option Explicit

Sub createTadpoleChart()
'Creates a tadpole chart using two sets of data. The length of the tadpole _
 tails is determined by the number of rows in the data matrices. The number _
 of tadpoles is equal to the number of columns in the data matrices.


    Dim xDataRange, yDataRange As String
    xDataRange = "xAxisDataMatrix"
    yDataRange = "YAxisDataMatrix"


    Dim ChartObj As ChartObject
    Set ChartObj = ActiveSheet.ChartObjects.add(Left:=20, Width:=800, Top:=20, Height:=500)
    ChartObj.Chart.ChartType = xlXYScatterSmoothNoMarkers

    Dim teamId As Integer
    
    Dim teamCount, tailLength As Integer
    Dim xRng, yRng As range
    Dim ChartSeries As Series
    
    teamCount = range(xDataRange).Columns.Count
    tailLength = range(xDataRange).Rows.Count
    
    tailLength = 4
    
    For teamId = 1 To teamCount

        With range(xDataRange).Cells(0, 0)
            Set xRng = range(.offset(1, teamId), .offset(tailLength, teamId))
        End With
        With range(yDataRange).Cells(0, 0)
            Set yRng = range(.offset(1, teamId), .offset(tailLength, teamId))
        End With
        
        Debug.Print "----" & teamId & "----"
        Debug.Print "x Range: " & xRng.Address
        Debug.Print "y Range: " & yRng.Address
        

        
        Set ChartSeries = ChartObj.Chart.SeriesCollection.NewSeries
        With ChartSeries
            .XValues = xRng.Cells
            .values = yRng.Cells
            .Name = "Team " & teamId
            .Format.Line.DashStyle = msoLineSolid
            .Format.Line.Transparency = 0.25
        End With
        
        Dim headPoint As Point
        Set headPoint = ChartSeries.Points(1)
        headPoint.MarkerStyle = xlMarkerStyleDiamond
        headPoint.MarkerForegroundColor = ColorConstants.vbBlack
        headPoint.MarkerSize = 8
        
    Next teamId

    'ChartObj.Activate
'    With ChartObj.Chart
'        .SetElement (msoElementLegendBottom)
'        .Axes(xlValue).MajorUnit = 1
'        .Axes(xlValue).MinorUnit = 0.5
'        .Axes(xlValue).MinorTickMark = xlOutside
'        '.Axes(xlCategory).TickLabels.NumberFormat = "#,##000"
'        .Axes(xlCategory).TickLabels.NumberFormat = "#,##0"
'        '.Location Where:=xlLocationAsObject, Name:="Plot"
'    End With

End Sub
