Sub SetChartSeriesToStacked()

    Dim CHART_NAME As Variant   
    Dim seriesNames As Variant
    Dim seriesName As Variant
    Dim chartObj As ChartObject
    Dim ser As Series
    Dim serExists As Boolean
    Dim chartExists As Boolean

    ' 👇 UPDATE CHART NAME HERE 👇:
    CHART_NAME = "Chart 7"
    
    ' Check if chart exists
    chartExists = False
    On Error Resume Next ' Ignore errors temporarily
    Set chartObj = ActiveSheet.ChartObjects(CHART_NAME)
    If Not chartObj Is Nothing Then chartExists = True
    On Error GoTo 0 ' Turn back on normal error handling
       

    If chartExists Then
        ' Count total series in the chart
        totalSeries = chartObj.Chart.SeriesCollection.Count
        
        ' Loop through each series in the chart
        For seriesCounter = 1 To totalSeries
            Set chartSeries = chartObj.Chart.SeriesCollection(seriesCounter)
            
            ' Set series to Column Stacked, except the last three
            If seriesCounter <= totalSeries - 1 Then
                chartSeries.ChartType = xlColumnStacked
            Else
                chartSeries.ChartType = xlLine
            End If
        Next seriesCounter
    Else
        MsgBox "Chart 7 does not exist on the active sheet.", vbExclamation, "Chart Not Found"
    End If
End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    ' Call your macro when the slicer changes
    Call SetChartSeriesToStacked
End Sub