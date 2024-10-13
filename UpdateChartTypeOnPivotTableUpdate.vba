Sub SetChartSeriesToStacked()

    Dim CHART_NAME As Variant  ' Chart that is affected by the change
    Dim LINE_SERIES_NAME As Variant ' Series name that will be changed to xlLine
    Dim seriesNames As Variant
    Dim seriesName As Variant
    Dim chartObj As ChartObject
    Dim ser As Series
    Dim serExists As Boolean
    Dim chartExists As Boolean
    

    ' 👇 UPDATE CHART NAME HERE 👇:
    CHART_NAME = "Chart 7"
    LINE_SERIES_NAME = "my_series"
    
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
            If ser.Name = LINE_SERIES_NAME Then
                ser.ChartType = xlLine
            Else
                chartSeries.ChartType = xlColumnStacked
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