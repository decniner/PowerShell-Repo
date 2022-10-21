#Reference necessary assemblies to get the chart and form type into the scope
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
#Instantiate the chart types
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart_area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$series = New-Object  System.Windows.Forms.DataVisualization.Charting.Series
$legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
#Compose the instantiated objects and set the properties
$chart.Series.Add($series)
$chart.ChartAreas.Add($chart_area)
$chart.Legends.Add($legend)
$chart.Dock = [System.Windows.Forms.DockStyle]::Fill

$series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
$series.Points.DataBindXY(@('a', 'b', 'a'), @(5, 7, 2))
#Set up the window and add the chart to the window
$window = New-Object System.Windows.Forms.Form 
$window.Text = "Beautiful Pie Charts"
$window.Controls.Add($chart);
#Show the window
$window.ShowDialog();
