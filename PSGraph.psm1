#version 2.1
#jst
#changelog v1.0 - Initial release
#changelog v2.0 - optimization
#changelog v2.1
#- consolidation of module using member methods
#- fixed scope issues with this module

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
function New-PSGraphChart {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$false)][string]$Width=1200,
    [Parameter(Mandatory=$false)][string]$Height=600,
    [Parameter(Mandatory=$True)][string]$AxisXTitle,
    [Parameter(Mandatory=$True)][string]$AxisYTitle,
    [Parameter(Mandatory=$false)][string]$AxisXInterval = 5,
    [Parameter(Mandatory=$True)][string]$ChartTitle

  )

  $chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
  $chart.Width = $Width
  $chart.Height = $Height
  $chart.BackColor = [System.Drawing.Color]::White


  # chart area
  $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
  $chartarea.Name = "ChartArea1"
  $chartarea.AxisY.Title = $AxisYTitle
  $chartarea.AxisX.Title = $AxisXTitle
  #$chartarea.AxisY.Interval = 100
  $chartarea.AxisX.Interval = $AxisXInterval
  $chart.ChartAreas.Add($chartarea)

  # legend
  $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
  $legend.name = "Legend1"

  $chart.Legends.Add($legend)
  [void]$chart.Titles.Add($ChartTitle)
  $chart.Titles[0].Font = "Arial,20pt"
  $chart.Titles[0].Alignment = "topLeft"

  #add a scriptmethod to be able to save the chart to disk
  $chart | Add-Member -MemberType ScriptMethod -Name "SaveChart" -Value {
  Param(
    [Parameter(Mandatory=$True)][string]$FileName,
    [Parameter(Mandatory=$false)][string]$fileformat="png",
    [Parameter(Mandatory=$false)][string]$OpenImages
  )


  $this.SaveImage($FileName,$fileformat) |out-null
  if($OpenImages -eq $true){Invoke-Item $FileName }

  }

  #add a scriptmethod to be able to add a graph serie
  $chart | Add-Member -MemberType Scriptmethod -Name "AddSerie" -Value {
    Param(
      [Parameter(Mandatory=$True)][string]$SerieName,
      [Parameter(Mandatory=$false)][string]$Type="Line"
    )
      #$chart = Get-Variable -Name "$ChartName" -ValueOnly
      $this.Series.Add($SerieName)|out-null
      $this.Series[$SerieName].ChartType = $Type
      $this.Series[$SerieName].IsVisibleInLegend = $true
      $this.Series[$SerieName].BorderWidth  = 3
      $this.Series[$SerieName].chartarea = "ChartArea1"

  }

  $chart | Add-Member -MemberType Scriptmethod -Name "AddSerieValues" -Value {
    Param(
      [Parameter(Mandatory=$True)][string]$SerieName,
      [Parameter(Mandatory=$True)][string]$XValue,
      [Parameter(Mandatory=$True)][string]$YValue
    )
    $this.Series[$SerieName].Points.addxy( $XValue , $YValue)  |Out-Null
  }

  $chart
}
Export-ModuleMember -Function New-PSGraphChart
