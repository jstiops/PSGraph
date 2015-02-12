#version 1.0
#jst
#changelog v1.0 - Initial release
 
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


  $this.SaveImage($FileName,$fileformat)
  if($OpenImages -eq $true){Invoke-Item $FileName }

  }

  $chart
}

#Takes object collection and adds a serie for each object piped into it
function Add-PSGraphSerie {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$True)][string]$Name,
    [Parameter(Mandatory=$false)][string]$Type="Line",
    [Parameter(Mandatory=$True)][string]$XPropertyName,
    [Parameter(Mandatory=$True)][string]$YPropertyName,
    [Parameter(Mandatory=$True)][string]$ChartName,
    [Parameter(ValueFromPipeline=$True)][string[]]$datasource
    
  )
  BEGIN {
    
    $chart = Get-Variable -Name "$ChartName" -ValueOnly
    $chart.Series.Add($Name)
    $chart.Series[$Name].ChartType = $Type
    $chart.Series[$Name].IsVisibleInLegend = $true
    $chart.Series[$Name].BorderWidth  = 3
    $chart.Series[$Name].chartarea = "ChartArea1"
    
  }
  
  PROCESS{
    
    $chart.Series[$Name].Points.addxy( $_.($XPropertyName) , $_.($YPropertyName))  |Out-Null
  }
  
  END {}
}


function Out-PSGraphImageFile
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$True)][string]$FileName,
    [Parameter(Mandatory=$false)][string]$fileformat="png",
    [Parameter(Mandatory=$True)][string]$ChartName,
    [Parameter(Mandatory=$false)][string]$OpenImages
  )

  $chart = Get-Variable -Name "$ChartName" -ValueOnly
  
  $chart.SaveImage($FileName,$fileformat)
  if($OpenImages -eq $true){Invoke-Item $FileName }
}

Export-ModuleMember -Function New-PSGraphChart, Add-PSGraphSerie
