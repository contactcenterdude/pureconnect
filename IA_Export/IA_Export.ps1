cls


# ManagedWinapi.dll only works in 32-bit envirnoment.
if ([Environment]::Is64BitProcess -eq $True)
{
    Write-Host "ERROR: This script can only run in 32-bit environment."
    Write-Host "Please make sure that you run it in Powershell(x86)"
    break
}


#name of the view in Interaction Attendant
$IA_view=$args[0]

#if input parameter is empty - exit
if ($IA_view -eq $null)
{
	Write-Host "Please provide argument"
	Write-Host " "
	Write-Host "Example:"
	Write-Host " "
	Write-Host ".\IA_Export.ps1 ""Managed IP Phones"""
	
	break
	
}


#all views in IA have prefix IAListView
$syslistview32Name="IAListView:"+$IA_view


$scriptPath = if (-not $PSScriptRoot) {  # $PSScriptRoot not defined?
    # Get the path of the executable *as invoked*, via
    # [environment]::GetCommandLineArgs()[0],
    # resolve it to a full path with Convert-Path, then get its directory path
    Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0])) 
  } 
  else {
    # Use the automatic variable.
    $PSScriptRoot 
  }


#load ManagedWinapi library
Add-Type -Path "$scriptPath\ManagedWinapi.dll"

# find Interaction Administrator window
$connectionStatusWindow = [ManagedWinapi.Windows.SystemWindow]::AllToplevelWindows | Where-Object { $_.Title.Contains("Interaction Administrator") }


#Interaction Administrator window is not found

if ($connectionStatusWindow -eq $null)
{
	Write-Host "ERROR: Interaction Attendant window is not found. Please make sure that you have Interaction Attendant open."	
	break	
}

 
#find IA view that we need
$listViewWindow = $connectionStatusWindow.AllDescendantWindows | Where-Object { $_.ClassName.Equals("SysListView32") -and $_.Title.Equals($syslistview32Name)}

if ($listViewWindow -eq $null)
{
	Write-Host "ERROR: Cannot find view ""$IA_view"". Make sure you entered correct name of the view."	
	break	
}

#Extract data from the view
$listView = [ManagedWinapi.Windows.SystemListView]::FromSystemWindow($listViewWindow)

if ($listView -eq $null)
{
	Write-Host "ERROR: View ""$IA_view"" is empty. Make sure it is currently open."	
	break	
}



#convert it to Data Table 
$data = New-Object System.Data.DataTable



ForEach ($col in $listView.Columns)
{
    $data.Columns.Add($col.Title) | Out-Null
}


 
For ($i = 0; $i -lt $listView.Count; $i++)
{
    $newRow = @()
    For ($j = 0; $j -lt $data.Columns.Count; $j++)
    {
        $newRow += $listView[$i,$j].Title
    }
    $data.Rows.Add($newRow)
}



#export to CSV file  -- filename=  {ViewName}__YYYY-MM-DD__HH:MM:ss.csv
$CurDate=(Get-Date).ToString().Replace(' ','__').Replace(":","-").Replace("__PM","PM")
$file_name="$scriptPath\"+$IA_view.Replace(' ','_')+"__"+$CurDate+".csv"
$data | Export-Csv  -NoTypeInformation $file_name

Write-Host "Extraction completed"
Write-Host 'Output:'$file_name
 

Remove-Variable connectionStatusWindow
Remove-Variable listViewWindow
Remove-Variable listView
Remove-Variable data
