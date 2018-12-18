Clear
#Data Request

$NameOfExcel = Read-Host 'Name of the Excel?'
Clear
$ChangeNumber = Read-Host 'Enter Change Number'
Clear
$StgOrPrd = Read-Host 'Is for Staging or Production? Press 1 for Staging, 2 for production'
Clear
$NameOfMp = Read-Host 'Enter the name of the MP'
Clear
$VersionNumber = Read-Host 'Enter the Version of MP'
Clear
$ImportingMP = "$($NameOfMp) $($VersionNumber)"
$LocationOfMp = Read-Host 'Enter the location of the MP'
Clear

#Monitors
$MonitorValues = @()
$MonitorValue = Read-Host "Enter the Monitors separated by a comma"
$MonitorValueSplit = $MonitorValue.Split(',')
$MonitorValues += $MonitorValueSplit
Clear

#Groups
$GroupValues = @()
$GroupValue = Read-Host "Enter the Groups separated by a comma"
$GroupValueSplit = $GroupValue.Split(',')
$GroupValues += $GroupValueSplit
Clear

#Server Assignation to groups
$ServerValues = @()
foreach($GroupServer in $GroupValues){
    $ServerValue = Read-Host "Enter the Server list for group"$GroupServer
    $ServerValues += $ServerValue
}
Clear
#Comments for Excel
$CommentValue = Read-Host 'Want to add comments? Press 1 to add comments, 2 to skip'
if($CommentValue -eq 1){
    $Comments = Read-Host 'Write comments. Press Enter to finish'
}
Clear

#Creating the Excel
$f = 2
$excel = New-Object -ComObject excel.application
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$Data= $workbook.Worksheets.Item(1)

#Excel Data Values
$Data.Cells.Item(1,1) = 'Change Number'
$Data.Cells.Item(1,2) = 'Staging or Production'
$Data.Cells.Item(1,3) = 'Name of MP'
$Data.Cells.Item(1,4) = 'Version'
$Data.Cells.Item(1,5) = 'Location'
$Data.Cells.Item(1,6) = 'Monitor / Rules'
$Data.Cells.Item(1,7) = 'Importing MP'
$Data.Cells.Item(1,8) = 'Groups on MP'
$Data.Cells.Item(1,9) = 'Servers to be added'
$Data.Cells.Item(1,10) = 'Overrides to be performed'
$Data.Cells.Item(1,11) = 'Comments'
$Data.Cells.Item(2,1) = $ChangeNumber
$Data.Cells.Item(2,2) = $StgOrPrd
$Data.Cells.Item(2,3) = $NameOfMp
$Data.Cells.Item(2,4) = $VersionNumber
$Data.Cells.Item(2,5) = $LocationOfMp
$Data.Cells.Item(2,6) = "Import MP $($NameOfMp) $($VersionNumber)"
$Data.Cells.Item(2,11) = $Comments

if($StgOrPrd -eq 1){
    $Data.Cells.Item(2,2) = 'Staging'
    } else { $Data.Cells.Item(2,2) = 'Production'
    }

foreach($Monitor in $MonitorValues){
    $Data.Cells.Item($f,7) = $Monitor
    $f++
}
$f = 2
foreach($Group in $GroupValues){
    $Data.Cells.Item($f,8) = $Group
    $f++
}
$f = 2
foreach($Server in $ServerValues){
    $Data.Cells.Item($f,9) = $Server
    $f++
}
$f = 2

for($i = 0; $i -lt $GroupValues.Length; $i++){
    $Override = "On $($MonitorValues[$i]) - Enable the monitor for group $($GroupValues[$i])"
    $Data.Cells.Item($f,10) = $Override
    $f++
}

#Format, Save and Quit Excel 

$user_name = [Environment]::UserName
$usedRange = $Data.UsedRange
$usedRange.Application.ActiveWindow.SplitRow = 1
$usedRange.Application.ActiveWindow.FreezePanes = $true
$Selection = $Data.Range('a1:k1')
$Selection.Interior.ColorIndex = 33
$Selection.BorderAround(1,4,30)
$Selection.Font.Bold = $True
$Selection.Font.Size = 12
$usedRange.EntireColumn.AutoFit() | Out-Null
$excel.Quit()

#Stop all Excel Processes

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)