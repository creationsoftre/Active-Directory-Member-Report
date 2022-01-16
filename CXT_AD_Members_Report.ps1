#to get script to work. Complete a find and replace add domain & Add AD SamAccounts

$cxt_SamAccounts = "Add AD SamAccounts"

#create excel application
$excel = New-Object -ComObject excel.application

#Make Excel Visable 
$excel.application.Visible = $true
$excel.DisplayAlerts = $false

#Create WorkBook
$workBook = $excel.Workbooks.Add()

$row = 1
$column = 1

foreach($cxt_samAccount in $cxt_SamAccounts)
{
    #Update Default Sheet Name
    $workSheets = $workBook.worksheets.add()
    $workSheets.name = $cxt_samAccount


    $workSheets.Activate() | Out-Null
    
    #Headers
    $workSheets.Cells.Item($row,$column) = "Name"
    $column++
    $workSheets.Cells.Item($row,$column) = "Account ID"
    $column++
    $workSheets.Cells.Item($row,$column) = "Creation of Account"
    $column++
    $workSheets.Cells.Item($row,$column) = "Last Log-On Date"
    $column++
    $workSheets.Cells.Item($row,$column) = "Enabled"
    
    #Resets column back to 1 for each worksheet
    $column = 1
    if($cxt_samAccount -like '*DEV*' -or $cxt_samAccount -like '*QA*'){
        $info = Get-ADGroupMember -Identity $cxt_samAccount -Server 'add domain' | foreach{ get-aduser $_ -Properties *} | Select DisplayName, SamAccountName, whenCreated, LastLogonDate, Enabled, MemberOf

        $row++
        for($i = 0 ; $i -lt $info.Length; $i++){
            $workSheets.Cells.Item($row,$column) = $info.DisplayName[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.SamAccountName[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.whenCreated[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.LastLogonDate[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.Enabled[$i]
            $column++
            #Resets column back to 1 for each worksheet
            $column = 1
            $row++
        }
    }

    if($cxt_samAccount -like '*PROD*'){
        $info = Get-ADGroupMember -Identity $cxt_samAccount -Server 'add domain'| foreach{ get-aduser $_ -Properties *} | Select DisplayName, SamAccountName, whenCreated, LastLogonDate, Enabled, MemberOf

        $row++
        for($i = 0 ; $i -lt $info.Length; $i++){
            $workSheets.Cells.Item($row,$column) = $info.DisplayName[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.SamAccountName[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.whenCreated[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.LastLogonDate[$i]
            $column++
            $workSheets.Cells.Item($row,$column) = $info.Enabled[$i]
            $column++
            #Resets column back to 1 for each worksheet
            $column = 1
            $row++
        }
    }

    $row = 1
}

#Delete Default Sheet
$workbook.worksheets.item("Sheet1").Delete()