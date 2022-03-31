$cxt_SamAccounts = "TP_GBD_Cxt_Dev_UserDev", "TP_GBD_Cxt_Dev_BusAdminDev", "TP_GBD_Cxt_Dev_SysAdminDev", "TP_GBD_Cxt_QA_User", "TP_GBD_Cxt_QA_SysAdmin", "TP_GBD_Cxt_QA_BusAdmin", "TP_GBD_Cxt_Prod_User", "TP_GBD_Cxt_Prod_SysAdmin", "TP_GBD_Cxt_Prod_BusAdmin"

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
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Account ID"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Creation of Account"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Last Log-On Date"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    $column++
    $workSheets.Cells.Item($row,$column) = "Enabled"
    $workSheets.Cells.Item($row,$column).Font.Bold = $true
    $workSheets.Cells.Item($row,$column).Font.Color = 8210719
    $workSheets.Cells.Item($row,$column).Font.Size = 15
    
    #Resets column back to 1 for each worksheet
    $column = 1
    if($cxt_samAccount -like '*DEV*' -or $cxt_samAccount -like '*QA*'){
        $info = Get-ADGroupMember -Identity $cxt_samAccount -Server 'devad.wellpoint.com' | foreach{ get-aduser $_ -Properties *} | Select DisplayName, SamAccountName, whenCreated, LastLogonDate, Enabled, MemberOf

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
        $info = Get-ADGroupMember -Identity $cxt_samAccount -Server 'us.ad.wellpoint.com'| foreach{ get-aduser $_ -Properties *} | Select DisplayName, SamAccountName, whenCreated, LastLogonDate, Enabled, MemberOf

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

    #Auto fit everything so it looks better
    $usedRange = $workSheets.UsedRange
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

#Delete Default Sheet
$workbook.worksheets.item("Sheet1").Delete()

#Save the file
$workbook.SaveAs("\\va01pstodfs003.corp.agp.ads\apps\Local\EMT\COTS\McKesson\ClaimsXten\Active Directory\CXT_AD_Users_List.xlsx")

#close workbook
#$workbook.Close

#Quit the application
$excel.Quit()

#Release COM Object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null