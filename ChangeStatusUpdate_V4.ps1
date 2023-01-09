<# Script made by 
 .----------------.  .----------------.  .----------------.  .----------------.  .----------------.  .-----------------.
| .--------------. || .--------------. || .--------------. || .--------------. || .--------------. || .--------------. |
| |  ________    | || |      __      | || |   _____      | || |  _________   | || |     ____     | || | ____  _____  | |
| | |_   ___ `.  | || |     /  \     | || |  |_   _|     | || | |  _   _  |  | || |   .'    `.   | || ||_   \|_   _| | |
| |   | |   `. \ | || |    / /\ \    | || |    | |       | || | |_/ | | \_|  | || |  /  .--.  \  | || |  |   \ | |   | |
| |   | |    | | | || |   / ____ \   | || |    | |   _   | || |     | |      | || |  | |    | |  | || |  | |\ \| |   | |
| |  _| |___.' / | || | _/ /    \ \_ | || |   _| |__/ |  | || |    _| |_     | || |  \  `--'  /  | || | _| |_\   |_  | |
| | |________.'  | || ||____|  |____|| || |  |________|  | || |   |_____|    | || |   `.____.'   | || ||_____|\____| | |
| |              | || |              | || |              | || |              | || |              | || |              | |
| '--------------' || '--------------' || '--------------' || '--------------' || '--------------' || '--------------' |
 '----------------'  '----------------'  '----------------'  '----------------'  '----------------'  '----------------' 
 let me know if you have any questions
 #>

#Username variable
$username = $env:username.substring(0, $env:username.length - 1)

#Import-Excel Module Install - Keyloni Big Moni
$psGalleryTrust = Get-PSRepository -name PSGallery | Select-Object -ExpandProperty Trusted
if (-not($psGalleryTrust)) {
    Set-PSRepository -name PSGallery -InstallationPolicy Trusted
}
import-module *
$moduleCheck = Get-Module ImportExcel
if ($null -eq $moduleCheck) {
    Install-Module -Name ImportExcel -Scope CurrentUser | Out-Null
    import-module ImportExcel
}

#folder permission check
$pathcheck = test-path C:\users\$username\Documents -errorvariable failed
if ($pathcheck -eq $true) {

    #Excel Data
    $filepath = "C:\Users\$username\OneDrive - Humana\Command\Nightly Report Final.xlsx"
    $ExcelPkg = Open-ExcelPackage -path $filepath
    $WorkSheet = $ExcelPkg.Workbook.Worksheets["Nightly"].Cells

    #Change Details
    Write-Host "Update the Change with the reason it is unable to be completed and what the requestor's next steps are" -ForegroundColor Red
    $Change = Read-Host "Please enter the Change Number`n"
    $Summary = Read-Host "Please enter the short description of the change`n"
    $Reason = Read-Host "Please enter the reason why the Change was unable to be completed`n"
    $Status = Read-Host "Please enter the status of the change from the following selection. Canceled, Hold, or Failed`n"
    $EmailTo = Read-Host "Please enter the requestor's email`n"

    #Target Cell start
    $x = 1
    $y = 35

    #finding an available row entry
    $cellvalue = "test"
    while ($null -ne $cellvalue) {
        $cellvalue = $WorkSheet[$y, $x] | Select-Object value
        if ($null -ne $cellvalue) {
            $y++
        }
        elseif ($y -eq '45') {
            write-host "Table is full, please manually add entry or clear space"
            exit
        }
        else {
            Write-Host "Available Row Located" -ForegroundColor Green
        }

    }
    #updating the row
    $worksheet[$y, $x].value = $Change.ToUpper()
    $WorkSheet[$y, ($x + 1)].value = $Status.ToUpper()
    $WorkSheet[$y, ($x + 2)].value = $Reason
    $WorkSheet[$y, ($x + 3)].value = $Summary


    Close-ExcelPackage $ExcelPkg
    #required to push the update up in onedrive
    Move-Item  $filepath $filepath

    ##Email Notification##
    #Configuration Variables for E-mail
    $SmtpServer = "pobox.humana.com" #or IP Address such as "10.125.150.250"
    $EmailFrom = "NOCEscalation@humana.com"
    $cc = "Ddecker2@humana.com, dneely1@humana.com, cgriffin28@humana.com, jervin1@humana.com, mclare@humana.com, rkorosec@humana.com, KDeLancey@humana.com, Tmurrell1@humana.com, vbandavong1@humana.com, cburns22@humana.com, NOC@humana.com"
    $EmailSubject = "$Change has been moved to $Status | $Summary"
 
    #HTML Template
    $EmailBody = @"
<p> Hello, <p>
<table style="width: 68%" style="border-collapse: collapse; border: 1px solid #008080;">
 <tr>
    <td colspan="2" bgcolor="#008080" style="color: #FFFFFF; font-size: large; height: 35px;"> 
        Change Status Report  
    </td>
 </tr>
 <tr style="border-bottom-style: solid; border-bottom-width: 1px; padding-bottom: 1px">
    <td style="width: 201px; height: 35px">  Change Number:</td>
    <td style="text-align: center; height: 35px; width: 233px;">
    <b>VarChange</b></td>
 </tr>
  <tr style="height: 39px; border: 1px solid #008080">
  <td style="width: 201px; height: 39px">  Change Status:</td>
 <td style="text-align: center; height: 39px; width: 233px;">
  <b>VarStatus</b></td>
 </tr>
</table>
<p> The TOC has determined that your Change could not be continued for the following reason, <b>VarReason</b><p> 
<P>Please review the change in SerivceNOW for further details <p>
<p>Thank you, <p>
<p>Humana TOC<p>
"@
 

    #Replace the Variables Change, Status, and Reason
    $EmailBody = $EmailBody.Replace("VarChange", $Change)
    $EmailBody = $EmailBody.Replace("VarStatus", $Status)
    $EmailBody = $EmailBody.Replace("VarReason", $Reason)

  
    #Send E-mail from PowerShell script
    Send-MailMessage -To $EmailTo -From $EmailFrom -CC $cc -Subject $EmailSubject -Body $EmailBody -BodyAsHtml -SmtpServer $SmtpServer -WarningAction Ignore
    Write-Host "Notification Email sent and Nightly Report has been updated" -ForegroundColor Green
}
else {
    Write-Host "Permssion to the User folder has been denied. Please refer to the README.txt for this script for further instruction"
    exit
}