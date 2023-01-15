<#
.SYNOPSIS
cls

Trying to Achieve:     Utility Tool for opening files on regular basis - using Powershell script 


#### Utility tool for opening files ####

Goals:  On daily basis if you are opening files, 
        Instead of going to different locations each time, 
        Have the utility program open files for u

Author: Morphine

-Contains Test-Path in IF ELSE for file validation if file doesn't exist.
-Opens the required file.

** *** ******
Need to remove hard coded paths and add something for it.
---------------######################--------------
Now implementation is using Teams here to get files and stuff. - For some other day.


#>

Write-Output "List of folders/files in my system available to access.5"
Write-Output "1.Excel 
2.Word
3.PPT
4.Outlook"
$excel = 'D:\Test_1\Excel\Test.csv'

$user = Read-Host "Which Folder/Folder No. are u looking for?"

$testExcel = Test-Path -Path $excel  #validating the path.

if ($user -eq "Excel" -or $user -eq 1) {
    if($testExcel -eq $true)
    {
        # D:\Test_1\Excel\Test.csv
        $exc = Get-Content -Path "D:\Practice\PowerShell\Utility Tool\Files-Path.txt"
        Start-Process $exc
    }
}
elseif ($user -eq "Word" -or $user -eq 2) {
    $word
}
elseif ($user -eq "PPT" -or $user -eq 3) {
    $PPT
}
elseif ($user -eq "Outlook" -or $user -eq 4) {
    $Outlook
}
# elseif($user -eq 1 -or $user -eq 2 -or $user -eq 3 -or $user -eq 4){

# }
else {
    echo "Go Home"
}