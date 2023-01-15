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

More Changes - Instead of going through each inside loop (line no 37), 
will put a function so its more convenient - still same as below string.
-function was a tad tedious, now the user can put the path of the location in external file (for now txt , csv can be used).
2) also thinking of splitting the $excel variable w/o ' '- would look coool may be
-splitting wont work, but replace will work., no need will still return a string.
- no need , Start Process helps.
** *** ******
all was rubbish, used a simple start cmdlet to initiate the process
start $excel , but since I moved the file to some external file to make it easy for the user. here we are :)/\
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