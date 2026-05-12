##############################################################################
#  Script: password_manager.ps1
#    Date: 05.11.2026
# Version: 1.0
#  Author: Adrian Munoz
# Purpose: Password Manager utilizing PowerShell and Excel.Application COM Object
#   Legal: Script provided "AS IS" without warranties or guarantees of any
#          kind.  USE AT YOUR OWN RISK.  Public domain, no rights reserved.
##############################################################################

$pslist = "$env:USERPROFILE\Documents\pslist.xlsx"

if ($(Test-Path $pslist) -eq $false)
{
    Write-Host "Password list not found!" -ForegroundColor Red
    Start-Sleep 1
    Write-Host "Creating Password list at $env:USERPROFILE\Documents\pslist.xlsx..." -ForegroundColor Cyan
    $check = $true
    While ($check -eq $true)
    {
        Start-Sleep 1
        Write-Host "Enter Password to lock Password List:" -ForegroundColor Yellow
        $pass1 = Read-Host -AsSecureString
        Write-Host "Re-enter Password:" -ForegroundColor Yellow
        $pass2 = Read-Host -AsSecureString
        $pass1_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass1))
        $pass2_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass2))

        if ($pass2_txt -ne $pass1_txt)
        {
            Write-Host "[!] ERROR: The passwords you entered do not match!" -ForegroundColor Red
            Start-Sleep 1
            $check = $true
        }
        elseif ($pass2_txt -eq $pass1_txt)
        {
            Write-Host "Continuing..." -ForegroundColor Cyan
            $check = $false
        }
    }

    $fecs = @("Site1-ELK", "Site2-ELK", "Site1-SPL", "Site2-SPL")

    $excel = New-Object -ComObject "Excel.Application"
    $excel.Visible = $false
    $wb = $excel.workbooks.add()
    $ws = $wb.worksheets.item(1)
    $ws.name = "PasswordList"
    $ws.Cells.Item(1,1).value = "Sites"
    $ws.Cells.Item(1,2).value = "Passwords"
    $row = 2
    ForEach ($fec in $fecs)
    {
        $ws.Cells.Item($row,1).value = $fec
        $row++
    }
    $excel.ActiveWorkbook.SaveAs($pslist, 51, $pass1_txt)
    Start-Sleep 1
    $excel.quit()
    Get-Process -Name *excel* | Stop-Process
    Write-Host "[+] Password list created at $env:USERPROFILE\Documents\pslist.xlsx" -ForeGroundColor Green
}

if ($(Test-Path $pslist) -eq $true)
{
    $excel = New-Object -ComObject "Excel.Application"
    $excel.Visible = $false
    $workbooks = $excel.Workbooks

    $attempts = 0
    while ($attempts -lt 3)
    {
        Write-Host "Enter Password to unlock Password List:" -ForegroundColor Yellow
        $pass = Read-Host -AsSecureString
        $pass_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))
        try 
        {
            $workbook = $workbooks.Open($pslist, $false, $false, [type]::Missing, $pass_txt)
            $worksheet = $workbook.Worksheets.item(1)
            $attempts = 5
        }
        catch
        {
            Write-Host "[!] Incorrect Password!" -ForegroundColor Red
            $attempts++
        }
    }

    if (($attempts -ge 3) -and ($attempts -lt 5))
    {
        Write-Host "[!] Too many incorrect attempts... If you forgot your password, delete `"$pslist`" and start over." -ForegroundColor Magenta 
        Start-Sleep 1
        $excel.quit()
        Get-Process -Name *excel* | Stop-Process
        exit
    }
    Write-Host "Successfully authenticated!" -ForegroundColor Green
    {}
    Write-Host "Select-Site:" -ForegroundColor Yellow
    Write-Host "[1] - Site1-ELK" -ForegroundColor Green
    Write-Host "[2] - Site2-ELK" -ForegroundColor Green
    Write-Host "[3] - Site1-SPL" -ForegroundColor Green
    Write-Host "[4] - Site2-SPL" -ForegroundColor Green
    Write-Host "[0] - Exit" -ForegroundColor Red
    
    $response = Read-Host

    Switch($response)
    {
        {$_ -eq 0} {Write-Host "Cya!" -ForegroundColor Magenta; $excel.Quit(); exit}
        {$_ -eq 1} {$site = "Site1-Elk"}
        {$_ -eq 2} {$site = "Site2-Elk"}
        {$_ -eq 3} {$site = "Site1-SPL"}
        {$_ -eq 4} {$site = "Site2-SPL"}
    }

    $sitecell = $worksheet.Cells.Find("$site")
    $row = $sitecell.row
    $column = $sitecell.column
    $password = $worksheet.cells.item($row, $($column+1)).text
    
    if ($password -eq "")
    {
        Write-Host "$site does not contain a password. Would you like to add one? [y/n]" -ForegroundColor Yellow
        $reply = Read-Host

        if (($reply -eq "y") -or ($reply -eq "Y"))
        {
            
            $attempts = 0

            While ($attempts -lt 3)
            {
                Write-Host "Enter Password:" -ForegroundColor Green
                $siteps = Read-Host -AsSecureString
                Write-Host "Re-enter Password:" -ForegroundColor Green
                $siteps2 = Read-Host -AsSecureString
                $siteps_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($siteps))
                $siteps2_txt = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($siteps2))

                if ($siteps_txt -ne $siteps2_txt)
                {
                    Write-Host "[!] ERROR: The passwords you entered do not match!" -ForegroundColor Red
                    $attempts++
                    
                }

                elseif ($siteps_txt -eq $siteps2_txt)
                {
                    Write-Host "Adding password for $site..." -ForegroundColor Cyan
                    $worksheet.Cells.Item($row, $($column+1)).value = $siteps2_txt
                    Start-Sleep 1
                    Write-Host "[+] Password successfully added to $site!" -ForegroundColor Green
                    Start-Sleep 1
                    Write-Host "[+] $site's password set to clipboard" -ForegroundColor Green
                    Set-Clipboard $siteps2_txt
                    $excel.ActiveWorkbook.Save()
                    $attempts = 5
                    $excel.quit()
                    Get-Process -Name *excel* | Stop-Process
                    
                
                }
            }
            
            if (($attempts -ge 3) -and ($attempts -lt 5))
            {
                Write-Host "[!] Too many failed attempts... Try again later..." -ForegroundColor Magenta
                Start-Sleep 1
                $excel.quit()
                Get-Process -Name *excel* | Stop-Process
                exit
            }
        }
        elseif (($reply -eq "n") -or ($reply -eq "N"))
        {
            Write-Host "Cya!" -ForegroundColor Magenta 
            $excel.Quit()
            Get-Process -Name *excel* | Stop-Process
            exit
        }
    }
    
    elseif ($password -ne "")
    {
        Start-Sleep 1
        Write-Host "[+] $site's password set to clipboard" -ForegroundColor Green
        Set-Clipboard $password
        $excel.Quit()
        Get-Process -Name *excel* | Stop-Process
    }
        
}
