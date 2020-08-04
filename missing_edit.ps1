##############################################################################

$input = '.\flagged.csv'													## File to run script against, 1 or 2 columns seperated by semicolon
$output = '.\foundusers.csv'												## File to write findings to

$mailSender = 'Servicedesk <john.doe@yourcompany.com>'						## Email address to send emails from
$mailServer = 'smtp.yourcompany.local'										## SMTP Server
$mailAdmin = 'john.doe@yourcompany.com'										## Address used for email copies, samples etc

$mailExceptions = $true                                                     ## Mails list of users that cannot be contacted to $mailAdmin
$mailCopy = $false															## Sends copy of every email to $mailAdmin
$daysToRemove = 3															## Days to remove findings from $input and $output (dependant on $input refresh)

$testing = $false															## Disables writing to files and sending emails if $true
$sampleEmails = 0															## Forwards emails to $mailAdmin, $testing has to be $true

##############################################################################


$countUsers = 1
$today = Get-Date
$getMatch = Get-Content -Path $output
$importInput = Import-Csv -Path $input -Delimiter ';' -Encoding UTF8 -Header columnOne, columnTwo
Import-Module ActiveDirectory
## Find logged on users and import them to $output
foreach ($i in $importInput) {
    ## Prepping our file for Get-ADComputer filter to work properly
    ## An attempt to take care off any discrepancies in flagged.csv
    $first, $second =  $i.columnOne, $i.columnTwo
    if ([string]::IsNullOrEmpty($first)) { $first = $false }
    if ([string]::IsNullOrEmpty($second)) { $second = $false }
    if ($first) { $first = $first.Trim() }
    if ($second) { $second = $second.Trim() }
    $select = Get-ADComputer -Filter "Name -eq '$first' -or Name -eq '$second'" | Select-Object -First 1 -ExpandProperty Name
    Try {
        #Write-Host($select)
        if (Test-Connection -ComputerName $select -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        #$currentUser = (Get-WMIObject -Class Win32_ComputerSystem -ComputerName $select).UserName -replace ".*\\"
        $currentUser = (Get-CimInstance -ComputerName $select -ClassName Win32_ComputerSystem).UserName -replace ".*\\"
        $grabUser = (Get-ADUser -Identity $currentUser -Properties DisplayName, Name, Mail)
        $collectData = "$select;$($grabUser.Name);$($grabUser.DisplayName);$($grabUser.Mail);$today"
        }
        else { Write-Host 'Connection failed:' $select -ForegroundColor Red }
    }
    Catch { Write-Host 'Connection failed:' $select -ForegroundColor Red }
	## Check so we have everything we need
    if ( (![string]::IsNullOrEmpty($currentUser)) -and (![string]::IsNullOrEmpty($select)) -and ($null -ne $collectData) ) {
        $findMatch = $getMatch | Where-Object{$_.Contains($currentUser.Name) -and $_.Contains($select)}
		## check if unique, if yes, write to $output
        if ( ([string]::IsNullOrEmpty($findMatch)) ) {
			Out-File -FilePath $output -InputObject $collectData -Encoding UTF8 -Append
			Write-Host("$($countUsers.ToString()):$collectData") -ForegroundColor Green
			$countUsers++
			}
        else {
			Write-Host("$select already exists") -ForegroundColor Yellow
			}
		}
    $findMatch = $null
    $collectData = $null
}


$countMails = 1
$exceptionsHandle = @()
$exceptionsCount = 1
$importOutput = Import-Csv -Path $output -Delimiter ';' -Encoding UTF8 -Header 'Dator', 'Användare', 'Namn', 'Mail', 'Datum'
## Send emails to logged on users
foreach ($u in $importOutput) {
    $userDevice = $u.Dator
    $userAccount = $u.Användare
    $userName = $u.Namn
    $userMail = $u.Mail
    $userDate = [datetime]$u.Datum
	## Check if user has an email etc
    if ( (![string]::IsNullOrEmpty($userMail)) -and ($userMail.Contains(".")) -and ($userMail.Contains("@")) ) {
        if ( ($userDate.AddHours(2) -gt $today) ) {
            ## Finnish to finnish users
            if ( ($userMail.EndsWith('fi')) ) {
                $mailSubject = 'Tietokoneiden inventaario tarkastus'
                $mailBody = "Hei $userName,
                <p> .... </p>"
                }
            ## Swedish to everyone else
            else {
                $mailSubject = 'Kontroll av datorägare'
                $mailBody = "Hej $userName
                <p> .... </p>"
                }
				## if $false, send out emails
				if ( (!$testing) ) {
					Try {
                        if ( (!$mailCopy) ) {
                        Send-MailMessage -From $mailSender -To $userMail -Subject $mailSubject -SmtpServer $mailServer -Body $mailBody -BodyAsHtml -Encoding UTF8
                        }
                        else {
                        Send-MailMessage -From $mailSender -To $userMail -Subject $mailSubject -SmtpServer $mailServer -Body $mailBody -Bcc $mailAdmin -BodyAsHtml -Encoding UTF8
                        }
					Write-Host("$($countMails.ToString()): Sending mail to $userMail, $userName ($userAccount) regarding device $userDevice") -ForegroundColor Green
					}
					Catch {
					Write-Host("Failed to send email to $userMail")
					}
				}
				## if $true, don't send out emails
				else {
					## if $sampleEmails is 0 don't send out any emails
					if ( ($sampleEmails -eq 0) ) {
					Write-Host("$($countMails.ToString()): Skipping mail to $userMail, $userName ($userAccount) regarding device $userDevice") -ForegroundColor Yellow
					}
					## If $sampleEmails is greater than 0, send out emails to $mailAdmin
					elseif ( ($sampleEmails -gt 0) ) {
						## If $sampleEmails is greater than $countMails, set $sampleEmails to $countMails
						if ( ($sampleEmails) -gt ($countMails-1) ) { $sampleEmails = $countMails-1 }
							Try {
								Send-MailMessage -From $mailSender -To $mailAdmin -Subject $mailSubject -SmtpServer $mailServer -Body $mailBody -BodyAsHtml -Encoding UTF8
							    Write-Host("$($countMails.ToString()): Sending mail to $userMail, $userName ($userAccount) regarding device $userDevice (diverted to $mailAdmin)") -ForegroundColor Yellow
							}
							Catch {
							    Write-Host("Failed to send email to $userMail")
							}
						}
					else {
					Write-Host("Unknown value set. ( $sampleEmails )")
					}
				}
				$countMails++
			}
		}
    ## If user doesn't have email or etc, send data to array $exceptionsHandle if due removal
    else {
        if ( $today.AddDays(-$daysToRemove) -ge $userDate ) {
        $exceptionsHandle += ($exceptionsCount.ToString() + ";$userDevice;$userName;$userMail;$userAccount<br>")
        $exceptionsCount++
        }
    }
}
Write-Host("A total of $(($countMails-1).ToString()) emails sent.") -ForeGroundColor Green

## Mail $exceptionsHandle to $mailAdmin
if ( (![string]::IsNullOrEmpty($exceptionsHandle)) -and (!$testing) -and ($mailExceptions) ) {
	$mailSubject = "Kontroll av datorägare"
    $mailBody = "Dator;Användare;Namn;Mail;Konto<br>$exceptionsHandle"
	Send-MailMessage -From $mailSender -To $mailAdmin -Subject $mailSubject -SmtpServer $mailServer -Body $mailBody -BodyAsHtml -Encoding UTF8
	Write-Host("A total of $(($exceptionsCount-1).ToString()) exceptions sent. to $mailAdmin") -ForeGroundColor Green
}




## Remove -$daystoRemove day(s) old devices from $input and $output
if ( (!$testing) ) {
    $filterMatches = get-content -path $input
    $upd = foreach ($v in $importOutput) {
        $varDate = [datetime]$v.Datum
        ## Compile all that isn't -$daysToRemove days old
        if ($today.AddDays(-$daysToRemove) -le $varDate) {
			$v = "$($v.Dator);$($v.Användare);$($v.Namn);$($v.Mail);$($v.Datum)"
			$v
        }
        ## Remove all that is -$daysToRemove days old from $input
        else {
			Write-Host("Removing  $($v.Dator) from our Searchbase") -ForegroundColor Yellow
			$filterMatches = $filterMatches | Where-Object{!$_.Contains($v.Dator)}
        }
    }
    ## Update $input and $output
    $upd | Out-File -Filepath $output -Encoding utf8
    Out-File -FilePath $input -InputObject $filterMatches
}