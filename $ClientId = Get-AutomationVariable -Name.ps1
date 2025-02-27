$ClientId = Get-AutomationVariable -Name 'Azure Automation ClientID'
$TenantId = Get-AutomationVariable -Name 'Azure Automation TenantID'
$Cert = Get-AutomationCertificate -Name 'Azure Automation'
$Tenant = 'forthepeople0.onmicrosoft.com'
$recipientEmail = @("jack.stokes@forthepeople.com","jack.stokes@forthepeople.com")
$fromEmail = "EmployeeSurveys@forthepeople.com"

#Import-Module Microsoft.Graph.Users.Actions
$TriggerDay = 1
$SharePointURL  = "https://forthepeople0.sharepoint.com"
$SharePointSite = "$SharePointURL/teams/HREmployeeSurveys-dev"
$Force = $true

function Send-TestEmail {
    # Get current date information
    $today = Get-Date
    $currentDay = $today.Day
    $targetMonth = $today

    # If we're in the first 7 days of the month, use previous month's folder
    if ($currentDay -le 7) {
        $targetMonth = $today.AddMonths(-1)
    }

    # Check if we should proceed based on trigger day
    if (-not $Force -and $currentDay -ne $TriggerDay) {
        Write-Verbose "Not trigger day. Use -Force to override."
        return @{
            Success = $false
            Reason = "Not trigger day"
            EmailsSent = 0
        }
    }

        # Format folder name properly (e.g., "January 2025")
        $folderName = $targetMonth.ToString("MMMM yyyy")
        $folderPath = "Shared Documents/Survey Results/Exit Surveys - dev/$folderName"

        # Get folder URL
        #$web = Get-PnPWeb
        $folderUrl = "$SharePointSite/$folderPath"
        Write-Output "Sending email to $($recipientEmail.Count) recipients..."
        $emailSubject = "$($targetMonth.ToString("MMMM")) Exit Surveys"
        # Create HTML email body
        $htmlBody =@"
<p>Hello,</p>
<p>Here's the link to the Exit Survey folder for $($folderName):</p>
<p><a href='$folderUrl'>$folderName Exit Survey Folder</a></p>
<p>Please ensure all exit surveys are completed and uploaded to this location.</p>
<p>Best regards,<br>IT Team</p>
"@
     
    try {
        # Connect to Microsoft Graph
        Connect-MgGraph -ClientId $ClientId -TenantId $tenantId -CertificateThumbprint $Cert.Thumbprint -NoWelcome
        # Create email message
        $params = @{
            Message = @{
                Subject = $emailSubject
                Body = @{
                    ContentType = "HTML"
                    Content = $htmlBody
                }
                ToRecipients = @(
                    foreach($recipient in $recipientEmail) {
                        @{
                            EmailAddress = @{
                                Address = $recipient
                            }
                        }
                    }
                )
                From = @{
                    EmailAddress = @{
                        Address = $fromEmail
                    }
                }
            }
            SaveToSentItems = $true
        }

        # Send email using Microsoft Graph
        Send-MgUserMail -UserId $fromEmail -BodyParameter $params
        
        # Log success
        Write-Output "Email sent successfully with subject: Test Email"
        return $true
    }
    catch {
        # Log any errors
        Write-output "Failed to send email: $_ Error: $error "
        return $false
    }
    finally {
        # Disconnect from Microsoft Graph
        Disconnect-MgGraph
    }
}

 Send-TestEmail