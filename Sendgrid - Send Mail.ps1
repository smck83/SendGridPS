########################################################################################
##################### Send E-mail with SendGrid via Powershell #########################
########################################################################################
#################                                                     ##################
################# Use the SendGrid API to send e-mail from Powershell ##################
################# Author: https://github.com/smckellar/SendGridPS     ##################
################# Last Updated: 11-Apr-20                             ##################
################# Sendgrid API Refernece Index:                       ##################
################# https://sendgrid.com/docs/API_Reference/index.html  ##################
#################                                                     ##################
########################################################################################
################################# SET E-MAIL VARIABLES #################################
########################################################################################

# Set your Sendgrid API Key e.g. starting with 'SG.' - you can get this from : https://app.sendgrid.com/settings/api_keys
$apikey = ""

# Set the Recipient TO: Address
$recipientEmail = "billgates@contoso.com"

# Set the Recipient TO: Address
$recipientDisplayName = "William Gates"

# Set the Header FROM: Address
$fromEmail = "stevenjobs@isender.local"

# Set the Header FROM: Display Name
$fromDisplayName = "Steven Paul Jobs"

# Set the Subject Line
$subject = "This is my test message $(get-date)"

#Set the message body here
$messagebody = "
                Dear $recipientDisplayName,<br/>
                Test email sent at $(Get-Date)
                Regards,<br/>
                $fromDisplayName
               "

# To attach a bad file set to 1, otherwise set to 0 for safe file
$includeBadFile = 0

########################################################################################
######################## DON'T CHANGE ANYTHING BELOW THIS LINE #########################
########################################################################################
clear
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
write-host "Running Sendgrid E-mail Script"

$date = $(get-date).ToString('dd-MM-yyyy')

if ($apikey -eq ""){
"No API Key Set, please set `$apikey
"
exit
}
elseif ($recipientEmail -eq ""){
"No Recipient Email Set, please set `$recipientEmail
"
exit
}

if ($includeBadFile -eq 1){
$filename = "this-is-a-bad-file $date.xlsm"
# Base64 hash of Eicar Virus file
$filecontent = "WDVPIVAlQEFQWzRcUFpYNTQoUF4pN0NDKTd9JEVJQ0FSLVNUQU5EQVJELUFOVElWSVJVUy1URVNULUZJTEUhJEgrSCo="
}
else
{
# FileContent = base64 of a file. You can use the function below : base64File <File Path Here>
$filecontent = "VGhpcyBpcyBhIHNhZmUgZmlsZQ=="
$filename = "Safe File $date.txt"
}

$url = "https://api.sendgrid.com/v3/mail/send"


$header = @{
        "Content-Type" = "application/json"
        "Authorization" = "Bearer $apikey"
        }


function makeBody (){
$jsonbody = @{
           personalizations = @(
                                @{
                                    to = @(
                                            @{
                                                email = $recipientEmail
                                                name = $recipientDisplayName
                                            }
                                          )
                                    subject = $subject
                                }
                                )
                        from = @{
                                    email = $fromEmail
                                    name = $fromDisplayName
                                }
                     content = @(
                                    @{
                                         type = "text/html"
                                         value = $messagebody
                            
                                        }
                                )

                           attachments = @(
                                    @{
                                         filename = $filename
                                         content = $filecontent
                            
                                        }
                                ) 
                                
        

        } | ConvertTo-Json -Depth 4

return $jsonbody

}

function sendEmail (){

Invoke-RestMethod -Uri $url -Method Post -Body $(makeBody) -Headers $header

}

#Call the function sendEmail to send the message
sendEmail
