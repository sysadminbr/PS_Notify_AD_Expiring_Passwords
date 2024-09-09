##################################################################################################################
# CITRA IT CONSULTING
# SCRIPT PARA NOTIFICAR AO USUÁRIO POR EMAIL QUE SUA SENHA ESTÁ PRÓXIMA DE ESPERAÇÃO.
# Author: luciano@citrait.com.br
# Date: 17/10/2021
# version: 1.0
# Agendar a execução deste script em um controlador de domínio.
#   O Script irá localizar usuários com a senha próxima de expiração e enviar um 
#   e-mail a estes usuários para realizarem a troca da senha o quanto antes.
#
# Nota: Quando editar o template do email, as seguintes variáveis serão substituídas automaticamente:
# Lembrando que o usuário deverá ter estes atributos definidos no ActiveDirectory.
#   {GivenName}		-> nome visível, geralmente nome e sobrenome
#   {Title}			-> título
#   {mail}			-> o e-mail do colaborador
#   {Manager}		-> O nome do gerente do colaborador
#   {Company}		-> O nome da empresa
#   {Department}	-> Departamento
##################################################################################################################


# Variables regarding office365 graph api to send email
$TENANT_ID     = "you-tenant-id-here"
$APP_ID        = "you-azure-app-id-here"
$APP_SECRET    = "your-app-secret-here"
$SEND_EMAIL_AS = "id@youtcompany.co"


# Amount of days that passwords expires by default on your network
# 42 defaults to Active Directory. Check the current value on your default domain group policy
$password_expire_days = 42

# Amount of days before the user passwords expires se we start warning them
$start_notify_days = 3

# Mail subject
$mail_subject = "[AVISO DA TI] Sua senha de rede está próxima de expirar."

# Name of the email message template file
$mail_template_filename = "mail_template.html"

# Name of the mime template
$mime_template_filename = "mime_template.txt"

# Extra User properties to retrieve from AD and set available for use on e-mail template
$user_custom_properties = @("GivenName", "Title", "Mail", "Manager", "Company","Department", "PasswordLastSet")




#---------------------------------------------------------------------------------------------
# Do not modify from here below, unless you know exactly what you are doing.
# I've told you -.-
#---------------------------------------------------------------------------------------------

#
# Detecting from which folder this script was invoked
#
$ScriptPath = Split-Path -Parent $MyInvocation.mycommand.Path

#
# Screen logging information messages
#
Function Log()
{
	Param([String]$text)
	$timestamp = Get-Date -Format G
	Write-Host -ForegroundColor Green "$timestamp`: $text"
	
}

#
# Screen logging error messages
#
Function LogError()
{
	Param([String]$text)
	$timestamp = Get-Date -Format G
	Write-Host -ForegroundColor Red "$timestamp`: $text"
	
}

#
# Function to streamline mail sending
#
Function SendOffice365Mail
{
	Param(
		[Parameter(Mandatory=$true)][String] $TENANT_ID,
		[Parameter(Mandatory=$true)][String] $APP_ID,
		[Parameter(Mandatory=$true)][String] $APP_SECRET,
		[Parameter(Mandatory=$true)][String] $SEND_EMAIL_AS,
		[Parameter(Mandatory=$true)][String] $SEND_EMAIL_TO,
		[Parameter(Mandatory=$true)][String] $EMAIL_SUBJECT,
		[Parameter(Mandatory=$true)][String] $EMAIL_CONTENT_TYPE, # text, html
		[Parameter(Mandatory=$true)][String] $EMAIL_BODY
	)

	
	# Important URL's
	$authorize_url = "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token"
	$send_mail_url = "https://graph.microsoft.com/v1.0/users/$SEND_EMAIL_AS/sendMail"
	
	
	# Loading email message (as mime) template
	# Loading template from file
	$EMAIL_MIME = New-Object System.Text.StringBuilder
	$template_mime_path = Join-Path -Path $ScriptPath -ChildPath $mime_template_filename
	If([System.IO.File]::Exists($template_mime_path))
	{
		try{
			$sr = New-Object System.IO.StreamReader -ArgumentList $template_mime_path
			$EMAIL_MIME.Append( $sr.ReadToEnd() ) | Out-Null
			$EMAIL_MIME.Replace("{SEND_EMAIL_AS}", $SEND_EMAIL_AS) | Out-Null
			$EMAIL_MIME.Replace("{SEND_EMAIL_TO}", $SEND_EMAIL_TO) | Out-Null
			$EMAIL_MIME.Replace("{EMAIL_SUBJECT}", $EMAIL_SUBJECT) | Out-Null
			If($EMAIL_CONTENT_TYPE.tolower() -eq "text")
			{
				$EMAIL_MIME.Replace("{MAIL_CONTENT_TYPE}", "text/plain") | Out-Null
			}Else{
				$EMAIL_MIME.Replace("{MAIL_CONTENT_TYPE}", "text/html") | Out-Null
			}
			$EMAIL_MIME.Replace("{EMAIL_BODY}", $EMAIL_BODY) | Out-Null
		}catch {
			LogError("Error reading mime template at $template_mime_path")
			LogError("Details")
			LogError($_.Exception)
			Exit(0)
		}finally{
			$sr.Close()
		}
		
	}else{
		LogError("MIME template file not found.")
		LogError("It's was excepcted in the path $template_mime_path")
		[System.Threading.Thread]::Sleep(5000)
		Exit(0)
	}


	# Getting the auth token
	$token_request = Invoke-RestMethod -URI $authorize_url `
	  -Method Post `
	  -ContentType "application/x-www-form-urlencoded" `
	  -Body @{
		"tenant"        =$TENANT_ID
		"client_id"     =$APP_ID
		"scope"         ="https://graph.microsoft.com/.default"
		"client_secret" =$APP_SECRET
		"grant_type"    ="client_credentials"
		
	  }
	  

	# Preparing request headers
	$email_send_headers = @{
		"Authorization" = "Bearer $($token_request.access_token)"
		"Content-Type"  = "text/plain"
	}

	
	# Preparing Message
	# convert to base64 and then send
	$message_base64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($EMAIL_MIME.toString()))
	try{
		Invoke-RestMethod -Method Post -uri $send_mail_url -Headers $email_send_headers -Body $message_base64
	}Catch{
		Write-Host -ForegroundColor Red $_.Exception
	}

}


# Calculating the date that 
Log "passwords default expires in $password_expire_days days"
Log "We will start notifying users $start_notify_days days before it expires"
Log "Today is $(Get-Date -Format G)"


# Processing the e-mail body message
Log "Loading default e-mail message template from file $mail_template_filename"
$base_mail_body = New-Object System.Text.StringBuilder

# Loading template from file
$template_file_path = Join-Path -Path $ScriptPath -ChildPath $mail_template_filename
If([System.IO.File]::Exists($template_file_path))
{
	try{
		$sr = New-Object System.IO.StreamReader -ArgumentList $template_file_path,$True
		$template_text = $sr.ReadToEnd()
		$base_mail_body.Append( $template_text ) | Out-Null
	}catch {
		LogError("Error reading from template file at $template_file_path")
		LogError("Details")
		LogError($_.Exception)
		Exit(0)
	}finally{
		$sr.Close()
	}
	
}else{
	LogError("Template e-mail file not found.")
	LogError("It's was excepcted in the path $template_file_path")
	[System.Threading.Thread]::Sleep(5000)
	Exit(0)
}




# Filtering users who's passwords is about to expire in the next [$start_notify_days] days
$today = Get-Date
$password_lifetime_before_warn = $password_expire_days - $start_notify_days
$password_start_warning_from = $today.AddDays($password_lifetime_before_warn * -1)
$valid_users = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -lt $password_start_warning_from } `
	-Properties $user_custom_properties -ResultSetSize $null -ResultPageSize ([int]::MaxValue)
Log([String]::Format("Found {0} users that are Enabled, Marked to not expire password and password is about to expire", $valid_users.Count))


# Processing Users
ForEach($user in $valid_users)
{
	$mail_body = New-Object System.Text.StringBuilder $base_mail_body.toString()
	
	# Processing user variables in template
	ForEach($uservar in $user_custom_properties)
	{
		$mail_body.replace("`{$uservar`}", $user.Item($uservar)) | Out-Null
	}
	If($user.PasswordLastSet -ne $null)
	{
		$mail_body.replace("`{PwExpireDate`}", $user.PasswordLastSet.AddDays($password_expire_days).toString("dd/MM/yyyy HH:mm")) | Out-Null
	}Else{
		$mail_body.replace("`{PwExpireDate`}", "*senha nunca definida") | Out-Null
	}
	
	
	
	# Notifying the target user
	Log "Sending notification (email) to $($user.Name) to change his/her password ASAP."
	SendOffice365Mail -TENANT_ID $TENANT_ID `
		-APP_ID $APP_ID `
		-APP_SECRET $APP_SECRET `
		-SEND_EMAIL_AS $SEND_EMAIL_AS `
		-SEND_EMAIL_TO $user.mail `
		-EMAIL_SUBJECT $mail_subject  `
		-EMAIL_CONTENT_TYPE "html" `
		-EMAIL_BODY $mail_body.toString()
	
}
