##################################################################################################################
# CITRA IT CONSULTING
# SCRIPT PARA NOTIFICAR AO USUÁRIO POR EMAIL QUE SUA SENHA ESTÁ PRÓXIMA DE ESPERAÇÃO.
# Author: luciano@citrait.com.br
# Date: 19/08/2020
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


# Smtp variables to send e-mail
$smtp_server = "yout.youcompany.co"
$smtp_mail_from = "notification@youcompany.co" 
$smtp_username = "notification@youcompany.co"
# smtp password should be encoded in base64.
$smtp_password = "QUJDREVGR0hJSktMTU5PUFFSU1RVVldYWVo="

# Amount of days that passwords expires by default on your network
# 42 defaults to Active Directory. Check the current value on your default domain group policy
$password_expire_days = 42

# Amount of days before the user passwords expires se we start warning them
$start_notify_days = 3

# Mail subject
$mail_subject = "[IT FREE DONUTS] Your network password is about to expire."

# Name of the email template file
$mail_template_filename = "mail_template.html"

# Extra User properties to retrieve from AD and set available for use on e-mail template
$user_custom_properties = @("GivenName", "Title", "Mail", "Manager", "Company","Department")




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
# ! gets encoded password and convert to plain
#
Function Send-Email {
    Param(
		[Parameter(Mandatory=$True)][string]$To, 
		[Parameter(Mandatory=$True)][string]$Subject, 
		[Parameter(Mandatory=$True)][string]$Body,
		[Parameter(Mandatory=$False)][string[]]$Attachments
	)

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.To.Add($to)       
    $mail.From = $smtp_mail_from
    
    $mail.IsBodyHtml = $true 
    $mail.Body = $body
    $mail.Subject = $subject
	
	foreach($anexo in $attachments)
	{
		If(! (Test-Path $anexo) ){ Write-Host "Erro ao anexar arquivo $anexo"; Continue}
		$mail.Attachments.Add($anexo)
	}
    
    $smtpclient = new-object System.Net.Mail.SmtpClient
    $smtpclient.Host = $smtp_server
    $smtpclient.Timeout = 30000  # timeout defaults to 30 seconds

    $smtpclient.EnableSSL = $true #Default all big providers like google and microsoft365 force use off SSL/TLS
    $smtpclient.Port = 587  # default smtp port.

	$final_password = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($smtp_password))
    $smtpclient.Credentials = new-object System.Net.NetworkCredential($smtp_username, $final_password) 
 
    try{
		$smtpclient.Send($mail)
	}Catch{
		LogError $_.Exception
	}
   
}


# Calculating the date that 
Log "passwords default expires in $password_expire_days days"
Log "We will start notifying users $start_notify_days days before it expires"
Log "Today is $(Get-Date -Format G)"


# Processing the e-mail body message
Log "Loading default e-mail message template from file $mail_template_filename"
$mail_body = New-Object System.Text.StringBuilder

# Loading template from file
$template_file_path = Join-Path -Path $ScriptPath -ChildPath $mail_template_filename
If([System.IO.File]::Exists($template_file_path))
{
	try{
		$sr = [System.IO.StreamReader]::New($template_file_path)
		$template_text = $sr.ReadToEnd()
		$mail_body.Append( $template_text ) | Out-Null
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
	# Processing user variables in template
	ForEach($uservar in $user_custom_properties)
	{
		$mail_body.replace("`{$uservar`}", $user.Item($uservar)) | Out-Null
	}
	
	
	# Notifying the target user
	Log "Sending notification (email) to $($user.Name) to change his/her password ASAP."
	Send-Email -To $user.mail -Subject $mail_subject -Body $mail_body.toString()
	
}
