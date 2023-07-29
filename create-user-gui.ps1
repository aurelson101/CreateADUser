Add-Type -AssemblyName System.Windows.Forms
 # Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Account Creation"
$form.Size = New-Object System.Drawing.Size(400,500)
 
 # Create labels and textboxes for user input
$labelFirstName = New-Object System.Windows.Forms.Label
$labelFirstName.Text = "First Name:"
$labelFirstName.Location = New-Object System.Drawing.Point(10,10)
$form.Controls.Add($labelFirstName)
 $textBoxFirstName = New-Object System.Windows.Forms.TextBox
$textBoxFirstName.Location = New-Object System.Drawing.Point(150,10)
$textBoxFirstName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxFirstName)

 $labelSecondFirstName = New-Object System.Windows.Forms.Label
$labelSecondFirstName.Text = "Second First Name:"
$labelSecondFirstName.Location = New-Object System.Drawing.Point(10,40)
$form.Controls.Add($labelSecondFirstName)
 $textBoxSecondFirstName = New-Object System.Windows.Forms.TextBox
$textBoxSecondFirstName.Location = New-Object System.Drawing.Point(150,40)
$textBoxSecondFirstName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxSecondFirstName)

 $labelLastName = New-Object System.Windows.Forms.Label
$labelLastName.Text = "Last Name:"
$labelLastName.Location = New-Object System.Drawing.Point(10,70)
$form.Controls.Add($labelLastName)
 $textBoxLastName = New-Object System.Windows.Forms.TextBox
$textBoxLastName.Location = New-Object System.Drawing.Point(150,70)
$textBoxLastName.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxLastName)

 $labelManagerADAccount = New-Object System.Windows.Forms.Label
$labelManagerADAccount.Text = "Manager AD Account:"
$labelManagerADAccount.Location = New-Object System.Drawing.Point(10,100)
$form.Controls.Add($labelManagerADAccount)
 $textBoxManagerADAccount = New-Object System.Windows.Forms.TextBox
$textBoxManagerADAccount.Location = New-Object System.Drawing.Point(150,100)
$textBoxManagerADAccount.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxManagerADAccount)

 $labelEmployeeID = New-Object System.Windows.Forms.Label
$labelEmployeeID.Text = "Employee ID:"
$labelEmployeeID.Location = New-Object System.Drawing.Point(10,130)
$form.Controls.Add($labelEmployeeID)
 $textBoxEmployeeID = New-Object System.Windows.Forms.TextBox
$textBoxEmployeeID.Location = New-Object System.Drawing.Point(150,130)
$textBoxEmployeeID.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxEmployeeID)

 $labelJob = New-Object System.Windows.Forms.Label
$labelJob.Text = "Job:"
$labelJob.Location = New-Object System.Drawing.Point(10,160)
$form.Controls.Add($labelJob)
 $textBoxJob = New-Object System.Windows.Forms.TextBox
$textBoxJob.Location = New-Object System.Drawing.Point(150,160)
$textBoxJob.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxJob)

 $labelFunction = New-Object System.Windows.Forms.Label
$labelFunction.Text = "Function:"
$labelFunction.Location = New-Object System.Drawing.Point(10,190)
$form.Controls.Add($labelFunction)
 $textBoxFunction = New-Object System.Windows.Forms.TextBox
$textBoxFunction.Location = New-Object System.Drawing.Point(150,190)
$textBoxFunction.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxFunction)

 $labelContractType = New-Object System.Windows.Forms.Label
$labelContractType.Text = "Contract Type:"
$labelContractType.Location = New-Object System.Drawing.Point(10,220)
$form.Controls.Add($labelContractType)
 $comboBoxContractType = New-Object System.Windows.Forms.ComboBox
$comboBoxContractType.Location = New-Object System.Drawing.Point(150,220)
$comboBoxContractType.Size = New-Object System.Drawing.Size(200,20)
$comboBoxContractType.Items.Add("CDI")
$comboBoxContractType.Items.Add("CDD")
$comboBoxContractType.Items.Add("EXT")
$form.Controls.Add($comboBoxContractType)
 $labelContractEndDate = New-Object System.Windows.Forms.Label
$labelContractEndDate.Text = "Contract End Date:"
$labelContractEndDate.Location = New-Object System.Drawing.Point(10,250)
$form.Controls.Add($labelContractEndDate)
 $textBoxContractEndDate = New-Object System.Windows.Forms.TextBox
$textBoxContractEndDate.Location = New-Object System.Drawing.Point(150,250)
$textBoxContractEndDate.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxContractEndDate)

 $buttonCreateAccount = New-Object System.Windows.Forms.Button
$buttonCreateAccount.Text = "Create Account"
$buttonCreateAccount.Location = New-Object System.Drawing.Point(150,280)
$buttonCreateAccount.Size = New-Object System.Drawing.Size(100,30)
$form.Controls.Add($buttonCreateAccount)
# Event handler for the Create Account button
$buttonCreateAccount.Add_Click({
    Import-Module ActiveDirectory
    #EDIT your exchange or delete if not need
    $Session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://myexchange.home.net/powershell
    Import-PSSession $Session
    Set-ADServerSettings -ViewEntireForest $True
    $upnsuffix = "@home.net" #EDIT by your domaine directory
    $securepwd = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force #EDIT the password by default
    $firstName = $textBoxFirstName.Text
    $secondFirstName = $textBoxSecondFirstName.Text
    $lastName = $textBoxLastName.Text
    $managerADAccount = $textBoxManagerADAccount.Text
    $employeeID = $textBoxEmployeeID.Text
    $job = $textBoxJob.Text
    $function = $textBoxFunction.Text
    $contractType = $comboBoxContractType.SelectedItem.ToString()
    $contractEndDate = $textBoxContractEndDate.Text
    $lastNamenom = $lastName.substring(0,1).toupper()+$lastName.substring(1).tolower()
    $secondNamenom = $secondFirstName.substring(0,1).toupper()+$secondFirstName.substring(1).tolower()
    $firstNamenom = $firstName.substring(0,1).toupper()+$firstName.substring(1).tolower()
if ($contractType -eq "CDI" -or $contractType -eq "CDD") {
    $displayname = "$lastNamenom, $secondNamenom $firstNamenom (PHS)" #EDIT (PHS) by your organisation
    
} else {
    $displayname = "$lastNamenom, $secondNamenom $firstNamenom (PHS-EXT)" #EDIT (PHS) by your organisation
}


#EDIT where your compagny localisation 
    $description = "CDI"
    $bureau = "TECH"
    $adresse = "8 Rue Chaptal"
    $ville = "Paris"
    $company = "Bytoprof"
    $countryCode = "75013"
    $country = "FR"
    $monpoint = "."
    $monmoin = "-"
     #Display the email address
 if ($secondFirstName -ne "") {
    $email = "$firstName-$secondFirstName.$lastName@exemple.fr" #Edit by your email
} else {
    $email = "$firstName.$lastName@exemple.fr" #Edit by your email
}
 Write-Host "Email address: $email"
     # Create the login
    $login = $lastName + $firstName.Substring(0,1)
     # Check if the login exists, if so, add a number after the first letter of the first name
$i=2
$origin = $login
while (get-aduser -identity $login)
{
    $login = $origin
    $login = "$login" + "$i"
    $i = $i + 1
}
$login2 = "$login"+"$upnsuffix"
     # Create the user account in the specified OU
     
    $ou = "OU=Users,OU=User Accounts,DC=b2b,DC=home,DC=net" # Replace with your actual OU path
    $userParams = @{
        SamAccountName = $login
        Name = $displayname
        GivenName = $firstNamenom
        Surname = $lastNamenom
        DisplayName = "$displayname"
        UserPrincipalName = "$login"+"$upnsuffix"
        EmailAddress = $email
        Manager = $managerADAccount
        EmployeeID = $employeeID
        Title = $job
        Description = $contractType
        Department = $function
        Path = $ou
        AccountPassword=  $securepwd
        Enabled = $True
        Country = $country
        Office = $bureau 
        StreetAddress = $adresse 
        City = $ville 
        PostalCode = $countryCode 
        Company = $company 

    }
     if ($contractType -eq "EXT" -or $contractType -eq "CDD") {
        $userParams.Add("AccountExpirationDate", $contractEndDate)
    }
New-ADUser @userParams
#Edit by your GROUP AD
Add-ADPrincipalGroupMembership -Identity "$login" -MemberOf "GROUPA","GROUPB","GROUPC","GROUPD","GROUPE","GROUPF","GROUPG","GROUPH","GROUPI","GROUPJ" 
#Edit by your waiting minute you can remove if not need
$minutes = 40
$seconds = $minutes * 60
 $progressbar = New-Object System.Windows.Forms.ProgressBar
$progressbar.Location = New-Object System.Drawing.Point(10,380)
$progressbar.Size = New-Object System.Drawing.Size(280,30)
$progressbar.Minimum = 0
$progressbar.Maximum = $seconds
$progressbar.Step = 1
$form.Controls.Add($progressbar)
 $label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,360)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = "Starting..."
$form.Controls.Add($label)
 $form.Show()
 for ($i = 0; $i -le $seconds; $i++) {
    Start-Sleep -Seconds 1
    $progressbar.PerformStep()
    $label.Text = "Sleeping for $minutes minutes - $([math]::Round($i / $seconds * 100))% Complete"
    $form.Refresh()
}

#Edit by your email configuration, or delete if not need

Enable-RemoteMailbox -Identity $login@b2b.home.net -PrimarySmtpAddress $email -RemoteRoutingAddress o365phs-$login@mycompagny.mail.onmicrosoft.com -DomainController DCCONTROLLER.b2b.home.net
Enable-RemoteMailbox -Identity $login@b2b.home.net -Archive -DomainController DCCONTROLLER.b2b.home.net
Set-RemoteMailbox -Identity $login@b2b.home.net -EmailAddresses @{add=”o365phs-$login@mycompagny.mail.onmicrosoft.com”} -DomainController DCCONTROLLER.b2b.home.net

Remove-PSSession $session
     # Display a success message
    [System.Windows.Forms.MessageBox]::Show("User account created successfully!", "Success")

})
 # Show the form
$form.ShowDialog()