Import-Module ActiveDirectory
#EDIT your exchange or delete if not need
$Session = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri http://myexchange.home.net/powershell
Import-PSSession $Session
$ou = "OU=User Accounts,DC=home,DC=net" # Replace with your actual OU path
$upnsuffix = "@home.net" # Replace with your actual OU path
$securepwd = ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force #EDIT the password by default
$contractType = Read-Host "Entrez le type de contrat (CDI, CDD, EXT)"
if ($contractType -eq "CDI") {
    $endDate = "Never"
}
elseif ($contractType -eq "CDD") {
    $endDate = Read-Host "Entrez la date de fin du CDD (format : JJ/MM/AAAA)"
}
elseif ($contractType -eq "EXT") {
    $endDate = Read-Host "Entrez la date de fin de l'EXT (format : JJ/MM/AAAA)"
}
$prenom = Read-Host "Entrez le prenom"
$secondname = Read-Host "Il y a un deuxieme prenom ? (oui ou non)"
if ($secondname -eq "oui") {
    $secondname2 = Read-Host "Entrez le deuxieme prenom"
    $nom = Read-Host "Entrez le nom de famille"
    $email = "$prenom-$secondname2.$nom@mydomain.fr" #Edit by your email
}

if ($secondname -eq "non") {
$nom = Read-Host "Entrez le nom de famille"
    $email = "$prenom.$nom@mydomain.fr" #Edit by your email
}
else {

}

Write-Host "ladresse e-mail est : $email"
$lastNamenom = $nom.substring(0,1).toupper()+$nom.substring(1).tolower()
$secondfirstname = $secondname2.substring(0,1).toupper()+$secondname2.substring(1).tolower()
$firstNamenom = $prenom.substring(0,1).toupper()+$prenom.substring(1).tolower()

if ($contractType -eq "CDI" -or $contractType -eq "CDD") {
    $displayname = "$lastNamenom, $secondfirstname $firstNamenom (B2B)"
    
}# else {
#    $displayname = "$lastNamenom, $secondfirstname $firstNamenom (RX-EXT)"
if ($contractType -eq "EXT") {
    $displayname = "$lastNamenom, $secondfirstname $firstNamenom (B2B-EXT)"
    

}

#EDIT where your compagny localisation 
$managerADAccount = Read-Host "Ecrire l'identifiant AD du manager"
$service = Read-Host "Ecrire la Fonction"
$Department = Read-Host "Ecrire le Departement"
$ID = Read-Host "Employee ID"
$description = "CDI"
$bureau = "MyCompagny"
$adresse = "18 Rue du sandwich"
$ville = "Paris"
$company = "MyCompagny"
$countryCode = "75013"
$country = "FR"
$monpoint = "."
$monmoin = "-"
# Create the login
$login = $lastNamenom + $firstNamenom.Substring(0,1).tolower()
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
$GiveName = "$lastNamenom"
$GiveSurname = "$firstNamenom"

 if ($contractType -eq "EXT" -or $contractType -eq "CDD") {
    New-ADUser -Name "$displayname" -UserPrincipalName "$login"  -SamAccountName "$login" -GivenName "$GiveName" -Surname "GiveSurname" -DisplayName "$displayname" -Path "$ou" -AccountPassword $securepwd -Enabled $True -Description $contractType -Office $bureau -StreetAddress $adresse -City $ville -EmailAddress $email -PostalCode $countryCode -Country "FR" -Company $company -title $service -Department $Department -Manager $managerADAccount -EmployeeID $ID -AccountExpirationDate "$endDate 23:00 PM"
}
if ($contractType -eq "CDI") {
    New-ADUser -Name "$displayname" -UserPrincipalName "$login"  -SamAccountName "$login" -GivenName "$GiveName" -Surname "GiveSurname" -DisplayName "$displayname" -Path "$ou" -AccountPassword $securepwd -Enabled $True -Description $contractType -Office $bureau -StreetAddress $adresse -City $ville -EmailAddress $email -PostalCode $countryCode -Country "FR" -Company $company -title $service -Department $Department -Manager $managerADAccount -EmployeeID $ID
}
#Edit by your GROUP AD
Add-ADPrincipalGroupMembership -Identity "$login" -MemberOf "MYGROUPAD","MYGROUPAD","MYGROUPAD"

#Edit by your waiting minute you can remove if not need
$minutes = 40
$seconds = $minutes * 60
for ($i = 0; $i -le $seconds; $i++) {
    Start-Sleep -Seconds 1
    Write-Progress -Activity "Sleeping for $minutes minutes" -Status "$([math]::Round($i / $seconds * 100))% Complete" -PercentComplete ($i / $seconds * 100)
}


#Edit by your email configuration, or delete if not need
Enable-RemoteMailbox -Identity $login@b2b.home.net -PrimarySmtpAddress $email -RemoteRoutingAddress o365phs-$login@mycompagny.mail.onmicrosoft.com -DomainController DCCONTROLLER.b2b.home.net
Enable-RemoteMailbox -Identity $login@b2b.home.net -Archive -DomainController DCCONTROLLER.b2b.home.net
Set-RemoteMailbox -Identity $login@b2b.home.net -EmailAddresses @{add=”o365phs-$login@mycompagny.mail.onmicrosoft.com”} -DomainController DCCONTROLLER.b2b.home.net

write-host "Attendre 2h pour que la boite mail fonctionne"
Remove-PSSession $session
