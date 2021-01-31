###-----------------------------------------------------------------------------------------------###
### ONGLET AD
###-----------------------------------------------------------------------------------------------###

# ---------------------------------------------------------
#Déclaration des Controleur AD
# ---------------------------------------------------------
#SUPPRIME

#---------------------------------------------------------------------------------------------------#
# Fonctions
#---------------------------------------------------------------------------------------------------#
Function AgeMDP {
$lbl_ad_agepassword.Text = ""
[datetime]$today = (get-date)
$passdate2 = [datetime]::fromfiletime((Get-ADUser -Filter {samAccountName -eq $Trigramme} -Properties * | 
    Select-Object -expand pwdlastset))
$PwdAge = ($today - $passdate2).Days
$script:TrigMAJ = $Trigramme.toupper()

If ($PwdAge -gt 90){
		$Global:Age =  "Age du mot de passe : $PwdAge jours (EXPIRE)"


}else{
        $Global:Age  = "Age du mot de passe : $PwdAge jours"
	}
	$lbl_ad_agepassword.Text = "$Global:Age"
	echl "[ii]`t Age du mot de passe de l'UT: $Global:Age"
}

Function VerifCompteWind  {
ControlAD
$script:Lock1 = Get-ADUser -Filter {samAccountName -eq $Trigramme } -Properties lockedout,enabled


If($Lock1.enabled -match "False"){$Global:Etat = "Compte désactivé"} 
Else {
If($Lock1.lockedout -match "False"){$Global:Etat = "Le compte n'est pas verrouillé"}
if($Lock1.lockedout -match "True"){$Global:Etat = "Le compte est verrouillé"}
	}
$lbl_ad_verifad.Text = "$Global:Etat"
echl "[ii]`t $Global:Etat"
}

Function DevAD {
$Global:cred_glb = Get-Cred("glb.intra.groupama.fr")
Unlock-ADAccount -Identity $trigramme -Credential $cred_glb

$script:Lock = Get-ADUser -Filter {samAccountName -eq $Trigramme } -Properties * | 
            Select-Object -expand lockedout | Out-String

If($Lock -match "False"){$Global:Etat ="Le compte $Trigramme a été déverrouillé."}
if($Lock -match "True"){$Global:Etat ="Le compte $Trigramme est toujours verrouillé"} }

Function ReinitAD {
ControlAD
$Global:cred_glb = Get-Cred("glb.intra.groupama.fr")
$new_pass = $txtb_ad_mdp.text

Set-ADAccountPassword $trigramme -NewPassword (ConvertTo-SecureString -AsPlainText $new_pass -Force) -PassThru -Reset -Credential $cred_glb
				Set-ADuser -Identity $trigramme -ChangePasswordAtLogon $True -Credential $cred_glb

}

Function ConsAD {
$Global:cred_glb = Get-Cred("domain")
Start-Process “C:\Windows\System32\cmd.exe” -workingdirectory $PSHOME -ArgumentList “/c dsa.msc” -Credential $cred_glb -NoNewWindow
}
