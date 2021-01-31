<#
.SYNOPSIS
PopupMDP - Alerte Utilisateur sur Age du mot de passe - 15 jours avant Expiration
Possibilité de Fermer la fenêtre, de Reporter ou de changer le mot de passe Windows

.DESCRIPTION
Affichage d'une POPUP d'alerte 15 jours avant l'expiration du mot de passe :
Si Changer : Affichage d'une fenêtre permettant de changer le mot de passe Windows
			Avec vérification de la conformité du mot de passe 
			Et fermeture de la session une fois le mot de passe changé
Si Reporter : Report de la POPUP à intervalle prédéfini : 1h/2h/4h/8h
Si Fermer : Fermeture de la POPUP - réafficahge le lendemain

Ne s'affiche pas en TRAMO/TELETRAVAIL

.NOTES
Version : 2.0 - Refonte de l'interface graphique avec Popup et Fenêtre
Date de création : 27/01/2020
Compatibilité : Powershell V5

Version 2.1 - 03/12/2020 : Ajout d'un contrôle pour TRAMO/TELETRAVAIL
	=> SI TRAMO/TELETRAVAIL : Affichage de la POPUP
		Si clique sur Changer : Affichage de la fenêtre avec une alerte et un lien
		pour modification du mot de passe en TRAMO/TELETRAVAIL
		=> Blocage du changement de mot de passe via la fenêtre

#>


Param
 (
 [String]$Restart,
 [int]$time_to_sleep = 3600
 
 )
 
If ($Restart -ne "") 
 {
  Start-Sleep $time_to_sleep
 }


$Global:Current_Folder = $PSScriptRoot 
Set-Location $Current_Folder
$Assembly_Folder = Join-Path -Path $Current_Folder -ChildPath .\assembly
$Resources_Folder = Join-Path -Path $Current_Folder -ChildPath .\Resources

# Import Module Log
Import-Module Logs_TA
Import-Module ActiveDirectory
Create_Log -Path_Log $Current_Folder -Appli PopupMDP

$script:count = 0
Log -Add "Lancement du script"
Log -Add "Chargement des composants"
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null

[System.Windows.Forms.Application]::EnableVisualStyles()


If(Test-Path $Assembly_Folder) {
	Foreach ($Assembly in (Dir $Assembly_Folder -Filter *.dll)) {
		$null = [System.Reflection.Assembly]::LoadFrom($Assembly.fullName)
	}
}
 
$Global:Current_Folder = $PSScriptRoot

Function LoadXaml($Global:filename) {
	$xamlLoader = (New-Object System.Xml.XmlDocument)
	$xamlLoader.Load($filename)
	$xamlLoader
}

#
# Chargement xaml des fenetres
#
$xaml = LoadXaml("$Current_Folder\Notif_Window.xaml") 
$xamlReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Users_Window = [Windows.Markup.XamlReader]::Load($xamlReader)

$xaml_change = LoadXaml("$Current_Folder\Password_Window.xaml") 
$xamlReader_change = (New-Object System.Xml.XmlNodeReader $xaml_change)
$Change_Window = [Windows.Markup.XamlReader]::Load($xamlReader_change)

#	FONCTIONS
#-------------------------------------------------------------------------

Function Launch_Toast {

			# AFFICHAGE BOUTON ANNULER OU PAS => SI PASSWORD EXPIRE J=0
		#$Time_limite = Verif_Password
		if ( $Time_limite -le 1 ) { $btn_cancel.IsEnabled = $false
			$btn_close.IsEnabled = $false
			$btn_report.IsEnabled = $false }
		Else {
		$btn_cancel.IsEnabled = $true
		$btn_close.IsEnabled = $true
		$btn_report.IsEnabled = $true 
		}

		if (($Time_limite -le 15) -and ($Time_limite -ge 10)){  $icon = "$Resources_Folder\shield_alert_green.ico"}
		if (($Time_limite -le 9) -and ($Time_limite -ge 5)){  $icon = "$Resources_Folder\shield_alert_orange.ico"}
		if (($Time_limite -le 4) -and ($Time_limite -ge 0)){  $icon = "$Resources_Folder\shield_alert_red.ico"}

		# AFFICHAGE TEXTBOX + LABEL
		$lbl_name.Content = "Bonjour $($user.GivenName) $($user.Surname)"
		if ($Expire_date_show -eq (Get-Date)) {
		$txtb_run_1.Text = "Votre mot de passe de session Windows va expirer"
		$txtb_run_2.Text = "aujourd'hui. Pour le changer dès maintenant,"
		$txtb_run_3.Text = "cliquez sur le bouton 'Changer'"
		$lbl_expire_date.Content = "à $($expire_hour)"
		}
		else {
		$txtb_run_1.Text = "Votre mot de passe de session Windows va bientôt "
		$txtb_run_2.Text = "expirer. Pour le changer dès maintenant, cliquez"
		$txtb_run_3.Text = "sur le bouton 'Changer'"
		$lbl_expire_date.Content = "le $($Expire_date_show)"
		}

	#Cacher fenetre powershell
	$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
	$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
	$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
	
	# Limite l'utilisation de la RAM
	[System.GC]::Collect()
	
	$Users_Window.WindowStartupLocation = "Manual"	
	$Users_Window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Users_Window.Width)
	$Users_Window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Users_Window.Height)	
	
	$Users_Window.Show()
	$Users_Window.Activate()	
	
	$appContext = New-Object System.Windows.Forms.ApplicationContext
	[void][System.Windows.Forms.Application]::Run($appContext)
	}

Function Launch_Or_Not_Launch {
	$script:check_pwd_age = $script:Time_limite =  Verif_Password
	#$script:Time_limite = Verif_Password

    Log -Add "Lancement Fonction Launch_Or_Not_Launch"
    #if ( $script:check_pwd_age -le 90) { Log -Add "Lancement du Toast => Pour test - limite d'affichage mise à 90 jours" # A COMMENTER AVANT PASSAGE EN PROD
    if (  $script:check_pwd_age -le 15) { Log -Add "Lancement du Toast => Expiration dans 15 jours ou moins" # A DECOMMENTER AVANT PASSAGE EN PROD
    Launch_Toast }
    Else { Log -Add "Expiration du mot de passe supérieur à 15 jours - pas de lancement"
        exit }
}

Function Verif_Win { 
    if ([System.Environment]::OSVersion.Version.Major -eq 10 ){ 
        Log -Add "Poste en W10 - Poursuite du script"
        return 10 }
    Else { Log -Add "Poste en W7 - Arret du script"
        return 7 }
}


Function Snooze {
	Param(
		[Parameter(Position=0,Mandatory=$false)]
		[int]$time_to_snooze = 3600
	)
	$Users_Window.Hide()
	$Change_Window.Hide()
	$Main_Tool_Icon.Visible = $false
	Log -Add "Lancement Snooze - duree : $($time_to_snooze) secondes"
	Start-Sleep -Seconds $time_to_snooze
	Log -Add "Fin Snooze - duree : $($time_to_snooze) secondes"
	
	try{
	powershell -file "$($Current_Folder)\Popupmdp_notif.ps1"
	} catch { Log -Add $Error[0] }
	
}

Function Check_Old {
	Param (
		[Parameter(Position=0,Mandatory=$true,HelpMessage='Renseigner votre mot de passe')]
		[string]$old_pwd
		)
	Log -Add "Verification du mot de passe"	
	$UserName = $env:USERDOMAIN +"\"+$env:USERNAME
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement
	$contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
	$principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $contextType,$env:USERDOMAIN
	$principalContext.ValidateCredentials($UserName,$old_pwd)


}

Function Verif_Password {
    Log -Add "Verification Age du mot de passe"
    $script:user = Get-ADUser $env:USERNAME -Properties passwordlastset
    $Expire_date = ($user.passwordlastset).adddays(90)
    $script:Expire_date_show = Get-Date $Expire_date -Format dd/MM/yyyy
    $script:expire_hour = Get-Date $Expire_date -Format "HH:mm"
    $limite_date = ($Expire_date - (Get-Date)).days
    #if ($limite_date -le 90) { Log -Add "Expiration du mot de passe dans $($limite_date) jours"}
    Log -Add "Expiration du mot de passe dans $($limite_date) jours"
    return $limite_date
    }

Function Test_IP {
	
	try {
		if (Verif_Win -eq 10) {
			Log -Add "Poste en W10 - Poursuite du script"
	
			Log -Add "Verification Reseau"
			$Connexion = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null } 
			if ($null -eq $Connexion){ Log -Add "Pas IP"
							Start-Sleep -Seconds 20
							$count++
							Log -Add "Execution : $($count)"
							Do {		
							Test_IP }
							 until ($count -eq 10 )
	
						}
			Elseif ( $Connexion.IPAddress.Count -ne 1 ) {
				Log -Add "Adresse IP : $($Connexion.IPAddress) "
				switch -Wildcard ($Connexion.IPAddress[1]) {
					"10.227.*" { Log -Add "En Teletravail - EndUser - Exécution du script"
							Launch_Or_Not_Launch }
					"10.210.*" { Log -Add "En Teletravail - Pulse 210 - Exécution du script"
								Launch_Or_Not_Launch }
					"10.208.*" { Log -Add "En Teletravail - Pulse 208  - Exécution du script"
								Launch_Or_Not_Launch }
					Default { Log -Add "En Télétravail - hors réseau Groupama - Boucle"
						Start-Sleep -Seconds 20
						$count++
						Log -Add "Execution : $($count)"
						Do {		
						Test_IP }
						until ($count -eq 10 )}
				}
			}
			else { 
	
			Log -Add "Adresse IP : $($Connexion.IPAddress) "
			#$script:Time_limite = Verif_Password
			switch -Wildcard ($Connexion.IPAddress) {
				"10.1*" { Log -Add "Sur Site - Exécution du script" 
							Launch_Or_Not_Launch }
				"10.22.*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch }
				"10.29*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch }
				"10.35*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch }
				"10.44*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch }
				"10.49*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch  }
				"10.56*" { Log -Add "En Agence - Exécution du script"
							Launch_Or_Not_Launch }
			
				Default { Log -Add "En Télétravail ou hors réseau Groupama - Pas de lancement du script"
							Start-Sleep -Seconds 20
							$count++
							Do {		
							Test_IP }
							until ($count -eq 10 ) }
			} # SWITCH
		} #ELSE IP
			 } # IF
		Else { Log -Add "Poste en W7 - Arret du script" 
				exit  } # ELSE
		} # TRY
		catch {Log -Add $Error[0] }
	}
	
	



$btn_close = $Users_Window.FindName("btn_close")
$btn_report = $Users_Window.FindName("btn_report")
$btn_change = $Users_Window.FindName("btn_change")

$cbx_snooze = $Users_Window.FindName("cbx_snooze")

$txtb_msg = $Users_Window.FindName("txtb_msg")
	$txtb_run_1 = $Users_Window.FindName("txtb_run_1")
	$txtb_run_2 = $Users_Window.FindName("txtb_run_2")
	$txtb_run_3 = $Users_Window.FindName("txtb_run_3")
	$lbl_expire_date = $Users_Window.FindName("lbl_expire_date")
	$lbl_name = $Users_Window.FindName("lbl_name")

#Load Template pour Dialog MAHAPPS
$xamlDialog  = LoadXaml("$Current_Folder\template\Dialog.xaml")
$read = (New-Object System.Xml.XmlNodeReader $xamlDialog)
$DialogForm=[Windows.Markup.XamlReader]::Load( $read )
$CustomDialog = [MahApps.Metro.Controls.Dialogs.CustomDialog]::new($Change_Window)
$CustomDialog.AddChild($DialogForm)

##########################################################################################
##########################################################################################
                                #WINDOW CHANGE PASSWORD#
##########################################################################################
##########################################################################################

$pwd_old = $Change_Window.FindName("Pb_old")
$pwd_new = $Change_Window.FindName("Pb_new")
$pwd_conf = $Change_Window.FindName("Pb_confirm")
$btn_ok = $Change_Window.FindName("btn_ok")
$btn_cancel = $Change_Window.FindName("btn_cancel")

$btn_help = $Change_Window.FindName("btn_help")

$btn_help.ToolTip = "Votre mot de passe doit respecter les principes suivants :

•contenir 8 caractères minimum 

• contenir au moins 3 catégories des 4 catégories listées ci-dessous : 

	- contenir au moins une majuscule (A-Z)
	- contenir au moins une minuscule (a-z)
	- contenir au moins un chiffre (0-9)
	- contenir au moins un caractère spécial ( ! ? / *  etc…)

•ne pas contenir :
	- tout ou partie de votre nom d’utilisateur (exemple : FKM18022 ou 18022)
	- tout ou partie de votre nom complet (prénom et/ou nom)

•ne pas réutiliser les 14 précédents mots de passe
"
######################################################
$ico_old_ok = $Change_Window.FindName("ico_old_ok")
$ico_new_ok = $Change_Window.FindName("ico_new_ok")
$ico_conf_ok = $Change_Window.FindName("ico_conf_ok")
$lbl_old_ko  = $Change_Window.FindName("lbl_old_ko")
$lbl_new_ko  = $Change_Window.FindName("lbl_new_ko")
$lbl_conf_ko  = $Change_Window.FindName("lbl_conf_ko")
######################################################
$BtnClose = $DialogForm.FindName("BtnClose")
$BtnClose_All = $DialogForm.FindName("BtnClose_All")
$Title_Alerte = $DialogForm.FindName("Title_Alerte")
$Content_Alerte = $DialogForm.FindName("Content_Alerte")
$Content_Alerte_2 = $DialogForm.FindName("Content_Alerte_2")
$Content_Alerte_3 = $DialogForm.FindName("Content_Alerte_3")
######################################################
$flyout = $Change_Window.FindName("flyout")
$window_grid = $Change_Window.FindName("window_grid")
$window_grid.add_MouseDown({ if ($flyout.isopen -eq $true)
	{ $flyout.isopen = $false }
})
$flyout.add_MouseDown({ if ($flyout.isopen -eq $true)
	{ $flyout.isopen = $false
	} })



###########################################



##########################################################################################

##########################################################################################





#	ACTIONS
#-------------------------------------------------------------------------
$pwd_old.add_LostFocus({ 
	[String]$old = $pwd_old.get_password()
	if ($old -eq "" -or $old -eq $null) { 
		$ico_old_ok.Kind = "ErrorOutline"
		$ico_old_ok.Foreground = "DarkGoldenrod"
		$ico_old_ok.ToolTip = "Le champs ne doit pas être vide !"
		$ico_old_ok.Visibility = "Visible"
		$lbl_old_ko.Content = "Le champs ne doit pas être vide !"
		$lbl_old_ko.Foreground = "DarkGoldenrod"
		$lbl_old_ko.Visibility = "Visible"
		Log -Add "PWD Actuel : Erreur de saisie => Champs vide"
	}
	Else {
		if (( Check_Old -old_pwd $old ) -eq $false) { #Write-Host "Bad Password"
						$ico_old_ok.Kind = "Close"
						$ico_old_ok.Foreground= "Red"
						$ico_old_ok.ToolTip = "Erreur de saisie !"
						$ico_old_ok.Visibility = "Visible"
						$lbl_old_ko.Foreground = "Red"
						$lbl_old_ko.Content = "Erreur de saisie"
						$lbl_old_ko.Visibility = "Visible" 
						Log -Add "PWD Actuel : Erreur de saisie => Mot de passe errone" }
		Else { #Write-Host "Good Password"
				$lbl_old_ko.Visibility = "Hidden"
				$ico_old_ok.Kind = "Check"
				$ico_old_ok.Foreground= "Green"
				$ico_old_ok.visibility = "Visible"
				$ico_old_ok.ToolTip = "Mot de passe valide !"
				Log -Add "PWD Actuel : Mot de passe valide" }
	}
})

$pwd_new.add_LostFocus({
	[String]$New = $pwd_new.get_password()
	$pass_complex = '^((?=.*[a-z])(?=.*[A-Z])(?=.*\d)|(?=.*[a-z])(?=.*[A-Z])(?=.*[^A-Za-z0-9])|(?=.*[a-z])(?=.*\d)(?=.*[^A-Za-z0-9])|(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z0-9]))([A-Za-z\d@#$%^&£*éèêëîïùûüàâäöô\-_+=[\]{}|\\:,?/`~"()''\s;!]|\.(?!@)){8,}$'
	$user = Get-ADUser $env:USERNAME
	if ($New -eq "" -or $New -eq $null) {$ico_new_ok.Kind = "ErrorOutline"
											$ico_new_ok.Foreground = "DarkGoldenrod"
											$ico_new_ok.ToolTip = "Le champs ne doit pas être vide !"
											$ico_new_ok.Visibility = "Visible"
											$lbl_new_ko.Foreground = "DarkGoldenrod"
											$lbl_new_ko.Content = "Le champs ne doit pas être vide !"
											$lbl_new_ko.Visibility = "Visible"
											Log -Add "PWD New : Erreur de saisie => Champs vide"
										}
	Else {
		if (($new -cmatch $pass_complex) -and ($new -notmatch $env:USERNAME ) -and ($new -notmatch $user.GivenName) -and ($new -notmatch $user.Surname) )
		{ $ico_new_ok.Kind = "Check"
			$ico_new_ok.Foreground= "Green"
			$ico_new_ok.visibility = "Visible"
			$ico_new_ok.ToolTip = "Bravo !" 
			$lbl_new_ko.Visibility = "Hidden"
			Log -Add "PWD New : Mot de passe valide" }
		Else { $ico_new_ok.Kind = "Close"
			$ico_new_ok.Foreground= "Red"
			$ico_new_ok.ToolTip = "Le mot de passe ne respecte pas les règles de complexité définies"
			$ico_new_ok.Visibility = "Visible"
			$lbl_new_ko.Content = "Ne respecte pas la politique de complexité définie"
			$lbl_new_ko.Foreground = "Red"
			$lbl_new_ko.Visibility = "Visible"
			Log -Add "PWD New : Erreur de saisie => Ne respecte pas la politique de complexite definie"
		}
	}
})

$pwd_conf.add_LostFocus({
	[String]$New = $pwd_new.get_password()
	[String]$Confirm = $pwd_conf.get_password()
	Write-Host $New
	Write-Host $Confirm
	if (($New -eq "" -or $New -eq $null) -and ($Confirm -eq "" -or $Confirm -eq $null)) {
					$ico_conf_ok.Kind = "ErrorOutline"
					$ico_conf_ok.Foreground= "DarkGoldenrod"
					$ico_conf_ok.ToolTip = "Le champs ne doit pas être vide !"
					$ico_conf_ok.Visibility = "Visible"
					$lbl_conf_ko.Content = "Le champs ne doit pas être vide."
					$lbl_conf_ko.Foreground= "DarkGoldenrod"
					$lbl_conf_ko.Visibility = "Visible"
					Log -Add "PWD Conf : Erreur de saisie => Champs vide"
	}
	else {
		if ($Confirm -eq $new ) { 
			$ico_conf_ok.Kind = "Check"
			$ico_conf_ok.Foreground= "Green"
			$ico_conf_ok.visibility = "Visible"
			$ico_conf_ok.ToolTip = "Bravo !"
			$lbl_conf_ko.Visibility = "Hidden"
			Log -Add "PWD Conf : Confirmation confirmé - les mots de passe concordent"
			
		}
		Else { $ico_conf_ok.Kind = "Close"
				$ico_conf_ok.Foreground= "Red"
				$ico_conf_ok.visibility = "Visible"
				$ico_conf_ok.ToolTip = "Les mots de passe saisies ne correspondent pas !"
				$lbl_conf_ko.Content = "Les mots de passe ne correspondent pas."
				$lbl_conf_ko.Foreground= "Red"
				$lbl_conf_ko.Visibility = "Visible"
				Log -Add "PWD Conf : Erreur de saisie => Les mots de passe ne correspondent pas"
			}
	}

})

$btn_ok.add_click({

	[String]$old = $pwd_old.get_password()
	[String]$New = $pwd_new.get_password()
	[String]$Confirm = $pwd_conf.get_password()
	if (($old -eq $null) -or ($old -eq "") -or ($new -eq $null) -or ($new -eq "") -or ($Confirm -eq $null) -or ($Confirm -eq "") ) {
		$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
		$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
		
		$Title_Alerte.Content = "Erreur de saisie"
		$Content_Alerte.Content = "Tous les champs doivent être renseignés !"
		$Content_Alerte_2.Content = "Merci de réessayer."
		$Content_Alerte_3.Content = ""

		$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
		Log -Add "Btn OK : Erreur de saisie => Un ou plusieurs champs sont vides"

	}
	Else {
			if (( Check_Old -old_pwd $old ) -eq $false) { 
				$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
				$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
				
				$Title_Alerte.Content = "Erreur de saisie"
				$Content_Alerte.Content = "Ancien mot de passe renseigné erroné !"
				$Content_Alerte_2.Content = "Merci de réessayer."
				$Content_Alerte_3.Content = ""

				$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
				Log -Add "Btn OK : Erreur de saisie => Ancien mot de passe renseigne errone"
			} 
			Else {
				If ($New -cne $Confirm) {
					$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
					$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
			
					$Title_Alerte.Content = "Erreur de saisie"
					$Content_Alerte.Content = "Les mots de passe ne correspondent pas."
					$Content_Alerte_2.Content = "Merci de réessayer."
					$Content_Alerte_3.Content = ""

					$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
					Log -Add "Btn OK : Erreur de saisie => Les mots de passe ne correspondent pas"
				}
				Else {
					Try {
						$error.clear()
						Log -Add "Btn OK : Changement du mot de passe - Set-ADAccountPassword "
						#Get-ADUser MBG17590
						Set-ADAccountPassword $env:USERNAME -OldPassword (ConvertTo-SecureString -AsPlainText $Old -Force) -NewPassword  (ConvertTo-SecureString -AsPlainText $New -Force) -PassThru -ErrorVariable ErrorVar
						$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
						$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
						
						$BtnClose.Visibility = "Hidden"
						$BtnClose_All.visibility = "Visible"
						$Title_Alerte.Content = "Succès changement de mot de passe"
						$Content_Alerte.Content = "Le mot de passe a été changé. La session va être fermée."
						$Content_Alerte_2.Content = "Merci de vous reloguer avec votre nouveau mot de passe."
						$Content_Alerte_3.Content = ""
						Log -Add "Btn OK : Succès changement de mot de passe"
	
						$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
	
						Log -Add "Btn OK : Fermeture de session"

					}catch{
						$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
						$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
				
						$Title_Alerte.Content = "Erreur changement de mot de passe"
						#$Content_Alerte.Content = "$($Error[0])"
						$Content_Alerte.Content = "Le mot de passe ne répond pas aux spécifications de "
						$Content_Alerte_2.Content = "longueur, de complexité ou d’historique du domaine."
						$Content_Alerte_3.Content = "Merci de réessayer"

						Log -Add "Btn OK : Erreur changement de mot de passe"
						Log -Add "$($Error[0])"
						$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
					} #catch
					
				} #Else Changement
			} # Else mdp_old OK
	} # Else NULL
})

$BtnClose.add_click({ 
	$CustomDialog.RequestCloseAsync() 

	if (($Connexion.IPAddress.count -ne 1) -and (($Connexion.IPAddress -match "10.227.*") -or ($Connexion.IPAddress -match "10.208.*") -or ($Connexion.IPAddress -match "10.210.*"))) {
	Log -Add "Poste en TRAMO : Fermeture de la fenetre"
	[System.Windows.Forms.Application]::Exit(); Stop-Process $pid
	}
	
})

$BtnClose_All.add_click({		
	Log -Add "Btn OK : Fermeture de session"	
	Logoff
	
	$CustomDialog.RequestCloseAsync()
	[System.Windows.Forms.Application]::Exit(); Stop-Process $pid

})

$btn_help.add_click({
	Log -Add "Btn Help : Ouverture de l'aide"
	$flyout.isopen = $true	
})

$btn_cancel.add_click({ 
	Log -Add "Btn Cancel : Fermeture de la fenêtre"	
	Snooze
})

$btn_close.add_Click({
	 $Users_Window.Hide()
	 [System.Windows.Forms.Application]::Exit(); Stop-Process $pid
})


$btn_report.add_click({ 
	Log -Add "Snooze défini par utilisateur"

	switch -Wildcard ($cbx_snooze.SelectedIndex){
		0 { [int]$time_to_sleep = 900 }
		1 { [int]$time_to_sleep = 1800 }
		2 { [int]$time_to_sleep = 3600 }
		3 { [int]$time_to_sleep = 14400 }
		4 { [int]$time_to_sleep = 28800 }
	}
	Log -Add "Rappel de la notification dans $($time_to_sleep) secondes"
	$Restart = "Snooze"

	start-process -WindowStyle hidden powershell.exe ".\Popupmdp_notif.ps1 -restart '$Restart' -time_to_sleep '$time_to_sleep'"  
	$Users_Window.Close()
	Stop-Process $pid
	 
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
	 {
	  $timer.Stop() 
	 }  
 })

 $btn_change.add_Click({ 

	Log -Add "Ouverture Windows_Change"
    $Change_Window.WindowStartupLocation = "CenterScreen"
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Change_Window)
	
	$Users_Window.Hide()

	$Change_Window.Show()
	$Change_Window.Activate()
	
	Log -Add "Verification si TELETRAVAIL/TRAMO"
	$Connexion = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null } 
	if (($Connexion.IPAddress.count -ne 1) -and (($Connexion.IPAddress -match "10.227.*") -or ($Connexion.IPAddress -match "10.208.*") -or ($Connexion.IPAddress -match "10.210.*"))) {
		Log -Add "TELETRAVAIL/TRAMO - Affichage message avec procedure changement mot de passe"
		$settings = [MahApps.Metro.Controls.Dialogs.MetroDialogSettings]::new()
		$settings.ColorScheme = [MahApps.Metro.Controls.Dialogs.MetroDialogColorScheme]::Theme
		
		$BtnClose.Visibility = "Visible"
		$BtnClose_All.visibility = "Hidden"
		$Title_Alerte.Content = "INFORMATION TELETRAVAIL/TRAMO"
		$Content_Alerte.Content = "Pour changer votre mot de passe a domicile, merci de"
		$Content_Alerte_2.Content = "suivre la procédure décrite dans le lien suivant :"
		
		$Content_Alerte_3.Content = "Lien Procedure Guid Ouest"
		$Content_Alerte_3.Foreground = "CornflowerBlue"
		$Content_Alerte_3.Add_MouseLeftButtonUp({
			[system.Diagnostics.process]::start('http://')
			Log -Add "Poste en TRAMO : Fermeture de la fenetre"
			[System.Windows.Forms.Application]::Exit(); Stop-Process $pid
		})
	
		$result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($Change_Window,$CustomDialog, $settings)
	
	}
})

# ----------------------------------------------------
# MENU SYSTRAY
# ----------------------------------------------------		
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "Votre mot de passe va expirer"
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true

$Menu_Users = New-Object System.Windows.Forms.MenuItem
$Menu_Users.Text = "Afficher la notification"

$Menu_Change = New-Object System.Windows.Forms.MenuItem
$Menu_Change.Text = "Modifier votre mot de passe"

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"


$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Users)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Change)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)


# ---------------------------------------------------------------------
# ACTION SYSTRAY
# ---------------------------------------------------------------------
$Main_Tool_Icon.Add_Click({

        $Users_Window.WindowStartupLocation = "Manual"	
        $Users_Window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Users_Window.Width)
        $Users_Window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Users_Window.Height)	
    
		$Users_Window.Show()
		$Users_Window.Activate()	
			
})

# ---------------------------------------------------------------------
# AFFICHAGE/DESAFFICHAGE FENETRE
# ---------------------------------------------------------------------
$Menu_Users.Add_Click({
    $Users_Window.WindowStartupLocation = "Manual"
    $Users_Window.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Users_Window.Width)
    $Users_Window.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Users_Window.Height)	
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Users_Window)
	$Users_Window.Show()
	$Users_Window.Activate()	
	
	
})

$Menu_Change.Add_Click({
	Log -Add "Ouverture Windows_Change"
    $Change_Window.WindowStartupLocation = "CenterScreen"
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Change_Window)
	
	$Users_Window.Hide()
    $Change_Window.Show()
    $Change_Window.Activate()
 })


$Menu_Exit.add_Click({
	$Main_Tool_Icon.Visible = $false
	$Users_Window.Close()
	Stop-Process $pid
 })





Log -Add "Lancement Fonction Test_IP"
Test_IP
