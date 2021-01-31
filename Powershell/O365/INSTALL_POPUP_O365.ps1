<#
.SYNOPSIS
Popup Installation OFFICE 365 - A Valider par Utilisateur

.DESCRIPTION
Affichage Popup pour permettre à l'utilisateur de lancer l'installation de OFFICE 365
Possibilités : INSTALLATION // REPORT // FERMER

Script lancé par Tâche planifiée à l'ouverture de session - Activée par le script INSTALL_WOL_O365.ps1

.NOTES
Version : 1
Date de création : 17/06/2020
Compatibilité : Powershell V5
#>



#
#	INIT
#-------------------------------------------------------------------------
Param
 (
 [String]$Restart,
 [int]$time_to_sleep = 3600
 
 )
 
If ($Restart -ne "") 
 {
  Start-Sleep $time_to_sleep
 }


 # VARIABLES // ASSEMBLIES
 #-------------------------------------------------------------------------
$Global:Current_Folder = "C:\AppDSI\EXPGLB\OFFICE365\O365_GLB"

Set-Location $Current_Folder
$Assembly_Folder = Join-Path -Path $Current_Folder -ChildPath .\assembly

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null
[System.Windows.Forms.Application]::EnableVisualStyles()
If(Test-Path $Assembly_Folder) {
	Foreach ($Assembly in (Get-ChildItem $Assembly_Folder -Filter *.dll)) {
		$null = [System.Reflection.Assembly]::LoadFrom($Assembly.fullName)
	}
}
#-------------------------------------------------------------------------

$script:source_path = 'C:\AppDSI\EXPGLB\OFFICE365\1908\','C:\ProgramData\RepoOfficeO365V1908X64\setup.exe'
#-------------------------------------------------------------------------
# Chargement modules
Function Create_Log {
<#
.Synopsis
Création du fichier de logs

.Description
Création du dossier logs dans le chemin fourni en paramètre.
Création du fichier de logs avec nom de l'appli et date du jour dans le dossier logs.

.Parameter Path_Logs
Le chemin où doit être stocké l'appli - (C:\APPDSI\EXPGLB...)

.Parameter Appli
Nom de l'application choisi

.Example
Create_logs -Path_log C:\APPDSI\EXPGLB\REP_IE -Appli Reparation_IE
=> Création de C:\APPDSI\EXPGLB\REP_IE\logs\20190122_Reparation_IE.txt

#>
	[CmdletBinding()]
	Param (	[Parameter(Position=0,Mandatory=$true,HelpMessage='Renseigner le chemin où doivent être stockés les logs')]
			[string]$Path_Log ,
			[Parameter(Position=1,Mandatory=$true,HelpMessage='Renseigner le nom de l application')]
			[string]$Appli )
			
$Date_day = Get-Date -Format yyyyMMdd

$script:directory_log = "$($Path_Log)\logs\"
$file_log = "$($Date_day)_$($Appli).txt"
$script:Error_log = $path_log+"\error_log.txt"
$script:path_to_log = $directory_log+$file_log


try {
if((Test-Path $path_to_log  ) -ne $true){New-Item $path_to_log  -Force -ItemType File
										Add-Content $path_to_log  -Value "#############################################################################"
										Add-Content $path_to_log -Value "                        LOGS $($Appli.toUpper())"
										Add-Content $path_to_log  -Value "        Poste $env:ComputerName - Utilisateur $env:UserName - Date $Date_day"
										Add-Content $path_to_log  -Value "#############################################################################"
										}
if ((Test-Path $Error_log) -ne $true) { New-Item $Error_log -Force -ItemType File
										Add-Content $Error_log  -Value "#############################################################################"
										Add-Content $Error_log -Value "                        ERRORS $($Appli.toUpper())"
										Add-Content $Error_log  -Value "        Poste $env:ComputerName - Utilisateur $env:UserName - Date $Date_day"
										Add-Content $Error_log  -Value "#############################################################################"}
										
} catch {Add-Content $Error_log -value $Error[0] }

}

Function Log {
<#
.Synopsis
Génération des logs

.Description
Ajout de ligne dans le fichier de log créer via la fonction Create_log 

.Parameter Add
Renseigner le contenu voulu à ajouter au fichier de log

.Example
Log -Add "Arrêt du processus iexplore.exe"
=> Ajout de la ligne 20190122_10:14:58 Arrêt du processus iexplore.exe dans le fichier de log
#>
	[CmdletBinding()]
	Param ([Parameter(Mandatory=$true,HelpMessage='Renseigner le contenu à ajouter aux logs')][string]$Add)
	
$Date = Get-Date -Format yyyyMMdd_hh:mm:ss	
try { Add-Content $path_to_log -Value "[$($Date)] $($Add)" }
catch { Add-Content $Error_log -value "$Date - Erreur : $Error[0]" }
}

#-------------------------------------------------------------------------
# Creation des logs
Create_Log -Path_Log $Current_Folder -Appli Install_O365
Log -Add "#################################################################"
Log -Add "###         Lancement du script"
Log -Add "#################################################################"
Log -Add "Chargement des composants"
#-------------------------------------------------------------------------



#	SPLASH SCREEN
#-------------------------------------------------------------------------
function Start_SplashScreen {
	$Pwshell.Runspace = $runspace
	$script:handle = $Pwshell.BeginInvoke() 
}

function Close_SplashScreen {
	$hash.window.Dispatcher.Invoke("Normal",[action]{ $hash.window.close() })
	$null = $Pwshell.EndInvoke($handle)
	#$null = $runspace.Close()
}

$hash = [Hashtable]::Synchronized(@{})
$runspace = [Runspacefactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("hash",$hash) 
$Pwshell = [PowerShell]::Create()	

$null = $Pwshell.AddScript({
$xml = [xml]@"
<Window
xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
x:Name="WindowSplash" Title="SplashScreen" WindowStyle="None" WindowStartupLocation="CenterScreen"
AllowsTransparency="True" ShowInTaskbar ="False" Width="600" Height="300" ResizeMode ="NoResize" Topmost="True">
   <Grid Background="#FFEA3E00">
		<StackPanel Orientation="Vertical" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="1" Background="White" Width="599" Height="299" >   
			<Image Source="C:\AppDSI\EXPGLB\OFFICE365\O365_GLB\Resources\O365.png" Width="350" Margin="0,20,0,0"/>
			<Controls:ProgressRing IsActive="{Binding IsActive}" Foreground="#FFEA3E00" HorizontalAlignment="Center"  Margin = "0,0,0,30" Width="40" Height="8"/>
			<Label Name="LoadingLabel" Content="Installation en cours. Merci de patienter"  Foreground="#FFEA3E00" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="24" Margin = "0,0,0,0"/>
			<Label Content="Merci de ne pas utiliser votre poste pendant le temps de l'installation."  Foreground="#FFEA3E00" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14" Margin = "0,0,0,0"/>
		</StackPanel>
   </Grid>
</Window> 
"@
$reader = New-Object System.Xml.XmlNodeReader $xml
$hash.window = [Windows.Markup.XamlReader]::Load($reader) 
$hash.window.ShowDialog() 
})


#	XAML
#-------------------------------------------------------------------------
Function LoadXaml($Global:filename) {
	$xamlLoader = (New-Object System.Xml.XmlDocument)
	$xamlLoader.Load($filename)
	$xamlLoader
}


# Chargement xaml des fenetres
#--------------------------------------------------------------------------

# Fenetre Popup
$notif_xaml = LoadXaml("$Current_Folder\Notif.xaml") 
$notif_xamlReader = (New-Object System.Xml.XmlNodeReader $notif_xaml)
$Window_Notif = [Windows.Markup.XamlReader]::Load($notif_xamlReader)

#Fenetre Fin
$validation_xaml = LoadXaml("$Current_Folder\Window_validation.xaml") 
$validation_xamlReader = (New-Object System.Xml.XmlNodeReader $validation_xaml)
$Window_Valid = [Windows.Markup.XamlReader]::Load($validation_xamlReader)

$btn_report = $Window_Notif.FindName("btn_report")
$btn_close = $Window_Notif.FindName("btn_close")
$btn_change = $Window_Notif.FindName("btn_change")
$cbx_snooze = $Window_Notif.FindName("cbx_snooze")
$lbl_name = $Window_Notif.FindName("lbl_name")

$btn_ok = $Window_Valid.FindName("btn_ok")
$lbl_state_install = $Window_Valid.FindName("lbl_state_install")
$lbl_consigne = $Window_Valid.FindName("lbl_consigne")


#	EVENEMENTS
#-------------------------------------------------------------------------

# BOUTON INSTALLER
$btn_change.add_Click({
	Log -Add "Lancement installation par utilisateur"
	Log -Add "Masquage fenetre Window_Notif"
	$Window_Notif.Hide()
	
	Log -Add "Lancement SplashScreen"
	Start_SplashScreen

	try {
		$Application = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"})
		$SCCM_Arg = @{EnforcePreference = [UINT32] 0
			Id = "$($Application.id)"
			IsMachineTarget = $Application.IsMachineTarget
			IsRebootIfNeeded = $False
			Priority = 'High'
			Revision = "$($Application.Revision)" }

		Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Application -MethodName Install -Arguments $SCCM_Arg

		$count = 0
		
		Do {
			$check_install = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"})
		
			if (($null -eq $check_install.InProgressActions[0]) -and ($check_install.installState -eq "Installed"))
				{ $check = 0 }
			elseif(($null -eq $check_install.InProgressActions[0]) -and ($check_install.installState -eq "NotInstalled"))
				{ 	Log -Add "InProgressActions a null et InstallState a NotInstalled - mise en attente"
					Log -Add "Compteur = $($count)"
					start-sleep -Seconds 150
					$count++ }
			
			Else { $check = 1 }

		} until (($check -eq 0) -or ($count -eq 2))


		Log -Add "Installation terminee"
		
		
		### VERIFICATION INSTALL ###
		Log -Add "Verification Installation"
		$Check_Result = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"}).InstallState
		
		if ($Check_Result -eq "Installed") {
			Log -Add "Statut : $($Check_Result)"
			Log -Add "Suppression des Tâches Planifiées O365"
			#Suppression des Tâches Planifiées O365
			Get-ScheduledTask -TaskName Task_WOL_Install_O365 | Unregister-ScheduledTask -Confirm:$false
			Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Unregister-ScheduledTask -Confirm:$false
		}
		else {		
			Log -Add "Statut : $($Check_Result)"
			#Modification des tâches planifiés
			Log -Add "Desactivation de la tache planifiee Task_POPUP_Install_O365"
			Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Disable-ScheduledTask
		}


		Log -Add "Fermeture SplashScreen"
		Close_SplashScreen
		
		Start-Sleep -Seconds 1
		
		Log -Add "Affichage de la fenetre Window_Valid"
		$Window_Valid.Show()
		$Window_Valid.Activate()

	} catch {
		Log -Add "Une erreur s'est produite"
		Log -Add "$($Error[0])"
		
		Log -Add "Fermeture SplashScreen"
		Close_SplashScreen

		Start-Sleep -Seconds 1
		
		Log -Add "Affichage de la fenetre Window_Valid"
		$lbl_state_install.Content = "Une erreur s'est produite pendant la mise à jour..."
		$lbl_consigne.Content = "Merci de contacter votre support informatique."
		$Window_Valid.Show()
		$Window_Valid.Activate()

		Log -Add "Desactivation de la tache planifiee Task_POPUP_Install_O365"
		Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Disable-ScheduledTask
	}

})

#BOUTON SNOOZE
$btn_report.add_Click({
	Log -Add "Snooze defini par utilisateur"
	Log -Add "Desactivation de la tache planifiee Task_POPUP_Install_O365"
	Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Disable-ScheduledTask

	$recup_time = $cbx_snooze.SelectedIndex
	switch -Wildcard ($recup_time){
		0 { $time_to_sleep = 3600 }
		1 { $time_to_sleep = 7200 }
		2 { $time_to_sleep = 14400 }
		3 { $time_to_sleep = 28800 }
	}
	Log -Add "Rappel de l'install dans $($time_to_sleep) secondes"

	$Restart = "Snooze"
	start-process -WindowStyle hidden powershell.exe ".\INSTALL_POPUP_O365.ps1 -restart '$Restart' -time_to_sleep '$time_to_sleep'"  
	$Window_Notif.Close()
	Stop-Process $pid
	 
	$Global:Timer_Status = $timer.Enabled
	If ($Timer_Status -eq $true)
	 {
	  $timer.Stop() 
	 }  

})

#BOUTON FERMER
$btn_close.add_Click({ 
	Log -Add "Fenetre Window_Notif fermee par utilisateur"
	$Window_Notif.Close()
	Log -Add "Desactivation de la tache planifiee Task_POPUP_Install_O365"
	Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Disable-ScheduledTask
	Log -Add "Arret du programme"
	Stop-Process $pid
})

#BOUTON OK - FIN INSTALLATION
$btn_ok.add_Click({
	Log -Add "Fenetre Window_Valid fermee par utilisateur"
	$Window_Valid.Close()
	Log -Add "Arret du programme"
	Stop-Process $pid
})

#PRECHARGEMENT DE LA POPUP
$Window_Notif.Add_Loaded({
	#Recuperation Nom d'utilisateur
	$user_connected = Get-CimInstance -classname Win32_ComputerSystem -Property UserName | Select-Object Username -ExpandProperty Username
	$User = Get-ADUser $user_connected.Split('\')[1] | Select-Object GivenName,Surname
	$prenom = $user.GivenName[0]+($user.GivenName.Split($user.GivenName[0])[1]).toLower()

	$lbl_name.Content = "Bonjour $($prenom) $($user.Surname)"
	Log -Add "$($prenom) $($user.Surname)"
})

###########################################################################################
###########################################################################################
###########################################################################################

# LANCEMENT DU SCRIPT

$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

$Window_Notif.Left = $([System.Windows.SystemParameters]::WorkArea.Width-$Window_Notif.Width)
$Window_Notif.Top = $([System.Windows.SystemParameters]::WorkArea.Height-$Window_Notif.Height)	

Log -Add "Affichage de la fenetre Window_Notif"
$Window_Notif.Show()
$Window_Notif.Activate()	

$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)