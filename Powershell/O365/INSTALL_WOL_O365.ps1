<#
.SYNOPSIS
Installation OFFICE 365

.DESCRIPTION
Installation de OFFICE 365 : 
Script déclenché par tâche planifiée WOL_Popup_O365
Session fermée et tranche horaire définie : si les conditions sont réunies, lancement de l'installation - aucune action utilisateur
Si conditions non réunies : activation tache planifiée USER_Popup_O365 pour validation par l'utilisateur

Si installation complétée : suppression des 2 tâches planifiées WOL_Popup_O365 et USER_Popup_O365


.NOTES
Version : 1
Date de création : 17/06/2020
Compatibilité : Powershell V5
#>

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
    
$path_logs = "C:\AppDSI\EXPGLB\OFFICE365\O365_GLB"

Create_Log -Path_Log $path_logs -Appli Wol_install_O365
Log -Add "#################################################################"
Log -Add "###         Lancement du script"
Log -Add "#################################################################"

#Verification version Office
$Office_version_key = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Common\ProductVersion).LastProduct
$script:source_path = 'C:\AppDSI\EXPGLB\OFFICE365\1908\','C:\ProgramData\RepoOfficeO365V1908X64\setup.exe'

if (($Office_version_key -ne '16.0.11929.20562' ) -and ((Test-Path $source_path[0]) -eq "True") -and ((Test-Path $source_path[1]) -eq "True")){
    Log -Add "Sources presentes sur le poste"
    Log -Add "Verification tranche horaire et etat session :"
    $user_connected = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Username
    if ( ($null -eq $user_connected.username) -and  ((Get-Date) -ge (Get-Date -Hour 7 -Minute 0 ))) { 
        Log -Add "Session fermée et tranche horaire invalide"
        Log -Add "Tâche Planifiée SESSION : passage en actif"
        # Lancement du script Popup_O365 // Tâche Planifiée SESSION : passage en actif
        Get-ScheduledTask -TaskName Task_POPUP_Install_O365  | Enable-ScheduledTask
    }
    elseif ( ($null -ne $user_connected.username) -and  ((Get-Date) -le (Get-Date -Hour 7 -Minute 0 ))) {
        Log -Add "Session ouverte et tranche horaire valide"
        Log -Add "Tâche Planifiée SESSION : passage en actif"
        #Lancement du script Popup_O365 // Tâche Planifiée SESSION : passage en actif
        Get-ScheduledTask -TaskName Task_POPUP_Install_O365  | Enable-ScheduledTask
    }
    elseif ( ($null -ne $user_connected.username) -and  ((Get-Date) -ge (Get-Date -Hour 7 -Minute 0 ))) {
        Log -Add "Session ouverte et tranche horaire invalide"
        Log -Add "Tâche Planifiée SESSION : passage en actif"
        #Lancement du script Popup_O365 // Tâche Planifiée SESSION : passage en actif
        Get-ScheduledTask -TaskName Task_POPUP_Install_O365  | Enable-ScheduledTask
    }
    elseif ( ($null -eq $user_connected.username) -and  ((Get-Date) -le (Get-Date -Hour 7 -Minute 0 )) ) { 
        Log -Add "Session fermée et tranche horaire valide" 
        try {
            Log -Add "Lancement installation O365 sans confirmation utilisateur"
            # Lancement installation  O365
            #Log -Add "Désactivation de la tâche Planifiée Task_POPUP_Install_O365 "
            #Get-ScheduledTask -TaskName Task_POPUP_Install_O365  | Disable-ScheduledTask


            Log -Add "Recuperation info Appli O365 sur SCCM - get-ciminstance"
            $Application = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"})
            Log -Add "Creation Argument pour installation"
            $SCCM_Arg = @{EnforcePreference = [UINT32] 0
                Id = "$($Application.id)"
                IsMachineTarget = $Application.IsMachineTarget
                IsRebootIfNeeded = $False
                Priority = 'High'
                Revision = "$($Application.Revision)" }

            Log -Add "Lancement Installation via Invoke-cimmethode"
            Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Application -MethodName Install -Arguments $SCCM_Arg

            Log -Add "Lancement boucle"
            Do {
                Log -Add "tempo 5 secondes"
                Start-Sleep -Seconds 5
                $check_install = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" -ErrorAction SilentlyContinue | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"}).InProgressActions
            } until ( $null -eq $check_install[0] )

              Log -Add "Installation terminee"


            Log -Add "Verification Installation"
            $Check_Result = (Get-CimInstance -ClassName CCM_Application -Namespace "root\ccm\clientSDK" | Where-Object {$_.Name -like "OFFICE365_EXP_OFFICESTDO365V1908X64_10001"}).InstallState
            
            if ($Check_Result -eq "Installed") {
                Log -Add "Statut : $($Check_Result)"
                Log -Add "Suppression des Taches Planifiees O365"
                #Suppression des Tâches Planifiées O365
                Get-ScheduledTask -TaskName Task_WOL_Install_O365 | Unregister-ScheduledTask -Confirm:$false
                Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Unregister-ScheduledTask -Confirm:$false
            }
            else {		
                Log -Add "Statut Installation : $($Check_Result)"
                # A DEFINIR
    
                #Modification des tâches planifiés
                Log -Add "Activation de la tache planifiee Task_POPUP_Install_O365"
                Get-ScheduledTask -TaskName Task_POPUP_Install_O365 | Enable-ScheduledTask
            }

            Log -Add "Fin du script"
        } catch {
            Log -Add "Installation en erreur :"
            Log -Add "$($Error[0])"
                }
    }
} else { 
    if ($Office_version_key -eq '16.0.11929.20562' ) { Log -Add "Version OFFICE 365 16.0.11929.20562 deja installee sur le poste"}
    elseif (!(Test-Path $source_path[1]) -or !(Test-Path $source_path[0])) { Log -Add "Sources non presentes sur le poste"}
    Log -Add "Pas de lancement du script"
}