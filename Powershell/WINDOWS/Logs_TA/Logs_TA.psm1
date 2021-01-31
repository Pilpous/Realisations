### MODULE LOG TA ###

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
