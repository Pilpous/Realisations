##### Script SSOX #####
## Surveillance process watcher.exe + relance SSOC si KO
## Arrêt SSOX
## Démarrage SSOX
$scriptRoot = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($scriptRoot -eq $PSHOME.TrimEnd('\') -or $scriptRoot -eq 'C:\Program Files (x86)\PowerGUI')
{
    $scriptRoot = $PSScriptRoot
}
Import-Module Logs_TA
Create_Log -Path_Log $scriptRoot -Appli WAIT_WATCHER

Function LaunchWatcher {
Log -add "Lancement de la fonction LaunchWatcher"
$check_process = Get-Process -Name watcher -ErrorAction SilentlyContinue
if (!$check_process) { 	start "$ssox_path\watcher.exe"
						Log -Add "Lancement du watcher.exe"
						WaitWatcher
						Log -Add "Lancement de la fonction WaitWatcher" }
else {	Log -add "Watcher.exe déjà en exécution - Lancement de la fonction WaitWatcher" 
		WaitWatcher}
}

Function WaitWatcher {
$check_process = Get-Process -Name watcher -ErrorAction SilentlyContinue
if (!$check_process) { 	Log -Add "Processus watcher.exe n'est pas en exécution"
						LaunchWatcher
						Log -add "Lancement de la fonction LaunchWatcher"}
else {	Log -Add "Watcher.exe en exécution - surveillance du processus"
		Wait-Process -Name watcher
		LaunchWatcher
}
}

Function Poste {
If ([IntPtr]::Size -eq 8){	$ssox_path = 'C:\Program Files (x86)\Avencis\SSOX'
							Log -add "Poste 64 bits - SSOX : $($ssox_path)" }
Else {	$ssox_path = 'C:\Program Files\Avencis\SSOX'
		Log -add "Poste 32 bits - SSOX : $($ssox_path)"}

	
LaunchWatcher
}

Poste