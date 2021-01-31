### Lancement/Arrêt SSOX ###

#region VARIABLES
$scriptRoot = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($scriptRoot -eq $PSHOME.TrimEnd('\') -or $scriptRoot -eq 'C:\Program Files (x86)\PowerGUI')
{
    $scriptRoot = $PSScriptRoot
}
$check_watcher = Get-Process -Name watcher -ErrorAction SilentlyContinue
$check_wait_watcher = Get-Process -Name wait_watcher -ErrorAction SilentlyContinue
#endregion
Import-Module Logs_TA
Create_Log -Path_Log $scriptRoot -Appli AR_SSOX

If ((!$check_wait_watcher) -and (!$check_watcher)) 
		{Log -Add "Watcher et Wait_Watcher non lancé - Lancement !" 
		start 'C:\AppDSI\EXPGLB\SSOX\GLB\Applications\wait_watcher.exe'}

Else { 	Log -Add "Watcher et Wait_Watcher lancés - Arrêt des 2 processus !"
		$check_wait_watcher | Stop-Process -ErrorAction SilentlyContinue
		Start-Sleep -Seconds 1
		$check_watcher | Stop-Process -ErrorAction SilentlyContinue }
		
		


