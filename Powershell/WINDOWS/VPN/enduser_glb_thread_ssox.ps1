<#
.SYNOPSIS
Lancement de EndUser avec arrêt et relance des applis

.DESCRIPTION
Arret de Pulse.exe SSOX, IE, Outlook et Skype
Lancement de EndUser
Relance de SSOX, IE, Outlook et Skype

.NOTES
Version : 1
Date de création : 22/04/2020
Compatibilité : Powershell V5
#>
Import-Module Logs_TA
$path_logs ='C:\AppDsi\EXPGLB\PULSE\END_USER'
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null

[System.Windows.Forms.Application]::EnableVisualStyles()
$PSDefaultParameterValues = @{ '*:Encoding' = 'utf8' }

Create_Log -Path_Log $path_logs -Appli EndUser_GLB
Log -Add "#################################################################"
Log -Add "###         Lancement du script"
Log -Add "#################################################################"


$xml = [xml]@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
Title="Connexion EndUser" Height="229.744" Width="600"  WindowStartupLocation="CenterScreen" Topmost="True" ResizeMode="NoResize" >
<Grid>
<StackPanel HorizontalAlignment="Center" Margin="10">
    <Label Content="Veuillez patienter 2 - 3 minutes :"/>
    <Separator Height="5" Background="Transparent"/>
    <Label Content="- les applications SSOX, Phare Ouest, Outlook, Skype et SIGMA vont se connecter automatiquement"/>
    <Separator Height="5" Background="Transparent"/>
    <Label Content="- vous pourrez ensuite lancer les autres applications : IMPORTANT, merci de laisser SSOX remplir vos"/>
    <Label Content=" identifiants et mots de passe" Margin="0,-10,0,0"/>
    <Separator Height="10" Background="Transparent"/>
    <Label Content="Ce message va se fermer automatiquement au bout de 1 - 2 minutes"/>
</StackPanel>
</Grid>
</Window>
"@

Try {
    Log -Add "Arret des process"
    $process = @("wait_watcher","watcher","iexplore","outlook","lync","pulse")
    foreach ($proc in $process) {
       Get-Process $proc -ErrorAction SilentlyContinue | Stop-Process
       Log -Add "Arret $proc"
    }
    Log -Add "Arret des process termine"
    $d_day = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'

    Log -Add "Lancement EndUser"
    Start-Job -ScriptBlock { Start-Process 'C:\Program Files (x86)\Common Files\Pulse Secure\JamUI\Pulse.exe'  -ArgumentList '-tray' }
    #Start-Job -ScriptBlock { C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -file "C:\AppDSI\EXPGLB\PULSE\SECURE\Pulse Secure.ps1" }

    
    Log -Add "Surveillance Event Windows pour EndUser" 
    $count_event = 0
    Do {

    $event_pulse = Get-WinEvent -LogName 'Pulse Secure/Operational' -MaxEvents 10 | Where-Object {(Get-Date $_.TimeCreated -Format 'dd/MM/yyyy HH:mm:ss') -ge $d_day }
    #$event_pulse = Get-WinEvent -LogName 'Junos Pulse/Operational' -MaxEvents 10 | Where-Object {(Get-Date $_.TimeCreated -Format 'dd/MM/yyyy HH:mm:ss') -ge $d_day }
    Start-Sleep -Seconds 1
    $count_event ++

    } until (($event_pulse.ID -match 312) -or ($count_event -eq 3600 ) )

    If ($count_event -eq 3600) { Log -Add "Attente de 1h : arret du script" }
    Else {
        Log -Add "Event Connexion EndUser OK"

        Log -Add "Lancement des Applis"
        Log -Add "Lancement SSOX"

        Start-Process C:\AppDsi\EXPGLB\SSOX\GLB\Applications\wait_watcher.exe 
        Log -Add "En attente de SSOX"
        
        $xamlReader = (New-Object System.Xml.XmlNodeReader $xml)
        $Popup =  [Windows.Markup.XamlReader]::Load($xamlReader)
        $Popup.Show()

        Do {
            $check_ssox = Get-Process watcher
            Start-Sleep -Seconds 2
            #Write-Host $check_ssox.Threads.Count
        }
        while ($check_ssox.Threads.Count -le 11) 
            
        
        #Start-Sleep -Seconds 90

        $Popup.Close()

        Log -Add "Lancement PhareOuest"
        Start-Sleep -Seconds 2
        Start-Process "C:\Program Files (x86)\Internet Explorer\iexplore.exe" http://phareouest.glb.intra.groupama.fr/
        

        Log -Add "Detection Version Office"
        $office = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer\
        if ($office.'(default)' -eq 'Word.Application.16') { 
            Log -Add "Lancement Skype"
            Start-Process "C:\Program Files\Microsoft Office\root\Office16\lync.exe"
            Log -Add "Lancement Outlook"
            Start-Process "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        }
        else {
            Log -Add "Lancement Skype"
            Start-Process "C:\Program Files (x86)\Microsoft Office\Office16\lync.exe"
            Log -Add "Lancement Outlook"
            Start-Process "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE" 
        }

        Start-Sleep -Seconds 2
        Log -Add "Lancement Sigma"
        Start-Process C:\Applics\TN3270\TN3270.exe -ArgumentList '/Session GRC3270'
            if ( !(Get-Process TN3270 -ErrorAction SilentlyContinue)) { Log -Add "Sigma non lance - relance"
                Start-Process C:\Applics\TN3270\TN3270.exe -ArgumentList '/Session GRC3270'
                Start-Sleep -Seconds 10
                Log -Add "Sigma - mise en attente terminee"
                Get-Process TN3270 | Stop-Process
                Log -Add "Arret du process TN3270"  }
            Else { Log -Add "Sigma en execution - mise en attente pour login"
                Start-Sleep -Seconds 10
                Log -Add "Sigma - mise en attente terminee"
                Get-Process TN3270 | Stop-Process
                Log -Add "Arret du process TN3270" }
            
        Start-Process C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe 'C:\AppDsi\DEVGLB\mbg\projets_perso\login_session.ps1'        
        Log -Add "Fin du script"
    } # ELSE
} # TRY
Catch { Log -Add $Error[0] }
