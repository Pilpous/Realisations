<#
.SYNOPSIS
BRIADAEMON - Surveillance BRIA / GPHONE

.DESCRIPTION
Surveillance des process BriaStretto.exe, BriaEnterprise.exe et GenesysSoftphone.exe en lien avec les process Genesys: Genesys.exe et javaw.exe
Si SOFTPHONE lancÃ© => Rien
Si GPHONE lancÃ© seul => message d'alerte : merci de lancer votre softphone
Si ArrÃªt SOFTPHONE alors que GPHONE lancÃ© => message d'alerte : merci de lancer votre softphone

Passage des processus BriaStretto.exe, BriaEnterprise.exe et GenesysSoftphone.exe en prioritÃ© HAUTE

.NOTES
Version : 2.0
Date de crÃ©ation : 18/01/2020
        Ajout de la surveillance des process BriaEnterprise.exe et GenesysSoftphone.exe
        Ajout de la surveillance du process Genesys.exe
        Passage des processus Genesys.exe ou javaw.exe(Gphone)en priorité HAUTE

#>

## DEFINITION VARIABLE ET IMPORT MODULE
$Boucle_infinie = 0 
$Current_Folder = "C:\Temp\BriaDaemon\"
Import-Module Logs_TA

Create_Log -Path_Log $Current_Folder -Appli BriaDaemon
Log -Add "#########################################################"
Log -Add "###         Lancement du script"
Log -Add "#########################################################"

Function Get_High {
    Param (
        $process
    )
    ## PASSAGE DU PROCESS SOFTPHONE EN PRIORITE HAUTE
    Foreach ($proc in $process) {
        if ( ($proc.GetType()).Name -eq 'Int32') {
            if ((Get-Process -id $proc).PriorityClass -ne 'High') {
                Get-Process -id $proc | ForEach-Object { $_.ProcessorAffinity=128
                                                    $_.PriorityClass='High' }
                Log -Add "Passage du Process ID $($proc) en priorite HAUTE"

            }
            #else { Log -Add "Process ID $($proc) en priorite HAUTE"
            #}
        }
        else {
            if ((Get-Process $proc).PriorityClass -ne 'High') {
                Get-Process $proc | ForEach-Object { $_.ProcessorAffinity=128
                                                    $_.PriorityClass='High' }
                Log -Add "Passage du Process $($proc) en priorite HAUTE"

            }
            #else { Log -Add "Process $($proc) en priorite HAUTE"
            #}
        }
    }

}

Function Phone_Surv {
    $process = @()
    if ( $sp_process = Get-Process briastretto,briaenterprise,genesyssoftphone -ErrorAction SilentlyContinue ) {
        $process += $sp_process.Name
        }
    else {
        $process =   0
    }
    # PRIORISATION PROCESS HAUTE
    if ($process -ne 0) {
        Get_High -process $process
    }
    return $process
}


Function GPhone_Surv {
    if ($gphone = Get-Process Genesys -ErrorAction SilentlyContinue) {
        $gphone_ID = $gphone.Id
    }
    elseif ($gphone = Get-Process javaw -ErrorAction SilentlyContinue | Where-Object Path -NE "C:\Program Files (x86)\GRCWS21\GRCClient\jre\bin\javaw.exe") {
        $gphone_ID = $gphone.Id
    }
    else {
        $gphone_ID = 0
    }
    # PRIORISATION PROCESS HAUTE
    if ($gphone_ID -ne 0) {
        Get_High -process $gphone_ID
    }
    return $gphone_ID
}

Function Message {
    Param (
        [string]$MessageboxTitle = "Information",
        [string]$Messageboxbody
    )
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::Ok
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning

    [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$messageicon)

    Log -Add "Message a Utilisateur : $($Messageboxbody)"
}


Do {
    ## TANT QUE GPHONE NON LANCE - SURVEILLANCE PROCESS GPHONE
    Log -Add "Boucle sur Surveillance GPHONE"
    $Pproc = Phone_Surv
    Do {
        Start-Sleep -Seconds 5 # Tempo de 5 secondes
        if ($Pproc -eq 0) { # Si Softphone dÃ©tectÃ©, la fonction Phone_Surv ne sera pas exÃ©cutÃ©
            $Pproc = Phone_Surv }
        $GPhone_Surv = GPhone_Surv
    } Until ($GPhone_Surv -ne 0)

    ## SI LANCEMENT GPHONE / E-GPHONE SANS BRIA
    $Pproc = Phone_Surv #VÃ©rification prÃ©sence Softphone / mise Ã  jour
    $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour

    if ($Pproc -eq 0 ){
        Do { 
            # Si plus/pas de Softphone
            $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour
            if ( $Gsurv -ne 0) {
                Get-Process -Id $Gsurv -ErrorAction SilentlyContinue | Stop-Process # ArrÃªt process Gphone
                Log -Add "Lancement du Gphone Hors presence Softphone"
                Message -Messageboxbody "Veuillez lancer votre Softphone Bria !"
                $SoftPhone_Surv = Phone_Surv
                break
            }
            $SoftPhone_Surv = Phone_Surv #VÃ©rification prÃ©sence Softphone / mise Ã  jour
        } While ($SoftPhone_Surv -eq 0) # Tant que absence Softphone
        
    }
    ## SI ARRET BRIA ALORS QUE GPHONE / E-GPHONE PRESENT
    $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour
    $Pproc = Phone_Surv #VÃ©rification prÃ©sence Softphone / mise Ã  jour
    if (($Pproc -ne 0) -and ($Gsurv -ne 0)) { # Si prÃ©sence Softphone et Gphone - Surveillance du process Softphone
        #Wait-Process $Pproc -ErrorAction SilentlyContinue
        Do {
            Start-Sleep -Seconds 5
            $Pproc = Phone_Surv #VÃ©rification prÃ©sence Softphone / mise Ã  jour
            $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour
        } While ($Pproc -ne 0)
        $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour
        if ($Gsurv -ne 0){ # Si prÃ©sence GPHONE sans Softphone - ArrÃªt du GPHONE
            Get-Process -Id $Gsurv -ErrorAction SilentlyContinue | Stop-Process
            Log -Add "Arret du Softphone alors que le Gphone est ouvert - Arret du GPHONE"
            Message -Messageboxbody "Votre softphone Bria n'est plus ouvert, le bandeau ne peut plus fonctionner."
        } 
    }
    $Pproc = Phone_Surv #VÃ©rification prÃ©sence Softphone / mise Ã  jour
    $Gsurv = GPhone_Surv #VÃ©rification prÃ©sence GPhone / mise Ã  jour

    [System.GC]::Collect()  # Purge charge mÃ©moire utilisÃ©e
    
} While ($Boucle_infinie -eq 0) # Boucle infinie