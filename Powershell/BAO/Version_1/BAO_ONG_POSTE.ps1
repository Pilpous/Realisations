###-----------------------------------------------------------------------------------------------###
### ONGLET POSTE
###-----------------------------------------------------------------------------------------------###

function Test-ADComputer {
   Param([Parameter(Mandatory=$true)][string]$Identity)
   $filter = 'Name -eq "'+$Identity+'" -or DisplayName -eq "'+$Identity+'" -or UserPrincipalName -eq "'+$Identity+'"'
   $Comps = Get-ADComputer -Filter $filter
   ($Comps  -ne $null)
}

Function Check-ModuleConfigurationManager { 
If ([IntPtr]::Size -eq 8){$SCCMModulePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'}
Else {$SCCMModulePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'}
     If (-not(Test-Path -Path $SCCMModulePath)) 
     {
         Write-Host "Impossible de charger le module SCCM, le fichier $SCCCMModulePath est introuvable"
		 echl "[ii]`tImpossible de charger le module SCCM, le fichier $SCCCMModulePath est introuvable"

     } 
     Else 
     {
         If ((Get-Module ConfigurationManager) -eq $null)
          {
            Import-Module $SCCMModulePath

           Write-Host "Chargement du module ConfigurationManager"
			 echl "[ii]`tChargement du module ConfigurationManager"
                        } 
          Else 
          {
           Write-Host "Module ConfigurationManager déjà chargé"
  			 echl "[ii]`tModule ConfigurationManager déjà chargé"

         }
    }
}

Function RechTitulairePoste {
Check-ModuleConfigurationManager
New-PSDrive -Name P29 -PSProvider CMSite -Root "server" -ErrorAction SilentlyContinue
Set-Location 'P29:'

$global:titulaire = Get-CMUserDeviceAffinity -DeviceName $computer

$lbl_poste_infos_titsccmres.text = $titulaire.UniqueUserName.split("\")[1]

Set-Location 'C:'
}

Function RechPosteAssocie {
New-PSDrive -Name P29 -PSProvider CMSite -Root "server" -ErrorAction SilentlyContinue
Set-Location 'P29:'

$global:poste_asso = Get-CMUserDeviceAffinity -UserName $titulaire.UniqueUserName

if ( $poste_asso.ResourceName -eq $null){$lbl_poste_infos_posccmres.text = "Aucun poste associé"}
Else {$lbl_poste_infos_posccmres.text = $poste_asso.ResourceName }

Set-Location 'C:'
}

Function LastBoot {
$Variable=Get-WmiObject -comp $computer win32_operatingsystem | select csname, @{LABEL=’LastBootUpTime’;EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
Foreach ($Caracteristique in $Variable){
$boot = $Caracteristique.lastbootuptime} 
$testdate = Get-date $boot -Format 'dd/MM/yyyy hh:mm:ss'
$lbl_poste_infos_bootres.text = $testdate
}

Function Verifdouappel {
switch -Wildcard ($IP)
{
    "10.1*"{$script:PathSCCM = "C:\Temp\"
#			$lbl_poste_infos_ipres.text = $IP
			$lbl_poste_infos_sitres.text = "sur Site"
			$lbl_poste_infos_etstares.text = "-"
			$lbl_poste_infos_ipstares.text = "-"
				        }
    
    "10.227*"{$lbl_poste_infos_sitres.text = "Télétravail"
#			$lbl_poste_infos_ipres.text = $IP
			$lbl_poste_infos_etstares.text = "-"
			$lbl_poste_infos_ipstares.text = "-"
			}

    "192.*"{$lbl_poste_infos_sitres.text = "Télétravail"
#			$lbl_poste_infos_ipres.text = $IP
			$lbl_poste_infos_etstares.text = "-"
			$lbl_poste_infos_ipstares.text = "-"
			}
    
    Default {$ip = [regex]::match($IP,'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b').value
             $ipsta = [regex]::match($ip, '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.').ToString()
             $script:ipsta2 = $ipsta+"70"
             $script:ipagence = $ipsta+"2"
			$agence = Get-Content \\servicepilot\EXPGLB\SCEPILOT\ScePilotData\Import\Agances.txt| Select-String $ipagence | Out-String
			$Agencecut =$Agence.indexof(":")
			$Agencecut = $AgenceCut -7
			$script:AgenceN = $Agence.Substring(6,$AgenceCut)
			$script:AgenceNom = $AgenceN.ToUpper()
            $lbl_poste_infos_sitres.text = "Agence  : $AgenceNom"
#			$lbl_poste_infos_ipres.text = $IP
			$lbl_poste_infos_ipstares.text = $ipsta2
		    

            $TestSTA = Test-Connection -ComputerName $ipsta2 -Count 1 -Quiet
            If ($TestSTA -eq "True"){$lbl_poste_infos_etstares.text = "en ligne"
									#$script:PathSCCM = "\\$ipsta2\c$\sasloc\numerisation"
                                    $script:PathSCCM = "C:\temp\"	}
            Else {$lbl_poste_infos_etstares.text = "hors ligne"
				$script:PathSCCM = "C:\temp\" }
           
    }
}}

function Get-ProcessInfo { 
$global:array_proc = New-Object System.Collections.ArrayList 
$global:Process_Option = New-CimSessionoption -Protocol Dcom 
$global:Process_session = New-CimSession -ComputerName $computer -SessionOption $Process_Option 

$global:procInfo = Get-CimInstance -ClassName win32_process -Property * -CimSession $Process_session | Select ProcessId,Name,Path,VM,WS,PSComputerName | sort -Property Name 
$array_proc.AddRange($procInfo) 
$dtgv_poste_proc_proc.DataSource = $array_proc
} 

Function Kill_Proc { 
$selectedRow = $dtgv_poste_proc_proc.selectedcells.value
Write-Host $selectedRow.count

if ($selectedRow.count -ne 1) { 
Write-Host $selectedRow[0]
$procline = $selectedRow[0]
Get-CimInstance -ClassName win32_process -CimSession $Process_session -Filter "ProcessID='$procline'" | Invoke-CimMethod -MethodName Terminate
Get-ProcessInfo
}
Else { Write-Host $selectedRow

Get-CimInstance -ClassName win32_process -CimSession $Process_session -Filter "ProcessID='$selectedRow'" | Invoke-CimMethod -MethodName Terminate
Get-ProcessInfo
}
 }

function Get-ServiceInfo { 
$array_serv = New-Object System.Collections.ArrayList 
$script:Services_option = New-CimSessionoption -Protocol Dcom 
$script:Services_session = New-CimSession -ComputerName $computer -SessionOption $Services_option 

$Script:ServInfo = Get-CimInstance -ClassName win32_service -Property * -CimSession $Services_session | Select ProcessId,Name,Caption,State,Status,SystemName | sort -Property Name 
$array_serv.AddRange($ServInfo) 
$dtgv_poste_serv_serv.DataSource = $array_serv 
#$form1.refresh() 
} 

Function AR_Serv {
$selected_serv = $dtgv_poste_serv_serv.selectedcells.value
if ($selected_serv.count -ne 1) { 
			$selserv = $selected_serv[0]
			Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='$selserv'" | Invoke-CimMethod -MethodName StopService
			timeout /t 10
			Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='$selserv'" | Invoke-CimMethod -MethodName StartService
			}
Else { 		Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='$selected_serv'" | Invoke-CimMethod -MethodName StopService
			timeout /t 10
			Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='$selected_serv'" | Invoke-CimMethod -MethodName StartService
			}
Get-ServiceInfo
}

function fastping {
[CmdletBinding()]
param(
[String]$computer = "127.0.0.1",
[int]$delay = 100
)
$ping = new-object System.Net.NetworkInformation.Ping
try {
if ($ping.send($computer,$delay).status -ne "Success") { #return $false
															$lbl_poste_infos_ipres.text = "Poste injoignable"
															echl "[ii]`tPoste injoignable" }
else { 		
		RechTitulairePoste
		RechPosteAssocie
		LastBoot
		Verifdouappel
#		Get-Infosystem
		InfoSyst
		
		Get-ProcessInfo
		Get-ServiceInfo
		
		#$lbl_poste_infos_netbios.Text = "Information sur le poste $TempComputer :"
		$rtxtb_bao_act.appendtext("`rLancement des fonctions Poste")
		echl "[ii]`tLancement des fonctions Poste"
		}
} catch {
#return $false
$lbl_poste_infos_ipres.text = "Poste injoignable"
echl "[ii]`tPoste injoignable"
}
}

Function MonterNumérisation  {
start explorer.exe "\\$ipsta2\c$\sasloc\numerisation"
}

Function PDMSTA {
If ([IntPtr]::Size -eq 8){start "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$ipsta2}
Else {start "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$ipsta2}
 }

Function RebootSTA {
Restart-Computer -ComputerName $ipsta2 -Force
}

Function RemotePoste {
#$lbl_poste_infos_ipres.text
If ([IntPtr]::Size -eq 8){start "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$lbl_poste_infos_ipres.text}
Else {start "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$lbl_poste_infos_ipres.text}
}

Function PDMPoste {
msra.exe /offerra $lbl_poste_infos_ipres.text
}

Function InfoSyst {
$InfoSys_option = New-CimSessionoption -Protocol Dcom 
$InfoSys_session = New-CimSession -ComputerName $computer -SessionOption $InfoSys_option 
$InfoSys = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $InfoSys_session
$InfoOS = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $InfoSys_session
$InfoDisk = Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $InfoSys_session
$InfoDiskStatus = Get-WMIObject -Computer $computer -Class Win32_DiskDrive |where -Property Name -Like "\\.\PHYSICALDRIVE0"


foreach($disk in $InfoDisk){
if ($disk.DriveType -eq 3){
$size = "{0:n0} GB" -f (($disk | Measure-Object -Property Size -Sum).sum/1gb)
$FreeSpace = "{0:n0} GB" -f (($disk | Measure-Object -Property FreeSpace -Sum).sum/1gb)}}

$RAM = "{0:n2} GB" -f ($InfoSys.TotalPhysicalMemory/1gb)


$lbl_poste_infos_modelres.text = $InfoSys.Model
$lbl_poste_infos_osres.text = $InfoOS.Caption
$lbl_poste_infos_archires.text =  $InfoOS.OSArchitecture
#Write-Host $InfoDisk.DeviceID
$lbl_poste_infos_sizeres.text = $Size
$lbl_poste_infos_freespaceres.text = $FreeSpace
$lbl_poste_infos_ramres.text = $RAM
$lbl_poste_infos_diskstatres.text = $InfoDiskStatus.status
}

Function ConsSCCM {
If ([IntPtr]::Size -eq 8){start "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"}
Else {start "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"}
}

Function RepSCCM {
$oSCCM = [wmiclass] “\\$computer\root\ccm:sms_client”
$oSCCM.RepairClient()
}

Function ConfLogon {
Set-Location C:\Applics\CFLogon
start C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ArgumentList "C:\APPLICS\CFLogon\ConfigurationLogon.PS1" -Credential $cred_glb
Set-Location C:\AppDSI\DEVGLB\OUTILSTA\BAO
}

Function ExtractIP{

$Global:Extract = [System.Net.Dns]::GetHostEntry($TempComputer) | select hostname,AddressList
if($Extract.AddressList.IPAddressToString.count -eq 1 ) {
$GLOBAl:IP = $Global:Extract.AddressList.IPAddressToString }
Else { $GLOBAl:IP = $Global:Extract.AddressList.IPAddressToString[0] }

$GLOBAL:computer = $Global:Extract.HostName.split(".")[0].toupper()

$lbl_poste_infos_netbios.Text = "Information sur le poste $computer"

$lbl_poste_infos_ipres.text = $IP
}