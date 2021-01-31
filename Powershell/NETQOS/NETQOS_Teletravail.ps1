<#
.Synopsis
NETQOS - Test de connexion TELETRAVAIL / TRAVAIL MOBILE

.Description
Lancement du script à l'ouverture de session - Vérification de la connexion
Si IP EndUser/PulseSecure : poursuite du script
Lancement d'un test de débit toutes les 30min : 
=> fichier de 5Mo copié du poste vers le NAS + calcul du temps
=> fichier de 5Mo copié du NAS vers le poste + calcul du temps

Lancement d'un test Temps de réponse toutes les 10min :
=> 100 ping vers le NAS
=> si 1 ping sur 100 KO : 100 ping vers Google

Inscription des résultats dans la base SQL NETQOS

Version 1.0 - 19/07/2018
Version 1.1 - 01/08/2018 - Correction Bugs
Version 1.2 - 31/07/2019 - W10 - Correction Bugs W10 - Ajout du compte b_netqos comme Admin W10
Version 1.3 - 23/09/2019 - Modification de la cible NAS suite demande de R2T
        => l'ensemble des postes pointaient sur le NAS22
        => les postes pointent désormais vers le NAS du département du poste
        => Réduction de la taille du fichier utilisé pour le SpeedTest : passe de 10 à 5Mo
Version 2.0 - 21/04/2020 - Refonte complète du script : utilisation de nouvelle cmdlet dispo depuis W10  
        + mise en place de nouvelle plage d'adresses pour PulseSecure - 10.210.XX.XX - suite demande de R2T
Version 2.1 - 19/10/2020 - Modification de la cible NAS suite demande de R2T :
        => l'ensemble des postes va faire ses SpeedTests sur le NAS35
        => fréquence du SpeedTest modifiée : de 30min à 60min
        => l'ensemble des PING INTRA se fera vers le NAS35

#>


$iguazu_server = "Server"


Import-Module Logs_TA
$Global:Current_Folder = 'C:\AppDsi\EXPGLB\NETQOS\'
Set-Location $Current_Folder

Create_Log -Path_Log $Current_Folder -Appli NETQOS_Teletravail
Log -Add "#################################################################"
Log -Add "###         Lancement du script"
Log -Add "#################################################################"

### FONCTION SQL
Function Connect_SQL {
    Param (
        [Parameter(Mandatory=$true)] [string]$command_sql,
        [Parameter(Mandatory=$true)] [object[]]$data_sql
    )

    Try {
        $cOK = 0
        $cErr = 1
        $gRetour = $cOK
        Log -add  "Controle de la connexion a la base LOG_TELETRAVAIL en OLEDB"
        $gConnect = "Server=$iguazu_server;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
        $oConnMSSQL = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $oConnMSSQL.ConnectionString = $gConnect
        $wret = $oConnMSSQL.open()
        #Ouverture d'une transaction
        $oTrans = $oConnMSSQL.BeginTransaction()

        Log -Add "Mise a jour base SQL"
        switch -Wildcard ($command_sql) {
            "PDT" { $wret = Ecrire_PDT_TEST $data_sql $oConnMSSQL $oTrans }
            "PNG" { $wret = Ecrire_TEST_PDT_PING $data_sql $oConnMSSQL $oTrans }
            "DEB" { $wret = Ecrire_TEST_DEBIT_PDT $data_sql $oConnMSSQL $oTrans }
            "RET" { $wret = Ecrire_TEST_TYPE_DEBIT $data_sql $oConnMSSQL $oTrans}
            Default { }
        }
        $wret = $oTrans.Commit()
    } 
    Catch [System.Data.SqlClient.SqlException] {
        Log -Add "Pb acces base SQL : $($Error[0].Exception.ErrorRecord.FullyQualifiedErrorId)"
        
    }
    Catch { $gRetour=$cErr
            Log -add $Error[0].Exception.ErrorRecord.FullyQualifiedErrorId }
    Finally { $oConnMSSQL.Close() 
                Log -Add "Mise a jour base SQL terminee !"}

}
### FONCTIONS SQL - Fournie ###
Function Ecrire_PDT_TEST { param ( $oRow, $oConn, $oTrans )
    Try {
          $iRet=$cOK;  $wPr = $MyInvocation.MyCommand.Name + " : "; $wr=0; $wc=0
          #--Le ID_TEST est le même pour le tuplet ID_PDT, IP_PUBLIC, IP _PRIVATE, IP_PULSE sur une même journée
          #--on ne fait pas d'insertion car on estime que c'est le même identifiant de test, sinon on insère un nouveau test
          #--Si une IP change ou la journée change sur le même pdt, le ID_TEST sera incrémenté car on estime que c'est un nouveau test
          $Req = "SELECT ID_TEST FROM PDT_TEST "
          $Req = $Req + "WHERE ID_PDT = '" + $oRow.ID_PDT + "' "
          $Req = $Req + "AND IP_PUBLIC = '" + $oRow.IP_PUBLIC + "' "
          $Req = $Req + "AND IP_PRIVATE = '" + $oRow.IP_PRIVATE + "' "
          $Req = $Req + "AND IP_PULSE = '" + $oRow.IP_PULSE +  "' "
          $Req = $Req + "AND convert(date, DATE_TEST) = convert(date, '" + $oRow.DATE_TEST + "')"
          $oCmdc = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdc = $oConn.CreateCommand()
          $oCmdc.CommandText = $Req
          $oCmdc.Transaction = $oTrans
          $wc = $oCmdc.ExecuteScalar()
          
          if ($wc -eq '' -or $wc -eq $null) {
                 $Reqi = "INSERT INTO PDT_TEST "
                 $Reqi = $Reqi + " ( ID_PDT, MODELE_PDT, DESC_PDT, USER_TEST, IP_PRIVATE, IP_PULSE, IP_PUBLIC, DATE_TEST"
                 $Reqi = $Reqi + ") Values ("
                 $Reqi = $Reqi + "'" + $oRow.ID_PDT + "'"
                 $Reqi = $Reqi + ",'" + $oRow.MODELE_PDT + "'"
                 $Reqi = $Reqi + ",'" + $oRow.DESC_PDT + "'"
                 $Reqi = $Reqi + ",'" + $oRow.USER_TEST + "'"
                 $Reqi = $Reqi + ",'" + $oRow.IP_PRIVATE + "'"
                 $Reqi = $Reqi + ",'" + $oRow.IP_PULSE + "'"
                 $Reqi = $Reqi + ",'" + $oRow.IP_PUBLIC + "'"
                 $Reqi = $Reqi + ",'" + $oRow.DATE_TEST + "'"
                 $Reqi = $Reqi + ")"
                 
                 $oCmdm = New-Object -TypeName System.Data.SqlClient.SqlCommand
                 $oCmdm = $oConn.CreateCommand()
                 $oCmdm.Transaction = $oTrans
                 $oCmdm.CommandText = $Reqi 
                 Log -add ($wpr + "requete :" + $oCmdm.CommandText)
                 $wr = $oCmdm.ExecuteNonQuery()
          }
          return $iRet
          }
   Catch  {
          $iRet=$cErr
          Log -add $Req
          Log -add  $wPr 
          Log -add  $_.exception 
          return $iRet
    }
}

Function Ecrire_TEST_PDT_PING { param ( $oRow, $oConn, $oTrans )
    Try {
          
          #--On récupère l'identifiant de test
          $iRet=$cOK;  $wPr = $MyInvocation.MyCommand.Name + " : "; $wr=0; $wc=0
          $Req = "SELECT ID_TEST FROM PDT_TEST "
          $Req = $Req + "WHERE ID_PDT = '" + $oRow.ID_PDT + "' "
          $Req = $Req + "AND IP_PUBLIC = '" + $oRow.IP_PUBLIC + "' "
          $Req = $Req + "AND IP_PRIVATE = '" + $oRow.IP_PRIVATE + "' "
          $Req = $Req + "AND IP_PULSE = '" + $oRow.IP_PULSE + "' "
          $Req = $Req + "AND convert(date, DATE_TEST) = convert(date, '" + $oRow.DATE_TEST + "')"
          $oCmdc = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdc = $oConn.CreateCommand()
          $oCmdc.CommandText = $Req
          $oCmdc.Transaction = $oTrans
          $wc = $oCmdc.ExecuteScalar()
          
          $Reqi = "INSERT INTO TEST_PDT_PING "
          $Reqi = $Reqi + " ( ID_TEST, PING_TYPE, PING_PCT_OK, TPS_MAX_MS, TPS_MIN_MS, TPS_AVG_MS, DATE_TEST_PING"
          $Reqi = $Reqi + ") Values ("
          $Reqi = $Reqi + "'" + $wc + "'"
          $Reqi = $Reqi + ",'" + $oRow.PING_TYPE + "'"
          $Reqi = $Reqi + ",'" + $oRow.PING_PCT_OK + "'"
          $Reqi = $Reqi + ",'" + $oRow.TPS_MAX_MS + "'"
          $Reqi = $Reqi + ",'" + $oRow.TPS_MIN_MS + "'"
          $Reqi = $Reqi + ",'" + $oRow.TPS_AVG_MS + "'"
          $Reqi = $Reqi + ",'" + $oRow.DATE_TEST_PING + "'"
          $Reqi = $Reqi + ")"

          $oCmdm = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdm = $oConn.CreateCommand()
          $oCmdm.Transaction = $oTrans
          $oCmdm.CommandText = $Reqi 
          Log -add  ($wpr + "requete :" + $oCmdm.CommandText)
          $wr = $oCmdm.ExecuteNonQuery()
          return $iRet
          }
   Catch  {
          $iRet=$cErr
          Log -add $wPr 
          Log -add $_.exception 
          return $iRet
    }
}

Function Ecrire_TEST_DEBIT_PDT { param ( $oRow, $oConn, $oTrans )
    Try {
          $iRet=$cOK;  $wPr = $MyInvocation.MyCommand.Name + " : "; $wr=0; $wc=0
          #--On récupère le ID_TEST
          $Req = "SELECT ID_TEST FROM PDT_TEST "
          $Req = $Req + "WHERE ID_PDT = '" + $oRow.ID_PDT + "'"
          $Req = $Req + " AND IP_PUBLIC = '" + $oRow.IP_PUBLIC + "'"
          $Req = $Req + " AND IP_PRIVATE = '" + $oRow.IP_PRIVATE + "'"
          $Req = $Req + " AND IP_PULSE = '" + $oRow.IP_PULSE + "'"
          $Req = $Req + " AND convert(date, DATE_TEST) = convert(date, '" + $oRow.DATE_TEST + "')"
            $oCmdc = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdc = $oConn.CreateCommand()
          $oCmdc.CommandText = $Req
          $oCmdc.Transaction = $oTrans
          $wc = $oCmdc.ExecuteScalar()
          
          $Reqi = "INSERT INTO TEST_DEBIT_PDT "
          $Reqi = $Reqi + " ( ID_TEST, DESC_MBPS, ASC_MBPS, DATE_TEST_DEBIT"
          $Reqi = $Reqi + ") Values ("
          $Reqi = $Reqi + "'" + $wc + "'"
          $Reqi = $Reqi + ",'" + $oRow.DESC_MBPS + "'"
          $Reqi = $Reqi + ",'" + $oRow.ASC_MBPS + "'"
          $Reqi = $Reqi + ",'" + $oRow.DATE_TEST_DEBIT + "'"
          $Reqi = $Reqi + ")"
          
          $oCmdm = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdm = $oConn.CreateCommand()
          $oCmdm.Transaction = $oTrans
          $oCmdm.CommandText = $Reqi 
          Log -add ($wpr + "requete :" + $oCmdm.CommandText)
          $wr = $oCmdm.ExecuteNonQuery()
          return $iRet
          }
   Catch  {
          $iRet=$cErr
          Log -add $wPr 
          Log -add $_.exception 
          return $iRet
    }
}

Function Ecrire_TEST_TYPE_DEBIT { param ( $oRow, $oConn, $oTrans )
    Try {
          $iRet=$cOK;  $wPr = $MyInvocation.MyCommand.Name + " : "; $wr=0; $wc=0

          $Req = "SELECT TEST_DEBIT_PDT.ID_DEBIT_TEST FROM PDT_TEST "
          $Req = $Req + "INNER JOIN TEST_DEBIT_PDT ON TEST_DEBIT_PDT.ID_TEST = PDT_TEST.ID_TEST "
          $Req = $Req + "WHERE ID_PDT = '" + $oRow.ID_PDT + "'"
          $Req = $Req + " AND IP_PUBLIC = '" + $oRow.IP_PUBLIC + "'"
          $Req = $Req + " AND IP_PRIVATE = '" + $oRow.IP_PRIVATE + "'"
          $Req = $Req + " AND IP_PULSE = '" + $oRow.IP_PULSE + "'"
          $Req = $Req + " AND DATE_TEST_DEBIT = '" + $oRow.DATE_TEST + "'"
            $oCmdc = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdc = $oConn.CreateCommand()
          $oCmdc.CommandText = $Req
          $oCmdc.Transaction = $oTrans
          $wc = $oCmdc.ExecuteScalar()

          
          $Reqi = "INSERT INTO TEST_TYPE_DEBIT "
          $Reqi = $Reqi + " (  ID_DEBIT_TEST, TYPE_TEST, RETOUR, DATE_RETOUR"
          $Reqi = $Reqi + ") Values ("
          $Reqi = $Reqi + "'" + $wc + "'"
          $Reqi = $Reqi + ",'" + $oRow.TYPE_TEST + "'"
          $Reqi = $Reqi + ",'" + $oRow.RETOUR + "'"
          $Reqi = $Reqi + ",'" + $oRow.DATE_RETOUR + "'"
          $Reqi = $Reqi + ")"
          
          $oCmdm = New-Object -TypeName System.Data.SqlClient.SqlCommand
          $oCmdm = $oConn.CreateCommand()
          $oCmdm.Transaction = $oTrans
          $oCmdm.CommandText = $Reqi 
          Log -add ($wpr + "requete :" + $oCmdm.CommandText)
          $wr = $oCmdm.ExecuteNonQuery()
          return $iRet
          }
   Catch  {
          $iRet=$cErr
          Log -add $wPr 
          Log -add $_.exception 
          return $iRet
    }
}


### FONCTIONS W10
Function Kiss_my_NAS {
    $Nas_Status = 0
    Log -Add "Lancement Fonction detection departement pour cibler NAS + test de connexion"
    $script:nas_dep = "nas35"
        $script:cible_nas = "\\nas35\Logs\Logs_Teletravail\"
        if (Test-Path $cible_nas){
        Log -add "DEP 35 - NAS OK"	
        $Nas_Status = 0}
        else { Log -Add "DEP 35 - Acces NAS KO"
        $Nas_Status = 1}

        return $Nas_Status
}

Function First_Blood {
    # Recolte info
    Log -Add "Fonction : recuperation des infos pour initialisation de la recolte"
    $script:Trigramme = (Get-CimInstance -ClassName Win32_ComputerSystem).username.split("\")[1] 
    $Date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model.trim()

    $oTest = New-Object System.Object
    $oTest | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
    $oTest | Add-Member -Type NoteProperty -Name MODELE_PDT -Value "$Model"
    $oTest | Add-Member -Type NoteProperty -Name DESC_PDT -Value "$($script:Connexion.interfaceAlias[1])"
    $oTest | Add-Member -Type NoteProperty -Name USER_TEST -Value "$script:trigramme"
    $oTest | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
    $oTest | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
    $oTest | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
    $oTest | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$Date"

    # Mise à jour SQL
    Connect_SQL -command_sql PDT -data_sql $oTest
    Log -Add "$($oTest.ID_PDT) ; $($oTest.MODELE_PDT) ; $($oTest.DESC_PDT) ; $($oTest.USER_TEST) ; $($oTest.IP_PRIVATE) ; $($oTest.IP_PULSE) ; $($oTest.IP_PUBLIC) ; $($oTest.DATE_TEST)"
}

Function Tracert {
Log -Add "Lancement Function Tracert"
if ($null -eq $script:gateway){
    $script:liste_adresse = New-Object System.Object
    $connexion_info = Test-NetConnection -ComputerName www.google.fr -TraceRoute #-Hops 3
    $script:gateway = $connexion_info.traceroute[0]
    Log -Add "Gateway : $($connexion_info.traceroute[0])"  
    Log -Add "IP Publique : $($connexion_info.traceroute[2])"
    Log -Add $connexion_info
        $script:liste_adresse  | Add-Member -Type NoteProperty -Name GATEWAY -Value $connexion_info.Traceroute[0]
        $script:liste_adresse  | Add-Member -Type NoteProperty -Name PUBLIQUE -Value $connexion_info.Traceroute[2]
        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IFINDEX -Value $connexion_info.InterfaceIndex
    } else { Log -Add "Demande deja traitee"
             Log -Add "Gateway : $($liste_adresse.GATEWAY)"
             Log -Add "IP Publique : $($liste_adresse.PUBLIQUE)"
             }              
}

Function Add_Road {
Log -Add "Lancement Function Add_Road"
$on_the_road = Get-NetRoute | Select-Object DestinationPrefix
if ($on_the_road -match "8.8.8.8/32") { Log -Add "Route deja existante"}
else { 
    try {
    New-NetRoute -DestinationPrefix "8.8.8.8/32" -InterfaceIndex $liste_adresse.IFINDEX -NextHop $liste_adresse.GATEWAY -PolicyStore ActiveStore -InformationAction SilentlyContinue 

    Log -Add "Creation route Google OK"
        }
    catch { Log -Add $Error[0] }
}
}

Function Delete_Road {
Log -Add "Lancement Function Delete_Road"
    $on_the_road = Get-NetRoute | Select-Object DestinationPrefix
    if ($on_the_road -match "8.8.8.8/32") { 
                Remove-NetRoute -DestinationPrefix "8.8.8.8/32" -Confirm:$false -InformationAction SilentlyContinue
                Log -Add "Suppression Route existante"}
    else {Log -Add "Route non presente" }
}

Function Get_IP {
Log -Add "Lancement fct Get_IP"
if ($null -eq $script:liste_adresse ) { $script:liste_adresse = New-Object System.Object }

    $get_ip = Get-NetIPAddress -AddressState Preferred -AddressFamily IPv4 | Where-Object IPAddress -NE "127.0.0.1" | Select-Object IPAddress,InterfaceAlias

    ForEach ($ip in $get_ip.IPAddress ) {
        switch -Wildcard ($ip) {
            "10.122.*" { Log -Add "IP SITE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.22.*"  { Log -Add "IP AGENCE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.129.*" { Log -Add "IP SITE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.29.*"  { Log -Add "IP AGENCE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.144.*" { Log -Add "IP SITE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.44.*"  { Log -Add "IP AGENCE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.149.*" { Log -Add "IP SITE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.49.*"  { Log -Add "IP AGENCE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.156.*" { Log -Add "IP SITE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.56.*"  { Log -Add "IP AGENCE : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.135.*" { Log -Add "IP SITE : $($ip)"
                         $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_SITE -Value $ip}
            "10.35.*"  { Log -Add "IP AGENCE : $($ip)"
                         $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AG -Value $ip  }
            "10.227.*" { Log -Add "IP ENDUSER : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_PULSE -Value $ip  }
            "10.208.*" { Log -Add "IP PULSE : $($ip)" 
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_PULSE -Value $ip  }
            "10.210.*" { Log -Add "IP PULSE : $($ip)" 
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_PULSE -Value $ip  }
            Default { Log -Add "IP : $($ip)"
                        $script:liste_adresse  | Add-Member -Type NoteProperty -Name IP_AUTRE -Value $ip  }
        } # SWITCH
    }# FOREACH
} #FUNCTION

Function Surv_IP {
    Param (
        [int]$time_to_wait = 15,
        [int]$nb_execution = 0
     )
    Do {
        $script:Connexion = Get-NetConnectionProfile
        switch -Wildcard ($script:Connexion.count){
                $null {  if ($script:Connexion.Name -eq "intra.groupama.fr") { Log -Add "NULL - Connected to GLB - Site/Agence"
                                                                        $connected = 3 } 
                    else { Log -Add "Connected to $($script:Connexion.Name)"
                            Log -Add "En attente connexion GLB"
                            Tracert 
                            $connected = 0 }
                    }
                0 { Log -Add "Pas de reseau" 
                    Log -Add "En attente connexion"
                    $connected = 0 }
                1 { Log -Add "Connected"
                    if ($script:Connexion.Name -eq "intra.groupama.fr") { Log -Add "Connected to GLB - Site/Agence"
                                                                        $connected = 3  } 
                    else { Log -Add "Connected to $($script:Connexion.Name)"
                            Log -Add "En attente connexion GLB"
                            Tracert
                            $connected = 0 
                            }
                    }
                2 { Log -Add "Connected" 

                    if ($script:Connexion.Name -eq "intra.groupama.fr") { Log -Add "Connected to GLB - Teletravail"
                        $connected = 5 } 
                    # VERIF ADRESSE IP 
                    else {
                        $ip = @()
                        foreach ($interface in $script:Connexion.InterfaceIndex ) {
                            $ip += Get-NetIPAddress | Where-Object { $_.InterfaceIndex -eq $interface}
                        }
                        if (($ip.IPAddress -match '10.227.') -or ($ip.IPAddress -match '10.208.') -or ($ip.IPAddress -match '10.210.')){ Log -Add "Connected to GLB - Teletravail" 
                            $connected = 5 }
                    }
                    }
                Default {  if ($script:Connexion.Name -eq "intra.groupama.fr") { Log -Add "DEFAULT - Connected to GLB - Site/Agence"
                                                                        $connected = 3 } }
        } # switch

        [System.GC]::Collect()
        Start-Sleep -Seconds $time_to_wait
        $nb_execution ++
        Log -Add "Execution numero $($nb_execution)"

        } # DO
        until (($connected -eq 5) -or ($connected -eq 3) -or ($nb_execution -eq 500 ))

        if ($nb_execution -eq 500) { return $nb_execution}
        else { return $connected }

} # FUNCTION

Function Ping_intra {
Log -Add "Lancement fct Ping_intra"
##Fonction Ping GLB
$heureping = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$intra_connexion=@()
for ($i=1; $i -le 100; $i++)
    { $intra_ping = Test-NetConnection $script:nas_dep  -InformationLevel Detailed -InformationAction SilentlyContinue
       $intra_connexion += $intra_ping }

$intra_ping_min = ($intra_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Minimum).minimum
$intra_ping_max = ($intra_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Maximum).maximum
$intra_ping_avg = [Math]::Round(($intra_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Average).average , 2)

$script:intra_ping_OK = ($intra_connexion | Where-Object { $_.PingSucceeded -eq $true } ).count

$oTestPing = New-Object System.Object
$oTestPing | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
$oTestPing | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$heureping"
$oTestPing | Add-Member -Type NoteProperty -Name PING_TYPE -Value "GLB"
$oTestPing | Add-Member -Type NoteProperty -Name PING_PCT_OK -Value "$intra_ping_OK "
$oTestPing | Add-Member -Type NoteProperty -Name TPS_MAX_MS -Value "$intra_ping_max"
$oTestPing | Add-Member -Type NoteProperty -Name TPS_MIN_MS -Value "$intra_ping_min"
$oTestPing | Add-Member -Type NoteProperty -Name TPS_AVG_MS -Value "$intra_ping_avg"
$oTestPing | Add-Member -Type NoteProperty -Name DATE_TEST_PING -Value "$heureping"

Log -Add "$($oTestPing.ID_PDT);$($oTestPing.IP_PRIVATE);$($oTestPing.IP_PULSE);$($oTestPing.IP_PUBLIC);$($oTestPing.DATE_TEST);$($oTestPing.PING_TYPE);$($oTestPing.PING_PCT_OK);$($oTestPing.TPS_MAX_MS);$($oTestPing.TPS_MIN_MS);$($oTestPing.TPS_AVG_MS);$($oTestPing.DATE_TEST_PING)"
Log -Add "PING - Verification connexion"
$verif_connexion = Surv_IP -time_to_wait 5
    if ($verif_connexion -eq 5) {
        Log -Add "PING - Connexion OK - Mise a jour table SQL"
        Connect_SQL -command_sql PNG -data_sql $oTestPing
    }
    elseif ($verif_connexion -eq 500 -or $verif_connexion -eq 3 ) {
        Log -Add "Delai depasse ou connexion sur site/agence"
        Log -Add "Arret du script"
        $script:nb_ping = 3
        exit
    }

}

Function Ping_inter {
Log -Add "Lancement fct Ping_inter"
##Fonction Ping Google
$heureping = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$inter_connexion=@()
for ($i=1; $i -le 100; $i++)
    {  $inter_ping = Test-NetConnection 8.8.8.8 -InformationLevel Detailed -InformationAction SilentlyContinue -WarningAction SilentlyContinue
       $inter_connexion += $inter_ping }

$inter_ping_min = ($inter_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Minimum).minimum
$inter_ping_max = ($inter_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Maximum).maximum
$inter_ping_avg = [Math]::Round(($inter_connexion.PingReplyDetails.Roundtriptime | Measure-Object -Average).average , 2)

$script:inter_ping_OK = ($inter_connexion | Where-Object { $_.PingSucceeded -eq $true } ).count

$oTestPing = New-Object System.Object
$oTestPing | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
$oTestPing | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
$oTestPing | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$heureping"
$oTestPing | Add-Member -Type NoteProperty -Name PING_TYPE -Value "Google"
$oTestPing | Add-Member -Type NoteProperty -Name PING_PCT_OK -Value "$inter_ping_OK "
$oTestPing | Add-Member -Type NoteProperty -Name TPS_MAX_MS -Value "$inter_ping_max"
$oTestPing | Add-Member -Type NoteProperty -Name TPS_MIN_MS -Value "$inter_ping_min"
$oTestPing | Add-Member -Type NoteProperty -Name TPS_AVG_MS -Value "$inter_ping_avg"
$oTestPing | Add-Member -Type NoteProperty -Name DATE_TEST_PING -Value "$heureping"

Log -Add "$($oTestPing.ID_PDT);$($oTestPing.IP_PRIVATE);$($oTestPing.IP_PULSE);$($oTestPing.IP_PUBLIC);$($oTestPing.DATE_TEST);$($oTestPing.PING_TYPE);$($oTestPing.PING_PCT_OK);$($oTestPing.TPS_MAX_MS);$($oTestPing.TPS_MIN_MS);$($oTestPing.TPS_AVG_MS);$($oTestPing.DATE_TEST_PING)"

Log -Add "PING - Verification connexion"
$verif_connexion = Surv_IP -time_to_wait 5
    if ($verif_connexion -eq 5) {
        Log -Add "PING - Connexion OK - Mise a jour table SQL"
        Connect_SQL -command_sql PNG -data_sql $oTestPing
    }
    elseif ($verif_connexion -eq 500 -or $verif_connexion -eq 3 ) {
        Log -Add "Delai depasse ou connexion sur site/agence"
        Log -Add "Arret du script"
        $script:nb_ping = 3
        exit
    }

}

Function Check_SpeedTest {
    Param ( [int]$acces_nas )

    if ( $acces_nas -eq 0 ) { 
        Log -Add "Acces NAS OK - Lancement SpeedTest"
        Speed_Test }

    else { Log -Add "Acces NAS KO - Pas de lancement SpeedTest"
        $Date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        #$DBNull = [System.DBNull]::Value 
        
        $oTestDebit = New-Object System.Object
        $oTestDebit | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
        $oTestDebit | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
        $oTestDebit | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
        $oTestDebit | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
        $oTestDebit | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$Date"
        $oTestDebit | Add-Member -Type NoteProperty -Name DESC_MBPS -Value 0
        $oTestDebit | Add-Member -Type NoteProperty -Name ASC_MBPS -Value 0
        $oTestDebit | Add-Member -Type NoteProperty -Name DATE_TEST_DEBIT -Value "$Date"

        Log -Add "Ajout TEST DEBIT NULL"
        Connect_SQL -command_sql DEB -data_sql $oTestDebit
        Start-Sleep -Seconds 1

        $oTestRetour =  New-Object System.Object
        $oTestRetour | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$Date"
        $oTestRetour | Add-Member -Type NoteProperty -Name TYPE_TEST "SPEEDTEST_ERROR"
        $oTestRetour | Add-Member -Type NoteProperty -Name RETOUR "ACCES NAS KO"
        $oTestRetour | Add-Member -Type NoteProperty -Name DATE_RETOUR "$Date"
    
        Log -Add "$($oTestRetour.TYPE_TEST);$($oTestRetour.RETOUR);$($oTestRetour.DATE_RETOUR)"

        Log -Add "DEBIT - Erreur rencontree - Mise a jour table SQL"
        Connect_SQL -command_sql RET -data_sql $oTestRetour


    }


}
Function Speed_Test {
Log -Add "Lancement Fonction SpeedTest"
$Source = "C:\AppDSI\EXPGLB\NETQOS\Logs"

 Try {
    $TotalSize = (Get-ChildItem $Source\Ne_Pas_Supprimer.txt -ErrorAction Stop).Length
    $WriteTest = Measure-Command { 
        Log -Add "SpeedTest : Copie Local vers $($Cible_Nas) en cours"
        Copy-Item $Source\Ne_Pas_Supprimer.txt $Cible_Nas\SpeedTest_$($script:trigramme).txt -ErrorAction Stop -Force
        Log -Add "SpeedTest : Copie Local vers $($Cible_Nas) terminee"
    }
            
    $ReadTest = Measure-Command {
        Log -Add "SpeedTest : Copie $($Cible_Nas) vers Local en cours"
        Copy-Item $Cible_Nas\SpeedTest_$($script:trigramme).txt $Source\SpeedTestRead_$($script:trigramme).txt -ErrorAction Stop -Force
        Log -Add "SpeedTest : Copie $($Cible_Nas) vers Local terminee"
    }

    $WriteMbps = [Math]::Round((($TotalSize * 8) / $WriteTest.TotalSeconds) / 1048576,2)
    $ReadMbps = [Math]::Round((($TotalSize * 8) / $ReadTest.TotalSeconds) / 1048576,2)
    $RunTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    

    $oTestDebit = New-Object System.Object
    $oTestDebit | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
    $oTestDebit | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
    $oTestDebit | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
    $oTestDebit | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
    $oTestDebit | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$RunTime"
    $oTestDebit | Add-Member -Type NoteProperty -Name DESC_MBPS -Value "$ReadMbps"
    $oTestDebit | Add-Member -Type NoteProperty -Name ASC_MBPS -Value "$WriteMbps"
    $oTestDebit | Add-Member -Type NoteProperty -Name DATE_TEST_DEBIT -Value "$RunTime"

    Log -add "$($oTestDebit.ID_PDT) ; $($oTestDebit.IP_PRIVATE) ; $($oTestDebit.IP_PULSE);$($oTestDebit.IP_PUBLIC); SpeedTest ;$($oTestDebit.DATE_TEST);$($oTestDebit.DESC_MBPS);$($oTestDebit.ASC_MBPS)"

    }
    Catch {
        $Date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Log -Add "SpeedTest en Erreur : $($Error[0].Exception.ErrorRecord.FullyQualifiedErrorId)"

        $oTestRetour =  New-Object System.Object
        $oTestRetour | Add-Member -Type NoteProperty -Name ID_PDT -Value "$env:COMPUTERNAME"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PRIVATE -Value "$($liste_adresse.IP_AUTRE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PULSE -Value "$($liste_adresse.IP_PULSE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name IP_PUBLIC -Value "$($liste_adresse.PUBLIQUE)"
        $oTestRetour | Add-Member -Type NoteProperty -Name DATE_TEST -Value "$Runtime"
        $oTestRetour | Add-Member -Type NoteProperty -Name TYPE_TEST "SPEEDTEST_ERROR"
        $oTestRetour | Add-Member -Type NoteProperty -Name RETOUR "$($Error[0].FullyQualifiedErrorId)"
        $oTestRetour | Add-Member -Type NoteProperty -Name DATE_RETOUR "$Date"

        Log -Add "$($oTestRetour.TYPE_TEST);$($oTestRetour.RETOUR);$($oTestRetour.DATE_RETOUR);"
    }
    Finally {
        Log -Add "DEBIT - Verification connexion"
        $verif_connexion = Surv_IP -time_to_wait 5

        if ($verif_connexion -eq 5) {
            Log -Add "DEBIT - Connexion OK - Mise a jour table SQL"
            Connect_SQL -command_sql DEB -data_sql $oTestDebit

            if ($null -ne $oTestRetour ) {
                Log -Add "DEBIT - Erreur rencontree - Mise a jour table SQL"
                Connect_SQL -command_sql RET -data_sql $oTestRetour
            }
        }
        elseif ($verif_connexion -eq 500 -or $verif_connexion -eq 3 ) {
            Log -Add "Delai depasse ou connexion sur site/agence"
            Log -Add "Arret du script"
            $script:nb_ping = 3
        }
    }
}

Function Delete_proof {
    $Source = "C:\AppDSI\EXPGLB\NETQOS\Logs"
                ## Suppression des fichiers créés pour effectuer le SpeedTest : C:\ et \\nas22
                $Suppr_Fichier = "$($Cible_Nas)\SpeedTest_$($script:trigramme).txt","$($Source)\SpeedTestRead_$($script:trigramme).txt"
                foreach ($fichier in $Suppr_Fichier) {
                        try {
                                if ($null -ne (Get-ChildItem $fichier -ErrorAction SilentlyContinue)) {
                                    $Date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                                    Log -Add "Suppression du fichier $($fichier)"
                                    Remove-Item $fichier -Confirm:$false -Force
                                    }
                                } catch { 
                                    Log -Add "Suppression fichier Error : $($Error[0])" }
                }
}

Function Stop_Script_or_not {
    #Si presence fichier Arret_script.txt => arret de l execution du script
    if ((Test-Path $Cible_Nas\Arret_script.txt) -eq $true){
        Log -Add "Presence du fichier Arret_script.txt => Arret du script"
        exit
        }

}


### LANCEMENT SCRIPT

Do {
    $connected = 0
    $connected = Surv_IP

    if ($connected -eq 500) { Log -Add "Temps d'attente limite atteint - Arret du script"}
    if ($connected -eq 3 ) { Log -Add "ARRET DU SCRIPT" }
    if ($connected -eq 5 ) { Log -Add "TELETRAVAIL - Lancement du script"
        Get_IP
        First_Blood
        $script:nb_ping = 0
        
        $verif_NAS = Kiss_my_NAS
        Stop_Script_or_not 

        Check_SpeedTest -acces_nas $verif_NAS
       
        Delete_proof

        #VERIF SI TOUJOURS CONNECTE
        $connected = Surv_IP -time_to_wait 1
        if ($connected -eq 3 -or $connected -eq 500) { $script:nb_ping = 5 }

        Do { Ping_intra
                if ($intra_ping_ok -ne 100 ) 
                        { 
                        Add_Road
                        Ping_inter
                        Delete_Road
                        }

            Start-Sleep 600
            $nb_ping ++
            
            #VERIF SI TOUJOURS CONNECTE
            $connected = Surv_IP -time_to_wait 1
            if ($connected -eq 3) { $script:nb_ping = 5 }

        } Until ($script:nb_ping -eq 5 )
    } # IF TELETRAVAIL

 } while ($connected -eq 5 )