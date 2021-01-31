<#
.Synopsis
IMPORT SERVICE DANS CONTAINER SSOX
  + GENERATION DU CONTAINER SI INEXISTANT

.Description
Génération automatique des containers SSOX et des Services depuis base DataBase - table dbo.SSOX_CONTAINER_INIT
Génération de log horodaté

Version 1.0 => mise en PROD le 10/09/2019

Version 2.0 => Modification du 18/11/2019
            => Mise en place de l'utilisation des API SOAP Avensys
            => Modification de la prise en compte du mot de passe Admin SSOX : mot de passe crypté dans un fichier

Version 2.1 => Modification du 05/03/2020
            => Correction : ligne 342 - Ajout d'un controle sur le compte AD : suite détection pb de gestion des nouveaux comptes AD

Version 2.2 => Modification du 22/10/2020
            => Correction : certains comptes déjà créés  avec containers existants n'étaient pas traités
                    => modification des controles : vérification de l'existence ou non du container avant vérification  de la propriété SSOXUserK du compte AD
                    => l'ensemble des comptes - en CREATE ou MODIFY - sont désormais traités

.Parameter admin_ssox
Identifiant de l'Admin SSOX


.Parameter envir
Environnement à renseigner : PROD / PPROD / REC

.Example
*Chemin du script*\SSOX_Admin1_IMPORT.ps1 -admin_ssox *identifiant*
SCRIPT : D:\APPDSI\EXPGLB\MAJID\
LOG => G:\LOG\APPDSI\EXPGLB\MAJID\
#>
Param (
    [Parameter(Position=0,Mandatory=$true,HelpMessage='Renseigner votre identifiant')]
    [string]$admin_ssox,
    #[Parameter(Position=1,Mandatory=$true,HelpMessage='Saisissez votre mot de passe')]
    #[String]$password,
    [Parameter(Position=2,Mandatory=$true,HelpMessage='Environnement à préciser : RCT / PPR / EXP')]
    [String]$envir = "RCT"
)

#Code Retour VEGA : 
$return_vega = 0
# $return_vega = 0 => Execution script OK - aucune erreur
# $return_vega = 3 => Execution script KO - manque prérequis - Arret du script
# $return_vega = 5 => Execution script KO - Fichier non trouvé - Arret du script
# $return_vega = 1 => Execution script KO - erreur fatale - Arret du script


Import-Module ActiveDirectory -Cmdlet Get-ADUser
$script:DBNull = [System.DBNull]::Value 
switch -Wildcard ($envir) {
    "RCT" {     $script:Serv_SSOX = "ad.re.ss.ip"
                $script:Base_SQL = "Server\DSIDEVDEVB"
                $script:serv_GRC = "GRC1"
                $script:ScriptRoot = 'G:\LOG\APPDSI\RCTGLB\MAJID\'
                #$script:ScriptRoot = 'C:\temp\'
             }
    "PPR" {     $script:Serv_SSOX = "ad.re.ss.ip"
                $script:Base_SQL = "Server\ASRDPPR"
                $script:serv_GRC = "GRC_PPROD1" 
				$script:ScriptRoot = 'G:\LOG\APPDSI\PPRGLB\MAJID\'	}
    "EXP" {     $script:Serv_SSOX = "ad.re.ss.ip"
                $script:Base_SQL = "Server\ASRDEXP"
                $script:serv_GRC = "GRC1"
                $script:ScriptRoot = 'C:\temp\'
                #$script:ScriptRoot = 'G:\LOG\APPDSI\EXPGLB\MAJID\'
            	}
    Default {throw "Environnement invalide"}
}

Function Decrypt ($SIDFile) {  
    [String] $SID = Get-Content -Path $('C:\temp\' + $SIDFile)
    #[String] $SID = Get-Content -Path $("D:\AppDSI\" + $Envir + "GLB\MAJID\" + $SIDFile)
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($($SID | ConvertTo-SecureString -Key $key))
    $DecryptWord = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Return $DecryptWord
}
Function Log {
[CmdletBinding()]
Param ([Parameter(Mandatory=$true,HelpMessage='Renseigner le contenu à ajouter aux logs')][string]$Add)
$Date = Get-Date -Format 'yyyyMMdd_HH:mm:ss'
$Log_File = "Log_SSOX\$($Date.Split('_')[0])_Log_SSOX_IMPORT.txt"

if (!(Test-path $ScriptRoot\$Log_File)){ New-Item $ScriptRoot\$Log_File -Force -ItemType File
                                            Add-Content $ScriptRoot\$Log_File  -Value "#############################################################################"
                                            Add-Content $ScriptRoot\$Log_File -Value "                        LOGS IMPORT SERVICE CONTAINER SSOX"
                                            Add-Content $ScriptRoot\$Log_File  -Value "#############################################################################"
        }

try { Add-Content $ScriptRoot\$Log_File -Value "[$($Date)] $($Add)" }
catch { Write-Verbose $Error[0] } 
}
function Invoke_sql1 {
param( [string]$sql,
    [System.Data.SQLClient.SQLConnection]$connection
)

$cmd = new-object System.Data.SQLClient.SQLCommand($sql,$connection)
$ds = New-Object system.Data.DataSet
$da = New-Object System.Data.SQLClient.SQLDataAdapter($cmd)
$da.fill($ds) | Out-Null
return $ds.tables[0].DefaultView
}
Function Decrypt_GRC {
param(
[Parameter(Mandatory=$true)][string]$data
)

$key = [System.Text.Encoding]::UTF8.GetBytes("key")
$provider = New-Object System.Security.Cryptography.AesCryptoServiceProvider
$decryptor = $provider.CreateDecryptor($key, $(New-Object byte[] 16))
$encrypted = [System.Convert]::FromBase64String($data)
return [System.Text.Encoding]::UTF8.GetString($decryptor.TransformFinalBlock($encrypted, 0, $encrypted.Length))
}
Function Add_GRC {
try {
Log -Add "$($user.CD_LOGIN) - Lancement Fonction Add_GRC"
Write-Verbose "$($user.CD_LOGIN) - Lancement Fonction Add_GRC"
    if (($user.CD_GRC_LOGIN -ne $null) -and ($user.CD_MOT_DE_PASSE -ne $null)){
        $SSOX_Proxy.AddUserService($user.CD_LOGIN,$script:serv_GRC,$user.CD_GRC_LOGIN,$Password_GRC,"","")
        Log -Add "$($user.CD_LOGIN) - Ajout dans le Container : GRC - $($user.CD_GRC_LOGIN) - $($Password_GRC)"
        Write-Verbose "$($user.CD_LOGIN) - Ajout dans le Container : GRC - $($user.CD_GRC_LOGIN) - $($Password_GRC)"
        $script:ii++
        $script:iii++
        } else {Log -Add "$($user.CD_LOGIN) - GRC - ID et PWD vide dans la table SQL"
                Write-Verbose "$($user.CD_LOGIN) - GRC - ID et PWD vide dans la table SQL"}
} catch { 
	Log -add "$($user.CD_LOGIN) - Plantage Fonction Add_GRC"
	Log -add $error[0]
	Log -add $error	}
#finally { continue }
}
Function Add_SIGMA {
Log -Add "$($user.CD_LOGIN) - Lancement Fonction Add_SIGMA"
Write-Verbose "$($user.CD_LOGIN) - Lancement Fonction Add_SIGMA"
if ($envir -eq "EXP"){    
#    if (($lastlogon.lastlogon -le (Get-Date).AddDays((-45)))) 
#            {
		#	Log -Add "$($user.CD_LOGIN) - lastlogon > 1 mois => SIGMA : Mise à jour Container"
        #    Write-Verbose "$($user.CD_LOGIN) - lastlogon > 1 mois => SIGMA : Mise à jour Container"
        #Controle Ajouté : pour SIGMA, l'accès aux différents ENV se fait avec le même MDP
        # ce controle permet d'éviter d'écraser le mot de passe enregistré dans le container si celui-ci à déjà été changé suite au premier lancement de SIGMA
        # OU (SURTOUT) si le script est déjà passé pour cet utilisateur sur un environnement différent
        #   Si lastlogon est supérieur à 30 jours (ici 45 par sécurité), alors on peut l'écraser sans soucis : l'utilisateur devra de toute façon contacter la TA pour se reconnecter
        # => A faire valider !

            if ($user.CD_SIGMA_LOGIN -ne $null){
                $SSOX_Proxy.AddUserService($user.CD_LOGIN,"SIGMA",$user.CD_SIGMA_LOGIN,$user.CD_SIGMA_LOGIN,"","")
                Log -Add "$($user.CD_LOGIN) - Ajout dans le Container : SIGMA - $($user.CD_SIGMA_LOGIN) - $($user.CD_SIGMA_LOGIN)"
                Write-Verbose "$($user.CD_LOGIN) - Ajout dans le Container : SIGMA - $($user.CD_SIGMA_LOGIN) - $($user.CD_SIGMA_LOGIN)"
                $script:ii++
        
                }else {Log -Add "$($user.CD_LOGIN) - SIGMA - ID vide dans la table SQL"
                    Write-Verbose "$($user.CD_LOGIN) - SIGMA - ID vide dans la table SQL"} 

#    }
#    else {  Log -Add "$($user.CD_LOGIN) - Compte deja connecte dans le mois => SIGMA : Pas de mise à jour Container "
#            Write-Verbose "$($user.CD_LOGIN) - Compte deja connecte dans le mois => SIGMA : Pas de mise à jour Container"}

} #IF ENVIRONNEMENT
Else {Log -Add "$($user.CD_LOGIN) - ENV = $($envir) - SIGMA : Pas de mise à jour Container - MAJ seulement en PROD"
    Write-Verbose "$($user.CD_LOGIN) - ENV = $($envir) - SIGMA : Pas de mise à jour Container - MAJ seulement en PROD"}  
}
Function Update_SQL {
    #$update_time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #$query_up = "UPDATE dbo.SSOX_CONTAINER_INIT SET FL_TRAITE_SSOX = '$($ii+$iii)',DT_PASSAGE_SSOX = '$($update_time)' WHERE CD_LOGIN = '$($user.CD_LOGIN)'"
    #invoke_sql1 $query_up $con
    [int]$action_nb = $ii + $iii

    switch ($action_nb) {
        0 { Log -Add "$($user.CD_LOGIN) - Aucune mise a jour de la table"
            Write-Verbose "$($user.CD_LOGIN) - Aucune mise a jour de la table"}
        1 { Log -Add "$($user.CD_LOGIN) - SIGMA Traite - Mise a jour de la table"
            Write-Verbose "$($user.CD_LOGIN) - SIGMA Traite - Mise a jour de la table"
			    $update_time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				[int]$flag = $user.FL_TRAITE_SSOX.ToString()
				[int]$maj_flag = $flag + $action_nb
				$query_up = "UPDATE dbo.SSOX_CONTAINER_INIT SET FL_TRAITE_SSOX = '$($maj_flag)',DT_PASSAGE_SSOX = '$($update_time)' WHERE CD_LOGIN = '$($user.CD_LOGIN)'"
				invoke_sql1 $query_up $con}
        2 { Log -Add "$($user.CD_LOGIN) - GRC Traité - Mise a jour de la table"
            Write-Verbose "$($user.CD_LOGIN) - GRC Traite - Mise a jour de la table"
				$update_time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				[int]$flag = $user.FL_TRAITE_SSOX.ToString()
				[int]$maj_flag = $flag + $action_nb
				$query_up = "UPDATE dbo.SSOX_CONTAINER_INIT SET FL_TRAITE_SSOX = '$($maj_flag)',DT_PASSAGE_SSOX = '$($update_time)' WHERE CD_LOGIN = '$($user.CD_LOGIN)'"
				invoke_sql1 $query_up $con}
        3 { Log -Add "$($user.CD_LOGIN) - GRC et SIGMA Traite - Mise a jour de la table"
            Write-Verbose "$($user.CD_LOGIN) - GRC et SIGMA Traite - Mise a jour de la table"
				$update_time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				[int]$flag = $user.FL_TRAITE_SSOX.ToString()
				[int]$maj_flag = $flag + $action_nb
				$query_up = "UPDATE dbo.SSOX_CONTAINER_INIT SET FL_TRAITE_SSOX = '$($maj_flag)',DT_PASSAGE_SSOX = '$($update_time)' WHERE CD_LOGIN = '$($user.CD_LOGIN)'"
                invoke_sql1 $query_up $con}    
        9 { Log -Add "$($user.CD_LOGIN) - MaJ du Flag a 9 pour identification"
            Write-Verbose "$($user.CD_LOGIN) - MaJ du Flag a 9 pour identification"
            $update_time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            [int]$maj_flag = 9
            $query_up = "UPDATE dbo.SSOX_CONTAINER_INIT SET FL_TRAITE_SSOX = '$($maj_flag)',DT_PASSAGE_SSOX = '$($update_time)' WHERE CD_LOGIN = '$($user.CD_LOGIN)'"
            invoke_sql1 $query_up $con}        
        Default { Log -Add "$($user.CD_LOGIN) - Incoherence resultat - voir variable : ii+iii = $($action_nb)"
                    Write-Verbose "$($user.CD_LOGIN) - Incoherence resultat - voir variable : ii+iii = $($action_nb)"
                }
    }
<# Pour info : 
	FL_TRAITE_SSOX = NULL = 0 : A traiter
    FL_TRAITE_SSOX = 1 : SIGMA traité
    FL_TRAITE_SSOX = 2 : GRC traité
	FL_TRAITE_SSOX = 3 : GRC et SIGMA traités
#>
    
}
Function Lancement_Fonctions {
#LANCEMENT DES FONCTIONS POUR AJOUT DANS LES CONTAINERS + MAJ SQL
		switch ($user.FL_TRAITE_SSOX) {
			$DBNull {  
                write-verbose "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA et GRC a traiter"
				Log -add "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA et GRC a traiter"
                #Déchiffrement du mot de passe GRC : 
                $script:Password_GRC = Decrypt_GRC -data $user.CD_MOT_DE_PASSE
                Log -Add "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
                Write-Verbose "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
				Add_GRC
				Start-Sleep -Seconds 1
				Add_SIGMA
				Update_SQL
			}
            0 { write-verbose "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA et GRC a traiter"
				Log -add "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA et GRC a traiter"
                #Déchiffrement du mot de passe GRC : 
                $script:Password_GRC = Decrypt_GRC -data $user.CD_MOT_DE_PASSE
                Log -Add "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
                Write-Verbose "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
				Add_GRC
				Start-Sleep -Seconds 1
				Add_SIGMA
				Update_SQL }

            1 { write-verbose "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - GRC a traiter"
				Log -add "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - GRC a traiter"
                #Déchiffrement du mot de passe GRC : 
                $script:Password_GRC = Decrypt_GRC -data $user.CD_MOT_DE_PASSE
                Log -Add "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
                Write-Verbose "$($user.CD_LOGIN) - Dechiffrement du mot de passe GRC"
				Add_GRC
				Update_SQL }

			2 { write-verbose "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA a traiter"
				Log -add "$($user.CD_LOGIN) - FLAG : $($user.FL_TRAITE_SSOX) - SIGMA a traiter"
				Add_SIGMA
				Update_SQL }

			3{ Log -Add "$($user.CD_LOGIN) - SIGMA et GRC deja traite"
                Write-Verbose "$($user.CD_LOGIN) - SIGMA et GRC deja traite"}    
                
            9 {Log -Add "$($user.CD_LOGIN) - Chiffrement AES "
                Write-Verbose "$($user.CD_LOGIN) - Chiffrement AES "}    

			Default { 	Log -Add "$($user.CD_LOGIN) - Incoherence sur la valeur de FL_TRAITE_SSOX => $($user.FL_TRAITE_SSOX)"
						Write-Verbose "$($user.CD_LOGIN) - Incoherence sur la valeur de FL_TRAITE_SSOX => $($user.FL_TRAITE_SSOX)"}
		}

}

try {
    Log -Add "Lancement du script SSOX_IMPORT"
    Write-Verbose "Lancement du script SSOX_IMPORT"

    Log -Add "Connexion table SQL DataBase - dbo.SSOX_CONTAINER_INIT"
    Write-Verbose "Connexion table SQL DataBase - dbo.SSOX_CONTAINER_INIT"

    $con = New-Object System.Data.SqlClient.SqlConnection
    $con.ConnectionString="Server=$script:Base_SQL;Database=DataBase;Trusted_Connection=True;"
    $con.open()

    Log -Add "Recuperation des donnees a traiter"
    Write-Verbose "Recuperation des donnees a traiter"

    $query = "SELECT * FROM dbo.SSOX_CONTAINER_INIT WHERE FL_TRAITE_SSOX <> 3 AND FL_TRAITE_SSOX <> 9 OR FL_TRAITE_SSOX IS NULL"


    $data_DataBase = Invoke_sql1 $query $con
    <#Pour info : 
        FL_TRAITE_SSOX = NULL = 0 : A traiter
        FL_TRAITE_SSOX = 1 : SIGMA traité
        FL_TRAITE_SSOX = 2 : GRC traité
        FL_TRAITE_SSOX = 3 : GRC et SIGMA traités
        FL_TRAITE_SSOX = 9 : Flag passéé à 9 manuellement => suite erreur script
    #>

    #Vérif existence de donnees a traiter dans la table 
Log -Add "Verif existence de donnees a traiter"
Write-Verbose "Verif existence de donnees a traiter"
if ($null -ne $data_DataBase){
    Log -Add "Donnees OK"
    Write-Verbose "Donnees OK"

    ### TEST WEBSERVICE
    Log -Add "Connexion Admin SSOX"
    Write-Verbose "Connexion Admin SSOX"
    $Key = (1..16)
    $SSOX_WebService_Port = "99"
    $SSOX_Serveur_Proxy = "http://" + $script:Serv_SSOX + ":" + $SSOX_WebService_Port + "/service.asmx"
    $SSOX_Proxy = New-WebServiceProxy $SSOX_Serveur_Proxy
    $Authent_Proxy = $SSOX_Proxy.ToString() -replace "\.Service",".CredentialHeader"
    $credv = New-Object ($Authent_Proxy)
    $credv.UserName = $admin_ssox
    $cPassword =  Decrypt password_bssox.txt
    $credv.password = $cPassword
    $SSOX_Proxy.CredentialHeaderValue = $credv
    Log -Add "Connexion Admin SSOX OK"
    Write-Verbose "Connexion Admin SSOX OK"
    ### TEST WEBSERVICE 
    
   	Log -Add "Debut du traitement"
	Write-Verbose "Debut du traitement"
	foreach ($user in $data_DataBase) {
            $return_vega = 0
            $script:ii = $script:iii = 0
            Log -Add "TRAITEMENT : $($user.CD_LOGIN) -  $($user.LB_NOM_COLLABORATEUR) $($user.LB_PRENOM_COLLABORATEUR) - $($user.DT_CREATION_ENREGISTREMENT) - $($user.LB_OPERATION)"
            Write-Verbose "TRAITEMENT : $($user.CD_LOGIN) -  $($user.LB_NOM_COLLABORATEUR) $($user.LB_PRENOM_COLLABORATEUR) - $($user.DT_CREATION_ENREGISTREMENT) - $($user.LB_OPERATION)"
    <#Pour info : 
        FL_TRAITE_SSOX = NULL = 0 : A traiter
        FL_TRAITE_SSOX = 1 : SIGMA traité
        FL_TRAITE_SSOX = 2 : GRC traité
        FL_TRAITE_SSOX = 3 : GRC et SIGMA traités
    #>
                try { 
                    #Verification de l'existence du compte AD :
                    $user_AD = Get-ADUser $user.CD_LOGIN -properties *
                    ### WEBSERVICE SSOX
                    $is_SSOX_User = $SSOX_Proxy.isSSOXUser($user.CD_LOGIN)
                    $SSOX_User = $is_SSOX_User.FirstChild.Attributes["ISSSOX"].Value
                    ### WEBSERVICE SSOX
                    
                    #Verification de la présence d'un container pour l'utilisateur
                    Log -Add "$($user.CD_LOGIN) - Verification de la presence d'un container pour l'utilisateur"
                    Write-Verbose "$($user.CD_LOGIN) - Verification de la presence d'un container pour l'utilisateur"
                    if ($SSOX_User -eq "yes") { Log -Add "$($user.CD_LOGIN) - Container SSOX existant"
                                Write-Verbose "$($user.CD_LOGIN) - Container SSOX existant"


                                if (($null -ne $user_AD.SSOXUserk) -and ([System.Text.Encoding]::UTF8.GetString($user_AD.SSOXUserk) -match "KeyAES:KeyAES") ) {

                                                Log -Add "$($user.CD_LOGIN) - Verification de l'existence du compte AD : OK"
                                                Write-Verbose "$($user.CD_LOGIN) - Verification de l'existence du compte AD : OK"
                                                Log -Add "$($user.CD_LOGIN) - Compte avec chiffrement AES active - passage du script annule pour ce compte"
                                                Write-Verbose "$($user.CD_LOGIN) - Compte avec chiffrement AES active - passage du script annule pour ce compte"
                                                
                                                $script:ii = 9
                                                Update_SQL
                                    } # IF KEY AES - PAS DE TRAITEMENT
                                    else {
                                        Log -Add "$($user.CD_LOGIN) - Verification de l'existence du compte AD : OK"
                                        Write-Verbose "$($user.CD_LOGIN) - Verification de l'existence du compte AD : OK"
                                        Log -Add "$($user.CD_LOGIN) - Compte SANS chiffrement AES active - passage du script pour ce compte"
                                        Write-Verbose "$($user.CD_LOGIN) - Compte SANS chiffrement AES active - passage du script pour ce compte"
                                    
                                        Lancement_Fonctions
                                    } # ELSE - PAS AES - TRAITEMENT OK
                                } # IF CONTAINER EXISTE
                                                        
                    else {Log -Add "$($user.CD_LOGIN) - Container SSOX n'existe pas - Creation du Container"
                            Write-Verbose "$($user.CD_LOGIN) - Container SSOX n'existe pas - Creation du Container"
                            $SSOX_Proxy.CreateUserContainer($user_AD.SamAccountName,"")
                            #VERIF CREATION OK
                            #Start-Sleep -seconds 5
                            $is_SSOX_User = $SSOX_Proxy.isSSOXUser($user.CD_LOGIN)
                            $SSOX_User = $is_SSOX_User.FirstChild.Attributes["ISSSOX"].Value
                            if ($SSOX_User -eq "yes") {Log -Add "$($user.CD_LOGIN) - CONTAINER SSOX Genere avec SUCCES !"
                                                                Write-Verbose "$($user.CD_LOGIN) - CONTAINER SSOX Genere avec SUCCES !"
                                                                $user_AD = Get-ADUser $user.CD_LOGIN -properties *
                                                                Lancement_Fonctions  }
                            else { Log -Add "$($user.CD_LOGIN) - GENERATION CONTAINER SSOX KO"
                                Write-Verbose "$($user.CD_LOGIN) - GENERATION CONTAINER SSOX KO"
                                }
                    } # ELSE - CREATION CONTAINER SSOX
                } # TRY
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { 
                        Write-Verbose "$($user.CD_LOGIN) - Utilisateur non trouve dans l'AD - Ligne suivante !"
                        Log -Add "$($user.CD_LOGIN) - Utilisateur non trouve dans l'AD - Ligne suivante !"
                        $return_vega = 3 }
                catch {Write-Verbose $Error[0]
                        $return_vega = 1
                        Log -Add $return_vega
                        Log -Add $Error[0] }
                
        } # FOREACH DATA
                Log -Add "Fin du script !"
                Write-Verbose "Fin du script !"
    } # IF DATA OK
    Else {  Write-Verbose "Aucune donnee a traiter" 
            $return_vega = 3
            Log -Add "Aucune donnee a traiter"}

}
Catch {Write-Verbose $Error[0]
        $return_vega = 1
        Log -Add $return_vega
        Log -Add $Error[0] }

Finally { $return_vega }