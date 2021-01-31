###-----------------------------------------------------------------------------------------------###
### ONGLET GENESYS
###-----------------------------------------------------------------------------------------------###

Function hist_UT{
##--------------------------------------------------------------------------------------------
## Fonction historique Utilisateur
##--------------------------------------------------------------------------------------------

#ParamÃ¨tres de connexion.

$GLOBAL:SQLSERVER = "SERVER\GIMEXP"  	
$GLOBAL:SQLBASE = "GEN_GIM"             	
$GLOBAL:SqlQuery = "USE GEN_GIM
DECLARE @Trigramme varchar(50)
SET @Trigramme = '$Trigramme'
SELECT
      GIDB_G_LOGIN_SESSION_V.ID,
      GIDB_GC_AGENT.EMPLOYEEID Trigramme,
      GIDB_GC_AGENT.FIRSTNAME Prenom,
      GIDB_GC_AGENT.LASTNAME Nom,
      GIDB_GC_PLACE.NAME Place,
      GIDB_GC_ENDPOINT.DN Telephone_Associe,
      GIDB_GC_FOLDER.NAME Localisation,
      GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
      GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_AGENT.EMPLOYEEID=@Trigramme)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC"


$retour1 = Sql_QueryTel $SQLSERVER $SQLBASE $SqlQuery

$array_hist_UT = New-Object System.Collections.ArrayList
$Global:Hist_UT =  $retour1 | select Trigramme,Nom,Prenom, Place,@{n='Numero';e={$_.Telephone_Associe}},DebutDeSession,FinDeSession

$array_hist_UT.AddRange($Hist_UT) 
$dtgv_genesys_histut_hist.DataSource = $array_hist_UT 

echl "[ii]`tConsultation Historique Genesys de $Trigramme"
}

Function Hist_No{
##--------------------------------------------------------------------------------------------
## Fonction historique NumÃ©ro de tÃ©l
##--------------------------------------------------------------------------------------------

$Global:Sqlquery2 = "USE GEN_GIM
DECLARE @TelephoneAssocie varchar(50)
SET @TelephoneAssocie = '$num_tel'
SELECT
      GIDB_G_LOGIN_SESSION_V.ID,
      GIDB_GC_AGENT.EMPLOYEEID Trigramme,
      GIDB_GC_AGENT.FIRSTNAME Prenom,
      GIDB_GC_AGENT.LASTNAME Nom,
      GIDB_GC_PLACE.NAME Place,
      GIDB_GC_ENDPOINT.DN Telephone_Associe,
      GIDB_GC_FOLDER.NAME Localisation,
      GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
      GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_ENDPOINT.DN=@TelephoneAssocie)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC"  
 
$retour2 = Sql_QueryTel $SQLSERVER $SQLBASE $SqlQuery2

$array_Hist_No = New-Object System.Collections.ArrayList
$Global:Hist_No =  $retour2 | select Trigramme,Nom,Prenom, Place,@{n='Numero';e={$_.Telephone_Associe}},DebutDeSession,FinDeSession

$array_Hist_No.AddRange($Hist_No) 
$dtgv_genesys_histnum_hist.DataSource = $array_Hist_No

echl "[ii]`tConsultation Historique Genesys du numéro $num_tel"

}

Function Num_associe {
##--------------------------------------------------------------------------------------------
## Fonction de rÃ©cupÃ©ration du numÃ©roi associÃ© Ã  un trigrmame
##--------------------------------------------------------------------------------------------
$GLOBAL:SQLSERVER = "SERVER\GIMEXP"  	
$GLOBAL:SQLBASE = "GEN_GIM"             	
$GLOBAL:SqlQuery = "USE GEN_GIM
DECLARE @Trigramme varchar(50)
SET @Trigramme = '$Trigramme'
SELECT
      GIDB_G_LOGIN_SESSION_V.ID,
      GIDB_GC_AGENT.EMPLOYEEID Trigramme,
      GIDB_GC_AGENT.FIRSTNAME Prenom,
      GIDB_GC_AGENT.LASTNAME Nom,
      GIDB_GC_PLACE.NAME Place,
      GIDB_GC_ENDPOINT.DN Telephone_Associe,
      GIDB_GC_FOLDER.NAME Localisation,
      GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
      GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_AGENT.EMPLOYEEID=@Trigramme)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC"

$Global:retourGen = Sql_QueryTel $SQLSERVER $SQLBASE $SqlQuery


if ($retourGen -eq 0) {$GLOBAL:NumNonGen = Get-ADUser -Filter {samAccountName -eq $Trigramme } -Properties * | Select-Object -ExpandProperty HomePhone
	if ($NumNonGen -eq $null){  $lbl_genesys_num.text = "Téléphone non connu"
								echl "[ii]`tLe numéro de $Trigramme est non connu" }
	else{ $lbl_genesys_num.text =  "Tél par défaut : $NumNonGen"
		$txtb_genesys_delog_place.text = ""
		echl "[ii]`tLe numéro de $Trigramme est $NumNonGen"	}
}
Else {
	If ($retourGen.Telephone_Associe.count -ne 1 ){
	$GLOBAL:num_tel = $retourGen.Telephone_Associe[0] }
	Else { $GLOBAL:num_tel = $retourGen.Telephone_Associe }
		$lbl_genesys_num.text = "Numéro GENESYS : $num_tel "
		echl "[ii]`tLe numéro de $Trigramme est $num_tel" 
		Hist_UT
		Hist_No			

$txtb_genesys_rechnum_trig.text = $Trigramme
$txtb_genesys_delog_trig.text = $Trigramme
$txtb_genesys_rechnum_num.text = $num_tel
$txtb_genesys_delog_place.text = $num_tel
If ($Rech_Netbios.ID_RESSOURCE.count -eq 1) {$cbx_genesys_delog_poste.text = $Rech_Netbios.ID_RESSOURCE	}
	Else { 	$cbx_genesys_delog_poste.text = $Rech_Netbios.ID_RESSOURCE[0]	} 

}
$lbl_genesys_trig.text = "Trigramme : $Trigramme "
}

Function Delog_Bandeau {
$cbx_genesys_delog_poste.Items.Add($cbx_genesys_delog_poste.Text)

if ($rdbtn_genesys_delog_site.Checked -eq $true) {

$newplace = $cbx_genesys_delog_poste.Text
}
if ($rdbtn_genesys_delog_ag.Checked -eq $true) {
$newplace = $txtb_genesys_delog_place.text

}

# VARIABLES
$newuser = $txtb_genesys_rechnum_trig.text
$newuser = $newuser.tolower()

$Path = ".\BANDEAUGENESYS\SITE\conf\groupamaPhone.properties"

# Effacer la place existante ainsi que le user dans le fichier groupamaPhone.properties 
$place = Get-Content -Path $PATH | where {$_ -notmatch "bandeau.genesys.place"}
Set-content $PATH $place
$user = Get-Content -Path $PATH | where {$_ -notmatch "bandeau.genesys.loginAgent"}
Set-content $PATH  $user

# Inscription des nouvelles données
Add-content $PATH "bandeau.genesys.place=$newplace"
Add-content $PATH "bandeau.genesys.loginAgent=$newuser"

Timeout /T 1

# Lancement du bandeau
start-process -FilePath ".\BANDEAUGENESYS\SITE\groupamaPhone.jar"

}