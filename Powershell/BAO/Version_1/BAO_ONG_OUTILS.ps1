###-----------------------------------------------------------------------------------------------###
### ONGLET OUTILS
###-----------------------------------------------------------------------------------------------###

#---------------------------------------------------------------------------------------------------#
# Chargement des librairies et définition des variables associées
#---------------------------------------------------------------------------------------------------#
#C:\AppDSI\DEVGLB\OUTILSTA\BAO\libs\libSQL.ps1
. .\libs\libSQL.ps1
#---------------------------------------------------------------------------------------------------#
# Déclarations des Fonctions
#---------------------------------------------------------------------------------------------------#


Function Recherche {
Param ( [string]$Netbios  )

Write-Host $Netbios1

echl "[ii]`t Requete IGUAZU sur $Netbios1"

$SQUERY = "
	SELECT 
	TRIGRAMME,
	FIRSTNAME,
	LASTNAME,
	ID_RESSOURCE,
	SERVICE,
	ADRESSE1,
	ADRESSE2,
	CP,
	VILLE,
	TELEPHONE,
	PORTABLE,
	EMAIL,
	MATRICULE_RH,
	LOGIN_SIGMA,
	LOGIN_GRC
	FROM PERSONNES
	INNER JOIN RESSOURCE ON RESSOURCE.SID_PERS = PERSONNES.SID_PERS
	INNER JOIN SITEENTREPRISE on SITEENTREPRISE.ID_SITEENTREPRISE = PERSONNES.ID_SITEENTREPRISE
	where TRIGRAMME = '$Netbios1' OR ID_RESSOURCE = '$Netbios1' 
"
 
$global:Rech_Netbios = SQLExecuteReader $SQUERY 
#$global:Result = $Rech_Netbios.syncroot[5]

$Global:Result = $Rech_Netbios | 
			Select-Object  TRIGRAMME,
							FIRSTNAME,
							LASTNAME,
							ID_RESSOURCE,
							SERVICE,
							ADRESSE1,
							ADRESSE2,
							CP,
							VILLE,
							TELEPHONE,
							PORTABLE,
							EMAIL,
							MATRICULE_RH,
							LOGIN_SIGMA,
							LOGIN_GRC -last 1
						
$global:result1 = $result | Get-Unique
$lbl_poste_infos_netbios.Text = "Information sur le poste $($Rech_Netbios.ID_RESSOURCE) :"
}

Function PrePDM { 
$cbx_outils_verifntb.Items.Clear()

If ($Rech_Netbios.ID_RESSOURCE.count -gt 1){
		foreach($computer in $Rech_Netbios.ID_RESSOURCE)
		{   $cbx_outils_verifntb.Items.add($computer)
			$cbx_poste_verifposte.Items.add($computer)
			$cbx_genesys_delog_poste.Items.add($computer)}
		$cbx_outils_verifntb.Items.Add($cbx_outils_verifntb.Text)
		$cbx_outils_verifntb.visible = $true
}
		
Else {
$cbx_outils_verifntb.visible = $false
}
$cbx_outils_verifntb.refresh()

}

Function MSRA {
If ($Rech_Netbios.ID_RESSOURCE.count -gt 1){
msra.exe /offerra $cbx_outils_verifntb.Text }
Else { msra.exe /offerra $Rech_Netbios.ID_RESSOURCE }
 }

Function Remote {
If ($Rech_Netbios.ID_RESSOURCE.count -gt 1){
$GLOBAL:Netbios1 =  $cbx_outils_verifntb.Text }
Else { $GLOBAL:Netbios1 =  $Rech_Netbios.ID_RESSOURCE }

If ([IntPtr]::Size -eq 8){start "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$Netbios1}
Else {start "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$Netbios1}
}


Function ClipBoard {
"$($rtxtb_bao_info.text)",
"Etat du compte : $($lbl_ad_verifad.text)",
"$($lbl_ad_agepassword.text)",
"$($lbl_genesys_num.text)", 
"",
"IP : $($lbl_poste_infos_ipres.text)",
"Dernier Demarrage : $($lbl_poste_infos_bootres.text)",
"Titulaire SCCM : $($lbl_poste_infos_titsccmres.text)",
"Poste associe SCCM : $($lbl_poste_infos_posccmres.text)",
"Situation : $($lbl_poste_infos_sitres.text)",
"IP STA : $($lbl_poste_infos_ipstares.text)",
"Etat STA : $($lbl_poste_infos_etstares.text)",
"",
"Infos Materiel :",
"Modele : $($lbl_poste_infos_modelres.text)",
"OS : $($lbl_poste_infos_osres.text)",
"Archi : $($lbl_poste_infos_archires.text)",
"DD : $($lbl_poste_infos_sizeres.text)",
"Espace libre : $($lbl_poste_infos_freespaceres.text)",
"Etat Disque : $($lbl_poste_infos_diskstatres.text)",
"RAM : $($lbl_poste_infos_ramres.text)" | clip 

#$lbl_genesys_trig.text | clip 

}

##
## REINITIALISATION SIGMA
##

Function SIGMA {
#Start ".\Scripts\SIGMA_MDP\SIGMA_MDP.exe" $result1.LOGIN_SIGMA

Start ".\Scripts\SIGMA_MDP\SIGMA_MDP.exe" $txtb_outils_sigma.text
}

##
## LIENS
##

Function Wiki {
start "http://"
}

Function GRC {
start "https://"
}

Function DSI {
start "http://"
}

Function Phareouest {
start "http://phareouest"
}

Function Guidouest {
start "http://"
}

Function SMART {
start "https://"
}

Function RICOH {
Start "https://"
}

Function CANON {
Start "https://"
}

Function WifiGuest {
Start "https://"
}

Function Iguazu {
Start "http://$($result1.TRIGRAMME)"
}

Function ServicePilot {
start "http://"
}



##
## PURGE VARIABLES
##

Function Purge_var {
Clear-Variable Netbios1 -ErrorAction SilentlyContinue
Clear-Variable Trigramme -ErrorAction SilentlyContinue
Clear-Variable Num_tel -Scope global -ErrorAction SilentlyContinue
Clear-Variable retourGen -Scope global -ErrorAction SilentlyContinue
Clear-Variable NumNonGen -ErrorAction SilentlyContinue
Clear-Variable Computer -ErrorAction SilentlyContinue
 

$rtxtb_bao_info.text = ""
$txtb_ad_trig.text = ""

$lbl_poste_infos_ramres.text = "-"
$lbl_poste_infos_diskstatres.text = "-"
$lbl_poste_infos_freespaceres.text = "-"
$lbl_poste_infos_sizeres.text = "-"
$lbl_poste_infos_archires.text = "-"
$lbl_poste_infos_osres.text = "-"
$lbl_poste_infos_modelres.text = "-"
$lbl_poste_infos_etstares.text = "-"
$lbl_poste_infos_ipstares.text = "-"
$lbl_poste_infos_sitres.text = "-"
$lbl_poste_infos_posccmres.text = "-"
$lbl_poste_infos_titsccmres.text = "-"
$lbl_poste_infos_bootres.text = "-"
$lbl_poste_infos_ipres.text = "-"

$lbl_ad_verifad.text = ""

$rtxtb_bao_act.text = ""
$rtxtb_bao_info.text = ""

$dtgv_poste_proc_proc.DataSource = $null
$dtgv_poste_serv_serv.DataSource = $null
$dtgv_genesys_histut_hist.DataSource = $null
$dtgv_genesys_histnum_hist.DataSource = $null

$txtb_genesys_delog_place.text = ""
$txtb_genesys_delog_trig.text = ""
#$txtb_genesys_delog_poste.text = ""
$txtb_genesys_rechnum_trig.text = ""
$txtb_genesys_rechnum_num.text = ""
$txtb_outils_sigma.text = ""

$lbl_genesys_rechnum_trigag.text = ""
$lbl_genesys_rechnum_numnom.text = ""
$lbl_genesys_rechnum_numtrig.text = ""
#$lbl_genesys_delog_poste.text = ""
$lbl_genesys_num.text = ""
$lbl_genesys_trig.text = ""

$cbx_genesys_delog_poste.Items.Clear()
$cbx_genesys_delog_poste.text = ""
$cbx_poste_verifposte.Items.Clear()
$cbx_poste_verifposte.text = ""
}

Function Purge_label {
$lbl_poste_infos_ramres.text = "-"
$lbl_poste_infos_diskstatres.text = "-"
$lbl_poste_infos_freespaceres.text = "-"
$lbl_poste_infos_sizeres.text = "-"
$lbl_poste_infos_archires.text = "-"
$lbl_poste_infos_osres.text = "-"
$lbl_poste_infos_modelres.text = "-"
$lbl_poste_infos_etstares.text = "-"
$lbl_poste_infos_ipstares.text = "-"
$lbl_poste_infos_sitres.text = "-"
$lbl_poste_infos_posccmres.text = "-"
$lbl_poste_infos_titsccmres.text = "-"
$lbl_poste_infos_bootres.text = "-"
$lbl_poste_infos_ipres.text = "-"

$dtgv_poste_proc_proc.DataSource = $null
$dtgv_poste_serv_serv.DataSource = $null
}


##
## TEXTBOX ACTIONS
##

Function AutoScroll {
$rtxtb_bao_act.SelectionStart = $rtxtb_bao_act.TextLength;
$rtxtb_bao_act.ScrollToCaret()
}
