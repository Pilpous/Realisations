<#
 
.Synopsis

.Description

Le niveau du mode verbose se défini en fonction du nombre de "v" ou "d"
 vide => [!!], [WW]
 v    => [!!], [WW], [OK], [TT]
 vv   => [!!], [WW], [OK], [TT], [ii]
 d    => [!!], [WW], [OK], [TT], [ii], [DBG]

.Example

#>

#---------------------------------------------------------------------------------------------------#
# Lecture des arguments
#---------------------------------------------------------------------------------------------------#
param 
(
	[switch] $noLastLog, 
	[switch] $v,
	[switch] $vv,
	[switch] $d
	
)

#---------------------------------------------------------------------------------------------------#
# Chargement des librairies et définition des variables associées
#---------------------------------------------------------------------------------------------------#

. .\libs\libLog.ps1
. .\libs\libCred.ps1
. .\libs\libSQL.ps1

. .\BAO_ONG_OUTILS.ps1
. .\BAO_ONG_AD.ps1
. .\BAO_ONG_POSTE.ps1
. .\BAO_ONG_GENESYS.ps1
. .\BAO_ONG_HISTO.ps1


#Set-Location $location
Get-ChildItem .\log | Remove-Item -ErrorAction SilentlyContinue

# ---------------------------------------------------------
# Variables du script
# ---------------------------------------------------------
[String] $global:NomProc		= $MyInvocation.MyCommand.Name
[String] $global:Param  		= $MyInvocation.BoundParameters.getenumerator() |ForEach-Object { Write-Output "$($_.key)=$($_.value) " }
[String] $global:Version_no		= '1.0'
[String] $global:Version_date	= '14/09/2016'
[switch] $global:lastLogEnabled	= $noLastLog



function GenerateForm {
#region Import the Assemblies
[Void] [reflection.assembly]::loadwithpartialname("System.Windows.Forms") 
[Void] [reflection.assembly]::loadwithpartialname("System.Drawing") 
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

$PSDefaultParameterValues['*:Encoding'] = 'utf8'
#region Generated Form Objects
$BAO = New-Object System.Windows.Forms.Form
$tbct_bao_ongl = New-Object System.Windows.Forms.TabControl
$gbx_outils_liens = New-Object System.Windows.Forms.GroupBox
$cbx_outils_verifntb = New-Object System.Windows.Forms.ComboBox

$cbx_poste_verifposte = New-Object System.Windows.Forms.ComboBox
$cbx_genesys_delog_poste = New-Object System.Windows.Forms.ComboBox

#$rbtn_outils_verifntb = New-Object System.Windows.Forms.RadioButton
$rdbtn_genesys_delog_site = New-Object System.Windows.Forms.RadioButton
$rdbtn_genesys_delog_ag = New-Object System.Windows.Forms.RadioButton

$lbl_bao_tp = New-Object System.Windows.Forms.Label
$lbl_ad_reinitinfo = New-Object System.Windows.Forms.Label
$lbl_ad_reinit = New-Object System.Windows.Forms.Label
$lbl_ad_verifad = New-Object System.Windows.Forms.Label
$lbl_ad_agepassword = New-Object System.Windows.Forms.Label
$lbl_hist_sugg = New-Object System.Windows.Forms.Label


$lbl_poste_infos_ramres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_ram = New-Object System.Windows.Forms.Label
$lbl_poste_infos_diskstatres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_diskstat = New-Object System.Windows.Forms.Label
$lbl_poste_infos_freespaceres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_freespace = New-Object System.Windows.Forms.Label
$lbl_poste_infos_sizeres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_size = New-Object System.Windows.Forms.Label
$lbl_poste_infos_archires = New-Object System.Windows.Forms.Label
$lbl_poste_infos_archi = New-Object System.Windows.Forms.Label
$lbl_poste_infos_osres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_os = New-Object System.Windows.Forms.Label
$lbl_poste_infos_modelres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_model = New-Object System.Windows.Forms.Label
$lbl_poste_infos_etstares = New-Object System.Windows.Forms.Label
$lbl_poste_infos_ipstares = New-Object System.Windows.Forms.Label
$lbl_poste_infos_sitres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_posccmres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_titsccmres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_bootres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_ipres = New-Object System.Windows.Forms.Label
#$lbl_poste_infos_netbiosres = New-Object System.Windows.Forms.Label
$lbl_poste_infos_infsys = New-Object System.Windows.Forms.Label
$lbl_poste_infos_etsta = New-Object System.Windows.Forms.Label
$lbl_poste_infos_ipsta = New-Object System.Windows.Forms.Label
$lbl_poste_infos_sit = New-Object System.Windows.Forms.Label
$lbl_poste_infos_posccm = New-Object System.Windows.Forms.Label
$lbl_poste_infos_titsccm = New-Object System.Windows.Forms.Label
$lbl_poste_infos_boot = New-Object System.Windows.Forms.Label
$lbl_poste_infos_ip = New-Object System.Windows.Forms.Label
$lbl_poste_infos_netbios = New-Object System.Windows.Forms.Label
$lbl_genesys_delog_place = New-Object System.Windows.Forms.Label
$lbl_genesys_delog_trig = New-Object System.Windows.Forms.Label
$lbl_genesys_delog_poste = New-Object System.Windows.Forms.Label
$lbl_genesys_rechnum_trig = New-Object System.Windows.Forms.Label
$lbl_genesys_rechnum_num = New-Object System.Windows.Forms.Label
$lbl_genesys_num = New-Object System.Windows.Forms.Label
$lbl_genesys_trig = New-Object System.Windows.Forms.Label
$lbl_genesys_rechnum_trigag = New-Object System.Windows.Forms.Label
#$lbl_genesys_rechnum_trignum = New-Object System.Windows.Forms.Label
$lbl_genesys_rechnum_numnom = New-Object System.Windows.Forms.Label
$lbl_genesys_rechnum_numtrig = New-Object System.Windows.Forms.Label
$lbl_hist_histo = New-Object System.Windows.Forms.Label

$btn_ad_consad = New-Object System.Windows.Forms.Button
$btn_ad_dev = New-Object System.Windows.Forms.Button
$btn_ad_verif = New-Object System.Windows.Forms.Button
$btn_outils_wg = New-Object System.Windows.Forms.Button
$btn_outils_ric = New-Object System.Windows.Forms.Button
$btn_outils_go = New-Object System.Windows.Forms.Button
$btn_outils_grc = New-Object System.Windows.Forms.Button
$btn_outils_smart = New-Object System.Windows.Forms.Button
$btn_outils_sp = New-Object System.Windows.Forms.Button
$btn_outils_igua = New-Object System.Windows.Forms.Button
$btn_outils_po = New-Object System.Windows.Forms.Button
$btn_outils_dsi = New-Object System.Windows.Forms.Button
$btn_outils_wiki = New-Object System.Windows.Forms.Button
$btn_outils_sigma = New-Object System.Windows.Forms.Button
$btn_outils_remote = New-Object System.Windows.Forms.Button
$btn_outils_pdm = New-Object System.Windows.Forms.Button
$btn_ad_reinit = New-Object System.Windows.Forms.Button
$btn_poste_infos_pdm = New-Object System.Windows.Forms.Button
$btn_poste_infos_remote = New-Object System.Windows.Forms.Button
$btn_poste_infos_num = New-Object System.Windows.Forms.Button
$btn_poste_infos_rebsta = New-Object System.Windows.Forms.Button
$btn_poste_infos_pdmsta = New-Object System.Windows.Forms.Button
$btn_poste_sccm_copy = New-Object System.Windows.Forms.Button
$btn_poste_sccm_rech = New-Object System.Windows.Forms.Button
$btn_poste_proc_kill = New-Object System.Windows.Forms.Button
$btn_poste_proc_ref = New-Object System.Windows.Forms.Button
$btn_poste_serv_arrel = New-Object System.Windows.Forms.Button
$btn_poste_serv_ref = New-Object System.Windows.Forms.Button
$btn_poste_rech = New-Object System.Windows.Forms.Button
$btn_genesys_delog_lancer = New-Object System.Windows.Forms.Button
$btn_genesys_rechnum_trig = New-Object System.Windows.Forms.Button
$btn_genesys_rechnum_num = New-Object System.Windows.Forms.Button
$btn_hist_sugg = New-Object System.Windows.Forms.Button
$btn_bao_rech = New-Object System.Windows.Forms.Button
$btn_poste_infos_conssccm = New-Object System.Windows.Forms.Button
$btn_poste_infos_repsccm = New-Object System.Windows.Forms.Button
$btn_poste_infos_cflog = New-Object System.Windows.Forms.Button
$btn_outils_clip = New-Object System.Windows.Forms.Button

$ong_bao_outils = New-Object System.Windows.Forms.TabPage
$ong_bao_ad = New-Object System.Windows.Forms.TabPage
$ong_bao_poste = New-Object System.Windows.Forms.TabPage
$ong_bao_poste_infos = New-Object System.Windows.Forms.TabPage
#$ong_bao_poste_sccm = New-Object System.Windows.Forms.TabPage
$ong_bao_poste_proc = New-Object System.Windows.Forms.TabPage
$ong_bao_poste_serv = New-Object System.Windows.Forms.TabPage
$ong_genesys_histut = New-Object System.Windows.Forms.TabPage
$ong_genesys_histnum = New-Object System.Windows.Forms.TabPage
$ong_genesys_delog = New-Object System.Windows.Forms.TabPage
$ong_genesys_rechnum = New-Object System.Windows.Forms.TabPage
$ong_bao_hist = New-Object System.Windows.Forms.TabPage
$tabPage4 = New-Object System.Windows.Forms.TabPage

$tabControl2 = New-Object System.Windows.Forms.TabControl
$tabControl1 = New-Object System.Windows.Forms.TabControl

$txtb_outils_verifntb = New-Object System.Windows.Forms.TextBox
$txtb_ad_mdp = New-Object System.Windows.Forms.TextBox
$txtb_ad_trig = New-Object System.Windows.Forms.TextBox
$txtb_poste_sccm_rech = New-Object System.Windows.Forms.TextBox
$txtb_poste_netbios = New-Object System.Windows.Forms.TextBox
$txtb_genesys_delog_place = New-Object System.Windows.Forms.TextBox
$txtb_genesys_delog_trig = New-Object System.Windows.Forms.TextBox
#$txtb_genesys_delog_poste = New-Object System.Windows.Forms.TextBox
$txtb_genesys_rechnum_trig = New-Object System.Windows.Forms.TextBox
$txtb_genesys_rechnum_num = New-Object System.Windows.Forms.TextBox
$txtb_bao_tp = New-Object System.Windows.Forms.TextBox
$txtb_outils_sigma = New-Object System.Windows.Forms.TextBox

#$rtxtb_poste_infos_mat = New-Object System.Windows.Forms.RichTextBox
#$rtxtb_hist_sugg = New-Object System.Windows.Forms.RichTextBox
$rtxtb_hist_histo = New-Object System.Windows.Forms.RichTextBox
$rtxtb_bao_act = New-Object System.Windows.Forms.RichTextBox
$rtxtb_bao_info = New-Object System.Windows.Forms.RichTextBox

#$dtgv_poste_sccm_result = New-Object System.Windows.Forms.DataGridView
$dtgv_poste_proc_proc = New-Object System.Windows.Forms.DataGridView
$dtgv_poste_serv_serv = New-Object System.Windows.Forms.DataGridView
$dtgv_genesys_histut_hist = New-Object System.Windows.Forms.DataGridView
$dtgv_genesys_histnum_hist = New-Object System.Windows.Forms.DataGridView

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects


#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#
#$handler_button1_Click = { }
#$handler_button9_Click = { }

##
## Bouton Rechercher
##

$btn_bao_rech_OnClick = {
Purge_var
$cbx_outils_verifntb.Text = ""
$cbx_poste_verifposte.Items.Clear()
$lbl_ad_reinitinfo.text = ""
$txtb_ad_mdp.text = "Password"

		$global:Netbios1 = $txtb_bao_tp.get_text() -replace '\s',''
		Recherche -Netbios $Netbios1
		echl "[ii]`tLe poste $($Rech_Netbios.ID_RESSOURCE) appartient a $($Rech_Netbios.TRIGRAMME)"
		$rtxtb_bao_info.text = "Informations EasyVista "
		$rtxtb_bao_info.appendtext("`r------------------------------------------")

				If ($result1) {
						$rtxtb_bao_info.appendtext("`nNom: $($result1.FIRSTNAME) $($result1.LASTNAME)")
						$rtxtb_bao_info.appendtext("`nTrigramme: $($result1.TRIGRAMME)")
						$rtxtb_bao_info.appendtext("`nPoste: $($Rech_Netbios.ID_RESSOURCE)")
						$rtxtb_bao_info.appendtext("`nService: $($result1.SERVICE)")
						$rtxtb_bao_info.appendtext("`n")
						$rtxtb_bao_info.appendtext("`nAdresse: $($result1.ADRESSE1) `n$($result1.ADRESSE2)")
						$rtxtb_bao_info.appendtext("`n$($result1.CP) $($result1.VILLE)")
						$rtxtb_bao_info.appendtext("`n")
						$rtxtb_bao_info.appendtext("`nTéléphone: $($result1.TELEPHONE)")
						$rtxtb_bao_info.appendtext("`nPortable: $($result1.PORTABLE)")
						$rtxtb_bao_info.appendtext("`nEmail: $($result1.EMAIL)")
						$rtxtb_bao_info.appendtext("`n")
						$rtxtb_bao_info.appendtext("`nMatricule: $($result1.MATRICULE_RH)")
						$rtxtb_bao_info.appendtext("`nSigma: $($result1.LOGIN_SIGMA)")
						$rtxtb_bao_info.appendtext("`nGRC: $($result1.LOGIN_GRC)")
						$rtxtb_bao_info.appendtext("`n")
#						$rtxtb_bao_info.appendtext("`nModèle PC: $($result1.ID_MODELE)")
#						$rtxtb_bao_info.appendtext("`nVersion OS: $($result1.ID_VERSION_OS)")
#						$rtxtb_bao_info.appendtext("`nEtat: $($result1.FICHE_DESC)")
						PrePDM
						$cbx_outils_verifntb.Text = $($Rech_Netbios.ID_RESSOURCE[0])
						$cbx_poste_verifposte.Text = $($Rech_Netbios.ID_RESSOURCE[0])
						$txtb_ad_trig.text = $result1.TRIGRAMME
						$txtb_outils_sigma.visible = $true
						$txtb_outils_sigma.text = $($result1.LOGIN_SIGMA) 
						#$btn_outils_sigma = $($result1.LOGIN_SIGMA)
						$Global:Trigramme = $result1.TRIGRAMME
							VerifCompteWind
							AgeMDP

							
							If ($Rech_Netbios.ID_RESSOURCE.count -eq 1) {
							$cbx_poste_verifposte.text = $Rech_Netbios.ID_RESSOURCE
#							$cbx_genesys_delog_poste.text = $Rech_Netbios.ID_RESSOURCE
#							$global:computer = $Rech_Netbios.ID_RESSOURCE
#							fastping -computer $Rech_Netbios.ID_RESSOURCE 
							}
							Else { 
							$cbx_poste_verifposte.text = $Rech_Netbios.ID_RESSOURCE[0]
#							$cbx_genesys_delog_poste.text = $Rech_Netbios.ID_RESSOURCE[0]
#							$global:computer = $Rech_Netbios.ID_RESSOURCE[0]
#							fastping -computer $Rech_Netbios.ID_RESSOURCE[0]
							} 
							
							Num_associe
#							hist_UT
#							Hist_No						
							Historique
							$rtxtb_bao_act.text = "Recherche sur $Netbios1"
						}
				Else {
						$rtxtb_bao_info.appendtext("`nPoste ou utilisateur non trouvé dans EasyVista") }




}

##
##Onglet Outils
##

$btn_outils_igua_OnClick = { Iguazu 
echl "[ii]`t Ouverture d'Iguazu"
$rtxtb_bao_act.appendtext("`rOuverture d'Iguazu")
AutoScroll
}

$btn_outils_sp_OnClick = { ServicePilot 
echl "[ii]`t Ouverture de Service Pilot"
$rtxtb_bao_act.appendtext("`rOuverture de Service Pilot")
AutoScroll
}

$btn_outils_pdm_OnClick = { MSRA 
echl "[ii]`t Prise de main via msra sur le poste"
$rtxtb_bao_act.appendtext("`rPrise de main via msra sur le poste")
AutoScroll
}

$btn_outils_sigma_OnClick = {SIGMA
echl "[ii]`t Reinit Sigma pour le $($txtb_outils_sigma.text) )"
$rtxtb_bao_act.appendtext("`rReinit Sigma pour le $($txtb_outils_sigma.text)")
AutoScroll
}

$btn_outils_wiki_OnClick = {wiki
echl "[ii]`t Ouverture du Wiki"
$rtxtb_bao_act.appendtext("`rOuverture du Wiki")
AutoScroll
}

$btn_outils_grc_OnClick = {GRC
echl "[ii]`t Ouverture de la GRC"
$rtxtb_bao_act.appendtext("`rOuverture de la GRC")
AutoScroll
}

$btn_outils_dsi_OnClick = {DSI
echl "[ii]`t Ouverture du Portail DSI"
$rtxtb_bao_act.appendtext("`rOuverture du Portail DSI")
AutoScroll
}

$btn_outils_ric_OnClick = {CANON
echl "[ii]`t Ouverture du Portail Canon"
$rtxtb_bao_act.appendtext("`rOuverture du Portail CANON")
AutoScroll
}

$btn_outils_wg_OnClick = {WifiGuest
echl "[ii]`t Ouverture du portail Wifi Guest"
$rtxtb_bao_act.appendtext("`rOuverture du portail Wifi Guest")
AutoScroll
}

$btn_outils_remote_OnClick = { Remote
echl "[ii]`t Prise de main Remote"
$rtxtb_bao_act.appendtext("`rPrise de main Remote")
AutoScroll
}

$btn_outils_go_OnClick = {Guidouest
echl "[ii]`t Ouverture de Guidouest"
$rtxtb_bao_act.appendtext("`rOuverture de Guidouest")
AutoScroll
}

$btn_outils_smart_OnClick = {SMART
echl "[ii]`t Ouverture de Smart"
$rtxtb_bao_act.appendtext("`rOuverture de Smart")
AutoScroll
}

$btn_outils_po_OnClick = {Phareouest
echl "[ii]`t Ouverture de Phareouest"
$rtxtb_bao_act.appendtext("`rOuverture de Phareouest")
AutoScroll
}


$btn_outils_clip_OnClick = {
ClipBoard
echl "[ii]`t Copie donnees dans Presse-Papier"
$rtxtb_bao_act.appendtext("`rDonnees copiees dans Presse-Papier")
AutoScroll

}

##
## Onglet AD
##

$btn_ad_verif_OnClick = {
If ($txtb_ad_trig.text -eq ""){$lbl_ad_verifad.Text = "Le champs est vide !" 
								echl "[ii]`t Recherche sur champs vide"}
Else {
$Trigramme  = $txtb_ad_trig.text
VerifCompteWind $Trigramme
AgeMDP $Trigramme
echl "[ii]`t Verification etat du compte et age du mot de passe"
$rtxtb_bao_act.appendtext("`rVerif compte et age du mot de passe")
AutoScroll
}}

$btn_ad_dev_OnClick = {
If ($txtb_ad_trig.text -eq ""){$lbl_ad_verifad.Text = "Le champs est vide !" 
								echl "[ii]`t Recherche sur champs vide"}
Else {
$Trigramme  = $txtb_ad_trig.text
DevAD $Trigramme
echl "[ii]`t $Global:Etat"
$rtxtb_bao_act.appendtext("`rDéverouillage du compte $($txtb_ad_trig.text)")
AutoScroll
}}

$btn_ad_reinit_OnClick = {
If ($txtb_ad_trig.text -eq ""){$lbl_ad_verifad.Text = "Le champs est vide !"
								echl "[ii]`t Recherche sur champs vide"}
Else {
$Trigramme  = $txtb_ad_trig.text
ReinitAD
echl "[ii]`t Le mot de passe de $trigramme a été réinitialisé."
$lbl_ad_reinitinfo.Text = "Le mot de passe a été réinitialisé !"
$rtxtb_bao_act.appendtext("`rRéinitialisation du mot de passe de $Trigramme")
AutoScroll

VerifCompteWind $txtb_ad_trig.text
AgeMDP $txtb_ad_trig.text
} }

$btn_ad_consad_OnClick = {
ConsAD
echl "[ii]`t Ouverture de la console AD"
$rtxtb_bao_act.appendtext("`rOuverture de la console AD")
AutoScroll
}

##
## Onglet Poste
##

$btn_poste_proc_ref_OnClick = {
Get-ProcessInfo
echl "[ii]`t Refresh Process"
$rtxtb_bao_act.appendtext("`rRefresh Process")
AutoScroll
}

$btn_poste_infos_rebsta_OnClick = {RebootSTA
echl "[ii]`t Reboot du STA $ipsta2"
$rtxtb_bao_act.appendtext("`rReboot du STA $ipsta2")
AutoScroll
}

$btn_poste_serv_ref_OnClick = {
Get-ServiceInfo
echl "[ii]`t Refresh Services sur $computer"
$rtxtb_bao_act.appendtext("`rRefresh Services sur $computer")
AutoScroll
}

$btn_poste_infos_pdm_OnClick = {
PDMPoste
echl "[ii]`t Prise de main MSRA sur $($lbl_poste_infos_ipres.text)"
$rtxtb_bao_act.appendtext("`rPrise de main MSRA sur $($lbl_poste_infos_ipres.text)")
AutoScroll
}

$btn_poste_infos_remote_OnClick = {
RemotePoste
echl "[ii]`t Prise de main Remote sur $($lbl_poste_infos_ipres.text)"
$rtxtb_bao_act.appendtext("`rPrise de main Remote sur $($lbl_poste_infos_ipres.text)")
AutoScroll
}

$btn_poste_rech_OnClick = {
Purge_label
if ($cbx_poste_verifposte.text -eq ""){$lbl_poste_infos_netbios.Text = "Le champs est vide !"
										echl "[ii]`t Recherche sur champs vide"}
Else {
$cbx_poste_verifposte.Items.Add($cbx_poste_verifposte.Text)
$global:TempComputer = $cbx_poste_verifposte.Text

switch -Wildcard ($cbx_poste_verifposte.Text)
{
    "10.*"{		ExtractIP		
				fastping -computer $computer
				echl "[ii]`t Recherche sur $TempComputer"
				AutoScroll}
	Default{
			if ((Test-ADComputer $TempComputer) -eq $false) {Write-Host "Le poste n'existe pas dans l'AD"
				$rtxtb_bao_act.appendtext("`rRecherche sur $TempComputer KO")
				$lbl_poste_infos_netbios.Text = "Le $TempComputer n'existe pas dans l'AD"
				}
			Else {Write-host "Le poste existe dans l'AD"
				ExtractIP
				fastping -computer $computer
				$rtxtb_bao_act.appendtext("`rRecherche sur $TempComputer")
				echl "[ii]`t Recherche sur $TempComputer"
				AutoScroll}
			}
}
}}

$btn_poste_serv_arrel_OnClick = {
AR_Serv
echl "[ii]`t Arrêt/relance Service"
$rtxtb_bao_act.appendtext("`rArrêt/relance Service")
AutoScroll
}

$btn_poste_proc_kill_OnClick = {
Kill_Proc
echl "[ii]`t Arret process"
$rtxtb_bao_act.appendtext("`rArret process")
AutoScroll
}

$btn_poste_infos_num_OnClick = {
MonterNumérisation
echl "[ii]`t Montage du $ipsta2\Numérisation"
$rtxtb_bao_act.appendtext("`rAccès au lecteur Numerisation")
AutoScroll
}

$btn_poste_infos_pdmsta_OnClick = {PDMSTA
echl "[ii]`t Prise de main sur le STA $ipsta2"
$rtxtb_bao_act.appendtext("`rPrise de main sur le STA $ipsta2")
AutoScroll
}

$btn_poste_infos_repsccm_OnClick = { RepSCCM
echl "[ii]`t Réparation Client SCCM"
$rtxtb_bao_act.appendtext("`rRéparation Client SCCM")
AutoScroll
}

$btn_poste_infos_conssccm_OnClick = {ConsSCCM
echl "[ii]`t Ouverture Console SCCM"
$rtxtb_bao_act.appendtext("`rOuverture Console SCCM")
AutoScroll
}

$btn_poste_infos_cflog_OnClick = {ConfLogon
echl "[ii]`t Ouverture Configuration Logon"
$rtxtb_bao_act.appendtext("`rOuverture Configuration Logon")
AutoScroll
}

##
## Onglet Genesys
##

$btn_genesys_rechnum_trig_OnClick = {
If ($txtb_genesys_rechnum_trig.text -eq ""){$lbl_genesys_rechnum_trigag.Text = "Le champs est vide !"
											echl "[ii]`t Recherche sur champs vide"}
Else {
$Trigramme = $txtb_genesys_rechnum_trig.text
if ($Trigramme -eq $null ) { $lbl_genesys_rechnum_trigag.text = ""
							 $rtxtb_bao_act.appendtext("`rRecherche numero associe") }
Else {
Num_associe
$lbl_genesys_rechnum_trigag.text = "Place : $($retourgen.Place[0])"
$rtxtb_bao_act.appendtext("`rRecherche numero associe") 
}
AutoScroll 
}}

$btn_genesys_delog_lancer_OnClick = {
if ($rdbtn_genesys_delog_site.Checked -eq $true) {
            if (($cbx_genesys_delog_poste.Text -eq "") -or ($txtb_genesys_delog_place.text -eq "") -or ($cbx_genesys_delog_poste.Text -eq "" -and $txtb_genesys_delog_place.text -eq "" ) ) {$rtxtb_bao_act.appendtext("`rChamps requis manquant") 
            AutoScroll}
            Else { $newplace = $cbx_genesys_delog_poste.Text 
            Delog_Bandeau
            $rtxtb_bao_act.appendtext("`rDeloguage du Bandeau")
            AutoScroll}
}

Elseif ($rdbtn_genesys_delog_ag.Checked -eq $true) {
            if (($txtb_genesys_delog_place.text -eq "") -or ($txtb_genesys_delog_trig.text -eq "") -or ($txtb_genesys_delog_place.text -eq "" -and $txtb_genesys_delog_trig.text -eq "" ) ){$rtxtb_bao_act.appendtext("`rChamps requis manquant")
             AutoScroll}
            Else {$newplace = $txtb_genesys_delog_place.text
            Delog_Bandeau
            $rtxtb_bao_act.appendtext("`rDeloguage du Bandeau")
            AutoScroll}
            }

Else {$rtxtb_bao_act.appendtext("`rMerci de choisir Site ou Agence, je suis pas devin !") 
		AutoScroll}
#Delog_Bandeau
#$rtxtb_bao_act.appendtext("`rDeloguage du Bandeau")
#AutoScroll
}

$btn_genesys_rechnum_num_OnClick = {
If ($txtb_genesys_rechnum_num.text -eq ""){$lbl_genesys_rechnum_trigag.Text = "Le champs est vide !"
											echl "[ii]`t Recherche sur champs vide"}
Else {
$Global:num_tel = $txtb_genesys_rechnum_num.text
Hist_No
$lbl_genesys_rechnum_numtrig.Text = "Trigramme: $($Hist_No.Trigramme[1])"
$lbl_genesys_rechnum_numnom.Text = "Nom: $($Hist_No.Nom[1]+" "+$Hist_No.Prenom[1])"
echl "[ii]`t Recherche sur le numéro $($txtb_genesys_rechnum_num.text)"
$rtxtb_bao_act.appendtext("`rRecherche sur le numéro $($txtb_genesys_rechnum_num.text)")
AutoScroll
}}

##
## Onglet Historique
##

$btn_hist_sugg_OnClick = {
Suggestions
echl "[ii]`t Suggestion envoyée"
$rtxtb_bao_act.appendtext("`rSuggestion envoyée")
AutoScroll
}

$OnLoadForm_StateCorrection = {	$BAO.WindowState = $InitialFormWindowState }

#----------------------------------------------
#region Generated Form Code
$BAO.BackColor = [System.Drawing.Color]::FromArgb(255,185,209,234)
$BAO.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\AppDSI\EXPGLB\OUTILSTA\BAO\Images\bao.ico')
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 476
$System_Drawing_Size.Width = 730
$BAO.ClientSize = $System_Drawing_Size
$BAO.DataBindings.DefaultDataSourceUpdateMode = 0
$BAO.FormBorderStyle = 1
$BAO.MaximizeBox = $False
$BAO.Name = "BAO"
$BAO.Text = "BAO"
#_____________________________________
$tbct_bao_ongl.DataBindings.DefaultDataSourceUpdateMode = 0
$tbct_bao_ongl.ImeMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 270
$System_Drawing_Point.Y = 13
$tbct_bao_ongl.Location = $System_Drawing_Point
$tbct_bao_ongl.Name = "tbct_bao_ongl"
$tbct_bao_ongl.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 451
$System_Drawing_Size.Width = 448
$tbct_bao_ongl.Size = $System_Drawing_Size
$tbct_bao_ongl.SizeMode = 2
$tbct_bao_ongl.TabIndex = 5
$BAO.Controls.Add($tbct_bao_ongl)
#_____________________________________
$ong_bao_outils.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_outils.Location = $System_Drawing_Point
$ong_bao_outils.Name = "ong_bao_outils"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_outils.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 425
$System_Drawing_Size.Width = 440
$ong_bao_outils.Size = $System_Drawing_Size
$ong_bao_outils.TabIndex = 0
$ong_bao_outils.Text = "Outils"
$ong_bao_outils.UseVisualStyleBackColor = $True
$tbct_bao_ongl.Controls.Add($ong_bao_outils)
#_____________________________________
$gbx_outils_liens.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 212
$gbx_outils_liens.Location = $System_Drawing_Point
$gbx_outils_liens.Name = "gbx_outils_liens"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 205
$System_Drawing_Size.Width = 409
$gbx_outils_liens.Size = $System_Drawing_Size
$gbx_outils_liens.TabIndex = 3
$gbx_outils_liens.TabStop = $False
$gbx_outils_liens.Text = "Liens"
$ong_bao_outils.Controls.Add($gbx_outils_liens)
#_____________________________________
$btn_outils_wg.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 230
$System_Drawing_Point.Y = 156
$btn_outils_wg.Location = $System_Drawing_Point
$btn_outils_wg.Name = "btn_outils_wg"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 31
$System_Drawing_Size.Width = 109
$btn_outils_wg.Size = $System_Drawing_Size
$btn_outils_wg.TabIndex = 9
$btn_outils_wg.Text = "WifiGuest"
$btn_outils_wg.UseVisualStyleBackColor = $True
$btn_outils_wg.add_Click($btn_outils_wg_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_wg)
#_____________________________________
$btn_outils_ric.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 230
$System_Drawing_Point.Y = 121
$btn_outils_ric.Location = $System_Drawing_Point
$btn_outils_ric.Name = "btn_outils_ric"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 29
$System_Drawing_Size.Width = 109
$btn_outils_ric.Size = $System_Drawing_Size
$btn_outils_ric.TabIndex = 8
$btn_outils_ric.Text = "Canon"
$btn_outils_ric.UseVisualStyleBackColor = $True
$btn_outils_ric.add_Click($btn_outils_ric_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_ric)
#_____________________________________
$btn_outils_go.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 230
$System_Drawing_Point.Y = 86
$btn_outils_go.Location = $System_Drawing_Point
$btn_outils_go.Name = "btn_outils_go"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 28
$System_Drawing_Size.Width = 109
$btn_outils_go.Size = $System_Drawing_Size
$btn_outils_go.TabIndex = 7
$btn_outils_go.Text = "GuidOuest"
$btn_outils_go.UseVisualStyleBackColor = $True
$btn_outils_go.add_Click($btn_outils_go_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_go)
#_____________________________________
$btn_outils_grc.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 230
$System_Drawing_Point.Y = 51
$btn_outils_grc.Location = $System_Drawing_Point
$btn_outils_grc.Name = "btn_outils_grc"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 28
$System_Drawing_Size.Width = 109
$btn_outils_grc.Size = $System_Drawing_Size
$btn_outils_grc.TabIndex = 6
$btn_outils_grc.Text = "GRC"
$btn_outils_grc.UseVisualStyleBackColor = $True
$btn_outils_grc.add_Click($btn_outils_grc_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_grc)
#_____________________________________
$btn_outils_smart.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 230
$System_Drawing_Point.Y = 18
$btn_outils_smart.Location = $System_Drawing_Point
$btn_outils_smart.Name = "btn_outils_smart"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 27
$System_Drawing_Size.Width = 109
$btn_outils_smart.Size = $System_Drawing_Size
$btn_outils_smart.TabIndex = 5
$btn_outils_smart.Text = "Smart"
$btn_outils_smart.UseVisualStyleBackColor = $True
$btn_outils_smart.add_Click($btn_outils_smart_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_smart)
#_____________________________________
$btn_outils_sp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 71
$System_Drawing_Point.Y = 156
$btn_outils_sp.Location = $System_Drawing_Point
$btn_outils_sp.Name = "btn_outils_sp"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 31
$System_Drawing_Size.Width = 123
$btn_outils_sp.Size = $System_Drawing_Size
$btn_outils_sp.TabIndex = 4
$btn_outils_sp.Text = "ServicePilot"
$btn_outils_sp.UseVisualStyleBackColor = $True
$btn_outils_sp.add_Click($btn_outils_sp_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_sp)
#_____________________________________
$btn_outils_igua.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 70
$System_Drawing_Point.Y = 120
$btn_outils_igua.Location = $System_Drawing_Point
$btn_outils_igua.Name = "btn_outils_igua"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 124
$btn_outils_igua.Size = $System_Drawing_Size
$btn_outils_igua.TabIndex = 3
$btn_outils_igua.Text = "Iguazu"
$btn_outils_igua.UseVisualStyleBackColor = $True
$btn_outils_igua.add_Click($btn_outils_igua_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_igua)
#_____________________________________
$btn_outils_po.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 69
$System_Drawing_Point.Y = 85
$btn_outils_po.Location = $System_Drawing_Point
$btn_outils_po.Name = "btn_outils_po"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 29
$System_Drawing_Size.Width = 124
$btn_outils_po.Size = $System_Drawing_Size
$btn_outils_po.TabIndex = 2
$btn_outils_po.Text = "PhareOuest"
$btn_outils_po.UseVisualStyleBackColor = $True
$btn_outils_po.add_Click($btn_outils_po_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_po)
#_____________________________________
$btn_outils_dsi.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 69
$System_Drawing_Point.Y = 51
$btn_outils_dsi.Location = $System_Drawing_Point
$btn_outils_dsi.Name = "btn_outils_dsi"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 28
$System_Drawing_Size.Width = 124
$btn_outils_dsi.Size = $System_Drawing_Size
$btn_outils_dsi.TabIndex = 1
$btn_outils_dsi.Text = "Portail DSI"
$btn_outils_dsi.UseVisualStyleBackColor = $True
$btn_outils_dsi.add_Click($btn_outils_dsi_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_dsi)
#_____________________________________
$btn_outils_wiki.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 70
$System_Drawing_Point.Y = 18
$btn_outils_wiki.Location = $System_Drawing_Point
$btn_outils_wiki.Name = "btn_outils_wiki"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 27
$System_Drawing_Size.Width = 124
$btn_outils_wiki.Size = $System_Drawing_Size
$btn_outils_wiki.TabIndex = 0
$btn_outils_wiki.Text = "Wiki"
$btn_outils_wiki.UseVisualStyleBackColor = $True
$btn_outils_wiki.add_Click($btn_outils_wiki_OnClick)
$gbx_outils_liens.Controls.Add($btn_outils_wiki)
#_____________________________________
$btn_outils_sigma.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 95
$btn_outils_sigma.Location = $System_Drawing_Point
$btn_outils_sigma.Name = "btn_outils_sigma"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 131
$btn_outils_sigma.Size = $System_Drawing_Size
$btn_outils_sigma.TabIndex = 2
$btn_outils_sigma.Text = "Sigma"
$btn_outils_sigma.UseVisualStyleBackColor = $True
$btn_outils_sigma.add_Click($btn_outils_sigma_OnClick)
$ong_bao_outils.Controls.Add($btn_outils_sigma)
#_____________________________________
$btn_outils_remote.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 56
$btn_outils_remote.Location = $System_Drawing_Point
$btn_outils_remote.Name = "btn_outils_remote"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 131
$btn_outils_remote.Size = $System_Drawing_Size
$btn_outils_remote.TabIndex = 1
$btn_outils_remote.Text = "Remote"
$btn_outils_remote.UseVisualStyleBackColor = $True
$btn_outils_remote.add_Click($btn_outils_remote_OnClick)
$ong_bao_outils.Controls.Add($btn_outils_remote)
#_____________________________________
$btn_outils_pdm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 17
$btn_outils_pdm.Location = $System_Drawing_Point
$btn_outils_pdm.Name = "btn_outils_pdm"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 131
$btn_outils_pdm.Size = $System_Drawing_Size
$btn_outils_pdm.TabIndex = 0
$btn_outils_pdm.Text = "Prise de main"
$btn_outils_pdm.UseVisualStyleBackColor = $True
$btn_outils_pdm.add_Click($btn_outils_pdm_OnClick)
$ong_bao_outils.Controls.Add($btn_outils_pdm)
#_____________________________________

#
$txtb_outils_sigma.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 100
$txtb_outils_sigma.Location = $System_Drawing_Point
$txtb_outils_sigma.Name = "textBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 70
$txtb_outils_sigma.Size = $System_Drawing_Size
$txtb_outils_sigma.TabIndex = 2
$ong_bao_outils.Controls.Add($txtb_outils_sigma)
#_____________________________________
$btn_outils_clip.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 17
$btn_outils_clip.Location = $System_Drawing_Point
$btn_outils_clip.Name = "btn_outils_clip"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 70
$System_Drawing_Size.Width = 131
$btn_outils_clip.Size = $System_Drawing_Size
$btn_outils_clip.TabIndex = 0
$btn_outils_clip.Text = "Copier les données dans le Presse Papier"
$btn_outils_clip.UseVisualStyleBackColor = $True
$btn_outils_clip.add_Click($btn_outils_clip_OnClick)
$ong_bao_outils.Controls.Add($btn_outils_clip)

##________________________________________
#$rbtn_outils_verifntb.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 16
#$System_Drawing_Point.Y = 50
#$rbtn_outils_verifntb.Location = $System_Drawing_Point
#$rbtn_outils_verifntb.Name = "radioButton1"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 24
#$System_Drawing_Size.Width = 62
#$rbtn_outils_verifntb.Size = $System_Drawing_Size
#$rbtn_outils_verifntb.TabIndex = 1
#$rbtn_outils_verifntb.TabStop = $True
#$rbtn_outils_verifntb.Text = "Autre :"
#$rbtn_outils_verifntb.UseVisualStyleBackColor = $True
#$ong_bao_outils.Controls.Add($rbtn_outils_verifntb)
#________________________________________
$cbx_outils_verifntb.DataBindings.DefaultDataSourceUpdateMode = 0
$cbx_outils_verifntb.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 16
$System_Drawing_Point.Y = 20
$cbx_outils_verifntb.Location = $System_Drawing_Point
$cbx_outils_verifntb.Name = "comboBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 75
$cbx_outils_verifntb.Size = $System_Drawing_Size
$cbx_outils_verifntb.TabIndex = 0
$ong_bao_outils.Controls.Add($cbx_outils_verifntb)
$cbx_outils_verifntb.visible = $false
#________________________________________
#________________________________________
#________________________________________
#$cbx_poste_verifposte
$cbx_poste_verifposte.DataBindings.DefaultDataSourceUpdateMode = 0
$cbx_poste_verifposte.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 17
$cbx_poste_verifposte.Location = $System_Drawing_Point
$cbx_poste_verifposte.Name = "comboBox2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 113
$cbx_poste_verifposte.Size = $System_Drawing_Size
$cbx_poste_verifposte.TabIndex = 0
$ong_bao_poste.Controls.Add($cbx_poste_verifposte)



#$txtb_poste_netbios.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 6
#$System_Drawing_Point.Y = 17
#$txtb_poste_netbios.Location = $System_Drawing_Point
#$txtb_poste_netbios.Name = "txtb_poste_netbios"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 20
#$System_Drawing_Size.Width = 113
#$txtb_poste_netbios.Size = $System_Drawing_Size
#$txtb_poste_netbios.TabIndex = 0
#$ong_bao_poste.Controls.Add($txtb_poste_netbios)
#________________________________________
#________________________________________
#________________________________________
$ong_bao_ad.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_ad.Location = $System_Drawing_Point
$ong_bao_ad.Name = "ong_bao_ad"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_ad.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 423
$System_Drawing_Size.Width = 440
$ong_bao_ad.Size = $System_Drawing_Size
$ong_bao_ad.TabIndex = 1
$ong_bao_ad.Text = "AD"
$ong_bao_ad.UseVisualStyleBackColor = $True
$tbct_bao_ongl.Controls.Add($ong_bao_ad)
#_____________________________________
$lbl_ad_reinitinfo.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 29
$System_Drawing_Point.Y = 288
$lbl_ad_reinitinfo.Location = $System_Drawing_Point
$lbl_ad_reinitinfo.Name = "lbl_ad_reinitinfo"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 180
$lbl_ad_reinitinfo.Size = $System_Drawing_Size
$lbl_ad_reinitinfo.TabIndex = 7
$lbl_ad_reinitinfo.Text = ""
$ong_bao_ad.Controls.Add($lbl_ad_reinitinfo)
#_____________________________________
$btn_ad_reinit.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 206
$System_Drawing_Point.Y = 230
$btn_ad_reinit.Location = $System_Drawing_Point
$btn_ad_reinit.Name = "btn_ad_reinit"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 122
$btn_ad_reinit.Size = $System_Drawing_Size
$btn_ad_reinit.TabIndex = 6
$btn_ad_reinit.Text = "Réinitialiser"
$btn_ad_reinit.UseVisualStyleBackColor = $True
$btn_ad_reinit.add_Click($btn_ad_reinit_OnClick)
$ong_bao_ad.Controls.Add($btn_ad_reinit)
#_____________________________________
$txtb_ad_mdp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 29
$System_Drawing_Point.Y = 236
$txtb_ad_mdp.Location = $System_Drawing_Point
$txtb_ad_mdp.Name = "txtb_ad_mdp"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 133
$txtb_ad_mdp.Size = $System_Drawing_Size
$txtb_ad_mdp.TabIndex = 5
$ong_bao_ad.Controls.Add($txtb_ad_mdp)
$txtb_ad_mdp.text = "Groupama1"
#_____________________________________
$lbl_ad_reinit.DataBindings.DefaultDataSourceUpdateMode = 0
$lbl_ad_reinit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9.75,1,3,1)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 94
$System_Drawing_Point.Y = 183
$lbl_ad_reinit.Location = $System_Drawing_Point
$lbl_ad_reinit.Name = "lbl_ad_reinit"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 36
$System_Drawing_Size.Width = 253
$lbl_ad_reinit.Size = $System_Drawing_Size
$lbl_ad_reinit.TabIndex = 4
$lbl_ad_reinit.Text = "Réinitialisation du mot de passe : "
$ong_bao_ad.Controls.Add($lbl_ad_reinit)
#_____________________________________
$btn_ad_consad.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 97
$btn_ad_consad.Location = $System_Drawing_Point
$btn_ad_consad.Name = "btn_ad_consad"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 35
$System_Drawing_Size.Width = 135
$btn_ad_consad.Size = $System_Drawing_Size
$btn_ad_consad.TabIndex = 3
$btn_ad_consad.Text = "Console AD"
$btn_ad_consad.UseVisualStyleBackColor = $True
$btn_ad_consad.add_Click($btn_ad_consad_OnClick)
$ong_bao_ad.Controls.Add($btn_ad_consad)
#_____________________________________
$btn_ad_dev.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 57
$btn_ad_dev.Location = $System_Drawing_Point
$btn_ad_dev.Name = "btn_ad_dev"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 135
$btn_ad_dev.Size = $System_Drawing_Size
$btn_ad_dev.TabIndex = 2
$btn_ad_dev.Text = "Déverouiller"
$btn_ad_dev.UseVisualStyleBackColor = $True
$btn_ad_dev.add_Click($btn_ad_dev_OnClick)
$ong_bao_ad.Controls.Add($btn_ad_dev)
#_____________________________________
$btn_ad_verif.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 300
$System_Drawing_Point.Y = 17
$btn_ad_verif.Location = $System_Drawing_Point
$btn_ad_verif.Name = "btn_ad_verif"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 33
$System_Drawing_Size.Width = 135
$btn_ad_verif.Size = $System_Drawing_Size
$btn_ad_verif.TabIndex = 1
$btn_ad_verif.Text = "Vérification"
$btn_ad_verif.UseVisualStyleBackColor = $True
$btn_ad_verif.add_Click($btn_ad_verif_OnClick)
$ong_bao_ad.Controls.Add($btn_ad_verif)
#_____________________________________
$lbl_ad_verifad.DataBindings.DefaultDataSourceUpdateMode = 0
#$lbl_ad_verifad.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9.75,1,3,1)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 70
$lbl_ad_verifad.Location = $System_Drawing_Point
$lbl_ad_verifad.Name = "lbl_ad_verifad"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 36
$System_Drawing_Size.Width = 253
$lbl_ad_verifad.Size = $System_Drawing_Size
$lbl_ad_verifad.TabIndex = 4
$lbl_ad_verifad.Text = ""
$ong_bao_ad.Controls.Add($lbl_ad_verifad)
#_____________________________________
$lbl_ad_agepassword.DataBindings.DefaultDataSourceUpdateMode = 0
#$lbl_ad_agepassword.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9.75,1,3,1)
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 120
$lbl_ad_agepassword.Location = $System_Drawing_Point
$lbl_ad_agepassword.Name = "lbl_ad_verifad"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 36
$System_Drawing_Size.Width = 253
$lbl_ad_agepassword.Size = $System_Drawing_Size
$lbl_ad_agepassword.TabIndex = 4
$lbl_ad_agepassword.Text = ""
$ong_bao_ad.Controls.Add($lbl_ad_agepassword)
#_____________________________________
$txtb_ad_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 24
$txtb_ad_trig.Location = $System_Drawing_Point
$txtb_ad_trig.Name = "txtb_ad_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 123
$txtb_ad_trig.Size = $System_Drawing_Size
$txtb_ad_trig.TabIndex = 0
$ong_bao_ad.Controls.Add($txtb_ad_trig)
#_____________________________________
$ong_bao_poste.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$ong_bao_poste.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_poste.Location = $System_Drawing_Point
$ong_bao_poste.Name = "ong_bao_poste"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_poste.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 423
$System_Drawing_Size.Width = 440
$ong_bao_poste.Size = $System_Drawing_Size
$ong_bao_poste.TabIndex = 2
$ong_bao_poste.Text = "Poste"
$tbct_bao_ongl.Controls.Add($ong_bao_poste)
#_____________________________________
$tabControl2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = -4
$System_Drawing_Point.Y = 46
$tabControl2.Location = $System_Drawing_Point
$tabControl2.Multiline = $True
$tabControl2.Name = "tabControl2"
$tabControl2.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 381
$System_Drawing_Size.Width = 448
$tabControl2.Size = $System_Drawing_Size
$tabControl2.SizeMode = 2
$tabControl2.TabIndex = 2
$ong_bao_poste.Controls.Add($tabControl2)
#_____________________________________
$ong_bao_poste_infos.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_poste_infos.Location = $System_Drawing_Point
$ong_bao_poste_infos.Name = "ong_bao_poste_infos"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_poste_infos.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 355
$System_Drawing_Size.Width = 440
$ong_bao_poste_infos.Size = $System_Drawing_Size
$ong_bao_poste_infos.TabIndex = 0
$ong_bao_poste_infos.Text = "Infos"
$ong_bao_poste_infos.UseVisualStyleBackColor = $True
$tabControl2.Controls.Add($ong_bao_poste_infos)
#_____________________________________
$lbl_poste_infos_etstares.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 143
$lbl_poste_infos_etstares.Location = $System_Drawing_Point
$lbl_poste_infos_etstares.Name = "lbl_poste_infos_etstares"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 152
$lbl_poste_infos_etstares.Size = $System_Drawing_Size
$lbl_poste_infos_etstares.TabIndex = 21
$lbl_poste_infos_etstares.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_etstares)
#_____________________________________
$lbl_poste_infos_ipstares.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 126
$lbl_poste_infos_ipstares.Location = $System_Drawing_Point
$lbl_poste_infos_ipstares.Name = "lbl_poste_infos_ipstares"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 152
$lbl_poste_infos_ipstares.Size = $System_Drawing_Size
$lbl_poste_infos_ipstares.TabIndex = 20
$lbl_poste_infos_ipstares.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ipstares)
#_____________________________________
$lbl_poste_infos_sitres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 108
$lbl_poste_infos_sitres.Location = $System_Drawing_Point
$lbl_poste_infos_sitres.Name = "lbl_poste_infos_sitres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 200
$lbl_poste_infos_sitres.Size = $System_Drawing_Size
$lbl_poste_infos_sitres.TabIndex = 19
$lbl_poste_infos_sitres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_sitres)
$lbl_poste_infos_posccmres.DataBindings.DefaultDataSourceUpdateMode = 0
#_____________________________________
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 91
$lbl_poste_infos_posccmres.Location = $System_Drawing_Point
$lbl_poste_infos_posccmres.Name = "lbl_poste_infos_posccmres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 152
$lbl_poste_infos_posccmres.Size = $System_Drawing_Size
$lbl_poste_infos_posccmres.TabIndex = 18
$lbl_poste_infos_posccmres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_posccmres)
#_____________________________________
$lbl_poste_infos_titsccmres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 73
$lbl_poste_infos_titsccmres.Location = $System_Drawing_Point
$lbl_poste_infos_titsccmres.Name = "lbl_poste_infos_titsccmres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 152
$lbl_poste_infos_titsccmres.Size = $System_Drawing_Size
$lbl_poste_infos_titsccmres.TabIndex = 17
$lbl_poste_infos_titsccmres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_titsccmres)
#_____________________________________
$lbl_poste_infos_bootres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 56
$lbl_poste_infos_bootres.Location = $System_Drawing_Point
$lbl_poste_infos_bootres.Name = "lbl_poste_infos_bootres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 152
$lbl_poste_infos_bootres.Size = $System_Drawing_Size
$lbl_poste_infos_bootres.TabIndex = 16
$lbl_poste_infos_bootres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_bootres)
#_____________________________________
$lbl_poste_infos_ipres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 38
$lbl_poste_infos_ipres.Location = $System_Drawing_Point
$lbl_poste_infos_ipres.Name = "lbl_poste_infos_ipres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 152
$lbl_poste_infos_ipres.Size = $System_Drawing_Size
$lbl_poste_infos_ipres.TabIndex = 15
$lbl_poste_infos_ipres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ipres)
#_____________________________________
#$lbl_poste_infos_netbiosres.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 133
#$System_Drawing_Point.Y = 6
#$lbl_poste_infos_netbiosres.Location = $System_Drawing_Point
#$lbl_poste_infos_netbiosres.Name = "lbl_poste_infos_netbiosres"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 19
#$System_Drawing_Size.Width = 152
#$lbl_poste_infos_netbiosres.Size = $System_Drawing_Size
#$lbl_poste_infos_netbiosres.TabIndex = 14
#$lbl_poste_infos_netbiosres.Text = "-"
#$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_netbiosres)
#_____________________________________
#_____________________________________
#
#
#
#
#
#
#
#_____________________________________
$lbl_poste_infos_ramres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 299
$lbl_poste_infos_ramres.Location = $System_Drawing_Point
$lbl_poste_infos_ramres.Name = "lbl_poste_infos_ramres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_ramres.Size = $System_Drawing_Size
$lbl_poste_infos_ramres.TabIndex = 5
$lbl_poste_infos_ramres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ramres)
#_____________________________________
$lbl_poste_infos_ram.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 299
$lbl_poste_infos_ram.Location = $System_Drawing_Point
$lbl_poste_infos_ram.Name = "lbl_poste_infos_ram"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_ram.Size = $System_Drawing_Size
$lbl_poste_infos_ram.TabIndex = 5
$lbl_poste_infos_ram.Text = "RAM : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ram)
#_____________________________________
$lbl_poste_infos_diskstatres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 281
$lbl_poste_infos_diskstatres.Location = $System_Drawing_Point
$lbl_poste_infos_diskstatres.Name = "lbl_poste_infos_diskstatres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_diskstatres.Size = $System_Drawing_Size
$lbl_poste_infos_diskstatres.TabIndex = 5
$lbl_poste_infos_diskstatres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_diskstatres)
#_____________________________________
$lbl_poste_infos_diskstat.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 281
$lbl_poste_infos_diskstat.Location = $System_Drawing_Point
$lbl_poste_infos_diskstat.Name = "lbl_poste_infos_diskstat"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_diskstat.Size = $System_Drawing_Size
$lbl_poste_infos_diskstat.TabIndex = 5
$lbl_poste_infos_diskstat.Text = "Etat du Disque Dur : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_diskstat)
#_____________________________________
$lbl_poste_infos_freespaceres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 263
$lbl_poste_infos_freespaceres.Location = $System_Drawing_Point
$lbl_poste_infos_freespaceres.Name = "lbl_poste_infos_freespaceres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_freespaceres.Size = $System_Drawing_Size
$lbl_poste_infos_freespaceres.TabIndex = 5
$lbl_poste_infos_freespaceres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_freespaceres)
#_____________________________________
$lbl_poste_infos_freespace.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 263
$lbl_poste_infos_freespace.Location = $System_Drawing_Point
$lbl_poste_infos_freespace.Name = "lbl_poste_infos_freespace"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_freespace.Size = $System_Drawing_Size
$lbl_poste_infos_freespace.TabIndex = 5
$lbl_poste_infos_freespace.Text = "Espace Libre :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_freespace)
#_____________________________________
$lbl_poste_infos_sizeres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 245
$lbl_poste_infos_sizeres.Location = $System_Drawing_Point
$lbl_poste_infos_sizeres.Name = "lbl_poste_infos_sizeres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_sizeres.Size = $System_Drawing_Size
$lbl_poste_infos_sizeres.TabIndex = 5
$lbl_poste_infos_sizeres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_sizeres)
#_____________________________________
$lbl_poste_infos_size.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 245
$lbl_poste_infos_size.Location = $System_Drawing_Point
$lbl_poste_infos_size.Name = "lbl_poste_infos_size"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_size.Size = $System_Drawing_Size
$lbl_poste_infos_size.TabIndex = 5
$lbl_poste_infos_size.Text = "Disque Dur :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_size)
#_____________________________________
$lbl_poste_infos_archires.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 227
$lbl_poste_infos_archires.Location = $System_Drawing_Point
$lbl_poste_infos_archires.Name = "lbl_poste_infos_archires"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_archires.Size = $System_Drawing_Size
$lbl_poste_infos_archires.TabIndex = 5
$lbl_poste_infos_archires.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_archires)
#_____________________________________
$lbl_poste_infos_archi.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 227
$lbl_poste_infos_archi.Location = $System_Drawing_Point
$lbl_poste_infos_archi.Name = "lbl_poste_infos_archi"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 250
$lbl_poste_infos_archi.Size = $System_Drawing_Size
$lbl_poste_infos_archi.TabIndex = 5
$lbl_poste_infos_archi.Text = "Archi du poste :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_archi)
#_____________________________________
$lbl_poste_infos_osres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 209
$lbl_poste_infos_osres.Location = $System_Drawing_Point
$lbl_poste_infos_osres.Name = "lbl_poste_infos_osres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 200
$lbl_poste_infos_osres.Size = $System_Drawing_Size
$lbl_poste_infos_osres.TabIndex = 5
$lbl_poste_infos_osres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_osres)
#_____________________________________
$lbl_poste_infos_os.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 209
$lbl_poste_infos_os.Location = $System_Drawing_Point
$lbl_poste_infos_os.Name = "lbl_poste_infos_os"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_os.Size = $System_Drawing_Size
$lbl_poste_infos_os.TabIndex = 5
$lbl_poste_infos_os.Text = "OS du poste :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_os)
#_____________________________________
$lbl_poste_infos_modelres.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 133
$System_Drawing_Point.Y = 191
$lbl_poste_infos_modelres.Location = $System_Drawing_Point
$lbl_poste_infos_modelres.Name = "lbl_poste_infos_modelres"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_modelres.Size = $System_Drawing_Size
$lbl_poste_infos_modelres.TabIndex = 5
$lbl_poste_infos_modelres.Text = "-"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_modelres)
#_____________________________________
$lbl_poste_infos_model.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 191
$lbl_poste_infos_model.Location = $System_Drawing_Point
$lbl_poste_infos_model.Name = "lbl_poste_infos_model"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 150
$lbl_poste_infos_model.Size = $System_Drawing_Size
$lbl_poste_infos_model.TabIndex = 5
$lbl_poste_infos_model.Text = "Modèle du poste :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_model)
#_____________________________________


#_____________________________________
$lbl_poste_infos_infsys.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 170
$lbl_poste_infos_infsys.Location = $System_Drawing_Point
$lbl_poste_infos_infsys.Name = "lbl_poste_infos_infsys"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 150
$lbl_poste_infos_infsys.Size = $System_Drawing_Size
$lbl_poste_infos_infsys.TabIndex = 13
$lbl_poste_infos_infsys.Text = "Infos Système : "
$lbl_poste_infos_infsys.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,1,3,0)
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_infsys)
#_____________________________________
$lbl_poste_infos_etsta.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 143
$lbl_poste_infos_etsta.Location = $System_Drawing_Point
$lbl_poste_infos_etsta.Name = "lbl_poste_infos_etsta"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 100
$lbl_poste_infos_etsta.Size = $System_Drawing_Size
$lbl_poste_infos_etsta.TabIndex = 12
$lbl_poste_infos_etsta.Text = "Etat du STA : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_etsta)
#_____________________________________
$lbl_poste_infos_ipsta.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 126
$lbl_poste_infos_ipsta.Location = $System_Drawing_Point
$lbl_poste_infos_ipsta.Name = "lbl_poste_infos_ipsta"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 100
$lbl_poste_infos_ipsta.Size = $System_Drawing_Size
$lbl_poste_infos_ipsta.TabIndex = 11
$lbl_poste_infos_ipsta.Text = "Adresse IP STA :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ipsta)
#_____________________________________
$lbl_poste_infos_sit.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 108
$lbl_poste_infos_sit.Location = $System_Drawing_Point
$lbl_poste_infos_sit.Name = "lbl_poste_infos_sit"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 15
$System_Drawing_Size.Width = 100
$lbl_poste_infos_sit.Size = $System_Drawing_Size
$lbl_poste_infos_sit.TabIndex = 10
$lbl_poste_infos_sit.Text = "Situation : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_sit)
#_____________________________________
$lbl_poste_infos_posccm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 91
$lbl_poste_infos_posccm.Location = $System_Drawing_Point
$lbl_poste_infos_posccm.Name = "lbl_poste_infos_posccm"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 17
$System_Drawing_Size.Width = 116
$lbl_poste_infos_posccm.Size = $System_Drawing_Size
$lbl_poste_infos_posccm.TabIndex = 9
$lbl_poste_infos_posccm.Text = "Postes associés :"
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_posccm)
#_____________________________________
$lbl_poste_infos_titsccm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 73
$lbl_poste_infos_titsccm.Location = $System_Drawing_Point
$lbl_poste_infos_titsccm.Name = "lbl_poste_infos_titsccm"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 116
$lbl_poste_infos_titsccm.Size = $System_Drawing_Size
$lbl_poste_infos_titsccm.TabIndex = 8
$lbl_poste_infos_titsccm.Text = "Titulaire du poste : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_titsccm)
#_____________________________________
$lbl_poste_infos_boot.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 56
$lbl_poste_infos_boot.Location = $System_Drawing_Point
$lbl_poste_infos_boot.Name = "lbl_poste_infos_boot"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 124
$lbl_poste_infos_boot.Size = $System_Drawing_Size
$lbl_poste_infos_boot.TabIndex = 7
$lbl_poste_infos_boot.Text = "Dernier démarrage : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_boot)
#_____________________________________
$lbl_poste_infos_ip.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 39
$lbl_poste_infos_ip.Location = $System_Drawing_Point
$lbl_poste_infos_ip.Name = "lbl_poste_infos_ip"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 100
$lbl_poste_infos_ip.Size = $System_Drawing_Size
$lbl_poste_infos_ip.TabIndex = 6
$lbl_poste_infos_ip.Text = "Adresse IP : "
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_ip)
#_____________________________________

#_____________________________________
$lbl_poste_infos_netbios.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 17
$lbl_poste_infos_netbios.Location = $System_Drawing_Point
$lbl_poste_infos_netbios.Name = "lbl_poste_infos_netbios"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 250
$lbl_poste_infos_netbios.Size = $System_Drawing_Size
$lbl_poste_infos_netbios.TabIndex = 5
$lbl_poste_infos_netbios.Text = "Information sur le poste:"
$lbl_poste_infos_netbios.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,1,3,0)
$ong_bao_poste_infos.Controls.Add($lbl_poste_infos_netbios)
#_____________________________________
$btn_poste_infos_pdm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 130
$btn_poste_infos_pdm.Location = $System_Drawing_Point
$btn_poste_infos_pdm.Name = "btn_poste_infos_pdm"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_pdm.Size = $System_Drawing_Size
$btn_poste_infos_pdm.TabIndex = 4
$btn_poste_infos_pdm.Text = "Prise de main"
$btn_poste_infos_pdm.UseVisualStyleBackColor = $True
$btn_poste_infos_pdm.add_Click($btn_poste_infos_pdm_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_pdm)
#_____________________________________
#_____________________________________
#_____________________________________
$btn_poste_infos_repsccm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 250
$btn_poste_infos_repsccm.Location = $System_Drawing_Point
$btn_poste_infos_repsccm.Name = "btn_poste_infos_remote"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_repsccm.Size = $System_Drawing_Size
$btn_poste_infos_repsccm.TabIndex = 4
$btn_poste_infos_repsccm.Text = "Répa Client SCCM"
$btn_poste_infos_repsccm.UseVisualStyleBackColor = $True
$btn_poste_infos_repsccm.add_Click($btn_poste_infos_repsccm_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_repsccm)
#_____________________________________
$btn_poste_infos_conssccm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 280
$btn_poste_infos_conssccm.Location = $System_Drawing_Point
$btn_poste_infos_conssccm.Name = "btn_poste_infos_remote"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_conssccm.Size = $System_Drawing_Size
$btn_poste_infos_conssccm.TabIndex = 4
$btn_poste_infos_conssccm.Text = "Console SCCM"
$btn_poste_infos_conssccm.UseVisualStyleBackColor = $True
$btn_poste_infos_conssccm.add_Click($btn_poste_infos_conssccm_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_conssccm)
#_____________________________________
$btn_poste_infos_cflog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 310
$btn_poste_infos_cflog.Location = $System_Drawing_Point
$btn_poste_infos_cflog.Name = "btn_poste_infos_remote"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_cflog.Size = $System_Drawing_Size
$btn_poste_infos_cflog.TabIndex = 4
$btn_poste_infos_cflog.Text = "ConfLogon"
$btn_poste_infos_cflog.UseVisualStyleBackColor = $True
$btn_poste_infos_cflog.add_Click($btn_poste_infos_cflog_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_cflog)
#_____________________________________
$btn_poste_infos_remote.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 160
$btn_poste_infos_remote.Location = $System_Drawing_Point
$btn_poste_infos_remote.Name = "btn_poste_infos_remote"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_remote.Size = $System_Drawing_Size
$btn_poste_infos_remote.TabIndex = 4
$btn_poste_infos_remote.Text = "Remote"
$btn_poste_infos_remote.UseVisualStyleBackColor = $True
$btn_poste_infos_remote.add_Click($btn_poste_infos_remote_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_remote)
#_____________________________________
$btn_poste_infos_num.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 83
$btn_poste_infos_num.Location = $System_Drawing_Point
$btn_poste_infos_num.Name = "btn_poste_infos_num"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_num.Size = $System_Drawing_Size
$btn_poste_infos_num.TabIndex = 3
$btn_poste_infos_num.Text = "Numérisation STA"
$btn_poste_infos_num.UseVisualStyleBackColor = $True
$btn_poste_infos_num.add_Click($btn_poste_infos_num_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_num)
#_____________________________________
$btn_poste_infos_rebsta.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 53
$btn_poste_infos_rebsta.Location = $System_Drawing_Point
$btn_poste_infos_rebsta.Name = "btn_poste_infos_rebsta"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_rebsta.Size = $System_Drawing_Size
$btn_poste_infos_rebsta.TabIndex = 2
$btn_poste_infos_rebsta.Text = "Reboot STA"
$btn_poste_infos_rebsta.UseVisualStyleBackColor = $True
$btn_poste_infos_rebsta.add_Click($btn_poste_infos_rebsta_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_rebsta)
#_____________________________________
$btn_poste_infos_pdmsta.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 295
$System_Drawing_Point.Y = 23
$btn_poste_infos_pdmsta.Location = $System_Drawing_Point
$btn_poste_infos_pdmsta.Name = "btn_poste_infos_pdmsta"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 132
$btn_poste_infos_pdmsta.Size = $System_Drawing_Size
$btn_poste_infos_pdmsta.TabIndex = 1
$btn_poste_infos_pdmsta.Text = "Prise de main STA"
$btn_poste_infos_pdmsta.UseVisualStyleBackColor = $True
$btn_poste_infos_pdmsta.add_Click($btn_poste_infos_pdmsta_OnClick)
$ong_bao_poste_infos.Controls.Add($btn_poste_infos_pdmsta)
#_____________________________________
#$rtxtb_poste_infos_mat.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 7
#$System_Drawing_Point.Y = 169
#$rtxtb_poste_infos_mat.Location = $System_Drawing_Point
#$rtxtb_poste_infos_mat.Name = "rtxtb_poste_infos_mat"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 180
#$System_Drawing_Size.Width = 278
#$rtxtb_poste_infos_mat.Size = $System_Drawing_Size
#$rtxtb_poste_infos_mat.TabIndex = 0
#$rtxtb_poste_infos_mat.Text = ""
#$ong_bao_poste_infos.Controls.Add($rtxtb_poste_infos_mat)
#_____________________________________
#$ong_bao_poste_sccm.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 4
#$System_Drawing_Point.Y = 22
#$ong_bao_poste_sccm.Location = $System_Drawing_Point
#$ong_bao_poste_sccm.Name = "ong_bao_poste_sccm"
#$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
#$System_Windows_Forms_Padding.All = 3
#$System_Windows_Forms_Padding.Bottom = 3
#$System_Windows_Forms_Padding.Left = 3
#$System_Windows_Forms_Padding.Right = 3
#$System_Windows_Forms_Padding.Top = 3
#$ong_bao_poste_sccm.Padding = $System_Windows_Forms_Padding
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 355
#$System_Drawing_Size.Width = 440
#$ong_bao_poste_sccm.Size = $System_Drawing_Size
#$ong_bao_poste_sccm.TabIndex = 1
#$ong_bao_poste_sccm.Text = "SCCM"
#$ong_bao_poste_sccm.UseVisualStyleBackColor = $True
#$tabControl2.Controls.Add($ong_bao_poste_sccm)
#_____________________________________
#$dtgv_poste_sccm_result.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 7
#$System_Drawing_Point.Y = 34
#$dtgv_poste_sccm_result.Location = $System_Drawing_Point
#$dtgv_poste_sccm_result.Name = "dtgv_poste_sccm_result"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 315
#$System_Drawing_Size.Width = 427
#$dtgv_poste_sccm_result.Size = $System_Drawing_Size
#$dtgv_poste_sccm_result.TabIndex = 3
#$ong_bao_poste_sccm.Controls.Add($dtgv_poste_sccm_result)
#_____________________________________
#$btn_poste_sccm_copy.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 332
#$System_Drawing_Point.Y = 7
#$btn_poste_sccm_copy.Location = $System_Drawing_Point
#$btn_poste_sccm_copy.Name = "btn_poste_sccm_copy"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 23
#$System_Drawing_Size.Width = 102
#$btn_poste_sccm_copy.Size = $System_Drawing_Size
#$btn_poste_sccm_copy.TabIndex = 2
#$btn_poste_sccm_copy.Text = "Copier"
#$btn_poste_sccm_copy.UseVisualStyleBackColor = $True
#$btn_poste_sccm_copy.add_Click($btn_poste_sccm_copy_OnClick)
#$ong_bao_poste_sccm.Controls.Add($btn_poste_sccm_copy)
#_____________________________________
#$btn_poste_sccm_rech.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 220
#$System_Drawing_Point.Y = 7
#$btn_poste_sccm_rech.Location = $System_Drawing_Point
#$btn_poste_sccm_rech.Name = "btn_poste_sccm_rech"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 23
#$System_Drawing_Size.Width = 106
#$btn_poste_sccm_rech.Size = $System_Drawing_Size
#$btn_poste_sccm_rech.TabIndex = 1
#$btn_poste_sccm_rech.Text = "Rechercher"
#$btn_poste_sccm_rech.UseVisualStyleBackColor = $True
#$btn_poste_sccm_rech.add_Click($btn_poste_sccm_rech_OnClick)
#$ong_bao_poste_sccm.Controls.Add($btn_poste_sccm_rech)
#_____________________________________
#$txtb_poste_sccm_rech.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 7
#$System_Drawing_Point.Y = 7
#$txtb_poste_sccm_rech.Location = $System_Drawing_Point
#$txtb_poste_sccm_rech.Name = "txtb_poste_sccm_rech"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 20
#$System_Drawing_Size.Width = 197
#$txtb_poste_sccm_rech.Size = $System_Drawing_Size
#$txtb_poste_sccm_rech.TabIndex = 0
#$ong_bao_poste_sccm.Controls.Add($txtb_poste_sccm_rech)
#_____________________________________
$ong_bao_poste_proc.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_poste_proc.Location = $System_Drawing_Point
$ong_bao_poste_proc.Name = "ong_bao_poste_proc"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_poste_proc.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 355
$System_Drawing_Size.Width = 440
$ong_bao_poste_proc.Size = $System_Drawing_Size
$ong_bao_poste_proc.TabIndex = 2
$ong_bao_poste_proc.Text = "Processus"
$ong_bao_poste_proc.UseVisualStyleBackColor = $True
$tabControl2.Controls.Add($ong_bao_poste_proc)
#_____________________________________
$btn_poste_proc_kill.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 329
$System_Drawing_Point.Y = 322
$btn_poste_proc_kill.Location = $System_Drawing_Point
$btn_poste_proc_kill.Name = "btn_poste_proc_kill"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 105
$btn_poste_proc_kill.Size = $System_Drawing_Size
$btn_poste_proc_kill.TabIndex = 2
$btn_poste_proc_kill.Text = "Arrêt"
$btn_poste_proc_kill.UseVisualStyleBackColor = $True
$btn_poste_proc_kill.add_Click($btn_poste_proc_kill_OnClick)
$ong_bao_poste_proc.Controls.Add($btn_poste_proc_kill)
#_____________________________________
$btn_poste_proc_ref.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 322
$btn_poste_proc_ref.Location = $System_Drawing_Point
$btn_poste_proc_ref.Name = "btn_poste_proc_ref"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 115
$btn_poste_proc_ref.Size = $System_Drawing_Size
$btn_poste_proc_ref.TabIndex = 1
$btn_poste_proc_ref.Text = "Refresh"
$btn_poste_proc_ref.UseVisualStyleBackColor = $True
$btn_poste_proc_ref.add_Click($btn_poste_proc_ref_OnClick)
$ong_bao_poste_proc.Controls.Add($btn_poste_proc_ref)
#_____________________________________
$dtgv_poste_proc_proc.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 0
$dtgv_poste_proc_proc.Location = $System_Drawing_Point
$dtgv_poste_proc_proc.Name = "dtgv_poste_proc_proc"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 307
$System_Drawing_Size.Width = 444
$dtgv_poste_proc_proc.Size = $System_Drawing_Size
$dtgv_poste_proc_proc.TabIndex = 0
$ong_bao_poste_proc.Controls.Add($dtgv_poste_proc_proc)
#_____________________________________
$ong_bao_poste_serv.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_poste_serv.Location = $System_Drawing_Point
$ong_bao_poste_serv.Name = "ong_bao_poste_serv"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_poste_serv.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 355
$System_Drawing_Size.Width = 440
$ong_bao_poste_serv.Size = $System_Drawing_Size
$ong_bao_poste_serv.TabIndex = 3
$ong_bao_poste_serv.Text = "Services"
$ong_bao_poste_serv.UseVisualStyleBackColor = $True
$tabControl2.Controls.Add($ong_bao_poste_serv)
#_____________________________________
$btn_poste_serv_arrel.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 322
$System_Drawing_Point.Y = 322
$btn_poste_serv_arrel.Location = $System_Drawing_Point
$btn_poste_serv_arrel.Name = "btn_poste_serv_arrel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 112
$btn_poste_serv_arrel.Size = $System_Drawing_Size
$btn_poste_serv_arrel.TabIndex = 2
$btn_poste_serv_arrel.Text = "Arrêt/Relance"
$btn_poste_serv_arrel.UseVisualStyleBackColor = $True
$btn_poste_serv_arrel.add_Click($btn_poste_serv_arrel_OnClick)
$ong_bao_poste_serv.Controls.Add($btn_poste_serv_arrel)
#____________________________________
$btn_poste_serv_ref.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 322
$btn_poste_serv_ref.Location = $System_Drawing_Point
$btn_poste_serv_ref.Name = "btn_poste_serv_ref"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 112
$btn_poste_serv_ref.Size = $System_Drawing_Size
$btn_poste_serv_ref.TabIndex = 1
$btn_poste_serv_ref.Text = "Refresh"
$btn_poste_serv_ref.UseVisualStyleBackColor = $True
$btn_poste_serv_ref.add_Click($btn_poste_serv_ref_OnClick)
$ong_bao_poste_serv.Controls.Add($btn_poste_serv_ref)
#_____________________________________
$dtgv_poste_serv_serv.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 0
$dtgv_poste_serv_serv.Location = $System_Drawing_Point
$dtgv_poste_serv_serv.Name = "dtgv_poste_serv_serv"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 307
$System_Drawing_Size.Width = 444
$dtgv_poste_serv_serv.Size = $System_Drawing_Size
$dtgv_poste_serv_serv.TabIndex = 0
$ong_bao_poste_serv.Controls.Add($dtgv_poste_serv_serv)
#_____________________________________
$btn_poste_rech.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 136
$System_Drawing_Point.Y = 17
$btn_poste_rech.Location = $System_Drawing_Point
$btn_poste_rech.Name = "btn_poste_rech"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 103
$btn_poste_rech.Size = $System_Drawing_Size
$btn_poste_rech.TabIndex = 1
$btn_poste_rech.Text = "Rechercher"
$btn_poste_rech.UseVisualStyleBackColor = $True
$btn_poste_rech.add_Click($btn_poste_rech_OnClick)
$ong_bao_poste.Controls.Add($btn_poste_rech)
#_____________________________________
$txtb_poste_netbios.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 17
$txtb_poste_netbios.Location = $System_Drawing_Point
$txtb_poste_netbios.Name = "txtb_poste_netbios"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 113
$txtb_poste_netbios.Size = $System_Drawing_Size
$txtb_poste_netbios.TabIndex = 0
$ong_bao_poste.Controls.Add($txtb_poste_netbios)
#_____________________________________
$tabPage4.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$tabPage4.Location = $System_Drawing_Point
$tabPage4.Name = "tabPage4"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$tabPage4.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 425
$System_Drawing_Size.Width = 440
$tabPage4.Size = $System_Drawing_Size
$tabPage4.TabIndex = 3
$tabPage4.Text = "Genesys"
$tabPage4.UseVisualStyleBackColor = $True
$tbct_bao_ongl.Controls.Add($tabPage4)
#_____________________________________
$tabControl1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = -4
$System_Drawing_Point.Y = 43
$tabControl1.Location = $System_Drawing_Point
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 384
$System_Drawing_Size.Width = 448
$tabControl1.Size = $System_Drawing_Size
$tabControl1.TabIndex = 2
$tabPage4.Controls.Add($tabControl1)
#_____________________________________
$ong_genesys_histut.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_genesys_histut.Location = $System_Drawing_Point
$ong_genesys_histut.Name = "ong_genesys_histut"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_genesys_histut.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 358
$System_Drawing_Size.Width = 432
$ong_genesys_histut.Size = $System_Drawing_Size
$ong_genesys_histut.TabIndex = 0
$ong_genesys_histut.Text = "Historique UT"
$ong_genesys_histut.UseVisualStyleBackColor = $True
#_____________________________________
$tabControl1.Controls.Add($ong_genesys_histut)
$dtgv_genesys_histut_hist.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 0
$dtgv_genesys_histut_hist.Location = $System_Drawing_Point
$dtgv_genesys_histut_hist.Name = "dtgv_genesys_histut_hist"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 321
$System_Drawing_Size.Width = 436
$dtgv_genesys_histut_hist.Size = $System_Drawing_Size
$dtgv_genesys_histut_hist.TabIndex = 0
$ong_genesys_histut.Controls.Add($dtgv_genesys_histut_hist)
#_____________________________________
$ong_genesys_histnum.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_genesys_histnum.Location = $System_Drawing_Point
$ong_genesys_histnum.Name = "ong_genesys_histnum"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_genesys_histnum.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 358
$System_Drawing_Size.Width = 432
$ong_genesys_histnum.Size = $System_Drawing_Size
$ong_genesys_histnum.TabIndex = 1
$ong_genesys_histnum.Text = "Historique Numéro"
$ong_genesys_histnum.UseVisualStyleBackColor = $True
#_____________________________________
$tabControl1.Controls.Add($ong_genesys_histnum)
$dtgv_genesys_histnum_hist.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = -1
$System_Drawing_Point.Y = 0
$dtgv_genesys_histnum_hist.Location = $System_Drawing_Point
$dtgv_genesys_histnum_hist.Name = "dtgv_genesys_histnum_hist"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 334
$System_Drawing_Size.Width = 433
$dtgv_genesys_histnum_hist.Size = $System_Drawing_Size
$dtgv_genesys_histnum_hist.TabIndex = 0
$ong_genesys_histnum.Controls.Add($dtgv_genesys_histnum_hist)
#_____________________________________
$ong_genesys_delog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_genesys_delog.Location = $System_Drawing_Point
$ong_genesys_delog.Name = "ong_genesys_delog"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_genesys_delog.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 358
$System_Drawing_Size.Width = 432
$ong_genesys_delog.Size = $System_Drawing_Size
$ong_genesys_delog.TabIndex = 2
$ong_genesys_delog.Text = "Délogage"
$ong_genesys_delog.UseVisualStyleBackColor = $True
$tabControl1.Controls.Add($ong_genesys_delog)
#_____________________________________
$btn_genesys_delog_lancer.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 286
$System_Drawing_Point.Y = 128
$btn_genesys_delog_lancer.Location = $System_Drawing_Point
$btn_genesys_delog_lancer.Name = "btn_genesys_delog_lancer"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 69
$System_Drawing_Size.Width = 111
$btn_genesys_delog_lancer.Size = $System_Drawing_Size
$btn_genesys_delog_lancer.TabIndex = 4
$btn_genesys_delog_lancer.Text = "Lancer Bandeau"
$btn_genesys_delog_lancer.UseVisualStyleBackColor = $True
$btn_genesys_delog_lancer.add_Click($btn_genesys_delog_lancer_OnClick)
$ong_genesys_delog.Controls.Add($btn_genesys_delog_lancer)
#_____________________________________
$lbl_genesys_delog_place.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 184
$System_Drawing_Point.Y = 186
$lbl_genesys_delog_place.Location = $System_Drawing_Point
$lbl_genesys_delog_place.Name = "lbl_genesys_delog_place"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 42
$lbl_genesys_delog_place.Size = $System_Drawing_Size
$lbl_genesys_delog_place.TabIndex = 6
$lbl_genesys_delog_place.Text = "Place"
$ong_genesys_delog.Controls.Add($lbl_genesys_delog_place)
#_____________________________________
$lbl_genesys_delog_poste.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 184
$System_Drawing_Point.Y = 120
$lbl_genesys_delog_poste.Location = $System_Drawing_Point
$lbl_genesys_delog_poste.Name = "lbl_genesys_delog_place"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 42
$lbl_genesys_delog_poste.Size = $System_Drawing_Size
$lbl_genesys_delog_poste.TabIndex = 6
$lbl_genesys_delog_poste.Text = "Poste"
$ong_genesys_delog.Controls.Add($lbl_genesys_delog_poste)
#_____________________________________
$lbl_genesys_delog_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 172
$System_Drawing_Point.Y = 65
$lbl_genesys_delog_trig.Location = $System_Drawing_Point
$lbl_genesys_delog_trig.Name = "lbl_genesys_delog_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 62
$lbl_genesys_delog_trig.Size = $System_Drawing_Size
$lbl_genesys_delog_trig.TabIndex = 5
$lbl_genesys_delog_trig.Text = "Trigramme"
$ong_genesys_delog.Controls.Add($lbl_genesys_delog_trig)
#_____________________________________
$txtb_genesys_delog_place.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 154
$System_Drawing_Point.Y = 212
$txtb_genesys_delog_place.Location = $System_Drawing_Point
$txtb_genesys_delog_place.Name = "txtb_genesys_delog_place"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$txtb_genesys_delog_place.Size = $System_Drawing_Size
$txtb_genesys_delog_place.TabIndex = 3
$ong_genesys_delog.Controls.Add($txtb_genesys_delog_place)
#_____________________________________
$cbx_genesys_delog_poste.DataBindings.DefaultDataSourceUpdateMode = 0
$cbx_genesys_delog_poste.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 154
$System_Drawing_Point.Y = 152
$cbx_genesys_delog_poste.Location = $System_Drawing_Point
$cbx_genesys_delog_poste.Name = "comboBox3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$cbx_genesys_delog_poste.Size = $System_Drawing_Size
$cbx_genesys_delog_poste.TabIndex = 0
$ong_genesys_delog.Controls.Add($cbx_genesys_delog_poste)
#
#$cbx_genesys_delog_poste.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 154
#$System_Drawing_Point.Y = 152
#$cbx_genesys_delog_poste.Location = $System_Drawing_Point
#$cbx_genesys_delog_poste.Name = "txtb_genesys_delog_place"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 20
#$System_Drawing_Size.Width = 100
#$cbx_genesys_delog_poste.Size = $System_Drawing_Size
#$cbx_genesys_delog_poste.TabIndex = 3
#$ong_genesys_delog.Controls.Add($cbx_genesys_delog_poste)
#_____________________________________
$txtb_genesys_delog_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 154
$System_Drawing_Point.Y = 91
$txtb_genesys_delog_trig.Location = $System_Drawing_Point
$txtb_genesys_delog_trig.Name = "txtb_genesys-delog_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$txtb_genesys_delog_trig.Size = $System_Drawing_Size
$txtb_genesys_delog_trig.TabIndex = 2
$ong_genesys_delog.Controls.Add($txtb_genesys_delog_trig)
#_____________________________________
$rdbtn_genesys_delog_site.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 28
$System_Drawing_Point.Y = 209
$rdbtn_genesys_delog_site.Location = $System_Drawing_Point
$rdbtn_genesys_delog_site.Name = "rdbtn_genesys_delog_place"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$rdbtn_genesys_delog_site.Size = $System_Drawing_Size
$rdbtn_genesys_delog_site.TabIndex = 1
$rdbtn_genesys_delog_site.TabStop = $True
$rdbtn_genesys_delog_site.Text = "Site"
$rdbtn_genesys_delog_site.UseVisualStyleBackColor = $True
$ong_genesys_delog.Controls.Add($rdbtn_genesys_delog_site)
#_____________________________________
$rdbtn_genesys_delog_ag.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 28
$System_Drawing_Point.Y = 88
$rdbtn_genesys_delog_ag.Location = $System_Drawing_Point
$rdbtn_genesys_delog_ag.Name = "rdbtn_genesys_delog_ag"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 104
$rdbtn_genesys_delog_ag.Size = $System_Drawing_Size
$rdbtn_genesys_delog_ag.TabIndex = 0
$rdbtn_genesys_delog_ag.TabStop = $True
$rdbtn_genesys_delog_ag.Text = "Agence"
$rdbtn_genesys_delog_ag.UseVisualStyleBackColor = $True
$ong_genesys_delog.Controls.Add($rdbtn_genesys_delog_ag)
#_____________________________________
$ong_genesys_rechnum.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_genesys_rechnum.Location = $System_Drawing_Point
$ong_genesys_rechnum.Name = "ong_genesys_rechnum"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_genesys_rechnum.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 358
$System_Drawing_Size.Width = 440
$ong_genesys_rechnum.Size = $System_Drawing_Size
$ong_genesys_rechnum.TabIndex = 3
$ong_genesys_rechnum.Text = "Recherche par Numéro"
$ong_genesys_rechnum.UseVisualStyleBackColor = $True
$tabControl1.Controls.Add($ong_genesys_rechnum)
#_____________________________________
$lbl_genesys_rechnum_trigag.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 265
$System_Drawing_Point.Y = 275
$lbl_genesys_rechnum_trigag.Location = $System_Drawing_Point
$lbl_genesys_rechnum_trigag.Name = "lbl_genesys_rechnum_trigag"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 121
$lbl_genesys_rechnum_trigag.Size = $System_Drawing_Size
$lbl_genesys_rechnum_trigag.TabIndex = 9
$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_trigag)
#_____________________________________
#$lbl_genesys_rechnum_trignum.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 265
#$System_Drawing_Point.Y = 212
#$lbl_genesys_rechnum_trignum.Location = $System_Drawing_Point
#$lbl_genesys_rechnum_trignum.Name = "lbl_genesys_rechnum_trignum"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 23
#$System_Drawing_Size.Width = 121
#$lbl_genesys_rechnum_trignum.Size = $System_Drawing_Size
#$lbl_genesys_rechnum_trignum.TabIndex = 8
#$lbl_genesys_rechnum_trignum.Text = "Numéro : "
#$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_trignum)
##_____________________________________
$lbl_genesys_rechnum_numnom.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 46
$System_Drawing_Point.Y = 275
$lbl_genesys_rechnum_numnom.Location = $System_Drawing_Point
$lbl_genesys_rechnum_numnom.Name = "lbl_genesys_rechnum_numnom"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 160
$lbl_genesys_rechnum_numnom.Size = $System_Drawing_Size
$lbl_genesys_rechnum_numnom.TabIndex = 7
$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_numnom)
#_____________________________________
$lbl_genesys_rechnum_numtrig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 46
$System_Drawing_Point.Y = 212
$lbl_genesys_rechnum_numtrig.Location = $System_Drawing_Point
$lbl_genesys_rechnum_numtrig.Name = "lbl_genesys_rechnum_numtrig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 121
$lbl_genesys_rechnum_numtrig.Size = $System_Drawing_Size
$lbl_genesys_rechnum_numtrig.TabIndex = 6
$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_numtrig)
#_____________________________________
$btn_genesys_rechnum_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 265
$System_Drawing_Point.Y = 105
$btn_genesys_rechnum_trig.Location = $System_Drawing_Point
$btn_genesys_rechnum_trig.Name = "btn_genesys_rechnum_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 69
$System_Drawing_Size.Width = 121
$btn_genesys_rechnum_trig.Size = $System_Drawing_Size
$btn_genesys_rechnum_trig.TabIndex = 5
$btn_genesys_rechnum_trig.Text = "Où est logué ?"
$btn_genesys_rechnum_trig.UseVisualStyleBackColor = $True
$btn_genesys_rechnum_trig.add_Click($btn_genesys_rechnum_trig_OnClick)
#_____________________________________
$ong_genesys_rechnum.Controls.Add($btn_genesys_rechnum_trig)
$txtb_genesys_rechnum_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 265
$System_Drawing_Point.Y = 38
$txtb_genesys_rechnum_trig.Location = $System_Drawing_Point
$txtb_genesys_rechnum_trig.Name = "txtb_genesys_rechnum_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$txtb_genesys_rechnum_trig.Size = $System_Drawing_Size
$txtb_genesys_rechnum_trig.TabIndex = 4
$ong_genesys_rechnum.Controls.Add($txtb_genesys_rechnum_trig)
#_____________________________________
$btn_genesys_rechnum_num.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 46
$System_Drawing_Point.Y = 105
$btn_genesys_rechnum_num.Location = $System_Drawing_Point
$btn_genesys_rechnum_num.Name = "btn_genesys_rechnum_num"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 69
$System_Drawing_Size.Width = 121
$btn_genesys_rechnum_num.Size = $System_Drawing_Size
$btn_genesys_rechnum_num.TabIndex = 3
$btn_genesys_rechnum_num.Text = "Qui est logué ?"
$btn_genesys_rechnum_num.UseVisualStyleBackColor = $True
$btn_genesys_rechnum_num.add_Click($btn_genesys_rechnum_num_OnClick)
#_____________________________________
$ong_genesys_rechnum.Controls.Add($btn_genesys_rechnum_num)
$txtb_genesys_rechnum_num.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 59
$System_Drawing_Point.Y = 38
$txtb_genesys_rechnum_num.Location = $System_Drawing_Point
$txtb_genesys_rechnum_num.Name = "txtb_genesys_rechnum_num"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 100
$txtb_genesys_rechnum_num.Size = $System_Drawing_Size
$txtb_genesys_rechnum_num.TabIndex = 2
$ong_genesys_rechnum.Controls.Add($txtb_genesys_rechnum_num)
#_____________________________________
$lbl_genesys_rechnum_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 284
$System_Drawing_Point.Y = 12
$lbl_genesys_rechnum_trig.Location = $System_Drawing_Point
$lbl_genesys_rechnum_trig.Name = "lbl_genesys_rechnum_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 59
$lbl_genesys_rechnum_trig.Size = $System_Drawing_Size
$lbl_genesys_rechnum_trig.TabIndex = 1
$lbl_genesys_rechnum_trig.Text = "Trigramme"
$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_trig)
#_____________________________________
$lbl_genesys_rechnum_num.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 67
$System_Drawing_Point.Y = 13
$lbl_genesys_rechnum_num.Location = $System_Drawing_Point
$lbl_genesys_rechnum_num.Name = "lbl_genesys_rechnum_num"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 92
$lbl_genesys_rechnum_num.Size = $System_Drawing_Size
$lbl_genesys_rechnum_num.TabIndex = 0
$lbl_genesys_rechnum_num.Text = "Numéro de tél"
$ong_genesys_rechnum.Controls.Add($lbl_genesys_rechnum_num)
#_____________________________________
$lbl_genesys_num.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 231
$System_Drawing_Point.Y = 10
$lbl_genesys_num.Location = $System_Drawing_Point
$lbl_genesys_num.Name = "lbl_genesys_num"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 174
$lbl_genesys_num.Size = $System_Drawing_Size
$lbl_genesys_num.TabIndex = 1
$lbl_genesys_num.Text = "Numéro GENESYS :"
$tabPage4.Controls.Add($lbl_genesys_num)
#_____________________________________
$lbl_genesys_trig.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 54
$System_Drawing_Point.Y = 10
$lbl_genesys_trig.Location = $System_Drawing_Point
$lbl_genesys_trig.Name = "lbl_genesys_trig"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 30
$System_Drawing_Size.Width = 168
$lbl_genesys_trig.Size = $System_Drawing_Size
$lbl_genesys_trig.TabIndex = 0
$lbl_genesys_trig.Text = "Trigramme : "
$tabPage4.Controls.Add($lbl_genesys_trig)
#_____________________________________
$ong_bao_hist.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$ong_bao_hist.Location = $System_Drawing_Point
$ong_bao_hist.Name = "ong_bao_hist"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$ong_bao_hist.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 423
$System_Drawing_Size.Width = 440
$ong_bao_hist.Size = $System_Drawing_Size
$ong_bao_hist.TabIndex = 5
$ong_bao_hist.Text = "Historique"
$ong_bao_hist.UseVisualStyleBackColor = $True
#_____________________________________
$lbl_hist_sugg.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 300
$lbl_hist_sugg.Location = $System_Drawing_Point
$lbl_hist_sugg.Name = "lbl_hist_sugg"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 224
$lbl_hist_sugg.Size = $System_Drawing_Size
$lbl_hist_sugg.TabIndex = 1
$lbl_hist_sugg.Text = "Des suggestions ou remarques?"
$ong_bao_hist.Controls.Add($lbl_hist_sugg)
#_____________________________________
$tbct_bao_ongl.Controls.Add($ong_bao_hist)
#$rtxtb_hist_sugg.DataBindings.DefaultDataSourceUpdateMode = 0
#$System_Drawing_Point = New-Object System.Drawing.Point
#$System_Drawing_Point.X = 7
#$System_Drawing_Point.Y = 309
#$rtxtb_hist_sugg.Location = $System_Drawing_Point
#$rtxtb_hist_sugg.Name = "rtxtb_hist_sugg"
#$System_Drawing_Size = New-Object System.Drawing.Size
#$System_Drawing_Size.Height = 90
#$System_Drawing_Size.Width = 234
#$rtxtb_hist_sugg.Size = $System_Drawing_Size
#$rtxtb_hist_sugg.TabIndex = 3
#$rtxtb_hist_sugg.Text = ""
#$ong_bao_hist.Controls.Add($rtxtb_hist_sugg)
#_____________________________________
$btn_hist_sugg.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 267
$System_Drawing_Point.Y = 309
$btn_hist_sugg.Location = $System_Drawing_Point
$btn_hist_sugg.Name = "btn_hist_sugg"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 90
$System_Drawing_Size.Width = 114
$btn_hist_sugg.Size = $System_Drawing_Size
$btn_hist_sugg.TabIndex = 2
$btn_hist_sugg.Text = "Envoyer un Email"
$btn_hist_sugg.UseVisualStyleBackColor = $True
$btn_hist_sugg.add_Click($btn_hist_sugg_OnClick)
$ong_bao_hist.Controls.Add($btn_hist_sugg)
#_____________________________________
$lbl_hist_histo.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 17
$lbl_hist_histo.Location = $System_Drawing_Point
$lbl_hist_histo.Name = "lbl_hist_histo"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 224
$lbl_hist_histo.Size = $System_Drawing_Size
$lbl_hist_histo.TabIndex = 1
$lbl_hist_histo.Text = "Historiques des derniers appelants"
$ong_bao_hist.Controls.Add($lbl_hist_histo)
#_____________________________________
$rtxtb_hist_histo.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 43
$rtxtb_hist_histo.Location = $System_Drawing_Point
$rtxtb_hist_histo.Name = "rtxtb_hist_histo"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 221
$System_Drawing_Size.Width = 427
$rtxtb_hist_histo.Size = $System_Drawing_Size
$rtxtb_hist_histo.TabIndex = 0
$rtxtb_hist_histo.Text = ""
$rtxtb_hist_histo.readonly = "False"
$ong_bao_hist.Controls.Add($rtxtb_hist_histo)
#_____________________________________
$rtxtb_bao_act.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 375
$rtxtb_bao_act.Location = $System_Drawing_Point
$rtxtb_bao_act.Name = "rtxtb_bao_act"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 87
$System_Drawing_Size.Width = 241
$rtxtb_bao_act.Size = $System_Drawing_Size
$rtxtb_bao_act.TabIndex = 4
$rtxtb_bao_act.Text = ""
$rtxtb_bao_act.readonly = "False"
$rtxtb_bao_act.SelectionStart = $rtxtb_bao_act.TextLength;
$rtxtb_bao_act.ScrollToCaret()
$BAO.Controls.Add($rtxtb_bao_act)
#_____________________________________
$rtxtb_bao_info.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 94
$rtxtb_bao_info.Location = $System_Drawing_Point
$rtxtb_bao_info.Name = "rtxtb_bao_info"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 266
$System_Drawing_Size.Width = 241
$rtxtb_bao_info.Size = $System_Drawing_Size
$rtxtb_bao_info.TabIndex = 3
$rtxtb_bao_info.Text = ""
$rtxtb_bao_info.readonly = "False"
$BAO.Controls.Add($rtxtb_bao_info)
#_____________________________________
$lbl_bao_tp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$lbl_bao_tp.Location = $System_Drawing_Point
$lbl_bao_tp.Name = "lbl_bao_tp"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 116
$lbl_bao_tp.Size = $System_Drawing_Size
$lbl_bao_tp.TabIndex = 2
$lbl_bao_tp.Text = "Trigramme ou Poste :"
$BAO.Controls.Add($lbl_bao_tp)
#_____________________________________
$btn_bao_rech.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 165
$System_Drawing_Point.Y = 12
$btn_bao_rech.Location = $System_Drawing_Point
$btn_bao_rech.Name = "btn_bao_rech"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 73
$System_Drawing_Size.Width = 89
$btn_bao_rech.Size = $System_Drawing_Size
$btn_bao_rech.TabIndex = 1
$btn_bao_rech.Text = "Rechercher"
$btn_bao_rech.UseVisualStyleBackColor = $True
$btn_bao_rech.add_Click($btn_bao_rech_OnClick)
$BAO.Controls.Add($btn_bao_rech)
#_____________________________________
$txtb_bao_tp.DataBindings.DefaultDataSourceUpdateMode = 0
$BAO.AcceptButton = $btn_bao_rech
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 41
$txtb_bao_tp.Location = $System_Drawing_Point
$txtb_bao_tp.Name = "txtb_bao_tp"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 116
$txtb_bao_tp.Size = $System_Drawing_Size
$txtb_bao_tp.TabIndex = 0
$BAO.Controls.Add($txtb_bao_tp)
#_____________________________________
#endregion Generated Form Code

$InitialFormWindowState = $BAO.WindowState
$BAO.add_Load($OnLoadForm_StateCorrection)
[Void]$BAO.ShowDialog()

}


# ---------------------------------------------------------
# Gestion des répertoire et des Logs
# ---------------------------------------------------------

Try
{
	GetMyIguazu "EXP"
	InitScript
	$script:grrr = Get-Content env:username
	$global:user = $grrr.substring(0,3)
	Credentials $user "domain"
	$cred_glb = Get-Cred("domain")
	GenerateForm
	
	
}
Catch
{
	echl "[!!]`t $($Error[0])"
	echl $_.Exception
}
Finally
{
	# Cloture le log et renvoi le code d'erreur
	LogClose
}


