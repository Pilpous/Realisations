###-----------------------------------------------------------------------------------------------###
### ONGLET HISTORIQUE
###-----------------------------------------------------------------------------------------------###

Function Historique {
$rtxtb_hist_histo.appendtext("$($result1.FIRSTNAME) $($result1.LASTNAME) - $($result1.TRIGRAMME) - $($Rech_Netbios.ID_RESSOURCE)`r")

$rtxtb_hist_histo.SelectionStart = $rtxtb_hist_histo.TextLength;
$rtxtb_hist_histo.ScrollToCaret()}

Function Suggestions {
$olFolderDrafts = 16
$ol = New-Object -comObject Outlook.Application 
$ns = $ol.GetNameSpace("MAPI")

$mail = $ol.CreateItem(0)
$null = $Mail.Recipients.Add("mail")  
$Mail.Subject = "[BAO - SUGGESTIONS]:"  
$Mail.display()


$draft.Send()
}

