#################################################
##												#
## SCRIPT FICHE PLANNING						#
## 		Réalisateur : ScripTeam		#
##												#
#################################################


Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms, System.Drawing
### Runspace 1 - Form
$global:syncHash = [hashtable]::Synchronized(@{})
$global:FPlanRunspace =[runspacefactory]::CreateRunspace()
$global:FPlanRunspace.ApartmentState = "STA"
$global:FPlanRunspace.ThreadOptions = "ReuseThread"         
$global:FPlanRunspace.Open()
$global:FPlanRunspace.SessionStateProxy.SetVariable("syncHash",$global:syncHash)          
 
$global:script_code = {
### Fonction Excel
Function OpenFiche {
$date_fiche = Get-Date -Format yyyy_MM_dd
$PathToXlsx = "C:\AppDSI\DEVGLB\mbg\Fiche_Planning\Fiche de modification du planning.xlsx"

$U_TA = "U:\SIT\TA\_Fiches planning\A SAISIR AU PLANNING\"
$local_temp = "C:\temp\"
if (!(test-path $U_TA)) {$ExcelPath = "$($local_temp)$date_fiche $utilisateur Fiche de modification du planning.xlsx"
						 $saved_fiche = $local_temp 
						 
						  	$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
							$rect_alerte = $xml_al_form.FindName('rect_alerte')
							$btn_ok_alerte.Margin = "256,30,0,0"
							$xml_al_form.Height = 70
							$txtbck_alerte.Height = 70
							$rect_alerte.Height = 70
							
							$txtbck_alerte.text = "Attention ! Le lecteur U: n'est pas accessible, la fiche sera sauvegardée en local, et devra être déplacée manuellement par vos soins !"
					        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					        $xml_al_form.showdialog()
						 
						 
						 }
Else {$ExcelPath = "$($local_temp)$date_fiche $utilisateur Fiche de modification du planning.xlsx"
		#$saved_fiche = $U_TA
		$saved_fiche = $local_temp }

$SheetName = "Feuil1"

$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false
$workBook = $objExcel.Workbooks.Open($PathToXlsx)
$workSheet = $workBook.Worksheets.Item(1)
$workSheet.Cells.Item(1,6) = $global:syncHash.cbbox_nom.get_text()
$global:checked = $global:checked_rdbtn | select -Unique

switch -wildcard ($global:checked ) {
#Inversion
	"rbt_inv"{	for ($count_rbt = 0; $count_rbt -le $count_inv ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_inv,1) = $rbt_inv_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_inv,2) = $rbt_inv_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_inv,4) = $rbt_inv_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_inv,6) = $rbt_inv_infos[$count_rbt][3]
				
				$global:line_rbt_inv++ } }
	"rbt_aug"  {	for ($count_rbt = 0; $count_rbt -le $count_aug ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_aug ,1) = $rbt_aug_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_aug ,2) = $rbt_aug_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_aug ,4) = $rbt_aug_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_aug ,6) = $rbt_aug_infos[$count_rbt][3]
				
				$global:line_rbt_aug ++ }  }
				
	"rbt_dec" {	for ($count_rbt = 0; $count_rbt -le $count_dec ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_dec ,1) = $rbt_dec_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_dec ,2) = $rbt_dec_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_dec ,4) = $rbt_dec_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_dec ,6) = $rbt_dec_infos[$count_rbt][3]
				
				$global:line_rbt_dec ++ } }
				
	"rbt_modif" {for ($count_rbt = 0; $count_rbt -le $count_modif ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_modif ,1) = $rbt_modif_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_modif ,2) = $rbt_modif_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_modif ,4) = $rbt_modif_infos[$count_rbt][3]
				$workSheet.Cells.Item($global:line_rbt_modif ,5) = $rbt_modif_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_modif ,6) = $rbt_modif_infos[$count_rbt][4]

				$global:line_rbt_modif ++
				} }
				
	"rbt_reclun" { for ($count_rbt = 0; $count_rbt -le $count_reclun ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_reclun ,1) = $rbt_reclun_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_reclun ,2) = $rbt_reclun_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_reclun ,4) = $rbt_reclun_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_reclun ,6) = $rbt_reclun_infos[$count_rbt][3]
				
				$global:line_rbt_reclun++ } }
				
	"rbt_pent" {for ($count_rbt = 0; $count_rbt -le $count_pent ; $count_rbt++ ) {
				$workSheet = $workBook.Worksheets.Item(1)
				$workSheet.Cells.Item($global:line_rbt_pent ,1) = $rbt_pent_infos[$count_rbt][0]
				$workSheet.Cells.Item($global:line_rbt_pent ,2) = $rbt_pent_infos[$count_rbt][1]
				$workSheet.Cells.Item($global:line_rbt_pent ,4) = $rbt_pent_infos[$count_rbt][2]
				$workSheet.Cells.Item($global:line_rbt_pent ,6) = $rbt_pent_infos[$count_rbt][3]
				$workSheet.Cells.Item($global:line_rbt_pent ,7) = $rbt_pent_infos[$count_rbt][4]
				
				$global:line_rbt_pent++ } }
	}		

$workSheet.SaveAs($ExcelPath)
$workBook.close()
$objExcel.quit()

#$workSheet.Close($ExcelPath)
}


### XAML
$ThemeFile = 'C:\AppDSI\DEVGLB\mbg\Fiche_Planning\Theme.xaml'

[xml]$alerte_msg = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow"  WindowStartupLocation="CenterScreen" Height="50" Width="350" WindowStyle="None" BorderBrush="Black" Cursor="Arrow" FontFamily="Calibri"
        AllowsTransparency="True"
        ResizeMode="NoResize">
		<Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile" /> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
         </Window.Resources>
    <Grid>
        <Button Name="btn_ok_alerte" Content="OK" HorizontalAlignment="Left" Margin="256,10,0,0" VerticalAlignment="Top" Width="75" Grid.Column="1" Height="22" ClickMode="Press" IsHitTestVisible="True"/>
        <TextBlock Name="txtbck_alerte" HorizontalAlignment="Left" Margin="11,12,0,0" TextWrapping="Wrap" Text="Tous les champs notés d'une * sont obligatoires !" VerticalAlignment="Top" Height="38" Width="250"/>
        <Rectangle Name="rect_alerte" HorizontalAlignment="Left" Height="50" Stroke="Black" VerticalAlignment="Top" Width="350"/>
    </Grid>
</Window>
"@

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FichePlanning" Height="390" Width="554.787" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" WindowStyle="None">
		<Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile" /> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
         </Window.Resources>
    <Grid>
        <Button Name="btn_quit" Content="X" HorizontalAlignment="Left" Margin="525,10,0,0" VerticalAlignment="Top" Width="22"/>
        <Button Name="btn_mini" Content="_" HorizontalAlignment="Left" Margin="498,10,0,0" VerticalAlignment="Top" Width="22"/>
        <Label Content="Fiche Planning" HorizontalAlignment="Left" Height="32" Margin="8,1,0,0" VerticalAlignment="Top" Width="460" FontWeight="Bold" FontSize="16"/>
        <Label Name="lbl_nom" Content="Nom : " HorizontalAlignment="Left" Height="27" Margin="10,38,0,0" VerticalAlignment="Top" Width="55"/>
        <RadioButton Name="rbt_inv" Content="Inversion d'horaire" HorizontalAlignment="Left" Margin="34,100,0,0" VerticalAlignment="Top" Width="194"/>
        <RadioButton Name="rbt_aug" Content="Augmentation Pause déjeuner" HorizontalAlignment="Left" Margin="34,121,0,0" VerticalAlignment="Top" Width="194"/>
        <RadioButton Name="rbt_dec" Content="Décalage Pause déjeuner" HorizontalAlignment="Left" Margin="34,142,0,0" VerticalAlignment="Top" Width="194"/>
        <RadioButton Name="rbt_modif" Content="Modification Horaire - RecupE" HorizontalAlignment="Left" Margin="34,163,0,0" VerticalAlignment="Top" Width="194"/>
        <RadioButton Name="rbt_reclun" Content="Récup Lundi" HorizontalAlignment="Left" Margin="34,184,0,0" VerticalAlignment="Top" Width="194"/>
        <RadioButton Name="rbt_pent" Content="Récup Pentecote" HorizontalAlignment="Left" Margin="34,205,0,0" VerticalAlignment="Top" Width="194"/>
        <Label Name="lbl_modif" Content="Modification à faire : " HorizontalAlignment="Left" Height="30" Margin="10,73,0,0" VerticalAlignment="Top" Width="239"/>
        <ComboBox Name="cbbox_nom" HorizontalAlignment="Left" Margin="65,43,0,0" VerticalAlignment="Top" Width="198"/>
        <DatePicker Name="datepick" HorizontalAlignment="Left" Height="27" Margin="249,73,0,0" VerticalAlignment="Top" Width="154" Visibility="Hidden" SelectedDateFormat="Short"/>
        <ComboBox Name="cbbox_remp" HorizontalAlignment="Left" Margin="342,109,0,0" VerticalAlignment="Top" Width="161" Height="26" Visibility="Hidden"/>
        <Label Name="lbl_remp" Content="Remplacant : *" HorizontalAlignment="Left" Margin="249,109,0,0" VerticalAlignment="Top" Width="93" Visibility="Hidden"/>
        <Label Name="lbl_nouvhor" Content="Nouvel Horaire : *" HorizontalAlignment="Left" Height="26" Margin="249,140,0,0" VerticalAlignment="Top" Width="104" Visibility="Hidden"/>
        <CheckBox Name="ckbx_8" IsEnabled="True" Content="8h-16h42" HorizontalAlignment="Left" Margin="358,145,0,0" VerticalAlignment="Top" Height="24" Width="97" Visibility="Hidden"/>
        <CheckBox Name="ckbx_17" IsEnabled="True" Content="8h18-17h" HorizontalAlignment="Left" Margin="358,166,0,0" VerticalAlignment="Top" Width="97" Visibility="Hidden"/>
        <Label Name="lbl_durpause" Content="Durée de la pause : *" HorizontalAlignment="Left" Height="34" Margin="249,187,0,0" VerticalAlignment="Top" Width="122" Visibility="Hidden"/>
        <TextBox Name="txtbx_durpause" HorizontalAlignment="Left" Height="23" Margin="371,189,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_horpause" Content="Horaire de la pause : *" HorizontalAlignment="Left" Margin="249,216,0,0" VerticalAlignment="Top" Width="126" Visibility="Hidden"/>
        <TextBox Name="txtbx_horpause" HorizontalAlignment="Left" Height="23" Margin="371,217,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_horreel" Content="Horaire réel de la jounée : *" HorizontalAlignment="Left" Margin="249,438,0,0" VerticalAlignment="Top" Width="154" Visibility="Hidden"/>
        <TextBox Name="txtbx_horreel" HorizontalAlignment="Left" Height="23" Margin="396,439,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_inc" Content="N° Incident : *" HorizontalAlignment="Left" Margin="249,245,0,0" VerticalAlignment="Top" Width="122" Visibility="Hidden"/>
        <TextBox Name="txtbx_inc" HorizontalAlignment="Left" Height="23" Margin="371,246,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_rais" Content="Raison : *" HorizontalAlignment="Left" Margin="249,272,0,0" VerticalAlignment="Top" Width="122" Visibility="Hidden"/>
        <TextBox Name="txtbx_raison" HorizontalAlignment="Left" Height="23" Margin="371,273,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_recupe" Content="Temps de RecupE : *" HorizontalAlignment="Left" Margin="249,304,0,0" VerticalAlignment="Top" Width="122" Visibility="Hidden"/>
        <CheckBox Name="ckbx_plus" Content="En Plus" HorizontalAlignment="Left" Margin="373,311,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <CheckBox Name="ckbx_moins" Content="En moins" HorizontalAlignment="Left" Margin="373,336,0,0" VerticalAlignment="Top" Visibility="Hidden"/>
        <TextBox Name="txtbx_plus" HorizontalAlignment="Left" Height="23" Margin="449,310,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="71" Visibility="Hidden"/>
        <TextBox Name="txtbx_moins" HorizontalAlignment="Left" Height="23" Margin="449,333,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="71" Visibility="Hidden"/>
        <Label Name="lbl_tpssupp" Content="Temps Suppl (min) : *" HorizontalAlignment="Left" Margin="249,360,0,0" VerticalAlignment="Top" Width="126" Visibility="Hidden"/>
        <TextBox Name="txtbx_tpssup" HorizontalAlignment="Left" Height="23" Margin="371,363,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_debfin" Content="Hor.Début/Hor.Fin Solidarité : *" HorizontalAlignment="Left" Margin="249,387,0,0" VerticalAlignment="Top" Width="179" Visibility="Hidden"/>
        <TextBox Name="txtbx_debfin" HorizontalAlignment="Left" Height="23" Margin="416,390,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Label Name="lbl_tpsrest" Content="Temps restant à faire : " HorizontalAlignment="Left" Margin="249,463,0,0" VerticalAlignment="Top" Width="154" Visibility="Hidden"/>
        <TextBox Name="txtbx_tpsrest" HorizontalAlignment="Left" Height="23" Margin="396,466,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <Button Name="btn_save" Content="Sauvegarder la Fiche Planning" HorizontalAlignment="Left" Height="58" Margin="34,309,0,0" VerticalAlignment="Top" Width="194"/>
        <Button Name="btn_retry" Content="Nouvelle Fiche Planning ?" HorizontalAlignment="Left" Margin="396,323,0,0" VerticalAlignment="Top" Width="140" Height="44" Visibility="Hidden"/>
        <Button Name="btn_quitter" Content="Quitter" HorizontalAlignment="Left" Height="58" Margin="34,309,0,0" VerticalAlignment="Top" Width="194" Visibility="Hidden"/>
        <Button Name="btn_ok" Content="Valider la saisie" HorizontalAlignment="Left" Margin="396,323,0,0" VerticalAlignment="Top" Width="140" Height="44"/>
        <RichTextBox Name="txtbl_recap" HorizontalAlignment="Left" Height="47" Margin="277,271,0,0" VerticalAlignment="Top" Width="259" FontStyle="Italic" Foreground="#FF939393" FontSize="9" IsReadOnly="True" VerticalScrollBarVisibility="Auto"/>
        <Rectangle HorizontalAlignment="Left" Height="47" Margin="277,271,0,0" Stroke="Black" VerticalAlignment="Top" Width="259"/>
        <Rectangle HorizontalAlignment="Left" Height="390" Stroke="Black" VerticalAlignment="Top" Width="555"/>
    </Grid>
</Window>
"@

#region Declaration_Control	
    $global:reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $global:syncHash.window=[Windows.Markup.XamlReader]::Load($reader)

	$global:syncHash.lbl_nom = $global:syncHash.window.FindName("lbl_nom")
    $global:syncHash.lbl_modif = $global:syncHash.window.FindName("lbl_modif")
    $global:syncHash.lbl_remp = $global:syncHash.window.FindName("lbl_remp")
    $global:syncHash.lbl_nouvhor = $global:syncHash.window.FindName("lbl_nouvhor")
    $global:syncHash.lbl_durpause = $global:syncHash.window.FindName("lbl_durpause")
    $global:syncHash.lbl_horpause = $global:syncHash.window.FindName("lbl_horpause")
    $global:syncHash.lbl_horreel = $global:syncHash.window.FindName("lbl_horreel")
    $global:syncHash.lbl_inc = $global:syncHash.window.FindName("lbl_inc")
    $global:syncHash.lbl_rais = $global:syncHash.window.FindName("lbl_rais")
    $global:syncHash.lbl_recupe = $global:syncHash.window.FindName("lbl_recupe")
    $global:syncHash.lbl_tpssupp = $global:syncHash.window.FindName("lbl_tpssupp")
    $global:syncHash.lbl_debfin = $global:syncHash.window.FindName("lbl_debfin")
    $global:syncHash.lbl_tpsrest = $global:syncHash.window.FindName("lbl_tpsrest")
	
    $global:syncHash.btn_ok = $global:syncHash.window.FindName("btn_ok")
    $global:syncHash.btn_quit = $global:syncHash.window.FindName("btn_quit")
    $global:syncHash.btn_mini = $global:syncHash.window.FindName("btn_mini")
    $global:syncHash.btn_save = $global:syncHash.window.FindName("btn_save")
    $global:syncHash.btn_quitter = $global:syncHash.window.FindName("btn_quitter")
	$global:syncHash.btn_retry = $global:syncHash.window.FindName("btn_retry")
	
	$global:syncHash.txtbx_tpsrest = $global:syncHash.window.FindName("txtbx_tpsrest")
    $global:syncHash.txtbx_debfin = $global:syncHash.window.FindName("txtbx_debfin")
    $global:syncHash.txtbx_tpssup = $global:syncHash.window.FindName("txtbx_tpssup")
    $global:syncHash.txtbx_moins = $global:syncHash.window.FindName("txtbx_moins")
    $global:syncHash.txtbx_plus = $global:syncHash.window.FindName("txtbx_plus")
	$global:syncHash.txtbx_raison = $global:syncHash.window.FindName("txtbx_raison")
    $global:syncHash.txtbx_inc = $global:syncHash.window.FindName("txtbx_inc")
    $global:syncHash.txtbx_horreel = $global:syncHash.window.FindName("txtbx_horreel")
    $global:syncHash.txtbx_horpause = $global:syncHash.window.FindName("txtbx_horpause")
	$global:syncHash.txtbx_durpause = $global:syncHash.window.FindName("txtbx_durpause")
	$global:syncHash.txtbl_recap = $global:syncHash.window.FindName("txtbl_recap")

    $global:syncHash.ckbx_moins = $global:syncHash.window.FindName("ckbx_moins")
    $global:syncHash.ckbx_plus = $global:syncHash.window.FindName("ckbx_plus")
    $global:syncHash.ckbx_17 = $global:syncHash.window.FindName("ckbx_17")
    $global:syncHash.ckbx_8 = $global:syncHash.window.FindName("ckbx_8")
	
    $global:syncHash.cbbox_remp = $global:syncHash.window.FindName("cbbox_remp")
    $global:syncHash.cbbox_nom = $global:syncHash.window.FindName("cbbox_nom")
	$global:syncHash.cbbox_nom.tabindex = 0
    $global:syncHash.datepick = $global:syncHash.window.FindName("datepick")
	$global:syncHash.datepick.tabindex = 7
	$global:syncHash.datepick.selecteddate = $global:syncHash.datepick.displaydate
	
    $global:syncHash.rbt_pent = $global:syncHash.window.FindName("rbt_pent")
	$global:syncHash.rbt_pent.tabindex = 6
    $global:syncHash.rbt_reclun = $global:syncHash.window.FindName("rbt_reclun")
	$global:syncHash.rbt_reclun.tabindex = 5
    $global:syncHash.rbt_modif = $global:syncHash.window.FindName("rbt_modif")
	$global:syncHash.rbt_modif.tabindex = 4
    $global:syncHash.rbt_dec = $global:syncHash.window.FindName("rbt_dec")
	$global:syncHash.rbt_dec.tabindex = 3
    $global:syncHash.rbt_aug = $global:syncHash.window.FindName("rbt_aug")
	$global:syncHash.rbt_aug.tabindex = 2
    $global:syncHash.rbt_inv = $global:syncHash.window.FindName("rbt_inv")
	$global:syncHash.rbt_inv.tabindex = 1
	


#endregion	

#region Actions_Control
## Actions Controles
$noms = ("")
foreach ($nom in $noms){
$global:syncHash.cbbox_nom.items.add($nom)
$global:syncHash.cbbox_remp.items.add($nom) }

$utilisateur = $ENV:USERNAME
$TeleAssistance = @{
###
}

$nominus = $TeleAssistance.GetEnumerator() | where {$_.key -eq $ENV:USERNAME }
$global:syncHash.cbbox_nom.text = $nominus.Value

$global:syncHash.rbt_pent.add_Checked({
	$global:rbt_checked = "rbt_pent"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	  	$global:syncHash.lbl_tpssupp.visibility = "Visible" 
	    $global:syncHash.lbl_debfin.visibility = "Visible" 
	    $global:syncHash.lbl_tpsrest.visibility = "Visible"  
	    $global:syncHash.lbl_horreel.visibility = "Visible"
		$global:syncHash.cbbox_nom.visibility = "Visible"
		$global:syncHash.txtbx_tpsrest.visibility = "Visible"
	    $global:syncHash.txtbx_debfin.visibility = "Visible"
	    $global:syncHash.txtbx_tpssup.visibility = "Visible"
	    $global:syncHash.txtbx_horreel.visibility = "Visible"
	## Emplacement
	    $global:syncHash.lbl_debfin.margin = "249,109,0,0"
	    $global:syncHash.txtbx_debfin.margin = "420,109,0,0"
		$global:syncHash.lbl_horreel.margin = "249,140,0,0" 
		$global:syncHash.txtbx_horreel.margin = "400,140,0,0"
		$global:syncHash.lbl_tpssupp.margin = "249,170,0,0"
		$global:syncHash.txtbx_tpssup.margin = "371,170,0,0"
	    $global:syncHash.lbl_tpsrest.margin = "249,200,0,0"
	    $global:syncHash.txtbx_tpsrest.margin = "385,200,0,0" 
	## Invisible : 	
		$global:syncHash.cbbox_remp.visibility = "Hidden"
	    $global:syncHash.ckbx_moins.visibility = "Hidden"
	    $global:syncHash.ckbx_plus.visibility = "Hidden"
	    $global:syncHash.ckbx_17.visibility = "Hidden"
	    $global:syncHash.ckbx_8.visibility = "Hidden"
	    $global:syncHash.txtbx_moins.visibility = "Hidden"
	    $global:syncHash.txtbx_plus.visibility = "Hidden"
		$global:syncHash.txtbx_raison.visibility = "Hidden"
	    $global:syncHash.txtbx_inc.visibility = "Hidden"
	    $global:syncHash.txtbx_horpause.visibility = "Hidden"
		$global:syncHash.txtbx_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_remp.visibility = "Hidden"
	    $global:syncHash.lbl_nouvhor.visibility = "Hidden"
	    $global:syncHash.lbl_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_horpause.visibility = "Hidden"
	    $global:syncHash.lbl_inc.visibility = "Hidden"
	    $global:syncHash.lbl_rais.visibility = "Hidden"
	    $global:syncHash.lbl_recupe.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.txtbx_debfin.tabindex = 10
		$global:syncHash.txtbx_horreel.tabindex = 11
		$global:syncHash.txtbx_tpssup.tabindex = 12
		$global:syncHash.txtbx_tpsrest.tabindex = 13
		$global:syncHash.btn_ok.tabindex = 14
		$global:syncHash.btn_save.tabindex = 15
	},"Normal")
})

$global:syncHash.rbt_reclun.add_Checked({
	$global:rbt_checked = "rbt_reclun"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	  	$global:syncHash.lbl_tpssupp.visibility = "Visible" 
	    $global:syncHash.lbl_tpsrest.visibility = "Visible"  
	    $global:syncHash.lbl_horreel.visibility = "Visible"
		$global:syncHash.cbbox_nom.visibility = "Visible"
		$global:syncHash.txtbx_tpsrest.visibility = "Visible"
	    $global:syncHash.txtbx_tpssup.visibility = "Visible"
	    $global:syncHash.txtbx_horreel.visibility = "Visible"
	## Emplacement
	    $global:syncHash.lbl_tpssupp.margin = "249,109,0,0"
		$global:syncHash.txtbx_tpssup.margin = "371,109,0,0"
		$global:syncHash.lbl_horreel.margin = "249,140,0,0" 
		$global:syncHash.txtbx_horreel.margin = "400,140,0,0"
	    $global:syncHash.lbl_tpsrest.margin = "249,170,0,0"
	    $global:syncHash.txtbx_tpsrest.margin = "385,170,0,0" 
	## Invisible : 	
		$global:syncHash.cbbox_remp.visibility = "Hidden"
	    $global:syncHash.ckbx_moins.visibility = "Hidden"
	    $global:syncHash.ckbx_plus.visibility = "Hidden"
	    $global:syncHash.ckbx_17.visibility = "Hidden"
	    $global:syncHash.ckbx_8.visibility = "Hidden"
	    $global:syncHash.txtbx_moins.visibility = "Hidden"
	    $global:syncHash.txtbx_plus.visibility = "Hidden"
		$global:syncHash.txtbx_raison.visibility = "Hidden"
	    $global:syncHash.txtbx_inc.visibility = "Hidden"
	    $global:syncHash.txtbx_horpause.visibility = "Hidden"
		$global:syncHash.txtbx_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_remp.visibility = "Hidden"
	    $global:syncHash.lbl_nouvhor.visibility = "Hidden"
	    $global:syncHash.lbl_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_horpause.visibility = "Hidden"
	    $global:syncHash.lbl_inc.visibility = "Hidden"
	    $global:syncHash.lbl_rais.visibility = "Hidden"
	    $global:syncHash.lbl_recupe.visibility = "Hidden"
	    $global:syncHash.lbl_debfin.visibility = "Hidden" 
	    $global:syncHash.txtbx_debfin.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.txtbx_tpssup.tabindex = 10
		$global:syncHash.txtbx_horreel.tabindex = 11
		$global:syncHash.txtbx_tpsrest.tabindex = 12
		$global:syncHash.btn_ok.tabindex = 13
		$global:syncHash.btn_save.tabindex = 14
	},"Normal")
})

$global:syncHash.rbt_modif.add_Checked({
	$global:rbt_checked = "rbt_modif"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	    $global:syncHash.lbl_horreel.visibility = "Visible"
	    $global:syncHash.cbbox_nom.visibility = "Visible"
	    $global:syncHash.lbl_rais.visibility = "Visible"
		$global:syncHash.txtbx_raison.visibility = "Visible"
	    $global:syncHash.ckbx_moins.visibility = "Visible"
	    $global:syncHash.ckbx_plus.visibility = "Visible"
	    $global:syncHash.txtbx_moins.visibility = "Visible"
	    $global:syncHash.txtbx_plus.visibility = "Visible"
	    $global:syncHash.txtbx_horreel.visibility = "Visible"
	    $global:syncHash.lbl_recupe.visibility = "Visible"
	## Emplacement
	    $global:syncHash.lbl_rais.margin = "249,109,0,0"
		$global:syncHash.txtbx_raison.margin = "371,109,0,0"
		$global:syncHash.lbl_recupe.margin = "249,140,0,0" 
		$global:syncHash.ckbx_moins.margin = "371,170,0,0"
	    $global:syncHash.ckbx_plus.margin = "371,145,0,0" 
	    $global:syncHash.txtbx_moins.margin = "450,170,0,0"
	    $global:syncHash.txtbx_plus.margin = "450,145,0,0" 
		$global:syncHash.lbl_horreel.margin = "249,200,0,0"
	    $global:syncHash.txtbx_horreel.margin = "400,200,0,0"
	## Invisible : 	
		$global:syncHash.cbbox_remp.visibility = "Hidden"
	    $global:syncHash.ckbx_17.visibility = "Hidden"
	    $global:syncHash.ckbx_8.visibility = "Hidden"
	    $global:syncHash.txtbx_inc.visibility = "Hidden"
	    $global:syncHash.txtbx_horpause.visibility = "Hidden"
		$global:syncHash.txtbx_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_remp.visibility = "Hidden"
	    $global:syncHash.lbl_nouvhor.visibility = "Hidden"
	    $global:syncHash.lbl_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_horpause.visibility = "Hidden"
	    $global:syncHash.lbl_inc.visibility = "Hidden"
	  	$global:syncHash.lbl_tpssupp.visibility = "Hidden" 
	    $global:syncHash.lbl_debfin.visibility = "Hidden" 
	    $global:syncHash.lbl_tpsrest.visibility = "Hidden"  
		$global:syncHash.txtbx_tpsrest.visibility = "Hidden"
	    $global:syncHash.txtbx_debfin.visibility = "Hidden"
	    $global:syncHash.txtbx_tpssup.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.txtbx_raison.tabindex = 10
		$global:syncHash.ckbx_plus.tabindex = 11
		$global:syncHash.txtbx_plus.tabindex = 12
		$global:syncHash.ckbx_moins.tabindex = 13
		$global:syncHash.txtbx_moins.tabindex = 14
		$global:syncHash.txtbx_horreel.tabindex = 15
		$global:syncHash.btn_ok.tabindex = 16
		$global:syncHash.btn_save.tabindex = 17
	},"Normal")
})

$global:syncHash.rbt_dec.add_Checked({
	$global:rbt_checked = "rbt_dec"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	    $global:syncHash.lbl_horreel.visibility = "Visible"
	    $global:syncHash.txtbx_horreel.visibility = "Visible"
		$global:syncHash.cbbox_nom.visibility = "Visible"
	    $global:syncHash.lbl_horpause.visibility = "Visible"
	    $global:syncHash.lbl_inc.visibility = "Visible"
	    $global:syncHash.txtbx_inc.visibility = "Visible"
	    $global:syncHash.txtbx_horpause.visibility = "Visible"
		$global:syncHash.btn_ok.visibility = "Visible"
	    $global:syncHash.btn_save.visibility = "Visible"
	## Emplacement
	    $global:syncHash.lbl_horpause.margin = "249,109,0,0"
	    $global:syncHash.txtbx_horpause.margin = "371,109,0,0"
	    $global:syncHash.lbl_inc.margin = "249,140,0,0" 
	    $global:syncHash.txtbx_inc.margin = "371,140,0,0"
		$global:syncHash.lbl_horreel.margin = "249,175,0,0"
	    $global:syncHash.txtbx_horreel.margin = "400,175,0,0"
	## Invisible : 	
		$global:syncHash.cbbox_remp.visibility = "Hidden"
	    $global:syncHash.ckbx_17.visibility = "Hidden"
	    $global:syncHash.ckbx_8.visibility = "Hidden"
		$global:syncHash.txtbx_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_remp.visibility = "Hidden"
	    $global:syncHash.lbl_nouvhor.visibility = "Hidden"
	    $global:syncHash.lbl_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_recupe.visibility = "Hidden"
	  	$global:syncHash.lbl_tpssupp.visibility = "Hidden" 
	    $global:syncHash.lbl_debfin.visibility = "Hidden" 
	    $global:syncHash.lbl_tpsrest.visibility = "Hidden"  
		$global:syncHash.txtbx_tpsrest.visibility = "Hidden"
	    $global:syncHash.txtbx_debfin.visibility = "Hidden"
	    $global:syncHash.txtbx_tpssup.visibility = "Hidden"
	    $global:syncHash.lbl_rais.visibility = "Hidden"
		$global:syncHash.txtbx_raison.visibility = "Hidden"
	    $global:syncHash.ckbx_moins.visibility = "Hidden"
	    $global:syncHash.ckbx_plus.visibility = "Hidden"
	    $global:syncHash.txtbx_moins.visibility = "Hidden"
	    $global:syncHash.txtbx_plus.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.txtbx_horpause.tabindex = 10
		$global:syncHash.txtbx_inc.tabindex = 11
		$global:syncHash.txtbx_horreel.tabindex = 12
		$global:syncHash.btn_ok.tabindex = 13
		$global:syncHash.btn_save.tabindex = 14
	},"Normal")
})

$global:syncHash.rbt_aug.add_Checked({
	$global:rbt_checked = "rbt_aug"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	    $global:syncHash.lbl_horreel.visibility = "Visible"
	    $global:syncHash.txtbx_horreel.visibility = "Visible"
		$global:syncHash.cbbox_nom.visibility = "Visible"
		$global:syncHash.txtbx_durpause.visibility = "Visible"
		$global:syncHash.lbl_durpause.visibility = "Visible"
	    $global:syncHash.lbl_horpause.visibility = "Visible"
	    $global:syncHash.txtbx_horpause.visibility = "Visible"
	## Emplacement
		$global:syncHash.txtbx_durpause.margin = "371,109,0,0"
		$global:syncHash.lbl_durpause.margin = "249,109,0,0"
	    $global:syncHash.lbl_horpause.margin = "249,140,0,0" 
	    $global:syncHash.txtbx_horpause.margin = "371,140,0,0"
	    $global:syncHash.lbl_horreel.margin = "249,175,0,0"
	    $global:syncHash.txtbx_horreel.margin = "400,175,0,0"
	## Invisible : 	
		$global:syncHash.cbbox_remp.visibility = "Hidden"
	    $global:syncHash.ckbx_17.visibility = "Hidden"
	    $global:syncHash.ckbx_8.visibility = "Hidden"
	    $global:syncHash.lbl_remp.visibility = "Hidden"
	    $global:syncHash.lbl_nouvhor.visibility = "Hidden"
	    $global:syncHash.lbl_recupe.visibility = "Hidden"
	  	$global:syncHash.lbl_tpssupp.visibility = "Hidden" 
	    $global:syncHash.lbl_debfin.visibility = "Hidden" 
	    $global:syncHash.lbl_tpsrest.visibility = "Hidden"  
		$global:syncHash.txtbx_tpsrest.visibility = "Hidden"
	    $global:syncHash.txtbx_debfin.visibility = "Hidden"
	    $global:syncHash.txtbx_tpssup.visibility = "Hidden"
	    $global:syncHash.lbl_rais.visibility = "Hidden"
		$global:syncHash.txtbx_raison.visibility = "Hidden"
	    $global:syncHash.ckbx_moins.visibility = "Hidden"
	    $global:syncHash.ckbx_plus.visibility = "Hidden"
	    $global:syncHash.txtbx_moins.visibility = "Hidden"
	    $global:syncHash.txtbx_plus.visibility = "Hidden"
		$global:syncHash.lbl_inc.visibility = "Hidden"
	    $global:syncHash.txtbx_inc.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.txtbx_durpause.tabindex = 10
		$global:syncHash.txtbx_horpause.tabindex = 11
		$global:syncHash.txtbx_horreel.tabindex = 12
		$global:syncHash.btn_ok.tabindex = 13
		$global:syncHash.btn_save.tabindex = 14
	},"Normal")
})

$global:syncHash.rbt_inv.add_Checked({
	$global:rbt_checked = "rbt_inv"
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
### Visible : 
		$global:syncHash.datepick.visibility = "Visible" 
	    $global:syncHash.lbl_remp.visibility = "Visible"
	    $global:syncHash.ckbx_17.visibility = "Visible"
	    $global:syncHash.ckbx_8.visibility = "Visible"
		$global:syncHash.cbbox_remp.visibility = "Visible"
	    $global:syncHash.lbl_nouvhor.visibility = "Visible"
	## Invisible : 	
	    $global:syncHash.lbl_recupe.visibility = "Hidden"
	  	$global:syncHash.lbl_tpssupp.visibility = "Hidden" 
	    $global:syncHash.lbl_debfin.visibility = "Hidden" 
	    $global:syncHash.lbl_tpsrest.visibility = "Hidden"  
		$global:syncHash.txtbx_tpsrest.visibility = "Hidden"
	    $global:syncHash.txtbx_debfin.visibility = "Hidden"
	    $global:syncHash.txtbx_tpssup.visibility = "Hidden"
	    $global:syncHash.lbl_rais.visibility = "Hidden"
		$global:syncHash.txtbx_raison.visibility = "Hidden"
	    $global:syncHash.ckbx_moins.visibility = "Hidden"
	    $global:syncHash.ckbx_plus.visibility = "Hidden"
	    $global:syncHash.txtbx_moins.visibility = "Hidden"
	    $global:syncHash.txtbx_plus.visibility = "Hidden"
		$global:syncHash.lbl_inc.visibility = "Hidden"
	    $global:syncHash.txtbx_inc.visibility = "Hidden"
	    $global:syncHash.lbl_horreel.visibility = "Hidden"
	    $global:syncHash.txtbx_horreel.visibility = "Hidden"
		$global:syncHash.txtbx_durpause.visibility = "Hidden"
		$global:syncHash.lbl_durpause.visibility = "Hidden"
	    $global:syncHash.lbl_horpause.visibility = "Hidden"
	    $global:syncHash.txtbx_horpause.visibility = "Hidden"
	## TabIndex
		$global:syncHash.datepick.tabindex = 9
		$global:syncHash.cbbox_remp.tabindex = 10
		$global:syncHash.ckbx_8.tabindex = 11
		$global:syncHash.ckbx_17.tabindex = 12
		$global:syncHash.btn_ok.tabindex = 13
		$global:syncHash.btn_save.tabindex = 14
	},"Normal")
})

$global:syncHash.ckbx_17.add_Checked({$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.ckbx_8.Visibility="Hidden" },"Normal") })
$global:syncHash.ckbx_8.add_Checked({$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.ckbx_17.Visibility="Hidden"},"Normal")  })
$global:syncHash.ckbx_17.add_UnChecked({$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.ckbx_8.Visibility="Visible" },"Normal") })
$global:syncHash.ckbx_8.add_UnChecked({$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.ckbx_17.Visibility="Visible"},"Normal")  })

$global:syncHash.ckbx_moins.add_Checked({$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.ckbx_plus.Visibility="Hidden"
																							$global:syncHash.txtbx_plus.visibility="Hidden"},"Normal") })
$global:syncHash.ckbx_plus.add_Checked({$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.ckbx_moins.Visibility="Hidden"
																							$global:syncHash.txtbx_moins.visibility="Hidden" },"Normal")  })
$global:syncHash.ckbx_moins.add_UnChecked({$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.ckbx_plus.Visibility="Visible" 
																							$global:syncHash.txtbx_plus.visibility="Visible" },"Normal") })
$global:syncHash.ckbx_plus.add_UnChecked({$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.ckbx_moins.Visibility="Visible"
																							$global:syncHash.txtbx_moins.visibility="Visible"},"Normal")  })
#endregion

#region Fonctions_Actions_Boutons
### Remise A Zéro des données/variables
Function RaZ {
$global:count_inv = 0
$global:count_aug = 0
$global:count_dec = 0
$global:count_modif = 0
$global:count_reclun = 0
$global:count_pent = 0
$global:line_rbt_inv = 9 
$global:line_rbt_aug = 12
$global:line_rbt_dec = 16
$global:line_rbt_modif = 22 
$global:line_rbt_reclun = 31 
$global:line_rbt_pent = 14 

	$global:syncHash.txtbx_tpsrest.text = ""
    $global:syncHash.txtbx_debfin.text = ""
    $global:syncHash.txtbx_tpssup.text = ""
    $global:syncHash.txtbx_moins.text = ""
    $global:syncHash.txtbx_plus.text = ""
	$global:syncHash.txtbx_raison.text = ""
    $global:syncHash.txtbx_inc.text = ""
    $global:syncHash.txtbx_horreel.text = ""
    $global:syncHash.txtbx_horpause.text = ""
	$global:syncHash.txtbx_durpause.text = ""
	$global:syncHash.txtbl_recap.text = ""
	
	$global:rbt_inv_infos = @{}
	$global:rbt_aug_infos = @{}
	$global:rbt_dec_infos = @{}
	$global:rbt_modif_infos = @{}
	$global:rbt_reclun_infos = @{}
	$global:rbt_pent_infos = @{}
	$global:checked_rdbtn = @()
	
		$global:rbt_inv_infos.clear()
		$global:rbt_aug_infos.clear()
		$global:rbt_dec_infos.clear()
		$global:rbt_modif_infos.clear()
		$global:rbt_reclun_infos.clear()
		$global:rbt_pent_infos.clear()
		$global:checked_rdbtn.clear()
}

RaZ

$global:syncHash.btn_ok.add_click({
switch -wildcard ($rbt_checked){
	"rbt_pent" {	$global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
					$debfin = $global:syncHash.txtbx_debfin.get_text()
					$tps_rest = $global:syncHash.txtbx_tpsrest.get_text()
	 				$tps_sup = $global:syncHash.txtbx_tpssup.get_text()
					$horreel = $global:syncHash.txtbx_horreel.get_text()

					if (($date -eq $null)-or ($debfin -eq "") -or ($tps_sup -eq "") -or ($horreel -eq "") )  
							{ $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					          $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					          $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					          $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
							  $txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
					          $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					          $xml_al_form.showdialog()}
		
					else {
						$global:checked_rdbtn += "rbt_pent"
						$global:rbt_pent_infos.$global:count_pent = @()

						$global:rbt_pent_infos.$global:count_pent += $global:date 
						$global:rbt_pent_infos.$global:count_pent += $debfin
						$global:rbt_pent_infos.$global:count_pent += $horreel
						$global:rbt_pent_infos.$global:count_pent += $tps_sup
						$global:rbt_pent_infos.$global:count_pent += $tps_rest
						
						$global:count_pent++
						
						$global:date_fiche = $global:date.split(" ")[0]
						$global:syncHash.txtbl_recap.appendtext("Saisie : $date - Pentecote - $tps_sup min effectuées`r")
						$global:syncHash.txtbl_recap.scrolltoend()
					}
				 }
	"rbt_reclun" {	$global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
					$tps_rest = $global:syncHash.txtbx_tpsrest.get_text()
	 				$tps_sup = $global:syncHash.txtbx_tpssup.get_text()
					$horreel = $global:syncHash.txtbx_horreel.get_text()

					if (($date -eq $null)-or ($tps_sup -eq "") -or ($horreel -eq ""))  
							{ $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					          $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					          $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					          $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
							  $txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
					          $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					          $xml_al_form.showdialog()}
					else {		  
	
					$global:checked_rdbtn += "rbt_reclun"
					$global:rbt_reclun_infos.$global:count_reclun = @()

					$global:rbt_reclun_infos.$global:count_reclun += $global:date
					$global:rbt_reclun_infos.$global:count_reclun += $tps_sup
					$global:rbt_reclun_infos.$global:count_reclun += $horreel
					$global:rbt_reclun_infos.$global:count_reclun += $tps_rest

					$global:count_reclun++
					$global:date_fiche = $global:date.split(" ")[0]
					$global:syncHash.txtbl_recap.appendtext("Saisie : $date - RecupLundi - $tps_sup min effectuées`r")
					$global:syncHash.txtbl_recap.scrolltoend()
					}
				 }
	"rbt_modif" {	$global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
					$raison = $global:syncHash.txtbx_raison.get_text()
					$tps_plus = $global:syncHash.txtbx_moins.get_text()
					$tps_moins = $global:syncHash.txtbx_plus.get_text()
					$horreel = $global:syncHash.txtbx_horreel.get_text()

				if (($date -eq $null)-or ($raison -eq "") -or ($horreel -eq "") -or ($global:syncHash.ckbx_moins.IsChecked -eq $false -and  $global:syncHash.ckbx_plus.IsChecked -eq $false) )  
						{ $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					      $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					      $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					      $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
						  $txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
					      $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					      $xml_al_form.showdialog()}
				elseif ($tps_plus -eq "" -and $tps_moins -eq "") { $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
											                        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
											                        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
											                        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
																	$txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
											                        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
											                        $xml_al_form.showdialog() }
				else {	$global:checked_rdbtn += "rbt_modif"
						$global:rbt_modif_infos.$global:count_modif = @()
						
						$global:rbt_modif_infos.$global:count_modif += $global:date
						$global:rbt_modif_infos.$global:count_modif += $raison
						$global:rbt_modif_infos.$global:count_modif += $tps_plus
						$global:rbt_modif_infos.$global:count_modif += $tps_moins
						$global:rbt_modif_infos.$global:count_modif += $horreel

						$global:count_modif++

						$global:date_fiche = $global:date.split(" ")[0]
						if ($tps_plus -ne "") {
								$global:syncHash.txtbl_recap.appendtext("Saisie : $date - $raison ; Modif : - $tps_moins ; $horreel`r")
								$global:syncHash.txtbl_recap.scrolltoend()}
						else { $global:syncHash.txtbl_recap.appendtext("Saisie : $date - $raison ; Modif : + $tps_plus ; $horreel`r")
								$global:syncHash.txtbl_recap.scrolltoend()}
					} 
					######## Reset donnees + et -
					$global:syncHash.txtbx_moins.text = ""
    				$global:syncHash.txtbx_plus.text = ""
#					$global:syncHash.txtbl_recap.appendtext("$global:count_modif `r")
#					$global:syncHash.txtbl_recap.scrolltoend()
					########
					}
	"rbt_dec" {	$global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
				 $horpause = $global:syncHash.txtbx_horpause.get_text()
				 $inc = $global:syncHash.txtbx_inc.get_text()
				 $horreel = $global:syncHash.txtbx_horreel.get_text()
	
				if (($date -eq $null)-or ($horpause -eq "") -or ($inc -eq "") -or ($horreel -eq ""))  
						{ $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					      $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					      $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					      $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
						  $txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
					      $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					      $xml_al_form.showdialog()}
				else {
					$global:checked_rdbtn += "rbt_dec"
					$global:rbt_dec_infos.$global:count_dec = @()
					 
					$global:rbt_dec_infos.$global:count_dec += $global:date
					$global:rbt_dec_infos.$global:count_dec += $horpause
					$global:rbt_dec_infos.$global:count_dec += $inc
					$global:rbt_dec_infos.$global:count_dec += $horreel
					 
					$global:count_dec++ 

					$global:date_fiche = $global:date.split(" ")[0]
					$global:syncHash.txtbl_recap.appendtext("Saisie : $date - Décalage - Horaire : $horpause - $inc`r")
					$global:syncHash.txtbl_recap.scrolltoend()
					}
				 }
	"rbt_aug" {	$global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
				$durpause = $global:syncHash.txtbx_durpause.get_text()
				$horpause = $global:syncHash.txtbx_horpause.get_text()
				$horreel = $global:syncHash.txtbx_horreel.get_text() 

				if (($date -eq $null)-or ($durpause -eq "") -or ($horpause -eq "") -or ($horreel -eq ""))  
						{ $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
					      $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
					      $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
					      $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
						  $txtbck_alerte.text = "Tous les champs notés d'une * sont obligatoires !"
					      $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
					      $xml_al_form.showdialog()}

				else {
						$global:checked_rdbtn += "rbt_aug"
						$global:rbt_aug_infos.$global:count_aug = @()
						
						$global:rbt_aug_infos.$global:count_aug += $global:date
						$global:rbt_aug_infos.$global:count_aug += $durpause
						$global:rbt_aug_infos.$global:count_aug += $horpause
						$global:rbt_aug_infos.$global:count_aug += $horreel
						
						$global:count_aug++ 
						
					$global:date_fiche = $global:date.split(" ")[0]
					$global:syncHash.txtbl_recap.appendtext("Saisie : $date - Pause - Horaire : $horpause -Durée : $durpause`r")
					$global:syncHash.txtbl_recap.scrolltoend()
					}}
	"rbt_inv" { $global:date = get-date ($global:syncHash.datepick.selecteddate.tostring()) -Format dd/MM/yyyy
				$global:remplacant = $global:syncHash.cbbox_remp.get_text()
	
				if (($date -eq $null) -or ($remplacant -eq "") -or ($global:syncHash.ckbx_8.IsChecked -eq $false -and $global:syncHash.ckbx_17.IsChecked -eq $false) )
					{ 	$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                        $xml_al_form.showdialog()
                         }
				else {
					if ($global:syncHash.ckbx_8.IsChecked -eq $true) {$global:nouvhor = $global:syncHash.ckbx_8.content
																	$global:nouvremp = $global:syncHash.ckbx_17.content}
					if ($global:syncHash.ckbx_17.IsChecked -eq $true) {$global:nouvhor =  $global:syncHash.ckbx_17.content
																	$global:nouvremp = $global:syncHash.ckbx_8.content}

					$global:checked_rdbtn += "rbt_inv"
					$global:rbt_inv_infos.$global:count_inv = @()
					$global:rbt_inv_infos.$global:count_inv += $global:date
					$global:rbt_inv_infos.$global:count_inv += $global:remplacant
					$global:rbt_inv_infos.$global:count_inv += $global:nouvhor
					$global:rbt_inv_infos.$global:count_inv += $global:nouvremp
					
					$global:count_inv++
					
					$global:date_fiche = $global:date.split(" ")[0]
					$global:syncHash.txtbl_recap.appendtext("Saisie : $date - Inversion - Horaire : $nouvhor -Remp : $remplacant`r")
					$global:syncHash.txtbl_recap.scrolltoend()	}
					}
}

$global:syncHash.Window.Dispatcher.invoke([action]{ 
				$global:syncHash.btn_quitter.visibility = "Hidden"
				$global:syncHash.btn_save.visibility = "Visible"},"Normal")
})

$global:syncHash.btn_save.add_click({
	OpenFiche
	RaZ
	$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$global:syncHash.btn_retry.visibility = "Visible"
					$global:syncHash.btn_ok.visibility = "Hidden"				
					$global:syncHash.btn_quitter.visibility = "Visible"
					$global:syncHash.btn_save.visibility = "Hidden"},"Normal")
					
	$global:syncHash.txtbl_recap.appendtext("Fiche enregistrée dans $($saved_fiche) `r")
	$global:syncHash.txtbl_recap.scrolltoend()
})

$global:syncHash.btn_retry.add_click({
		Raz
		$global:syncHash.Window.Dispatcher.invoke([action]{ 
				$global:syncHash.btn_retry.visibility = "Hidden"
				$global:syncHash.btn_ok.visibility = "Visible"
				
				$global:syncHash.btn_quitter.visibility = "Hidden"
				$global:syncHash.btn_save.visibility = "Visible"},"Normal")

})

$global:syncHash.btn_quit.add_click({
	$global:syncHash.window.close() 
	$global:FPlanRunspace.close()
})

$global:syncHash.btn_quitter.add_click({
	$global:syncHash.window.close() 
	$global:FPlanRunspace.close()
})

$global:syncHash.btn_mini.add_click({
	$syncHash.Window.WindowState = 'Minimized'
})

$global:syncHash.Window.add_MouseLeftButtonDown({$this.DragMove() })

#endregion	
	
## Generation form    
	$global:syncHash.window.ShowDialog() 
	$global:FPlanRunspace.close()
	$global:FPlanRunspace.dispose()

}

$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
$global:psCmd.Runspace = $global:FPlanRunspace
$global:data = $global:psCmd.BeginInvoke()