##
##
## FONCTIONS ANALYSE LOG BANDEAU
##
##


#Fonction Parcourir pour import fichier de log
function Select-FileDialog {
    param([string]$Titre,[string]$Dossier,[string]$Filtre="Tous les fichiers *.*|*.*")
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $objForm = New-Object System.Windows.Forms.OpenFileDialog
    $objForm.InitialDirectory = $Directory
    $objForm.Filter = $Filter
    $objForm.Title = $Title
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK")
    {
        Return $objForm.FileName
    }
    Else
    {
        Write-Error "Opération annulé"
        return exit
    }
}



#Fonction pour récupérer les data brut du BT.log et export dans un txt
#Permet d'insérer un séparateur pour faciliter la mise en forme
Function RecupDataBrut { 
$file_data = new-item "c:\temp\BTLog_Data.txt" –type file -force
$data = Get-Content $File
$sep = "Testdata_"

Foreach ($line in $data){
			switch -wildcard ($line) {
				"*VOICE_DRIVER.SIP_GLBR*" {$linedate = $line.split(',')
											$line_sep = $sep + $linedate[0]
											$line_sep |	Out-File -Force $file_data -Encoding UTF8 -Append
											}
											
				"*'Event*" { $line |	Out-File -Force $file_data -Encoding UTF8 -Append	}

				"*'GCTI_NOT_READY_ACTIVATION'*" { $line |	Out-File -Force $file_data -Encoding UTF8 -Append	}


				"*AttributeConnID*" {$line |	Out-File -Force $file_data -Encoding UTF8 -Append }
									
				"*'DispoUSER11'*" {	$line |	Out-File -Force $file_data -Encoding UTF8 -Append }					

				"*AttributeCallType*" {	$line |	Out-File -Force $file_data -Encoding UTF8 -Append }
										

				"*AttributeOtherDN *" { $line |	Out-File -Force $file_data -Encoding UTF8 -Append }
				
				}
}}


#Fonction : création du fichier excel + import et mise en forme des données du .txt créé par la fonction RecupDataBrut
Function CreateExcelFinal {
try{
    $objExcel = new-object -comobject excel.application
    $ExcelTest = $true
}catch{
    $ExcelTest = $false
}
$data_check = Get-Content "c:\temp\BTLog_Data.txt"
# Génération de la date du jour pour le nom du fichier d'export
$Date = Get-Date -Format ddMMyyyyhhmmss
	$objExcel.Visible =$true
    # Génération du chemin du fichier d'export
    Write-Host "Creation fichier excel"
    $ExcelPath = "C:\temp\Test_BTlog_$date.xlsx"
    $finalWorkBook = $objExcel.Workbooks.Add()
	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
    $finalWorkBook.Worksheets.Item(1).Name = "Analyse Bandeau"
	$finalWorkBook.Worksheets.Item(1).Columns.item('a').NumberFormat = "jj/mm/aaaa hh:mm:ss" # modif format colonne 1
    Write-Host "Création titres"	
	$finalWorkSheet.Cells.Item(1,1) = "Date et heure"
    $finalWorkSheet.Cells.Item(1,1).Font.Bold = $True # Met le texte en gras
	$finalWorkSheet.Cells.Item(1,1).Interior.ColorIndex = 48 # Couleur de la cellule
    $finalWorkSheet.Cells.Item(1,2) = "Action";
    $finalWorkSheet.Cells.Item(1,2).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,2).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,3) = "ID Appel";
    $finalWorkSheet.Cells.Item(1,3).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,3).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,4) = "Agence Appelée"
    $finalWorkSheet.Cells.Item(1,4).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,4).Interior.ColorIndex = 48
	$finalWorkSheet.Cells.Item(1,5) = "Type d'appel"
    $finalWorkSheet.Cells.Item(1,5).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,5).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,6) = "N° Compo /N° Appelé"
    $finalWorkSheet.Cells.Item(1,6).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,6).Interior.ColorIndex = 48
	$finalWorkSheet.Cells.Item(1,7) = "Type de retrait"
    $finalWorkSheet.Cells.Item(1,7).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,7).Interior.ColorIndex = 48
	
	
    $finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
    $finalWorkBook.Worksheets.Item(2).Name = "Analyse BTlog"
	$finalWorkBook.Worksheets.Item(2).Columns.item('a').NumberFormat = "jj/mm/aaaa hh:mm:ss"
    Write-Host "Création titres"
   # Rempli la première ligne
    $finalWorkSheet.Cells.Item(1,1) = "Date et heure"
    $finalWorkSheet.Cells.Item(1,1).Font.Bold = $True # Met le texte en gras
	$finalWorkSheet.Cells.Item(1,1).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,2) = "Event";
    $finalWorkSheet.Cells.Item(1,2).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,2).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,3) = "ID Appel";
    $finalWorkSheet.Cells.Item(1,3).Font.Bold = $True	
	$finalWorkSheet.Cells.Item(1,3).Interior.ColorIndex = 48
	$finalWorkSheet.Cells.Item(1,4) = "Agence Appelée"
    $finalWorkSheet.Cells.Item(1,4).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,4).Interior.ColorIndex = 48
	$finalWorkSheet.Cells.Item(1,5) = "Type d'Appel"
    $finalWorkSheet.Cells.Item(1,5).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,5).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,6) = "N° Compo / N° Appelé"
    $finalWorkSheet.Cells.Item(1,6).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,6).Interior.ColorIndex = 48
    $finalWorkSheet.Cells.Item(1,7) = "Type de Retrait"
    $finalWorkSheet.Cells.Item(1,7).Font.Bold = $True
	$finalWorkSheet.Cells.Item(1,7).Interior.ColorIndex = 48

Write-Host "Récupération des données..." -ForegroundColor Green
$FinalExcelRow = 2 

Foreach ($line in $data_check){
			switch -wildcard ($line) {
				"*Testdata*" {	$FinalExcelRow++
							$line_test = $line.split('_')
							$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
							$finalWorkSheet.Cells.Item($FinalExcelRow,1) = $line_test[1]
											
							$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
							$finalWorkSheet.Cells.Item($FinalExcelRow,1) = $line_test[1] }
											
				"*'Event*" { switch -wildcard ($line){
									"*(73)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Agent logué"
												$finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = 4
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventAgentLogin"
												}
												
									"*(75)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Agent Prêt"
												$finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = 4
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventAgentReady" 
												}
												
									"*(74)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "LogOut" 
												$finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = 3
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventAgentLogout"
												 }
												
									"*(60)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Sonnerie" 
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventRinging"
												}
									
									"*(86)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Décroché"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventOffHook"
												}
												
									"*(87)*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Raccroché" 
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventOnHook"
												}
												
									"*(85)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												#[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "DataChanged"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												#[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "DataChanged"
												
												}	
									
									"*(64)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Communication établie"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventEstablished"
												}
									
									"*(65)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Action aboutie/terminée"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventReleased" }
									
									"*(76)*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Agent non prêt" 
												$finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = 15
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventAgentNotReady"
												}
												
									"*(59)*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Appel abandonné" 
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventAbandoned"}
									
									"*(61)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Numérotation"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventDialing"}
									
									"*(66)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Mise en attente"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Eventheld" }
												
									"*(52)*" { 	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Erreur"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventError"}
												
									"*(62)*" {  $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Passerelle de sortie atteinte"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventNetworkReached"}
											
									"*(67)*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "Action récupérée"
												
												$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
												$finalWorkSheet.Cells.Item($FinalExcelRow,2) = "EventRetrieved" 
												}
									
									}
							
							
							}

				"*'GCTI_NOT_READY_ACTIVATION'*" {$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
												$Retrait = $line.split('=')
												[string]$retraite = $Retrait[1]
												switch -wildcard ($Retrait[1]){
													"*1*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait Post Appel / RIA"
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 3
															}
													"*2*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Wrap Up"
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 53
															 }
													"*3*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Pause" 
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 46
															 }
													"*4*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait BO / AVA" 
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 3
															 }
													"*5*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait Management"
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 3
															}
													"*8*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait suite non réponse"
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 44
															}
													"*9*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait Injoignabilité"
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 44
															}
													"*no-answer*" {$finalWorkSheet.Cells.Item($FinalExcelRow,7) = "Retrait suite non réponse" 
															$finalWorkSheet.Cells.Item($FinalExcelRow,7).Interior.ColorIndex = 44
															}
													}
											$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
											$finalWorkSheet.Cells.Item($FinalExcelRow,7) = $line
												}


				"*AttributeConnID*" {$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
									$ConnID = $line.split('=')
									$finalWorkSheet.Cells.Item($FinalExcelRow,3) = $ConnID[1]
									if ($finalWorkSheet.Cells.Item($FinalExcelRow,2).value() -like "DataChanged")
											{[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
											$FinalExcelRow--}
									
									$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
									$finalWorkSheet.Cells.Item($FinalExcelRow,3) = $line 
									if ($finalWorkSheet.Cells.Item($FinalExcelRow,2).value() -like "DataChanged")
											{[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
											$FinalExcelRow--}
									}
									
				"*'DispoUSER11'*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
									$DispoUSER = $line.split('=')
									$finalWorkSheet.Cells.Item($FinalExcelRow,4) = $DispoUSER[1]
									
									if ($finalWorkSheet.Cells.Item($FinalExcelRow,2).value() -like "DataChanged")
											{[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
											$FinalExcelRow--}
									
									$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
									$finalWorkSheet.Cells.Item($FinalExcelRow,4) = $line 
																
									}					

				"*AttributeCallType*" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
									switch -wildcard ($line){
									"*Inbound*" { $finalWorkSheet.Cells.Item($FinalExcelRow,5) = "Appel Entrant"
													$finalWorkSheet.Cells.Item($FinalExcelRow,5).Interior.ColorIndex = 8}
									"*Outbound*" { $finalWorkSheet.Cells.Item($FinalExcelRow,5) = "Appel Sortant"
													$finalWorkSheet.Cells.Item($FinalExcelRow,5).Interior.ColorIndex = 33}
									"*Internal*" { $finalWorkSheet.Cells.Item($FinalExcelRow,5) = "Appel interne" 
													$finalWorkSheet.Cells.Item($FinalExcelRow,5).Interior.ColorIndex = 42}
												}
										
										if ($finalWorkSheet.Cells.Item($FinalExcelRow,2).value() -like "DataChanged")
											{[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
											$FinalExcelRow--}
									
										$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
										$finalWorkSheet.Cells.Item($FinalExcelRow,5) = $line
										}
										

				"*AttributeOtherDN *" {	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
										$OtherDN = $line.split('=')
										$finalWorkSheet.Cells.Item($FinalExcelRow,6) = $OtherDN[1]
										
										if ($finalWorkSheet.Cells.Item($FinalExcelRow,2).value() -like "DataChanged")
											{[void]$finalWorkSheet.Cells.Item($FinalExcelRow, 2).EntireRow.Delete()
											$FinalExcelRow--}
										
										$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
										$finalWorkSheet.Cells.Item($FinalExcelRow,6) = $line
										
										}
							
				}
				
			}

#Suppression de la première ligne vide
$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
[void]$finalWorkSheet.Cells.Item(2, 1).EntireRow.Delete()

$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
[void]$finalWorkSheet.Cells.Item(2, 1).EntireRow.Delete()

Write-Host "Sauvegarde du fichier" -ForegroundColor Green
    # Sélectionne les cellules utilisées
	$finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
    $UR = $finalWorkSheet.UsedRange
   # Auto ajustement de la taille de la colonne    
    $null = $UR.EntireColumn.AutoFit()
	
	$finalWorkSheet = $finalWorkBook.Worksheets.Item(2)
	$UR_1 = $finalWorkSheet.UsedRange
   # Auto ajustement de la taille de la colonne    
    $null = $UR_1.EntireColumn.AutoFit()
    if (Test-Path $ExcelPath) {
        # Si le fichier existe déjà, on le sauvegarde
        $finalWorkBook.Save()
    }else{
        # Sinon on lui donne un nom de fichier au moment de la sauvegarde
        $finalWorkBook.SaveAs($ExcelPath)
    }
    # On ferme le fichier
    #$finalWorkBook.Close()

# Le processus Excel utilisé pour traiter l'opération est arrêté
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
}


##
##
## Génération de la fenêtre
##
##

function WindowAnalyse_BTLog {
#region Import the Assemblies
[Void][reflection.assembly]::loadwithpartialname("System.Windows.Forms")
[Void][reflection.assembly]::loadwithpartialname("System.Drawing")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Generated Form Objects
$wind_bt_log = New-Object System.Windows.Forms.Form
$btn_quit = New-Object System.Windows.Forms.Button
$btn_go_btlog = New-Object System.Windows.Forms.Button
$rtxtb_btlog = New-Object System.Windows.Forms.RichTextBox
$btn_parc = New-Object System.Windows.Forms.Button
$txtb_btlog = New-Object System.Windows.Forms.TextBox
$lbl_btlog = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#______________________________________________
#______________________________________________
##
## Actions des boutons
##

$btn_parc_OnClick = {
$script:File = Select-FileDialog -Titre "Choisir le fichier de LOG" -Dossier "C:\"
$txtb_btlog.text = $file
$rtxtb_btlog.text = "Fichier sélectionné : $($txtb_btlog.get_text())"
}

$btn_quit_OnClick = {
$wind_bt_log.close()
}

$btn_go_btlog_OnClick = {
$rtxtb_btlog.appendtext("`r`rLancement de l'analyse. `rVeuillez patientez, cela peut prendre du temps...`r")
$rtxtb_btlog.appendtext("`rRécupération des données brutes...")
RecupDataBrut
$rtxtb_btlog.appendtext("`rMise en forme des données...")
CreateExcelFinal
$rtxtb_btlog.appendtext("`rMise en forme terminée !")
$rtxtb_btlog.appendtext("`r`rLe fichier excel a été enregistré dans le dossier C:\temp\")

}

$OnLoadForm_StateCorrection=
{ $wind_bt_log.WindowState = $InitialFormWindowState }

#______________________________________________
#______________________________________________
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 290
$System_Drawing_Size.Width = 460
$wind_bt_log.ClientSize = $System_Drawing_Size
$wind_bt_log.DataBindings.DefaultDataSourceUpdateMode = 0
$wind_bt_log.FormBorderStyle = 3
$wind_bt_log.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\AppDSI\EXPGLB\BANDEAUGENESYS\SITE\icone_Groupama.ico')
$wind_bt_log.MaximizeBox = $False
$wind_bt_log.Name = "wind_bt_log"
$wind_bt_log.Text = "Analyse des logs du Bandeau Genesys"
#______________________________________________
$btn_quit.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 347
$System_Drawing_Point.Y = 254
$btn_quit.Location = $System_Drawing_Point
$btn_quit.Name = "btn_quit"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 96
$btn_quit.Size = $System_Drawing_Size
$btn_quit.TabIndex = 5
$btn_quit.Text = "Quitter"
$btn_quit.UseVisualStyleBackColor = $True
$btn_quit.add_Click($btn_quit_OnClick)
$wind_bt_log.Controls.Add($btn_quit)
#______________________________________________
$btn_go_btlog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 347
$System_Drawing_Point.Y = 112
$btn_go_btlog.Location = $System_Drawing_Point
$btn_go_btlog.Name = "btn_go_btlog"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 73
$System_Drawing_Size.Width = 96
$btn_go_btlog.Size = $System_Drawing_Size
$btn_go_btlog.TabIndex = 4
$btn_go_btlog.Text = "Lancer l'analyse"
$btn_go_btlog.UseVisualStyleBackColor = $True
$btn_go_btlog.add_Click($btn_go_btlog_OnClick)
#______________________________________________
$wind_bt_log.Controls.Add($btn_go_btlog)
$rtxtb_btlog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 112
$rtxtb_btlog.Location = $System_Drawing_Point
$rtxtb_btlog.Name = "rtxtb_btlog"
$rtxtb_btlog.ReadOnly = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 166
$System_Drawing_Size.Width = 312
$rtxtb_btlog.Size = $System_Drawing_Size
$rtxtb_btlog.TabIndex = 3
$rtxtb_btlog.Text = ""
$wind_bt_log.Controls.Add($rtxtb_btlog)
#______________________________________________
$btn_parc.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 347
$System_Drawing_Point.Y = 39
$btn_parc.Location = $System_Drawing_Point
$btn_parc.Name = "btn_parc"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 96
$btn_parc.Size = $System_Drawing_Size
$btn_parc.TabIndex = 2
$btn_parc.Text = "Parcourir"
$btn_parc.UseVisualStyleBackColor = $True
$btn_parc.add_Click($btn_parc_OnClick)
$wind_bt_log.Controls.Add($btn_parc)
#______________________________________________
$txtb_btlog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 40
$txtb_btlog.Location = $System_Drawing_Point
$txtb_btlog.Name = "txtb_btlog"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 312
$txtb_btlog.Size = $System_Drawing_Size
$txtb_btlog.TabIndex = 1
$wind_bt_log.Controls.Add($txtb_btlog)
#______________________________________________
$lbl_btlog.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 13
$lbl_btlog.Location = $System_Drawing_Point
$lbl_btlog.Name = "lbl_btlog"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 269
$lbl_btlog.Size = $System_Drawing_Size
$lbl_btlog.TabIndex = 0
$lbl_btlog.Text = "Sélectionner le fichier de log à analyser : "
$wind_bt_log.Controls.Add($lbl_btlog)
#______________________________________________
#endregion Generated Form Code

$InitialFormWindowState = $wind_bt_log.WindowState
$wind_bt_log.add_Load($OnLoadForm_StateCorrection)
[Void]$wind_bt_log.ShowDialog()

}


WindowAnalyse_BTLog