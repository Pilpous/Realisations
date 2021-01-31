####
#### REPARATION EDITO
####	
####	Réalisateur : ScripTeam
####
#### Arrêt PDFCREATOR + Suppression C:\data\Edito\.l + Message pour relancer un mail libre
####

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms, System.Drawing
### Runspace 1 - Form
$global:syncHash = [hashtable]::Synchronized(@{})
$global:newRunspace =[runspacefactory]::CreateRunspace()
$global:newRunspace.ApartmentState = "STA"
$global:newRunspace.ThreadOptions = "ReuseThread"         
$global:newRunspace.Open()
$global:newRunspace.SessionStateProxy.SetVariable("syncHash",$global:syncHash)          
#$psCmd = [PowerShell]::Create().AddScript({  
$global:edito_directory = "C:\data\Edito\logs\"
$global:edito_lock = "C:\data\Edito\.lock"
$global:File_log = "C:\Windows\Log\Réparation\Log_Edito.log"
if ( (Test-path $File_log) -eq $false) {New-Item $File_log -Force -ItemType File}

$global:script_code = {



### Runspace 2 - Maj Form + Action
	$maj_Runspace = [runspacefactory]::CreateRunspace()
	$maj_Runspace.ApartmentState = "MTA"
    $maj_Runspace.ThreadOptions = "ReuseThread"          
    $maj_Runspace.Open()
    $maj_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

    $psScript = [PowerShell]::Create().AddScript({
## Action à définir
			try { 	
				$global:File_log = "C:\Windows\Log\Réparation\Log_Edito.log"
				Function Log {
							Param ([string]$Add)
						$Date = Get-Date -Format yyyyMMdd_hh:mm:ss	

						Add-Content $File_log -Value "[$($Date)] $($Add)"
						#Write-Host "[$($Date)] $($Add)"
						}
				$global:edito_directory = "C:\data\Edito"
				$global:syncHash.Window.Dispatcher.invoke([action]{ },"Normal")
				Test-Path $edito_directory

				$Check_PdfCreator = Get-Process -Name PdfCreator -ErrorAction SilentlyContinue
				If (!$Check_PdfCreator)  
						{ Log -Add "PdfCreator non lancé - Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rPdfCreator non lancé ")},"Normal") 
						} 
						
				Else { 	Log -Add "PdfCreator lancé - Arrêt du processus !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rPdfCreator lancé - Arrêt du processus !")},"Normal") 
						$Check_PdfCreator | Stop-Process -Force -ErrorAction SilentlyContinue
						Log -Add "Processus arrêtés - Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rProcessus arrêté")},"Normal") 
						}
				

				Start-Sleep -Seconds 1
			
				$global:edito_lock = "C:\data\Edito\.lock"
				if ((Test-Path $edito_lock) -eq $True) 
						{Remove-item	$edito_lock -Force -Debug -Verbose
						Log -Add  "Suppression fichier Lock"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rSuppression fichier Lock")},"Normal") 
						}
						
				Else { 	Log -Add "Pas de fichier Lock"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rPas de fichier Lock")},"Normal") 
						}	
						
				$edito_directory = "C:\data\Edito\Logs\"
				if ((Test-Path $edito_directory) -eq $true) 
						{Log -Add  "Purge des Logs Edito sous: $($edito_directory)"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rPurge des Logs Edito sous: $($edito_directory)")},"Normal") 
						Get-ChildItem -Path $edito_directory -Filter *.log | Where-Object {$_.LastWriteTime -lt ((get-date).AddMonths(-1))} | Remove-Item -Force
						}
						
				Else {	Log -Add "Dossier $($edito_directory) non présent sur le poste"
						$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rDossier $($edito_directory) non présent sur le poste")},"Normal")  }

			} catch { Add-Content C:\temp\log.txt -value $Error }

			Start-Sleep -Seconds 1
			Log -Add "Réparation terminée !"
			$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`nRéparation terminée !")
																 $global:syncHash.btn_launch.visibility="Hidden"
																 $global:syncHash.btn_quit.visibility="Visible" },"Normal")
			$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`nMerci de lancer un Mail Libre afin de faire remonter les documents en attente")
																 $global:syncHash.btn_launch.visibility="Hidden"
																 $global:syncHash.btn_quit.visibility="Visible" },"Normal") 



    })

#Fin Runspace 2

#$global:ThemeFile = 'C:\temp\packages\ExpressionDark\Theme.xaml'
[xml]$global:xa_ml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="350" Width="525" HorizontalAlignment="Center" VerticalAlignment="Center" ResizeMode="NoResize" Topmost="True" WindowStartupLocation="CenterScreen" IsManipulationEnabled="True" WindowStyle="None">
    <Grid>
        <Label Content="&#x9;       Télé Assistance - Groupama Loire Bretagne" Margin="0,38,0,0" VerticalAlignment="Top" Height="43" FontSize="20" Width="525" Background="#FFFCA60D" HorizontalAlignment="Right" HorizontalContentAlignment="Center"/>
		<Label Content="Réparation EDITO" HorizontalAlignment="Left" Margin="143,92,0,0" VerticalAlignment="Top" Width="326" FontWeight="Bold" FontSize="16"/>
        <Image Margin="17,10,408,240" Source="C:\AppDSI\EXPGLB\POPUPMDP\Images\Picto-DSI-3.png  " Stretch="Fill" Width="100" Height="100"/>
        <Button Name="btn_quit" Content="Quitter" Visibility="Hidden" HorizontalAlignment="Left"  Margin="390,274,0,0" VerticalAlignment="Top" Width="113" Height="55"/>
        <Button Name="btn_launch" Content="Lancer la réparation" HorizontalAlignment="Left"  Margin="390,274,0,0" VerticalAlignment="Top" Width="113" Height="55"/>
		<Button Name="btn_close" Content="X" HorizontalAlignment="Left" Margin="496,10,0,0" VerticalAlignment="Top" Width="19"/>
        <RichTextBox Name="rtxtb_infos" HorizontalAlignment="Left" Height="145" Margin="143,128,0,0" VerticalAlignment="Top" Width="360" IsReadOnly="True" BorderThickness="0">
            <FlowDocument>
                <Paragraph>
                    <Run Text=""/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <!-- <ProgressBar Name="pgbar" HorizontalAlignment="Left" Height="22" Margin="143,307,0,0" VerticalAlignment="Top" Width="218"/> -->
		<Rectangle HorizontalAlignment="Left" Height="350" Stroke="Black" StrokeThickness="2" VerticalAlignment="Top" Width="525"/>
    </Grid>
</Window>
"@  
    $global:reader=(New-Object System.Xml.XmlNodeReader $xa_ml)
    $global:syncHash.window=[Windows.Markup.XamlReader]::Load($reader)
	
    $global:syncHash.rtxtb_infos = $global:syncHash.window.FindName("rtxtb_infos")
    $global:syncHash.btn_launch = $global:syncHash.window.FindName("btn_launch")
	$global:syncHash.btn_close = $global:syncHash.window.FindName("btn_close")
	$global:syncHash.btn_quit = $global:syncHash.window.FindName("btn_quit")

	
	$global:syncHash.btn_launch.add_click({
		$psScript.Runspace = $maj_Runspace
      	$psScript.BeginInvoke() 
		})
	
	$global:syncHash.btn_close.add_click({ 
		$global:syncHash.Window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close() })
		
	$global:syncHash.btn_quit.add_click({ 
		Get-Process Reparation_Edito | Stop-Process -force
		$global:syncHash.Window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close()
		})

	
    
	$global:syncHash.window.ShowDialog() 
	$global:newRunspace.close()
	$global:newRunspace.dispose()

}

$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
$global:psCmd.Runspace = $global:newRunspace
$global:data = $global:psCmd.BeginInvoke()