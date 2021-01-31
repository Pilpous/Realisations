####
#### REPARATION SSOX
####	
####	Réalisateur : ScripTeam
####
#### Arrêt SSOX + Suppression C:\Users\$ENV:USERNAME\SSOX + Relance SSOX
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
$global:ssox_directory = "C:\Users\$env:USERNAME\SSOX"
$global:File_log = "C:\AppDSI\DEVGLB\OUTILSTA\Reparation\SSOX\Log_SSOX.txt"
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
				$global:File_log = "C:\AppDSI\DEVGLB\OUTILSTA\Reparation\SSOX\Log_SSOX.txt"
				Function Log {
							Param ([string]$Add)
						$Date = Get-Date -Format yyyyMMdd_hh:mm:ss	

						Add-Content $File_log -Value "[$($Date)] $($Add)"
						#Write-Host "[$($Date)] $($Add)"
						}
				$global:ssox_directory = "C:\Users\$env:USERNAME\SSOX"
				$global:syncHash.Window.Dispatcher.invoke([action]{ },"Normal")
				Test-Path $ssox_directory

				$check_watcher = Get-Process -Name watcher -ErrorAction SilentlyContinue
				$check_wait_watcher = Get-Process -Name wait_watcher -ErrorAction SilentlyContinue
				If ((!$check_wait_watcher) -and (!$check_watcher)) 
						{ Log -Add "Watcher et Wait_Watcher non lancé - Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rWatcher et Wait_Watcher non lancé")},"Normal") 
						} 
						
				Else { 	Log -Add "Watcher et Wait_Watcher lancés - Arrêt des 2 processus !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rWatcher et Wait_Watcher lancés - Arrêt des 2 processus")},"Normal") 
						$check_wait_watcher | Stop-Process -Force -ErrorAction SilentlyContinue
						Start-Sleep -Seconds 1
						$check_watcher | Stop-Process -Force -ErrorAction SilentlyContinue
						Log -Add "Processus arrêtés - Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rProcessus arrêtés")},"Normal") 
						}
				

				Start-Sleep -Seconds 1
				$ssox_directory = "C:\Users\$env:USERNAME\SSOX"
				if ((Test-Path $ssox_directory) -eq $true) {Log -Add  "Suppression du dossier $($ssox_directory)"
															$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rSuppression du dossier $($ssox_directory)")},"Normal") 
															#Remove-Item $ssox_directory -Force -ErrorAction SilentlyContinue 
															}
				Else {	Log -Add "Dossier $($ssox_directory) non présent sur le poste"
						$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rDossier $($ssox_directory) non présent sur le poste")},"Normal")  }

				Start-Sleep -Seconds 1
				Log -Add "Lancement de SSOX"
				$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rLancement de SSOX")},"Normal") 
				#start 'C:\AppDSI\DEVGLB\OUTILSTA\Reparation\wait_watcher.exe'
			

			} catch { Add-Content C:\temp\log.txt -value $Error }

			Start-Sleep -Seconds 1
			Log -Add "Réparation terminée !"
			$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`nRéparation terminée !")
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
        <Label Content="&#x9;       Télé Assistance" Margin="0,38,0,0" VerticalAlignment="Top" Height="43" FontSize="20" Width="525" Background="#FFFCA60D" HorizontalAlignment="Right" HorizontalContentAlignment="Center"/>
		<Label Content="Réparation SSOX" HorizontalAlignment="Left" Margin="143,92,0,0" VerticalAlignment="Top" Width="326" FontWeight="Bold" FontSize="16"/>
        <Image Margin="17,10,408,240" Source="C:\AppDSI\EXPGLB\POPUPMDP\Images\Picto-DSI-3.png  " Stretch="Fill" Width="100" Height="100"/>
        <Button Name="btn_quit" Content="Quitter" Visibility="Hidden" HorizontalAlignment="Left"  Margin="390,274,0,0" VerticalAlignment="Top" Width="113" Height="55"/>
        <Button Name="btn_launch" Content="Lancer la réparation" HorizontalAlignment="Left"  Margin="390,274,0,0" VerticalAlignment="Top" Width="113" Height="55"/>
		<Button Name="btn_close" Content="X" HorizontalAlignment="Left" Margin="496,10,0,0" VerticalAlignment="Top" Width="19"/>
        <RichTextBox Name="rtxtb_infos" HorizontalAlignment="Left" Height="141" Margin="143,128,0,0" VerticalAlignment="Top" Width="360" IsReadOnly="True" BorderThickness="0">
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
		$global:syncHash.Window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close() })
	
    
	$global:syncHash.window.ShowDialog() 
	$global:newRunspace.close()
	$global:newRunspace.dispose()

}

$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
$global:psCmd.Runspace = $global:newRunspace
$global:data = $global:psCmd.BeginInvoke()