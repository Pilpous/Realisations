####
#### REPARATION SSOX
####	
####	Réalisateur : ScripTeam
####
#### Arrêt SSOX + Suppression C:\Users\$ENV:USERNAME\SSOX + Relance SSOX
####

$global:ssox_directory = "C:\Users\$env:USERNAME\SSOX"
$archi_poste = [IntPtr]::Size
$scriptRoot = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($scriptRoot -eq $PSHOME.TrimEnd('\') -or $scriptRoot -eq 'C:\Program Files (x86)\PowerGUI')
{
    $scriptRoot = $PSScriptRoot
}
Write-Host $scriptRoot
Write-Host $PSScriptRoot
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms, System.Drawing
### Runspace 1 - Form
$global:syncHash = [hashtable]::Synchronized(@{})
$global:syncHash.scriptroot = $scriptRoot
$global:syncHash.ssox_directory = $ssox_directory
$global:syncHash.archi_poste = $archi_poste


$global:newRunspace =[runspacefactory]::CreateRunspace()
$global:newRunspace.ApartmentState = "STA"
$global:newRunspace.ThreadOptions = "ReuseThread"         
$global:newRunspace.Open()
$global:newRunspace.SessionStateProxy.SetVariable("syncHash",$global:syncHash)          

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
				Import-Module Logs_TA
				Create_Log -Path_Log $global:syncHash.scriptroot -Appli SSOX


				$check_watcher = Get-Process -Name watcher -ErrorAction SilentlyContinue
				$check_wait_watcher = Get-Process -Name wait_watcher -ErrorAction SilentlyContinue
				$check_xehll = Get-Process -Name Xehll3 -ErrorAction SilentlyContinue
				If ((!$check_wait_watcher) -and (!$check_watcher) -and (!$check_xehll)) 
						{ Log -Add "Watcher et Wait_Watcher non lancés"
						  Log -Add "Processus Xehll3.exe absents"
						  Log -Add "Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rXehll3, Watcher et Wait_Watcher non lancé")},"Normal") 
						} 
						
				Else { 	Log -Add "Xehll3, Watcher et Wait_Watcher lancés - Arrêt des processus !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rWatcher, Wait_Watcher et Xehll3 lancés `rArrêt des processus")},"Normal") 
						$check_wait_watcher | Stop-Process -Force -ErrorAction SilentlyContinue
						Start-Sleep -Seconds 1
						$check_watcher | Stop-Process -Force -ErrorAction SilentlyContinue
						Start-Sleep -Seconds 1
						$check_xehll | Stop-Process -Force -ErrorAction SilentlyContinue
						Log -Add "Processus arrêtés - Lancement Réparation !"
						$global:syncHash.Window.Dispatcher.invoke([action]{ $global:syncHash.rtxtb_infos.AppendText("`rProcessus arrêtés")},"Normal") 
						}
				

				Start-Sleep -Seconds 1
				if ((Test-Path $global:syncHash.ssox_directory) -eq $true) {Log -Add  "Suppression du dossier $($global:syncHash.ssox_directory)"
															$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rSuppression du dossier $($global:syncHash.ssox_directory)")},"Normal") 
															Remove-Item $global:syncHash.ssox_directory -Recurse -Force -ErrorAction SilentlyContinue 
															}
				Else {	Log -Add "Dossier $($global:syncHash.ssox_directory) non présent sur le poste"
						$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rDossier $($global:syncHash.ssox_directory) non présent sur le poste")},"Normal")  }

				Start-Sleep -Seconds 1
				Log -Add "Lancement de SSOX"
				$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`rLancement de SSOX")},"Normal") 
				start 'C:\AppDSI\EXPGLB\SSOX\GLB\Applications\wait_watcher.exe'
				start 'C:\Applics\TN3270\Xehll3.exe'
				if ($global:syncHash.archi_poste -eq 8) { start 'C:\Program Files (x86)\Avencis\SSOX\x86\Xehll3.exe' }
				else {start 'C:\Program Files\Avencis\SSOX\Xehll3.exe' }

			

			} catch { Add-Content C:\temp\log.txt -value $Error }

			Start-Sleep -Seconds 1
			Log -Add "Réparation terminée !"
			$global:syncHash.Window.Dispatcher.invoke([action]{$global:syncHash.rtxtb_infos.AppendText("`nRéparation terminée !")
																 $global:syncHash.btn_launch.visibility="Hidden"
																 $global:syncHash.btn_quit.visibility="Visible" },"Normal") 



    })

#Fin Runspace 2

$global:ThemeFile = 'C:\AppDSI\EXPGLB\SSOX\GLB\images\Theme.xaml'
[xml]$global:xa_ml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow" Height="350" Width="525" HorizontalAlignment="Center" VerticalAlignment="Center" ResizeMode="NoResize" Topmost="True" WindowStartupLocation="CenterScreen" IsManipulationEnabled="True" WindowStyle="None">
		<Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile"/> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
        </Window.Resources>
    <Grid>
        <Label Content="&#x9;       Télé Assistance" Margin="0,38,0,0" VerticalAlignment="Top" Height="43" FontSize="20" Width="525" Background="#FFFCA60D" HorizontalAlignment="Right" HorizontalContentAlignment="Center"/>
		<Label Content="Réparation SSOX" HorizontalAlignment="Left" Margin="143,92,0,0" VerticalAlignment="Top" Width="326" FontWeight="Bold" FontSize="16"/>
        <Image Margin="17,10,408,240" Source="C:\AppDSI\EXPGLB\SSOX\GLB\images\Picto_DSI.png" Stretch="Fill" Width="100" Height="100"/>
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
		$maj_Runspace.Close()
		Get-Process -Name SSOX_Reparation | Stop-Process})
		
	$global:syncHash.btn_quit.add_click({ 
		$global:syncHash.Window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close() })
	
    
	$global:syncHash.window.ShowDialog() 
	$global:newRunspace.close()
	$global:newRunspace.dispose()
	
	Get-Process -Name SSOX_Reparation | Stop-Process

}

$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
$global:psCmd.Runspace = $global:newRunspace
$global:data = $global:psCmd.BeginInvoke()