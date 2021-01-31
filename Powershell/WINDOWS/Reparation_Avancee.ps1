####
#### REPARATION AVANCEE
####	
####	Réalisateur : ScripTeam
####
#### 
####

# Déclaration $PSScriptRoot pour Runspace
$scriptRoot = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
if ($scriptRoot -eq $PSHOME.TrimEnd('\') -or $scriptRoot -eq 'C:\Program Files (x86)\PowerGUI')
{
    $scriptRoot = $PSScriptRoot
}

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms, System.Drawing
### Runspace 1 - Form
$global:syncHash = [hashtable]::Synchronized(@{})
$global:syncHash.scriptroot = $scriptRoot
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
			try {
			$name = $global:syncHash.scriptroot.Split('\')[$global:syncHash.scriptroot.Split('\').count - 1]
			Import-Module Logs_TA
			Create_Log -Path_Log $global:syncHash.scriptroot -Appli $name 
			Log -add "Lancement de la réparation"
			Log -Add "PSSCriptRoot : $($global:syncHash.scriptroot)"
			}
			catch {Add-Content "$($global:syncHash.scriptroot)\error_log.txt" -value $Error[0] }

			try { 	
				$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$syncHash.lbl_go_rep.Visibility = "Visible"
					$syncHash.lbl_stop_process.Visibility = "Visible"},"Normal")
				get-process iexplore -ErrorAction SilentlyContinue  | stop-process -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				log -add "Arrêt process iexplore.exe"
				get-process chrome -ErrorAction SilentlyContinue  | stop-process -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				log -add "Arrêt process chrome.exe"
				get-process rundll32 -ErrorAction SilentlyContinue  | stop-process -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				log -add "Arrêt process rundll32.exe"

				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__1.Visibility = "Visible"
					$syncHash.lbl__suppr_temp_ie.Visibility = "Visible" },"Normal")
			
				RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351
				log -add "Suppression fichiers temporaires IE"
			
				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__2.Visibility = "Visible"
					$syncHash.lbl_suppr_temp_win.Visibility = "Visible"},"Normal")
				try {
				Get-process rundll32 -ErrorAction SilentlyContinue  | stop-process -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				log -add "Arrêt process rundll32.exe"
				
				
				$cache_cache = Get-Content "$($global:syncHash.scriptroot)\path_temp.txt"
				## Si nouveaux éléments à purger, rajouter le chemin dans le fichier texte
				foreach ($path_cache in $cache_cache){ 
						$test_path = Test-Path $path_cache -ErrorAction SilentlyContinue
						if ($test_path -eq $true ) {Get-ChildItem $path_cache | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue  }
						}
				<#
				Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Log -Add "Suppression des fichiers ds C:\Windows\Temp\"
								
				Remove-Item -path "C:\Users\$env:username\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Recurse -Force -ErrorAction SilentlyContinue
				Remove-Item -path "C:\Users\$env:username\AppData\Local\Temp\*" -Recurse -Force -ErrorAction Ignore 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Microsoft\Internet Explorer\Recovery\High\Active\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Microsoft\Internet Explorer\Recovery\High\Last Active\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Roaming\Microsoft\Windows\Cookies\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Microsoft\Windows\Explorer\ThumbCacheToDelete\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\CrashDumps\*" -Recurse -Force -ErrorAction SilentlyContinue 
				
				##test##
				Remove-Item -Path "C:\biduletruc\machinchose\*" -Recurse -Force	-ErrorAction SilentlyContinue
				
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default\Media Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default\GPUCache\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default\Application Cache\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\Default\Service Worker\CacheStorage\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Google\Chrome\User Data\ShaderCache\GPUCache\*" -Recurse -Force -ErrorAction SilentlyContinue
				
				Remove-Item -Path "C:\Users\$env:username\AppData\Local\Adobe\Acrobat\11.0\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
				Remove-Item -Path "C:\Users\$env:username\AppData\LocalLow\Adobe\Acrobat\11.0\Search\*" -Recurse -Force -ErrorAction SilentlyContinue
				Remove-Item -Path "C:\Users\$env:username\.thumbnails\normal\*" -Recurse -Force -ErrorAction SilentlyContinue
				Remove-item -Path "C:\Users\$env:username\AppData\Local\Microsoft\Terminal Server Client\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue 
				Remove-Item -Path "C:\$Recycle.Bin" -Recurse -Force -ErrorAction SilentlyContinue
				
				Log -Add "Suppression des fichiers temporaires présents dans C:\Users\$env:username\"
				#>
				
				Get-ChildItem -Path C:\data\Edito\Logs\ -Filter *.log | Where-Object {$_.LastWriteTime -lt ((get-date).AddMonths(-1))} | Remove-Item -Force
				Log -Add "Suppression des fichiers de logs dans C:\data\EDITO\Logs\ supérieur à 1 mois"
				
				#Vider la corbeille
				$shell = New-Object -comObject Shell.Application
				$recycler = $shell.Namespace(0xA)
					foreach ($item in $recycler.Items()){
					remove-item $item.path -Force -Recurse }
				Log -Add "Vidage de la corbeille"

				
				$log_gen = Get-ChildItem -Path C:\AppDSI\EXPGLB\BANDEAUGENESYS\SITE\log , C:\AppDSI\EXPGLB\BANDEAUGENESYS\AGENCE\log  | where {$_.LastWriteTime -LE (Get-Date).addDays(-15) -AND $_.Name -ne 'aconserver.txt'}
				$log_bao = Get-ChildItem -Path C:\AppDSI\EXPGLB\OUTILSTA\log ,C:\AppDSI\EXPGLB\OUTILSTA\log\syslog ,C:\AppDSI\EXPGLB\OUTILSTA\log\CMtrace | where {$_.LastWriteTime -LE (Get-Date).addDays(-15) }
				$log_rep_IE = Get-ChildItem -Path C:\AppDSI\DEVGLB\OUTILSTA\Reparation\RepIE | where {$_.LastWriteTime -LE (Get-Date).addDays(-15) }
				$log_netqos = Get-ChildItem -Path C:\AppDSI\EXPGLB\NETQOS\Logs , C:\AppDSI\EXPGLB\NETQOS\LogLoc | where {($_.LastWriteTime -LE (Get-Date).addDays(-15)) -and ($_.Name -ne "Ne_Pas_Supprimer.txt")  } 
				if ($log_gen -ne $null) { $log_gen | Remove-Item }
				if ($log_bao -ne $null) { $log_bao | Remove-Item }
				if ($log_rep_IE -ne $null) { $log_rep_IE | Remove-Item }
				if ($log_netqos -ne $null) { $log_netqos | Remove-Item }
				
				Log -Add "Suppression des fichiers de logs GPHONE, NETQOS, BAO et Réparation IE supérieurs à 15 jours"

				Wait-Process -name rundll32 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
				} catch {Add-Content "$($run_path)\error_log.txt" -value $Error[0] }
				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__3.Visibility = "Visible"
					$syncHash.lbl_reinit_ie.Visibility = "Visible"},"Normal")
					
				Set-itemproperty -Path Registry::'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Zoom' -Name "ZoomFactor" -Value 100000
				Remove-ItemProperty -Path Registry::'HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main' -Name "Window_Placement"
				Log -Add "Modification de la clef de registre ZOOM IE"
				Log -Add "Suppression de la clef de registre Window_Placement IE"
				
				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__4.Visibility = "Visible"
					$global:syncHash.lbl_flush_dns.Visibility = "Visible" },"Normal")
				
				ipconfig /flushdns
				Log -Add "Lancement ipconfig /flushdns"
				
				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__5.Visibility = "Visible"
					$global:syncHash.lbl_rep_edito.Visibility = "Visible" },"Normal")
					
				Get-process -Name PdfCreator -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
				Log -Add  "Arret des processus PDFCreator"				
				$edito_lock = "C:\data\Edito\.lock"
				if ((Test-Path $edito_lock) -eq $True) 
						{Remove-item $edito_lock -Force
						Log -Add  "Suppression fichier C:\data\Edito\.lock" }
				Else { 	Log -Add "Pas de fichier C:\data\Edito\.lock" }	

				$global:syncHash.Window.Dispatcher.invoke([action]{
					$syncHash.img__6.Visibility = "Visible" },"Normal")

			} catch { Add-Content "$($global:syncHash.scriptroot)\error_log.txt" -value $Error[0] }

			Start-Sleep -Seconds 1
			$global:syncHash.Window.Dispatcher.invoke([action]{
			$global:syncHash.txtb_rep_ok.Visibility = "Visible"
			$syncHash.btn__fermer.Visibility = "Visible"
			$syncHash.btn_go.Visibility = "Hidden" },"Normal")
			Log -add "Fin de la réparation"
    })


$global:ThemeFile = $global:syncHash.scriptroot+'\Images\Theme.xaml'
$global:img_check = $global:syncHash.scriptroot+'\Images\check_ok.ico'

[xml]$global:xa_ml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   	    Title="Reparation Avancee" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Height="360" Width="525" WindowStyle="None" Foreground="Black">
		<Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile"/> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
        </Window.Resources>
    <Grid>
        <Label Content="&#x9;       Réparation Avancée" Margin="0,12,0,0" VerticalAlignment="Top" Height="43" FontSize="20" Width="525" Background="#FFFCA60D" HorizontalAlignment="Right"/>
        <Image Margin="10,6,415,254" Source="C:\AppDSI\EXPGLB\POPUPMDP\Images\Picto-DSI-3.png  " Stretch="Fill" Width="100" Height="100"/>
        <Button Content="Lancer la Réparation" Name="btn_go" HorizontalAlignment="Left" Margin="341,273,0,0" VerticalAlignment="Top" Width="166" Height="68"/>
        <Button Content="X" Name="btn_quit" HorizontalAlignment="Left" Height="22" Margin="486,7,0,0" VerticalAlignment="Top" Width="21"/>
        <Button Content="_" Name="btn_mini" HorizontalAlignment="Left" Height="22" Margin="460,7,0,0" VerticalAlignment="Top" Width="21"/>
        <Button Name="btn__fermer" Content="Quitter" HorizontalAlignment="Left" Height="68" Margin="341,273,0,0" VerticalAlignment="Top" Width="166" Visibility="Hidden"/>
        <TextBlock HorizontalAlignment="Left" Height="36" Margin="115,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="400"><Run Text="La Réparation va fermer "/><Run Text="votre/vos navigateur/s internet"/><Run Text=". Etes-vous certain de vouloir continuer ?"/></TextBlock>

        <Label Name="lbl_go_rep" Content="Lancement de la réparation..." HorizontalAlignment="Left" Margin="94,111,0,0" VerticalAlignment="Top" Width="179" FontWeight="Bold" Visibility="Hidden"/>
        <Label Name="lbl_stop_process" Content="Arrêt des processus IE et Chrome" HorizontalAlignment="Left" Margin="177,140,0,0" VerticalAlignment="Top" Width="263" Foreground="Black" Visibility="Hidden"/>
        <Label Name="lbl__suppr_temp_ie" Content="Suppression des fichiers temporaires internet" HorizontalAlignment="Left" Margin="177,160,0,0" VerticalAlignment="Top" Width="263" Foreground="Black" Visibility="Hidden"/>
        <Label Name="lbl_suppr_temp_win" Content="Suppression des fichiers temporaires Windows" HorizontalAlignment="Left" Margin="177,180,0,0" VerticalAlignment="Top" Width="263" Height="31" Visibility="Hidden"/>
        <Label Name="lbl_reinit_ie" Content="Réinitialisation des paramètres IE" HorizontalAlignment="Left" Margin="177,200,0,0" VerticalAlignment="Top" Width="263" Visibility="Hidden"/>
        <Label Name="lbl_flush_dns" Content="Purge du cache DNS" HorizontalAlignment="Left" Margin="177,220,0,0" VerticalAlignment="Top" Width="263" Visibility="Hidden"/>
        <Label Name="lbl_rep_edito" Content="Réparation Edito" HorizontalAlignment="Left" Margin="177,240,0,0" VerticalAlignment="Top" Width="263" Visibility="Hidden"/>

        <TextBlock Name="txtb_rep_ok" HorizontalAlignment="Left" Height="36" Margin="22,277,0,0" TextWrapping="Wrap" Text="Réparation Terminée ! Vous pouvez de nouveau utiliser votre navigateur préféré." VerticalAlignment="Top" Width="314" Visibility="Hidden" FontWeight="Bold"/>

        <Image Name="img__1" HorizontalAlignment="Left" Height="13" Margin="153,146,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>
        <Image Name="img__2" HorizontalAlignment="Left" Height="13" Margin="153,166,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>
        <Image Name="img__3" HorizontalAlignment="Left" Height="13" Margin="153,186,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>
        <Image Name="img__4" HorizontalAlignment="Left" Height="13" Margin="153,206,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>
        <Image Name="img__5" HorizontalAlignment="Left" Height="13" Margin="153,226,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>
        <Image Name="img__6" HorizontalAlignment="Left" Height="13" Margin="153,246,0,0" VerticalAlignment="Top" Width="13" Source="$img_check" Visibility="Hidden"/>

        <Rectangle HorizontalAlignment="Left" Height="360" Width="525" Stroke="Black" StrokeThickness="2" VerticalAlignment="Top"/>

    </Grid>
</Window>
"@  
    $global:reader=(New-Object System.Xml.XmlNodeReader $xa_ml)
    $global:syncHash.window=[Windows.Markup.XamlReader]::Load($reader)
	
    $global:syncHash.btn_go = $global:syncHash.window.FindName("btn_go")
    $global:syncHash.btn_quit = $global:syncHash.window.FindName("btn_quit")
    $global:syncHash.btn_mini = $global:syncHash.window.FindName("btn_mini")
	
    $global:syncHash.lbl_stop_process = $global:syncHash.window.FindName("lbl_stop_process")
    $global:syncHash.lbl__suppr_temp_ie = $global:syncHash.window.FindName("lbl__suppr_temp_ie")
	$global:syncHash.lbl_suppr_temp_win = $global:syncHash.window.FindName("lbl_suppr_temp_win")
    $global:syncHash.lbl_reinit_ie = $global:syncHash.window.FindName("lbl_reinit_ie")
	$global:syncHash.lbl_go_rep = $global:syncHash.window.FindName("lbl_go_rep")
	$global:syncHash.lbl_flush_dns = $global:syncHash.window.FindName("lbl_flush_dns")
	$global:syncHash.lbl_rep_edito = $global:syncHash.window.FindName("lbl_rep_edito")

	$global:syncHash.txtb_rep_ok = $global:syncHash.window.FindName("txtb_rep_ok")

    $global:syncHash.img__1 = $global:syncHash.window.FindName("img__1")
    $global:syncHash.img__2 = $global:syncHash.window.FindName("img__2")
    $global:syncHash.img__3 = $global:syncHash.window.FindName("img__3")
    $global:syncHash.img__4 = $global:syncHash.window.FindName("img__4")
    $global:syncHash.img__5 = $global:syncHash.window.FindName("img__5")
    $global:syncHash.img__6 = $global:syncHash.window.FindName("img__6")
	
    $global:syncHash.btn__fermer = $global:syncHash.window.FindName("btn__fermer")
	
	$global:syncHash.btn_quit.add_click({
		$global:syncHash.window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close()
		
		Get-Process Reparation_Avancee | Stop-Process -Force -ErrorAction SilentlyContinue
		})
	
	$global:syncHash.btn__fermer.add_click({ 
		$global:syncHash.Window.Close()
		$newRunspace.Close()
		$maj_Runspace.Close()

		Get-Process Reparation_Avancee | Stop-Process -Force -ErrorAction SilentlyContinue
		})
	
	$global:syncHash.btn_mini.add_click({ 
		$syncHash.Window.WindowState = 'Minimized'})

	$global:syncHash.btn_go.add_click({ 
		$psScript.Runspace = $maj_Runspace
      	$psScript.BeginInvoke() 
		
		})
	
    
	$global:syncHash.window.ShowDialog() 
	$global:newRunspace.close()
	$global:newRunspace.dispose()

}

$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
$global:psCmd.Runspace = $global:newRunspace
$global:data = $global:psCmd.BeginInvoke()