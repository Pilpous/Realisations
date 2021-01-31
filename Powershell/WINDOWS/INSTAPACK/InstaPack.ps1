##
##
## SCRIPT Install Package SCCM
##		Nom : InstaPack
## 		Réalisateur : ScripTeam
##
##


#region LOG
##LOGS##
$Path_Logs = 'C:\AppDSI\EXPGLB\SCCM\INSTAPACK\Logs\Log_SCCM.txt'
$script:path_sccm = '\\SCCMP029\SCCMContener\Packages\EXPGLB'
if((Test-Path $Path_Logs) -ne $true){New-Item $Path_Logs -Force -ItemType File }
$Date = Get-Date -Format ddMMyyyy_hh:mm:ss
$count_copie = 0
$Trigramme = (Get-CimInstance -ClassName Win32_ComputerSystem).username.split("\")[1]
$iii = 1
$script:environnement = "_EXP"
#endregion

#region Déclaration des fonctions
Function Log {
	Param ([string]$Add)
$Date = Get-Date -Format yyyyMMdd_hh:mm:ss	

Add-Content $Path_Logs -Value "[$($Date)] $($Add)"
Write-Host "[$($Date)] $($Add)"
}

Function Rech_Package {
	try {
	$script:RechSCCM = Get-ChildItem -Force -PATH $path_sccm -ErrorAction SilentlyContinue | Where-Object {($_.Name -like "*$package*")} | Select-Object -Property Name,FullName -ExpandProperty FullName
	if ($RechSCCM -eq $null) { $list_result.AddText("Aucun package trouvé")
								Log -Add "Aucun package trouvé"}
	Else {
		foreach ($script:rez in $RechSCCM ) 	{$list_result.AddText($($rez.Name))
												$list_result.RenderSize
												$list_result.VerticalContentAlignment }
		}
	}
	catch { Write-Host $_.exception.message }
}

Function Affichage { 
	$list_result.items.Clear()
	$list_detail.items.Clear()
	$script:aff_sccm = Get-ChildItem -Force -PATH $path_sccm -ErrorAction SilentlyContinue | Select-Object -Property Name,FullName -ExpandProperty FullName
	foreach ($script:line in $script:aff_sccm) { 
		$list_result.AddText($($line.Name))
		$list_result.RenderSize
		$list_result.VerticalContentAlignment } 
}

Function InstallPackage {
	#Set-Location $dest
	try {
	$script:launch = Get-ChildItem -force -Path $dest | Where-Object {($_.Name -like "*_Install.bat")}
	Start-Process powershell.exe -NoNewWindow -ArgumentList "start-process $dest\$launch"
	#Write-Host $dest+"\"+$launch
	timeout /t 2
	} catch { Write-Host $_.exception.message }
}

Function UninstallPackage {
	#Set-Location $dest
	try {
	$script:launch = Get-ChildItem -force -Path $dest | Where-Object {($_.Name -like "*_Uninstall.bat")}
	Start-Process powershell.exe -NoNewWindow -ArgumentList "start-process $dest\$launch"
	timeout /t 2
	} catch { Log -add $_.exception.message }
}

Function EffacerLesPreuves {
	Clear-Variable package_selected -Scope  script
	$var = Get-CimInstance -ClassName win32_computersystem | select Username # -Property Username
	$userne = ($var.Username.Split("\"))[1]
	Set-Location C:\Users\$($userne)
	$Date_suppr = Get-Date -Format yyyyMMdd
	$suppr = Get-Content C:\AppDSI\EXPGLB\SCCM\INSTAPACK\Logs\Log_SCCM.txt | where {($_ -match "Copie") -and ($_ -match "$($Date_suppr)")}
	foreach ( $li in $suppr){
		$Asuppr = ($li.Split(" "))[$li.Split(" ").count -1 ]
		if ((Test-Path $Asuppr) -eq $true){
						if ($Asuppr -ne 'C:\temp\') {
						Log -add "Suppression du dossier $($Asuppr)"
						Remove-Item $Asuppr.trim() -Force -Recurse
									}
						}
		}
}

Function Select_Package {
	$Date= Get-Date -Format yyyyMMdd #ddMMyyyy
	$temp = Get-Content C:\AppDSI\EXPGLB\SCCM\INSTAPACK\Logs\Log_SCCM.txt | where {($_ -match "Copie") -and ($_ -match "$($Date)")}
	foreach ($dos in $temp){
		$doss = ($dos.Split("\"))[$dos.Split("\").count - 1] 
		if ((Test-Path "C:\temp\$($doss)") -eq $true) {
			$verif_list = $list_detail.items
			if ($verif_list -notcontains "C:\temp\$($doss)") {
			$list_detail.AddText("C:\temp\$($doss)")
			$list_detail.RenderSize
			$list_detail.VerticalContentAlignment }
			}	
	}
}

Function New-ProgressBar {
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace =[runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $newRunspace
    $syncHash.AdditionalInfo = ''
    $newRunspace.ApartmentState = "STA" 
    $newRunspace.ThreadOptions = "ReuseThread"           
    $data = $newRunspace.Open() | Out-Null
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)           
    $PowerShellCommand = [PowerShell]::Create().AddScript({    
        [string]$xaml = @" 
			<Window Name="Window"
			        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
			        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
			        Title="Copie en cours..." Height="90" Width="403.348" ResizeMode="NoResize" Background="#FFD3D0D0
					" WindowStartupLocation="CenterScreen" Cursor="Wait" Topmost="True" WindowStyle="None">
			    <Grid>
			        <ProgressBar Height="17" Margin="23,15,23,0" VerticalAlignment="Top" RenderTransformOrigin="-0.042,-0.116" Name="ProgressBar"/>
			        <TextBlock HorizontalAlignment="Left" Margin="23,40,0,0" TextWrapping="Wrap" Text="Copie en cours, veuillez patienter !" VerticalAlignment="Top" Width="349"/>
                    <Rectangle HorizontalAlignment="Left" Height="90" Width="403.348" Stroke="Black" VerticalAlignment="Top"/>
			    </Grid>
			</Window>
"@ 

   
	$syncHash.Window=[Windows.Markup.XamlReader]::parse( $xaml ) 
    ([xml]$xaml).SelectNodes("//*[@Name]") | %{ $SyncHash."$($_.Name)" = $SyncHash.Window.FindName($_.Name)}
    

    $updateBlock = { $SyncHash.ProgressBar.Value = $SyncHash.PercentComplete }

        $syncHash.Window.Add_SourceInitialized( {            
            $timer = new-object System.Windows.Threading.DispatcherTimer            
            $timer.Interval = [TimeSpan]"0:0:0.01"            
            $timer.Add_Tick( $updateBlock )            
            $timer.Start()            
        } )

    $syncHash.Window.ShowDialog() | Out-Null 
    $syncHash.Error = $Error  }) 
	
    $PowerShellCommand.Runspace = $newRunspace 
    $data = $PowerShellCommand.BeginInvoke() 
   
    return $syncHash
}

function Close-ProgressBar {
    Param (
        [Parameter(Mandatory=$true)]
        [System.Object[]]$ProgressBar
    )

	$ProgressBar.Window.Dispatcher.Invoke([action]{ 
      $ProgressBar.Window.close() }, "Normal")
 }

function Write-ProgressBar {
    Param (
        [Parameter(Mandatory=$true)]
        $ProgressBar,
        [Parameter(Mandatory=$true)]
        [int]$PercentComplete
    ) 
   
   if($PercentComplete)
   { $ProgressBar.PercentComplete = $PercentComplete }
}

function Get-FolderSize {
    [CmdletBinding()]
    param(  
        [Parameter(ValueFromPipeline=$True)]
        [String]$Name,
 
        [Parameter(ValueFromPipeline=$true)]
        [System.IO.FileSystemInfo]$Folder,
 
        [switch]$TB,
        [switch]$GB,
        [switch]$MB
        )
        
    #begin { }
    process 
    {	if ($name) {$items = Get-ChildItem  $name -Recurse }
        elseif ($Folder){$items = Get-ChildItem  $folder -Recurse}
        $length = 0
        foreach ($i in $items){ $length += $i.length }
    }
    end {
        if ($TB) {$length/1TB}
        elseif ($GB) {$length/1GB}
        elseif ($MB){$length/1MB}
        else{$length}
    }
}

Function Detail {
$list_detail.items.Clear()
$det_path = $path_sccm+"\"+$list_result.SelectedItem
$get_detail = Get-ChildItem -Force -PATH $det_path -ErrorAction SilentlyContinue | Select-Object -Property FullName -ExpandProperty FullName

foreach ($line_det in $get_detail ) 	{$list_detail.AddText($($line_det.FullName))
											$list_detail.RenderSize
											$list_detail.VerticalContentAlignment }

}

Function Window_Verif_AD {
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
    $syncHash_AD = [hashtable]::Synchronized(@{})
    $newRunspace_AD =[runspacefactory]::CreateRunspace()
    $syncHash_AD.Runspace = $newRunspace_AD
    $syncHash_AD.AdditionalInfo = ''
    $newRunspace_AD.ApartmentState = "STA" 
    $newRunspace_AD.ThreadOptions = "ReuseThread"           
    $data_AD = $newRunspace_AD.Open() | Out-Null
    $newRunspace_AD.SessionStateProxy.SetVariable("syncHash_AD",$syncHash_AD)           
    $PowerShellCommand_AD = [PowerShell]::Create().AddScript({    
        [string]$xaml = @" 
			<Window Name="Window"
			        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
			        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
			        Title="Verif_AD" Height="60" Width="350" ResizeMode="NoResize" Background="#FFD3D0D0
					" WindowStartupLocation="CenterScreen" Cursor="Wait" Topmost="True" WindowStyle="None">
			    <Grid>
			        <TextBlock HorizontalAlignment="Left" Margin="30,20,0,0" TextWrapping="Wrap" Text="Vérification des habilitations, veuillez patienter..." VerticalAlignment="Top" Width="349"/>
                    <Rectangle HorizontalAlignment="Left" Height="60" Width="350" Stroke="Black" VerticalAlignment="Top"/>
			    </Grid>
			</Window>
"@ 

   
	$syncHash_AD.Window=[Windows.Markup.XamlReader]::parse( $xaml ) 
    ([xml]$xaml).SelectNodes("//*[@Name]") | %{ $SyncHash_AD."$($_.Name)" = $SyncHash_AD.Window.FindName($_.Name)}
    
    $syncHash_AD.Window.ShowDialog() | Out-Null 
    $syncHash_AD.Error = $Error  }) 
	
    $PowerShellCommand_AD.Runspace = $newRunspace_AD
    $data_AD = $PowerShellCommand_AD.BeginInvoke() 
   
    return $syncHash_AD
}

function Close_Window_Verif_AD {
    Param (
        [Parameter(Mandatory=$true)]
        [System.Object[]]$Window_Verif_AD
    )

	$Window_Verif_AD.Window.Dispatcher.Invoke([action]{ 
      $Window_Verif_AD.Window.close() }, "Normal")
 }

Function Verif_AD {
	if ($conscience -eq "OK") {$script:test_hab = "OK" }
	Else {
	$script:test_hab = "KO"
	Clear-Variable -Name verif_GAD -ErrorAction SilentlyContinue
	$new_wind_AD = Window_Verif_AD

	$Group_AD = "AG_"+$list_result.SelectedItem+$environnement

	$verif_GAD = Get-ADGroup "$($Group_AD)" -ErrorAction SilentlyContinue
	if ($verif_GAD -ne $null) {log -Add "Vérification Habilitation : Groupe AD $($Group_AD) existe"
								$membre_AD = Get-ADGroupMember $Group_AD -Recursive | select SamAccountName
								if ($membre_AD -match $Trigramme) { Log -Add "Vérification terminée : $($Trigramme) habilité"
																	$script:test_hab = "OK"}
								Else { Log -Add "Vérification terminée : $($Trigramme) non habilité"
										$script:test_hab = "KO"}
	 }
	 Else {log -Add "Vérification Habilitation : Groupe AD $($Group_AD) n'existe pas"
			Log -Add "Recherche de groupe AD contenant $($Group_AD)"
			$Currentdomain = "XXX" 
			$check_AD = Get-ADGroup -Filter * -SearchBase $Currentdomain  | where Name -Like "*$($Group_AD)*"

			if ($check_AD -eq $null) { Log -add "Recherche de groupe AD contenant $($list_result.SelectedItem)"
										$check_AD = Get-ADGroup -Filter * -SearchBase $Currentdomain  | where Name -Like "*$($list_result.SelectedItem)*"
										if ($check_AD -eq $null) { Log -Add "Vérification terminée : pas de groupe d'habilitation trouvé contenant $($list_result.SelectedItem)"
																	$script:test_hab = "NOK" }
										else { 
												foreach ($GAD in $check_AD.SamAccountName) {
												$membre_AD = Get-ADGroupMember $GAD -Recursive | select SamAccountName
												if ($membre_AD -match $Trigramme) { Log -Add "Vérification terminée : $($Trigramme) fait parti du groupe $($GAD)"
																	$script:test_hab = "OK"
																	break}
												Else {Log -Add "Vérification terminée : $($Trigramme) ne fait pas parti du groupe $($GAD)"
												$script:test_hab = "KO"}


					} }
										}
			Else {
				 foreach ($GAD in $check_AD.SamAccountName) {
								$membre_AD = Get-ADGroupMember $GAD -Recursive | select SamAccountName
								if ($membre_AD -match $Trigramme) { Log -Add "Vérification terminée : $($Trigramme) fait parti du groupe $($GAD)"
																	$script:test_hab = "OK"
																	break}
								Else {Log -Add "Vérification terminée : $($Trigramme) ne fait pas parti du groupe $($GAD)"
									$script:test_hab = "KO"}


					}
			 }
		}
	#$xml_al_form.Close()
	Close_Window_Verif_AD -Window_Verif_AD $new_wind_AD
	}
}

Function Copy_Window { 
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
    $syncHash_Copy = [hashtable]::Synchronized(@{})
    $newRunspace_Copy =[runspacefactory]::CreateRunspace()
    $syncHash_Copy.Runspace = $newRunspace_Copy
    $syncHash_Copy.AdditionalInfo = ''
    $newRunspace_Copy.ApartmentState = "STA" 
    $newRunspace_Copy.ThreadOptions = "ReuseThread"           
    $data_Copy = $newRunspace_Copy.Open() | Out-Null
    $newRunspace_Copy.SessionStateProxy.SetVariable("syncHash_Copy",$syncHash_Copy)           
    $PowerShellCommand_Copy = [PowerShell]::Create().AddScript({    
        [string]$xaml = @" 
			<Window Name="Window"
			        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
			        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
			        Title="Verif_AD" Height="60" Width="350" ResizeMode="NoResize" Background="#FFD3D0D0
					" WindowStartupLocation="CenterScreen" Cursor="Wait" Topmost="True" WindowStyle="None">
			    <Grid>
			        <TextBlock HorizontalAlignment="Left" Margin="30,20,0,0" TextWrapping="Wrap" Text="Copie du dossier en cours, veuillez patienter..." VerticalAlignment="Top" Width="349"/>
                    <Rectangle HorizontalAlignment="Left" Height="60" Width="350" Stroke="Black" VerticalAlignment="Top"/>
			    </Grid>
			</Window>
"@ 

   
	$syncHash_Copy.Window=[Windows.Markup.XamlReader]::parse( $xaml ) 
    ([xml]$xaml).SelectNodes("//*[@Name]") | %{ $SyncHash_Copy."$($_.Name)" = $SyncHash_Copy.Window.FindName($_.Name)}
    
    $syncHash_Copy.Window.ShowDialog() | Out-Null 
    $syncHash_Copy.Error = $Error  }) 
	
    $PowerShellCommand_Copy.Runspace = $newRunspace_Copy
    $data_Copy = $PowerShellCommand_Copy.BeginInvoke() 
   
    return $syncHash_Copy}

Function Close_Copy_Window { 
    Param (
        [Parameter(Mandatory=$true)]
        [System.Object[]]$Copy_Window
    )

	$Copy_Window.Window.Dispatcher.Invoke([action]{ 
      $Copy_Window.Window.close() }, "Normal")}

function Invoke-sql1 {
    param( [string]$sql,
        [System.Data.SQLClient.SQLConnection]$connection
    )

    $cmd = new-object System.Data.SQLClient.SQLCommand($sql,$connection)
    $ds = New-Object system.Data.DataSet
    $da = New-Object System.Data.SQLClient.SQLDataAdapter($cmd)
    $da.fill($ds) | Out-Null
    return $ds.tables[0].DefaultView
}

Function Surv_IP {
$Connexion = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }
if ($Connexion -eq $null ){ Log -add "Pas de réseau" }
Else {
	if ($Connexion.IPAddress.count -ne 1) {$IPAddress = "$($Connexion.IPAddress[0]) ; $($Connexion.IPAddress[1])"}
	Else {$IPAddress = "$($Connexion.IPAddress)" }
	       
	switch -wildcard ($IPAddress) {
	       "ad.IP*" {$place_to_be = 1}
	       "ad.IP*" {$place_to_be = 1}
	       "ad.IP*" {$place_to_be = 1}
	       "ad.IP*" {$place_to_be = 1}
	       "ad.IP*" {$place_to_be = 1}
	       "ad.IP*" {$place_to_be = 1}

	       "ad.IP*" {$place_to_be = 2}
	       "ad.IP*" {$place_to_be = 2}
	       "ad.IP*" {$place_to_be = 2}
		   "ad.IP*" {$place_to_be = 2}
	       "ad.IP*" {$place_to_be = 2}
	       "ad.IP*" {$place_to_be = 2}

	       "ad.IP*" {$place_to_be = 1}
           "ad.IP*" {$place_to_be = 1}
	       "IP*" { if($Connexion.IPAddress.count -eq 2){$place_to_be = 1}
	                 else {Log -add "non connecté GLB " } } 
	       Default {$place_to_be = 1
                    Log -Add "Emplacement inconnu : Adresse IP $($IPAddress)" }
	             }
	       }
}

#endregion


#region XML

#Load Assembly and Library
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$ThemeFile = 'C:\AppDSI\EXPGLB\SCCM\INSTAPACK\xaml\theme.xaml'

[xml]$alerte_msg = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MainWindow"  WindowStartupLocation="CenterScreen" Height="50" Width="400" Background="#FFD3D0D0" Foreground="White" WindowStyle="None" BorderBrush="Black" Cursor="Arrow" FontFamily="Calibri"
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
        <Button Name="btn_ok_alerte" Content="OK" Background="Gray" HorizontalAlignment="Left" Margin="296,12,0,0" VerticalAlignment="Top" Width="75" Grid.Column="1" Height="22" ClickMode="Press" IsHitTestVisible="True"/>
        <TextBlock Name="txtbck_alerte" HorizontalAlignment="Left" Foreground="Black" Margin="11,12,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="38" Width="260"/>
		<Rectangle Name="Rec" HorizontalAlignment="Left" Height="50" Width="400" Stroke="Black" VerticalAlignment="Top"/>
    </Grid>
</Window>
"@

#Param de la Form
[xml]$Form = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Name="SCCM" Title="SCCM" Height="435.424" Width="838.061" WindowStyle="None" WindowStartupLocation="CenterScreen" Background="#FFD3D0D0" Foreground="White" ResizeMode="NoResize" Cursor="Arrow">
        <Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile" /> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
        </Window.Resources>
     <Grid>
        <TextBox Name="Txtbx_Rech" HorizontalAlignment="Left" Height="28" Margin="24,90,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="171"/>
        <Button Name="Btn_Rech" Content="Rechercher" HorizontalAlignment="Left" Margin="213,90,0,0" VerticalAlignment="Top" Width="120" Height="28" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press"/>
        <ListView Name="ListView" HorizontalAlignment="Left" Height="254" Margin="24,134,0,0" VerticalAlignment="Top" Width="192">
            <ListView.View>
                <GridView>
                    <GridViewColumn Width="192"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button Name="Btn_Copy" Content="Récuperer le package" HorizontalAlignment="Left" Height="51" Margin="674,138,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press"/>
        <Button Name="Btn_Quit" Content="Quitter" HorizontalAlignment="Left" Margin="674,354,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press" ForceCursor="True" Height="34"/>
        <Label Name="Lbl_appli" Content="Quel application/package ?" HorizontalAlignment="Left" Height="33" Margin="24,61,0,0" VerticalAlignment="Top" Width="309"/>
        <Button Name="Btn_Inst" Content="Installation" HorizontalAlignment="Left" Height="51" Margin="674,202,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press"/>
        <Button Name="Btn_Inst_2" Content="Lancer l'installation" HorizontalAlignment="Left" Height="51" Margin="674,202,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press" Visibility="Hidden"/>
        <Button Name="Btn_UnInst" Content="Désinstallation" HorizontalAlignment="Left" Height="51" Margin="674,266,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press"/>
        <Button Name="Btn_UnInst_2" Content="Désinstallation" HorizontalAlignment="Left" Height="51" Margin="674,266,0,0" VerticalAlignment="Top" Width="146" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press" Visibility="Hidden"/>
        <Label Name="Lbl_info_copy" Content="" HorizontalAlignment="Left" Margin="24,396,0,0" VerticalAlignment="Top" Width="700"/>
        <Button Name="Btn_XQuit" Content="X" HorizontalAlignment="Left" Margin="814,5,0,0" VerticalAlignment="Top" Width="18" Height="21" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press" ForceCursor="True" FontSize="9"/>
        <Button Name="Btn_mini" Content="_" HorizontalAlignment="Left" Margin="792,5,0,0" VerticalAlignment="Top" Width="18" Height="21" Background="#FFB6B6B6" BorderBrush="Black" FontFamily="Calibri" Cursor="Arrow" ClickMode="Press" ForceCursor="True" FontSize="9"/>
        <Label Content="SCCM - Installation de packages" HorizontalAlignment="Left" Height="41" Margin="17,-1,0,0" VerticalAlignment="Top" Width="611" FontWeight="Bold" FontSize="16"/>
        <Label Content="Environnement : " HorizontalAlignment="Left" Margin="24,40,0,0" VerticalAlignment="Top" Width="100"/>
        <RadioButton Name="rdbt_EXP" Content="EXP" HorizontalAlignment="Left" Margin="129,45,0,0" VerticalAlignment="Top"/>
        <RadioButton Name="rdbt_EFC" Content="EFC" HorizontalAlignment="Left" Margin="181,45,0,0" VerticalAlignment="Top"/>
        <RadioButton Name="rdbt_PPR" Content="PPR" HorizontalAlignment="Left" Margin="233,45,0,0" VerticalAlignment="Top"/>
        <RadioButton Name="rdbt_RCT" Content="RCT" HorizontalAlignment="Left" Margin="285,45,0,0" VerticalAlignment="Top"/>
        <ListView Name="list_detail" HorizontalAlignment="Left" Height="254" Margin="221,134,0,0" VerticalAlignment="Top" Width="439">
            <ListView.View>
                <GridView>
                    <GridViewColumn Width="439"/>
                </GridView>
            </ListView.View>
        </ListView>
   		<Rectangle HorizontalAlignment="Left" Height="435.424" Stroke="Black" VerticalAlignment="Top" Width="838.061"/>
    </Grid>
</Window>
"@

#Creation de la form
$XMLReader = (New-Object System.Xml.XmlNodeReader $Form)
$XMLForm = [Windows.Markup.XamlReader]::Load($XMLReader)

#Chargement des controles
$txtbx_rech = $XMLForm.FindName('Txtbx_Rech')
$btn_rech = $XMLForm.FindName('Btn_Rech')
$btn_copy = $XMLForm.FindName('Btn_Copy')
$list_result = $XMLForm.FindName('ListView')
$list_detail = $XMLForm.FindName('list_detail')
$btn_inst = $XMLForm.FindName('Btn_Inst')
$btn_inst_2 = $XMLForm.FindName('Btn_Inst_2')
$btn_uninst = $XMLForm.FindName('Btn_UnInst')
$btn_uninst_2 = $XMLForm.FindName('Btn_UnInst_2')
$btn_quit =$XMLForm.FindName('Btn_Quit')
$lbl_info = $XMLForm.FindName('Lbl_info_copy')
$btn_xquit =$XMLForm.FindName('Btn_XQuit')
$btn_mini =$XMLForm.FindName('Btn_mini')

$rdbt_EXP =$XMLForm.FindName('rdbt_EXP')
$rdbt_EFC =$XMLForm.FindName('rdbt_EFC')
$rdbt_PPR =$XMLForm.FindName('rdbt_PPR')
$rdbt_RCT =$XMLForm.FindName('rdbt_RCT')

$rdbt_EXP.IsChecked = "True"

#endregion


Affichage


#region Action des boutons
$rdbt_EXP.add_Checked({ $script:path_sccm = '\\SCCMP029\SCCMContener\Packages\EXPGLB'
                        Affichage
                        $script:environnement = "_EXP"
                        $btn_inst.visibility = "Visible"
                        $btn_inst_2.visibility  = "Hidden" })

$rdbt_EFC.add_Checked({ $script:path_sccm = '\\SCCMP029\SCCMContener\Packages\EFCGLB'
                        Affichage
                        $script:environnement = "_EFC"
                        $btn_inst.visibility = "Visible"
                        $btn_inst_2.visibility  = "Hidden" })

$rdbt_PPR.add_Checked({ $script:path_sccm = '\\SCCMP029\SCCMContener\Packages\PPRGLB'
                        Affichage 
                        $script:environnement = "_PPR"
                        $btn_inst.visibility = "Visible"
                        $btn_inst_2.visibility  = "Hidden"})

$rdbt_RCT.add_Checked({ $script:path_sccm = '\\SCCMP029\SCCMContener\Packages\RCTGLB'
                        Affichage 
                        $script:environnement = "_RCT"
                        $btn_inst.visibility = "Visible"
                        $btn_inst_2.visibility  = "Hidden"})




$list_result.add_mousedoubleclick({ Detail })

$btn_rech.add_Click({
$btn_inst.visibility = "Visible"
$btn_inst_2.visibility  = "Hidden"
#Write-Host $list_result.SelectedItem

#Write-Host $path_sccm
$list_result.items.Clear()
$script:package = $txtbx_rech.get_text()
if ($package -eq $null -or $package -eq "") {$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                        $txtbck_alerte.text = "Merci de saisir quelque chose avant de lancer la recherche !"
                        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                        $xml_al_form.showdialog()
						}
else { Rech_Package }
})

$btn_copy.add_Click({
if ($btn_inst.visibility -eq "Hidden"){$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                        $txtbck_alerte.text = "Copie impossible depuis cet écran !"
                        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                        $xml_al_form.showdialog() }
Else {


if ($list_detail.SelectedItem -eq $null) { $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                        $txtbck_alerte.text = "Aucune application selectionnée !"
                        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                        $xml_al_form.showdialog() }
Else {
####################################
$script:package_selected = $list_detail.selecteditem

    Verif_AD

	if ($test_hab -eq "KO") { $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
	        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
			$btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
			$txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
			$txtbck_alerte.text = "L'utilisateur n'a pas l'habilitation : installation impossible !"
			$btn_ok_alerte.add_Click({ $xml_al_form.Close() })
			$xml_al_form.showdialog()
							
			Log -add "Installation $($list_result.selecteditem) bloquée => habilitation KO"				
							}
    Elseif ($test_hab -eq "NOK") {$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
	        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
			$btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
			$txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
            $Rec = $xml_al_form.FindName('Rec')
            $xml_al_form.Height = "70"
            $rec.Height = "70"
            $txtbck_alerte.Height = "70"
            $btn_ok_alerte.Margin = "296,20,0,0"
			$txtbck_alerte.text = "Groupe d'habilitation pour $($list_result.selecteditem) non trouvé - Récupération possible en recliquant sur le même bouton !"
			$btn_ok_alerte.add_Click({ $xml_al_form.Close() })
			$xml_al_form.showdialog()

			Log -add "Groupe d'habilitation contenant $($list_result.selecteditem) non trouvé - Récupération non bloquée - Vérifs à faire manuellement."
            $script:conscience = "OK" }
	Else {
        $script:conscience = "KO"
		Surv_Ip
		$folder_size = Get-folderSize -MB $package_selected
		if ($folder_size -gt 20 -and $place_to_be -eq 1) {
			$xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
	        $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
	        $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
	        $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
	        $txtbck_alerte.text = "La taille du dossier dépasse la limite autorisée. Désolé !"
	        $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
	        $xml_al_form.showdialog()
			
			Log -Add "Situation : Agence/Teletravail => Installation $($package) bloquée => taille > 20mo"
			}
		Else {
        $script:count_dir = "{0:D2}" -f $iii
		$script:dest = "C:\temp\"+"$($count_dir)_"+$package_selected.Split("\")[6]+"_"+$package_selected.Split("\")[7]+$environnement
        if ((Test-Path $dest) -eq $true){$script:iii++
                                         $script:count_dir = "{0:D2}" -f $iii
                                         $script:dest = "C:\temp\"+"$($count_dir)_"+$package_selected.Split("\")[6]+"_"+$package_selected.Split("\")[7]+$environnement}

        $Launch_Copy = Copy_Window
        Copy-Item $package_selected -Recurse $dest
        Close_Copy_Window -Copy_Window $Launch_Copy

		$script:count_copie++
		$lbl_info.Content = "Copie de $package_selected terminée ! Installation possible !"
		Log -Add "Copie de $($package_selected) dans $dest"

		$btn_inst.visibility = "Visible"
		$btn_inst_2.visibility  = "Hidden"
		}
	}
}
}
})

$btn_inst.add_Click({
	if ($count_copie -ge 2) {	$btn_inst.visibility = "Hidden"
								$btn_inst_2.visibility  = "Visible"
								$list_detail.items.Clear()
								Select_Package 
								 }
	else {
if ($package_selected -eq $null) { $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                                    $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                                    $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                                    $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                                    $txtbck_alerte.text = "Aucune package selectionné !"
                                    $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                                    $xml_al_form.showdialog() }
else {
Set-Location $dest
InstallPackage
Log -Add "Install de $($package_selected) "
Start-Sleep -Seconds 2
$surv_install = Get-Process -Name wscript -ErrorAction SilentlyContinue
if ($surv_install -ne $null) {Wait-Process -Name wscript
							  $Lbl_info.Content = "Installation de $package_selected terminée !"
							  $Install_date = Get-Date -Format ddMMyyyy_hh:mm:ss
							  Log -add "Install de $($package_selected) terminée." }
		}
}

})

$btn_inst_2.add_Click({
$pack_sel_cop = $list_detail.selecteditem
#Write-Host $pack_sel_cop
Set-Location $pack_sel_cop
InstallPackage
$surv_install = Get-Process -Name wscript -ErrorAction SilentlyContinue
if ($surv_install -ne $null) {Wait-Process -Name wscript
							  $Lbl_info.Content = "Installation de $package_selected terminée !"
							  $Install_date = Get-Date -Format ddMMyyyy_hh:mm:ss
							  Log -add "Install de $($package_selected) terminée." }
})

$btn_uninst.add_Click({ 
	if ($count_copie -ge 2) {	$btn_uninst.visibility = "Hidden"
								$btn_uninst_2.visibility  = "Visible"
								$list_detail.items.Clear()
								Select_Package 
								 }
	else {
if ($package_selected -eq $null) { $xml_alerte = (New-Object System.Xml.XmlNodeReader $alerte_msg)
                                    $xml_al_form = [Windows.Markup.XamlReader]::Load($xml_alerte)
                                    $btn_ok_alerte = $xml_al_form.FindName('btn_ok_alerte')
                                    $txtbck_alerte = $xml_al_form.FindName('txtbck_alerte')
                                    $txtbck_alerte.text = "Aucune package selectionné !"
                                    $btn_ok_alerte.add_Click({ $xml_al_form.Close() })
                                    $xml_al_form.showdialog() }
else {
Set-Location $dest
UnInstallPackage
Log -Add "Désinstall de $($package_selected) "
Start-Sleep -Seconds 2
$surv_install = Get-Process -Name wscript -ErrorAction SilentlyContinue
if ($surv_install -ne $null) {Wait-Process -Name wscript
							  $Lbl_info.Content = "Désinstallation de $package_selected terminée !"
							  $Install_date = Get-Date -Format ddMMyyyy_hh:mm:ss
							  Log -add "Désinstall de $($package_selected) terminée." }
		}
}

})

$btn_uninst_2.add_Click({
$pack_sel_cop = $list_detail.selecteditem
Set-Location $pack_sel_cop
UninstallPackage
$surv_install = Get-Process -Name wscript -ErrorAction SilentlyContinue
if ($surv_install -ne $null) {Wait-Process -Name wscript
							  $Lbl_info.Content = "Désinstallation de $package_selected terminée !"
							  $Install_date = Get-Date -Format ddMMyyyy_hh:mm:ss
							  Log -add "Désinstall de $($package_selected) terminée." }

 })

$btn_quit.add_click({
EffacerLesPreuves
Log -Add "Clear sur les variables stockées"
Log -add "Suppression des fichiers copiés et fermeture programme"
$XMLForm.Close()
})

$XMLForm.add_MouseLeftButtonDown({$this.DragMove() })

$btn_mini.add_click({ $XMLForm.WindowState = "Minimized" })

$btn_xquit.add_click({ $XMLForm.Close() })

#endregion


#Lancement de la form
$XMLForm.ShowDialog()
