#
#	INIT
#-------------------------------------------------------------------------
$Global:Current_Folder = $PSScriptRoot
Set-Location $Current_Folder
$Assembly_Folder = Join-Path -Path $Current_Folder -ChildPath .\assembly
$Resource_Folder = Join-Path -Path $Current_Folder -ChildPath .\Resources
$Image_Folder = Join-Path -Path $Current_Folder -ChildPath .\views\images
$Script_Name = $MyInvocation.MyCommand.Name

Import-Module Logs_TA
create_log -Path_Log $Current_Folder -Appli NETQOS_Data 
$Image_splash = "$($Image_Folder)\programm.jpg"


#	ASSEMBLIES
#-------------------------------------------------------------------------
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

#
# Chargement des dll du dossier assembly
#

If(Test-Path $Assembly_Folder) {
	Foreach ($Assembly in (Dir $Assembly_Folder -Filter *.dll)) {
		$null = [System.Reflection.Assembly]::LoadFrom($Assembly.fullName)
	}
}


#	TOAST
#-------------------------------------------------------------------------
function Toast_Message {
Param([Parameter(Position=0,Mandatory=$true,HelpMessage='Renseigner le titre de la fenetre')]
		[string]$Title ,
		[Parameter(Position=1,Mandatory=$true,HelpMessage='Renseigner le texte de la fenetre')]
		[string]$Text,
		[Parameter(Position=2,Mandatory=$true,HelpMessage='Renseigner le chemin de l image à insérer')]
		[string]$Image,
		[Parameter(Position=3,Mandatory=$true,HelpMessage='Action au clic - 0=Rien 1=HTML 2=CSV')]
		[int]$do_click
	)
	
	$WindowHeight = 120
	$WindowWidth = 500
	$Timeout = 10
	$ImageHeight = 48
	$ImageWidth = 48

	$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
	$Bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
	$TopStart = $workingArea.Bottom
	$TopFinish = $workingArea.Bottom - ($WindowHeight + 10)
	$CloseFinish = $Bounds.Bottom

	$MainStackWidth = $WindowWidth - 10
	$SecondStackWidth = $WindowWidth - $ImageWidth -10
	$TextBoxWidth = $SecondStackWidth - 10

[XML]$Xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='Rapport Notification' Width='$WindowWidth' Height='$WindowHeight'
    WindowStyle='None' AllowsTransparency='True' Background='Transparent' Topmost='True' Opacity='0.9'>
    <Window.Resources>
        <Storyboard x:Name='ClosingAnimation' x:Key='ClosingAnimation' >
            <DoubleAnimation Duration='0:0:.5' Storyboard.TargetProperty='Top' From='$TopFinish' To='$CloseFinish' AccelerationRatio='.1'/>
        </Storyboard>
    </Window.Resources>
    <Window.Triggers>
        <EventTrigger RoutedEvent='Window.Loaded'>
            <BeginStoryboard>
                <Storyboard >
                    <DoubleAnimation Duration='0:0:.5' Storyboard.TargetProperty='Top' From='$TopStart' To='$TopFinish' AccelerationRatio='.1'/>
                </Storyboard>
            </BeginStoryboard>
        </EventTrigger>
    </Window.Triggers>
    <Grid>
        <Border BorderThickness='0' Background='#333333'>

            <StackPanel Margin='10,10,30,10' Orientation='Horizontal' Width='$MainStackWidth'>
                <Image Source="$Image" Margin="10,10,10,10" Width="$ImageWidth"></Image>
                <StackPanel Width='$SecondStackWidth'>
                    <TextBox Margin='5' MaxWidth='$TextBoxWidth' Background='#333333' BorderThickness='0' IsReadOnly='True' Foreground='White' FontSize='16' Text='$Title' FontWeight='Bold' HorizontalContentAlignment='Left' Width='Auto' HorizontalAlignment='Stretch' IsHitTestVisible='False'/>
                    <TextBox Margin='5' MaxWidth='$TextBoxWidth' Background='#333333' BorderThickness='0' IsReadOnly='True' Foreground='LightGray' FontSize='11' Text='$Text' HorizontalContentAlignment='Left' TextWrapping='Wrap' IsHitTestVisible='False'/>
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$Global:UI = @{}
$Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
$UI.ClosingAnimation = $Window.FindName('ClosingAnimation')


$Window.Add_Closing({
    $UI.DispatcherTimer.Stop()
})

$UI.ClosingAnimation.Add_Completed({
    $Window.Close()
})

$Window.Add_MouseEnter({
    $This.Cursor = 'Hand'
})

$Window.Add_MouseUp({
	switch -wildcard ($do_click){
	 0 { Log -Add "Toast - Aucune action" }
     1 {Start-Process $html_name
	 	Log -Add "Toast - Ouverture HTML"}
	 2 {Start-Process 'C:\temp'
	 	Log -Add "Toast - Ouverture C:\Temp\"}
	
	}
    $UI.DispatcherTimer.Stop()
    $This.Close()
})

$null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

}


#	SPLASH SCREEN
#-------------------------------------------------------------------------
#Region SplasScreen
function Start-SplashScreen {
	$Pwshell.Runspace = $runspace
	$script:handle = $Pwshell.BeginInvoke() 
}

function Close-SplashScreen {
	$hash.window.Dispatcher.Invoke("Normal",[action]{ $hash.window.close() })
	$null = $Pwshell.EndInvoke($handle)
	#$null = $runspace.Close()
}

$hash = [Hashtable]::Synchronized(@{})
$hash.Image_splash = $Image_splash
$runspace = [Runspacefactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("hash",$hash) 
$Pwshell = [PowerShell]::Create()	
	
$null = $Pwshell.AddScript({
$xml = [xml]@"
<Window
	xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	x:Name="WindowSplash" Title="SplashScreen" WindowStyle="None" WindowStartupLocation="CenterScreen"
	AllowsTransparency="True"  ShowInTaskbar ="False" BorderBrush="Transparent" 
	Width="378.053" Height="342.937" ResizeMode = "NoResize"  Background="Transparent" >

    <Border BorderBrush="AliceBlue" BorderThickness="1" CornerRadius="8" Background="#FF444444">
        <Grid Margin="0,0,0,0">
            <Image Source="$($hash.Image_splash)" Opacity="0.2" Margin="-92,-158,-132,-148"/>
            <StackPanel VerticalAlignment="Center" Height="340.906" Margin="0,0,0,0">
                <Label Name="LoadingLabel" Content="NETQOS"  Foreground="WhiteSmoke" VerticalAlignment="Center" FontSize="35" Margin = "97.526,45.468,97.837,215.469" Width="180.69" Height="80" HorizontalContentAlignment="Center" VerticalContentAlignment="Center"/>
                <Controls:ProgressRing IsActive="{Binding IsActive}" Foreground="WhiteSmoke" Margin="0,-250,0,0"   Width="58.271" Height="45"/>
                <Label Content="Version 2.0.0" Height="23.532" FontSize="10" Margin="0,-450,-70,0" Foreground="WhiteSmoke" HorizontalContentAlignment="Center"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@
$reader = New-Object System.Xml.XmlNodeReader $xml
$hash.window = [Windows.Markup.XamlReader]::Load($reader) 
$hash.window.ShowDialog() 
})
#endregion

#	VERIF CONNECTION
#-------------------------------------------------------------------------
Function Test_Connection {
if (!(Test-Connection nas22 -Count 1 )) { Log -Add "Déconnecté !"
								$Title = "NETQOS - Information"
								$Text = "Le poste n est pas connecté au réseau. L application va se fermer. Merci de vérifier votre connexion avant de relancer l application"
								Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\sad.png" -do_click 0
								exit	}
Else { Log -Add "Connecté !"}
}

#
# Ouverture Splash Screen
#

Start-SplashScreen
Log -Add "Lancement Start-SplashScreen"
Test_Connection



#	XAML
#-------------------------------------------------------------------------
Function LoadXaml($Global:filename) {
	$xamlLoader = (New-Object System.Xml.XmlDocument)
	$xamlLoader.Load($filename)
	$xamlLoader
}

#
# Chargement xaml des fenetres
#

# Fenetre principale
$xaml = LoadXaml("$Current_Folder\MainWindow.xaml") 
$xamlReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window_Main = [Windows.Markup.XamlReader]::Load($xamlReader)


#
# Tag des vues
#
$HamburgerMenuControl = $Window_Main.FindName("HamburgerMenuControl")

$HomeView      = $Window_Main.FindName("HomeView") 
$SettingsView  = $Window_Main.FindName("SettingsView")
$PrivateView   = $Window_Main.FindName("PrivateView") 
$AboutView     = $Window_Main.FindName("AboutView") 
$ChartsView	   = $Window_Main.FindName("ChartsView")
$QuitView	   = $Window_Main.FindName("QuitView")


$viewFolder = $Current_Folder+"\views"

$XamlChildWindow = LoadXaml($viewFolder+"\Home.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$HomeXaml        = [Windows.Markup.XamlReader]::Load($Childreader) 

$XamlChildWindow = LoadXaml($viewFolder+"\Private.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$PrivateXaml     = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Settings.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$SettingsXaml    = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\About.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$AboutXaml       = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Charts.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$ChartsXaml       = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Quit.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$QuitXaml       = [Windows.Markup.XamlReader]::Load($Childreader)

    
$HomeView.Children.Add($HomeXaml)          | Out-Null
$PrivateView.Children.Add($PrivateXaml)    | Out-Null  
$SettingsView.Children.Add($SettingsXaml)  | Out-Null    
$AboutView.Children.Add($AboutXaml)        | Out-Null
$ChartsView.Children.Add($ChartsXaml)	   | Out-Null
$QuitView.Children.Add($QuitXaml)	   | Out-Null

#******************************************************
#******************************************************

$HamburgerMenuControl.SelectedItem = $HamburgerMenuControl.ItemsSource[0]

#******************* Items Section  *******************

$HamburgerMenuControl.add_ItemClick({
   $HamburgerMenuControl.Content = $HamburgerMenuControl.SelectedItem
   $HamburgerMenuControl.IsPaneOpen = $false

})

#******************* Options Section  *******************

$HamburgerMenuControl.add_OptionsItemClick({
    $HamburgerMenuControl.Content = $HamburgerMenuControl.SelectedOptionsItem
    $HamburgerMenuControl.IsPaneOpen = $false
})

#*******************Home**************************
$lbl_netbios = $HomeXaml.FindName("lbl_netbios")
$lbl_model = $HomeXaml.FindName("lbl_model")
$lbl_ip_priv = $HomeXaml.FindName("lbl_ip_priv")
$lbl_ip_pub = $HomeXaml.FindName("lbl_ip_pub")
$cbx_user = $HomeXaml.FindName("cbx_user")
$btn_go = $HomeXaml.FindName("btn_go")
$dtpr_date = $HomeXaml.FindName("dtpr_date")
$dtpr_date_fin = $HomeXaml.FindName("dtpr_date_fin")
$dtg_speed = $HomeXaml.FindName("dtg_speed")
$dtg_ping = $HomeXaml.FindName("dtg_ping")
$cbx_hour_one = $HomeXaml.FindName("cbx_hour_one")
$cbx_hour_two = $HomeXaml.FindName("cbx_hour_two")
$ico_dev = $HomeXaml.FindName("ico_dev")
$dtpr_date.selecteddate = $dtpr_date.displaydate
$dtpr_date_fin.selecteddate = $dtpr_date_fin.displaydate
$btn_export_home = $HomeXaml.FindName("btn_export_home")
$mn_exp_csv_home = $HomeXaml.FindName("mn_exp_csv_home")
$mn_exp_html_home = $HomeXaml.FindName("mn_exp_html_home")
$ico_dtg_ping = $HomeXaml.FindName("ico_dtg_ping")

$hours = @("00:00:00";"01:00:00";"02:00:00";"03:00:00";"04:00:00";"05:00:00";"06:00:00";"07:00:00";"08:00:00";"09:00:00";"10:00:00";"11:00:00";"12:00:00";"13:00:00";"14:00:00";"15:00:00";"16:00:00";"17:00:00";"18:00:00";"19:00:00";"20:00:00";"21:00:00";"22:00:00";"23:00:00")
$cbx_hour_one.itemssource = $cbx_hour_two.itemssource = $hours
$cbx_hour_one.text = $hours[0]
$cbx_hour_two.text = $hours[23]
#*********************Private************************
$cbx_user_global = $PrivateXaml.FindName("cbx_user_global")
$dtpr_date_global = $PrivateXaml.FindName("dtpr_date_global")
$dtpr_date_fin_global = $PrivateXaml.FindName("dtpr_date_fin_global")
$dtg_global = $PrivateXaml.FindName("dtg_global")
$btn_go_global = $PrivateXaml.FindName("btn_go_global")
$ico_dev_priv = $PrivateXaml.FindName("ico_dev_priv")
$dtpr_date_global.selecteddate = $dtpr_date_global.displaydate
$dtpr_date_fin_global.selecteddate = $dtpr_date_fin_global.displaydate
$btn_export_glob = $PrivateXaml.FindName("btn_export_glob")

$mn_exp_csv_glob  = $PrivateXaml.FindName("mn_exp_csv_glob")
$mn_exp_html_glob  = $PrivateXaml.FindName("mn_exp_html_glob")
#*********************Settings************************
$rdtb_det = $SettingsXaml.FindName("rdtb_det")
$rdbt_avg = $SettingsXaml.FindName("rdbt_avg")
$rdbt_date = $SettingsXaml.FindName("rdbt_date")
$rdbt_prd = $SettingsXaml.FindName("rdbt_prd")
$rdbt_prod = $SettingsXaml.FindName("rdbt_prod")
$rdbt_dev = $SettingsXaml.FindName("rdbt_dev")

$DragDrop = $SettingsXaml.FindName("DragDrop")
$rdbt_AD = $SettingsXaml.FindName("rdbt_AD")
$rdbt_fic = $SettingsXaml.FindName("rdbt_fic")
$rdbt_one = $SettingsXaml.FindName("rdbt_one")
$ltv_list = $SettingsXaml.FindName("ltv_list")
$btn_go_sett = $SettingsXaml.FindName("btn_go_sett")
$lbl_count_poste = $SettingsXaml.FindName("lbl_count_poste")
$txtb_unit = $SettingsXaml.FindName("txtb_unit")
$btn_unit = $SettingsXaml.FindName("btn_unit")

$rdbt_ARRET = $SettingsXaml.FindName("rdbt_ARRET")
$rdbt_RELANCE = $SettingsXaml.FindName("rdbt_RELANCE")

#$ico_dev_test = $SettingsXaml.FindName("ico_dev_test")
#*******************Charts******************************
$donut = $ChartsXaml.FindName("Donut")
$pieChart = $ChartsXaml.FindName("pieChart")
	$C_Ping_ok = $ChartsXaml.FindName("C_Ping_ok")
	$C_Ping_ko = $ChartsXaml.FindName("C_Ping_ko")
$lineChart = $ChartsXaml.FindName("lineChart")
$axisX = $ChartsXaml.FindName("axisX")

$btn_refresh = $ChartsXaml.FindName("btn_refresh")
$btn_refresh_glob = $ChartsXaml.FindName("btn_refresh_glob")

$dtp_chart_glob = $ChartsXaml.FindName("dtp_chart_glob")
$dtp_chart_pers = $ChartsXaml.FindName("dtp_chart_pers")
$dtp_chart_glob.selecteddate = $dtp_chart_glob.displaydate
$dtp_chart_pers.selecteddate = $dtp_chart_pers.displaydate

$cb_chart_glob = $ChartsXaml.FindName("cb_chart_glob")
$cb_chart_pers = $ChartsXaml.FindName("cb_chart_pers")

$donut.Visibility = "Hidden"
$lineChart.Visibility = "Hidden"
$pieChart.Visibility = "Hidden"

$lbl_intro = $ChartsXaml.FindName("lbl_intro")
$lbl_count = $ChartsXaml.FindName("lbl_count")
$lbl_details = $ChartsXaml.FindName("lbl_details")
$txt_version = $ChartsXaml.FindName("txt_version")
$lbl_download  = $ChartsXaml.FindName("lbl_download")
$lbl_upload  = $ChartsXaml.FindName("lbl_upload")

########################################################
#$test_block = $AboutXaml.FindName("test_block")
$ico_notice = $AboutXaml.FindName("ico_notice")
$txt_name = $AboutXaml.FindName("txt_name")
$txt_email = $AboutXaml.FindName("txt_email")
$rtbx_comm = $AboutXaml.FindName("rtbx_comm")
$btn_send_comm = $AboutXaml.FindName("btn_send_comm")
$rtxt_run = $AboutXaml.FindName("rtxt_run")
########################################################
$btn_quit = $QuitXaml.FindName("btn_quit")


#	FONCTIONS
#-------------------------------------------------------------------------

[string]$script:Base_SQL = "ServerSQL"
function Invoke-sql1 {
	<#
	.SYNOPSIS
		 Connexion SQL
	.EXAMPLE
		 Invoke-sql1 $data_total $con
	#>
		param( [string]$sql,
			[System.Data.SQLClient.SQLConnection]$connection
		)
	
		$cmd = new-object System.Data.SQLClient.SQLCommand($sql,$connection)
		$ds = New-Object system.Data.DataSet
		$da = New-Object System.Data.SQLClient.SQLDataAdapter($cmd)
		$da.fill($ds) | Out-Null
		return $ds.tables[0].DefaultView
	}
	
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$users_for_search = " SELECT PERS.TRIGRAMME, PERS.FIRSTNAME, PERS.LASTNAME FROM TIBIA.dbo.PERSONNES AS PERS
     WHERE PERS.ID_ENTREPRISE = 'GLB' AND PERS.VISIBILITE_OUTLOOK = 'O'
     ORDER BY PERS.LASTNAME"
$Search_users = Invoke-sql1 $users_for_search $con
$con.close()

Log -Add "Chargement des comptes Utilisateurs"
foreach ($tt in $Search_users) {
	$cbx_user.items.add("$($tt.LASTNAME) $($tt.FIRSTNAME) - $($tt.TRIGRAMME)")
	$cb_chart_pers.items.add("$($tt.LASTNAME) $($tt.FIRSTNAME) - $($tt.TRIGRAMME)")

}
$cbx_user_global.items.add("Tout le monde")
$cbx_user_global.items.add("Teletravailleurs")

$cb_chart_glob.items.add("Tout le monde")
$cb_chart_glob.items.add("Teletravailleurs")
Log -Add "Fin du chargement"


function Get-StyleSheet {
    [CmdletBinding()]
    Param()
@"
<style>
body {
    font-family:Segoe,Tahoma,Arial,Helvetica;
    font-size:10pt;
    color:#333;
    background-color:#eee;
    margin:10px;
}
th {
    font-weight:bold;
    color:white;
    background-color:#333;
}
</style>
"@
}

$ico_notice.add_Click({start "$($Resource_Folder)\Notice_NETQOS.pdf"
						Log -Add "Acces notice"
						 })

#	VARIABLES
#-------------------------------------------------------------------------

$rdbt_dev.add_click({ $ico_dev.Visibility = "Visible"
						$ico_dev_priv.Visibility = "Visible"
						[string]$script:Base_SQL = "ServerSQL"
						Log -Add "Base SQL : ServerSQL"})
$rdbt_prod.add_click({ $ico_dev.Visibility = "Hidden"
						$ico_dev_priv.Visibility = "Hidden"
						[string]$script:Base_SQL = "ServerSQL"
						Log -Add "Base SQL : ServerSQL"})

$rdbt_avg.add_click({ })
$rdtb_det.add_click({ })

$rdbt_date.add_click({ $dtpr_date_fin_global.Visibility = "Hidden"
					$dtpr_date_fin.Visibility = "Hidden"
					Log -Add "Activation recherche sur une date"})
$rdbt_prd.add_click({ $dtpr_date_fin_global.Visibility = "Visible"
					$dtpr_date_fin.Visibility = "Visible"
					Log -Add "Activation recherche sur une periode"})
					

#	REQUETES SQL
#--------------------------------------------------------------------------

Function SQL_EveryBody {
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$data_total = "select PERS.TRIGRAMME,
	PERS.FIRSTNAME,
	PERS.LASTNAME,
	PERS.SERVICE,
	LOG.ID_PDT,
	LOG.DATE_TEST,
	CONVERT(varchar(10), LOG.DATE_TEST, 103) as DATE,
	CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) as AVG_DESC,
	CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) as AVG_ASC,
	CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) as AVG_TPS_AVG_MS,
	CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) as AVG_TPS_MAX_MS,
		CASE WHEN CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) > '2'
			THEN '1'
			ELSE '2'
		END AS ID_DESC,
		CASE WHEN CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) > '1'
			THEN '1'
			ELSE '2'
		END AS ID_ASC,
		CASE WHEN CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) < '100'
			THEN '1'
			ELSE '2'
		END AS ID_PING,
		CASE WHEN CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) < '100'
			THEN '1'
			ELSE '2'
		END AS ID_PING_MAX
	from TIBIA.dbo.PERSONNES AS PERS
	inner join LOG_TELETRAVAIL.dbo.PDT_TEST AS LOG on LOG.USER_TEST = PERS.TRIGRAMME
	inner join LOG_TELETRAVAIL.dbo.TEST_DEBIT_PDT AS DEB on DEB.ID_TEST = LOG.ID_TEST
	inner join LOG_TELETRAVAIL.dbo.TEST_PDT_PING AS PING on PING.ID_TEST = LOG.ID_TEST
	where LOG.DATE_TEST >= '$date 00:00:00' and LOG.DATE_TEST <= '$date_fin 23:00:00'
	GROUP BY PERS.TRIGRAMME,PERS.FIRSTNAME,PERS.LASTNAME,PERS.SERVICE,PERS.SID_PERS,LOG.ID_PDT,LOG.DATE_TEST,CONVERT(varchar(10), LOG.DATE_TEST, 103)
	ORDER BY PERS.LASTNAME,LOG.DATE_TEST DESC"	

$data_totaux = Invoke-sql1 $data_total $con
$script:data_totaux = $data_totaux
}

Function SQL_TeleTravail {
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$teletravail = "
	select PERS.TRIGRAMME,
	PERS.FIRSTNAME,
	PERS.LASTNAME,
	PERS.SERVICE,
	LOG.ID_PDT,
	LOG.DATE_TEST,
	CONVERT(varchar(10), LOG.DATE_TEST, 103) as DATE,
	CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) as AVG_DESC,
	CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) as AVG_ASC,
	CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) as AVG_TPS_AVG_MS,
	CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) as AVG_TPS_MAX_MS,
		CASE WHEN CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) > '2'
			THEN '1'
			ELSE '2'
		END AS ID_DESC,
		CASE WHEN CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) > '1'
			THEN '1'
			ELSE '2'
		END AS ID_ASC,
		CASE WHEN CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) < '100'
			THEN '1'
			ELSE '2'
		END AS ID_PING,
		CASE WHEN CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) < '100'
			THEN '1'
			ELSE '2'
		END AS ID_PING_MAX
	from TIBIA.dbo.PERSONNES AS PERS
	inner join TIBIA.dbo.TELETRAVAILLEURS AS TELE on TELE.SID_PERS = PERS.SID_PERS
	inner join LOG_TELETRAVAIL.dbo.PDT_TEST AS LOG on LOG.USER_TEST = PERS.TRIGRAMME
	inner join LOG_TELETRAVAIL.dbo.TEST_DEBIT_PDT AS DEB on DEB.ID_TEST = LOG.ID_TEST
	inner join LOG_TELETRAVAIL.dbo.TEST_PDT_PING AS PING on PING.ID_TEST = LOG.ID_TEST
	where LOG.DATE_TEST >= '$date 00:00:00' and LOG.DATE_TEST <= '$date_fin 23:00:00'
	GROUP BY PERS.TRIGRAMME,PERS.FIRSTNAME,PERS.LASTNAME,PERS.SERVICE,PERS.SID_PERS,LOG.DATE_TEST,LOG.ID_PDT,CONVERT(varchar(10), LOG.DATE_TEST, 103)
	ORDER BY PERS.LASTNAME, LOG.DATE_TEST DESC"
	
$result_tt = Invoke-sql1 $teletravail $con
$script:result_tt = $result_tt
}

Function SQL_TT_AllData {
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$sql_user = "select DESC_PDT,ID_PDT,MODELE_PDT,IP_PRIVATE,IP_PULSE,IP_PUBLIC
	from dbo.PDT_TEST
	where dbo.PDT_TEST.DATE_TEST >= '$date $hor_deb' and dbo.PDT_TEST.DATE_TEST <= '$date_fin $hor_fin'  and dbo.PDT_TEST.USER_TEST = '$teletravaileur' "
#################################################################################
$sql_ping = "select DATE_TEST_PING,IP_PULSE,DESC_PDT,PING_TYPE,PING_PCT_OK,
	TPS_MAX_MS,
        CASE WHEN TPS_MAX_MS < '100'
            THEN '1'
            ELSE '2'
        END AS ID_PING,
	TPS_MIN_MS,
	TPS_AVG_MS
	from dbo.TEST_PDT_PING
	inner join dbo.PDT_TEST on dbo.TEST_PDT_PING.ID_TEST = dbo.PDT_TEST.ID_TEST
	where dbo.TEST_PDT_PING.DATE_TEST_PING >= '$date $hor_deb' and dbo.TEST_PDT_PING.DATE_TEST_PING <= '$date_fin $hor_fin'  and dbo.PDT_TEST.USER_TEST = '$teletravaileur'
	order by DATE_TEST_PING"
#################################################################################
$sql_speed = "select DATE_TEST_DEBIT,IP_PULSE,DESC_PDT,
	ASC_MBPS,
        CASE WHEN ASC_MBPS > '1'
	            THEN '1'
	            ELSE '2'
        END AS ID_ASC,
	DESC_MBPS,
        CASE WHEN DESC_MBPS > '2'
                THEN '1'
                ELSE '2'
        END AS ID_DESC
	from dbo.TEST_DEBIT_PDT
	inner join dbo.PDT_TEST on dbo.TEST_DEBIT_PDT.ID_TEST = dbo.PDT_TEST.ID_TEST
	where dbo.TEST_DEBIT_PDT.DATE_TEST_DEBIT >= '$date $hor_deb' and dbo.TEST_DEBIT_PDT.DATE_TEST_DEBIT <= '$date_fin $hor_fin'  and dbo.PDT_TEST.USER_TEST = '$teletravaileur'
	order by DATE_TEST_DEBIT"
#################################################################################
$script:data_user = Invoke-sql1 $sql_user $con
$script:data_speed = Invoke-sql1 $sql_speed $con 
$script:data_ping = Invoke-sql1  $sql_ping $con
}

Function SQL_TT_AVG {
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$sql_user = "select DESC_PDT,ID_PDT,MODELE_PDT,IP_PRIVATE,IP_PULSE,IP_PUBLIC
	from dbo.PDT_TEST
	where dbo.PDT_TEST.DATE_TEST >= '$date $hor_deb' and dbo.PDT_TEST.DATE_TEST <= '$date_fin $hor_fin'  and dbo.PDT_TEST.USER_TEST = '$teletravaileur' "
#############################################################
$sql_avg_user = "select PERS.TRIGRAMME,
    PERS.FIRSTNAME,
    PERS.LASTNAME,
    PERS.SERVICE,
    LOG.ID_PDT,
	LOG.DATE_TEST,
    CONVERT(varchar(10), LOG.DATE_TEST, 103) as DATE,
    CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) as AVG_DESC,
    CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) as AVG_ASC,
    CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) as AVG_TPS_AVG_MS,
    CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) as AVG_TPS_MAX_MS,
        CASE WHEN CAST(ROUND(AVG(DEB.DESC_MBPS), 2)As decimal(10,2)) > '2'
            THEN '1'
            ELSE '2'
        END AS ID_DESC,
        CASE WHEN CAST(ROUND(AVG(DEB.ASC_MBPS),2)As decimal(10,2)) > '1'
            THEN '1'
            ELSE '2'
        END AS ID_ASC,
        CASE WHEN CAST(ROUND(AVG(PING.TPS_AVG_MS),2)As decimal(10,2)) < '100'
            THEN '1'
            ELSE '2'
        END AS ID_PING,
        CASE WHEN CAST(ROUND(AVG(PING.TPS_MAX_MS),2)As decimal(10,2)) < '100'
            THEN '1'
            ELSE '2'
        END AS ID_PING_MAX
    from TIBIA.dbo.PERSONNES AS PERS
    inner join LOG_TELETRAVAIL.dbo.PDT_TEST AS LOG on LOG.USER_TEST = PERS.TRIGRAMME
    inner join LOG_TELETRAVAIL.dbo.TEST_DEBIT_PDT AS DEB on DEB.ID_TEST = LOG.ID_TEST
    inner join LOG_TELETRAVAIL.dbo.TEST_PDT_PING AS PING on PING.ID_TEST = LOG.ID_TEST
    where LOG.DATE_TEST >= '$date $hor_deb' and LOG.DATE_TEST <= '$date_fin $hor_fin' and LOG.USER_TEST = '$teletravaileur'
    GROUP BY PERS.TRIGRAMME,PERS.FIRSTNAME,PERS.LASTNAME,PERS.SERVICE,PERS.SID_PERS,LOG.ID_PDT,LOG.DATE_TEST,CONVERT(varchar(10), LOG.DATE_TEST, 103)
    ORDER BY LOG.DATE_TEST DESC"

$script:data_user = Invoke-sql1 $sql_user $con
$script:data_user_avg = Invoke-sql1 $sql_avg_user $con

}

Function SQL_TT_Postes {
$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$script:Base_SQL;Database=LOG_TELETRAVAIL;Trusted_Connection=True;"
$con.open()
$sql_poste = "select ID_PDT from dbo.PDT_TEST"
$data_postes = Invoke-sql1 $sql_poste $con 
$script:data_poste = $data_postes | select ID_PDT -Unique -ExpandProperty ID_PDT | sort ID_PDT}

###########################################################################
############## CHARTS ############

$btn_refresh.add_Click({ 
if ($PSVersionTable.PSVersion.Major -lt 5) { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","La version de Powershell installée sur le poste n est pas compatible avec ces fonctionnalités. Merci d utiliser le formulaire sur l écran A propos pour indiquer que vous rencontrez ce blocage. Vous serez recontacté dans les plus brefs délais.")
			Log -Add "Version KO"}
Else {
	$date = $dtp_chart_glob.selecteddate.tostring().split(" ")[0]
	$date = Get-Date $date -Format yyyyMMdd
	$date_fin = $date
    $script:exp_date = $date
	$script:exp_data = $cb_chart_glob.get_text().trim()
	switch -wildcard ($exp_data) {
		"Tout le monde" { SQL_EveryBody
						if ($data_totaux  -eq $null){[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée à cette date.") }
						Else {
								$total = $data_totaux.Count
								$total_OK = ($data_totaux | where {($_.ID_DESC -EQ 1) -and ($_.ID_ASC -eq 1) }).count
								$total_KO = ($data_totaux | where {($_.ID_DESC -EQ 2) -and ($_.ID_ASC -eq 2) }).count
								$total_DESC_KO = ($data_totaux | where {($_.ID_DESC -EQ 2) -and ($_.ID_ASC -eq 1) }).count
								$total_ASC_KO =  ($data_totaux | where {($_.ID_DESC -EQ 1) -and ($_.ID_ASC -eq 2) }).count
								#$total_PING_KO = ($data_totaux | where {$_.ID_PING_MAX -eq 2}).count
								}#ELSE
							#####################################################################
							} #Tout le monde
		"Teletravailleurs" {SQL_TeleTravail
							if ($result_tt -eq $null){[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée à cette date.") }
								Else {
								$total = $result_tt.Count 
								$total_OK = ($result_tt | where {($_.ID_DESC -EQ 1) -and ($_.ID_ASC -eq 1) }).count
								$total_KO = ($result_tt | where {($_.ID_DESC -EQ 2) -and ($_.ID_ASC -eq 2) }).count
								$total_DESC_KO = ($result_tt | where {($_.ID_DESC -EQ 2) -and ($_.ID_ASC -eq 1) }).count
								$total_ASC_KO =  ($result_tt | where {($_.ID_DESC -EQ 1) -and ($_.ID_ASC -eq 2) }).count
								#$total_PING_KO = ($result_tt | where {$_.ID_PING_MAX -eq 2}).count
								} #ELSE
							} #Teletravailleur
		Default { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Merci de sélectionner une ligne dans le menu déroulant.")}					
		} #SWITCH

	########################
	#Region Donut
	$donutcollections = [LiveCharts.SeriesCollection]::new()

	$chartvalue1 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
	$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
	$chartvalue1.Add([LiveCharts.Defaults.ObservableValue]::new($total_OK))
	$pieSeries.Values = $chartvalue1
	$pieSeries.Title = "OK"
	$pieSeries.DataLabels = $true

	$donutcollections.Add($pieSeries)


	$chartvalue2 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
	$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
	$chartvalue2.Add([LiveCharts.Defaults.ObservableValue]::new($total_KO))
	$pieSeries.Values = $chartvalue2
	$pieSeries.Title = "KO"
	$pieSeries.DataLabels = $true

	$donutcollections.Add($pieSeries)


	$chartvalue3 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
	$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
	$chartvalue3.Add([LiveCharts.Defaults.ObservableValue]::new($total_DESC_KO))
	$pieSeries.Values = $chartvalue3
	$pieSeries.Title = "DESC_KO"
	$pieSeries.DataLabels = $true


	$donutcollections.Add($pieSeries)

	$chartvalue4 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
	$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
	$chartvalue4.Add([LiveCharts.Defaults.ObservableValue]::new($total_ASC_KO))
	$pieSeries.Values = $chartvalue4
	$pieSeries.Title = "ASC_KO"
	$pieSeries.DataLabels = $true


	$donutcollections.Add($pieSeries)
	#endRegion

	$lbl_count.Content = "Sur un total de : $($total) personnes"
	$donut.Series = $donutcollections
	$donut.Visibility = "Visible"
	Log -Add "Generation Donut - donnees globales sur $($exp_date)"
}#ELSE VERSION
})


$btn_refresh_glob.add_Click({
if ($PSVersionTable.PSVersion.Major -lt 5) { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","La version de Powershell installée sur le poste n est pas compatible avec ces fonctionnalités. Merci d utiliser le formulaire sur l écran A propos pour indiquer que vous rencontrez ce blocage. Vous serez recontacté dans les plus brefs délais.")
		Log -Add "Version KO"}
Else {		
	$date = $dtp_chart_pers.selecteddate.tostring().split(" ")[0]
	$date = Get-Date $date -Format yyyyMMdd
	$date_fin = $date
    $script:exp_date = $date
	$script:exp_data = $cb_chart_pers.get_text().Split("-")[1].trim()
	$teletravaileur = $cb_chart_pers.get_text().Split("-")[1].trim()
	$hor_deb = "00:00:00"
	$hor_fin = "23:00:00"
	if ($teletravaileur -eq $null -or $teletravaileur -eq "") {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Merci de sélectionner une ligne dans le menu déroulant.") }
	Else{
	SQL_TT_AllData
		if ($data_speed -eq $null -and $data_ping -eq $null) {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée disponible à cette date.") }
		Else {
			$malist = New-Object System.Collections.Generic.List[System.String]
			Foreach ($hour in $data_speed.DATE_TEST_DEBIT){
			    $hour_add = (Get-Date $hour -Format HH:mm)
			    $malist.Add($hour_add)
			    }

			$deb_asc = $data_speed | select ASC_MBPS
			$deb_desc = $data_speed | select DESC_MBPS

			$ping_chart_KO = $data_ping | Select PING_PCT_OK,ID_PING | where {$_.ID_PING -eq 2 }
			$ping_chart_OK = $data_ping | Select PING_PCT_OK,ID_PING | where {$_.ID_PING -eq 1 }
				if ($ping_chart_KO -eq $null) {$ping_t_ko = 0 }
				Else {$ping_t_ko = [Linq.Enumerable]::Sum([int[]]($ping_chart_KO.PING_PCT_OK )) }
				if ( $ping_chart_OK -eq $null) {$ping_t_ok = 0}
				Else {$ping_t_ok = [Linq.Enumerable]::Sum([int[]]($ping_chart_OK.PING_PCT_OK )) }
				$total_ping = ($ping_chart_KO.Count + $ping_chart_OK.count )*100
				$ping_total_ko = $total_ping - ($ping_t_ko + $ping_t_ok )


			$avg_desc = ([Linq.Enumerable]::Sum([decimal[]]($deb_desc.DESC_MBPS )) / $deb_desc.DESC_MBPS.count )
			$avg_asc = ( [Linq.Enumerable]::Sum([decimal[]]($deb_asc.ASC_MBPS )) / $deb_asc.ASC_MBPS.count)

			$lbl_download.content = [math]::Round($avg_desc,2)
			$lbl_upload.content = [math]::Round($avg_asc,2)

			#Region lineCharts
			$seriesCollection = [LiveCharts.SeriesCollection]::new()
			#== Serie 1 ===
			$lineserie1 = [LiveCharts.Wpf.LineSeries]::new()
			    $lineserie1.Title = "Debit Montant"
				
			    $chartValues = [LiveCharts.ChartValues[double]]::new( )
				
				foreach ($asc in $deb_asc) {
			    $chartValues.Add($asc.ASC_MBPS) }
			    $lineserie1.Values  = $chartValues
			#== Serie 2 ===	
			$lineserie2 = [LiveCharts.Wpf.LineSeries]::new()
			$lineserie2.Title = "Debit Descendant"
			$chartValues = [LiveCharts.ChartValues[double]]::new( )
			foreach ($desc in $deb_desc){
			    $chartValues.Add($desc.DESC_MBPS) }
			$lineserie2.Values  = $chartValues

			$axisX.Labels = $malist
			$seriesCollection.Add($lineserie1)
			$seriesCollection.Add($lineserie2)
			#EndRegion
			
			#Region PieChart
			$piecollections = [LiveCharts.SeriesCollection]::new()
			
			$chartvalue1 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
			$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
			$chartvalue1.Add([LiveCharts.Defaults.ObservableValue]::new($ping_t_ok))
			$pieSeries.Values = $chartvalue1
			$pieSeries.Title = "Ping OK"
			$pieSeries.DataLabels = $true
			$piecollections.Add($pieSeries)
			
			$chartvalue2 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
			$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
			$chartvalue2.Add([LiveCharts.Defaults.ObservableValue]::new($ping_t_ko))
			$pieSeries.Values = $chartvalue2
			$pieSeries.Title = "Ping Tps KO"
			$pieSeries.DataLabels = $true
			$piecollections.Add($pieSeries)

			$chartvalue3 = [LiveCharts.ChartValues[LiveCharts.Defaults.ObservableValue]]::new()
			$pieSeries = [LiveCharts.Wpf.PieSeries]::new()
			$chartvalue3.Add([LiveCharts.Defaults.ObservableValue]::new($ping_total_ko))
			$pieSeries.Values = $chartvalue3
			$pieSeries.Title = "Ping KO"
			$pieSeries.DataLabels = $true
			$piecollections.Add($pieSeries)
			
			$lineChart.Series = $seriesCollection
			$pieChart.Series = $piecollections
			$lineChart.Visibility = "Visible"
			$pieChart.Visibility = "Visible"
			Log -Add "Generation Stat - $($exp_date) - $($teletravaileur)"
			
				 
		
		}#ELSE
	}#ELSE Menu vide
}#Else Version	
})



############## SETTINGS ############
$DragDrop.Add_PreviewDragOver({
    [System.Object]$script:sender = $args[0]
    [System.Windows.DragEventArgs]$e = $args[1]

    $e.Effects = [System.Windows.DragDropEffects]::Move
$e.Handled = $true
})

$DragDrop.Add_Drop({
 [System.Object]$script:sender = $args[0]
    [System.Windows.DragEventArgs]$e = $args[1]

    If($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)){

        $Script:Files =  $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        #Log -Add $Files.Count
		$cont = Get-Content $files 
		$ltv_list.ItemsSource = $cont
		$lbl_count_poste.content = "Nombre de postes : $($cont.count)"

}
})

$rdbt_AD.add_click({
		$ltv_list.Items.clear()
		SQL_TT_Postes
		$ltv_list.ItemsSource = $data_poste
		$lbl_count_poste.content = "Nombre de postes : $($data_poste.count)"
		})
$rdbt_fic.add_click({ 
		$ltv_list.Items.clear()
		$ltv_list.ItemsSource = $null })

$rdbt_one.add_click({
		$ltv_list.Items.clear()
		$ltv_list.ItemsSource = $null })

$btn_unit.add_click({
	if ($rdbt_one.isChecked -eq $false) { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Information","Merci de cocher l'option correspondante !") }
	else {
	$ltv_list.items.add($txtb_unit.get_text())
	$lbl_count_poste.content = "Nombre de postes : $($ltv_list.items.count)"
	}
})

$btn_go_sett.add_click({
$user_connected = Get-CimInstance -ClassName Win32_ComputerSystem | select Username
$user_connected = ($user_connected.Username.Split("\"))[1]

if ($user_connected  -ne "mbg17590") {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Information","Vous n'êtes pas habilité à effectuer cette action.") }
Else {
	if (($rdbt_ARRET.IsChecked -eq $false) -and ($rdbt_RELANCE.IsChecked -eq $false)) {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Information","Vous devez choisir via les boutons radio : ARRET ou RELANCE.") }
	Else {
		if ($rdbt_ARRET.IsChecked -eq $true ) { $arret_relance = 0 }
		else {$arret_relance = 1 }
		
		$global:syncHash = [hashtable]::Synchronized(@{})
		$global:syncHash.Current_Folder = $Current_Folder
		$global:syncHash.list_poste = $ltv_list.items
		$global:syncHash.arret_relance = $arret_relance 
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
			Import-Module Logs_TA
			Create_log -Path_Log $global:syncHash.Current_Folder -Appli ARRET_NETQOS
			#Log -Add "Lancement de la procédure d'Arrêt"
				$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$global:syncHash.lst_actions.Items.add("Lancement de la procédure d'Arrêt/relance")
					$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
					$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
				
				$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$global:syncHash.ProgressBar.Maximum=$global:syncHash.list_poste.count },"Normal")
					$ii = 0
			
		
			# RAPPORT
			$global:rapport_poste_OK = @()
			$global:rapport_poste_KO = @()

			Foreach ($nom_machine in $global:syncHash.list_poste){
				$global:poste_properties = @{
						Name = $nom_machine
						Online = ""
						Status = ""
						Val_Clef = ""
						Service = ""
						Processus = ""
						IP_NETBIOS = ""
					}
				
				$test_co = Test-Connection $nom_machine -Count 1 -ErrorAction SilentlyContinue
				if (!$test_co) { Log -Add "$($nom_machine) - Hors ligne"
																$global:syncHash.Window.Dispatcher.invoke([action]{ 
																$global:syncHash.lst_actions.Items.add("$($nom_machine) - Hors ligne")
																$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
																$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
																$global:poste_properties.Status = "KO"
																$global:poste_properties.Online = "KO"}
				Else {	try {
							Log -Add "$($nom_machine) - Vérification Netbios/IP"
							$Services_option = New-CimSessionoption -Protocol Dcom 
							$Services_session = New-CimSession -ComputerName $test_co.IPV4Address.IPAddressToString -SessionOption $Services_option 
							$verif = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $Services_session | select Name
							if ($verif.Name -ne $nom_machine) {Log -Add "$($nom_machine) - Pas de concordance IP/NETBIOS - Abandon"
																$global:syncHash.Window.Dispatcher.invoke([action]{ 
																	$global:syncHash.lst_actions.Items.add("$($nom_machine) - Verif Concordance IP/NETBIOS - KO")
																	$global:syncHash.lst_actions.Items.add("$($nom_machine) - Abandon")
																	$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
																	$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 																
																$global:poste_properties.IP_NETBIOS = "Not Match"}
							Else {
				
							$global:poste_properties.Online = "OK"
							$global:poste_properties.IP_NETBIOS = "Match"

							#Activation Session DCOM à distance
							$Services_option = New-CimSessionoption -Protocol Dcom 
							$Services_session = New-CimSession -ComputerName $nom_machine -SessionOption $Services_option 
							Log -Add "$($nom_machine) - Activation Session DCOM"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Activation Session DCOM")
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 


							#Démarrage du service RemoteRegistry pour modif base de registre à distance
							$dem_serv = Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='remoteregistry'" | Invoke-CimMethod -MethodName StartService
							Log -Add "$($nom_machine) - Démarrage Service RemoteRegistry"
								$global:syncHash.Window.Dispatcher.invoke([action]{ 
									$global:syncHash.lst_actions.Items.add("$($nom_machine) - Démarrage Service RemoteRegistry")
									$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
									$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 

							$verif = Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='remoteregistry'" 
							Log -Add "$($nom_machine) - Etat du service : $($verif.state)"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Etat du service : $($verif.state)") 
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 


							if ($verif.State -eq 'Running'){
								$check_key=[microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE","N35R545").OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro", $true).GetSubKeyNames()
								if ($check_key -contains "NETQOS"){
								Log -Add "$($nom_machine) - Présence de la clef de registre NETQOS"
								$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Présence de la clef de registre NETQOS") },"Normal")
							
								if ($global:syncHash.arret_relance -eq 0) {
								#Modification Clef de Registre NETQOS pour arrêter le lancement auto au démarrage du poste
								$distant=[microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE",$nom_machine).OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS", $true).SetValue('Chemin',"Desactivation")
									Log -Add "$($nom_machine) - Modification Clef de Registre NETQOS pour arrêter le lancement auto"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Modification Clef de Registre NETQOS pour arrêter le lancement auto") },"Normal") 
									Log -Add "$($nom_machine) - Modif HKLM\SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS\Chemin => Desactivation"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Modif HKLM\SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS\Chemin => Desactivation") },"Normal") 
									#$rapport_poste_OK += $nom_machine
									$global:poste_properties.Status = "OK"
									$global:poste_properties.Val_Clef = [microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE",$nom_machine).OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS", $true).GetValue('Chemin')

								}	
								else {
								#Modification Clef de Registre NETQOS pour reprendre le lancement auto au démarrage du poste
								$distant=[microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE",$nom_machine).OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS", $true).SetValue('Chemin',"C:\AppDSI\EXPGLB\NETQOS\NETQOS_Launch_Teletravail.exe")
									Log -Add "$($nom_machine) - Modification Clef de Registre NETQOS pour reprendre le lancement auto"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Modification Clef de Registre NETQOS pour reprendre le lancement auto")
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 
									Log -Add "$($nom_machine) - Modif HKLM\SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS\Chemin => C:\AppDSI\EXPGLB\NETQOS\NETQOS_Launch_Teletravail.exe"
									$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Modif HKLM\SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS\Chemin => C:\AppDSI\EXPGLB\NETQOS\NETQOS_Launch_Teletravail.exe")
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 
									#$rapport_poste_OK += $nom_machine
									$global:poste_properties.Status = "OK"
									$global:poste_properties.Val_Clef = [microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE",$nom_machine).OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS", $true).GetValue('Chemin')

								}
								}#IF CHECK KEY
								else {Log -Add "$($nom_machine) - Absence de la clef de registre NETQOS - Appli Non installée sur le poste"
										$global:poste_properties.Val_Clef = "None"
										$global:syncHash.Window.Dispatcher.invoke([action]{ 
											$global:syncHash.lst_actions.Items.add("$($nom_machine) - Absence de la clef de registre NETQOS - Appli Non installée sur le poste") },"Normal") } #ELSE CHECK KEY
							} #IF SERVICE RUNNING
							else { $global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Modification de la Clef impossible - Service en STOP") },"Normal")
										Log -Add "$($nom_machine) - Modification de la Clef impossible - Service en STOP"
									#$rapport_poste_KO += $nom_machine
									$global:poste_properties.Status = "KO"
									$global:poste_properties.Val_Clef = [microsoft.win32.registrykey]::openremotebasekey("LOCALMACHINE",$nom_machine).OpenSubKey("SOFTWARE\Groupama\Wlogon\Logon\NonSynchro\NETQOS", $true).GetValue('Chemin')
									}
							#Arrêt du service RemoteRegistry
							$arret = Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='remoteregistry'" | Invoke-CimMethod -MethodName StopService
							Log -Add "$($nom_machine) - Arrêt du service RemoteRegistry"
							$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Arrêt du service RemoteRegistry") 
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") 
										
							$verif = Get-CimInstance -ClassName win32_service -CimSession $Services_session -Filter "Name='remoteregistry'" 
							if ($verif.State -eq 'Stopped'){ Log -Add "$($nom_machine) - Etat du service : $($verif.state)"
															$global:syncHash.Window.Dispatcher.invoke([action]{ 
																$global:syncHash.lst_actions.Items.add("$($nom_machine) - Etat du service : $($verif.state)")
																$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
																$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal") }
							$global:poste_properties.Service = $verif.State
							
							#### ARRET PROCESS NETQOS ####
							if ($global:syncHash.arret_relance -eq 0) {
							Log -Add "$($nom_machine) - Arret des processus NETQOS"
							$global:syncHash.Window.Dispatcher.invoke([action]{ 
								$global:syncHash.lst_actions.Items.add("$($nom_machine) - Arret des processus NETQOS")
								$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
								$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
								
							$process = Get-CimInstance -ClassName win32_process -CimSession $Services_session 
							foreach ($proc in $process) {
							$owner = $proc | Invoke-CimMethod -MethodName GetOwner #| where { $_.User -eq "mbg17590"}
								if ($owner.User -eq "b_netqos") {
									try { $result = $proc | Invoke-CimMethod -MethodName Terminate
											if ($result.ReturnValue -eq 0) {$global:poste_properties.Processus = "Stopped"
																				Log -Add "$($nom_machine) - Processus NETQOS arretes"
																				$global:syncHash.Window.Dispatcher.invoke([action]{ 
																						$global:syncHash.lst_actions.Items.add("$($nom_machine) - Processus NETQOS arretes")
																						$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
																						$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")}
											Else {$global:poste_properties.Processus = "Running"
													Log -Add "$($nom_machine) - Erreur - Processus NETQOS toujours en cours"
													$global:syncHash.Window.Dispatcher.invoke([action]{ 
															$global:syncHash.lst_actions.Items.add("$($nom_machine) - Erreur - Processus NETQOS toujours en cours")
															$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
															$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")} # ELSE
										} # TRY
									catch { log -Add $Error[0] }
								} #IF
								Else {$global:poste_properties.Processus = "None"}
							
						} #FOREACH PROCESS
							
							if ( $global:poste_properties.Processus -eq "None") {
								Log -Add "$($nom_machine) - Aucun processus NETQOS en cours"
								$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Aucun processus NETQOS en cours")
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
								}
						} # IF ARRET NETQOS => ARRET PROCESS
						
							#Fermeture Session DCOM
							Get-CimSession | where {$_.ComputerName -eq $Services_session.ComputerName} | Remove-CimSession
							Log -Add "$($nom_machine) - Fermeture Session DCOM"
							$global:syncHash.Window.Dispatcher.invoke([action]{ 
										$global:syncHash.lst_actions.Items.add("$($nom_machine) - Fermeture Session DCOM")
										$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
										$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
							} #ELSE NETBIOS/IP
						} #TRY
						Catch { $Error[0] | Out-File C:\temp\error.txt }
					}#ELSE TEST-CONNECTION
					
					
					
				$global:rapposte = New-Object PSObject -Property $global:poste_properties	
					
				if ($global:poste_properties.Status -eq "Ok"){$global:rapport_poste_OK += $global:rapposte }
				Else { $global:rapport_poste_KO += $global:rapposte }
				Log -Add $global:rapposte
				
		
				$ii++	
				$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$global:syncHash.ProgressBar.value = $ii },"Normal")	
					

				}#FOREACH COMPUTER
			
			Log -Add "Fin de la procédure d'Arrêt"
			$global:syncHash.Window.Dispatcher.invoke([action]{ 
				$global:syncHash.lst_actions.Items.add("Fin de la procédure d'Arrêt/Relance")
					$view_last = $global:syncHash.lst_actions.SelectedItem=$global:syncHash.lst_actions.Items[($global:syncHash.lst_actions.Items.count - 1)]
					$global:syncHash.lst_actions.ScrollIntoView($view_last)},"Normal")
			$global:syncHash.Window.Dispatcher.invoke([action]{ 
					$global:syncHash.rapport_poste_KO= $global:rapport_poste_KO 
					$global:syncHash.rapport_poste_OK= $global:rapport_poste_OK },"Normal")

			} catch { $Error[0] | Out-File C:\temp\error.txt }	

		    })



		$global:ThemeFile = "$($global:syncHash.Current_Folder)\Theme.xaml"
[xml]$global:xa_ml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
	    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
        Title="MainWindow" Height="440" Width="600" WindowStyle="ToolWindow">
		<Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$ThemeFile"/> 
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
        </Window.Resources>
    <Grid>
        <StackPanel>
			<StackPanel Orientation="Horizontal" Background="#FF3F5366">
            <Label Content="NETQOS - Détails des actions en cours" FontWeight="Bold" Foreground="White" Height="50" FontSize="14"/>
            </StackPanel>
			<Border>
                <ListView Name="lst_actions" Height="250" Margin="5"/>
            </Border>
            <Separator Height="10"/>
            <ProgressBar Name="ProgressBar" Height="10" Margin="5,0,5,0"/>
            <Separator Height="10"/>
            <StackPanel Orientation="Horizontal">
                <Button Name="btn_rapport_actions" Content="Rapport" Margin="10,0,0,0" Width="70"/>
                <Button Name="btn_close_actions" Content="Fermer" Width="70" Height="40" HorizontalAlignment="Right" Margin="430,0,0,0"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@ 

			$global:reader=(New-Object System.Xml.XmlNodeReader $xa_ml)
		    $global:syncHash.window=[Windows.Markup.XamlReader]::Load($reader)
			$global:syncHash.lst_actions = $global:syncHash.window.FindName("lst_actions")
			$global:syncHash.btn_rapport_actions = $global:syncHash.window.FindName("btn_rapport_actions")
			$global:syncHash.btn_close_actions = $global:syncHash.window.FindName("btn_close_actions")
			$global:syncHash.ProgressBar = $global:syncHash.window.FindName("ProgressBar")
			
			
				$psScript.Runspace = $maj_Runspace
		      	$psScript.BeginInvoke()
			
			$global:syncHash.btn_close_actions.add_Click({	
					$global:syncHash.window.Close() 
					$global:newRunspace.close() })
					
function Get-StyleSheet {
    [CmdletBinding()]
    Param()
@"
<style>
body {
    font-family:Segoe,Tahoma,Arial,Helvetica;
    font-size:10pt;
    color:#333;
    background-color:#eee;
    margin:10px;
}
th {
    font-weight:bold;
    color:white;
    background-color:#333;
}
</style>
"@
}
					
			$global:syncHash.btn_rapport_actions.add_Click({

					$global:syncHash.rapport_poste_KO | Select Name | Out-File C:\temp\NETQOS_AR_Rapport_KO.txt
					$global:syncHash.rapport_poste_OK | Select Name | Out-File C:\temp\NETQOS_AR_Rapport_OK.txt
				$frag_01 = $global:syncHash.rapport_poste_KO | Select Name,Online,Status,Val_Clef,Service,Processus | ConvertTo-Html -Fragment -As table -PreContent "<h2>Liste des postes KO</h2>" | Out-String
				$frag_02 = $global:syncHash.rapport_poste_OK | Select Name,Online,Status,Val_Clef,Service,Processus | ConvertTo-Html -Fragment -As table -PreContent "<h2>Liste des postes OK</h2>" | Out-String
					 			
			$style = Get-StyleSheet
			ConvertTo-HTML -Title "Rapport ARRET/RELANCE NETQOS" -Head $style -Body "<h1>Rapport NETQOS - $exp_date</h1>",$frag_01,$frag_02 | Out-File C:\temp\NETQOS_ARRET_RELANCE.html
			
			start C:\temp\NETQOS_ARRET_RELANCE.html
			})
		    
			$global:syncHash.window.ShowDialog() 
			$global:newRunspace.close()
			$global:newRunspace.dispose()
		 
		}

		$global:psCmd = [PowerShell]::Create().AddScript($global:script_code )
		$global:psCmd.Runspace = $global:newRunspace
		$global:data = $global:psCmd.BeginInvoke()
		}#ELSE ARRET/RELANCE
}#ELSE HABILLITATION
})


#	EVENEMENTS
#-------------------------------------------------------------------------

$Window_Main.add_Loaded({	
})

$Window_Main.add_ContentRendered({	
})

$btn_export_home.add_click({ })
$btn_export_glob.add_click({ })

$mn_exp_csv_home.add_Click({
	if ($rdbt_avg.isChecked -eq $true) { 
							$data_user | Export-Csv C:\temp\Netqos_infos_$($exp_data)_$($exp_date).csv
							$data_speed | Export-Csv C:\temp\Netqos_AVG_$($exp_data)_$($exp_date).csv 
							Log -Add "Export csv $($exp_data) - $($exp_date)"}
	else {
	$data_user | Export-Csv C:\temp\Netqos_infos_$($exp_data)_$($exp_date).csv
	$data_speed | Export-Csv C:\temp\Netqos_ST_$($exp_data)_$($exp_date).csv 
	$data_ping | Export-Csv C:\temp\Netqos_P_$($exp_data)_$($exp_date).csv 
	Log -Add "Export csv $($exp_data) - $($exp_date)"
	}
$Title = "NETQOS - Rapport créé !"
$Text = "Les fichiers CSV ont été enregistrés dans le dossier C:\Temp. Cliquez ici pour ouvrir le dossier de destination."
Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\happy.png"  -do_click 2
	
})
$mn_exp_csv_glob.add_Click({
$data_totaux | Export-Csv C:\temp\Netqos_$($exp_data)_$($exp_date).csv
Log -Add "Export csv $($exp_data) - $($exp_date)"
$Title = "NETQOS - Rapport créé !"
$Text = "Les fichiers CSV ont été enregistrés dans le dossier C:\Temp. Cliquez ici pour ouvrir le dossier de destination."
Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\happy.png"  -do_click 2})

$mn_exp_html_home.add_click({ 	
if ($rdbt_avg.isChecked -eq $true) { $frag01 = $data_user | Select DESC_PDT,ID_PDT,MODELE_PDT,IP_PRIVATE,IP_PULSE,IP_PUBLIC | ConvertTo-Html -Fragment -As table -PreContent "<h2>Informations sur le poste</h2>" | Out-String
									 $frag02 = $data_user_avg | Select TRIGRAMME,FIRSTNAME,LASTNAME,SERVICE,ID_PDT,DATE_TEST,DATE,AVG_DESC,AVG_ASC,AVG_TPS_AVG_MS,AVG_TPS_MAX_MS | ConvertTo-Html -Fragment -As table -PreContent "<h2>Données globales récoltées</h2>" | Out-String
									 $script:html_name = "C:\temp\NETQOS_rapport_$($exp_data)_AVG_$($exp_date).html"
									 
									 $style = Get-StyleSheet
										ConvertTo-HTML -Title "Rapport NETQOS" -Head $style -Body "<h1>Rapport NETQOS - $exp_date</h1>",$frag01,$frag02 | Out-File $html_name

										if ((Test-Path $html_name) -eq $true){
											$Title = "NETQOS - Rapport créé !"
											$Text = "Le rapport a été enregistré dans le dossier C:\Temp. Cliquez ici pour l ouvrir."
										Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\happy.png" -do_click 1 }
										Log -Add "Export HTML $($exp_data) - $($exp_date)"
}
	else { 
	$frag01 = $data_user | Select DESC_PDT,ID_PDT,MODELE_PDT,IP_PRIVATE,IP_PULSE,IP_PUBLIC | ConvertTo-Html -Fragment -As table -PreContent "<h2>Informations sur le poste</h2>" | Out-String
	$frag02 = $data_speed | Select DATE_TEST_DEBIT,IP_PULSE,DESC_PDT,ASC_MBPS,ID_ASC,DESC_MBPS | ConvertTo-Html -Fragment -As table -PreContent "<h2>Données des SpeedTest</h2>" | Out-String
	$frag03 = $data_ping | Select DATE_TEST_PING,IP_PULSE,DESC_PDT,PING_TYPE,PING_PCT_OK,TPS_MAX_MS,TPS_MIN_MS,TPS_AVG_MS | ConvertTo-Html -Fragment -As table -PreContent "<h2>Données des ping</h2>" | Out-String
	$script:html_name = "C:\temp\NETQOS_rapport_$($exp_data)_$($exp_date).html"
	
		$style = Get-StyleSheet
		ConvertTo-HTML -Title "Rapport NETQOS" -Head $style -Body "<h1>Rapport NETQOS - $exp_date</h1>",$frag01,$frag02,$frag03 | Out-File $html_name

		if ((Test-Path $html_name) -eq $true){
		$Title = "NETQOS - Rapport créé !"
		$Text = "Le rapport a été enregistré dans le dossier C:\Temp. Cliquez ici pour l ouvrir."
		Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\happy.png" -do_click 1}
		Log -Add "Export HTML $($exp_data) - $($exp_date)"
	}
	
	
})
$mn_exp_html_glob.add_click({ 
switch -wildcard ($exp_data){
	"Tout le monde" {$frag01 = $data_totaux | Select Trigramme,Firstname,Lastname,Service,ID_PDT,DATE_TEST,DATE,AVG_DESC,AVG_ASC,AVG_TPS_AVG_MS,AVG_TPS_MAX_MS |  ConvertTo-Html -Fragment -As table -PreContent "<h2>L'ensemble des personnes s'étant connecté à distance</h2>" | Out-String 
						$script:html_name = "C:\temp\NETQOS_rapport_all_$($exp_date).html"}
	"Teletravailleurs" {$frag01 = $result_tt | Select Trigramme,Firstname,Lastname,Service,ID_PDT,DATE_TEST,DATE,AVG_DESC,AVG_ASC,AVG_TPS_AVG_MS,AVG_TPS_MAX_MS |  ConvertTo-Html -Fragment -As table -PreContent "<h2>Télétravailleurs uniquement</h2>" | Out-String 
						$script:html_name = "C:\temp\NETQOS_rapport_TT_$($exp_date).html"}
}
$style = Get-StyleSheet
ConvertTo-HTML -Title "Rapport NETQOS" -Head $style -Body "<h1>Rapport NETQOS - $exp_date</h1>",$frag01 | Out-File $html_name

if ((Test-Path $html_name) -eq $true){
	$Title = "NETQOS - Rapport créé !"
	$Text = "Le rapport a été enregistré dans le dossier C:\Temp. Cliquez ici pour l ouvrir."
	Toast_Message -Title $Title -Text $Text -Image "$($Image_Folder)\happy.png" -do_click 1}
Log -Add "Export HTML $($exp_data) - $($exp_date)"
})


$btn_send_comm.add_click({
$date = Get-Date -Format ddMMyyyy
	$name_comm = $txt_name.text
	$email_comm = $txt_email.text
	$comm_comm = new-object System.Windows.Documents.TextRange($rtbx_comm.Document.ContentStart,$rtbx_comm.Document.ContentEnd  )

	if (($name_comm -eq $null) -or ($email_comm -eq $null) -or ($comm_comm.Text -eq $null)) {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Tous les champs sont obligatoires. Merci.") }
	Else {
		if ($email_comm -notmatch '[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$') {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","L'adresse mail renseignée ne respecte pas le format. Merci de vérifier votre saisie.") }
		Else {
				$encodingMail = [System.Text.Encoding]::UTF8
				$Subject = "NETQOS - Formulaire/Contact - $($name_comm) "
				$To =  "mail"
				$From =  "mail"
				$Relay = "smtp.res.local"
				
				$Body = "Email : $($email_comm)"
				$Body += "`nDate : ($date)"
				$Body += "`nMessage de $($name_comm) :"
				$Body += "`n$($comm_comm.text)"
				
				Send-MailMessage -To $To -From $From -Subject $Subject -SmtpServer $Relay -BodyAsHtml $Body -Encoding $encodingMail
				[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Information","Le mail a bien été envoyé. Merci pour votre participation.")
				$txt_name.text = $txt_email.text = ""
				$comm_comm.text = ""
				
				Log -Add "Mail envoyé"
		}#ELSE MAIL
	
	
	}#ELSE EMPTY
})


$btn_go.add_click({ 
$teletravaileur = $cbx_user.get_text()
$script:exp_data = $teletravaileur
#$teletravaileur | Out-File C:\temp\user.txt
### REINIT DTG/LABEL
$dtg_ping.ItemsSource = $null
$dtg_speed.ItemsSource = $null
$dtg_ping.Visibility = "Visible"
$dtg_speed.Height = "174.5"
$lbl_netbios.Content = ""
$lbl_ip_pub.Content = ""
$lbl_model.Content = ""
$lbl_ip_priv.Content = ""
$ico_dtg_ping.visibility = "Visible"

if ($teletravaileur -eq $null -or $teletravaileur -eq "") { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Merci de sélectionner un utilisateur dans la liste déroulante.") }
Else {
	$date = $dtpr_date.selecteddate.tostring().split(" ")[0]
	$date = Get-Date $date -Format yyyyMMdd
    $script:exp_date = $date
	if ($dtpr_date_fin.visibility -eq "Visible") { 
	$date_fin = $dtpr_date_fin.selecteddate.tostring().split(" ")[0]
	$date_fin = Get-Date $date_fin -Format yyyyMMdd }
	else { $date_fin = $date }
	$hordeb = $cbx_hour_one.get_text().Trim()
	$horfin = $cbx_hour_two.get_text().Trim()
	$hor_deb = Get-Date $hordeb -Format "HH:mm:ss"
	$hor_fin = Get-Date $horfin -Format "HH:mm:ss"
	$teletravaileur = $cbx_user.get_text().Split("-")[1].trim()
	$script:exp_data = $teletravaileur
	
	if ($rdbt_avg.isChecked -eq $true) 
					{ SQL_TT_AVG
						if ($data_user_avg -eq $null -and $data_ping -eq $null) {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée disponible à cette date.") }
						else { 
						    if ($data_user_avg.count -eq $null) { $dtg_speed.ItemsSource = $data_user_avg.DataView }
	                        Else {
							$dtg_speed.ItemsSource = $data_user_avg | Select @{n='Date';e={Get-date ($_.DATE) -Format yyyy/MM/dd}},@{n='Trigramme';e={$_.TRIGRAMME}},@{n='Prénom';e={$_.FIRSTNAME}},@{n='Nom';e={$_.LASTNAME}},@{n='Service';e={$_.SERVICE}},@{n='Poste';e={$_.ID_PDT}},@{n='Moy DESC';e={$_.AVG_DESC}},@{n='Moy ASC';e={$_.AVG_ASC}},@{n='Moy Ping';e={$_.AVG_TPS_AVG_MS}},@{n='Moy Ping Max';e={$_.AVG_TPS_MAX_MS}},ID_DESC,ID_ASC,ID_PING,ID_PING_MAX
	                              }	
							if ($data_user.ID_PDT.count -ne 1){
							$lbl_netbios.Content = $data_user.ID_PDT[0]
							$lbl_ip_pub.Content = $data_user.IP_PUBLIC[0]
							$lbl_model.Content = $data_user.MODELE_PDT[0]
							$lbl_ip_priv.Content = $data_user.IP_PRIVATE[0] }
							Else {
							$lbl_netbios.Content = $data_user.ID_PDT
							$lbl_ip_pub.Content = $data_user.IP_PUBLIC
							$lbl_model.Content = $data_user.MODELE_PDT
							$lbl_ip_priv.Content = $data_user.IP_PRIVATE }
							}
					$ico_dtg_ping.visibility = "Hidden"		
					$dtg_ping.Visibility = "Hidden"
					$dtg_speed.Height = "401"
					Log -Add "Collecte donnees personne $($exp_data) - $($exp_date)"
						 
					}
	Else {
		SQL_TT_AllData
		if ($data_speed -eq $null -and $data_ping -eq $null) {[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée disponible à cette date.") }
		Else { $dtg_speed.ItemsSource = $data_speed  | select @{n='Date';e={get-date ($_.DATE_TEST_DEBIT) -Format 'yyyy/MM/dd HH:mm:ss'}},@{n='IP Pulse';e={$_.IP_PULSE}},@{n='Carte Réseau';e={$_.DESC_PDT}},@{n='Débit Montant';e={$_.ASC_MBPS}},@{n='Débit Descendant';e={$_.DESC_MBPS}} ,ID_ASC,ID_DESC 
				$dtg_ping.ItemsSource = $data_ping | select @{n='Date';e={get-date ($_.DATE_TEST_PING) -Format 'yyyy/MM/dd HH:mm:ss'}},@{n='IP Pulse';e={$_.IP_PULSE}},@{n='Carte Réseau';e={$_.DESC_PDT}},@{n='Dest';e={$_.PING_TYPE}},@{n='Ping ok sur 100';e={$_.PING_PCT_OK}},@{n='Tps Max';e={$_.TPS_MAX_MS}},@{n='Tps Min';e={$_.TPS_MIN_MS}},@{n='Moyenne Tps';e={$_.TPS_AVG_MS}},ID_PING 
				if ($data_user.ID_PDT.count -ne 1){
							$lbl_netbios.Content = $data_user.ID_PDT[0]
							$lbl_ip_pub.Content = $data_user.IP_PUBLIC[0]
							$lbl_model.Content = $data_user.MODELE_PDT[0]
							$lbl_ip_priv.Content = $data_user.IP_PRIVATE[0]
						}
						Else {
						$lbl_netbios.Content = $data_user.ID_PDT
						$lbl_ip_pub.Content = $data_user.IP_PUBLIC
						$lbl_model.Content = $data_user.MODELE_PDT
						$lbl_ip_priv.Content = $data_user.IP_PRIVATE
						}
				Log -Add "Collecte donnees personne $($exp_data) - $($exp_date)"
				}
	}
}
})


$btn_go_global.add_click({ 
$teletravaileur = $cbx_user_global.get_text()
$script:exp_data = $teletravaileur
$dtg_global.ItemsSource = $null

if ($teletravaileur -eq $null -or $teletravaileur -eq "") { [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Merci de sélectionner un utilisateur dans la liste déroulante.") }
Else {
	$date = $dtpr_date_global.selecteddate.tostring().split(" ")[0]
	$date = Get-Date $date -Format yyyyMMdd
    $script:exp_date = $date
	if ($dtpr_date_fin_global.visibility -eq "Visible") { 
	$date_fin = $dtpr_date_fin_global.selecteddate.tostring().split(" ")[0]
	$date_fin = Get-Date $date_fin -Format yyyyMMdd }
	else { $date_fin = $date }
	#$teletravaileur = $cbx_user.get_text().Split("-")[1].trim()

		switch -wildcard ($teletravaileur) {
			"Tout le monde" {
					try { SQL_EveryBody
						if ($data_totaux  -eq $null){[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée à cette date.") }
						Else {
							 $dtg_global.ItemsSource = $data_totaux | select @{n='Date';e={Get-date ($_.DATE) -Format yyyy/MM/dd}},@{n='Trigramme';e={$_.TRIGRAMME}},@{n='Prénom';e={$_.FIRSTNAME}},@{n='Nom';e={$_.LASTNAME}},@{n='Service';e={$_.SERVICE}},@{n='Poste';e={$_.ID_PDT}},@{n='Moy DESC';e={$_.AVG_DESC}},@{n='Moy ASC';e={$_.AVG_ASC}},@{n='Moy Ping';e={$_.AVG_TPS_AVG_MS}},@{n='Moy Ping Max';e={$_.AVG_TPS_MAX_MS}},ID_DESC,ID_ASC,ID_PING,ID_PING_MAX
							}
					} catch {$Error | Add-Content "$($Current_Folder)\error_log.txt"}

			} #TOUT LE MONDE
			
			"Teletravailleurs" {
					try { SQL_TeleTravail
						if ($result_tt  -eq $null){[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Window_Main,"Attention","Aucune donnée à cette date.") }
						Else {
                             if ($result_tt.count -eq $null){ $dtg_data.ItemsSource = $result_tt.DataView }
                             Else {
    						$dtg_global.ItemsSource = $result_tt | select @{n='Date';e={Get-date ($_.DATE) -Format yyyy/MM/dd}},@{n='Trigramme';e={$_.TRIGRAMME}},@{n='Prénom';e={$_.FIRSTNAME}},@{n='Nom';e={$_.LASTNAME}},@{n='Service';e={$_.SERVICE}},@{n='Poste';e={$_.ID_PDT}},@{n='Moy DESC';e={$_.AVG_DESC}},@{n='Moy ASC';e={$_.AVG_ASC}},@{n='Moy Ping';e={$_.AVG_TPS_AVG_MS}},@{n='Moy Ping Max';e={$_.AVG_TPS_MAX_MS}},ID_DESC,ID_ASC,ID_PING,ID_PING_MAX
                            }
						}
					}
					catch {$Error | Add-Content "$($Current_Folder)\error_log.txt" } # C:\AppDSI\EXPGLB\NETQOS\launch.txt }
					} #TELETRAVAILLEURS

	}#SWITCH
	Log -Add "Collecte donnees globale $($teletravaileur) - $($exp_date)"
} #IF/ELSE
})

#
# Tout est chargé on ferme le Splash Screen
#

Start-sleep -Seconds 2
Close-SplashScreen
Log -Add "Fermeture SplashScreen"

#
# On ferme la runspacepool à la fermeture de l'applciation
#
$btn_quit.add_Click({ $Window_Main.Close()
						exit })

$Window_Main.add_Closing({
})

#
# Afficher la form
#

$null = $Window_Main.ShowDialog()