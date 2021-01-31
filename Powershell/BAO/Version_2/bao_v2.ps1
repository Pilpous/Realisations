#
#	INIT
#-------------------------------------------------------------------------

Param
 (
 [String]$Restart
 
 )
 
If ($Restart -ne "") 
 {
  Start-Sleep 2
 }

$Global:Current_Folder = $PSScriptRoot
Set-Location $Current_Folder
$Assembly_Folder = Join-Path -Path $Current_Folder -ChildPath .\assembly

$icon = "D:\Users\mbg17590\Pictures\Icones\champi.ico"


#	ASSEMBLIES
#-------------------------------------------------------------------------
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 	 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 		 | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | out-null

[System.Windows.Forms.Application]::EnableVisualStyles()

#
# Chargement des dll du dossier assembly
#

If(Test-Path $Assembly_Folder) {
	Foreach ($Assembly in (Get-ChildItem $Assembly_Folder -Filter *.dll)) {
		$null = [System.Reflection.Assembly]::LoadFrom($Assembly.fullName)
	}
}

#	XAML
#-------------------------------------------------------------------------
#	SPLASH SCREEN
#-------------------------------------------------------------------------
#Region SplashScreen
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
	Title="SplashScreen" WindowStyle="None" WindowStartupLocation="CenterScreen"
	AllowsTransparency="True" Background="Transparent" ShowInTaskbar ="False" 
    ResizeMode = "NoResize" Width="500" Height="300"  >
    <Grid Background="Black">
        <StackPanel Orientation="Vertical" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="5,5,5,5">
            <Label Content="TOOLS BOX" Foreground="SlateGray" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="32" FontWeight="Bold"  />
            <Label Name="LoadingLabel" Content="Loading"  Foreground="SlateGray" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="24" Margin = "0,0,0,0"/>
            <Controls:ProgressRing IsActive="{Binding IsActive}" Foreground="SlateGray" HorizontalAlignment="Center"  Margin = "0,30,0,0" Width="40" Height="8"/>
        </StackPanel>
    </Grid>
</Window> 
"@
$reader = New-Object System.Xml.XmlNodeReader $xml
$hash.window = [Windows.Markup.XamlReader]::Load($reader) 
$hash.window.ShowDialog() 
})

##
## Ouverture Splash Screen
##
Start-SplashScreen

#EndRegion

Function LoadXaml($Global:filename) {
	$xamlLoader = (New-Object System.Xml.XmlDocument)
	$xamlLoader.Load($filename)
	$xamlLoader
}

#
# Chargement xaml des fenetres
#

# Fenetre Login
$login_xaml = LoadXaml("$Current_Folder\Page1.xaml") 
$xamlReader_login = (New-Object System.Xml.XmlNodeReader $login_xaml)
$Login_Window =  [Windows.Markup.XamlReader]::Load($xamlReader_login)


# Fenetre principale
$xaml = LoadXaml("$Current_Folder\MainWindow_bis.xaml") 
$xamlReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window_Main = [Windows.Markup.XamlReader]::Load($xamlReader)

# Fenetre Config
$config_xaml = LoadXaml("$Current_Folder\config.xaml") 
$xamlReader_config = (New-Object System.Xml.XmlNodeReader $config_xaml)
$Config_Window = [Windows.Markup.XamlReader]::Load($xamlReader_config)

#Region ControlMainWindows
# CONTROLES
$title_bar = $Window_Main.FindName("title_bar")
$txt_search = $Window_Main.FindName("txt_search")
$btn_search = $Window_Main.FindName("btn_search")
$btn_flyout = $Window_Main.FindName("btn_flyout")
$flyout = $Window_Main.FindName("flyout")
###############INFORMATIONS#######################
$btn_copy = $Window_Main.FindName("btn_copy")
$lbl_name = $Window_Main.FindName("lbl_name")
$lbl_trig = $Window_Main.FindName("lbl_trig")
$mn_AD = $Window_Main.FindName("mn_AD")
$mn_unlock = $Window_Main.FindName("mn_unlock")
$mn_reinit = $Window_Main.FindName("mn_reinit")
$lbl_idsigma = $Window_Main.FindName("lbl_idsigma")
$btn_reinit_sigma = $Window_Main.FindName("btn_reinit_sigma")
$lbl_idgrc = $Window_Main.FindName("lbl_idgrc")
$btn_reinit_grc = $Window_Main.FindName("btn_reinit_grc")
#$lbl_netbios = $Window_Main.FindName("lbl_netbios")
$cbx_netbios = $Window_Main.FindName("cbx_netbios")
$mn_pdm = $Window_Main.FindName("mn_pdm")
$mn_remote = $Window_Main.FindName("mn_remote")
$mn_msra = $Window_Main.FindName("mn_msra")
$mn_exp_comp = $Window_Main.FindName("mn_exp_comp")
$lbl_service = $Window_Main.FindName("lbl_service")
$lbl_ville = $Window_Main.FindName("lbl_ville")
$lbl_phone = $Window_Main.FindName("lbl_phone")
$lbl_mail = $Window_Main.FindName("lbl_mail")
$btn_mail = $Window_Main.FindName("btn_mail")
##############COMPTE WINDOWS########################
$exp_ad = $Window_Main.FindName("exp_ad")
$txtb_pwd_trig = $Window_Main.FindName("txtb_pwd_trig")
$btn_pwd_trig = $Window_Main.FindName("btn_pwd_trig")
$lbl_pwd_age = $Window_Main.FindName("lbl_pwd_age")
$lbl_pwd_state = $Window_Main.FindName("lbl_pwd_state")
$btn_pwd_unlock = $Window_Main.FindName("btn_pwd_unlock")
$txtb_pwd_password = $Window_Main.FindName("txtb_pwd_password")
$btn_pwd_reinit = $Window_Main.FindName("btn_pwd_reinit")
$btn_pwd_console = $Window_Main.FindName("btn_pwd_console")
$lbl_confirm_reinit = $Window_Main.FindName("lbl_confirm_reinit")
################POSTE######################
$exp_computer = $Window_Main.FindName("exp_computer")
$pg_ring_comput = $Window_Main.FindName("pg_ring_comput")
$cbx_comp_computname = $Window_Main.FindName("cbx_comp_computname")
$btn_comp_comput = $Window_Main.FindName("btn_comp_comput")
$lbl_comp_netbios = $Window_Main.FindName("lbl_comp_netbios")
$lbl_comp_ip = $Window_Main.FindName("lbl_comp_ip")
$lbl_comp_modele = $Window_Main.FindName("lbl_comp_modele")
$lbl_comp_os = $Window_Main.FindName("lbl_comp_os")
$lbl_comp_lastboot = $Window_Main.FindName("lbl_comp_lastboot")
$lbl_comp_sccm_tit = $Window_Main.FindName("lbl_comp_sccm_tit")
$lbl_comp_localisation = $Window_Main.FindName("lbl_comp_localisation")
$lbl_comp_STA_ip = $Window_Main.FindName("lbl_comp_STA_ip")
$lbl_comp_STA_state = $Window_Main.FindName("lbl_comp_STA_state")
$btn_com_remote_ip = $Window_Main.FindName("btn_com_remote_ip")
$mn_sta = $Window_Main.FindName("mn_sta")
$mn_comp_sta_pdm = $Window_Main.FindName("mn_comp_sta_pdm")
$mn_comp_sta_reboot = $Window_Main.FindName("mn_comp_sta_reboot")
$mn_comp_sta_num = $Window_Main.FindName("mn_comp_sta_num")
############TELEPHONIE##########################
$txt_search_genesys_num = $Window_Main.FindName("txt_search_genesys_num")
$btn_search_genesys_num = $Window_Main.FindName("btn_search_genesys_num")
$txt_search_genesys = $Window_Main.FindName("txt_search_genesys")
$btn_search_genesys = $Window_Main.FindName("btn_search_genesys")
$lbl_search_genesys_num = $Window_Main.FindName("lbl_search_genesys_num")
$lbl_search_genesys = $Window_Main.FindName("lbl_search_genesys")
$dtg_hist_num = $Window_Main.FindName("dtg_hist_num")
$dtg_hist_ut = $Window_Main.FindName("dtg_hist_ut")
############FLYOUT##########################
$btn_fl_wiki = $Window_Main.FindName("btn_fl_wiki")
$btn_fl_pdsi = $Window_Main.FindName("btn_fl_pdsi")
$btn_fl_po = $Window_Main.FindName("btn_fl_po")
$btn_fl_ig = $Window_Main.FindName("btn_fl_ig")
$btn_fl_sp = $Window_Main.FindName("btn_fl_sp")
$btn_fl_smart = $Window_Main.FindName("btn_fl_smart")
$btn_fl_grc = $Window_Main.FindName("btn_fl_grc")
$btn_fl_go = $Window_Main.FindName("btn_fl_go")
$btn_fl_print = $Window_Main.FindName("btn_fl_print")
$btn_fl_wg = $Window_Main.FindName("btn_fl_wg")
$dtg_hist = $Window_Main.FindName("dtg_hist")
$hist_scroll = $Window_Main.FindName("hist_scroll")
$btn_config = $Window_Main.FindName("btn_config")
$btn_pin = $Window_Main.FindName("btn_pin")

#Endregion

#Region ControleLoginWindows
$txtb_username = $Login_Window.FindName("txtb_username")
$pb_pwd_confirm = $Login_Window.FindName("pb_pwd_confirm")
$btn_ok = $Login_Window.FindName("btn_ok")
$lbl_conf_ko = $Login_Window.FindName("lbl_conf_ko")
$btn_cancel = $Login_Window.FindName("btn_cancel")
#EndRegion
$btn_mode = $Config_Window.FindName("btn_mode")
$cbx_primary = $Config_Window.FindName("cbx_primary")
$cbx_secondary = $Config_Window.FindName("cbx_secondary")
$btn_config_ok = $Config_Window.FindName("btn_config_ok")
$tg_config_winlock = $Config_Window.FindName("tg_config_winlock")
$tg_config_hide = $Config_Window.FindName("tg_config_hide")
$btn_close_config = $Config_Window.FindName("btn_close_config")
#Region ControlConfigWindow

$btn_ad= $Window_Main.FindName("btn_ad")
$btn_context_ad = $Window_Main.FindName("btn_context_ad")
$btn_computer = $Window_Main.FindName("btn_computer")
$btn_context_comp = $Window_Main.FindName("btn_context_comp")
$mn_computer = $Window_Main.FindName("mn_computer")
$mn_context_comp = $Window_Main.FindName("mn_context_comp")

$btn_ad.add_click({
	$btn_context_ad.isopen = $true
})

$btn_computer.add_click({
	$btn_context_comp.isopen = $true
	$btn_context_comp.style = "{StaticResource MaterialDesignContextMenu}"
})

$mn_computer.add_click({
	$mn_context_comp.style = "{StaticResource MaterialDesignContextMenu}"
	$mn_context_comp.isopen = $true
	
})

#EndRegion

#	FONCTIONS
#-------------------------------------------------------------------------


Function Check_Password {
	Param (
		[Parameter(Position=0,Mandatory=$true)]
		[string]$password
		)
	$UserName = $env:USERDOMAIN +"\" + ("$($env:USERNAME[0..2])").Replace(" ","")
	Add-Type -AssemblyName System.DirectoryServices.AccountManagement
	$contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
	$principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $contextType,$env:USERDOMAIN
	$principalContext.ValidateCredentials($UserName,$password)
}

Function Create_Credential {
	$Pwd_File = "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt"
	[Byte[]] $key = (1..16)
	$UserName = $env:USERDOMAIN +"\" + ("$($env:USERNAME[0..2])").Replace(" ","")

	$script:MyCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content $Pwd_File | ConvertTo-SecureString -Key $key)
}

Function To_cred_or_not_to_cred {

	if ((Test-Path "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt") -eq $false) { 
			Close-SplashScreen 
				$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
				$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
				$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

				# Limite l'utilisation de la RAM
				#[System.GC]::Collect()

				$Login_Window.WindowStartupLocation = "CenterScreen"	

				[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Login_Window)
				$Login_Window.Show()
				$Login_Window.Activate()	
				$appContext = New-Object System.Windows.Forms.ApplicationContext
				[void][System.Windows.Forms.Application]::Run($appContext)
	}
	Else {
		$orig = Get-Content "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt"
		[Byte[]] $key = (1..16)
		$chaine_originale = ConvertTo-SecureString -key $key -string $orig
		$decode_pwd = [system.Runtime.InteropServices.Marshal]::SecureStringToBSTR($chaine_originale)

		$pwd_check_atlaunch = Check_Password -password ([System.Runtime.InteropServices.Marshal]::PtrToStringUni($decode_pwd))
		if ($pwd_check_atlaunch -eq $true) { 
			Close-SplashScreen 
			#Affichage Window_Main
			$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
			$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
			$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

			# Limite l'utilisation de la RAM
			#[System.GC]::Collect()
			
			$Window_Main.WindowStartupLocation = "Manual"	

			[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Window_Main)
			$Window_Main.Show()
			$Window_Main.Activate()	
			$appContext = New-Object System.Windows.Forms.ApplicationContext
			[void][System.Windows.Forms.Application]::Run($appContext)
		}
		Else { 
			Close-SplashScreen 
			#Affichage Login_Window
			#Cacher fenetre powershell
			$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
			$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
			$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

			# Limite l'utilisation de la RAM
			#[System.GC]::Collect()

			$Login_Window.WindowStartupLocation = "CenterScreen"	

			[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Login_Window)

			$Login_Window.Show()
			$Login_Window.Activate()	
			$appContext = New-Object System.Windows.Forms.ApplicationContext
			[void][System.Windows.Forms.Application]::Run($appContext)
		}	
	}
}

Function Check_ModuleConfigurationManager { 
	If ([IntPtr]::Size -eq 8){$SCCMModulePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'}
	Else {$SCCMModulePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'}
		 If (-not(Test-Path -Path $SCCMModulePath)) 
		 {
			 Write-Host "Impossible de charger le module SCCM, le fichier $SCCCMModulePath est introuvable"
		 } 
		 Else 
		 {
			 If ((Get-Module ConfigurationManager) -eq $null)
			  {
				Import-Module $SCCMModulePath
			  	Write-Host "Chargement du module ConfigurationManager"
			} 
			  Else 
			  {
			   Write-Host "Module ConfigurationManager déjà chargé"
			 }
		}
	}
function Invoke_sql1 {
	param( [string]$sql,
		[System.Data.SQLClient.SQLConnection]$connection
	)
	
	$cmd = new-object System.Data.SQLClient.SQLCommand($sql,$connection)
	$ds = New-Object system.Data.DataSet
	$da = New-Object System.Data.SQLClient.SQLDataAdapter($cmd)
	$da.fill($ds) | Out-Null
	return $ds.tables[0].DefaultView
	}

Function Search {
param ( 
	[Parameter(Mandatory=$true)] [string]$trigbios
)
$query_search = "SELECT TRIGRAMME,FIRSTNAME,LASTNAME,ID_RESSOURCE,SERVICE,CP,VILLE,TELEPHONE,EMAIL,LOGIN_SIGMA,LOGIN_GRC FROM PERSONNES
INNER JOIN RESSOURCE ON RESSOURCE.SID_PERS = PERSONNES.SID_PERS
INNER JOIN SITEENTREPRISE on SITEENTREPRISE.ID_SITEENTREPRISE = PERSONNES.ID_SITEENTREPRISE
where TRIGRAMME = '$trigbios' OR ID_RESSOURCE = '$trigbios'"

$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$global:iguazuSRV;Database=$global:iguazuDB;Trusted_Connection=True;"
$con.open()

Invoke_sql1 $query_search $con

$con.Close()
}

Function Origine {
	param ( 
	[Parameter(Mandatory=$true)] [string]$adresse_ip
)
$query_ip = "select TOP 1 COMMUNE,CODEPOSTAL,dbo.SUBNET.CODEINSEE from dbo.SUBNET
INNER JOIN dbo.CODEINSEE ON dbo.CODEINSEE.CODEINSEE = dbo.SUBNET.CODEINSEE
WHERE NETWORK LIKE '$adresse_ip'"

$con = New-Object System.Data.SqlClient.SqlConnection
$con.ConnectionString="Server=$global:iguazuSRV;Database=$global:iguazuDB;Trusted_Connection=True;"
$con.open()

Invoke_sql1 $query_ip $con

$con.Close()
}

Function Age_Password {
	Param ( [Parameter(Mandatory=$true)] [string]$user_trigramme )

    $script:user = Get-ADUser $user_trigramme -Properties passwordlastset,lockedout
	$pwd_age = ((get-date) - $user.passwordlastset).Days
    return $pwd_age 
	}
	
Function Dev_account {
	Param ( [Parameter(Mandatory=$true)] [string]$user_trigramme )

	
Unlock-ADAccount -Identity $user_trigramme -Credential $MyCredential
$Lock = Get-ADUser -Filter {samAccountName -eq $user_trigramme } -Properties Lockedout 

return $lock.LockedOut
}

Function Reinit_Win_password {
	Param ( 
		[Parameter(Position=0,Mandatory=$true)] [string]$user_account,
		[Parameter(Position=1,Mandatory=$true)] [string]$new_pwd
		)

	Set-ADAccountPassword $user_account -NewPassword (ConvertTo-SecureString -AsPlainText $new_pwd -Force) -PassThru -Reset -Credential $MyCredential
	Set-ADuser -Identity $user_account -ChangePasswordAtLogon $True -Credential $MyCredential

}

Function Genesys {
	Param (
		[Parameter(Mandatory=$true)] [string]$data_genesys,
        [Parameter(Mandatory=$true)] [string]$type
	)

	$SQL_Serv_Genesys = "Server"  	
	$SQL_DB_Genesys = "GEN_GIM"     

    switch -Wildcard ($type) {
        "user" {     $SqlQuery = "USE GEN_GIM
	                    DECLARE @Trigramme varchar(50)
	                    SET @Trigramme = '$data_genesys'
	                    SELECT TOP 4
		                      GIDB_G_LOGIN_SESSION_V.ID,
		                      GIDB_GC_AGENT.EMPLOYEEID Trigramme,
		                      GIDB_GC_AGENT.FIRSTNAME Prenom,
		                      GIDB_GC_AGENT.LASTNAME Nom,
		                      GIDB_GC_PLACE.NAME Place,
		                      GIDB_GC_ENDPOINT.DN Telephone_Associe,
		                      GIDB_GC_FOLDER.NAME Localisation,
		                      GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
		                      GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
	                    FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
	                    WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_AGENT.EMPLOYEEID=@Trigramme)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC"
}
        "phone" {  $SqlQuery = "USE GEN_GIM
                    DECLARE @TelephoneAssocie varchar(50)
                    SET @TelephoneAssocie = '$data_genesys'
                    SELECT TOP 4
                            GIDB_G_LOGIN_SESSION_V.ID,
                            GIDB_GC_AGENT.EMPLOYEEID Trigramme,
                            GIDB_GC_AGENT.FIRSTNAME Prenom,
                            GIDB_GC_AGENT.LASTNAME Nom,
                            GIDB_GC_PLACE.NAME Place,
                            GIDB_GC_ENDPOINT.DN Telephone_Associe,
                            GIDB_GC_FOLDER.NAME Localisation,
                            GIDB_G_LOGIN_SESSION_V.CREATED DebutDeSession,
                            GIDB_G_LOGIN_SESSION_V.TERMINATED FinDeSession
                    FROM GEN_GIM.dbo.GIDB_G_LOGIN_SESSION_V GIDB_G_LOGIN_SESSION_V, GEN_GIM.dbo.GIDB_GC_AGENT GIDB_GC_AGENT, GEN_GIM.dbo.GIDB_GC_ENDPOINT GIDB_GC_ENDPOINT, GEN_GIM.dbo.GIDB_GC_FOLDER GIDB_GC_FOLDER, GEN_GIM.dbo.GIDB_GC_PLACE GIDB_GC_PLACE
                    WHERE GIDB_G_LOGIN_SESSION_V.AGENTID = GIDB_GC_AGENT.ID AND GIDB_G_LOGIN_SESSION_V.PLACEID = GIDB_GC_PLACE.ID AND GIDB_G_LOGIN_SESSION_V.PRIMARYDEVICEID = GIDB_GC_ENDPOINT.ID AND GIDB_GC_AGENT.FOLDERID = GIDB_GC_FOLDER.ID AND ((GIDB_GC_ENDPOINT.DN=@TelephoneAssocie)) ORDER BY GIDB_G_LOGIN_SESSION_V.ID DESC"  
 }
        Default { }
    }    	
	
	

	
	
	$con = New-Object System.Data.SqlClient.SqlConnection
	$con.ConnectionString="Server=$SQL_Serv_Genesys;Database=$SQL_DB_Genesys;Trusted_Connection=True;"
	$con.open()
	
	Invoke_sql1 -sql $SqlQuery -connection $con
	$con.Close()
}	
Function Remote {
	Param ( [Parameter(Mandatory=$true)] [string]$computer )
	If ([IntPtr]::Size -eq 8){ Start-Process "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$computer }
	Else { Start-Process "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"$computer }
}

Function MSRA {
	Param ( [Parameter(Mandatory=$true)] [string]$computer )
	 msra.exe /offerra $computer
	 }

Function Launch_GRC {
Start-Process "URL"
}

Function Computer_Connected {
    param(
    [String]$computer,
    [int]$delay = 100
    )

	try {
	$script:Comput_AD = Get-ADComputer $computer -ErrorAction Stop
	$script:connected = Test-Connection -ComputerName $computer -Count 1 -ErrorAction Stop
	$computer_status = "OK"

	}
	catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] { Write-Host "poste n'existe pas"
																			   $computer_status = "AD" }
	catch [System.Net.NetworkInformation.PingException] { Write-Host "poste n'est pas en ligne !!!" 
														  $computer_status = "PG"}
	catch { Write-Host "erreur inconnue au bataillon"
			$computer_status = "KO" }
	finally { $computer_status }
}
Function Search_Computer {
	Param ( [Parameter(Mandatory=$true)] [string]$computer )

	$InfoSys_option = New-CimSessionoption -Protocol Dcom 
	$InfoSys_session = New-CimSession -ComputerName $computer -SessionOption $InfoSys_option 

	$script:InfoSys = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $InfoSys_session
	$script:InfoOS = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $InfoSys_session

}

Function Rech_Titulaire_Poste {
	Param ( [Parameter(Mandatory=$true)] [string]$computer )
    
    try {
	$sccm_drive = New-PSDrive -Name P29 -PSProvider CMSite -Root "server" -ErrorAction SilentlyContinue
	Set-Location 'P29:'
	$titulaire = Get-CMUserDeviceAffinity -DeviceName $computer
	Set-Location $Current_Folder
    Remove-PSDrive -Name P29

    $SCCM_Titul = $titulaire.UniqueUserName.split("\")[1]
    return $SCCM_Titul
    } catch { $SCCM_Titul = "Error" }

	}
Function Purge_label_infos {
	$lbl_name.Text = "-"
	$lbl_trig.Text = "-"
	$lbl_idsigma.Text = "-"
	$lbl_idgrc.Text = "-"
	#$cbx_netbios.ItemsSource = $null
	$lbl_service.Text = "-"
	$lbl_ville.Text = "-"
	$lbl_phone.Text = "-"
	$lbl_mail.Text = "-"
}
Function Purge_label_Comput {
	$lbl_comp_ip.Text = "-"
	$lbl_comp_modele.Text = "-"
	$lbl_comp_os.Text = "-"
	$lbl_comp_lastboot.Text = "-"
	$lbl_comp_sccm_tit.Text = "-"
	$lbl_comp_localisation.Text = "-"
	$lbl_comp_STA_ip.Text = "-"
	$lbl_comp_STA_state.Text = "-"
}

Function Add_History {
	Param (
		[Parameter(Mandatory=$true)] [string]$username,
		[Parameter(Mandatory=$true)] [string]$surname,
		[Parameter(Mandatory=$true)] [string]$name,
		[Parameter(Mandatory=$true)] $computer
	)
	$hour = Get-Date -Format "HH:mm"
	$searched = New-Object psobject
		Add-Member -InputObject $searched -MemberType NoteProperty -Name Heure -Value $hour
		Add-Member -InputObject $searched -MemberType NoteProperty -Name trigramme -Value $username
		Add-Member -InputObject $searched -MemberType NoteProperty -Name poste -Value $computer
		Add-Member -InputObject $searched -MemberType NoteProperty -Name prenom -Value $name
		Add-Member -InputObject $searched -MemberType NoteProperty -Name nom -Value $surname

	if ($global:hist_check -eq 0) { $dtg_hist.Text = "$($searched.heure) - $($searched.trigramme) - $($searched.poste) - $($searched.prenom) $($searched.nom)  `n" }
	else { $dtg_hist.Text += "$($searched.heure) - $($searched.trigramme) - $($searched.poste) - $($searched.prenom) $($searched.nom)  `n" }

	$global:hist_check++
	$hist_scroll.LineDown()
}

Function Config_theme {
 if (Test-Path "$Current_Folder\Resources\theme.txt" ) {
	 $Current_Folder = 'D:\ut\user\Projet_WPF\BAO_V2'
	 $to_apply = Get-Content "$Current_Folder\Resources\theme.txt"
	 $theme = $to_apply.Split(";")[0]
	 $primary = $to_apply.Split(";")[1]
	 $secondary = $to_apply.Split(";")[2]


	$theme_one = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Window_Main.Resources)
	   [MaterialDesignThemes.Wpf.ThemeExtensions]::SetBaseTheme($theme_one, [MaterialDesignThemes.Wpf.Theme]::$theme)
	   [MaterialDesignThemes.Wpf.ThemeExtensions]::SetPrimaryColor($theme_one, $primary)
	   [MaterialDesignThemes.Wpf.ThemeExtensions]::SetSecondaryColor($theme_one, $secondary)
	   [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Window_Main.Resources, $theme_one)

 }

}

#	VARIABLES
#-------------------------------------------------------------------------
[String] $global:iguazuDB		= "database"
[String] $global:iguazuSRV		= "server"
[int] $global:hist_check		= 0
[int] $global:pin_activated		= 0

## LANCEMENT FONCTIONS PRE-OUVERTURE
Create_Credential
Check_ModuleConfigurationManager
Config_theme



#	EVENEMENTS
#-------------------------------------------------------------------------
$btn_search.add_Click({
	Purge_label_infos
	Purge_label_Comput
	$cbx_netbios.ItemsSource = $null
	$cbx_comp_computname.ItemsSource = $null
	if ($lbl_confirm_reinit.Visibility -eq "Visible") {$lbl_confirm_reinit.Visibility = "Hidden" }
	$search_data = $txt_search.get_text()
	$script:result = Search -trigbios $search_data
		$result_unique = $script:result | Select-Object -Unique


	$lbl_name.Text = $result_unique.FIRSTNAME + " " + $result_unique.LASTNAME
	$lbl_trig.Text = $result_unique.TRIGRAMME
	$lbl_idsigma.Text = $result_unique.LOGIN_SIGMA
	$lbl_idgrc.Text = $result_unique.LOGIN_GRC

	if ($result.ID_RESSOURCE.Count -eq 1) {
		$cbx_netbios.Items.Add($result_unique.ID_RESSOURCE)
		$cbx_netbios.SelectedItem = $result_unique.ID_RESSOURCE
		
	}
	else {
		$cbx_netbios.ItemsSource = $result.ID_RESSOURCE
		$cbx_netbios.SelectedItem = $result.ID_RESSOURCE[$result.ID_RESSOURCE.Count -1]
		
	}

	$lbl_service.Text = $result_unique.SERVICE
	$lbl_ville.Text = $result_unique.CP + " " + $result_unique.VILLE
	
	$genesys_data = Genesys -data_genesys $result_unique.TRIGRAMME -type user
		if ($genesys_data -eq $null) { $lbl_phone.Text = $result_unique.TELEPHONE }
		else { $lbl_phone.Text = $genesys_data.Place[0]
			$txt_search_genesys_num.Text = $genesys_data.Place[0]
			$txt_search_genesys.Text = $result_unique.TRIGRAMME }
	
	$lbl_mail.Text = $result_unique.EMAIL

	$txtb_pwd_trig.Text = $result_unique.TRIGRAMME
	if ($result.ID_RESSOURCE.Count -eq 1) {
		$cbx_comp_computname.Items.Add($result_unique.ID_RESSOURCE)
		$cbx_comp_computname.SelectedItem = $result_unique.ID_RESSOURCE
	}
	else {
	$cbx_comp_computname.ItemsSource = $result.ID_RESSOURCE
		$cbx_comp_computname.SelectedItem = $result.ID_RESSOURCE[$result.ID_RESSOURCE.Count -1]
	}

	$lbl_pwd_age.Content = Age_Password -user_trigramme ($txtb_pwd_trig.get_text())
		if ($user.LockedOut -eq $false) { $lbl_pwd_state.Content = "Ok"
											$btn_pwd_unlock.IsEnabled = $false
											$mn_unlock.IsEnabled = $false }
		Else { $lbl_pwd_state.Content = "Verrouillé"
				$btn_pwd_unlock.IsEnabled = $true
				$mn_unlock.IsEnabled = $true }

		$btn_com_remote_ip.IsEnabled = $false
		$mn_comp_sta_pdm.IsEnabled = $false
		$mn_comp_sta_reboot.IsEnabled = $false
		$mn_comp_sta_num.IsEnabled = $false

		Add_History -computer $result.ID_RESSOURCE -username  $result_unique.TRIGRAMME -name $result_unique.LASTNAME -surname $result_unique.FIRSTNAME 
	
})

$btn_pwd_trig.add_Click({
	$lbl_pwd_age.Content = Age_Password -user_trigramme ($txtb_pwd_trig.get_text())
	if ($user.LockedOut -eq $false) { $lbl_pwd_state.Content = "Ok"
											$btn_pwd_unlock.IsEnabled = $false
											 }
		Else { $lbl_pwd_state.Content = "Verrouillé"
				$btn_pwd_unlock.IsEnabled = $true  
				}


 })

$btn_flyout.add_Click({ 
	$flyout.isopen = $true 
})

$mn_unlock.add_Click({ 
	$Unlock = Dev_account -user_trigramme $lbl_trig.get_text()

	if ($Unlock -eq $false) { $lbl_pwd_state.Content = "Compte déverrouillé !"
								$btn_pwd_unlock.IsEnabled = $false
								$mn_unlock.IsEnabled = $false }
	Else { $lbl_pwd_state.Content = "Echec du déverrouillage !"
			$btn_pwd_unlock.IsEnabled = $true
			$mn_unlock.IsEnabled = $true }
})

$btn_pwd_unlock.add_Click({ 
	$Unlock = Dev_account -user_trigramme $txtb_pwd_trig.get_text()

	if ($Unlock -eq $false) { $lbl_pwd_state.Content = "Compte déverrouillé !"
								$btn_pwd_unlock.IsEnabled = $false
								$mn_unlock.IsEnabled = $false }
	Else { $lbl_pwd_state.Content = "Echec du déverrouillage !"
			$btn_pwd_unlock.IsEnabled = $true
			$mn_unlock.IsEnabled = $true }
})

$mn_reinit.add_Click({ 
	$exp_ad.isExpanded = $true
})

$mn_remote.add_Click({
	Remote -computer $cbx_netbios.SelectedItem
 })

$mn_msra.add_Click({ 
	MSRA -computer $cbx_netbios.SelectedItem
})

$mn_exp_comp.add_Click({
	$exp_computer.IsExpanded = $true
})

$btn_reinit_grc.add_Click({
	Launch_GRC
})

$btn_reinit_sigma.add_Click({
$okOnly = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::Affirmative
[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Window_Main,"Information","Fonctionnalité non implémentées pour l'instant...",$okOnly)
})

$btn_mail.add_Click({
	$create_mail = New-Object -comObject Outlook.Application
	$new_mail = $create_mail.CreateItem(0)
	$new_mail.Subject = "[ToolBox]"
	$new_mail.To = $lbl_mail.get_text()
	$inspector = $new_mail.GetInspector
	$inspector.Activate()	
})

$btn_pwd_reinit.add_Click({
	$ok_cancel = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::AffirmativeAndNegative
	$return =[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Window_Main,"Validation","Réinitialisation du mot de passe de $($result_unique.TRIGRAMME) ?",$ok_cancel)
	#$return | Out-File C:\temp\test.txt

	try {
		if ($return -eq "Affirmative"){
		Reinit_Win_password -user_account $result_unique.TRIGRAMME -new_pwd $txtb_pwd_password.get_text()
		$lbl_confirm_reinit.Visibility = "Visible"
		}
		else {}
	} catch { $lbl_confirm_reinit.Content = "ERROR"
		$lbl_confirm_reinit.Visibility = "Visible" }
})

$btn_pwd_console.add_Click({
	Start-Process “C:\Windows\System32\cmd.exe” -workingdirectory $PSHOME -ArgumentList “/c dsa.msc” -Credential $MyCredential -NoNewWindow
})

$btn_comp_comput.add_Click({ 
	Purge_label_Comput
	$pg_ring_comput.IsActive = $true
		$mn_comp_sta_pdm.IsEnabled = $false
		$mn_comp_sta_reboot.IsEnabled = $false
		$mn_comp_sta_num.IsEnabled = $false
	if ($cbx_comp_computname.SelectedItem -eq $null	) { $computer_comp = $cbx_comp_computname.Text }
	else { $computer_comp = $cbx_comp_computname.SelectedItem }
	$result_comput = Computer_Connected -computer $computer_comp
	$lbl_comp_netbios.Content = $computer_comp.ToUpper() +" :"
	switch -Wildcard ($result_comput) {
		"OK" { Search_Computer	-computer $computer_comp
				$lbl_comp_ip.Text = $connected.IPV4Address.IPAddressToString
				$lbl_comp_modele.Text = $InfoSys.Model
				$lbl_comp_os.Text = $InfoOS.Caption
				$lbl_comp_lastboot.Text = (get-date ($InfoOS.LastBootUpTime) -Format 'dd/MM/yyyy hh:mm:ss' )
				$lbl_comp_sccm_tit.Text = Rech_Titulaire_Poste -computer $computer_comp
				$btn_com_remote_ip.IsEnabled = $true

				switch -wildcard ($connected.IPV4Address.IPAddressToString) {
					"10.1*" { $recup_ip = $connected.IPV4Address.IPAddressToString.split(".")
								$ip_ag = "$($recup_ip[0]).$($recup_ip[1]).$($recup_ip[2])"
								$situation = Origine -adresse_ip "$($ip_ag).%"
								$lbl_comp_localisation.Text = $situation.CODEPOSTAL +" "+ $situation.COMMUNE
								}
					"10.227.*" { $lbl_comp_localisation.Text = "Télétravail"
									 }
					
					"10.208.*" {  $lbl_comp_localisation.Text = "Télétravail"
									}

					Default {$recup_ip = $connected.IPV4Address.IPAddressToString.split(".")
								$ip_ag = "$($recup_ip[0]).$($recup_ip[1]).$($recup_ip[2])"
								$situation = Origine -adresse_ip "$($ip_ag).%"
								$lbl_comp_localisation.Text = $situation.CODEPOSTAL +" "+ $situation.COMMUNE
								$lbl_comp_STA_ip.Text = $ip_ag + ".70"
								$mn_comp_sta_pdm.IsEnabled = $true
								$mn_comp_sta_reboot.IsEnabled = $true
								$mn_comp_sta_num.IsEnabled = $true
								if (Test-Connection "$($ip_ag).70" -Count 1) { $lbl_comp_STA_state.Text = "En ligne"}
								else { $lbl_comp_STA_state.Text = "Hors ligne"}
							}
				}
				
			}
		"AD" { $lbl_comp_ip.Text = "Poste Inconnu"}
		"PG" {  $lbl_comp_ip.Text = "Poste Hors Ligne"}
		"KO" {  $lbl_comp_ip.Text = "Erreur Inconnue"}
	}
})

$btn_com_remote_ip.add_Click({
	Remote -computer $lbl_comp_ip.Text
})

$mn_comp_sta_pdm.add_Click({ Remote -computer $lbl_comp_STA_ip.Text })
$mn_comp_sta_reboot.add_Click({ Restart-Computer -ComputerName $lbl_comp_STA_ip.Text -Force })
$mn_comp_sta_num.add_Click({ Start-Process explorer.exe "\\$($lbl_comp_STA_ip.Text)\c$\sasloc\numerisation" })

$btn_search_genesys_num.add_Click({
	$genesys_data = Genesys -data_genesys $txt_search_genesys_num.get_text() -type phone
	if ($genesys_data -eq $null) { $lbl_search_genesys_num.Content = "Pas Genesys"
						$dtg_hist_num.ItemsSource = $null }
	else { $lbl_search_genesys_num.Content= $genesys_data.Trigramme[0]
		
		$array_genesys_data = New-Object System.Collections.ArrayList
		$array_genesys_data.AddRange( ($genesys_data | Select-Object @{n='Num';e={$_.Telephone_Associe}},Trigramme,Nom,Prenom,@{n='LogIn';e={Get-date ($_.DebutDeSession) -format 'dd/MM/yy HH:mm'}},@{n='LogOut';e={Get-date ($_.FinDeSession) -format 'dd/MM/yy HH:mm'}}) ) 
		$dtg_hist_num.ItemsSource = $array_genesys_data
 }

})

$btn_search_genesys.add_Click({
	$genesys_data = Genesys -data_genesys $txt_search_genesys.get_text() -type user
	if ($genesys_data -eq $null) { $lbl_search_genesys.Content = "Pas Genesys" 
						$dtg_hist_ut.ItemsSource = $null }
	else { $lbl_search_genesys.Content= $genesys_data.Place[0]
		$dtg_hist_ut.ItemsSource = $genesys_data

		$array_genesys_data = New-Object System.Collections.ArrayList
		$array_genesys_data.AddRange( ($genesys_data | Select-Object Trigramme,@{n='Num';e={$_.Telephone_Associe}},Nom,Prenom,@{n='LogIn';e={Get-date ($_.DebutDeSession) -format 'dd/MM/yy HH:mm'}},@{n='LogOut';e={Get-date ($_.FinDeSession) -format 'dd/MM/yy HH:mm'}}) ) 
		$dtg_hist_ut.ItemsSource = $array_genesys_data

 }
})

$btn_config.add_Click({
	$Config_Window.Show()
	$Config_Window.Activate()	

	$flyout.isopen = $false
})

$btn_pin.add_Click({
	$global:pin_activated++
	if ($btn_pin.Content.Kind -eq 'PinOff') { $btn_pin.Content.Kind = 'Pin' }
	else { $btn_pin.Content.Kind = 'PinOff' }

	$flyout.isopen = $false
})

$btn_mode.add_Click({
	$theme = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Window_Main.Resources)
   
    if ($btn_mode.IsChecked -eq $true) {
      [MaterialDesignThemes.Wpf.ThemeExtensions]::SetBaseTheme($theme, [MaterialDesignThemes.Wpf.Theme]::Light)
    }
    if ($btn_mode.IsChecked -eq $False) {
      [MaterialDesignThemes.Wpf.ThemeExtensions]::SetBaseTheme($theme, [MaterialDesignThemes.Wpf.Theme]::Dark)
    }
    [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Window_Main.Resources, $theme)
})

$cbx_primary.Add_SelectionChanged({
	$theme = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Window_Main.Resources)
    $PrimaryColors = [MaterialDesignColors.SwatchHelper]::Lookup[$cbx_primary.SelectedValue]
    [MaterialDesignThemes.Wpf.ThemeExtensions]::SetPrimaryColor($theme, $PrimaryColors)
	[MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Window_Main.Resources, $theme)
})

$cbx_secondary.Add_SelectionChanged({
	$theme = [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::GetTheme($Window_Main.Resources)
    $SecondaryColors = [MaterialDesignColors.SwatchHelper]::Lookup[$cbx_secondary.SelectedValue]
    [MaterialDesignThemes.Wpf.ThemeExtensions]::SetSecondaryColor($theme, $SecondaryColors)
    [MaterialDesignThemes.Wpf.ResourceDictionaryExtensions]::SetTheme($Window_Main.Resources, $theme)
})

$btn_config_ok.add_click({
	if ($btn_mode.IsChecked -eq $true) {$mode = "Light"}
	if ($btn_mode.IsChecked -eq $false) {$mode = "Dark"}
	$primary = [MaterialDesignColors.SwatchHelper]::Lookup[$cbx_primary.SelectedItem]
	$secondary = [MaterialDesignColors.SwatchHelper]::Lookup[$cbx_secondary.SelectedItem]

	"$($mode);$($primary);$($secondary)" | Out-File "$Current_Folder\Resources\theme.txt"

	$Config_Window.Hide()
})

$btn_close_config.add_Click({ 
	$Config_Window.Hide()
 })

$btn_fl_wiki.Add_Click({  
	Start-Process "http://"
	$flyout.isopen = $false 
})
$btn_fl_pdsi.Add_Click({ 
	Start-Process "http://"
	$flyout.isopen = $false 
})
$btn_fl_po.Add_Click({ 
	Start-Process "http://"
	$flyout.isopen = $false 
})
$btn_fl_ig.Add_Click({
	Start-Process "http://"
	$flyout.isopen = $false
})
$btn_fl_sp.Add_Click({ 
	Start-Process "http://" 
	$flyout.isopen = $false
})
$btn_fl_smart.Add_Click({ 
	Start-Process "https://"
	$flyout.isopen = $false 
})
$btn_fl_grc.Add_Click({ 
	Start-Process "https://"
	$flyout.isopen = $false 
})
$btn_fl_go.Add_Click({ 
	Start-Process "http://"
	$flyout.isopen = $false 
})
$btn_fl_print.Add_Click({
	Start-Process "https://"
	$flyout.isopen = $false
})
$btn_fl_wg.Add_Click({ 
	Start-Process "https://"
	$flyout.isopen = $false
})

#########################################################
### WINDOW LOGIN

$btn_ok.add_click({
	$chk_pwd = Check_Password -password $pb_pwd_confirm.get_password()
	if ($chk_pwd -eq $false){ $lbl_conf_ko.content = "Mot de passe erroné"
								$lbl_conf_ko.Visibility = "Visible"}
	else {

		$Pwd_File = "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt"
		[Byte[]] $key = (1..16)
		$Password = $pb_pwd_confirm.get_password() | ConvertTo-SecureString -AsPlainText -Force
		$Password | ConvertFrom-SecureString -key $key | Out-File $Pwd_File


		$Window_Main.WindowStartupLocation = "Manual"	
		[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Window_Main)
		
		$Login_Window.Close()

		$Window_Main.Show()
		$Window_Main.Activate()	
	}
})

$btn_cancel.add_Click({ 
	$Main_Tool_Icon.Visible = $false
	$Login_Window.Close()
	$Window_Main.Close()
	Stop-Process $pid
})

############################################################


# ----------------------------------------------------
# MENU SYSTRAY
# ----------------------------------------------------		
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "Toolbox 3.0"
$Main_Tool_Icon.Icon = $icon
$Main_Tool_Icon.Visible = $true

$Menu_Users = New-Object System.Windows.Forms.MenuItem
$Menu_Users.Text = "Open Toolbox"

$Menu_Config = New-Object System.Windows.Forms.MenuItem
$Menu_Config.Text = "Configuration"

$Menu_Reboot = New-Object System.Windows.Forms.MenuItem
$Menu_Reboot.Text = "Reboot Toolbox"

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"


$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Users)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Config)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Reboot)
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)


# ---------------------------------------------------------------------
# ACTION SYSTRAY
# ---------------------------------------------------------------------
$Main_Tool_Icon.Add_Click({
        $Window_Main.WindowStartupLocation = "Manual"	
		[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Window_Main)
		$Window_Main.Show()
		$Window_Main.Activate()	
			
})
$Menu_Users.Add_Click({
    $Window_Main.WindowStartupLocation = "Manual"
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Window_Main)
	$Window_Main.Show()
	$Window_Main.Activate()	
})

$Menu_Config.add_Click({
	$Config_Window.WindowStartupLocation = "Manual"
	[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Config_Window)
	$Config_Window.Show()
	$Config_Window.Activate()	
})

$Menu_Reboot.add_Click({
$Restart = "Yes"
 start-process -WindowStyle hidden powershell.exe ".\bao_V2.ps1 '$Restart'"  
 
 $Main_Tool_Icon.Visible = $false
 $Window_Main.Close()
 Stop-Process $pid
  
 $Global:Timer_Status = $timer.Enabled
 If ($Timer_Status -eq $true)
  {
   $timer.Stop() 
  }  
})

$Window_Main.Add_MouseDoubleClick({
	 #$Window_Main.Hide()
 })

$Window_Main.Add_Deactivated({
	if (($global:pin_activated % 2)	-ne 0 ) { $Window_Main.Hide() }
		
})

$Menu_Exit.add_Click({
	$Main_Tool_Icon.Visible = $false
	$Window_Main.Close()
	Stop-Process $pid
 })

$Window_Main.add_Loaded({	

})

$Config_Window.add_Loaded({
	$cbx_primary.ItemsSource = [System.Enum]::GetNames([MaterialDesignColors.PrimaryColor])
	$cbx_secondary.ItemsSource = [System.Enum]::GetNames([MaterialDesignColors.SecondaryColor])

	if (($global:pin_activated % 2)	-ne 0 ) { $tg_config_hide.IsChecked =  $true }
	else {$tg_config_hide.IsChecked =  $false }

})

$title_bar.add_MouseDown({
	$Window_Main.DragMove()
})

$Login_Window.add_Loaded({
	$txtb_username.Text =  $env:USERDOMAIN +"\" + ("$($env:USERNAME[0..2])").Replace(" ","")
})

$Window_Main.add_ContentRendered({	
})

$Window_Main.add_Closing({
	$_.Cancel = $true
	$Window_Main.Hide()
})


To_cred_or_not_to_cred

