#
#	INIT
#-------------------------------------------------------------------------
$Global:Current_Folder = 'D:\ut\MBG17590\Projet_WPF\BAO_V2' #$PSScriptRoot

Set-Location $Current_Folder
$Assembly_Folder = Join-Path -Path $Current_Folder -ChildPath .\assembly



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
	Foreach ($Assembly in (Dir $Assembly_Folder -Filter *.dll)) {
		$null = [System.Reflection.Assembly]::LoadFrom($Assembly.fullName)
	}
}


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
$xaml_login = LoadXaml("$Current_Folder\Page1.xaml") 
#$xaml = LoadXaml("$Current_Folder\Page1.xaml") 

$Login_xamlReader = (New-Object System.Xml.XmlNodeReader $xaml_login)
$Login_Window = [Windows.Markup.XamlReader]::Load($Login_xamlReader)


#SplashScreen
$splash_xaml = LoadXaml("$Current_Folder\Page2.xaml") 
$xamlReader_splash = (New-Object System.Xml.XmlNodeReader $splash_xaml)
$Splash_Window = [Windows.Markup.XamlReader]::Load($xamlReader_splash)

#Region ControleLoginWindows
$txtb_username = $Login_Window.FindName("txtb_username")
$pb_pwd_confirm = $Login_Window.FindName("pb_pwd_confirm")
$btn_ok = $Login_Window.FindName("btn_ok")
$lbl_conf_ko = $Login_Window.FindName("lbl_conf_ko")
$btn_cancel = $Login_Window.FindName("btn_cancel")
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

Function To_cred_or_not_to_cred {

	if ((Test-Path "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt") -eq $false) { 
				$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
				$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
				$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

				# Limite l'utilisation de la RAM
				[System.GC]::Collect()

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
			
			$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
			$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
			$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

			# Limite l'utilisation de la RAM
			[System.GC]::Collect()

			$Splash_Window.WindowStartupLocation = "CenterScreen"	

			[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Splash_Window)

			$Splash_Window.Show()
			$Splash_Window.Activate()	
			$appContext = New-Object System.Windows.Forms.ApplicationContext
			[void][System.Windows.Forms.Application]::Run($appContext)

			powershell -file "$($Current_Folder)\bao_v2.ps1"
			while (!(Test-Path C:\temp\launch_bao.txt)){ Start-Sleep -Seconds 2 }
			
		}
		Else { 
			#Affichage Login_Window
			#Cacher fenetre powershell
			$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
			$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
			$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

			# Limite l'utilisation de la RAM
			[System.GC]::Collect()

			$Login_Window.WindowStartupLocation = "CenterScreen"	

			[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Login_Window)

			$Login_Window.Show()
			$Login_Window.Activate()	
			$appContext = New-Object System.Windows.Forms.ApplicationContext
			[void][System.Windows.Forms.Application]::Run($appContext)
		}

	}
}

#	EVENEMENTS
#-------------------------------------------------------------------------
$btn_ok.add_click({
	$chk_pwd = Check_Password -password $pb_pwd_confirm.get_password()
	if ($chk_pwd -eq $false){ $lbl_conf_ko.content = "Mot de passe erroné"
								$lbl_conf_ko.Visibility = "Visible"}
	else {

		$Pwd_File = "$($Current_Folder)\Encrypted_pwd\Encrypted_pwd.txt"
		[Byte[]] $key = (1..16)
		$Password = $pb_pwd_confirm.get_password() | ConvertTo-SecureString -AsPlainText -Force
		$Password | ConvertFrom-SecureString -key $key | Out-File $Pwd_File

		
	}
})

$btn_cancel.add_Click({ 
	$Main_Tool_Icon.Visible = $false
	$Login_Window.Close()
	$Window_Main.Close()
	Stop-Process $pid
})

############################################################


To_cred_or_not_to_cred