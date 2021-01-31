class Department {
	[int]$dep
	[string]$nas_path
	[string]$groupe_ad
}

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

function departement {
	Param(
		[Parameter(Position=0,Mandatory=$true)]
		[int]$dept
	)

	$dep = New-Object Department

	switch ($dept) {
		22 { $nas_path = "\\nas22\save_pdt" 
				$groupe_ad = "GLB_ARC1001443" }
		29 { $nas_path = "\\nasri29\save_pdt"
				$groupe_ad = "GLB_ARC1001454" }
		35 { $nas_path = "\\naspdt35\save_pdt"
				$groupe_ad = "GLB_ARC1001456" }
		44 { $nas_path = "\\nas44\save_pdt" 
				$groupe_ad = "GLB_ARC1001459"}
		49 { $nas_path = "\\nas49\save_pdt"
				$groupe_ad = "GLB_ARC1001458" }
		56 { $nas_path = "\\nas56\save_pdt"
				$groupe_ad = "GLB_ARC1001457" }
		Default { Write-Host "merci de renseigner un numéro de département valide"}
	}
	

	$dep.dep = $dept
	$dep.groupe_ad = $groupe_ad
	$dep.nas_path = $nas_path
	return $dep
}

$get_dep = Read-Host("Merci de saisir le département (22/29/35/44/49/56) ")
$nas_dep = departement($get_dep)


$save_byname = Get-ChildItem -Path $nas_dep.nas_path | Select-Object Name
$list_user_KO = @()
$list_user_disable = @()
$list_user_OK = @()
$list_no_save = @()



foreach ($name in $save_byname){

    try { $test_user = Get-ADUser $name.Name -Properties SamAccountName,Name,enabled,whenchanged,PostalCode -ErrorAction SilentlyContinue

				}
    catch {
        if ($error[0].CategoryInfo.Reason -eq "ADIdentityNotFoundException") 
            {Write-Warning "$($error[0].CategoryInfo.TargetName) non trouvé dans l'AD" 
                $no_count = [ordered]@{
                Name = "$($name.name)"
                Etat = "Compte non trouvé dans l'AD"
                Action = "A Supprimer"
                }
             $result = New-Object psobject -Property $no_count
             $list_user_KO += $result }
          }

if ($test_user.SamAccountName -eq $name.name) {
	 $dir_save = Get-ChildItem "$($nas_dep.nas_path)\$($Name.Name)" | Select-Object Name,CreationTime,LastWriteTime,Length
			if ($null -eq $dir_save) {
				$ok_count = [ordered]@{
							Dep = "$($test_user.PostalCode)"
		                    Name = "$($name.name)"
		                    Nom_Complet = "$($test_user.Name)"
		                    Etat = ""
		                    Dossier = ""
							Taille = ""
		                    }
				    $result = New-Object psobject -Property $ok_count
					if ($test_user.Enabled -eq $false) { $list_user_disable += $result }
					else {	$list_user_OK += $result }
		            Write-Host "$($test_user.SamAccountName)_AD - $($name.name)_NAS "}
	
			else {
                foreach ($dir in $dir_save){
						$ok_count = [ordered]@{
								Dep = "$($test_user.PostalCode)"
			                    Name = "$($name.name)"
			                    Nom_Complet = "$($test_user.Name)"
			                    Etat = ""
			                    Dossier = "$($dir.Name) - Créé le $($Dir.CreationTime) - Modifié le $($dir.LastWriteTime)"
								Taille = "$($dir.Lenght)"
			                    }

					if ($test_user.Enabled -eq $false) { 
													$ok_count.Etat = "Compte désactivé"
													$result = New-Object psobject -Property $ok_count
													$list_user_disable += $result }
					else {	
						$ok_count.Etat = "Compte Actif"
						$result = New-Object psobject -Property $ok_count
						$list_user_OK += $result }
		           		Write-Host "$($test_user.SamAccountName)_AD - $($name.name)_NAS - $($result.Taille)"}
					}
	
	
}
}

$resume = New-Object psobject -Property @{
		Comptes_Actifs = "$($list_user_OK.count)"
		Comptes_Desactives = "$($list_user_disable.count)"
		Comptes_Inexistants = "$($list_user_KO.count)"
		
}

$SITE_DEP = Get-ADGroupMember $nas_dep.groupe_ad | Select-Object SamAccountName,Name
$comp = Compare-Object -ReferenceObject $SITE_DEP.SamAccountName -DifferenceObject $save_byname.name | Where-Object {$_.sideindicator -eq "<="}

		
		foreach ($name_no_save in $comp) {
		$name_AD = Get-ADUser $name_no_save.InputObject -Properties Name
		$no_save = [ordered]@{
			Comptes = ""
			Nom = ""
			Sauvegarde = "Aucune sauvegarde trouvée"
			}
		
		$no_save.Comptes = $name_no_save.InputObject
		$no_save.Nom = $name_AD.Name
		
		$result = New-Object psobject -Property $no_save
		
		$list_no_save += $result
		}


$frag01 = $list_no_save |  ConvertTo-Html -Fragment -As table -PreContent "<h2>Point d'attention !</h2>" | Out-String
$frag0 = $resume | ConvertTo-Html -Fragment -As Table -PreContent "<h2>En bref</h2>" | Out-String
$frag1 = $list_user_KO | Sort-Object Name | ConvertTo-Html -Fragment -as Table -PreContent "<h2>Liste des comptes KO</h2>" | Out-String
$frag2 = $list_user_disable | Sort-Object Name | ConvertTo-Html -Fragment -as Table -PreContent "<h2>Liste des comptes Disable</h2>" | Out-String
$frag3 = $list_user_OK | Sort-Object Name | ConvertTo-Html -Fragment -as Table -PreContent "<h2>Liste des comptes OK</h2>" | Out-String

$style = Get-StyleSheet

ConvertTo-HTML -Title "Rapport NAS $($nas_dep.dep)" `
                -Head $style `
                -Body "<h1>Rapport NAS $($nas_dep.dep)</h1>",$frag01,$frag0,$frag1,$frag2,$frag3 |
Out-File C:\temp\rapport.html

Start-Process "C:\temp\rapport.html"