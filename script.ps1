# Ask for credentials that have access to the servers
$cred = Get-Credential

# Loop over the servers (1 FQDN or IP by line)
foreach($serv in Get-Content .\ServCert.txt) {
    write-host "#################### `r"
    # Creation of a mapped drive J leading to destination C disk = \\destination\c$
    $desiredMappedDrive = 'J'
    $desiredMappedDrivePath = "\\"+$serv+"\c$" 

    # Source path of the PFX
    $source = "C:\PATH\PFX.PFX"
    # Destination path of the PFX
    $destination = "${desiredMappedDrive}:\PATH\" 
    
    # Test destination path, if it doesn't exist, it will be created
    Invoke-Command -ComputerName $serv -Credential $cred -ScriptBlock {if (-Not (Test-Path "C:\PATH")) {New-Item "C:\PATH" -itemType Directory}}

    # Creation of the mapped drive
    New-PSDrive -Name $desiredMappedDrive -PSProvider FileSystem -Root $desiredMappedDrivePath -Credential $cred

    # Copy of the file
    Copy-Item -Path $source -Destination $destination -Verbose

    # Removing the mapped drive
    Remove-PSDrive J


    # invoke commands that will be executed on the destination server to install the PFX
    Invoke-Command -ComputerName $serv -Credential $cred -ScriptBlock {

        # Installation of the Pfx
        Import-PfxCertificate -FilePath C:\PATH\PFX.pfx -CertStoreLocation "Cert:\LocalMachine\My" -Password (ConvertTo-SecureString -String 'PASSWORD-OF-THE-PFX-FILE'-AsPlainText -Force)

        # Edit of the bindings

	# Import of the module
        Import-Module WebAdministration

        # Loop over the IIS bindings to match the thumbprint of the old certificate
        # When it does, it will loop over the sites to switch the certificate by the newly installed one
        $siteThumbs = Get-ChildItem IIS:SSLBindings | Foreach-Object {
            $thumb = $_.Thumbprint 
            $refthumb = "THUMBPRINT-OF-THE-OLD-CERTIFICATE"
            if($thumb -eq $refthumb){
            # Loop over each sites
                write-host "#################### `r"
                foreach ($site in $_.Sites.Value) {
                    # Certificate switch
                    (Get-WebBinding -Name $site -protocol "https").AddSslCertificate("THUMBPRINT-OF-THE-NEW-CERTIFICATE", "my")
                    write-host "Certificat changed for this site : "$site
                }
                write-host "#################### `r"
            }

    }
}
# Pause to see the result, press any key to end
read-host
