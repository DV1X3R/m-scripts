# Tools: https://github.com/Microsoft/ReportingServicesTools
# Guide: https://www.mssqltips.com/sqlservertip/4738/powershell-commands-for-sql-server-reporting-services/
# Install Tools command: Invoke-Expression (Invoke-WebRequest https://aka.ms/rstools)

$a = new-object -comobject wscript.shell 

Function Select-Server {
    Write-Output '1. [Enter URL]'
    $select = Read-Host -Prompt 'Select'
    switch($select) {
        1 { $global:server = Read-Host -Prompt 'Input server address' }
    }
}

Write-Output '1. Download PBIRS Content'
Write-Output '2. Deploy .pbix to PBIRS'
$select = Read-Host -Prompt 'Select'
switch($select) {
    1 {
        Write-Output 'Choose source server'
        Select-Server
        $folder = Read-Host -Prompt 'Input destination folder path'
        
        New-Item -ItemType Directory -Force -Path $folder

        $intAnswer = $a.popup("Do you want to download content from " + $server +"`nAnd put it into the " + $folder + " ?", 0,"Download content",4) 
        if($intAnswer -eq 6) { 
            Out-RsRestFolderContent -ReportPortalUri $server -RsFolder / -Recurse -Destination $folder -ErrorAction SilentlyContinue -Verbose
        }

    }

    2 {
        Write-Output 'Choose destination server'
        Select-Server
        $folder = Read-Host -Prompt 'Input source folder path'

        New-Item -ItemType Directory -Force -Path $folder'\.deploy'  # create folder to deploy
        Get-ChildItem $folder | Where-Object{$_.Name -notin ".deploy"} | Copy-Item -Destination $folder'\.deploy' -Recurse  # fill folder to deploy
        dir $folder'\.deploy\*' -Recurse -File | Where-Object{$_.Extension -notin ".pbix"} | Remove-Item  # keep only .pbix
        
        $intAnswer = $a.popup("Do you want to deploy .pbix to the " + $server +"`nFrom the " + $folder + " ?", 0,"Deploy .pbix",4) 
        if($intAnswer -eq 6) { 
            Write-RsRestFolderContent -ReportPortalUri $server -RsFolder / -Recurse -Path $folder'\.deploy' -Overwrite -ErrorAction SilentlyContinue -Verbose
        }

        Remove-Item -recurse $folder'\.deploy' -Force

    }
}

Read-Host 'Press Enter to close window'
