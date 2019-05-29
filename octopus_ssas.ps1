# deploy-ssas-from-package 2.0
# https://library.octopus.com/step-templates/1409c3dd-e87d-49f1-9b4f-382af800b75d/actiontemplate-deploy-ssas-from-package

$ErrorActionPreference = 'Stop'
$regKeyFormat = 'HKLM:\Software\Wow6432Node\Microsoft\Microsoft SQL Server\{0}0\Tools\ClientSetup\'
$vsVersions = @( '12' )

function Validate-Argument($name, $value) {
    if (!$value) {
        throw ('Missing required value for parameter ''{0}''.' -f $name)
    }
    return $value
}

# Returns the Microsoft.AnalysisServices.Deployment.exe path
function Load-SsasDeploy {
    # display
    Write-Host "Beginning search of registry starting in HKLM:\Software\Wow6432Node\Microsoft\Microsoft SQL Server\ ..."
    
    # define base registry key
    $regKeyFormat = 'HKLM:\Software\Wow6432Node\Microsoft\Microsoft SQL Server\'

    # get all of the SQL registry keys
    $SQLRegistryKeys = Get-ChildItem -Path $regKeyFormat | Where-Object {$TestVar = 0; [int]::TryParse($_.Name.Substring($_.Name.LastIndexOf("\") + 1), [ref] $TestVar) -eq $true} | Sort-Object {[int]$_.Name.Substring($_.Name.LastIndexOf("\") + 1)} -Descending

    # display number of items found
    Write-Host "Found $($SQLRegistryKeys.Length) item(s) ..."

    # make sure something was returned
    if($SQLRegistryKeys -ne $null)
    {
        # display
        Write-Host "Searching keys for ClientSetup folder ..."
        
        # retrieve paths to the tools\binn path
        $SQLToolsPath = $SQLRegistryKeys | ForEach-Object {Get-ChildItem "HKLM:\$_" -Recurse | Where-Object {$_.Name -like "*ClientSetup" }} | Where-Object {$_.Property -eq "Path"} | ForEach-Object {Get-ItemProperty -Path "HKLM:\$_" -Name Path} | Select Path

        # make sure paths were found
        if($SQLToolsPath -ne $null)
        {
            # display
            Write-Host "Found $($SQLToolsPath.Length) item(s), searching in descending order ..."
            
            # loop through returned paths looking for Microsoft.AnalysisServices.Deployment.exe
            ForEach($Path in $SQLToolsPath)
            {
                # check to see if .exe exists
                if(($ReturnPath = Get-ChildItem -Path $Path.Path -Recurse | Where-Object {$_.Name -eq "Microsoft.AnalysisServices.Deployment.exe"}) -ne $null)
                {
                    # display
                    Write-Host "Found $($ReturnPath.FullName) ..."

                    # return
                    return $ReturnPath.FullName
                }
            }

            # display error
            Write-Error "Unable to find Microsoft.AnalysisServices.Deployment.exe! `r`nSearched $SQLToolsPath."
        }
        else
        {
            # display error
            Write-Error "No ClientSetup registry keys found!"
        }
    }
    else
    {
        # display error
        Write-Error "No SQL installations found!"
    }
}

# Update Deploy xml (.deploymenttargets)
function Update-Deploy {
	[xml]$deployContent = Get-Content $file
	$deployContent.DeploymentTarget.Database = $ssasDatabase 
	$deployContent.DeploymentTarget.Server = $ssasServer
	$deployContent.DeploymentTarget.ConnectionString = 'DataSource=' + $ssasServer + ';Timeout=0'
	$deployContent.Save($file)
}
# Update Config xml (.configsettings)
function Update-Config {
	[xml]$configContent = Get-Content $file
    $configContent.ConfigurationSettings.Database.DataSources.DataSource.ConnectionString = 'Provider=SQLNCLI11.1;Data Source=' + $dbServer + ';Integrated Security=SSPI;Initial Catalog=' + $dbDatabase
	$configContent.Save($file)
}
# Create Config xml (.configsettings) 2.0
function Create-Config {
    [xml]$configContent = '<ConfigurationSettings xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:ddl2="http://schemas.microsoft.com/analysisservices/2003/engine/2" xmlns:ddl2_2="http://schemas.microsoft.com/analysisservices/2003/engine/2/2" xmlns:ddl100_100="http://schemas.microsoft.com/analysisservices/2008/engine/100/100" xmlns:ddl200="http://schemas.microsoft.com/analysisservices/2010/engine/200" xmlns:ddl200_200="http://schemas.microsoft.com/analysisservices/2010/engine/200/200" xmlns:ddl300="http://schemas.microsoft.com/analysisservices/2011/engine/300" xmlns:ddl300_300="http://schemas.microsoft.com/analysisservices/2011/engine/300/300" xmlns:ddl400="http://schemas.microsoft.com/analysisservices/2012/engine/400" xmlns:ddl400_400="http://schemas.microsoft.com/analysisservices/2012/engine/400/400" xmlns:ddl500="http://schemas.microsoft.com/analysisservices/2013/engine/500" xmlns:ddl500_500="http://schemas.microsoft.com/analysisservices/2013/engine/500/500" xmlns:dwd="http://schemas.microsoft.com/DataWarehouse/Designer/1.0">
      <Database>
        <DataSources>
          <DataSource>
            <ID></ID>
            <ConnectionString></ConnectionString>
            <ManagedProvider></ManagedProvider>
            <ImpersonationInfo>
              <ImpersonationMode xmlns="http://schemas.microsoft.com/analysisservices/2003/engine">ImpersonateServiceAccount</ImpersonationMode>
              <Account xmlns="http://schemas.microsoft.com/analysisservices/2003/engine"></Account>
              <Password xmlns="http://schemas.microsoft.com/analysisservices/2003/engine"></Password>
              <ImpersonationInfoSecurity xmlns="http://schemas.microsoft.com/analysisservices/2003/engine">Unchanged</ImpersonationInfoSecurity>
            </ImpersonationInfo>
          </DataSource>
        </DataSources>
      </Database>
    </ConfigurationSettings>'
    $configContent.ConfigurationSettings.Database.DataSources.DataSource.ID = $dbDatasourceConnectionName
    $configContent.ConfigurationSettings.Database.DataSources.DataSource.ConnectionString = "Provider=SQLNCLI11.1;Data Source=" + $dbServer + ";Integrated Security=SSPI;Initial Catalog=" + $dbDatabase
    $configContent.Save($file)
}
# Update Config xml (.deploymentoptions)
function Update-Option {
	[xml]$optionContent = Get-Content $file
    $optionContent.DeploymentOptions.ProcessingOption = 'DoNotProcess'
	$optionContent.Save($file)
}

# Get arguments
$ssasPackageStepName = Validate-Argument 'SSAS Package Step Name' $OctopusParameters['SsasPackageStepName']
$ssasServer = Validate-Argument 'SSAS server name' $OctopusParameters['SsasServer']
$ssasDatabase = Validate-Argument 'SSAS database name' $OctopusParameters['SsasDatabase']
$dbServer = Validate-Argument 'SSAS source server' $OctopusParameters['SrcServer']
$dbDatabase = Validate-Argument 'SSAS source database' $OctopusParameters['SrcDatabase']
$dbDatasourceConnectionName = Validate-Argument 'SSAS source datasource connection name' $OctopusParameters['SrcDatasourceConnectionName']

# Set .NET CurrentDirectory to package installation path
$installDirPathFormat = 'Octopus.Action[{0}].Output.Package.InstallationDirectoryPath' -f $ssasPackageStepName
$installDirPath = $OctopusParameters[$installDirPathFormat]

#$ssasServer     = 'server2\md_dev'
#$ssasDatabase   = 'BusinessIntelligence'
#$dbServer       = 'server1\dev'
#$dbDatabase     = 'Warehouse'
#$installDirPath = 'c:\packages\v1'

Write-Verbose ('Setting CurrentDirectory to ''{0}''' -f $installDirPath)
[System.Environment]::CurrentDirectory = $installDirPath

$exe = Load-SsasDeploy

$files = Get-ChildItem –Path $installDirPath\* -Include *.deploymenttargets
foreach ($file in $files) {
  $name = [IO.Path]::GetFileNameWithoutExtension($file)

  Write-Host 'Updating' $file
  Update-Deploy
  
  $file = $installDirPath + '\' + $name + '.configsettings'
  #if(Test-Path $file) {
      #Write-Host 'Updating' $file
      #Update-Config
  #} else {
    #Write-Host "Config settings doesn't exist. Skipping."
  #}
  Write-Host 'Creating' $file
  Create-Config # 2.0
  
  $file = $installDirPath + '\' + $name + '.deploymentoptions'
  Write-Host 'Updating' $file
  Update-Option

  $arg = '"' + $installDirPath + '\' + $name + '.asdatabase" /s:"' + $installDirPath + '\Log.txt"'
  Write-Host $exe $arg
  $execute = [scriptblock]::create('& "' + $exe + '" ' + $arg)
  Invoke-Command -ScriptBlock $execute
  
  $log = [IO.File]::ReadAllLines('' + $installDirPath + '\Log.txt')
  if($log[-1] -eq "done") {
    Write-Host $log
  }
  else {
    Write-Error "$log"
  }
}
