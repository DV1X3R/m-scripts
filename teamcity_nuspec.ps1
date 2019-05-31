$workingDir = '%teamcity.build.workingDir%'
$ssisDir = 'SSIS\'
$ssasDir = 'SSAS\'
#$authors = 'Author'

foreach ($asdatabase in (Get-ChildItem -Recurse -Path $ssasDir -Include "*.asdatabase")) {
    [xml]$nuspec = '<?xml version="1.0" encoding="utf-8"?>
    <package xmlns="http://schemas.microsoft.com/packaging/2010/07/nuspec.xsd">
      <metadata>
        <id></id>
        <version>$version$</version>
        <authors></authors>
        <requireLicenseAcceptance>false</requireLicenseAcceptance>
        <description>NuGet package generated automatically</description>
      </metadata>
      <files> <!-- PLEASE DO NOT REMOVE --> </files>
    </package>'

    $nuspec.package.metadata.id = "SSAS." + $asdatabase.BaseName
    $nuspec.package.metadata.authors = $authors

    $asdatabaseBasePath = ($asdatabase | Resolve-Path -Relative | Split-Path) + "\" + $asdatabase.BaseName

    $file = $nuspec.CreateElement('file')
    $file.SetAttribute('src', $asdatabaseBasePath + ".asdatabase")
    $file.SetAttribute('target', $asdatabase.BaseName + ".asdatabase")
    $nuspec.package.files.AppendChild($file)
    
    $file = $nuspec.CreateElement('file')
    $file.SetAttribute('src', $asdatabaseBasePath + ".deploymentoptions")
    $file.SetAttribute('target', $asdatabase.BaseName + ".deploymentoptions")
    $nuspec.package.files.AppendChild($file)
    
    $file = $nuspec.CreateElement('file')
    $file.SetAttribute('src', $asdatabaseBasePath + ".deploymenttargets")
    $file.SetAttribute('target', $asdatabase.BaseName + ".deploymenttargets")
    $nuspec.package.files.AppendChild($file)
    
    $nuspec.save($workingDir + '\SSAS.' + $asdatabase.BaseName + '.nuspec')
}

foreach ($ispac in (Get-ChildItem -Recurse -Path $ssisDir -Include "*.ispac")) {
    [xml]$nuspec = '<?xml version="1.0" encoding="utf-8"?>
    <package xmlns="http://schemas.microsoft.com/packaging/2010/07/nuspec.xsd">
      <metadata>
        <id></id>
        <version>$version$</version>
        <authors></authors>
        <requireLicenseAcceptance>false</requireLicenseAcceptance>
        <description>NuGet package generated automatically</description>
      </metadata>
      <files> <!-- PLEASE DO NOT REMOVE --> </files>
    </package>'
    
    $nuspec.package.metadata.id = "SSIS." + $ispac.BaseName
    $nuspec.package.metadata.authors = $authors

    $file = $nuspec.CreateElement('file')
    $file.SetAttribute('src', ($ispac | Resolve-Path -Relative))
    $file.SetAttribute('target', $ispac.Name)
    $nuspec.package.files.AppendChild($file)

    $nuspec.save($workingDir + '\SSIS.' + $ispac.BaseName + '.nuspec')
}