using namespace System.Collections
using namespace System.Collections.Generic
using namespace System.Diagnostics
using namespace System.IO
using namespace System.Windows.Forms

Import-Module posh-git
$ErrorActionPreference = 'Stop'

$script:Products = @{
    m3 = 'Math3'
    m4 = 'Math4'
    m5 = 'Math5'
    m6 = 'Math6'
    m7 = 'Math7'
    pa = 'PreAlgebra'
    a1 = 'Algebra1'
    ge = 'Geometry'
    a2 = 'Algebra2'
    pc = 'PreCalculus'
}

# Color scheme
$script:EmphasisColor = @{ ForegroundColor = 'Green' }
$script:ExitColor = @{ ForegroundColor = 'Magenta' }
$script:HeadlineColor = @{ ForegroundColor = 'White' }

# Script-level variables
$script:RootPath = [System.Environment]::GetEnvironmentVariables('User')['V4_ROOT_PATH']
$script:SharedRepositoryName = 'v4shared_src'
$script:SharedRepositoryPath = [System.Environment]::GetEnvironmentVariables('User')['V4_SHARED_PATH']
$script:StashLog = [ArrayList]@()

# Establish if environment variables already exist and point to valid directories
$script:IsInitialized = [bool](
    $script:RootPath -and $script:SharedRepositoryPath -and
    (Test-Path $script:RootPath) -and (Test-Path $script:SharedRepositoryPath)
)

<#
.SYNOPSIS
    Tests if the necessary Math repository folders exist.
.DESCRIPTION
    Ensures that at least the root directory for all Math repositories and the shared codebase exist.
    These folders are tested before any further work is allowed by this module.
.PARAMETER Force
    Override the default folder checks and prompt for a new folder selection.
.EXAMPLE
    Initialize-MathData

    MathData is already initialized. To reset this data, type: "Initialize-MathData -Force"
.EXAMPLE
    Initialize-MathData -Force

    Check behind this window for a dialog.

    MATHDATA VALUES
    ----------------

    V4_ROOT_PATH   : C:\git\randy test
    V4_SHARED_PATH : C:\git\randy test\v4shared_src
#>
function Initialize-MathData {
    param(
        [switch]
        $Force
    )

    Write-Host $script:StartLocation

    if (($Force -eq $false) -and $script:IsInitialized) {
        Write-Host (Get-MathData)
        Write-Host "`nMathData is already initialized. To reset this data, type: 'Initialize-MathData -Force'"
        return
    }

    # Prompt the use for the root directory needed by the module
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "`nCheck behind this window for a dialog." -ForegroundColor Yellow
    $chooseFolderDialog = New-Object FolderBrowserDialog -Property @{
        Description            = 'Select the root directory for your Math repositories'
        UseDescriptionForTitle = $true
    }
    $chooseFolderDialog.ShowDialog() > $null

    # If result is valid and the Shared Repository can be located, store environment variables
    $hasSelection = ($null -ne $chooseFolderDialog.SelectedPath)
    $isValidSelection = (Test-Path (Join-Path $chooseFolderDialog.SelectedPath $script:SharedRepositoryName))

    if ($hasSelection -and $isValidSelection) {
        $rootPath = $chooseFolderDialog.SelectedPath
        $sharedPath = (Join-Path $rootPath $script:SharedRepositoryName)

        [System.Environment]::SetEnvironmentVariable('V4_ROOT_PATH', $rootPath, 'User')
        [System.Environment]::SetEnvironmentVariable('V4_SHARED_PATH', $sharedPath, 'User')

        $script:RootPath = $rootPath
        $script:SharedRepositoryPath = $sharedPath
        $script:IsInitialized = $true
        Write-Host (Get-MathData)
        return
    }

    # Throw error if directories can not be located
    $script:IsInitialized = $false
    Write-Error "Can't locate the Math Root and Shared Repository folders." -CategoryActivity 'Initialize-MathData'
}

function Get-MathData {
    $data = @{
        V4_ROOT_PATH   = $script:RootPath
        V4_SHARED_PATH = $script:SharedRepositoryPath
    }

    return $data
}

function ConvertTo-InvocablePath {
    param (
        [System.IO.DirectoryInfo]
        $Path
    )

    return $Path -replace ' ', '` '
}

function Build-MxmlcOutput {
    param (
        [ValidateSet('win', 'win-store', 'android-armv7', 'android-x86', 'android-armv8')]
        [string]
        $Target = 'win',

        [Parameter(Mandatory = $true)]
        [string]
        $FlexPath,

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectPath
    )

    $mxmlc = $FlexPath + '\bin\mxmlc.bat'
    $flexConfigName = ($Target -eq 'win-store') ? 'flex-config-win-store.xml' : 'flex-config-win.xml'
    $flexConfigFile = $ProjectPath + '\' + $flexConfigName

    Start-Process -FilePath $mxmlc -ArgumentList "-load-config=$flexConfigFile", '-debug=false', '-warnings=false' -Wait -NoNewWindow | Out-Null
    # Invoke-Expression "$mxmlc -load-config=$flexConfigFile -debug=false -warnings=false" | Out-Null
}

function Build-CaptiveRuntimeBundle {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ProductName,

        [ValidateSet('win', 'win-store', 'android-armv7', 'android-x86', 'android-armv8')]
        [string]
        $Target = 'win',

        [Parameter(Mandatory = $true)]
        [string]
        $FlexPath,

        [Parameter(Mandatory = $true)]
        [string]
        $SharedPath,

        [Parameter(Mandatory = $true)]
        [string]
        $ProductPath,

        [Parameter(Mandatory = $true)]
        [string]
        $ProjectPath,

        [Parameter(Mandatory = $true)]
        [string]
        $BuildPath
    )

    $adt = $FlexPath + '\bin\adt.bat'
    $arch = ($Target -match 'win') ? $null : ($Target -replace 'android-', '')
    $keystoreFile = $SharedPath + '\tool\temp_keystore.p12'

    $adtArgs = ''

    # TODO: DRY this up
    if ($arch) {
        $apkPath = Split-Path $BuildPath -Parent

        $adtArgs = '-package', `
            '-target', 'apk-captive-runtime', `
            '-arch', $arch, `
            '-storetype', 'PKCS12', `
            '-keystore', $keystoreFile, `
            '-storepass', 'keystore_password', `
            "$apkPath\$ProductName-$arch.apk", `
            "$ProjectPath\AppMobile-app-generated-win.xml", `
            '-extdir', "$SharedPath\ane", `
            '-C', "$BuildPath $ProductName.swf", `
            '-C', "$ProductPath icon", `
            '-C', "$ProductPath package", `
            '-C', "$SharedPath\src assets\global\sfx", `
            '-C', "$SharedPath\src assets\global\image", `
            '-C', "$SharedPath\ane com.distriqt.NetworkInfo.ane", `
            '-C', "$SharedPath\ane com.distriqt.Core.ane", `
            '-C', "$SharedPath\ane com.distriqt.Share.ane", `
            '-C', "$SharedPath\ane androidx.core.ane", `
            '-C', "$SharedPath\src V4Downloader.swf"

    } else {
        $adtArgs = '-package', `
            '-storetype', 'PKCS12', `
            '-keystore', $keystoreFile, `
            '-storepass', 'keystore_password', `
            '-tsa', 'none', `
            '-target', "bundle $BuildPath\$ProductName $ProjectPath\AppDesktop-app-generated-win.xml", `
            '-C', "$BuildPath $ProductName.swf", `
            '-extdir', "$BuildPath", `
            '-C', "$ProductPath icon", `
            '-C', "$SharedPath\src assets\global\sfx", `
            '-C', "$SharedPath\src assets\global\image", `
            '-C', "$ProductPath package", `
            '-C', "$SharedPath\src V4Downloader.swf"
    }

    Start-Process $adt -Wait -NoNewWindow -ArgumentList $adtArgs | Out-Null
    # Invoke-Expression "$adt $adtArgs" | Out-Null

    if ($Error.Count -gt 0) {
        Write-Error "Error building $ProductName for $Target" -CategoryActivity 'Build-CaptiveRuntimeBundle'
        Write-Error -CategoryActivity 'Build-CaptiveRuntimeBundle'
    }
}

<#
.SYNOPSIS
    Build Math v4 products

.DESCRIPTION
    Compile the output SWF using mxlmc, and then create a Captive Runtime Bundle with adt.
    Optionally, update the relevant Git repositories and deploy the payloads to Dropbox.

.PARAMETER Product
    Determines which Math products to build. (If left blank, defaults to All products.)

.PARAMETER Store
    Use to build for the Microsoft Store by disabling standard updated checks.

.PARAMETER Update
    Use to ensure the product-specific repos are updated on the master branch,
    and you are given the opportunity to choose a v4shared_src release/test branch.

.PARAMETER Deploy
    Use to copy the payloads to Dropbox when prequisite actions are finished.

.EXAMPLE
    Build-MathPayload m3

    Math3Desktop payload is building...
    - 01 of 08: Declared variables in 15 ms
    - 02 of 08: Created build path in 68 ms
    - 03 of 08: Acquired app version in 82 ms
    - 04 of 08: Updated AppDesktop XML values in 92 ms
    - 05 of 08: Removed AppDesktop XML comments in 114 ms
    - 06 of 08: Saved AppDesktop XML generated file in 119 ms
    - 07 of 08: Compiled Output SWF in 48998 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 58399 ms

    All builds completed in 58407 ms

.EXAMPLE
    Build-MathPayload m3 -Update

    Updating Git repositories for m3...
    - Checkout and pull the master branch for C:\git\v4\v4m3_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m3_src

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 49543 ms

    Math3Desktop payload is building...
    - 01 of 08: Declared variables in 9 ms
    - 02 of 08: Created build path in 16 ms
    - 03 of 08: Acquired app version in 28 ms
    - 04 of 08: Updated AppDesktop XML values in 37 ms
    - 05 of 08: Removed AppDesktop XML comments in 48 ms
    - 06 of 08: Saved AppDesktop XML generated file in 53 ms
    - 07 of 08: Compiled Output SWF in 49771 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 59148 ms

    All builds completed in 108706 ms

.EXAMPLE
    Build-MathPayload m5 m6 -Update -Deploy

    Updating Git repositories for m5 m6...
    - Checkout and pull the master branch for C:\git\v4\v4m5_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m5_src
    - Checkout and pull the master branch for C:\git\v4\v4m6_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m6_src

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4m5_src on branch PM5-183-update-m5-assets
    - C:\git\v4\v4m6_src on branch PM6-334-update-m6-assets
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 28213 ms

    Math5Desktop payload is building...
    - 01 of 08: Declared variables in 9 ms
    - 02 of 08: Created build path in 58 ms
    - 03 of 08: Acquired app version in 69 ms
    - 04 of 08: Updated AppDesktop XML values in 77 ms
    - 05 of 08: Removed AppDesktop XML comments in 113 ms
    - 06 of 08: Saved AppDesktop XML generated file in 118 ms
    - 07 of 08: Compiled Output SWF in 47196 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 57317 ms

    Math6Desktop payload is building...
    - 01 of 08: Declared variables in 10 ms
    - 02 of 08: Created build path in 73 ms
    - 03 of 08: Acquired app version in 81 ms
    - 04 of 08: Updated AppDesktop XML values in 82 ms
    - 05 of 08: Removed AppDesktop XML comments in 85 ms
    - 06 of 08: Saved AppDesktop XML generated file in 87 ms
    - 07 of 08: Compiled Output SWF in 45711 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 57461 ms

    All builds completed in 143012 ms

    Deploying payloads...
    - Deployed Math5Desktop in 3345 ms
    - Deployed Math6Desktop in 3408 ms

    All deployments completed in 6765 ms
#>
function Build-MathPayload {
    param (
        [Parameter(ValueFromRemainingArguments = $true)]
        [ValidateSet('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc')]
        [string[]]
        $Product = ('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc'),

        [ValidateSet('win', 'win-store', 'android-armv7', 'android-x86', 'android-armv8')]
        [string]
        $Target = 'win',

        [switch]
        $Update,

        [switch]
        $Deploy
    )

    if ($script:IsInitialized -eq $false) { Initialize-MathData }
    if ($Update) { Sync-MathRepository -Product $Product -Target $Target }

    $startLocation = Get-Location

    $batchTimer = [Stopwatch]::StartNew();

    Write-Host 'Building payloads for ' -NoNewline
    Write-Host $Product @script:EmphasisColor -NoNewline
    Write-Host ' on ' -NoNewline
    Write-Host $Target @script:EmphasisColor

    $flexVersion = 'apache-flex-sdk-4.16.1-air'
    $flexDirectory = Get-Item (Join-Path -Path $script:RootPath v4sdk_pc $flexVersion)

    if ($Target -in ('win', 'win-store')) {
        $targetName = 'Desktop'
        $targetPath = $targetName.ToLower()
    } else {
        $targetName = 'Mobile'
        $targetPath = 'android'
    }

    $Product | ForEach-Object {
        $buildPath = Join-Path "~\tt\builds\$targetPath\" $_
        $productName = $script:Products.$_ + $targetName
        $productRepositoryName = "v4$_" + '_src'
        $projectRepositoryName = "v4$_" + "_proj_$targetPath"
        $productDirectory = Get-Item (Join-Path $script:RootPath $productRepositoryName)
        $projectDirectory = Get-Item (Join-Path $script:RootPath $projectRepositoryName)
        $sharedDirectory = Get-Item $script:SharedRepositoryPath

        # Change location for relative references in XML config files
        Set-Location $projectDirectory

        # Overwrite any existing build directory or create a new one
        Write-Progress -Status 'Initializing build directory' -Id 1 -Activity $productName -PercentComplete ((1 / 7) * 100)
        Start-Sleep 0.2
        if (Test-Path $buildPath) { Remove-Item $buildPath -Recurse }
        $buildPath = New-Item $buildPath -ItemType Directory

        # Acquire the current app version from Config.as
        Write-Progress -Status 'Acquire current app version' -Id 1 -Activity $productName -PercentComplete ((2 / 7) * 100)
        Start-Sleep 0.2
        $appConfig = Select-String -Path "$script:SharedRepositoryPath\src\com\tt\constants\Config.as" -Pattern '.+VERSION_BUILD.+\"(.+)\"'
        $appVersion = $appConfig.Matches[0].Groups[1].Value

        # Load the existing App descriptor XML and acquire necessary values
        Write-Progress -Status "Process App$targetName descriptor XML" -Id 1 -Activity $productName -PercentComplete ((3 / 7) * 100)
        Start-Sleep 0.2
        $appXml = New-Object -TypeName xml
        $appXml.Load("$projectDirectory\App$targetName-app.xml")
        $appXml.application.versionNumber = $appVersion
        $appXml.application.initialWindow.content = "$productName.swf"

        # Remove all comments from the XML
        Write-Progress -Status 'Remove all comments from the XML' -Id 1 -Activity $productName -PercentComplete ((4 / 7) * 100)
        Start-Sleep 0.2
        Select-Xml -Xml $appXml -XPath '//comment()' `
        | ForEach-Object { $_.Node.ParentNode.RemoveChild($_.Node) } | Out-Null

        # Save the generated XML file
        Write-Progress -Status 'Saving generated XML file' -Id 1 -Activity $productName -PercentComplete ((5 / 7) * 100)
        Start-Sleep 0.2
        $appXml.Save("$projectDirectory\App$targetName-app-generated-win.xml")

        # Convert Paths for use with Invoke-Expression
        $flexPath = (ConvertTo-InvocablePath($flexDirectory))
        $sharedPath = (ConvertTo-InvocablePath($sharedDirectory))
        $productPath = (ConvertTo-InvocablePath($productDirectory))
        $projectPath = (ConvertTo-InvocablePath($projectDirectory))

        # Compile the application SWF using MXMLC
        Write-Progress -Status 'Compiling SWF output' -Id 1 -Activity $productName -PercentComplete ((6 / 7) * 100)
        Start-Sleep 0.2
        Build-MxmlcOutput -Target $Target -FlexPath $flexPath -ProjectPath $projectPath

        # Package a Captive Runtime Bundle using ADT
        Write-Progress -Status 'Packaging Captive Runtime Bundle' -Id 1 -Activity $productName -PercentComplete ((7 / 7) * 100)
        Start-Sleep 0.2
        Build-CaptiveRuntimeBundle -ProductName $productName -Target $Target -FlexPath $flexPath -SharedPath $sharedPath -ProductPath $productPath -ProjectPath $projectPath -BuildPath $buildPath

        Start-Sleep 0.2
        Write-Progress -Status 'Done' -Id 1 -Activity $productName -Completed
        Write-Host '- Finished building ' -NoNewline
        Write-Host $productName @script:EmphasisColor
    }

    Start-Sleep 1

    Write-Host "All builds completed in [$($batchTimer.ElapsedMilliseconds)] ms" @script:ExitColor

    # Deploy the payload if requested
    if ($Deploy) {
        Deploy-MathPayload -Product $Product -Target $Target
    }

    # Return to the directory this commandlet is called from
    Set-Location $startLocation
}

<#
.SYNOPSIS
    Deploy Math v4 products

.DESCRIPTION
    Deploy the Math v4 payloads to Dropbox.
    Optionally, update the relevant Git repositories and build the payloads.

.PARAMETER Product
    Determines which Math products to deploy. (If left blank, defaults to All products.)

.PARAMETER Update
    Use to ensure the product-specific repos are updated on the master branch,
    and you are given the opportunity to choose a v4shared_src release/test branch.

.PARAMETER Build
    Use to compile and package the payloads prior to deployment.

.EXAMPLE
    Deploy-MathPayload m3

    Deploying payloads...
    - Deployed Math3Desktop in 3345 ms

    All deployments completed in 3361 ms

.EXAMPLE
    Deploy-MathPayload m3 -Update

    Updating Git repositories for m3...
    - Checkout and pull the master branch for C:\git\v4\v4m3_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m3_src

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 11175 ms

    Deploying payloads...
    - Deployed Math3Desktop in 3345 ms

    All deployments completed in 3361 m

.EXAMPLE
    Deploy-MathPayload m5 m6 -Update -Build

    Updating Git repositories for m5 m6...
    - Checkout and pull the master branch for C:\git\v4\v4m5_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m5_src
    - Checkout and pull the master branch for C:\git\v4\v4m6_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m6_src

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4m5_src on branch PM5-183-update-m5-assets
    - C:\git\v4\v4m6_src on branch PM6-334-update-m6-assets
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 28213 ms

    Math5Desktop payload is building...
    - 01 of 08: Declared variables in 9 ms
    - 02 of 08: Created build path in 58 ms
    - 03 of 08: Acquired app version in 69 ms
    - 04 of 08: Updated AppDesktop XML values in 77 ms
    - 05 of 08: Removed AppDesktop XML comments in 113 ms
    - 06 of 08: Saved AppDesktop XML generated file in 118 ms
    - 07 of 08: Compiled Output SWF in 47196 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 57317 ms

    Math6Desktop payload is building...
    - 01 of 08: Declared variables in 10 ms
    - 02 of 08: Created build path in 73 ms
    - 03 of 08: Acquired app version in 81 ms
    - 04 of 08: Updated AppDesktop XML values in 82 ms
    - 05 of 08: Removed AppDesktop XML comments in 85 ms
    - 06 of 08: Saved AppDesktop XML generated file in 87 ms
    - 07 of 08: Compiled Output SWF in 45711 ms
    - 08 of 08: Packaged Captive Runtime Bundle in 57461 ms

    All builds completed in 143012 ms

    Deploying payloads...
    - Deployed Math5Desktop in 3345 ms
    - Deployed Math6Desktop in 3408 ms

    All deployments completed in 6765 ms
#>
function Deploy-MathPayload {
    [OutputType([boolean])]
    param (
        [Parameter(ValueFromRemainingArguments = $true)]
        [ValidateSet('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc')]
        [string[]]
        $Product = ('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc'),

        [switch]
        $Update,

        [switch]
        $Build
    )

    $DropboxParent = $Store ? 'windows-store' : 'windows'
    $DropboxRoot = Get-Item "~/Teaching Textbooks Dropbox/Development/payloads/$DropboxParent/"

    if ($Update) { Sync-MathRepository $Product }
    if ($Build) { Build-MathPayload $Product }

    $batchTimer = [System.Diagnostics.Stopwatch]::StartNew();

    $i = 0
    $total = $Product.Count
    $Product | ForEach-Object {
        $percentageDone = ($i++ / $total) * 100
        Write-Progress -Activity 'Deploying payloads' -Status "$_" -PercentComplete $percentageDone

        $productName = $script:Products.$_ + 'Desktop'
        $DropboxPath = $DropboxRoot.ToString() + $productName

        if (Test-Path $DropboxPath) { Remove-Item $DropboxPath -Recurse }

        $Payload = Get-Item "/tt/builds/desktop/$_/$productName"
        Copy-Item $Payload $DropboxRoot -Recurse
    }

    Write-Host "`nAll deployments completed in [$($batchTimer.ElapsedMilliseconds)] ms" @script:ExitColor
}

<#
.SYNOPSIS
    Tidies up Git branches
.DESCRIPTION
    Prunes remote refs before deleting any merged branches in a repository.
.PARAMETER Path
    Path containing a valid Git repository.
.EXAMPLE
    Optimize-Repository
    Optimize-Repo ~/git/my-Repository
#>
function Optimize-Repository {
    param (
        [System.IO.DirectoryInfo]
        $Path = (Get-Item .)
    )

    Write-Progress "Optimizing [$Path]" -PercentComplete -1

    git -C $Path fetch --prune --prune-tags &&
    git -C $Path branch --merged |
        Select-String -NotMatch '\*|master|develop' |
            ForEach-Object {
                $branch = $_.ToString().Trim()
                if (git branch -D $branch) {
                    Write-Host "Deleted branch '$branch'"
                }
            }
}

<#
 # Check each Git Directory and fetch changes if possible
#>
function Initialize-MathRepo {
    param (
        [Parameter(Mandatory = $true)]
        [List[DirectoryInfo]]
        $GitDirectory
    )

    $i = 1
    $total = $GitDirectory.Count

    $GitDirectory | ForEach-Object {
        # Collect Git Status information
        $status = git -C $_ status -b -s
        $currentBranch = git -C $_ branch --show-current
        $isAhead = [bool]($status -match 'ahead [0-9]')
        $shouldStash = $status.Count -gt 1

        # Bail out if repo is invalid or not on a branch
        if ($null -eq $status) { Write-Error "$_ is not a valid Git repository." }
        if ($null -eq $currentBranch) { Write-Error "$_ is not currently on a branch." }

        # If Ahead, bail out
        if ($isAhead) {
            $err = "The [$_] repository is currently ahead on branch [$currentBranch]. Can not safely continue."
            Write-Error $err -CategoryActivity 'Initialize-MathRepo'
        }

        # Stash any changes not yet staged and commited
        if ($shouldStash) {
            git -C $_ stash -u -q
            $script:StashLog.Add("$_ on branch $($status.Branch)") > $null
        }

        # Fetch any changes from the remote repository
        git -C $_ fetch

        Write-Progress -Activity 'Fetching updates' -Status "$_" -PercentComplete $(($i++ / $total) * 100)
    }

    Start-Sleep 1
    Write-Host '- Fetched updates successfully.'
}

function Get-Branch {
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Directory
    )

    $validBranches = New-Object List[string]

    # Get the current branch
    $currentBranch = git -C $Directory branch --show-current

    # Get all remote branches and filter out any that are not release, test, develop, or master
    $allBranches = (git -C $Directory branch -r |
            Select-String -Pattern 'origin/(((release|test)/.*)|(develop|master))' |
                Select-String -Pattern 'origin/HEAD' -NotMatch)

    # Clean up and filter remote results
    if ($allBranches) {
        $allBranches = $allBranches.Matches.Value.Split('origin/') | Where-Object { $_ -ne '' }

        'release', 'test', 'develop', 'master' | ForEach-Object {
            $branchGroup = $allBranches | Select-String $_

            $branch = ($branchGroup) ? $branchGroup[$branchGroup.Count - 1] : $null
            if ($branch) { $validBranches.Add($branch) }
        }
    }

    # Add the current branch to the list if necessary
    if ($validBranches -notcontains $currentBranch) {
        $validBranches.Insert(0, $currentBranch)
    }

    return $validBranches
}

function Get-BranchSelection {
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Branch
    )

    #TODO: Get branches from the Shared repository

    #TODO: Prompt user to select a branch

    switch ($Branch) {
        'Quit' {
            Write-Error 'Terminating process' -CategoryActivity 'Get-BranchSelection'
        }
        'Skip' {
            Write-Host '- Skipping branch selection for Shared repository'
            return $null
        }
        default {
            return $Branch
        }
    }
}

function Get-ProductDirectory {
    param (
        [ValidateSet('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc')]
        [string[]]
        $Product = ('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc'),

        [ValidateSet('win', 'win-store', 'android-armv7', 'android-x86', 'android-armv8')]
        [string]
        $Target = 'win'
    )

    $targetDirectory = ''

    if ($Target -in 'win', 'win-store' ) {
        $targetDirectory = 'desktop'
    }

    if ($Target -in 'android-armv7', 'android-x86', 'android-armv8') {
        $targetDirectory = 'android'
    }

    $directories = Get-ChildItem $script:RootPath -Directory |
        Where-Object {
            $_.Name -match '^v4' -and
            $_.Name -notmatch 'download|sdk|shared' -and
            $_.Name -match "$targetDirectory|src" -and
            $_.Name -match ($Product -join '|')
        }

    if ($directories.Count -ne $Product.Count * 2) {
        Write-Error 'Not all product directories found.' -CategoryActivity 'Get-ProductDirectory'
    }

    return $directories
}

<#
.SYNOPSIS
    Update Math v4 repositories
.DESCRIPTION
    Stash any changes on any designated project repositories, and then checkout and pull the 'master' branch for each.
    Stash any changes on the Shared repository, and then checkout and pull any designated test/release branch.

.PARAMETER Product
    Determines which Math product repositories to update. (If left blank, defaults to All products.)

.PARAMETER SkipProducts
    Use to bypass updating the product-specific project repositories and only work with the Shared respository.

.EXAMPLE
    Sync-MathRepository -SkipProjects

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 9245 ms

.EXAMPLE
    Sync-MathRepository

    Updating Git repositories for a1 a2 ge m3 m4 m5 m6 m7 pa pc...
    - Checkout and pull the master branch for C:\git\v4\v4a1_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4a1_src
    - Checkout and pull the master branch for C:\git\v4\v4a2_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4a2_src
    - Checkout and pull the master branch for C:\git\v4\v4ge_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4ge_src
    - Checkout and pull the master branch for C:\git\v4\v4m3_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m3_src
    - Checkout and pull the master branch for C:\git\v4\v4m4_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m4_src
    - Checkout and pull the master branch for C:\git\v4\v4m5_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m5_src
    - Checkout and pull the master branch for C:\git\v4\v4m6_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m6_src
    - Checkout and pull the master branch for C:\git\v4\v4m7_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4m7_src
    - Checkout and pull the master branch for C:\git\v4\v4pa_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4pa_src
    - Checkout and pull the master branch for C:\git\v4\v4pc_proj_desktop
    - Checkout and pull the master branch for C:\git\v4\v4pc_src

    Updating Git Shared Repository...
    - Checkout and pull the release/4.0.842 branch for C:\git\v4\v4shared_src\

    Changes stashed for these repositories:
    - C:\git\v4\v4m5_src on branch PM5-183-update-m5-assets
    - C:\git\v4\v4m6_src on branch PM6-334-update-m6-assets
    - C:\git\v4\v4m7_src on branch PM7-433-m7-assets-update
    - C:\git\v4\v4shared_src\ on branch test/4.0.842-A

    All repositories updated in 60384 ms

#>
function Sync-MathRepository {
    param (
        [Parameter(
            ValueFromRemainingArguments = $true
        )]
        [ValidateSet('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc')]
        [string[]]
        $Product = ('a1', 'a2', 'ge', 'm3', 'm4', 'm5', 'm6', 'm7', 'pa', 'pc'),

        [ValidateSet('win', 'win-store', 'android-armv7', 'android-x86', 'android-armv8')]
        [string]
        $Target = 'win'
    )

    if ($script:IsInitialized -eq $false) { Initialize-MathData }

    # Start a timer and get all repo directories
    $batchTimer = [System.Diagnostics.Stopwatch]::StartNew()

    # Create a List for all directories to be updated
    [List[DirectoryInfo]]$directories = @(Get-ProductDirectory $Product $Target)
    $directories.Add((Get-Item $script:SharedRepositoryPath)) > $null

    Write-Output 'Updating Math repositories '
    Write-Host 'v4shared_src ' @script:EmphasisColor -NoNewline
    if ($SkipProducts -ne $true) { Write-Host "$Product " @script:EmphasisColor -NoNewline }
    Write-Host 'for target ' -NoNewline
    Write-Host $Target @script:EmphasisColor

    # Perform preliminary checks on each repository, stash changes, and fetch updates
    Initialize-MathRepo $directories

    # TODO: Nest all of these in a single function
    $sharedBranches = Get-Branch $script:SharedRepositoryPath
    $sharedOptions = Read-BranchSelection $sharedBranches
    $sharedBranchSelection = Get-BranchSelection $sharedOptions
    Write-Host

    $i = 1
    $total = $directories.Count
    $branch = $null

    # Update each repository
    $directories | ForEach-Object {
        Write-Progress -Activity 'Applying updates' -Status "$_" -PercentComplete $(($i++ / $total) * 100)

        $branch = 'master'

        # Offer branch selection for the Shared Repo, otherwise default to the master branch
        if ($_.FullName -eq $script:SharedRepositoryPath) {
            $branch = $sharedBranchSelection
        }

        # Checkout and pull the selected branch
        if ($branch) {
            git -C $_ checkout $branch -q
            git -C $_ pull -q
        }

        Write-Host '- Updated repository: ' -NoNewline
        Write-Host "$_" @script:EmphasisColor
    }

    # TODO: Nest in a function
    # Log stashed changes
    if ($script:StashLog.Count -gt 0) {
        Write-Host '- Changes stashed for:'
        $script:StashLog | ForEach-Object {
            Write-Host '  - ' -NoNewline
            Write-Host "$_" @script:EmphasisColor
        }
    }

    Write-Host "All repositories updated in [$($batchTimer.ElapsedMilliseconds)] ms." @script:ExitColor
}

function Read-BranchSelection {
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string[]]
        $Branch
    )

    $caption = "Select the desired branch for the [$script:SharedRepositoryPath] repository"
    $message = $null
    $default = 0

    # Create an array of ChoiceDescriptions and populate it
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@()

    $Branch | ForEach-Object {
        $label = (Get-ChoiceLabel $_)
        $option = [System.Management.Automation.Host.ChoiceDescription]::new($label,
            "Check out and pull [$_] for the Shared repository")

        if (($choices.Count -eq 0) -or ($choices.Label.Contains($option.Label) -eq $false)) {
            $choices += $option
        }
    }

    $choices += [System.Management.Automation.Host.ChoiceDescription]::new(
        '&skip', 'Continue without updating the Shared repository')

    $choices += [System.Management.Automation.Host.ChoiceDescription]::new(
        '&quit', 'Terminate this process')

    # Prompt the user to make a selection
    $prompt = $host.UI.PromptForChoice($caption, $message, $choices, $default)
    return $choices[$prompt].Label -replace '\&', ''
}

function Get-ChoiceLabel {
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Branch
    )

    # For non-feature branches, return a default label
    if ($Branch -match '^(release|test|master|develop)') {
        return "&$Branch"
    } else {
        return "&current ($Branch)"
    }
}

function Read-MathGitStash {
    $mathDirectories = Get-ChildItem $script:RootPath -Directory
    $mathDirectories | ForEach-Object {
        $stash = git -C $_ stash list

        if ($stash) {
            Write-Host "`n$($_.Name)" @script:EmphasisColor

            $stash | ForEach-Object {
                if ($_ -match 'stash@{(?<index>[0-9]+)}:\s(?<message>.+)$') {
                    Write-Host "  |- $($Matches.index): $($Matches.message)"
                }
            }
        }
    }
}