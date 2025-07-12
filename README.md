# Deploy Math

Deploy Math is a PowerShell module/library for building and deploying the Teaching Textbooks Math v4 products.

## Installation

Using PowerShell, run the following commands. Use the Git Clone link provided form Bitbucket instead of the example.

```PowerShell
$ModulePath = '~/Documents/PowerShell/Modules'
if ((Test-Path $ModulePath) -eq $false) { mkdir $ModulePath }
cd $ModulePath
git clone ...
```

## Usage

### Example 1: Build all products

```PowerShell
Build-Math
```

### Exampe 2: Build Math 3 after updating the involved Git repositories

```PowerShell
Build-Math m3 -UpdateRepos
```

### Example 3: Build various products after updating their Git repositories, and then deploy the payloads to Dropbox

```PowerShell
Build-Math m6 a1 pc -UpdateRepos -Deploy
```
