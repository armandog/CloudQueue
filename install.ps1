# This script installs and configure a Windows Service to handle the
# Quickspace Cloud Queue printing service at warehouses

# check if the folder at Program Files exists %PROGAMFILES%\QsCloudQueue
# if not then create a new one
$programsFolder = Get-Content -Path Env:PROGRAMFILES
$cloudQueueFolder = $programsFolder + "\QSCloudQueue"
$cloudQueueBinFolder = $cloudQueueFolder + "\bin"
$cloudQueueTemplatesFolder = $cloudQueueFolder + "\Templates"

# get user/passwd to access the queue
$awsUser = Read-Host -Prompt 'Enter your Quickspace CloudQueue user'
$awsPasswd = Read-Host -Prompt 'Enter your Quickspace CloudQueue password' -MaskInput
$warehouseName = Read-Host -Prompt 'Enter you warehouse name:'

# set environment variables
Write-Progress -Id 0 -Activity 'Setting your environment'
[System.Environment]::SetEnvironmentVariable('AWS_ACCESS_KEY_ID',$awsUser,'Machine')
[System.Environment]::SetEnvironmentVariable('AWS_SECRET_ACCESS_KEY',$awsPasswd,'Machine')
[System.Environment]::SetEnvironmentVariable('QS_CLOUD_QUEUE_PATH',$cloudQueueFolder,'Machine')
[System.Environment]::SetEnvironmentVariable('QS_CLOUD_QUEUE_WAREHOUSE',$warehouseName,'Machine')
Write-Progress -Id 0 -Activity 'Environment set' -Completed

# guarantee folder structure
Write-Progress -Id 1 -Activity 'Downloading QSCloudQueue' -Status 'Setting up your folder' -PercentComplete 0
if (-Not (Test-Path -Path $cloudQueueFolder)) {
	# create the QS Cloud Queue folder
	[void](New-Item -Path $programsFolder -Name "QSCloudQueue" -ItemType "directory")
}

if (-Not (Test-Path -Path $cloudQueueBinFolder)) {
    [void](New-Item -Path $cloudQueueFolder -Name "bin" -ItemType "directory")
}

if (-Not (Test-Path -Path $cloudQueueTemplatesFolder)) {
    [void](New-Item -Path $cloudQueueFolder -Name "Templates" -ItemType "directory")
}

# download required binaries and templates
Write-Progress -Id 1 -Activity 'Downloading QSCloudQueue' -Status 'Downloading the service solution' -PercentComplete 25
$serviceSolution = $cloudQueueFolder + '\bin\CloudQueueSvc.exe'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/bin/CloudQueueSvc.exe -OutFile $serviceSolution
$clrcompression_dll = $cloudQueueFolder + '\bin\clrcompression.dll'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/bin/clrcompression.dll -OutFile $clrcompression_dll
$clrjit_dll = $cloudQueueFolder + '\bin\clrjit.dll'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/bin/clrjit.dll -OutFile $clrjit_dll
$coreclr_dll = $cloudQueueFolder + '\bin\coreclr.dll'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/bin/coreclr.dll -OutFile $coreclr_dll
$mscordaccore_dll = $cloudQueueFolder + '\bin\mscordaccore.dll'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/bin/mscordaccore.dll -OutFile $mscordaccore_dll

# download templates
Write-Progress -Id 1 -Activity 'Downloading QSCloudQueue' -Status 'Templates' -PercentComplete 75
$warehouseLabelbase = $cloudQueueFolder + '\Templates\WarehouseLabel_base.lbx'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/templates/WarehouseLabel_base.lbx -OutFile $warehouseLabelbase
$warehouseLabel385 = $cloudQueueFolder + '\Templates\WarehouseLabel_385.lbx'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/templates/WarehouseLabel_385.lbx -OutFile $warehouseLabel385
$warehouseLabel259 = $cloudQueueFolder + '\Templates\WarehouseLabel_259.lbx'
Invoke-WebRequest -Uri https://github.com/armandog/CloudQueue/raw/main/templates/WarehouseLabel_259.lbx -OutFile $warehouseLabel259
Write-Progress -Id 1 -Activity 'Downloading QSCloudQueue' -Completed

Write-Progress -Id 2 -Activity 'Configuring Windows Services'
# configure service
$serviceName = 'QuickspaceCloudQueue2'
$serviceDisplayName = 'Quickspace Cloud Queue'

# check if there the service already exists
$existingService = Get-Service -Name $serviceName
if (-Not ($null -eq $existingService)) {
    # check is it running
   if ((Get-Service -Name $serviceName).Status -eq 'Running') {
       # stop the service
       Stop-Service -Name $serviceName
   }
   # remove
   Remove-Service -Name $serviceName
}

# create the service
New-Service -Name $serviceName -DisplayName $serviceDisplayName -BinaryPathName $serviceSolution -StartupType 'Automatic'

# start the service
#Start-Service -Name $serviceName

