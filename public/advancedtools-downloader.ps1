#╔════════════════════════════════════════════════════════════════════════════════╗
#║                                                                                ║
#║   download-tool.ps1                                                            ║
#║                                                                                ║
#╟────────────────────────────────────────────────────────────────────────────────╢
#║   Guillaume Plante <codegp@icloud.com>                                         ║
#║   Code licensed under the GNU GPL v3.0. See the LICENSE file for details.      ║
#╚════════════════════════════════════════════════════════════════════════════════╝


[CmdletBinding(SupportsShouldProcess = $true)]
param()


function Read-FileHeader {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true, HelpMessage = 'Path to the file with header')]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$Path
    )

    process {
        $ExpectedMagic = [byte[]](0x42, 0x4D, 0x57, 0x21, 0x2A, 0x4D, 0x53, 0x47)

        $Reader = [System.IO.BinaryReader]::new([System.IO.File]::OpenRead($Path))
        try {
            $MagicStart = $Reader.ReadBytes(8)
            if (-not [System.Linq.Enumerable]::SequenceEqual($MagicStart, $ExpectedMagic)) {
                throw "Invalid or missing header magic number at start of file: $Path"
            }

            $PartID = $Reader.ReadInt32()
            $DataSize = $Reader.ReadInt64()
            $HashBytes = $Reader.ReadBytes(32)

            $MagicEnd = $Reader.ReadBytes(8)
            if (-not [System.Linq.Enumerable]::SequenceEqual($MagicEnd, $ExpectedMagic)) {
                throw "Invalid or missing header magic number at end of header: $Path"
            }
        }
        finally {
            $Reader.Close()
        }

        [pscustomobject]@{
            Path = $Path
            PartID = $PartID
            DataSize = $DataSize
            Hash = ([BitConverter]::ToString($HashBytes) -replace '-', '').ToLower()
        }
    }
}

function Remove-FileHeader {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true, HelpMessage = 'Path to the file with header')]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$Path
    )

    process {
        $Magic = [byte[]](0x42, 0x4D, 0x57, 0x21, 0x2A, 0x4D, 0x53, 0x47) # BMW!*MSG
        $LitteralPath = (Resolve-Path $Path).Path
        $TempPath = "$LitteralPath.raw"

        $Reader = [System.IO.BinaryReader]::new([System.IO.File]::OpenRead($LitteralPath))
        try {
            # Read and validate start magic number
            $MagicStart = $Reader.ReadBytes(8)
            if (-not [System.Linq.Enumerable]::SequenceEqual($MagicStart, $Magic)) {
                throw "Invalid or missing header magic number at start of file: $LitteralPath"
            }

            # Read header
            $PartID = $Reader.ReadInt32()
            $DataSize = $Reader.ReadInt64()
            $HashBytes = $Reader.ReadBytes(32)

            # Read and validate end magic number
            $MagicEnd = $Reader.ReadBytes(8)
            if (-not [System.Linq.Enumerable]::SequenceEqual($MagicEnd, $Magic)) {
                throw "Invalid or missing header magic number at end of header: $LitteralPath"
            }

            # Read actual data
            $RemainingBytes = $Reader.ReadBytes([int]$DataSize)

            # Write payload to new file
            $Writer = [System.IO.BinaryWriter]::new([System.IO.File]::Create($TempPath))
            try {
                $Writer.Write($RemainingBytes, 0, $RemainingBytes.Length)
            }
            finally {
                $Writer.Close()
            }
        }
        finally {
            $Reader.Close()
        }

        # Replace the original file with headerless copy
        Remove-Item -LiteralPath $LitteralPath -Force
        Rename-Item -LiteralPath $TempPath -NewName (Split-Path $LitteralPath -Leaf)

        # Return metadata object
        return [pscustomobject]@{
            Path = $LitteralPath
            PartID = $PartID
            DataSize = $DataSize
            Hash = ([BitConverter]::ToString($HashBytes) -replace '-', '').ToLower()
        }
    }
}

function Write-FileHeader {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true, HelpMessage = 'Path to the file to add header to')]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [int]$PartID
    )

    process {
        $Magic = [byte[]](0x42, 0x4D, 0x57, 0x21, 0x2A, 0x4D, 0x53, 0x47) # BMW!*MSG
        $Bytes = [System.IO.File]::ReadAllBytes($Path)
        $DataSize = $Bytes.Length
        $Hasher = [System.Security.Cryptography.SHA256]::Create()
        $HashBytes = $Hasher.ComputeHash($Bytes)

        $TempPath = "$Path.tmp"
        $Writer = [System.IO.BinaryWriter]::new([System.IO.File]::Create($TempPath))

        try {
            $Writer.Write($Magic)
            $Writer.Write([int]$PartID)
            $Writer.Write([long]$DataSize)
            $Writer.Write($HashBytes)
            $Writer.Write($Magic)
            $Writer.Write($Bytes, 0, $DataSize)
        }
        finally {
            $Writer.Close()
        }

        Remove-Item -LiteralPath $Path -Force
        Rename-Item -LiteralPath $TempPath -NewName (Split-Path $Path -Leaf)
    }
}


function Invoke-AutoUpdateProgress_FileUtils {
    [int32]$PercentComplete = (($Script:StepNumber / $Script:TotalSteps) * 100)
    if ($PercentComplete -gt 100) { $PercentComplete = 100 }
    Write-Progress -Activity $Script:ProgressTitle -Status $Script:ProgressMessage -PercentComplete $PercentComplete
    if ($Script:StepNumber -lt $Script:TotalSteps) { $Script:StepNumber++ }
}


function Sort-Lexically {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $false)]
        [string]$Property
    )

    begin {
        $items = @()
    }

    process {
        $items += $InputObject
    }

    end {
        $items | Sort-Object {
            $name = if ($Property) { $_.$Property } else { $_ }

            if ($name -match '(\d+)(?=\D*$)') {
                [int]$matches[1] # numeric suffix before non-digits at the end (e.g., .cpp)
            } else {
                $name
            }
        }
    }
}

function Sort-ByFileHeaderId {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true)]
        [string[]]$Path
    )

    begin {
        $items = @()
    }

    process {
        $items += $Path
    }

    end {
        $items |
        ForEach-Object {
            try {
                $header = Read-FileHeader -Path $_
                [pscustomobject]@{
                    Path = $_
                    PartID = $header.PartID
                }
            } catch {
                Write-Warning "Skipping invalid or corrupt header: $_"
            }
        } |
        Sort-Object PartID |
        Select-Object -ExpandProperty Path
    }
}


function Invoke-SplitDataFile {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [int64]$Newsize = 1MB,
        [Parameter(Mandatory = $false)]
        [string]$OutPath,
        [Parameter(Mandatory = $false)]
        [string]$Extension = "cpp",
        [Parameter(Mandatory = $false)]
        [switch]$AsString
    )

    if ($Newsize -le 0)
    {
        Write-Error "Only positive sizes allowed"
        return
    }

    $FileSize = (Get-Item $Path).Length
    $SyncStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Script:ProgressTitle = "Split Files"
    $TotalTicks = 0
    $Count = [math]::Round($FileSize / $Newsize)
    $Script:StepNumber = 1
    $Script:TotalSteps = $Count + 3
    if ($PSBoundParameters.ContainsKey('OutPath') -eq $False) {
        $OutPath = [IO.Path]::GetDirectoryName($Path)

        Write-Verbose "Using OutPath from Path $Path"
    } else {
        Write-Verbose "Using OutPath $OutPath"
    }
    $OutPath = $OutPath.TrimEnd('\')

    if (-not (Test-Path -Path "$OutPath")) {
        Write-Verbose "CREATING $OutPath"
        $Null = New-Item $OutPath -ItemType Directory -Force -ErrorAction Ignore
    }

    $FILENAME = [IO.Path]::GetFileNameWithoutExtension($Path)


    $MAXVALUE = 1GB # Hard maximum limit for Byte array for 64-Bit .Net 4 = [INT32]::MaxValue - 56, see here https://stackoverflow.com/questions/3944320/maximum-length-of-byte
    # but only around 1.5 GB in 32-Bit environment! So I chose 1 GB just to be safe
    $PASSES = [math]::Floor($Newsize / $MAXVALUE)
    $REMAINDER = $Newsize % $MAXVALUE
    if ($PASSES -gt 0) { $BUFSIZE = $MAXVALUE } else { $BUFSIZE = $REMAINDER }

    $OBJREADER = New-Object System.IO.BinaryReader ([System.IO.File]::Open($Path, 'Open', 'Read', 'Read')) # for reading)
    [Byte[]]$BUFFER = New-Object Byte[] $BUFSIZE
    $NUMFILE = 1

    do {
        $Extension = $Extension.TrimStart('.')
        $NEWNAME = "{0}\{1}{2:d4}.{3}" -f $OutPath, $FILENAME, $NUMFILE, $Extension
        $Script:ProgressMessage = "Split {0} of {1} files" -f $Script:StepNumber, $Script:TotalSteps
        Invoke-AutoUpdateProgress_FileUtils
        $Script:StepNumber++
        $COUNT = 0
        $OBJWRITER = $NULL
        [int32]$BYTESREAD = 0
        while (($COUNT -lt $PASSES) -and (($BYTESREAD = $OBJREADER.Read($BUFFER, 0, $BUFFER.Length)) -gt 0))
        {
            Write-Verbose "[Invoke-SplitDataFile] Reading $BYTESREAD bytes"
            if ($AsString) {
                $DataString = [convert]::ToBase64String($BUFFER, 0, $BYTESREAD)
                Write-Verbose "[Invoke-SplitDataFile] WRITING DataString to $NEWNAME"
                Set-Content $NEWNAME $DataString
            } else {
                if (!$OBJWRITER)
                {
                    $OBJWRITER = New-Object System.IO.BinaryWriter ([System.IO.File]::Create($NEWNAME))
                    Write-Verbose " + CREATING $NEWNAME"
                }
                Write-Verbose "[Invoke-SplitDataFile] WRITING $BYTESREAD bytes to $NEWNAME"
                $OBJWRITER.Write($BUFFER, 0, $BYTESREAD)
            }
            $COUNT++
        }
        if (($REMAINDER -gt 0) -and (($BYTESREAD = $OBJREADER.Read($BUFFER, 0, $REMAINDER)) -gt 0))
        {
            Write-Verbose "[Invoke-SplitDataFile] Reading $BYTESREAD bytes"
            if ($AsString) {
                $DataString = [convert]::ToBase64String($BUFFER, 0, $BYTESREAD)
                Write-Verbose "[Invoke-SplitDataFile] WRITING DataString to $NEWNAME"
                Set-Content $NEWNAME $DataString
            } else {
                if (!$OBJWRITER)
                {
                    $OBJWRITER = New-Object System.IO.BinaryWriter ([System.IO.File]::Create($NEWNAME))
                    Write-Verbose " + CREATING $NEWNAME"
                }
                Write-Verbose "[Invoke-SplitDataFile] WRITING $BYTESREAD bytes to $NEWNAME"
                $OBJWRITER.Write($BUFFER, 0, $BYTESREAD)
            }
        }

        if ($OBJWRITER) { $OBJWRITER.Close() }
        if ($BYTESREAD) {
            Write-FileHeader -Path $NEWNAME -PartId $NUMFILE
        }

        ++ $NUMFILE
    } while ($BYTESREAD -gt 0)

    $OBJREADER.Close()
}



function Invoke-CombineSplitFiles {

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Path to the folder containing split parts")]
        [string]$Path,
        [Parameter(Position = 1, Mandatory = $true, HelpMessage = "Path of the recombined file")]
        [string]$Destination,
        [Parameter(Mandatory = $false, HelpMessage = "Encoding type")]
        [ValidateSet('base64', 'raw')]
        [string]$Type = 'raw'
    )
    [bool]$EncodedAsString = ($Type -eq 'base64')

    $SyncStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Script:ProgressTitle = "Combine Split Files"
    $TotalTicks = 0
    $Basename = ''

    $Path = $Path.TrimEnd('\')
    Write-Verbose "Path is $Path"

    $Files = (Get-ChildItem $Path -File -Filter "$Basename*.cpp").FullName
    try {
        $SortedFiles = $Files | Sort-ByFileHeaderId
    } catch {
        Write-Error "$_"
    }
    $FilesCount = $SortedFiles.Count
    $Script:TotalSteps = $FilesCount
    $Script:StepNumber = 1

    if (![System.IO.File]::Exists("$Destination")) {
        New-Item -Path "$Destination" -ItemType File -Force -ErrorAction Ignore | Out-Null
    }

    # Open file stream for output
    $FileStream = [System.IO.File]::Open($Destination, 'Create', 'Write', 'Write') # for writing

    [bool]$RecombinedSuccessfully = $True

    try {
        foreach ($f in $SortedFiles) {
            if (-not (Test-Path -Path $f)) {
                throw "missing file: $f"
            }

            $HeaderData = Remove-FileHeader -Path $f

            if ($EncodedAsString) {
                [string]$Base64String = Get-Content -LiteralPath $f -Raw
                [byte[]]$ReadBytes = [Convert]::FromBase64String($Base64String)
                $FileStream.Write($ReadBytes, 0, $ReadBytes.Length)
            } else {
                [byte[]]$ReadBytes = [System.IO.File]::ReadAllBytes($f)
                $FileStream.Write($ReadBytes, 0, $ReadBytes.Length)
            }

            $Script:ProgressMessage = "Wrote part $Script:StepNumber of $Script:TotalSteps"
            Invoke-AutoUpdateProgress_FileUtils
        }
    } catch {
        $RecombinedSuccessfully = $False
        Write-Error "Error on $f . $_"
    } finally {
        $FileStream.Close()
    }

    if ($RecombinedSuccessfully) {
        Write-Host "Recombined Successfully ! Wrote combined file to $Destination"
    }

}


function Invoke-DoDecode {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Path to the folder containing split parts")]
        [string]$Path,
        [Parameter(Mandatory = $false, HelpMessage = "Encoding type")]
        [ValidateSet('base64', 'raw')]
        [string]$Type = 'raw',
        [Parameter(Mandatory = $false)]
        [string]$Password,
        [Parameter(Mandatory = $false, HelpMessage = "Was the source file encrypted or not")]
        [switch]$Encrypted
    )


    $RootPath = (Resolve-Path -Path "$Path").Path
    $DataCipherPath = Join-Path $RootPath 'data'
    $BinSrcPath = Join-Path $RootPath 'binsrc'
    $Filename = 'bmw_installer_package.rar'
    $EncryptedFilename = $Filename + '.aes'


    $SourcePackagePath = Join-Path $BinSrcPath $Filename
    $RecombinedEncryptedPackagePath = Join-Path $BinSrcPath $EncryptedFilename

    Invoke-CombineSplitFiles $DataCipherPath $RecombinedEncryptedPackagePath -Type $Type

    if ($Encrypted) {
        Write-Host "Invoke-AesBinaryEncryption `"$RecombinedEncryptedPackagePath`" `"$SourcePackagePath`""
        Invoke-AesBinaryEncryption $RecombinedEncryptedPackagePath $SourcePackagePath "$Password" -Mode Decrypt
    }
    if ([System.IO.File]::Exists("$SourcePackagePath")) {
        $Hash2 = (Get-FileHash -Algorithm SHA1 -Path "$SourcePackagePath").Hash
        Write-Host "[$Hash2] $SourcePackagePath" -f DarkCyan
    }
}


function Show-FunctionPath {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    $path = $null
    $folder = $null

    if ($MyInvocation.MyCommand.Path) {
        $path = $MyInvocation.MyCommand.Path
        $folder = Split-Path -Path $path -Parent
        Write-Verbose "This function is in script file: $path"
        Write-Verbose "Folder: $folder"
    }
    elseif ($MyInvocation.MyCommand.Module) {
        $folder = $MyInvocation.MyCommand.Module.ModuleBase
        Write-Verbose "Running from module folder: $folder"
    }
    else {
        $folder = (Get-Location).Path
        Write-Verbose "Running interactively in console. Current folder: $folder"
    }
    $folder
}


function Get-MyHome {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $false, HelpMessage = "Destination folder")]
        [string]$Destination
    )
    $MyHome = ''
    if (Test-Path "$HOME") {
        $MyHome = $HOME
    } elseif (Test-PAth "$("$ENV:HOMEDRIVE" + "$ENV:HOMEPATH")") {
        $MyHome = "$("$ENV:HOMEDRIVE" + "$ENV:HOMEPATH")"
    } else {
        pushd ~
        $MyHome = "$($PWD.PAth)"
        popd
    }
    $MyHome
}

function Get-PackagePath {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $false, HelpMessage = "Destination folder")]
        [string]$Destination
    )
    $MyHome = Get-MyHome
    $PackagePath = ''

    if (Test-Path "$ENV:TEMP") {
        $PackagePath = Join-Path "$ENV:TEMP" "dltool-extracted"

    } elseif (Test-Path "$ENV:LOCALAPPDATA") {
        $MyTemp = Join-Path "$ENV:LOCALAPPDATA" "Temp"
        $DlToolPath = Join-Path $MyTemp "dltool-extracted"
    } else {
        $PackagePath = Join-Path $MyHome "AppData\Local\Temp\dltool-extracted"
    }
    $PackagePath
}

function Install-7zPackage {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $false, HelpMessage = "Destination folder")]
        [string]$Destination,
        [Parameter(Mandatory = $false, HelpMessage = "Destination folder")]
        [switch]$Overwrite
    )

    # Determine installation path
    if (-not $Destination) {
        # Assume Get-PackagePath is defined elsewhere in your environment
        $InstallPath = Get-PackagePath
    } else {
        $InstallPath = $Destination
    }

    if ($Overwrite) {
        Remove-Item -Path "$InstallPath" -Recurse -Force -EA Ignore
    }

    # Construct expected paths
    $Expected7zFolder = Join-Path -Path $InstallPath -ChildPath '7z'
    $Expected7zExe = Join-Path -Path $Expected7zFolder -ChildPath '7z.exe'

    # Check if environment variables already exist
    $envSevenZipPath = [System.Environment]::GetEnvironmentVariable('SevenZipPath', 'User')
    $envSevenZipExe = [System.Environment]::GetEnvironmentVariable('SevenZipExe', 'User')

    $alreadyInstalled = $false

    if ($envSevenZipPath -and $envSevenZipExe) {
        if (Test-Path -Path $envSevenZipExe) {
            Write-Verbose "7-Zip appears to be already installed at:"
            Write-Verbose "SevenZipPath = $envSevenZipPath"
            Write-Verbose "SevenZipExe  = $envSevenZipExe"
            return $envSevenZipExe
        }
    }
    elseif (Test-Path -Path $Expected7zExe) {
        Write-Verbose "7-Zip already exists at $Expected7zExe"

        # Optionally set env vars if they weren't set yet
        if (-not $envSevenZipPath) {
            [System.Environment]::SetEnvironmentVariable('SevenZipPath', $Expected7zFolder, 'User')
        }
        if (-not $envSevenZipExe) {
            [System.Environment]::SetEnvironmentVariable('SevenZipExe', $Expected7zExe, 'User')
        }

        return $Expected7zExe
    }

    # Create the folder if it doesn't exist
    if (-not (Test-Path -Path $InstallPath)) {
        Write-Verbose "Creating installation path: $InstallPath"
        New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null
    }

    # Download the zip file
    $BaseURL = 'https://arsscriptum.github.io/files/7z/7z.zip'
    $ZipFile = Join-Path -Path $InstallPath -ChildPath '7z.zip'

    Write-Verbose "Downloading $BaseURL to $ZipFile ..."
    Invoke-WebRequest -Uri $BaseURL -OutFile $ZipFile -UseBasicParsing

    # Extract the zip file
    Write-Verbose "Extracting $ZipFile to $InstallPath ..."
    Expand-Archive -Path $ZipFile -DestinationPath $InstallPath -Force

    # Remove the zip afterwards
    Remove-Item -Path $ZipFile -Force

    # Resolve paths again
    $SevenZipPath = Resolve-Path -Path $Expected7zFolder -ErrorAction SilentlyContinue
    if (-not $SevenZipPath) {
        throw "7z folder not found in $InstallPath after extraction."
    }

    $SevenZipExe = Resolve-Path -Path $Expected7zExe -ErrorAction SilentlyContinue
    if (-not $SevenZipExe) {
        throw "7z.exe not found at expected location: $Expected7zFolder"
    }

    # Set environment variables
    [System.Environment]::SetEnvironmentVariable('SevenZipPath', $SevenZipPath.Path, 'User')
    [System.Environment]::SetEnvironmentVariable('SevenZipExe', $SevenZipExe.Path, 'User')

    Write-Verbose "7-Zip successfully installed."
    Write-Verbose "SevenZipPath = $($SevenZipPath.Path)"
    Write-Verbose "SevenZipExe  = $($SevenZipExe.Path)"

    # Return path to 7z.exe
    return $SevenZipExe.Path
}

function Get-My7zExe {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()


    # Construct expected paths
    $InstallPath = Get-PackagePath
    $Expected7zFolder = Join-Path -Path $InstallPath -ChildPath '7z'
    $Expected7zExe = Join-Path -Path $Expected7zFolder -ChildPath '7z.exe'

    # Check if environment variables already exist
    $envSevenZipPath = [System.Environment]::GetEnvironmentVariable('SevenZipPath', 'User')
    $envSevenZipExe = [System.Environment]::GetEnvironmentVariable('SevenZipExe', 'User')

    if ($envSevenZipPath -and $envSevenZipExe) {
        if (Test-Path -Path $envSevenZipExe) {
            return $envSevenZipExe
        }
    }
    elseif (Test-Path -Path $Expected7zExe) {
        return $Expected7zExe
    }
    return $NUll
}


function Test-7zDetect {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    [string]$SevenZipExe = Get-My7zExe
    Write-Host "SevenZipExe $SevenZipExe"
    # Validate 7z.exe exists

    if (-not (Test-Path $SevenZipExe)) {
        throw "7z.exe not found. Please install 7-Zip or specify the full path via -SevenZipExe."
    }

    $SevenZipExe
}

[bool]$LogToFileEnabled = $True
[bool]$LogToConsoleEnabled = $True
$ENV:DownloadLogsToFile = if ($LogToFileEnabled) { 1 } else { 0 }
$ENV:DownloadLogsToConsole = if ($LogToConsoleEnabled) { 1 } else { 0 }

$IsLegacy = ($PSVersionTable.PSVersion.Major -eq 5)
if ($IsLegacy) {
    # No need to load mscorlib; but if you want to:
    Add-Type -AssemblyName "mscorlib"
}

function Convert-Bytes {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [double]$Bytes,

        [string]$Suffix = ""
    )

    switch ($Bytes) {
        { $_ -ge 1TB } { return "{0:N2} TB$Suffix" -f ($Bytes / 1TB) }
        { $_ -ge 1GB } { return "{0:N2} GB$Suffix" -f ($Bytes / 1GB) }
        { $_ -ge 1MB } { return "{0:N2} MB$Suffix" -f ($Bytes / 1MB) }
        { $_ -ge 1KB } { return "{0:N2} KB$Suffix" -f ($Bytes / 1KB) }
        default { return "{0:N2} B$Suffix" -f $Bytes }
    }
}

function Update-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [pscustomobject]$JobStats
    )
    process {
        $CurrentStats = Get-GlobalJobsStats

        $jobCount = $CurrentStats.TotalJobs
        $minTime = $CurrentStats.MinTransferTimeSec
        $maxTime = $CurrentStats.MaxTransferTimeSec

        $totalFiles = $CurrentStats.TotalFilesTransferred
        $totalBytes = $CurrentStats.TotalBytesTransferred
        $totalTime = $CurrentStats.CurrentTotalTransferTime

        $totalFiles += $JobStats.TotalFiles
        $totalBytes += $JobStats.DownloadSize
        $totalTime += $JobStats.DownloadTime

        $globalAverageSpeedBps = [math]::Round($totalBytes / $totalTime, 2)



        # Human readable speed
        $globalAverageHumanSpeed = Convert-Bytes -Bytes $globalAverageSpeedBps -Suffix "/s"

        if (($minTime -eq $null) -or ($JobStats.DownloadTime -lt $minTime)) {
            $minTime = $JobStats.DownloadTime
        }
        if (($maxTime -eq $null) -or ($JobStats.DownloadTime -gt $maxTime)) {
            $maxTime = $JobStats.DownloadTime
        }

        $jobCount = $jobCount + 1

        $avgTime = if ($jobCount -gt 0) { $totalTime / $jobCount } else { 0 }

        $stats = [pscustomobject]@{
            TotalJobs = $jobCount -as [uint32]
            TotalFilesTransferred = $totalFiles -as [uint32]
            TotalBytesTransferred = $totalBytes -as [uint32]
            HumanReadableTotalSize = Convert-Bytes -Bytes $totalBytes
            AverageSpeed_Bps = [math]::Round($globalAverageSpeedBps, 2)
            HumanReadableAvgSpeed = Convert-Bytes -Bytes $globalAverageSpeedBps -Suffix "/s"
            AverageTransferTimeSec = [math]::Round($avgTime, 2)
            MinTransferTimeSec = $minTime
            MaxTransferTimeSec = $maxTime
            CurrentTotalTransferTime = $totalTime -as [double]
        }

        # Save stats as JSON to temp file
        $path = Join-Path $env:TEMP "GlobalJobStats.json"
        $stats | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
    }
}

function Reset-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $stats = [pscustomobject]@{
        TotalJobs = 0
        TotalFilesTransferred = 0
        TotalBytesTransferred = 0
        HumanReadableTotalSize = Convert-Bytes -Bytes 0
        AverageSpeed_Bps = [math]::Round(0, 2)
        HumanReadableAvgSpeed = Convert-Bytes -Bytes 0 -Suffix "/s"
        AverageTransferTimeSec = [math]::Round(0, 2)
        MinTransferTimeSec = 0
        MaxTransferTimeSec = 0
        CurrentTotalTransferTime = 0
    }

    # Save stats as JSON to temp file
    $path = Join-Path $env:TEMP "GlobalJobStats.json"
    $stats | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
}


function Get-GlobalJobsStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $path = Join-Path $env:TEMP "GlobalJobStats.json"

    if (-not (Test-Path $path)) {
        Write-Warning "No global job stats file found at $path."
        return $null
    }

    $json = Get-Content -Path $path -Raw -Encoding UTF8
    $stats = $json | ConvertFrom-Json
    return $stats
}


function Measure-JobStats {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "BITS Job Id (GUID)")]
        [guid]$JobId
    )

    # Fetch the job object
    $job = Get-BitsTransfer -JobId $JobId -ErrorAction Stop

    # --- Sanity checks ---
    if ($job.JobState -ne 'Transferred') {
        throw "BITS job state is not 'Transferred'. Current state: $($job.JobState)"
    }

    if ($job.BytesTransferred -ne $job.BytesTotal) {
        throw "BytesTransferred ($($job.BytesTransferred)) does not equal BytesTotal ($($job.BytesTotal))."
    }

    if ($job.TransferCompletionTime -le $job.CreationTime) {
        throw "TransferCompletionTime is not after CreationTime."
    }

    # Compute time delta in seconds

    [double]$durationSec = (New-TimeSpan -Start $jobptr.CreationTime -End $jobptr.TransferCompletionTime).TotalSeconds
    if ($durationSec -le 0) {
        throw "Calculated download time is zero or negative. Cannot compute speed."
    }
    Write-Verbose "[Measure-JobStats] BytesTransferred -> $($job.BytesTransferred)"
    Write-Verbose "[Measure-JobStats] durationSec -> $durationSec"

    # Compute speed
    $speedBps = [math]::Round($job.BytesTransferred / $durationSec, 2)

    # Human readable speed
    $humanSpeed = Convert-Bytes -Bytes $speedBps -Suffix "/s"

    # File path(s) - BITS supports multiple files in a job
    $localFilePath = $job.DisplayName

    $numberFilesTransfered = $job.FilesTransferred


    # Return stats
    $o = [pscustomobject]@{
        Speed = $speedBps -as [uint32]
        HumanReadableSpeed = $humanSpeed
        DownloadTime = $durationSec -as [double]
        DownloadSize = $job.BytesTransferred -as [uint32]
        FilePath = $localFilePath
        TotalFiles = $numberFilesTransfered -as [uint32]
        TotalJobs = 1
    }

    return $o
}

function Write-JobTransferStatsLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline)]
        [pscustomobject]$JobStats
    )

    process {
        if ($null -eq $JobStats) {
            Write-Host "No job statistics provided." -ForegroundColor Yellow
            return
        }

        Write-Host "===== BITS JOB TRANSFER STATISTICS =====" -ForegroundColor Cyan
        Write-Host ("File(s) Path           : {0}" -f $JobStats.FilePath) -ForegroundColor White
        Write-Host ("Download Size (bytes)  : {0}" -f $JobStats.DownloadSize) -ForegroundColor White
        Write-Host ("Download Time (sec)    : {0}" -f $JobStats.DownloadTime) -ForegroundColor White
        Write-Host ("Speed (Bytes/sec)      : {0}" -f $JobStats.Speed) -ForegroundColor White
        Write-Host ("Human-Readable Speed   : {0}" -f $JobStats.HumanReadableSpeed) -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
    }
}


function Write-GlobalTransferStatsLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    $stats = Get-GlobalJobsStats

    if ($null -eq $stats) {
        Write-Host "No global transfer statistics available." -ForegroundColor Yellow
        return
    }

    Write-Host "===== GLOBAL BITS TRANSFER STATISTICS =====" -ForegroundColor Cyan
    Write-Host ("Total Jobs              : {0}" -f $stats.TotalJobs) -ForegroundColor White
    Write-Host ("Total Transfer Time     : {0} seconds" -f $stats.CurrentTotalTransferTime) -ForegroundColor White
    Write-Host ("Total Files Transferred : {0}" -f $stats.TotalFilesTransferred) -ForegroundColor White
    Write-Host ("Total Bytes Transferred : {0} bytes" -f $stats.TotalBytesTransferred) -ForegroundColor White
    Write-Host ("Human-Readable Size     : {0}" -f $stats.HumanReadableTotalSize) -ForegroundColor Green
    Write-Host ("Average Speed (B/s)     : {0}" -f $stats.AverageSpeed_Bps) -ForegroundColor White
    Write-Host ("Human-Readable Avg Spd  : {0}" -f $stats.HumanReadableAvgSpeed) -ForegroundColor Green
    Write-Host ("Average Transfer Time   : {0} seconds" -f $stats.AverageTransferTimeSec) -ForegroundColor White
    Write-Host ("Min Transfer Time       : {0} seconds" -f $stats.MinTransferTimeSec) -ForegroundColor White
    Write-Host ("Max Transfer Time       : {0} seconds" -f $stats.MaxTransferTimeSec) -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan
}


function Write-DownloadLogs {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        $Object,

        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor,
        [switch]$NoNewline
    )

    # Compose the text exactly as Write-Host would output it
    $text = -join ($Object | ForEach-Object {
            if ($_ -is [System.Management.Automation.PSObject]) {
                $_.ToString()
            } else {
                $_
            }
        })

    # Write to console if enabled
    if ($ENV:DownloadLogsToConsole -eq "1") {
        $params = @{}
        if ($PSBoundParameters.ContainsKey("ForegroundColor")) {
            $params.ForegroundColor = $ForegroundColor
        }
        if ($PSBoundParameters.ContainsKey("BackgroundColor")) {
            $params.BackgroundColor = $BackgroundColor
        }
        if ($NoNewline) {
            $params.NoNewline = $true
        }
        Write-Host @params $text
    }

    # Write to file if enabled
    if ($ENV:DownloadLogsToFile -eq "1") {
        $logFile = Join-Path -Path $ENV:TEMP -ChildPath "DownloadLogs.log"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logLine = "$timestamp $text"
        Add-Content -Path $logFile -Value $logLine
    }
}


#  “define” KeyValuePair in PowerShell

function New-TmpDirectory {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param()

    # Get system temp path
    $tempPath = [System.IO.Path]::GetTempPath()

    # Generate a unique folder name
    $folderName = [System.IO.Path]::GetRandomFileName()

    # Combine to full path
    $fullPath = Join-Path -Path $tempPath -ChildPath $folderName

    # Create the directory
    $item = New-Item -ItemType Directory -Path $fullPath -Force

    return $item
}
function New-TmpFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param()

    # Use .NET method
    $tempFile = [System.IO.Path]::GetTempFileName()

    $item = New-Item -ItemType File -Path $tempFile -Force

    return $item
}

function Save-DataFiles {

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Base URL where the data is located")]
        [string]$BaseURL,
        [Parameter(Position = 1, Mandatory = $true, HelpMessage = "list of files, appended to baseurl")]
        [string[]]$Files,
        [Parameter(Position = 2, Mandatory = $false, HelpMessage = "Path to the folder containing split parts")]
        [string]$Destination,
        [Parameter(Mandatory = $false, HelpMessage = "how much files to download per batch")]
        [int]$BatchSize = 35,
        [Parameter(Mandatory = $false, HelpMessage = "transfer priority, Foreground is 10x faster")]
        [ValidateSet('Foreground', 'High', 'Low', 'Normal')]
        [string]$Priority = 'Foreground'
    )

    # Define a shorthand type alias
    import-module BitsTransfer -Force
    get-BitsTransfer | Remove-BitsTransfer -Confirm:$false


    # ------------------------------------------------------------
    # Dictionary: $UrlListType
    # Purpose  : Keep track of download URLs and whether each has
    #            already been processed (Start-BitsJob) or not.
    #
    # Type     : [hashtable] where
    #              - Key   = [string] URL to download
    #              - Value = [bool]   $true  = already processed
    #                                 $false = not yet processed
    # +-----------------------------------------------+----------+
    # |                    URL                        | Processed|
    # +-----------------------------------------------+----------+
    # | http://example.com/file1.zip                  | false    |
    # | http://example.com/file2.zip                  | true     |
    # +-----------------------------------------------+----------+
    #
    # Meaning:
    #   - file1.zip has NOT been downloaded yet.
    #   - file2.zip has ALREADY been downloaded.
    # ------------------------------------------------------------
    # Define generic List type
    $UrlListType = [System.Collections.Generic.List[[System.Collections.Generic.KeyValuePair[string,bool]]]]::new()
    # Create an empty list of URLs
    $UrlList = $UrlListType::new()
    $BitsJobListType = [System.Collections.Generic.List[string]]::new()

    $ShouldExit = $False

    $UrlList = $UrlListType::new()
    Write-DownloadLogs "[Save-DataFiles] Initialize Files Count $($Files.Count) " -f DarkYellow

    foreach ($f in $Files) {
        $RemoteUrl = '{0}/{1}' -f $BaseURL, $f
        Write-DownloadLogs "  add $RemoteUrl " -f DarkYellow
        $UrlList.Add([System.Collections.Generic.KeyValuePair[string,bool]]::new($RemoteUrl, $False))
    }

    $TotalNumberFiles = $UrlList.Count

    # How many jobs you want to run concurrently
    $completedDownloadsPath = (New-TmpFile).FullName
    $ENV:CompletedDownloadsPath = $completedDownloadsPath
    $runningJobs = $BitsJobListType::new()
    $completedJobs = $BitsJobListType::new()
    $suspendedJobs = $BitsJobListType::new()
    function Get-CompletedDownloads {
        [CmdletBinding()]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        if (-not (Test-Path -LiteralPath $path)) {
            return 0
        }

        $content = Get-Content -LiteralPath $path -ErrorAction Stop
        if ([int]::TryParse($content, [ref]$null)) {
            return [int]$content
        } else {
            throw "Invalid data in $path. Expected an integer."
        }
    }

    function Add-CompletedDownloads {
        [CmdletBinding(SupportsShouldProcess)]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        $current = 0
        if (Test-Path -LiteralPath $path) {
            $current = Get-CompletedDownloads
        }

        $new = $current + 1

        if ($PSCmdlet.ShouldProcess("File: $path", "Write value $new")) {
            Set-Content -LiteralPath $path -Value $new
        }
    }

    function Reset-CompletedDownloads {
        [CmdletBinding(SupportsShouldProcess)]
        param()

        $path = $ENV:CompletedDownloadsPath

        if (-not $path) {
            throw "Environment variable 'CompletedDownloadsPath' is not set."
        }

        if ($PSCmdlet.ShouldProcess("File: $path", "Reset to 0")) {
            Set-Content -LiteralPath $path -Value '0'
        }
    }

    Reset-CompletedDownloads
    Reset-GlobalJobsStats
    $TotalStartedJobs = 0
    $StartedJobsInBatch = 0
    $ProcessedBatch = 0
    $EstimatedNumberOfBatches = [math]::Round(($TotalNumberFiles / $BatchSize), [System.MidpointRounding]::ToPositiveInfinity)

    $StatsEmpty = $True

    while (!$ShouldExit) {
        Start-Sleep -Milliseconds 10

        # 1. Request to download X number of files ($BatchSize)
        # If we don't have any more active jobs, that means we are ready to download another batch of files.

        if ($runningJobs.Count -eq 0) {

            #if ($StatsEmpty -eq $False) {
            #    Write-GlobalTransferStatsLog
            #}
            $ProcessedBatch = $ProcessedBatch + 1
            $StartedJobsInBatch = 0
            Write-DownloadLogs " ★★★ Start Download in a new transfer group. This is Batch no $ProcessedBatch ★★★" -f Red

            $StartedAllBitsJobsInBatch = $False
            $globalId = $TotalStartedJobs

            # Looping until we started all the bits jobs for this batch

            while ($StartedAllBitsJobsInBatch -eq $False) {
                # Variables to track
                $nextUnprocessedUrlIndex = 0
                $foundNextUrl = $false

                while (($nextUnprocessedUrlIndex -lt $UrlList.Count) -and (-not $foundNextUrl)) {
                    if (-not $UrlList[$nextUnprocessedUrlIndex].Value) {
                        # Found an unprocessed URL
                        $foundNextUrl = $true
                        Write-Verbose "Next unprocessed URL index is: $nextUnprocessedUrlIndex"
                        Write-Verbose "URL: $($UrlList[$nextUnprocessedUrlIndex].Key)"
                    }
                    else {
                        # Move to next index
                        $nextUnprocessedUrlIndex++
                    }
                }

                if (-not $foundNextUrl) {
                    Write-Verbose "No unprocessed URLs found."
                    $StartedAllBitsJobsInBatch = $True
                    break
                }

                $IsReady = $u.Value -eq $False

                if (-not $UrlList[$nextUnprocessedUrlIndex].Value) {
                    $url = $UrlList[$nextUnprocessedUrlIndex].Key
                    # create a local save path for the file to download (remove the base url, and make the relative path: http://example.com/data/file1.zip -> data\file1.zip)
                    $fileName = $url.Replace("$BaseURL/", '').Replace('/', '\')
                    $dest = Join-Path $Destination $fileName
                    $globalId = $globalId + 1
                    Write-Verbose "[$globalId] downloading $url and saving to $dest"

                    $desc = '{0}/{1}|{2}/{3}|{4}|{5}' -f $ProcessedBatch, $EstimatedNumberOfBatches, $globalId, $TotalNumberFiles, $dest, $url
                    Write-Verbose "$desc"
                    # Safety checks
                    if (-not $url) { throw "URL is null or empty!" }
                    if (-not $dest) { throw "Destination path is null or empty!" }
                    if (-not $desc) { $desc = "BITS Transfer" }

                    # Ensure the target folder exists
                    $destDir = Split-Path -Parent $dest
                    if (-not (Test-Path $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                    }

                    $arguments = @{
                        Source = "$url"
                        Description = "$desc"
                        Destination = "$dest"
                        TransferType = "Download"
                        Asynchronous = $True
                        DisplayName = "$dest"
                        Priority = $Priority
                        RetryTimeout = 60
                        RetryInterval = 60
                    }
                    $StartedJobsInBatch = $StartedJobsInBatch + 1

                    $res = Start-BitsTransfer @arguments
                    $SmallGuid = $res.JobId.GUID.Substring(14, 4)
                    $TotalStartedJobs++
                    $NewJobGuid = $res.JobId.GUID
                    $log = " 🡲 Start-BitsTransfer [{0}] - {1} out of {2} in this batch, {3} out of {4} in total." -f $SmallGuid, $StartedJobsInBatch, $BatchSize, $globalId, $TotalNumberFiles
                    Write-DownloadLogs "$log" -f Blue
                    $UrlList[$nextUnprocessedUrlIndex] = [System.Collections.Generic.KeyValuePair[string,bool]]::new($url, $True)
                    $runningJobs.Add($NewJobGuid)
                    if (($runningJobs.Count) -ge $BatchSize) {
                        $StartedAllBitsJobsInBatch = $True
                    }
                }
            }
        }

        # 2. Wait until all the active transfers are completed ($runningJobs moved to -> $completedJobs)

        Write-DownloadLogs "⚠️ There are currently $($runningJobs.Count) active transfers... Waiting for jobs to complete." -f DarkGreen
        while ($runningJobs.Count -gt 0) {
            # give some resources for the transfer thread
            Start-Sleep -Milliseconds 10
            Get-BitsTransfer | % {
                $JobGuid = $_.JobId.GUID
                $jobptr = Get-BitsTransfer -JobId $JobGuid -ErrorAction Stop
                $state = $jobptr.JobState

                if ($state -eq 'Transferred') {
                    $SmallGuid = $jobptr.JobId.GUID.Substring(14, 4)
                    $DescriptionData = $jobptr.Description.Split('|')
                    $jobBatchId = $DescriptionData[0]
                    $jobGlobalId = $DescriptionData[1]

                    Write-Verbose "JOB $JobGuid is Transferred, Measure-JobStats..."
                    $JobStats = Measure-JobStats -JobId $JobGuid
                    #Write-JobTransferStatsLog $JobStats
                    Update-GlobalJobsStats $JobStats

                    Add-CompletedDownloads
                    $log = '✔️ Job [{0}] COMPLETED. BATCH {1} TOTAL {2}' -f $SmallGuid, $jobBatchId, $jobGlobalId
                    Write-DownloadLogs "$log" -f Gray
                    $jobptr | Complete-BitsTransfer
                    $completedJobs.Add($JobGuid)
                }
                foreach ($j in $completedJobs) {
                    [void]$runningJobs.Remove($j)
                }
            }
        }

        $CompletedDownloads = Get-CompletedDownloads
        if ($CompletedDownloads -ge $TotalNumberFiles) {
            $ShouldExit = $True
        } elseif ($ErrorOccured) {
            $ShouldExit = $True
        }
    }

    Write-GlobalTransferStatsLog

}



function Show-ExceptionDetails {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$Record,
        [Parameter(Mandatory = $false)]
        [switch]$ShowStack
    )
    $formatstring = "{0}`n{1}"
    $fields = $Record.FullyQualifiedErrorId, $Record.Exception.ToString()
    $ExceptMsg = ($formatstring -f $fields)
    $Stack = $Record.ScriptStackTrace
    Write-Host "`n[ERROR] -> " -NoNewline -ForegroundColor DarkRed;
    Write-Host "$ExceptMsg`n`n" -ForegroundColor DarkYellow
    if ($ShowStack) {
        Write-Host "--stack begin--" -ForegroundColor DarkGreen
        Write-Host "$Stack" -ForegroundColor Gray
        Write-Host "--stack end--`n" -ForegroundColor DarkGreen
    }
    if ((Get-Variable -Name 'ShowExceptionDetailsTextBox' -Scope Global -ErrorAction Ignore -ValueOnly) -eq 1) {
        Show-MessageBoxException $ExceptMsg $Stack
    }


}


function Invoke-AesBinaryEncryption {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string]$InputFile,
        [Parameter(Position = 1, Mandatory = $true)]
        [string]$OutputFile,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$Password,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Encrypt', 'Decrypt')]
        [string]$Mode,
        [Parameter(Mandatory = $false)]
        [switch]$TextMode
    )

    begin {
        $shaManaged = New-Object System.Security.Cryptography.SHA256Managed
        $aesManaged = New-Object System.Security.Cryptography.AesManaged
        $aesManaged.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aesManaged.Padding = [System.Security.Cryptography.PaddingMode]::Zeros
        $aesManaged.BlockSize = 128
        $aesManaged.KeySize = 256
    }

    process {
        try {
            $aesManaged.Key = $shaManaged.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Password))

            switch ($Mode) {
                'Encrypt' {

                    $File = Get-Item -Path $InputFile -ErrorAction SilentlyContinue
                    if (!$File.FullName) {
                        Write-Error -Message "File not found!"
                        break
                    }

                    $plainBytes = [System.IO.File]::ReadAllBytes($File.FullName)

                    $encryptor = $aesManaged.CreateEncryptor()
                    $encryptedBytes = $encryptor.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
                    $encryptedBytes = $aesManaged.IV + $encryptedBytes
                    $aesManaged.Dispose()

                    if ($TextMode) {
                        Write-Host "Writing TEXT data in $OutputFile..."
                        [System.Convert]::ToBase64String($encryptedBytes) | Set-Content -Path $OutputFile -Encoding ascii -Force
                    } else {
                        [System.IO.File]::WriteAllBytes($OutputFile, $encryptedBytes)
                        (Get-Item $OutputFile).LastWriteTime = $File.LastWriteTime
                        Write-Host "File encrypted to $OutputFile"

                    }
                }

                'Decrypt' {


                    $File = Get-Item -Path $InputFile -ErrorAction SilentlyContinue
                    if (!$File.FullName) {
                        Write-Error -Message "File not found!"
                        break
                    }

                    $tmpBytes = [System.IO.File]::ReadAllBytes($File.FullName)

                    if ($TextMode) {
                        $content = Get-Content -Path $File.FullName -Encoding ascii -Force
                        $cipherBytes = [System.Convert]::FromBase64String($content)
                    } else {
                        $cipherBytes = $tmpBytes;
                    }

                    $aesManaged.IV = $cipherBytes[0..15]
                    $decryptor = $aesManaged.CreateDecryptor()
                    $decryptedBytes = $decryptor.TransformFinalBlock($cipherBytes, 16, $cipherBytes.Length - 16)
                    $aesManaged.Dispose()

                    if ($TextMode) {
                        Write-Host "Writing TEXT data in $OutputFile..."
                        [System.Text.Encoding]::ASCII.GetString($decryptedBytes).Trim([char]0) | Set-Content -Path $OutputFile -Encoding ascii -Force
                    } else {
                        [System.IO.File]::WriteAllBytes($OutputFile, $decryptedBytes)
                        (Get-Item $OutputFile).LastWriteTime = $File.LastWriteTime
                        Write-Host "File decrypted to $OutputFile"
                    }
                }
            }
        } catch {
            Write-Error $_
        }
    }

    end {
        $shaManaged.Dispose()
        $aesManaged.Dispose()
    }
}

function Get-StringHash {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "String to hash.")]
        [string]$String,

        [Parameter(Mandatory = $false, HelpMessage = "Hash algorithm (MD5, SHA1, SHA256, SHA384, SHA512).")]
        [ValidateSet("MD5", "SHA1", "SHA256", "SHA384", "SHA512")]
        [string]$Algorithm = "SHA256"
    )

    # Convert string to byte array
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)

    # Create hasher instance
    switch ($Algorithm) {
        "MD5" { $hasher = [System.Security.Cryptography.MD5]::Create() }
        "SHA1" { $hasher = [System.Security.Cryptography.SHA1]::Create() }
        "SHA256" { $hasher = [System.Security.Cryptography.SHA256]::Create() }
        "SHA384" { $hasher = [System.Security.Cryptography.SHA384]::Create() }
        "SHA512" { $hasher = [System.Security.Cryptography.SHA512]::Create() }
        default { throw "Unsupported algorithm: $Algorithm" }
    }

    try {
        # Compute the hash
        $hashBytes = $hasher.ComputeHash($bytes)

        # Convert bytes to hex string
        $hashString = [BitConverter]::ToString($hashBytes) -replace '-', ''

        return $hashString
    }
    finally {
        # Clean up
        if ($hasher -is [System.IDisposable]) {
            $hasher.Dispose()
        }
    }
}



function Test-IsValidPassword {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = "String to hash.")]
        [string]$Password
    )
    #v2 pass
    $CorrectSha256ofPassword = '48E3FB372EED52ECB014AF5EAEC2F491A466D11C71507A18DA86878150B6E214'
    $Hash = Get-StringHash $Password
    if ($Hash -ne $CorrectSha256ofPassword) {
        return $False
    }
    return $True

}

function Get-PasswordWindow {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    Add-Type -AssemblyName System.Windows.Forms

    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Enter Password"
    $form.Size = New-Object System.Drawing.Size (300, 150)
    $form.StartPosition = "CenterScreen"

    # Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Password:"
    $label.Left = 10
    $label.Top = 20
    $label.AutoSize = $true
    $form.Controls.Add($label)

    # TextBox
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Left = 90
    $textBox.Top = 18
    $textBox.Width = 180
    $textBox.UseSystemPasswordChar = $true
    $form.Controls.Add($textBox)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Left = 90
    $okButton.Top = 60
    $okButton.Width = 80
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Left = 190
    $cancelButton.Top = 60
    $cancelButton.Width = 80
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Show the form
    $dialogResult = $form.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}


function Test-DriveFreeSpace220MB {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Full path to check (e.g. C:\SomeFolder)')]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )


    # Validate path exists
    if (-not (Test-Path -Path $Path)) {
        Write-Error "The specified path '$Path' does not exist."
        return $false
    }

    try {
        # Get drive root from the path
        $fullPath = (Resolve-Path $Path).Path
        $driveRoot = [System.IO.Path]::GetPathRoot($fullPath)

        # Use DriveInfo instead of WMI
        $driveInfo = New-Object System.IO.DriveInfo ($driveRoot)

        if ($driveInfo -eq $null) {
            Write-Error "Unable to retrieve drive information for $driveRoot."
            return $false
        }

        # Free space in MB
        $freeMB = [math]::Round($driveInfo.AvailableFreeSpace / 1MB, 2)


        return ($freeMB -ge 220)
    }
    catch {
        Write-Error "Error while checking free space: $_"
        return $false
    }
}


function Expand-RarFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "RAR file path")]
        [string]$RarFile,

        [Parameter(Mandatory = $true, HelpMessage = "Destination folder")]
        [string]$Destination
    )


    Install-7zPackage

    [string]$SevenZipExe = Get-My7zExe
    [string]$SevenZipPath = (Get-Item "$SevenZipExe").DirectoryName

    # Validate 7z.exe exists

    if (-not (Test-Path $SevenZipExe)) {
        throw "7z.exe not found. Please install 7-Zip or specify the full path via -SevenZipExe."
    }

    # Ensure destination folder exists
    if (-not (Test-Path $Destination)) {
        New-Item -Path $Destination -ItemType Directory | Out-Null
    }

    $arguments = @(
        "x" # extract
        "`"$RarFile`"" # rar file path
        "-o`"$Destination`"" # output folder
        "-y" # auto-confirm
    )

    Write-Verbose "Running: $SevenZipExe $($arguments -join ' ')"

    $process = Start-Process -FilePath "$SevenZipExe" -ArgumentList $arguments -Wait -NoNewWindow -Passthru -WorkingDirectory "$SevenZipPath"

    if ($process.ExitCode -ne 0) {
        throw "Extraction failed. 7z.exe exited with code $($process.ExitCode)."
    }
}

function Show-DownloadToolMainDialog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [string]$DownloadMethod = 'http'
    $IsLegacy = ($PSVersionTable.PSVersion.Major -eq 5)
    if ($IsLegacy) {
        Add-Type -AssemblyName "mscorlib"
    }

    # Create Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BMW Advanced Tools - Download Helper"
    $form.Size = New-Object System.Drawing.Size (500, 400)
    $form.StartPosition = "CenterScreen"

    # === GroupBox: Paths ===
    $groupPaths = New-Object System.Windows.Forms.GroupBox
    $groupPaths.Text = "Paths"
    $groupPaths.Size = New-Object System.Drawing.Size (460, 150)
    $groupPaths.Location = New-Object System.Drawing.Point (10, 10)

    # Label: Temp Path
    $labelTmp = New-Object System.Windows.Forms.Label
    $labelTmp.Text = "Temp Path:"
    $labelTmp.Size = New-Object System.Drawing.Size (70, 20)
    $labelTmp.Location = New-Object System.Drawing.Point (10, 25)
    $groupPaths.Controls.Add($labelTmp)

    # TextBox: Temp Path
    $textTmpPath = New-Object System.Windows.Forms.TextBox
    $textTmpPath.Size = New-Object System.Drawing.Size (300, 20)
    $textTmpPath.Location = New-Object System.Drawing.Point (80, 25)
    $textTmpPath.Text = "$ENV:TEMP\BMWAdvancedTools"
    $textTmpPath.Enabled = $False
    $groupPaths.Controls.Add($textTmpPath)

    # Button: Browse Temp
    $btnBrowseTmp = New-Object System.Windows.Forms.Button
    $btnBrowseTmp.Text = "..."
    $btnBrowseTmp.Size = New-Object System.Drawing.Size (50, 20)
    $btnBrowseTmp.Location = New-Object System.Drawing.Point (390, 25)

    $groupPaths.Controls.Add($btnBrowseTmp)

    # Label: Destination Path
    $labelDest = New-Object System.Windows.Forms.Label
    $labelDest.Text = "Path:"
    $labelDest.Size = New-Object System.Drawing.Size (70, 20)
    $labelDest.Location = New-Object System.Drawing.Point (10, 60)
    $groupPaths.Controls.Add($labelDest)

    # TextBox: Destination Path
    $textDestPath = New-Object System.Windows.Forms.TextBox
    $textDestPath.Size = New-Object System.Drawing.Size (300, 20)
    $textDestPath.Enabled = $False
    $textDestPath.Location = New-Object System.Drawing.Point (80, 60)
    $groupPaths.Controls.Add($textDestPath)

    # Button: Browse Destination
    $btnBrowseDest = New-Object System.Windows.Forms.Button
    $btnBrowseDest.Text = "..."
    $btnBrowseDest.Size = New-Object System.Drawing.Size (50, 20)
    $btnBrowseDest.Location = New-Object System.Drawing.Point (390, 60)
    $groupPaths.Controls.Add($btnBrowseDest)

    # Button: GO
    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = "GO"
    $btnGo.Size = New-Object System.Drawing.Size (440, 30)
    $btnGo.Location = New-Object System.Drawing.Point (10, 100)
    $groupPaths.Controls.Add($btnGo)

    $form.Controls.Add($groupPaths)

    # === GroupBox: Status ===
    $groupStatus = New-Object System.Windows.Forms.GroupBox
    $groupStatus.Text = "Status"
    $groupStatus.Size = New-Object System.Drawing.Size (460, 120)
    $groupStatus.Location = New-Object System.Drawing.Point (10, 180)

    # Label: State
    $labelState = New-Object System.Windows.Forms.Label
    $labelState.Text = "Ready"
    $labelState.Size = New-Object System.Drawing.Size (440, 20)
    $labelState.Location = New-Object System.Drawing.Point (10, 30)
    $groupStatus.Controls.Add($labelState)

    # ProgressBar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size (440, 25)
    $progressBar.Location = New-Object System.Drawing.Point (10, 60)
    $progressBar.Style = 'Continuous'
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $groupStatus.Controls.Add($progressBar)

    $form.Controls.Add($groupStatus)


    # === Browse Buttons Logic ===
    $btnBrowseTmp.Add_Click({
            $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($folderDlg.ShowDialog() -eq 'OK') {
                $textTmpPath.Text = $folderDlg.SelectedPath
            }
        })


    $btnBrowseDest.Add_Click({
            $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($folderDlg.ShowDialog() -eq 'OK') {
                $selPath = $folderDlg.SelectedPath
                $HasEnough = Test-DriveFreeSpace220MB -Path $selPath
                if (-not $HasEnough) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "The selected drive does not have at least 220MB of free space.",
                        "Insufficient Space",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                    return # Exit without setting the path
                }
                $textDestPath.Text = $selPath
            }
        })

    # === GO Button Logic ===
    $btnGo.Add_Click({
            $labelState.Text = "Processing..."
            $progressBar.Value = 0
            $form.Refresh()

            # Get paths from textboxes
            $tempPath = $textTmpPath.Text
            $destPath = $textDestPath.Text

            $password = Get-PasswordWindow
            $IsValidPass = Test-IsValidPassword $password
            if (-not $IsValidPass) {
                [System.Windows.Forms.MessageBox]::Show("INVALID PASSWORD!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $labelState.Text = "Error"
                return
            }
            if (-not (Test-Path $tempPath)) {
                [System.IO.Directory]::CreateDirectory($tempPath)
                #[System.Windows.Forms.MessageBox]::Show("Temporary path does not exist!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                #$labelState.Text = "Error"
                #return
            }

            try {
                $UseOnlineScript = $False
                $btnGo.Enabled = $False
                $DestPath = $textDestPath.Text
                if (-not (Test-Path $DestPath)) {
                    throw "$DestPath not found."
                }
                $labelState.Text = "Cloning repository..."
                $progressBar.Value = 20
                $form.Refresh()

                # Clone repo into temp path
                $clonePath = Join-Path -Path $tempPath -ChildPath "my.special.tools"
                if (Test-Path $clonePath) {
                    Remove-Item -Path $clonePath -Recurse -Force
                }
                if ($DownloadMethod -eq 'clone') {
                    $repoUrl = "https://github.com/arsscriptum/advanced-tools.git"
                    $gitCmd = "git clone $repoUrl `"$clonePath`""
                    $gitResult = & cmd /c $gitCmd
                } elseif ($DownloadMethod -eq 'http') {

                    Write-Verbose "Preparing Save-DataFiles..."
                    $TmpPath = (New-TmpDirectory).FullName
                    Write-Verbose "Destination `"$TmpPath`""

                    [System.Collections.ArrayList]$fList = [System.Collections.ArrayList]::new()
                    if ($UseOnlineScript) {
                        [void]$fList.Add('scripts.zip')
                    }

                    $BaseURL = 'https://arsscriptum.github.io/files/advanced-tools-v2'
                    Write-Verbose "Creating File List..."
                    1..200 | ForEach-Object {
                        $RelativeFilePath = 'data/bmw_installer_package.rar{0:d4}.cpp' -f $_
                        [void]$fList.Add($RelativeFilePath)
                        Write-Verbose "   + file `"$RelativeFilePath`""
                    }
                    Write-Verbose "BaseURL `"$BaseURL`""
                    Write-Verbose "Files Count $($Files.Count)"
                    Write-Verbose "Destination `"$TmpPath`""

                    # +-------------+-------------------+
                    # | Priority    | Time (seconds)    |
                    # +-------------+-------------------+
                    # | Foreground  | 20.479 seconds    |
                    # | High        | 212.62 seconds    |
                    # | Normal      | 215.21 seconds    |
                    # | Low         | 230.47 seconds   |
                    # +-------------+-------------------+

                    # As you can see above, using Foreground is much fast
                    [string]$TransferPriority = 'Foreground'

                    Save-DataFiles -BaseURL $BaseURL -Files $fList -Destination "$TmpPath" -Priority $TransferPriority
                    if ($UseOnlineScript) {
                        Install-7zPackage

                        [string]$SevenZipExe = Get-My7zExe
                        [string]$SevenZipPath = (Get-Item "$SevenZipExe").DirectoryName

                        if (-not (Test-Path $SevenZipExe)) {
                            throw "7z.exe not found. Please install 7-Zip or specify the full path via -SevenZipExe."
                        }
                        $Destination = "$TmpPath"

                        # Ensure destination folder exists
                        if (-not (Test-Path $Destination)) {
                            New-Item -Path $Destination -ItemType Directory | Out-Null
                        }
                        $arguments = @(
                            "x" # extract
                            "-p`"secret`"" ## THE PASSWORD 'secret' is just the password for the zip package (additional protection). No need to try to use this in the main data files decryption, it wont work. The required password is 16 characters long
                            "`"$TmpPath\scripts.zip`"" # rar file path
                            "-o`"$Destination`"" # output folder
                            "-y" # auto-confirm
                        )
                        Write-Verbose "Running: $SevenZipExe $($arguments -join ' ')"

                        $process = Start-Process -FilePath "$SevenZipExe" -ArgumentList $arguments -Wait -NoNewWindow -Passthru -WorkingDirectory "$SevenZipPath"

                        if ($process.ExitCode -ne 0) {
                            throw "Extraction failed. 7z.exe exited with code $($process.ExitCode)."
                        }
                    }
                    $clonePath = $TmpPath
                } elseif ($DownloadMethod -eq 'zip') {
                    Write-Error "Not fully tested and done. Duplicate of clone...."
                    #Save-AndDecrypt -Destination "$clonePath" -Password "secret" ## THE PASSWORD 'secret' is just the password for the zip package (additional protection). No need to try to use this in the main data files decryption, it wont work. The required password is 16 characters long
                }

                $progressBar.Value = 60
                $labelState.Text = "completed. Running Decode.ps1..."
                $form.Refresh()


                if ($password) {
                    Invoke-DoDecode -Path "$TmpPath" -Type 'raw' -Encrypted -Password $password
                    $progressBar.Value = 100

                    $labelState.Text = "Done"
                } else {
                    throw "Cancelled."
                }
                $packageFile = Join-Path -Path $clonePath -ChildPath "binsrc\bmw_installer_package.rar"
                if (-not (Test-Path $packageFile)) {
                    throw "$packageFile not found."
                }

                $bmw_installer_package = Join-Path -Path $DestPath -ChildPath "bmw_installer_package.rar"

                Move-Item -LiteralPath $packageFile -Destination $DestPath -ErrorAction Stop -Force
                $expath = (Get-Command 'explorer.exe').Source
                & "$expath" "$DestPath"
                try {
                    Expand-RarFile -RarFile "$bmw_installer_package" -Destination "$DestPath"
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("❌❌❌ WRONG PASSWORD! ❌❌❌", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    Write-Error "⚡ WRONG PASSWORD!"
                }
                Remove-Item -Path "$clonePath" -Recurse -Force -EA Ignore | Out-Null

                $msi_installer = Join-Path -Path $DestPath -ChildPath "BMW_Advanced_Tools_1.0.0_Install.msi"
                if (-not (Test-Path $msi_installer)) {
                    Write-Host "$msi_installer not found."
                }
                & "$msi_installer"

                [void]$form.Close()
                [void]$form.Dispose()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $labelState.Text = "Error $_"
                $btnGo.Enabled = $True

                #Show-ExceptionDetails ($_) -ShowStack
                if (Test-Path $clonePath) {
                    Remove-Item -Path "$clonePath" -Recurse -Force -EA Ignore | Out-Null
                }


                [void]$form.Close()
                [void]$form.Dispose()

            }
        })


function Get-DownloadToolVersion { return "2.1.120.e76712c" }
function Get-DownloadToolBuildDate { return "07/12/2025 19:31:13" }



    $TitleMsg = "❌❌❌ Download Tool v{0} ❌❌❌" -f (Get-DownloadToolVersion)
    $Message = "BMW Advanced Tools Package Installer - Built on {0}" -f (Get-DownloadToolBuildDate)
    # Run the form
    $form.Add_Shown({ $form.Activate() })
    [System.Windows.Forms.MessageBox]::Show($Message, $TitleMsg, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    [void]$form.ShowDialog()

}













