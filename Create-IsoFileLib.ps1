
function Create-IsoFile {
<#
.Description 
The Create-IsoFile cmdlet creates a new .iso file containing from chosen folders 
.Example 
Create-IsoFile -Source "c:\tools","c:Downloads\utils" 
This command creates a .iso file in current folder that contains c:\tools and c:\downloads\utils folders.
The folders themselves are included at the root of the .iso image. 
.Example
Create-IsoFile -FromClipboard -Verbose
Before running this command, select and copy (Ctrl-C) files/folders in Explorer first. 
.Example 
dir c:\WinPE | Create-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE"
This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx
#>
    [CmdletBinding(DefaultParameterSetName='Source')]
    Param(
        [string] $IsoFilePath = "$((Get-Date).ToString('yyyyMMdd-HHmmss')).iso", 
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [string] $BootFile = $null,
        [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')]
        [string] $Media = 'DVDPLUSRW_DUALLAYER',
        [string] $Title = (Get-Date).ToString("yyyyMMdd-HHmmss"), 
        [Switch] $Force,
        [Switch] $DotReport,
        [Parameter(ParameterSetName='Clipboard')]
        [Switch] $FromClipboard,
        [Alias('Sources')]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Source')]
        $Source)
    
    Begin { 
        Set-StrictMode -Version 3
        $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
        
        if(-not [string]::IsNullOrEmpty($BootFile)) {
            if('BDR','BDRE' -contains $Media) {
                Write-Warning "Bootable image doesn't seem to work with media type $Media"
            }
            $bootstream = New-Object -ComObject ADODB.Stream -Property @{Type=1}  # adFileTypeBinary
            $bootstream.Open()
            $bootstream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname)
            $boot = New-Object -ComObject IMAPI2FS.BootOptions
            $boot.AssignBootImage($bootstream)
        } else {
            $boot = $null
        }
        
        $mediatypelist = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE')
        $image = New-Object -com IMAPI2FS.MsftFileSystemImage
        $image.VolumeName=$Title
        $image.ChooseImageDefaultsForMediaType($mediatypelist.IndexOf($Media))
        
        $Target = New-Item -Path $IsoFilePath -ItemType File -Force:$Force
    }
    
    Process {
        if($FromClipboard) {
            if($PSVersionTable.PSVersion.Major -lt 5) {
                Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'
                break
            }
            $Source = Get-Clipboard -Format FileDropList
        }
        
        foreach($item in $Source) {
            if($item -isnot [System.IO.FileSystemInfo]) {
                $item = Get-Item -LiteralPath $item
            }
            
            if($item) {
                try {
                    $Image.Root.AddTree($item.FullName, $true)
                } catch {
                    Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.')
                }
            }
        }
    }
    
    End { 
        function Copy-IStreamToFile {
            param(
                [string] $Path,
                [System.MarshalByRefObject] $Stream,
                [int] $BlockSize,
                [int] $TotalBlock)
            
            $marshaledmemory = $null
            $output = $null
            try {
                $marshaledmemory = [System.Runtime.InteropServices.Marshal]::AllocHGlobal([System.Runtime.InteropServices.Marshal]::SizeOf([type][int]))
                $output = [System.IO.File]::OpenWrite($Path)
                $buffer = New-Object Byte[] $BlockSize
                    
                $mi = [System.Runtime.InteropServices.ComTypes.IStream].GetMethod('Read')
                    
                [object[]] $arguments = New-Object object[] 3
                $arguments[0] = $buffer.psobject.BaseObject
                $arguments[1] = $BlockSize.psobject.BaseObject
                $arguments[2] = $marshaledmemory.psobject.BaseObject
                    
                while($TotalBlock -gt 0) {
                    [int] $hr = $mi.Invoke($Stream, $arguments)
                    if($hr -lt 0) {throw [System.Runtime.InteropServices.Marshal]::GetExceptionForHR($hr)}
                        
                    $Length = [System.Runtime.InteropServices.Marshal]::ReadInt32($marshaledmemory)
                        
                    $output.Write($buffer, 0, $Length)
                    $TotalBlock = $TotalBlock -1
                    if($DotReport) {[Console]::Write('.')}
                }
                $output.Flush()
                $output.Close()
                if($DotReport) {[Console]::WriteLine('')}
            } finally {
                if($output) {$output.Dispose()}
                if($marshaledmemory) {
                    [System.Runtime.InteropServices.Marshal]::FreeHGlobal($marshaledmemory) | Out-Null
                }
            }
        }
        
        if($boot) {
            $image.BootImageOptions = $boot
        }
        $resultimg = $image.CreateResultImage()
        Copy-IStreamToFile $Target.FullName $resultimg.ImageStream $resultimg.BlockSize $resultimg.TotalBlocks
        
        Remove-Variable resultimg, image, boot, bootstream
    }
}
