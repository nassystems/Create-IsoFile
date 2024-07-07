function Create-IsoFile {
<#
.Description 
The Create-IsoFile cmdlet creates a new .iso file containing from chosen folders 
.Example 
Create-IsoFile -Source "c:\tools","c:\utils" 
This command creates a .iso file in current folder that contains c:\tools and c:\utils folders.
The folders themselves are included at the root of the .iso image.
.Note
Create-IsoFile version 1.00

MIT License

Copyright (c) 2024 Isao Sato

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>
    [CmdletBinding()]
    Param(
        [parameter(Position=0)]
        [string] $IsoFilePath = "$((Get-Date).ToString('yyyyMMdd-HHmmss')).iso", 
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [string] $BootFile = $null,
        [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DVDPLUSRW_DUALLAYER','BDR','BDRE')]
        [string] $MediaType = 'DVDPLUSRW_DUALLAYER',
        [string] $Title = (Get-Date).ToString("yyyyMMdd-HHmmss"), 
        [Switch] $DotReport,
        [Switch] $NoProgress,
        [Switch] $Force,
        [Alias('Sources')]
        [Parameter(ValueFromPipeline=$true)]
        $Source)
    
    Begin { 
        Set-StrictMode -Version 3
        $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
        
        function Open-FileStream {
            param(
                [Parameter(Mandatory=$true)]
                [string] $LiteralPath,
                [System.IO.FileMode]   $FileMode   = ([System.IO.FileMode]::Open),
                [System.IO.FileAccess] $FileAccess = ([System.IO.FileAccess]::Read),
                [System.IO.FileShare]  $FileShare  = ([System.IO.FileShare]::Read))
            
            $currentpath = [System.Environment]::CurrentDirectory
            try {
                if('FileSystem' -eq $pwd.Provider.Name) {
                    [System.Environment]::CurrentDirectory = $pwd.ProviderPath
                }
                $fstream = New-Object System.IO.FileStream $LiteralPath, $FileMode, $FileAccess, $FileShare
            } finally {
                [System.Environment]::CurrentDirectory = $currentpath
            }
            $fstream
        }
        
        $mediatypeid = @{
            CDROM = 0x1
            CDR = 0x2
            CDRW = 0x3
            DVDROM = 0x4
            DVDRAM = 0x5
            DVDPLUSR = 0x6
            DVDPLUSRW = 0x7
            DVDPLUSR_DUALLAYER = 0x8
            DVDDASHR = 0x9
            DVDDASHRW = 0xa
            DVDDASHR_DUALLAYER = 0xb
            DVDPLUSRW_DUALLAYER = 0xd
            HDDVDROM = 0xe
            HDDVDR = 0xf
            HDDVDRAM = 0x10
            BDROM = 0x11
            BDR = 0x12
            BDRE = 0x13
        }[$MediaType]
        
        if(-not [string]::IsNullOrEmpty($BootFile)) {
            if($mediatypeid -in 0x12, 0x13) {
                Write-Warning "Bootable image doesn't seem to work with media type $Media"
            }
            $bootstream = New-Object -ComObject ADODB.Stream -Property @{Type=1}  # adFileTypeBinary
            $bootstream.Open()
            $bootstream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname)
            $boot = New-Object -ComObject IMAPI2FS.BootOptions
            $boot.AssignBootImage($bootstream)
        } else {
            $bootstream = $null
            $boot = $null
        }
        
        $image = New-Object -com IMAPI2FS.MsftFileSystemImage
        $image.VolumeName=$Title
        $image.ChooseImageDefaultsForMediaType($mediatypeid)
        
        if($Force) {
            $outputstream = Open-FileStream -LiteralPath $IsoFilePath -FileMode Create    -FileAccess Write -FileShare None
        } else {
            $outputstream = Open-FileStream -LiteralPath $IsoFilePath -FileMode CreateNew -FileAccess Write -FileShare None
        }
    }
    
    Process {
        try {
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
        } catch {
            $outputstream.Dispose()
            throw
        }
    }
    
    End { 
        try {
            function Copy-IStreamToFile {
                param(
                    [System.IO.Stream] $Output,
                    [System.MarshalByRefObject] $Stream,
                    [int] $BlockSize,
                    [int] $TotalBlock)
                
                $marshaledmemory = $null
                try {
                    $marshaledmemory = [System.Runtime.InteropServices.Marshal]::AllocHGlobal([System.Runtime.InteropServices.Marshal]::SizeOf([type][int]))
                    $buffer = New-Object Byte[] $BlockSize
                    
                    $readmi = [System.Runtime.InteropServices.ComTypes.IStream].GetMethod('Read')
                    
                    [object[]] $arguments = New-Object object[] 3
                    $arguments[0] = $buffer.psobject.BaseObject
                    $arguments[1] = $BlockSize.psobject.BaseObject
                    $arguments[2] = $marshaledmemory.psobject.BaseObject
                    
                    $currentblock = 0
                    $lastreporttime = [datetime]::MinValue
                    while($currentblock -lt $TotalBlock) {
                        [int] $hr = $readmi.Invoke($Stream, $arguments)
                        if($hr -lt 0) {throw [System.Runtime.InteropServices.Marshal]::GetExceptionForHR($hr)}
                        
                        $Length = [System.Runtime.InteropServices.Marshal]::ReadInt32($marshaledmemory)
                        
                        $Output.Write($buffer, 0, $Length)
                        $currentblock = $currentblock +1
                        if(-not $NoProgress) {
                            if($DotReport) {
                                [Console]::Write('.')
                            } else {
                                if(([datetime]::Now -$lastreporttime) -gt [timespan]::FromSeconds(2)) {
                                    $lastreporttime = [datetime]::Now
                                    Write-Progress -Activity 'writing...' -Status ('{0}/{1}' -f $currentblock, $TotalBlock) -PercentComplete ($currentblock/$TotalBlock)
                                }
                            }
                        }
                    }
                    $Output.Flush()
                    
                    if(-not $NoProgress) {
                        if($DotReport) {
                            if($DotReport) {[Console]::WriteLine('')}
                        } else {
                            Write-Progress -Activity 'wrote.' -Completed
                        }
                    }
                } finally {
                    if($marshaledmemory) {
                        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($marshaledmemory) | Out-Null
                    }
                }
            }
            
            if($boot) {
                $image.BootImageOptions = $boot
            }
            $resultimg = $image.CreateResultImage()
            Copy-IStreamToFile $outputstream $resultimg.ImageStream $resultimg.BlockSize $resultimg.TotalBlocks
            $outputstream.Close()
        } finally {
            $outputstream.Dispose()
            Remove-Variable resultimg, image, boot, bootstream
            [system.gc]::collect(); [system.gc]::waitforpendingfinalizers(); [system.gc]::collect()
        }
    }
}
