<#
Get-ExifData.ps1
#>


[CmdletBinding()]
Param (
    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
      HelpMessage="`r`nPath: From which folder or directory would you like to check the pictures for their EXIF data? `r`n`r`nPlease enter a valid file system path to a directory (a full path name of a folder such as C:\Windows). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path in quotation marks (single or double). `r`n`t- To exit this script, please press [Ctrl] + C `r`n")]
    # [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [Alias("Directory","DirectoryPath","Folder","FolderPath")]
    [string[]]$Path,

    [Parameter(ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true,
        HelpMessage="`r`nFile: From which single file would you like to check the EXIF data? `r`n`r`nPlease enter a valid full file system path pointing to a file (with a full path name of a folder such as C:\Windows\explorer.exe). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path in quotation marks (single or double). `r`n`t- To exit this script, please press [Ctrl] + C `r`n")]
    # [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
    [Alias("SourceFile","FilePath","Files")]
    [string[]]$File,


    [Parameter(HelpMessage="`r`nOutput: In which folder or directory would you like to find the CSV logfile? `r`n`r`nPlease enter a valid file system path to a directory (a full path name of a folder such as C:\Windows). `r`n`r`nNotes:`r`n`t- If the path name includes space characters, please enclose the path in quotation marks (single or double). `r`n")]
    [Alias("OutputFolder","LogFileFolder")]
    [string]$Output = "$($env:USERPROFILE)\Pictures",

    [switch]$Recurse,
    [Alias("Silent","NoLog","DoNotCreateALog")]
    [switch]$SuppressLog,
    [switch]$Open,
    [switch]$Force,
    [switch]$Audio
)




Begin {


    # Function used to convert bytes to MB or GB or TB                                        # Credit: clayman2: "Disk Space"
    function ConvertBytes {
        Param (
            $size
        )
        If ($size -eq $null) {
            [string]'-'
        } ElseIf ($size -eq 0) {
            [string]'-'
        } ElseIf ($size -lt 1MB) {
            $file_size = $size / 1KB
            $file_size = [Math]::Round($file_size, 0)
            [string]$file_size + ' KB'
        } ElseIf ($size -lt 1GB) {
            $file_size = $size / 1MB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' MB'
        } ElseIf ($size -lt 1TB) {
            $file_size = $size / 1GB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' GB'
        } Else {
            $file_size = $size / 1TB
            $file_size = [Math]::Round($file_size, 1)
            [string]$file_size + ' TB'
        } # else
    } # function (ConvertBytes)




    # Set the common parameters
    # Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string
    $ErrorActionPreference = "Stop"
    $computer = $env:COMPUTERNAME
    $log_filename = "exif_log.csv"
    $date = Get-Date -Format g
    $timestamp = Get-Date -Format yyyyMMdd
    $discarded_files = @()
    $source_files = @()
    $horizontal = @()
    $new_files = @()
    $vertical = @()
    $results = @()
    $number_of_paths = $Path.Count
    $number_of_files = $File.Count
    $index = 0
    $num_images = 0
    $num_discarded = 0
    $path_verbose = "Please consider checking that the source folders '$Path', where the pictures containing EXIF data are supposed to be found (and which are set with the -Path parameter), were typed correctly, and that each of them is a valid file system path pointing to a directory. If a path name includes space characters, please enclose that individual path string in quotation marks (single or double). To enter multiple folderpaths, please separate them with a comma."
    $file_verbose = "Please consider checking that the paths to the source files '$File', which is set with the -File parameter, was typed correctly, and that they are valid full file system paths pointing to a file (with a full path name of a folder such as C:\Windows\explorer.exe). If the full filepath includes space characters, please enclose that individual path string in quotation marks (single or double). To enter multiple filepaths, please separate them with a comma."
    $empty_line = ""




    # Test if the Microsoft Windows Image Acquisition (WIA) service is enabled (used in Step 1)
    $test_wia = Get-Service | where { $_.ServiceName -eq "stisvc" } -ErrorAction SilentlyContinue
    $wia_startup_type = (Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "Name='stisvc'").StartMode

        If (($test_wia -eq $null) -or ($wia_startup_type -eq 'Disabled')) {

            # If the WIA service is not enabled, display an error message in console and exit
            $empty_line | Out-String
            Write-Warning "The Microsoft Windows Image Acquisition (WIA) service 'stisvc' doesn't seem to be enabled."
            $empty_line | Out-String
            Write-Verbose "The WIA service is needed for opening and saving the the image files. For futher instructions, how to enable the WIA service, please for example see http://kb.winzip.com/kb/entry/207/ " -verbose
            $empty_line | Out-String
            $exit_text = "Didn't open any files for further examination (Exit 1)."
            Write-Output $exit_text
            $empty_line | Out-String
            Exit

        } Else {
            $continue = $true
        } # Else (If $test_wia)




    # Set Image Swithes (used in Step 2)
    # Source: http://nicholasarmstrong.com/2010/02/exif-quick-reference/
    # Source: http://msdn.microsoft.com/en-us/library/ms630826%28v=vs.85%29.aspx
    # Source: https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
    # Source: https://stackoverflow.com/questions/7076958/read-exif-and-determine-if-the-flash-has-fired#7100717
    # Credit: Franck Richard: "Use PowerShell to Remove Metadata and Resize Images": http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
    $clnfaxdata = @{ 0 = "Clean"; 1 = "Regenerated"; 2 = "Unclean" }
    $compress   = @{ 1 = 'Uncompressed'; 2 = 'CCITT 1D'; 3 = 'T4/Group 3 Fax'; 4 = 'T6/Group 4 Fax'; 5 = 'LZW'; 6 = 'JPEG (old-style)'; 7 = 'JPEG'; 8 = 'Adobe Deflate'; 9 = 'JBIG B&W'; 10 = 'JBIG Color'; 99 = 'JPEG'; 262 = 'Kodak 262'; 32766 = 'Next'; 32767 = 'Sony ARW Compressed'; 32769 = 'Packed RAW'; 32770 = 'Samsung SRW Compressed'; 32771 = 'CCIRLEW'; 32772 = 'Samsung SRW Compressed 2'; 32773 = 'PackBits'; 32809 = 'Thunderscan'; 32867 = 'Kodak KDC Compressed'; 32895 = 'IT8CTPAD'; 32896 = 'IT8LW'; 32897 = 'IT8MP'; 32898 = 'IT8BL'; 32908 = 'PixarFilm'; 32909 = 'PixarLog'; 32946 = 'Deflate'; 32947 = 'DCS'; 34661 = 'JBIG'; 34676 = 'SGILog'; 34677 = 'SGILog24'; 34712 = 'JPEG 2000'; 34713 = 'Nikon NEF Compressed'; 34715 = 'JBIG2 TIFF FX'; 34718 = 'Microsoft Document Imaging (MDI) Binary Level Codec'; 34719 = 'Microsoft Document Imaging (MDI) Progressive Transform Codec'; 34720 = 'Microsoft Document Imaging (MDI) Vector'; 34892 = 'Lossy JPEG'; 65000 = 'Kodak DCR Compressed'; 65535 = 'Pentax PEF Compressed' }
    $contrast   = @{ 0 = "Normal"; 1 = "Low"; 2 = "High" }
    $custrender = @{ 0 = "Normal process"; 1 = "Custom process"; 3 = "HDR"; 6 = "Panorama"; 8 = "Portrait" }
    $exposurem  = @{ 0 = "Auto exposure"; 1 = "Manual exposure"; 2 = "Auto Bracket" }
    $exposurep  = @{ 0 = 'Not defined'; 1 = 'Manual'; 2 = 'Program AE (normal)'; 3 = 'Aperture-priority AE'; 4 = 'Shutter speed priority AE'; 5 = 'Creative (depth of field)'; 6 = 'Action (fast shutter speed)'; 7 = 'Portrait'; 8 = 'Landscape'; 9 = 'Bulb' }
    $extrasamp  = @{ 0 = "Unspecified"; 1 = "Associated Alpha"; 2 = "Unassociated Alpha" }
    $faxprofil  = @{ 0 = "Unknown"; 1 = "Minimal B&W lossless, S"; 2 = "Extended B&W lossless, F"; 3 = "Lossless JBIG B&W, J "; 4 = "Lossy color and grayscale, C"; 5 = "Lossless color and grayscale, L"; 6 = "Mixed raster content, M"; 7 = "Profile T"; 255 = "Multi Profiles" }
    $filesrc    = @{ 1 = "Film scanner"; 2 = "Reflection print scanner"; 3 = "Digital camera" }
    $fillord    = @{ 1 = "Normal"; 2 = "Reversed" }
    $flashvalue = @{ 0 = 'Flash not fired, no flash'; 1 = 'Flash was fired, fired'; 5 = 'Flash was fired, fired, return not detected'; 7 = 'Flash was fired, fired, return detected'; 8 = 'Flash not fired, on, did not fire'; 9 = 'Flash was fired, on, fired'; 13 = 'Flash was fired, on, return not detected'; 15 = 'Flash was fired, on, return detected'; 16 = 'Flash not fired, off, did not fire'; 20 = 'Flash not fired, off, did not fire, return not detected'; 24 = 'Flash not fired, auto, did not fire'; 25 = 'Flash was fired, auto, fired'; 29 = 'Flash was fired, auto, fired, return not detected'; 31 = 'Flash was fired, auto, fired, return detected'; 32 = 'Flash not fired, no flash function'; 48 = 'Flash not fired, off, no flash function'; 65 = 'Flash was fired, fired, red-eye reduction'; 69 = 'Flash was fired, fired, red-eye reduction, return not detected'; 71 = 'Flash was fired, fired, red-eye reduction, return detected'; 73 = 'Flash was fired, on, red-eye reduction'; 77 = 'Flash was fired, on, red-eye reduction, return not detected'; 79 = 'Flash was fired, on, red-eye reduction, return detected'; 80 = 'Flash not fired, off, red-eye reduction'; 88 = 'Flash not fired, auto, did not fire, red-eye reduction'; 89 = 'Flash was fired, auto, fired, red-eye reduction'; 93 = 'Flash was fired, auto, fired, red-eye reduction, return not detected'; 95 = 'Flash was fired, auto, fired, red-eye reduction, return detected' }
    $focal      = @{ 1 = "None"; 2 = "Inches"; 3 = "Centimetres" ; 4 = "Millimetres"; 5 = "Micrometres" }
    $focalprun  = @{ 1 = "None"; 2 = "Inches"; 3 = "cm" ; 4 = "mm"; 5 = "Micrometres" }
    $gain       = @{ 0 = "None"; 1 = "Low gain up"; 2 = "High gain up"; 3 = "Low gain down"; 4 = "High gain down" }
    $gpsaltref  = @{ 0 = "Above sea level"; 1 = "Below sea level" }
    $gpsdiffer  = @{ 0 = "No correction"; 1 = "Differential correction applied" }
    $gpsdirect  = @{ M = "Magnetic North"; T = "True North" }
    $gpsdistrf  = @{ K = "Kilometers"; M = "Miles"; N = "Nautical Miles" }
    $gpsmeasure = @{ 2 = "2-Dimensional Measurement"; 3 = "3-Dimensional Measurement" }
    $gpsspeedrf = @{ K = "km/h"; M = "mph"; N = "knots" }
    $gpsstat    = @{ A = "Measurement Active"; V = "Measurement Void" }
    $grayrespun = @{ 1 = '0.1'; 2 = '0.001'; 3 = '0.0001'; 4 = '0.00001'; 5 = '0.000001' }
    $incst      = @{ 1 = "CMYK"; 2 = "Not CMYK" }
    $indexd     = @{ 0 = "Not indexed"; 1 = "Indexed" }
    $jpegpro    = @{ 1 = "Baseline"; 14 = "Lossless" }
    $lightsrc   = @{ 0 = 'Unknown'; 1 = 'Daylight'; 2 = 'Fluorescent'; 3 = 'Tungsten (Incandescent)'; 4 = 'Flash'; 9 = 'Fine Weather'; 10 = 'Cloudy'; 11 = 'Shade'; 12 = 'Daylight Fluorescent (D 5700 - 7100K)'; 13 = 'Day White Fluorescent (N 4600 - 5500K)'; 14 = 'Cool White Fluorescent (W 3800 - 4500K)'; 15 = 'White Fluorescent (WW 3250 - 3800K)'; 16 = 'Warm White Fluorescent (L 2600 - 3250K)'; 17 = 'Standard Light A'; 18 = 'Standard Light B'; 19 = 'Standard Light C'; 20 = 'D55'; 21 = 'D65'; 22 = 'D75'; 23 = 'D50'; 24 = 'ISO Studio Tungsten'; 255 = 'Other' }
    $metering   = @{ 0 = "Unknown"; 1 = "Average"; 2 = "Center-weighted average"; 3 = "Spot"; 4 = "Multi-spot"; 5 = "Multi-segment"; 6 = "Partial"; 255 = "Other" }
    $opiprox    = @{ 0 = "Higher resolution image does not exist"; 1 = "Higher resolution image exists" }
    $orient     = @{ 1 = "Horizontal"; 2 = "Mirror horizontal"; 3 = "Rotate 180 degrees"; 4 = "Mirror vertical"; 5 = "Mirror horizontal and rotate 270 degrees clockwise"; 6 = "Rotate 90 degrees clockwise"; 7 = "Mirror horizontal and rotate 90 degrees clockwise"; 8 = "Rotate 270 degrees clockwise" }
    $photmtrint = @{ 0 = 'WhiteIsZero'; 1 = 'BlackIsZero'; 2 = 'RGB'; 3 = 'RGB Palette'; 4 = 'Transparency Mask'; 5 = 'CMYK'; 6 = 'YCbCr'; 8 = 'CIELab'; 9 = 'ICCLab'; 10 = 'ITULab'; 32803 = 'Color Filter Array'; 32844 = 'Pixar LogL'; 32845 = 'Pixar LogLuv'; 34892 = 'Linear Raw' }
    $planarconf = @{ 1 = "Chunky"; 2 = "Planar" }
    $predict    = @{ 1 = "None"; 2 = "Horizontal differencing" }
    $previewcol = @{ 0 = "Unknown"; 1 = "Gray Gamma 2.2"; 2 = "sRGB"; 3 = "Adobe RGB"; 4 = "ProPhoto RGB" }
    $profiletyp = @{ 0 = "Unspecified"; 1 = "Group 3 FAX" }
    $saturation = @{ 0 = "Normal"; 1 = "Low saturation"; 2 = "High saturation" }
    $scene      = @{ 0 = "Standard"; 1 = "Landscape"; 2 = "Portrait"; 3 = "Night" }
    $securitycl = @{ C = "Confidential"; R = "Restricted"; S = "Secret"; T = "Top Secret"; U = "Unclassified" }
    $sensing    = @{ 1 = "Not defined"; 2 = "One-chip colour area"; 3 = "Two-chip colour area" ; 4 = "Three-chip colour area"; 5 = "Colour sequential area"; 7 = "Trilinear"; 8 = "Colour sequential linear" }
    $sensittyp  = @{ 0 = "Unknown"; 1 = "Standard Output Sensitivity"; 2 = "Recommended Exposure Index"; 3 = "ISO Speed" ; 4 = "Standard Output Sensitivity and Recommended Exposure Index"; 5 = "Standard Output Sensitivity and ISO Speed"; 6 = "Recommended Exposure Index and ISO Speed"; 7 = "Standard Output Sensitivity, Recommended Exposure Index and ISO Speed" }
    $sensmethod = @{ 1 = "Monochrome area"; 2 = "One-chip colour area"; 3 = "Two-chip colour area" ; 4 = "Three-chip colour area"; 5 = "Colour sequential area"; 6 = "Monochrome linear"; 7 = "Trilinear"; 8 = "Colour sequential linear" }
    $sharpness  = @{ 0 = "Normal"; 1 = "Soft"; 2 = "Hard" }
    $subfiletyp = @{ 1 = "Full-resolution image"; 2 = "Reduced-resolution image"; 3 = "Single page of multi-page image" }
    $subjdist   = @{ 0 = "Unknown"; 1 = "Macro"; 2 = "Close" ; 3 = "Distant" }
    $threshold  = @{ 1 = "No dithering or halftoning"; 2 = "Ordered dither or halftone"; 3 = "Randomized dither" }
    $unit       = @{ 1 = "None"; 2 = "Inches"; 3 = "Centimetres" }
    $white      = @{ 0 = "Auto white balance"; 1 = "Manual white balance" }
    $ycbcrpos   = @{ 1 = "Centered"; 2 = "Co-sited" }




    # Display a welcome message in console    
    If ($number_of_files -gt 1) {
        $file_word = "files"
    } ElseIf ($number_of_files -eq 1) {
        $file_word = "file"
    } Else {
        $continue = $true
    } # Else ($number_of_files -gt 1)

        If ($number_of_paths -gt 1 -and -not $Recurse) {
            $folder_word = "folders"
        } ElseIf ($number_of_paths -eq 1 -and -not $Recurse) {
            $folder_word = "folder"
        } ElseIf ($number_of_paths -gt 1 -and $Recurse) {
            $folder_word = "folders recursively"
        } ElseIf ($number_of_paths -eq 1 -and $Recurse) {
            $folder_word = "folder recursively"
        } Else {
            $continue = $true
        } # Else ($number_of_paths -gt 1)    

                If (-not $Path -and -not $File) {
                    $welcome_text = "Searching for image files at the default -Path $($env:USERPROFILE)\Pictures"
                } ElseIf ($Path -and -not $File) {
                    $welcome_text = "Searching for image files at $number_of_paths $folder_word"
                } ElseIf (-not $Path -and $File) {
                    $welcome_text = "Reading $number_of_files image $file_word."
                } ElseIf ($Path -and $File) {
                    $welcome_text = "Searching for image files at $number_of_paths $folder_word and reading additional $number_of_files image $file_word."
                } Else {
                    $welcome_text = ""
                } # Else (-not $Path -and -not $File)

    $empty_line | Out-String
    Write-Output $welcome_text




    # Try to enumerate the files inputted trough the $File parameter
    If ($File) {

        ForEach ($file_candidate in $File) {

            If ((Test-Path $file_candidate -PathType 'Leaf') -eq $true) {

                # Add the individual filepath in question to the list of files to be processed further
                $new_files += Get-ChildItem -Path $file_candidate -Force -ErrorAction SilentlyContinue

            } ElseIf ((Test-Path $file_candidate -PathType 'Container') -eq $true) {

                # A directory is inputted through the $File parameter
                $empty_line | Out-String
                Write-Warning "-File: '$file_candidate' seems to be a directory."
                $empty_line | Out-String
                Write-Verbose "For checking the EXIF data on files inside a folder, please consider using the -Path parameter instead of the -File parameter." -verbose
                $empty_line | Out-String
                Exit

            } ElseIf ((Test-Path $file_candidate) -eq $false) {

                # A non-existing file is inputted through the $File parameter
                $empty_line | Out-String
                Write-Warning "-File: '$file_candidate' doesn't seem to exist."
                $empty_line | Out-String
                Write-Verbose $file_verbose -verbose
                $empty_line | Out-String
                Exit

            } Else {
                $continue = $true
            } # Else (If Test-Path $file_candidate -PathType 'Leaf')

        } # ForEach $file_candidate in $File

    } Else {
        $continue = $true
    } # Else (If $File)




    # Try to enumerate the folders inputted trough the $Path parameter
    If ($Path) {

        ForEach ($path_candidate in $Path) {

            If ((Test-Path $path_candidate -PathType 'Container') -eq $true) {

                # Omit the last character, if it's \
                # Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908
                If ((($path_candidate).EndsWith("\")) -eq $true)   { $path_candidate = $path_candidate -replace ".{1}$" }    Else { $continue = $true }

                # Add the files inside the folders to the list of files to be processed further according to the recursive setting
                If ($Recurse) {
                    $new_files += Get-ChildItem -Path $path_candidate -Recurse -Force -ErrorAction SilentlyContinue | where { $_.PsIsContainer -eq $false }
                } Else {
                    $new_files += Get-ChildItem -Path $path_candidate -Force -ErrorAction SilentlyContinue | where { $_.PsIsContainer -eq $false }
                } # Else If $Recurse

            } ElseIf ((Test-Path $path_candidate -PathType 'Leaf') -eq $true) {

                # A file is inputted through the $Path parameter
                $empty_line | Out-String
                Write-Warning "-Path: '$path_candidate' seems to be a file."
                $empty_line | Out-String
                Write-Verbose "For checking the EXIF data on individual files, please consider using the -File parameter instead of the -Path parameter." -verbose
                $empty_line | Out-String
                Exit

            } ElseIf ((Test-Path $path_candidate) -eq $false) {

                # A non-existing folder is inputted through the $Path parameter
                $empty_line | Out-String
                Write-Warning "-Path: '$path_candidate' doesn't seem to exist."
                $empty_line | Out-String
                Write-Verbose $path_verbose -verbose
                $empty_line | Out-String
                Exit

            } Else {
                $continue = $true
            } # Else (If Test-Path $path_candidate -PathType 'Container')

        } # ForEach $path_candidate in $Path

    } Else {
        $continue = $true
    } # Else (If $Path)




    # Try to find something if no source is specified
    If ( -not $File -and -not $Path ) {

        $Path = "$($env:USERPROFILE)\Pictures"

            If ((Test-Path $Path -PathType 'Container') -eq $true) {

                # Add the files inside the folders to the list of files to be processed further according to the recursive setting
                If ($Recurse) {
                    $new_files += Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | where { $_.PsIsContainer -eq $false }
                } Else {
                    $new_files += Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue | where { $_.PsIsContainer -eq $false }
                } # Else If $Recurse

            } Else {
                $continue = $true
            } # Else (If Test-Path $Path -PathType 'Container')

    } Else {
        $continue = $true
    } # Else (If -not $File -and -not $Path)




    # Filter the image files
    # The Select-Object -Unique may slow the process down quite considerably...
    # $unique_files = $new_files | Select-Object -Unique

        # ForEach ($item in $unique_files) {
        ForEach ($item in $new_files) {

            If ( (([System.IO.FileInfo]"$item").Extension) -like ".3FR" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".ARW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".BMP" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".CR2" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".CR3" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".CRW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".DNG" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".EPS" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".GIF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".JP2" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".JPEG" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".JPG" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".MDC" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".MRW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".NEF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".NRW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".ORF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PCT" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PEF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PGF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PNG" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PSD" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".PTX" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".R3D" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".RAF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".RAW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".RW2" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".RWL" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".SR2" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".SRF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".SRW" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".TIF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".TIFF" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".X3F" -or
            (([System.IO.FileInfo]"$item").Extension) -like ".XMP" ) {
                $num_images++
                $source_files += Get-ChildItem -Path $item.FullName -Force -ErrorAction SilentlyContinue
            } Else {
                $num_discarded++
                $discarded_files += Get-ChildItem -Path $item.FullName -Force -ErrorAction SilentlyContinue
            } # Else (If (([System.IO.FileInfo]"$item").Extension) -like)

        } # ForEach $item in $new_files


                    # Set the progress bar variables ($id denominates different progress bars, if more than one is being displayed)
                    $activity           = "Processing $($source_files.Count) image files"
                    $status             = " "
                    $task               = "Setting Initial Variables"
                    $operations         = (($source_files.Count) + 2)
                    $total_steps        = (($source_files.Count) + 3)
                    $task_number        = 0
                    $id                 = 1


                    # Start the progress bar if there is more than one unique file to process, and halt the procedure if it seems that there's nothing to do
                    If (($source_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete ((0.000002 / $total_steps) * 100)
                    } ElseIf (($source_files.Count) -lt 1) {
                        $empty_line | Out-String
                        $text = "Didn't find any image files (Exit 2)."
                        Write-Output $text
                        $empty_line | Out-String
                        Exit
                    } Else {
                        $continue = $true
                    } # If ($source_files.Count)



        # Test if the Output-path -LogFileFolder exists
        If ((Test-Path $Output -PathType Leaf) -eq $true) {

            # File: Display an error message in console and exit
            $empty_line | Out-String
            Write-Warning "-Output: '$Output' seems to point to a file."
            $empty_line | Out-String
            $exit_text = "Couldn't open the -Output folder '$Output' since it's a file (Exit 3)."
            Write-Output $exit_text
            $empty_line | Out-String
            Exit

        } ElseIf ((Test-Path $Output) -eq $false) {

            If ($Force) {

                # If the Force was used, create the destination folder ($Output)
                New-Item "$Output" -ItemType Directory -Force | Out-Null
                $continue = $true

            } Else {

                # No Destination: Display an error message in console
                $empty_line | Out-String
                Write-Warning "-Output: '$Output' doesn't seem to exist."

                # Offer the user an option to create the defined $Output -LogFileFolder                 # Credit: lamaar75: "Creating a Menu": http://powershell.com/cs/forums/t/9685.aspx
                # Source: "Adding a Simple Menu to a Windows PowerShell Script": https://technet.microsoft.com/en-us/library/ff730939.aspx
                $title_corner = "Create folder '$Output' with this script?"
                $message = " "

                $yes = New-Object System.Management.Automation.Host.ChoiceDescription    "&Yes",    "Yes:     tries to create a new folder and to write the CSV logfile there."
                $no = New-Object System.Management.Automation.Host.ChoiceDescription     "&No",     "No:      exits from this script (similar to Ctrl + C)."
                $exit = New-Object System.Management.Automation.Host.ChoiceDescription   "&Exit",   "Exit:    exits from this script (similar to Ctrl + C)."
                $abort = New-Object System.Management.Automation.Host.ChoiceDescription  "&Abort",  "Abort:   exits from this script (similar to Ctrl + C)."
                $cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel", "Cancel:  exits from this script (similar to Ctrl + C)."

                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $exit, $abort, $cancel)
                $choice_result = $host.ui.PromptForChoice($title_corner, $message, $options, 1)

                    switch ($choice_result)
                        {
                            0 {
                            "Yes. Creating the folder.";
                            New-Item "$Output" -ItemType Directory -Force | Out-Null
                            $continue = $true
                            }
                            1 {
                            $empty_line | Out-String
                            "No. Exiting from EXIF data retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            2 {
                            $empty_line | Out-String
                            "Exit. Exiting from EXIF data retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            3 {
                            $empty_line | Out-String
                            "Abort. Exiting from EXIF data retrieval script.";
                            $empty_line | Out-String
                            Exit
                            }
                            4 {
                            $empty_line | Out-String
                            "Cancel. Exiting from EXIF data retrieval script.";
                            $empty_line | Out-String
                            Exit
                            } # 4
                        } # switch
            } # Else If $Force)
        } Else {
            $continue = $true
        } # Else (Test-Path $Output -PathType Leaf)


                # Resolve the Output-path ("Destination") (if the Output-path is specified as relative) and remove the last character if it's \
                # Source: https://technet.microsoft.com/en-us/library/ee692804.aspx
                # Source: http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string#32608908
                $real_output_path = (Resolve-Path -Path $Output).Path
                If ((($real_output_path).EndsWith("\")) -eq $true) { $real_output_path = $real_output_path -replace ".{1}$" }


} # Begin




Process {

    # Process each image
    $image = New-Object -ComObject WIA.ImageFile
    ForEach ($picture in $source_files) {

                    # Increment the step counters and reset the comment section ($remarks)
                    $remarks = @()
                    $task_number++
                    $index++

                    # Update the progress bar if there is more than one unique file to process
                    $activity = "Processing $($source_files.Count) image files - Step $task_number/$operations"
                    If (($source_files.Count) -ge 2) {
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $picture.Name -PercentComplete (($task_number / $total_steps) * 100)
                    } # If ($source_files.Count)


        # Step 1
        # Load the image as an ImageFile COM object with Microsoft Windows Image Acquisition (WIA)
        # Note: To edit the image: WIA.ImageProcess... $ip = New-Object -ComObject WIA.ImageProcess  ...    $ip.FilterInfos | fl *
        # Note: For data retriaval also: System.Drawing.Image.PropertyItems
        # Note: In the default installations of Windows Server 2003 and 2012 the Windows Image Acquisition (WIA) service is not enabled by default.
        # Note: On Windows Server 2008 the WIA service is not installed by default. To add this feature, please see the WinZip Knowledgebase link below.
        # Source: http://kb.winzip.com/kb/entry/207/
        # Source: https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx
        # Source: https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/
        # Source: http://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell
        # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-wmiobject
        # Source: https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic
        # $disabled_wia = Get-WmiObject -Class Win32_Service -ComputerName $computer -Filter "Name='stisvc'" | Where-Object { $_.StartMode -eq 'Disabled' }
        # $test_wia = Get-Service | where {$_.DisplayName -like "*Windows Image Acquisition*" } -ErrorAction SilentlyContinue


        # $name = dir "C:\Users\Juha\Pictures\3.JPG" | select -ExpandProperty FullName
        # $name = dir "C:\Users\Juha\Pictures\philips_B21091H_A.jpg" | select -ExpandProperty FullName
        # $name = dir "C:\Temp\upload\fDSCN0433.JPG" | select -ExpandProperty FullName

        $name = $picture | select -ExpandProperty FullName
        $image.LoadFile($name)

            # Determine the picture orientation
            If (($image.Width -ge "320") -and ($image.Width -gt $image.Height)) {
                $type = "Landscape"
                $orientation = "Horizontal"
                $horizontal += "$name"
            } ElseIf (($image.Height -ge "320") -and ($image.Height -gt $image.Width)) {
                $type = "Portrait"
                $orientation = "Vertical"
                $vertical += "$name"
            } Else {
                $continue = $true
            } # Else (If $image.Width)


        # Step 2
        # Retrieve image properties (n ~360)
        # Source: https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012
        If ($image.IsIndexedPixelFormat -eq $true )     { $remarks += "Pixel data contains palette indexes" }                               Else { $continue = $true }
        If ($image.IsAlphaPixelFormat -eq $true )       { $remarks += "Pixel data has alpha information" }                                  Else { $continue = $true }
        If ($image.IsExtendedPixelFormat -eq $true )    { $remarks += "Pixel data has extended color information (16 bit/channel)" }        Else { $continue = $true }
        If ($image.IsAnimated -eq $true )               { $remarks += "Image is animated" }                                                 Else { $continue = $true }

                # Get SHA256
                Try {

                    If (($PSVersionTable.PSVersion).Major -ge 4) {
                        # Requires PowerShell version 4.0
                        # Source: https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/get-filehash
                        $hash = Get-FileHash -Path $name -Algorithm SHA256 | Select-Object -ExpandProperty Hash

                    } Else {
                        # Get SHA256 hash value in PowerShell version 2 regardless whether it is opened in another program or not
                        # Requires .NET Framework v3.5
                        # Source: http://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows
                        # Source: https://msdn.microsoft.com/en-us/library/system.security.cryptography.sha256cryptoserviceprovider(v=vs.110).aspx
                        # Source: http://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows
                        # Credit: Twon of An: "Get the SHA1,SHA256,SHA384,SHA512,MD5 or RIPEMD160 hash of a file" https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file
                        # Credit: Gisli: "Unable to read an open file with binary reader" http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader

                            # SHA256 (PowerShell v2)
                            $SHA256 = New-Object -TypeName System.Security.Cryptography.SHA256CryptoServiceProvider
                            $source_file_b = [System.IO.File]::Open("$name", [System.IO.Filemode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                            $hash = [System.BitConverter]::ToString($SHA256.ComputeHash($source_file_b)) -replace "-",""
                            $source_file_b.Close()

                    } # Else (If $PSVersionTable.PSVersion)

                } Catch { Write-Debug $_.Exception }


            # Source: https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
            # Source: https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell
            # Source: http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921

            # GPSProcessingMethod
            If ($image.Properties.Exists('27') -and $image.Properties.Item('27').Type -eq 1100) {
                $gpsproc = (($image.Properties.Item('27')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }
            ElseIf ($image.Properties.Exists('27') -and $image.Properties.Item('27').Type -eq 1002) {
                $gpsproc = $image.Properties.Item('27').Value }                                                                                             Else { $continue = $true }

            # GPSAreaInformation
            If ($image.Properties.Exists('28') -and $image.Properties.Item('28').Type -eq 1100) {
                $gpsarea = (($image.Properties.Item('28')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }
            ElseIf ($image.Properties.Exists('28') -and $image.Properties.Item('28').Type -eq 1002) {
                $gpsarea = $image.Properties.Item('28').Value }                                                                                             Else { $continue = $true }

            # UserComment
            If ($image.Properties.Exists('37510') -and $image.Properties.Item('37510').Type -eq 1100) {
                $usercomment = (($image.Properties.Item('37510')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }
            ElseIf ($image.Properties.Exists('37510') -and $image.Properties.Item('37510').Type -eq 1002) {
                $usercomment = $image.Properties.Item('37510').Value }                                                                                      Else { $continue = $true }

            # XPTitle
            If ($image.Properties.Exists('40091')) {
                $title = (($image.Properties.Item('40091')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }            Else { $continue = $true }

            # XPComment
            If ($image.Properties.Exists('40092')) {
                $comment = (($image.Properties.Item('40092')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }          Else { $continue = $true }

            # XPAuthor
            If ($image.Properties.Exists('40093')) {
                $author = (($image.Properties.Item('40093')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }           Else { $continue = $true }

            # XPKeywords
            If ($image.Properties.Exists('40094')) {
                $keywords = (($image.Properties.Item('40094')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }         Else { $continue = $true }

            # XPSubject
            If ($image.Properties.Exists('40095')) {
                $subject = (($image.Properties.Item('40095')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }          Else { $continue = $true }

            # ExifVersion
            If ($image.Properties.Exists('36864')) {
                $ExifVer = (($image.Properties.Item('36864')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }          Else { $continue = $true }

            # FlashpixVersion
            If ($image.Properties.Exists('40960')) {
                $FlashpixVer = (($image.Properties.Item('40960')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }      Else { $continue = $true }

            # Reading MakerNote EXIF data (camera maker specific information) requires decoding proprietary code, which has been compiled with different parameters by each vendor
            # c.f. https://stackoverflow.com/questions/26696074/read-makernote-exif-tag-in-c-sharp
            # c.f. http://tawbaware.com/990exif.htm
            # c.f. http://www.burren.cx/david/canon.html
            # Source: https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html Tag ID: 0x927c
            # If ($image.Properties.Exists('37500')) {
            # $MakerNote = (($image.Properties.Item('37500')).Value | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_) }) -join ("") }            Else { $continue = $true }


            $results += $obj_file = New-Object -TypeName PSCustomObject -Property @{

                        'Accessed'                      = $picture.LastAccessTime
                        'Accessed (UTC)'                = $picture.LastAccessTimeUtc
                        'ActiveFrame'                   = $image.ActiveFrame
                        'Aspect Ratio'                  = [math]::round(($image.Width / $image.Height), 4)
                        'Attributes'                    = $picture | select -ExpandProperty Attributes
                        'BaseName'                      = $picture.BaseName
                    #    'File Name Without Extension'   = $picture.BaseName
                        'Comments'                      = ($picture | select -ExpandProperty VersionInfo).Comments
                        'CompanyName'                   = ($picture | select -ExpandProperty VersionInfo).CompanyName
                        'Created'                       = $picture.CreationTime
                        'Created (UTC)'                 = $picture.CreationTimeUtc
                        'Directory'                     = $picture.Directory
                        'DirectoryName'                 = $picture | select -ExpandProperty DirectoryName
                        'File'                          = $picture.Name
                        'FileExtension'                 = $image.FileExtension
                        'Folder'                        = ($picture | select -ExpandProperty Directory).Name
                        'FormatID'                      = $image.FormatID
                        'FrameCount'                    = $image.FrameCount
                        'Height'                        = $image.Height
                        'Home'                          = $picture.FullName
                        'HorizontalResolution'          = $image.HorizontalResolution
                        'Index'                         = $index
                        'IsAlphaPixelFormat'            = $image.IsAlphaPixelFormat
                        'IsAnimated'                    = $image.IsAnimated
                        'IsExtendedPixelFormat'         = $image.IsExtendedPixelFormat
                        'IsIndexedPixelFormat'          = $image.IsIndexedPixelFormat
                        'IsPreRelease'                  = ($picture | select -ExpandProperty VersionInfo).IsPreRelease
                        'IsReadOnly'                    = $picture | select -ExpandProperty IsReadOnly
                        'IsSpecialBuild'                = ($picture | select -ExpandProperty VersionInfo).IsSpecialBuild
                        'LegalCopyright'                = ($picture | select -ExpandProperty VersionInfo).LegalCopyright
                        'LegalTrademarks'               = ($picture | select -ExpandProperty VersionInfo).LegalTrademarks
                        'Log Date'                      = $date
                        'Modified'                      = $picture.LastWriteTime
                        'Modified (UTC)'                = $picture.LastWriteTimeUtc
                        'Orientation'                   = $orientation
                        'Original File Extension'       = $picture.Extension
                        'PixelDepth'                    = $image.PixelDepth
                        'raw_size'                      = $picture.Length
                        'Remarks'                       = ($remarks -join ', ')
                        'SHA256'                        = $hash
                        'Size'                          = (ConvertBytes ($picture.Length))
                        'Source'                        = $name
                        'Type'                          = $type
                        'VerticalResolution'            = $image.VerticalResolution
                        'Width'                         = $image.Width


                        # Source: https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
                        'Author'                        = If ($image.Properties.Exists('40093')) { $author }                                                    Else {" "}
                        'Comment'                       = If ($image.Properties.Exists('40092')) { $comment }                                                   Else {" "}
                        'Keywords'                      = If ($image.Properties.Exists('40094')) { $keywords }                                                  Else {" "}
                        'Subject'                       = If ($image.Properties.Exists('40095')) { $subject }                                                   Else {" "}
                        'Title'                         = If ($image.Properties.Exists('40091')) { $title }                                                     Else {" "}


                        # Source: http://www.exiv2.org/tags.html
                        # Source: https://sno.phy.queensu.ca/~phil/exiftool/TagNames/GPS.html
                        'ActiveArea'                    = If ($image.Properties.Exists('50829')) { $image.Properties.Item('50829').Value }                      Else {" "}
                        'AnalogBalance'                 = If ($image.Properties.Exists('50727')) { $image.Properties.Item('50727').Value }                      Else {" "}
                        'AntiAliasStrength'             = If ($image.Properties.Exists('50738')) { $image.Properties.Item('50738').Value }                      Else {" "}
                        'Artist'                        = If ($image.Properties.Exists('315'))   { $image.Properties.Item('315').Value }                        Else {" "}
                        'AsShotICCProfile'              = If ($image.Properties.Exists('50831')) { $image.Properties.Item('50831').Value }                      Else {" "}
                        'AsShotNeutral'                 = If ($image.Properties.Exists('50728')) { $image.Properties.Item('50728').Value }                      Else {" "}
                        'AsShotPreProfileMatrix'        = If ($image.Properties.Exists('50832')) { $image.Properties.Item('50832').Value }                      Else {" "}
                        'AsShotWhiteXY'                 = If ($image.Properties.Exists('50729')) { $image.Properties.Item('50729').Value }                      Else {" "}
                        'BaselineExposure'              = If ($image.Properties.Exists('50730')) { $image.Properties.Item('50730').Value }                      Else {" "}
                        'BaselineNoise'                 = If ($image.Properties.Exists('50731')) { $image.Properties.Item('50731').Value }                      Else {" "}
                        'BaselineSharpness'             = If ($image.Properties.Exists('50732')) { $image.Properties.Item('50732').Value }                      Else {" "}
                        'BatteryLevel'                  = If ($image.Properties.Exists('33423')) { $image.Properties.Item('33423').Value }                      Else {" "}
                        'BayerGreenSplit'               = If ($image.Properties.Exists('50733')) { $image.Properties.Item('50733').Value }                      Else {" "}
                        'BestQualityScale'              = If ($image.Properties.Exists('50780')) { $image.Properties.Item('50780').Value }                      Else {" "}
                        'BitsPerSample'                 = If ($image.Properties.Exists('258'))   { $image.Properties.Item('258').Value }                        Else {" "}
                        'BlackLevel'                    = If ($image.Properties.Exists('50714')) { $image.Properties.Item('50714').Value }                      Else {" "}
                        'BlackLevelDeltaH'              = If ($image.Properties.Exists('50715')) { $image.Properties.Item('50715').Value }                      Else {" "}
                        'BlackLevelDeltaV'              = If ($image.Properties.Exists('50716')) { $image.Properties.Item('50716').Value }                      Else {" "}
                        'BlackLevelRepeatDim'           = If ($image.Properties.Exists('50713')) { $image.Properties.Item('50713').Value }                      Else {" "}
                        'BodySerialNumber'              = If ($image.Properties.Exists('42033')) { $image.Properties.Item('42033').Value }                      Else {" "}
                        'BrightnessValue'               = If ($image.Properties.Exists('37379')) { $image.Properties.Item('37379').Value }                      Else {" "}
                        'Byte_AsShotProfileName'        = If ($image.Properties.Exists('50934')) { $image.Properties.Item('50934').Value }                      Else {" "}
                        'Byte_CameraCalibrationSignat'  = If ($image.Properties.Exists('50931')) { $image.Properties.Item('50931').Value }                      Else {" "}
                        'Byte_CFAPattern'               = If ($image.Properties.Exists('33422')) { $image.Properties.Item('33422').Value }                      Else {" "}
                        'Byte_CFAPlaneColor'            = If ($image.Properties.Exists('50710')) { $image.Properties.Item('50710').Value }                      Else {" "}
                        'Byte_ClipPath'                 = If ($image.Properties.Exists('343'))   { $image.Properties.Item('343').Value }                        Else {" "}
                        'Byte_DNGBackwardVersion'       = If ($image.Properties.Exists('50707')) { $image.Properties.Item('50707').Value }                      Else {" "}
                        'Byte_DNGPrivateData'           = If ($image.Properties.Exists('50740')) { $image.Properties.Item('50740').Value }                      Else {" "}
                        'Byte_DNGVersion'               = If ($image.Properties.Exists('50706')) { $image.Properties.Item('50706').Value }                      Else {" "}
                        'Byte_DotRange'                 = If ($image.Properties.Exists('336'))   { $image.Properties.Item('336').Value }                        Else {" "}
                        'Byte_GPSAltitudeRef'           = If ($image.Properties.Exists('5'))     { $gpsaltref[[int]$image.Properties.Item('5').Value] }         Else {" "}
                        'Byte_GPSVersionID'             = If ($image.Properties.Exists('0'))     { ($image.Properties.Item('0').Value) -join ('.') }            Else {" "}
                        'Byte_ImageResources'           = If ($image.Properties.Exists('34377')) { $image.Properties.Item('34377').Value }                      Else {" "}
                        'Byte_LocalizedCameraModel'     = If ($image.Properties.Exists('50709')) { $image.Properties.Item('50709').Value }                      Else {" "}
                        'Byte_OriginalRawFileName'      = If ($image.Properties.Exists('50827')) { $image.Properties.Item('50827').Value }                      Else {" "}
                        'Byte_PreviewApplicationName'   = If ($image.Properties.Exists('50966')) { $image.Properties.Item('50966').Value }                      Else {" "}
                        'Byte_PreviewApplicationVersion'= If ($image.Properties.Exists('50967')) { $image.Properties.Item('50967').Value }                      Else {" "}
                        'Byte_PreviewSettingsDigest'    = If ($image.Properties.Exists('50969')) { $image.Properties.Item('50969').Value }                      Else {" "}
                        'Byte_PreviewSettingsName'      = If ($image.Properties.Exists('50968')) { $image.Properties.Item('50968').Value }                      Else {" "}
                        'Byte_ProfileCalibrationSignat' = If ($image.Properties.Exists('50932')) { $image.Properties.Item('50932').Value }                      Else {" "}
                        'Byte_ProfileCopyright'         = If ($image.Properties.Exists('50942')) { $image.Properties.Item('50942').Value }                      Else {" "}
                        'Byte_ProfileName'              = If ($image.Properties.Exists('50936')) { $image.Properties.Item('50936').Value }                      Else {" "}
                        'Byte_RawDataUniqueID'          = If ($image.Properties.Exists('50781')) { $image.Properties.Item('50781').Value }                      Else {" "}
                        'Byte_TIFFEPStandardID'         = If ($image.Properties.Exists('37398')) { $image.Properties.Item('37398').Value }                      Else {" "}
                        'Byte_XMLPacket'                = If ($image.Properties.Exists('700'))   { $image.Properties.Item('700').Value }                        Else {" "}
                        'CalibrationIlluminant1'        = If ($image.Properties.Exists('50778')) { $image.Properties.Item('50778').Value }                      Else {" "}
                        'CalibrationIlluminant2'        = If ($image.Properties.Exists('50779')) { $image.Properties.Item('50779').Value }                      Else {" "}
                        'CameraCalibration1'            = If ($image.Properties.Exists('50723')) { $image.Properties.Item('50723').Value }                      Else {" "}
                        'CameraCalibration2'            = If ($image.Properties.Exists('50724')) { $image.Properties.Item('50724').Value }                      Else {" "}
                        'CameraOwnerName'               = If ($image.Properties.Exists('42032')) { $image.Properties.Item('42032').Value }                      Else {" "}
                        'CameraSerialNumber'            = If ($image.Properties.Exists('50735')) { $image.Properties.Item('50735').Value }                      Else {" "}
                        'CellLength'                    = If ($image.Properties.Exists('265'))   { $image.Properties.Item('265').Value }                        Else {" "}
                        'CellWidth'                     = If ($image.Properties.Exists('264'))   { $image.Properties.Item('264').Value }                        Else {" "}
                        'CFALayout'                     = If ($image.Properties.Exists('50711')) { $image.Properties.Item('50711').Value }                      Else {" "}
                        'CFAPattern'                    = If ($image.Properties.Exists('41730')) { $image.Properties.Item('41730').Value }                      Else {" "}
                        'CFARepeatPatternDim'           = If ($image.Properties.Exists('33421')) { $image.Properties.Item('33421').Value }                      Else {" "}
                        'ChromaBlurRadius'              = If ($image.Properties.Exists('50737')) { $image.Properties.Item('50737').Value }                      Else {" "}
                        'ColorimetricReference'         = If ($image.Properties.Exists('50879')) { $image.Properties.Item('50879').Value }                      Else {" "}
                        'ColorMap'                      = If ($image.Properties.Exists('320'))   { $image.Properties.Item('320').Value }                        Else {" "}
                        'ColorMatrix1'                  = If ($image.Properties.Exists('50721')) { $image.Properties.Item('50721').Value }                      Else {" "}
                        'ColorMatrix2'                  = If ($image.Properties.Exists('50722')) { $image.Properties.Item('50722').Value }                      Else {" "}
                        'ComponentsConfiguration'       = If ($image.Properties.Exists('37121')) { (($image.Properties.Item('37121').Value) -join ("")).Replace("0","-").Replace("1","Y").Replace("2","Cb").Replace("3","Cr").Replace("4","R").Replace("5","G").Replace("6","B").Replace(" ","") }        Else {" "}
                        'CompressedBitsPerPixel'        = If ($image.Properties.Exists('37122')) { ($image.Properties.Item('37122').Value) | select -ExpandProperty Value } Else {" "}
                        'Compression'                   = If ($image.Properties.Exists('259'))   { $compress[[int]$image.Properties.Item('259').Value] }        Else {" "}
                        'Copyright'                     = If ($image.Properties.Exists('33432')) { $image.Properties.Item('33432').Value }                      Else {" "}
                        'CurrentICCProfile'             = If ($image.Properties.Exists('50833')) { $image.Properties.Item('50833').Value }                      Else {" "}
                        'CurrentPreProfileMatrix'       = If ($image.Properties.Exists('50834')) { $image.Properties.Item('50834').Value }                      Else {" "}
                        'DefaultCropOrigin'             = If ($image.Properties.Exists('50719')) { $image.Properties.Item('50719').Value }                      Else {" "}
                        'DefaultCropSize'               = If ($image.Properties.Exists('50720')) { $image.Properties.Item('50720').Value }                      Else {" "}
                        'DefaultScale'                  = If ($image.Properties.Exists('50718')) { $image.Properties.Item('50718').Value }                      Else {" "}
                        'DeviceSettingDescription'      = If ($image.Properties.Exists('41995')) { $image.Properties.Item('41995').Value }                      Else {" "}
                        'DocumentName'                  = If ($image.Properties.Exists('269'))   { $image.Properties.Item('269').Value }                        Else {" "}
                        'ExifTag'                       = If ($image.Properties.Exists('34665')) { $image.Properties.Item('34665').Value }                      Else {" "}
                        'ExifVersion'                   = If ($image.Properties.Exists('36864')) { $ExifVer }                                                   Else {" "}
                        'ExposureIndex_2'               = If ($image.Properties.Exists('41493')) { $image.Properties.Item('41493').Value }                      Else {" "}
                        'ExposureIndex'                 = If ($image.Properties.Exists('37397')) { $image.Properties.Item('37397').Value }                      Else {" "}
                        'ExposureProgram'               = If ($image.Properties.Exists('34850')) { $exposurep[[int]$image.Properties.Item('34850').Value] }     Else {" "}
                        'ExtraSamples'                  = If ($image.Properties.Exists('338'))   { $extrasamp[[int]$image.Properties.Item('338').Value] }       Else {" "}
                        'FillOrder'                     = If ($image.Properties.Exists('266'))   { $fillord[[int]$image.Properties.Item('266').Value] }         Else {" "}
                        'FlashEnergy_2'                 = If ($image.Properties.Exists('41483')) { $image.Properties.Item('41483').Value }                      Else {" "}
                        'FlashEnergy'                   = If ($image.Properties.Exists('37387')) { $image.Properties.Item('37387').Value }                      Else {" "}
                        'FlashpixVersion'               = If ($image.Properties.Exists('40960')) { $FlashpixVer }                                               Else {" "}
                        'FocalPlaneResolutionUnit'      = If ($image.Properties.Exists('37392')) { $focalprun[[int]$image.Properties.Item('37392').Value] }     Else {" "}
                        'FocalPlaneXResolution'         = If ($image.Properties.Exists('37390')) { $image.Properties.Item('37390').Value }                      Else {" "}
                        'FocalPlaneYResolution'         = If ($image.Properties.Exists('37391')) { $image.Properties.Item('37391').Value }                      Else {" "}
                        'ForwardMatrix1'                = If ($image.Properties.Exists('50964')) { $image.Properties.Item('50964').Value }                      Else {" "}
                        'ForwardMatrix2'                = If ($image.Properties.Exists('50965')) { $image.Properties.Item('50965').Value }                      Else {" "}
                        'FreeOffsets'                   = If ($image.Properties.Exists('288'))   { $image.Properties.Item('288').Value }                        Else {" "}
                        'FreeByteCounts'                = If ($image.Properties.Exists('289'))   { $image.Properties.Item('289').Value }                        Else {" "}
                        'GPSAltitude'                   = If ($image.Properties.Exists('6'))     { ($image.Properties.Item('6').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSAreaInformation'            = If ($image.Properties.Exists('28'))    { $gpsarea }                                                   Else {" "}
                        'GPSDateStamp'                  = If ($image.Properties.Exists('29'))    { $image.Properties.Item('29').Value }                         Else {" "}
                        'GPSDestBearing'                = If ($image.Properties.Exists('24'))    { $image.Properties.Item('24').Value }                         Else {" "}
                        'GPSDestBearingRef'             = If ($image.Properties.Exists('23'))    { $gpsdirect["$($image.Properties.Item('23').Value)"] }        Else {" "}
                        'GPSDestDistance'               = If ($image.Properties.Exists('26'))    { $image.Properties.Item('26').Value }                         Else {" "}
                        'GPSDestDistanceRef'            = If ($image.Properties.Exists('25'))    { $gpsdistrf["$($image.Properties.Item('25').Value)"] }        Else {" "}
                        'GPSDestLatitude'               = If ($image.Properties.Exists('20'))    { $image.Properties.Item('20').Value }                         Else {" "}
                        'GPSDestLatitudeRef'            = If ($image.Properties.Exists('19'))    { $image.Properties.Item('19').Value }                         Else {" "}
                        'GPSDestLongitude'              = If ($image.Properties.Exists('22'))    { $image.Properties.Item('22').Value }                         Else {" "}
                        'GPSDestLongitudeRef'           = If ($image.Properties.Exists('21'))    { $image.Properties.Item('21').Value }                         Else {" "}
                        'GPSDifferential'               = If ($image.Properties.Exists('30'))    { $gpsdiffer[[int]$image.Properties.Item('30').Value] }        Else {" "}
                   #     'GPSDOP'                        = If ($image.Properties.Exists('11'))    { $image.Properties.Item('11').Value }                         Else {" "}
                        'GPSHPositioningError'          = If ($image.Properties.Exists('31'))    { $image.Properties.Item('31').Value }                         Else {" "}
                        'GPSImgDirection'               = If ($image.Properties.Exists('17'))    { ($image.Properties.Item('17').Value) | select -ExpandProperty Value } Else {" "}
                        'GPSImgDirectionRef'            = If ($image.Properties.Exists('16'))    { $gpsdirect["$($image.Properties.Item('16').Value)"] }        Else {" "}
                        'GPSLatitude'                   = If ($image.Properties.Exists('2'))     { (($image.Properties.Item('2').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLatitudeRef'                = If ($image.Properties.Exists('1'))     { $image.Properties.Item('1').Value }                          Else {" "}
                        'GPSLongitude'                  = If ($image.Properties.Exists('4'))     { (($image.Properties.Item('4').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'GPSLongitudeRef'               = If ($image.Properties.Exists('3'))     { $image.Properties.Item('3').Value }                          Else {" "}
                        'GPSMapDatum'                   = If ($image.Properties.Exists('18'))    { $image.Properties.Item('18').Value }                         Else {" "}
                        'GPSMeasureMode'                = If ($image.Properties.Exists('10'))    { $gpsmeasure[[int]$image.Properties.Item('10').Value] }       Else {" "}
                        'GPSProcessingMethod'           = If ($image.Properties.Exists('27'))    { $gpsproc }                                                   Else {" "}
                        'GPSSatellites'                 = If ($image.Properties.Exists('8'))     { $image.Properties.Item('8').Value }                          Else {" "}
                        'GPSSpeed'                      = If ($image.Properties.Exists('13'))    { $image.Properties.Item('13').Value }                         Else {" "}
                        'GPSSpeedRef'                   = If ($image.Properties.Exists('12'))    { $gpsspeedrf["$($image.Properties.Item('12').Value)"] }       Else {" "}
                        'GPSStatus'                     = If ($image.Properties.Exists('9'))     { $gpsstat["$($image.Properties.Item('9').Value)"] }           Else {" "}
                        'GPSTag'                        = If ($image.Properties.Exists('34853')) { $image.Properties.Item('34853').Value }                      Else {" "}
                        'GPSTimeStamp'                  = If ($image.Properties.Exists('7'))     { (($image.Properties.Item('7').Value) | select -ExpandProperty Value) -join (':') } Else {" "}
                        'GPSTrack'                      = If ($image.Properties.Exists('15'))    { $image.Properties.Item('15').Value }                         Else {" "}
                        'GPSTrackRef'                   = If ($image.Properties.Exists('14'))    { $gpsdirect["$($image.Properties.Item('14').Value)"] }        Else {" "}
                        'GrayResponseCurve'             = If ($image.Properties.Exists('291'))   { $image.Properties.Item('291').Value }                        Else {" "}
                        'GrayResponseUnit'              = If ($image.Properties.Exists('290'))   { $grayrespun[[int]$image.Properties.Item('290').Value] }      Else {" "}
                        'HalftoneHints'                 = If ($image.Properties.Exists('321'))   { $image.Properties.Item('321').Value }                        Else {" "}
                        'HostComputer'                  = If ($image.Properties.Exists('316'))   { $image.Properties.Item('316').Value }                        Else {" "}
                        'ImageDescription'              = If ($image.Properties.Exists('270'))   { $image.Properties.Item('270').Value }                        Else {" "}
                        'ImageHistory'                  = If ($image.Properties.Exists('37395')) { $image.Properties.Item('37395').Value }                      Else {" "}
                        'ImageID'                       = If ($image.Properties.Exists('32781')) { $image.Properties.Item('32781').Value }                      Else {" "}
                        'ImageLength'                   = If ($image.Properties.Exists('257'))   { $image.Properties.Item('257').Value }                        Else {" "}
                        'ImageNumber'                   = If ($image.Properties.Exists('37393')) { $image.Properties.Item('37393').Value }                      Else {" "}
                        'ImageUniqueID'                 = If ($image.Properties.Exists('42016')) { $image.Properties.Item('42016').Value }                      Else {" "}
                        'ImageWidth'                    = If ($image.Properties.Exists('256'))   { $image.Properties.Item('256').Value }                        Else {" "}
                        'Indexed'                       = If ($image.Properties.Exists('346'))   { $indexd[[int]$image.Properties.Item('346').Value] }          Else {" "}
                        'InkNames'                      = If ($image.Properties.Exists('333'))   { $image.Properties.Item('333').Value }                        Else {" "}
                        'InkSet'                        = If ($image.Properties.Exists('332'))   { $incst[[int]$image.Properties.Item('332').Value] }           Else {" "}
                        'Interlace'                     = If ($image.Properties.Exists('34857')) { $image.Properties.Item('34857').Value }                      Else {" "}
                        'InteroperabilityTag'           = If ($image.Properties.Exists('40965')) { $image.Properties.Item('40965').Value }                      Else {" "}
                        'IPTCNAA'                       = If ($image.Properties.Exists('33723')) { $image.Properties.Item('33723').Value }                      Else {" "}
                        'ISOSpeed'                      = If ($image.Properties.Exists('34867')) { $image.Properties.Item('34867').Value }                      Else {" "}
                        'ISOSpeedLatitudeyyy'           = If ($image.Properties.Exists('34868')) { $image.Properties.Item('34868').Value }                      Else {" "}
                        'ISOSpeedLatitudezzz'           = If ($image.Properties.Exists('34869')) { $image.Properties.Item('34869').Value }                      Else {" "}
                        'JPEGACTables'                  = If ($image.Properties.Exists('521'))   { $image.Properties.Item('521').Value }                        Else {" "}
                        'JPEGDCTables'                  = If ($image.Properties.Exists('520'))   { $image.Properties.Item('520').Value }                        Else {" "}
                        'JPEGInterchangeFormat'         = If ($image.Properties.Exists('513'))   { $image.Properties.Item('513').Value }                        Else {" "}
                        'JPEGInterchangeFormatLength'   = If ($image.Properties.Exists('514'))   { $image.Properties.Item('514').Value }                        Else {" "}
                        'JPEGLosslessPredictors'        = If ($image.Properties.Exists('517'))   { $image.Properties.Item('517').Value }                        Else {" "}
                        'JPEGPointTransforms'           = If ($image.Properties.Exists('518'))   { $image.Properties.Item('518').Value }                        Else {" "}
                        'JPEGProc'                      = If ($image.Properties.Exists('512'))   { $jpegpro[[int]$image.Properties.Item('512').Value] }         Else {" "}
                        'JPEGQTables'                   = If ($image.Properties.Exists('519'))   { $image.Properties.Item('519').Value }                        Else {" "}
                        'JPEGRestartInterval'           = If ($image.Properties.Exists('515'))   { $image.Properties.Item('515').Value }                        Else {" "}
                        'JPEGTables'                    = If ($image.Properties.Exists('347'))   { $image.Properties.Item('347').Value }                        Else {" "}
                        'LensInfo'                      = If ($image.Properties.Exists('50736')) { $image.Properties.Item('50736').Value }                      Else {" "}
                        'LensMake'                      = If ($image.Properties.Exists('42035')) { $image.Properties.Item('42035').Value }                      Else {" "}
                        'LensModel'                     = If ($image.Properties.Exists('42036')) { $image.Properties.Item('42036').Value }                      Else {" "}
                        'LensSerialNumber'              = If ($image.Properties.Exists('42037')) { $image.Properties.Item('42037').Value }                      Else {" "}
                        'LensSpecification'             = If ($image.Properties.Exists('42034')) { $image.Properties.Item('42034').Value }                      Else {" "}
                        'LightSource'                   = If ($image.Properties.Exists('37384')) { $lightsrc[[int]$image.Properties.Item('37384').Value] }      Else {" "}
                        'LinearizationTable'            = If ($image.Properties.Exists('50712')) { $image.Properties.Item('50712').Value }                      Else {" "}
                        'LinearResponseLimit'           = If ($image.Properties.Exists('50734')) { $image.Properties.Item('50734').Value }                      Else {" "}
                    #    'MakerNote'                     = If ($image.Properties.Exists('37500')) { $MakerNote }                                                 Else {" "}
                        'MakerNoteSafety'               = If ($image.Properties.Exists('50741')) { $image.Properties.Item('50741').Value }                      Else {" "}
                        'MaskedAreas'                   = If ($image.Properties.Exists('50830')) { $image.Properties.Item('50830').Value }                      Else {" "}
                        'NewSubfileType'                = If ($image.Properties.Exists('254'))   { $image.Properties.Item('254').Value }                        Else {" "}
                        'Noise'                         = If ($image.Properties.Exists('37389')) { $image.Properties.Item('37389').Value }                      Else {" "}
                        'NoiseProfile'                  = If ($image.Properties.Exists('51041')) { $image.Properties.Item('51041').Value }                      Else {" "}
                        'NoiseReductionApplied'         = If ($image.Properties.Exists('50935')) { $image.Properties.Item('50935').Value }                      Else {" "}
                        'NumberOfInks'                  = If ($image.Properties.Exists('334'))   { $image.Properties.Item('334').Value }                        Else {" "}
                        'OECF'                          = If ($image.Properties.Exists('34856')) { $image.Properties.Item('34856').Value }                      Else {" "}
                        'OpcodeList1'                   = If ($image.Properties.Exists('51008')) { $image.Properties.Item('51008').Value }                      Else {" "}
                        'OpcodeList2'                   = If ($image.Properties.Exists('51009')) { $image.Properties.Item('51009').Value }                      Else {" "}
                        'OpcodeList3'                   = If ($image.Properties.Exists('51022')) { $image.Properties.Item('51022').Value }                      Else {" "}
                        'OPIProxy'                      = If ($image.Properties.Exists('351'))   { $opiprox[[int]$image.Properties.Item('351').Value] }         Else {" "}
                        'OriginalRawFileData'           = If ($image.Properties.Exists('50828')) { $image.Properties.Item('50828').Value }                      Else {" "}
                        'OriginalRawFileDigest'         = If ($image.Properties.Exists('50973')) { $image.Properties.Item('50973').Value }                      Else {" "}
                        'PageNumber'                    = If ($image.Properties.Exists('297'))   { $image.Properties.Item('297').Value }                        Else {" "}
                        'PhotometricInterpretation'     = If ($image.Properties.Exists('262'))   { $photmtrint[[int]$image.Properties.Item('262').Value] }      Else {" "}
                        'PlanarConfiguration'           = If ($image.Properties.Exists('284'))   { $planarconf[[int]$image.Properties.Item('284').Value] }      Else {" "}
                        'Predictor'                     = If ($image.Properties.Exists('317'))   { $predict[[int]$image.Properties.Item('317').Value] }         Else {" "}
                        'PreviewColorSpace'             = If ($image.Properties.Exists('50970')) { $previewcol[[int]$image.Properties.Item('50970').Value] }    Else {" "}
                        'PreviewDateTime'               = If ($image.Properties.Exists('50971')) { $image.Properties.Item('50971').Value }                      Else {" "}
                        'PrimaryChromaticities'         = If ($image.Properties.Exists('319'))   { (($image.Properties.Item('319').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'PrintImageMatching'            = If ($image.Properties.Exists('50341')) { $image.Properties.Item('50341').Value }                      Else {" "}
                        'ProcessingSoftware'            = If ($image.Properties.Exists('11'))    { ($image.Properties.Item('11').Value) | select -ExpandProperty Value } Else {" "}
                        'ProfileEmbedPolicy'            = If ($image.Properties.Exists('50941')) { $image.Properties.Item('50941').Value }                      Else {" "}
                        'ProfileHueSatMapData1'         = If ($image.Properties.Exists('50938')) { $image.Properties.Item('50938').Value }                      Else {" "}
                        'ProfileHueSatMapData2'         = If ($image.Properties.Exists('50939')) { $image.Properties.Item('50939').Value }                      Else {" "}
                        'ProfileHueSatMapDims'          = If ($image.Properties.Exists('50937')) { $image.Properties.Item('50937').Value }                      Else {" "}
                        'ProfileLookTableData'          = If ($image.Properties.Exists('50982')) { $image.Properties.Item('50982').Value }                      Else {" "}
                        'ProfileLookTableDims'          = If ($image.Properties.Exists('50981')) { $image.Properties.Item('50981').Value }                      Else {" "}
                        'ProfileToneCurve'              = If ($image.Properties.Exists('50940')) { $image.Properties.Item('50940').Value }                      Else {" "}
                        'Rating'                        = If ($image.Properties.Exists('18246')) { $image.Properties.Item('18246').Value }                      Else {" "}
                        'RatingPercent'                 = If ($image.Properties.Exists('18249')) { $image.Properties.Item('18249').Value }                      Else {" "}
                        'RawImageDigest'                = If ($image.Properties.Exists('50972')) { $image.Properties.Item('50972').Value }                      Else {" "}
                        'RecommendedExposureIndex'      = If ($image.Properties.Exists('34866')) { $image.Properties.Item('34866').Value }                      Else {" "}
                        'ReductionMatrix1'              = If ($image.Properties.Exists('50725')) { $image.Properties.Item('50725').Value }                      Else {" "}
                        'ReductionMatrix2'              = If ($image.Properties.Exists('50726')) { $image.Properties.Item('50726').Value }                      Else {" "}
                        'ReferenceBlackWhite'           = If ($image.Properties.Exists('532'))   { $image.Properties.Item('532').Value }                        Else {" "}
                        'RelatedSoundFile'              = If ($image.Properties.Exists('40964')) { $image.Properties.Item('40964').Value }                      Else {" "}
                        'RowInterleaveFactor'           = If ($image.Properties.Exists('50975')) { $image.Properties.Item('50975').Value }                      Else {" "}
                        'RowsPerStrip'                  = If ($image.Properties.Exists('278'))   { $image.Properties.Item('278').Value }                        Else {" "}
                        'SampleFormat'                  = If ($image.Properties.Exists('339'))   { $image.Properties.Item('339').Value }                        Else {" "}
                        'SamplesPerPixel'               = If ($image.Properties.Exists('277'))   { $image.Properties.Item('277').Value }                        Else {" "}
                        'SceneType'                     = If ($image.Properties.Exists('41729')) { $image.Properties.Item('41729').Value }                      Else {" "}
                        'SecurityClassification'        = If ($image.Properties.Exists('37394')) { $securitycl["$($image.Properties.Item('37394').Value)"] }    Else {" "}
                        'SelfTimerMode'                 = If ($image.Properties.Exists('34859')) { $image.Properties.Item('34859').Value }                      Else {" "}
                        'Sensing Method_2'              = If ($image.Properties.Exists('37399')) { $sensmethod[[int]$image.Properties.Item('37399').Value] }    Else {" "}
                        'SensitivityType'               = If ($image.Properties.Exists('34864')) { $sensittyp[[int]$image.Properties.Item('34864').Value] }     Else {" "}
                        'ShadowScale'                   = If ($image.Properties.Exists('50739')) { $image.Properties.Item('50739').Value }                      Else {" "}
                        'SMaxSampleValue'               = If ($image.Properties.Exists('341'))   { $image.Properties.Item('341').Value }                        Else {" "}
                        'SMinSampleValue'               = If ($image.Properties.Exists('340'))   { $image.Properties.Item('340').Value }                        Else {" "}
                        'Software'                      = If ($image.Properties.Exists('305'))   { $image.Properties.Item('305').Value }                        Else {" "}
                        'SpatialFrequencyResponse_2'    = If ($image.Properties.Exists('41484')) { $image.Properties.Item('41484').Value }                      Else {" "}
                        'SpatialFrequencyResponse'      = If ($image.Properties.Exists('37388')) { $image.Properties.Item('37388').Value }                      Else {" "}
                        'SpectralSensitivity'           = If ($image.Properties.Exists('34852')) { $image.Properties.Item('34852').Value }                      Else {" "}
                        'StandardOutputSensitivity'     = If ($image.Properties.Exists('34865')) { $image.Properties.Item('34865').Value }                      Else {" "}
                        'StripByteCounts'               = If ($image.Properties.Exists('279'))   { $image.Properties.Item('279').Value }                        Else {" "}
                        'StripOffsets'                  = If ($image.Properties.Exists('273'))   { $image.Properties.Item('273').Value }                        Else {" "}
                        'SubfileType'                   = If ($image.Properties.Exists('255'))   { $subfiletyp[[int]$image.Properties.Item('255').Value] }      Else {" "}
                        'SubIFDs'                       = If ($image.Properties.Exists('330'))   { $image.Properties.Item('330').Value }                        Else {" "}
                        'SubjectArea'                   = If ($image.Properties.Exists('37396')) { $image.Properties.Item('37396').Value }                      Else {" "}
                        'SubjectDistance'               = If ($image.Properties.Exists('37382')) { $image.Properties.Item('37382').Value }                      Else {" "}
                        'SubjectLocation'               = If ($image.Properties.Exists('41492')) { $image.Properties.Item('41492').Value }                      Else {" "}
                        'SubSecTime'                    = If ($image.Properties.Exists('37520')) { $image.Properties.Item('37520').Value }                      Else {" "}
                        'SubSecTimeDigitized'           = If ($image.Properties.Exists('37522')) { $image.Properties.Item('37522').Value }                      Else {" "}
                        'SubSecTimeOriginal'            = If ($image.Properties.Exists('37521')) { $image.Properties.Item('37521').Value }                      Else {" "}
                        'SubTileBlockSize'              = If ($image.Properties.Exists('50974')) { $image.Properties.Item('50974').Value }                      Else {" "}
                        'T4Options'                     = If ($image.Properties.Exists('292'))   { $image.Properties.Item('292').Value }                        Else {" "}
                        'T6Options'                     = If ($image.Properties.Exists('293'))   { $image.Properties.Item('293').Value }                        Else {" "}
                        'TargetPrinter'                 = If ($image.Properties.Exists('337'))   { $image.Properties.Item('337').Value }                        Else {" "}
                        'Thresholding'                  = If ($image.Properties.Exists('263'))   { $threshold[[int]$image.Properties.Item('263').Value] }       Else {" "}
                        'TileByteCounts'                = If ($image.Properties.Exists('325'))   { $image.Properties.Item('325').Value }                        Else {" "}
                        'TileLength'                    = If ($image.Properties.Exists('323'))   { $image.Properties.Item('323').Value }                        Else {" "}
                        'TileOffsets'                   = If ($image.Properties.Exists('324'))   { $image.Properties.Item('324').Value }                        Else {" "}
                        'TileWidth'                     = If ($image.Properties.Exists('322'))   { $image.Properties.Item('322').Value }                        Else {" "}
                        'TimeZoneOffset'                = If ($image.Properties.Exists('34858')) { $image.Properties.Item('34858').Value }                      Else {" "}
                        'TransferFunction'              = If ($image.Properties.Exists('301'))   { $image.Properties.Item('301').Value }                        Else {" "}
                        'TransferRange'                 = If ($image.Properties.Exists('342'))   { $image.Properties.Item('342').Value }                        Else {" "}
                        'UniqueCameraModel'             = If ($image.Properties.Exists('50708')) { $image.Properties.Item('50708').Value }                      Else {" "}
                        'UserComment'                   = If ($image.Properties.Exists('37510')) { $usercomment }                                              Else {" "}
                        'WhiteLevel'                    = If ($image.Properties.Exists('50717')) { $image.Properties.Item('50717').Value }                      Else {" "}
                        'WhitePoint'                    = If ($image.Properties.Exists('318'))   { (($image.Properties.Item('318').Value) | select -ExpandProperty Value) -join (" ") } Else {" "}
                        'XClipPathUnits'                = If ($image.Properties.Exists('344'))   { $image.Properties.Item('344').Value }                        Else {" "}
                        'YCbCrCoefficients'             = If ($image.Properties.Exists('529'))   { $image.Properties.Item('529').Value }                        Else {" "}
                        'YCbCrPositioning'              = If ($image.Properties.Exists('531'))   { $ycbcrpos[[int]$image.Properties.Item('531').Value] }        Else {" "}
                        'YCbCrSubSampling'              = If ($image.Properties.Exists('530'))   { $image.Properties.Item('530').Value }                        Else {" "}
                        'YClipPathUnits'                = If ($image.Properties.Exists('345'))   { $image.Properties.Item('345').Value }                        Else {" "}


                        # Source: http://nicholasarmstrong.com/2010/02/exif-quick-reference/
                        # Source: https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
                        # Credit: Franck Richard: "Use PowerShell to Remove Metadata and Resize Images": http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
                        'Acceleration'                  = If ($image.Properties.Exists('37892')) { $image.Properties.Item('37892').Value }                      Else {" "}
                        'Aperture'                      = If ($image.Properties.Exists('37378')) { ($image.Properties.Item('37378').Value).Value }              Else {" "}
                        'ApplicationNotes'              = If ($image.Properties.Exists('700'))   { $image.Properties.Item('700').Value }                        Else {" "}
                        'BadFaxLines'                   = If ($image.Properties.Exists('326'))   { $image.Properties.Item('326').Value }                        Else {" "}
                        'CameraElevationAngle'          = If ($image.Properties.Exists('37893')) { $image.Properties.Item('37893').Value }                      Else {" "}
                        'CameraLabel'                   = If ($image.Properties.Exists('51105')) { $image.Properties.Item('51105').Value }                      Else {" "}
                        'CleanFaxData'                  = If ($image.Properties.Exists('327'))   { $clnfaxdata[[int]$image.Properties.Item('327').Value] }      Else {" "}
                        'CodingMethods'                 = If ($image.Properties.Exists('403'))   { $image.Properties.Item('403').Value }                        Else {" "}
                        'Color Space'                   = If ($image.Properties.Exists('40961')) { (($image.Properties.Item('40961').Value) -join ("")).Replace("1","sRGB").Replace("FFFF.H","Uncalibrated") } Else {" "}
                        'ColorResponseUnit'             = If ($image.Properties.Exists('300'))   { $image.Properties.Item('300').Value }                        Else {" "}
                        'ConsecutiveBadFaxLines'        = If ($image.Properties.Exists('328'))   { $image.Properties.Item('328').Value }                        Else {" "}
                        'Contrast'                      = If ($image.Properties.Exists('41992')) { $contrast[[int]$image.Properties.Item('41992').Value] }      Else {" "}
                        'Custom Rendered'               = If ($image.Properties.Exists('41985')) { $custrender[[int]$image.Properties.Item('41985').Value] }    Else {" "}
                        'Date Digitized'                = If ($image.Properties.Exists('36868')) { $image.Properties.Item('36868').Value }                      Else {" "}
                        'Date Taken'                    = If ($image.Properties.Exists('36867')) { $image.Properties.Item('36867').Value }                      Else {" "}
                        'Decode'                        = If ($image.Properties.Exists('433'))   { $image.Properties.Item('433').Value }                        Else {" "}
                        'DefaultImageColor'             = If ($image.Properties.Exists('434'))   { $image.Properties.Item('434').Value }                        Else {" "}
                        'Digital Zoom Ratio'            = If ($image.Properties.Exists('41988')) { ($image.Properties.Item('41988').Value).Value }              Else {" "}
                        'FocalLength'                   = If ($image.Properties.Exists('37386')) { ($image.Properties.Item('37386').Value).Value }              Else {" "}
                        'Equipment Maker'               = If ($image.Properties.Exists('271'))   { $image.Properties.Item('271').Value }                        Else {" "}
                        'Equipment Model'               = If ($image.Properties.Exists('272'))   { $image.Properties.Item('272').Value }                        Else {" "}
                        'ExpandLens'                    = If ($image.Properties.Exists('44993')) { $image.Properties.Item('44993').Value }                      Else {" "}
                        'ExpandSoftware'                = If ($image.Properties.Exists('44992')) { $image.Properties.Item('44992').Value }                      Else {" "}
                        'Exposure Compensation'         = If ($image.Properties.Exists('37380')) { ($image.Properties.Item('37380').Value).Value }              Else {" "}
                        'Exposure Mode'                 = If ($image.Properties.Exists('41986')) { $exposurem[[int]$image.Properties.Item('41986').Value] }      Else {" "}
                        'Exposure Time'                 = If ($image.Properties.Exists('33434')) { ($image.Properties.Item('33434').Value).Value }              Else {" "}
                        'F Number'                      = If ($image.Properties.Exists('33437')) { ($image.Properties.Item('33437').Value).Value }              Else {" "}
                        'FaxProfile'                    = If ($image.Properties.Exists('402'))   { $faxprofil[[int]$image.Properties.Item('402').Value] }       Else {" "}
                        'File Source'                   = If ($image.Properties.Exists('41728')) { $filesrc[[int]$image.Properties.Item('41728').Value] }       Else {" "}
                    #    'Flash'                         = If ($image.Properties.Exists('37385')) { $image.Properties.Item('37385').Value }                      Else {" "}
                        'Flash'                         = If ($image.Properties.Exists('37385')) { $flashvalue[[int]$image.Properties.Item('37385').Value] }    Else {" "}
                        'Focal Length in 35 mm Format'  = If ($image.Properties.Exists('41989')) { $image.Properties.Item('41989').Value }                      Else {" "}
                        'Focal Plane Resolution Unit'   = If ($image.Properties.Exists('41488')) { $focal[[int]$image.Properties.Item('41488').Value] }         Else {" "}
                        'Focal Plane X Resolution'      = If ($image.Properties.Exists('41486')) { $image.Properties.Item('41486').Value }                      Else {" "}
                        'Focal Plane Y Resolution'      = If ($image.Properties.Exists('41487')) { $image.Properties.Item('41487').Value }                      Else {" "}
                        'Gain Control'                  = If ($image.Properties.Exists('41991')) { $gain[[int]$image.Properties.Item('41991').Value] }          Else {" "}
                        'Gamma'                         = If ($image.Properties.Exists('42240')) { $image.Properties.Item('42240').Value }                      Else {" "}
                        'Humidity'                      = If ($image.Properties.Exists('37889')) { $image.Properties.Item('37889').Value }                      Else {" "}
                        'ImageHistory_2'                = If ($image.Properties.Exists('41491')) { $image.Properties.Item('41491').Value }                      Else {" "}
                        'ImageNumber_2'                 = If ($image.Properties.Exists('41489')) { $image.Properties.Item('41489').Value }                      Else {" "}
                        'ISO Speed'                     = If ($image.Properties.Exists('34855')) { $image.Properties.Item('34855').Value }                      Else {" "}
                        'JPEGTables_2'                  = If ($image.Properties.Exists('437'))   { $image.Properties.Item('437').Value }                        Else {" "}
                        'Maximum Aperture'              = If ($image.Properties.Exists('37381')) { ($image.Properties.Item('37381').Value).Value }              Else {" "}
                        'Metering Mode'                 = If ($image.Properties.Exists('37383')) { $metering[[int]$image.Properties.Item('37383').Value] }      Else {" "}
                        'ModeNumber'                    = If ($image.Properties.Exists('405'))   { $image.Properties.Item('405').Value }                        Else {" "}
                        'Modified Date Time'            = If ($image.Properties.Exists('306'))   { $image.Properties.Item('306').Value }                        Else {" "}
                        'Noise_2'                       = If ($image.Properties.Exists('41485')) { ($image.Properties.Item('41485').Value).Value }              Else {" "}
                        'Image Orientation'             = If ($image.Properties.Exists('274'))   { $orient[[int]$image.Properties.Item('274').Value] }          Else {" "}
                        'OffsetTime'                    = If ($image.Properties.Exists('36880')) { $image.Properties.Item('36880').Value }                      Else {" "}
                        'OffsetTimeDigitized'           = If ($image.Properties.Exists('36882')) { $image.Properties.Item('36882').Value }                      Else {" "}
                        'OffsetTimeOriginal'            = If ($image.Properties.Exists('36881')) { $image.Properties.Item('36881').Value }                      Else {" "}
                        'PageName'                      = If ($image.Properties.Exists('285'))   { $image.Properties.Item('285').Value }                        Else {" "}
                        'Pixel X Dimension'             = If ($image.Properties.Exists('40962')) { $image.Properties.Item('40962').Value }                      Else {" "}
                        'Pixel Y Dimension'             = If ($image.Properties.Exists('40963')) { $image.Properties.Item('40963').Value }                      Else {" "}
                        'Pressure'                      = If ($image.Properties.Exists('37890')) { $image.Properties.Item('37890').Value }                      Else {" "}
                        'ProfileType'                   = If ($image.Properties.Exists('401'))   { $profiletyp[[int]$image.Properties.Item('401').Value] }      Else {" "}
                        'RelatedImageFileFormat'        = If ($image.Properties.Exists('4096'))  { $image.Properties.Item('4096').Value }                       Else {" "}
                        'RelatedImageHeight'            = If ($image.Properties.Exists('4098'))  { $image.Properties.Item('4098').Value }                       Else {" "}
                        'RelatedImageWidth'             = If ($image.Properties.Exists('4097'))  { $image.Properties.Item('4097').Value }                       Else {" "}
                        'Resolution Unit'               = If ($image.Properties.Exists('296'))   { $unit[[int]$image.Properties.Item('296').Value] }            Else {" "}
                        'Saturation'                    = If ($image.Properties.Exists('41993')) { $saturation[[int]$image.Properties.Item('41993').Value] }    Else {" "}
                        'Scene Capture Type'            = If ($image.Properties.Exists('41990')) { $scene[[int]$image.Properties.Item('41990').Value] }         Else {" "}
                        'SecurityClassification_2'      = If ($image.Properties.Exists('41490')) { $image.Properties.Item('41490').Value }                      Else {" "}
                        'Sensing Method'                = If ($image.Properties.Exists('41495')) { $sensing[[int]$image.Properties.Item('41495').Value] }       Else {" "}
                        'Sharpness'                     = If ($image.Properties.Exists('41994')) { $sharpness[[int]$image.Properties.Item('41994').Value] }     Else {" "}
                        'Shutter Speed'                 = If ($image.Properties.Exists('37377')) { ($image.Properties.Item('37377').Value).Value }              Else {" "}
                        'StripRowCounts'                = If ($image.Properties.Exists('559'))   { $image.Properties.Item('559').Value }                        Else {" "}
                        'Subject Distance Range'        = If ($image.Properties.Exists('41996')) { $subjdist[[int]$image.Properties.Item('41996').Value] }      Else {" "}
                        'T82Options'                    = If ($image.Properties.Exists('435'))   { $image.Properties.Item('435').Value }                        Else {" "}
                        'Temperature'                   = If ($image.Properties.Exists('37888')) { $image.Properties.Item('37888').Value }                      Else {" "}
                        'USPTOMiscellaneous'            = If ($image.Properties.Exists('999'))   { $image.Properties.Item('999').Value }                        Else {" "}
                        'VersionYear'                   = If ($image.Properties.Exists('404'))   { $image.Properties.Item('404').Value }                        Else {" "}
                        'WaterDepth'                    = If ($image.Properties.Exists('37891')) { $image.Properties.Item('37891').Value }                      Else {" "}
                        'White Balance'                 = If ($image.Properties.Exists('41987')) { $white[[int]$image.Properties.Item('41987').Value] }         Else {" "}
                        'X Position'                    = If ($image.Properties.Exists('286'))   { $image.Properties.Item('286').Value }                        Else {" "}
                        'X Resolution'                  = If ($image.Properties.Exists('282'))   { ($image.Properties.Item('282').Value).Value }                Else {" "}
                        'XP_DIP_XML'                    = If ($image.Properties.Exists('18247')) { $image.Properties.Item('18247').Value }                      Else {" "}
                        'Y Position'                    = If ($image.Properties.Exists('287'))   { $image.Properties.Item('287').Value }                        Else {" "}
                        'Y Resolution'                  = If ($image.Properties.Exists('283'))   { ($image.Properties.Item('283').Value).Value }                Else {" "}

                } # New-Object
    } # ForEach ($picture)


                    # Close the progress bar if it has been opened
                    If (($source_files.Count) -gt 1) {
                        $task = "Finished processing the image files."
                        Write-Progress -Id $id -Activity $activity -Status $status -CurrentOperation $task -PercentComplete (($total_steps / $total_steps) * 100) -Completed
                    } # If ($source_files.Count)


} # Process




End {


    If ($results.Count -ge 1) {

        # Discard the headers that have empty values and display the file properties in a pop-up window
        # Out-GridView doesn't support that many headers...
        # Credit: Fred: "select-object | where": https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell
        $results.PSObject.TypeNames.Insert(0,"Images")
        $results_selection = $results | Select-Object | ForEach-Object {
            $properties = $_.PsObject.Properties
            ForEach ($property in $properties) {
                If ($property.Value -eq " ") { $_.PSObject.Properties.Remove($property.Name) }
            } # ForEach
            $_
        } # ForEach-Object
        $results_selection | Sort File | Out-GridView


                        # Display the results in console
                        $empty_line | Out-String
                        $results_console = $results | Sort File | Select-Object 'File','Size','Type','Width','Height'
                        $results_console | Format-Table -Auto -Wrap

                        # Display rudimentary stats in console
                        If ($results.Count -ge 2) {
                            $text = "$($results.Count) image files found."
                        } ElseIf ($results.Count -eq 1) {
                            $text = "One image file found."
                        } Else {
                            $empty_line | Out-String
                            $text = "Didn't find any image files at -Path '$Path' or defined with the -File parameter '$File' (Exit 4)."
                        } # Else (If $results.Count)
                        Write-Output $text
                        $empty_line | Out-String


            # Sound the bell if set to do so with the -Audio parameter
            # Source: https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/
            If (($Audio) -and ($results.Count -ge 1)) {
                [console]::beep(2000,830)
            } Else {
                $continue = $true
            } # Else (If $Audio)


                        # Make the log entry if the -SuppressLog parameter is not present
                        # Note: Append parameter of Export-Csv was introduced in PowerShell 3.0.
                        # Source: http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
                        # Source: https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/

                        If ($SuppressLog) {
                            $continue = $true
                        } ElseIf ( $results.Count -ge 1 ) {
                            $data = $results | Sort File | Select-Object 'Source','File','Directory','Folder','BaseName','Original File Extension','FileExtension','Size','raw_size','Type','Orientation','Width','Height','Aspect Ratio','Created','Modified','Accessed','Log Date','HorizontalResolution','VerticalResolution','PixelDepth','FrameCount','ActiveFrame','IsIndexedPixelFormat','IsAlphaPixelFormat','IsExtendedPixelFormat','IsAnimated','IsReadOnly','IsPreRelease','IsSpecialBuild','Remarks','FormatID','SHA256','Acceleration','ActiveArea','AnalogBalance','AntiAliasStrength','Aperture','ApplicationNotes','Artist','AsShotICCProfile','AsShotNeutral','AsShotPreProfileMatrix','AsShotWhiteXY','Author','BadFaxLines','BaselineExposure','BaselineNoise','BaselineSharpness','BatteryLevel','BayerGreenSplit','BestQualityScale','BitsPerSample','BlackLevel','BlackLevelDeltaH','BlackLevelDeltaV','BlackLevelRepeatDim','BodySerialNumber','BrightnessValue','Byte_AsShotProfileName','Byte_CameraCalibrationSignat','Byte_CFAPattern','Byte_CFAPlaneColor','Byte_ClipPath','Byte_DNGBackwardVersion','Byte_DNGPrivateData','Byte_DNGVersion','Byte_DotRange','Byte_GPSAltitudeRef','Byte_GPSVersionID','Byte_ImageResources','Byte_LocalizedCameraModel','Byte_OriginalRawFileName','Byte_PreviewApplicationName','Byte_PreviewApplicationVersion','Byte_PreviewSettingsDigest','Byte_PreviewSettingsName','Byte_ProfileCalibrationSignat','Byte_ProfileCopyright','Byte_ProfileName','Byte_RawDataUniqueID','Byte_TIFFEPStandardID','Byte_XMLPacket','CalibrationIlluminant1','CalibrationIlluminant2','CameraCalibration1','CameraCalibration2','CameraElevationAngle','CameraLabel','CameraOwnerName','CameraSerialNumber','CellLength','CellWidth','CFALayout','CFAPattern','CFARepeatPatternDim','ChromaBlurRadius','CleanFaxData','CodingMethods','Color Space','ColorimetricReference','ColorMap','ColorMatrix1','ColorMatrix2','Comment','Comments','CompanyName','ComponentsConfiguration','CompressedBitsPerPixel','Compression','ConsecutiveBadFaxLines','Contrast','Copyright','CurrentICCProfile','CurrentPreProfileMatrix','Custom Rendered','Date Digitized','Date Taken','Decode','DefaultCropOrigin','DefaultCropSize','DefaultImageColor','DefaultScale','DeviceSettingDescription','Digital Zoom Ratio','DocumentName','Equipment Maker','Equipment Model','ExifTag','ExifVersion','ExpandLens','ExpandSoftware','Exposure Compensation','Exposure Mode','Exposure Time','ExposureIndex','ExposureIndex_2','ExposureProgram','ExtraSamples','F Number','FaxProfile','File Source','FillOrder','Flash','FlashEnergy','FlashEnergy_2','FlashpixVersion','Focal Length in 35 mm Format','Focal Plane Resolution Unit','Focal Plane X Resolution','Focal Plane Y Resolution','FocalLength','FocalPlaneResolutionUnit','FocalPlaneXResolution','FocalPlaneYResolution','ForwardMatrix1','ForwardMatrix2','FreeOffsets','FreeByteCounts','Gain Control','Gamma','GPSAltitude','GPSAreaInformation','GPSDateStamp','GPSDestBearing','GPSDestBearingRef','GPSDestDistance','GPSDestDistanceRef','GPSDestLatitude','GPSDestLatitudeRef','GPSDestLongitude','GPSDestLongitudeRef','GPSDifferential','GPSImgDirection','GPSHPositioningError','GPSImgDirectionRef','GPSLatitude','GPSLatitudeRef','GPSLongitude','GPSLongitudeRef','GPSMapDatum','GPSMeasureMode','GPSProcessingMethod','GPSSatellites','GPSSpeed','GPSSpeedRef','GPSStatus','GPSTag','GPSTimeStamp','GPSTrack','GPSTrackRef','GrayResponseCurve','GrayResponseUnit','HalftoneHints','HostComputer','Humidity','Image Orientation','ImageDescription','ImageHistory','ImageHistory_2','ImageID','ImageLength','ImageNumber','ImageNumber_2','ImageUniqueID','ImageWidth','Indexed','InkNames','InkSet','Interlace','InteroperabilityTag','IPTCNAA','ISO Speed','ISOSpeed','ISOSpeedLatitudeyyy','ISOSpeedLatitudezzz','JPEGACTables','JPEGDCTables','JPEGInterchangeFormat','JPEGInterchangeFormatLength','JPEGLosslessPredictors','JPEGPointTransforms','JPEGProc','JPEGQTables','JPEGRestartInterval','JPEGTables','JPEGTables_2','Keywords','LegalCopyright','LegalTrademarks','LensInfo','LensMake','LensModel','LensSerialNumber','LensSpecification','LightSource','LinearizationTable','LinearResponseLimit','MakerNoteSafety','MaskedAreas','Maximum Aperture','Metering Mode','ModeNumber','Modified Date Time','NewSubfileType','Noise','Noise_2','NoiseProfile','NoiseReductionApplied','NumberOfInks','OECF','OffsetTime','OffsetTimeDigitized','OffsetTimeOriginal','OpcodeList1','OpcodeList2','OpcodeList3','OPIProxy','OriginalRawFileData','OriginalRawFileDigest','PageName','PageNumber','PhotometricInterpretation','Pixel X Dimension','Pixel Y Dimension','PlanarConfiguration','Predictor','Pressure','PreviewColorSpace','PreviewDateTime','PrimaryChromaticities','PrintImageMatching','ProcessingSoftware','ProfileEmbedPolicy','ProfileHueSatMapData1','ProfileHueSatMapData2','ProfileHueSatMapDims','ProfileLookTableData','ProfileLookTableDims','ProfileToneCurve','ProfileType','Rating','RatingPercent','RawImageDigest','RecommendedExposureIndex','ReductionMatrix1','ReductionMatrix2','ReferenceBlackWhite','RelatedImageFileFormat','RelatedImageHeight','RelatedImageWidth','RelatedSoundFile','Resolution Unit','RowInterleaveFactor','RowsPerStrip','SampleFormat','SamplesPerPixel','Saturation','Scene Capture Type','SceneType','SecurityClassification','SecurityClassification_2','SelfTimerMode','Sensing Method','Sensing Method_2','SensitivityType','ShadowScale','Sharpness','Shutter Speed','SMaxSampleValue','SMinSampleValue','Software','SpatialFrequencyResponse_2','SpatialFrequencyResponse','SpectralSensitivity','StandardOutputSensitivity','StripByteCounts','StripOffsets','StripRowCounts','SubfileType','SubIFDs','Subject Distance Range','Subject','SubjectArea','SubjectDistance','SubjectLocation','SubSecTime','SubSecTimeDigitized','SubSecTimeOriginal','SubTileBlockSize','T4Options','T6Options','T82Options','TargetPrinter','Temperature','Thresholding','TileByteCounts','TileLength','TileOffsets','TileWidth','TimeZoneOffset','Title','TransferFunction','TransferRange','UniqueCameraModel','UserComment','USPTOMiscellaneous','VersionYear','WaterDepth','White Balance','WhiteLevel','WhitePoint','X Position','X Resolution','XClipPathUnits','XP_DIP_XML','Y Position','Y Resolution','YCbCrCoefficients','YCbCrPositioning','YCbCrSubSampling','YClipPathUnits','Index','Created (UTC)','Accessed (UTC)','Modified (UTC)','Attributes','DirectoryName','Home'
                            $logfile_path = "$real_output_path\$log_filename"

                                If ((Test-Path $logfile_path) -eq $false) {
                                    # $data | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8
                                    New-Item -ItemType File -Path "$logfile_path"
                                    $data | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Select-Object -First 1 | Out-File -FilePath $logfile_path -Append -Encoding UTF8
                                    $data | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $logfile_path -Append -Encoding UTF8
                                    (Get-Content $logfile_path) | ForEach-Object { $_ -replace ('"', '') } | Out-File -FilePath $logfile_path -Force -Encoding UTF8
                                } Else {
                                    # $data | Export-Csv $logfile_path -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Append
                                    $data | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $logfile_path -Append -Encoding UTF8
                                    (Get-Content $logfile_path) | ForEach-Object { $_ -replace ('"', '') } | Out-File -FilePath $logfile_path -Force -Encoding UTF8
                                } # Else (If Test-Path $logfile_path)
                        } Else {
                            $continue = $true
                        } # Else (If $Log)

            # Open the -Output location in the File Manager, if set to do so with the -Open parameter
            If (($Open) -and ($Force)) {
                Invoke-Item $real_output_path
            } ElseIf (($Open) -and ($results.Count -ge 1)) {
                Invoke-Item $real_output_path
            } Else {
                $continue = $true
            } # Else (If $Open)

    } Else {
        $text = "Didn't process any image files (Exit 5)."
        Write-Output $text
        $empty_line | Out-String
    } # Else (If $results.Count)

} # End





# [End of Line]




<#

   _____
  / ____|
 | (___   ___  _   _ _ __ ___ ___
  \___ \ / _ \| | | | '__/ __/ _ \
  ____) | (_) | |_| | | | (_|  __/
 |_____/ \___/ \__,_|_|  \___\___|


http://powershell.com/cs/media/p/7476.aspx                                                                                                          # clayman2: "Disk Space"
http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html                                                                # Franck Richard: "Use PowerShell to Remove Metadata and Resize Images"
http://powershell.com/cs/forums/t/9685.aspx                                                                                                         # lamaar75: "Creating a Menu"
https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file                                # Twon of An: "Get the SHA1,SHA256,SHA384,SHA512,MD5 or RIPEMD160 hash of a file"
http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader                                                           # Gisli: "Unable to read an open file with binary reader"
https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell    # Fred: "select-object | where"



  _    _      _
 | |  | |    | |
 | |__| | ___| |_ __
 |  __  |/ _ \ | '_ \
 | |  | |  __/ | |_) |
 |_|  |_|\___|_| .__/
               | |
               |_|
#>

<#
.SYNOPSIS
Retrieves EXIF data properties from digital image files and saves the info to 
a CSV-file in a defined directory.

.DESCRIPTION
Get-ExifData reads digital image files and tries to retrieve EXIF data from them and
write that info to a CSV-file (exif_log.csv). The console displays rudimentary info
about the gathering process, a reduced list is displayed in a pop-up window
(Out-GridView, about 30 categories) and the CSV-file is written/updated with over
350 categories, including the GPS tags.

The list of image files to be read is constructed in the command launching
Get-ExifData by adding a full path of a folder (after -Path parameter) or by adding
a full path of individual files (after -File parameter, multiple entries separated
with a comma). The search for image files may also be done recursively by adding
the -Recurse parameter the command launching Get-ExifData. If -Path and -File
parameters are not defined, Get-ExifData reads non-recursively the image files,
which reside in the "$($env:USERPROFILE)\Pictures" folder.

By default the CSV-file (exif_log.csv) is created into the User's own picture folder
"$($env:USERPROFILE)\Pictures" and the default CSV-file destination may be changed
with the -Output parameter. Shall the CSV-file already exist, Get-ExifData tries to
add new info to the bottom of the CSV-file rather than overwrite the CSV-file.
If the user wishes not to create any logs (exif_log.csv) or update any existing
(exif_log.csv) files, the -SuppressLog parameter may be added to the command
launching Get-ExifData.

The other available parameters (-Force, -Open and -Audio) are discussed in greater
detail below. Please note, that if any of the individual parameter values include
space characters, the individual value should be enclosed in quotation marks (single
or double), so that PowerShell can interpret the command correctly.

.PARAMETER Path
with aliases -Directory, -DirectoryPath, -Folder and -FolderPath. Specifies the
primary folder, from which the image files are checked for their EXIF data. The
default -Path parameter is "$($env:USERPROFILE)\Pictures", which will be used, if
no value for the -Path or the -File parameters is included in the command
launching Get-ExifData.

The value for the -Path parameter should be a valid file system path pointing to
a directory (a full path of a folder such as C:\Users\Dropbox\). Furthermore, if
the path includes space characters, please enclose the path in quotation marks
(single or double). Multiple entries may be entered, if they are separated with
a comma.

.PARAMETER File
with aliases -SourceFile, -FilePath and -Files. Specifies, which image files are
checked for their EXIF data.

The value for the -File parameter should be a valid full file system path pointing
to a file (with a full path name of a folder such as C:\Windows\explorer.exe).
Furthermore, if the path includes space characters, please enclose the path in
quotation marks (single or double). Multiple entries may be entered, if they are
separated with a comma.

.PARAMETER Output
with aliases -OutputFolder and -LogFileFolder. Defines the folder/directory, where
the CSV-file is created or updated.  The default -Output parameter is
"$($env:USERPROFILE)\Pictures", which will be used, if no value for the -Output is
included in the command launching Get-ExifData.

The value for the -Output parameter should be a valid file system path pointing to
a directory (a full path of a folder such as C:\Users\Dropbox\). Furthermore, if
the path includes space characters, please enclose the path in quotation marks
(single or double).

The log file file name (exif_log.csv) is defined on row 78 with $log_filename
variable and is thus "hard coded" into the script. The produced log file is UTF-8
encoded CSV-file with semi-colon as the separator.

.PARAMETER Recurse
The search for image files is done recursively, i.e. if a folder/directory is found,
all the subsequent subfolders and the image files that reside within those
subfolders (and in the subfolders of the subfolders' subfolders, and their
subfolders and so forth...) are included in the EXIF data gathering process. Please
note, that with great many image files, Get-ExifData may take some time to process
each and every file.

.PARAMETER SuppressLog
with aliases -Silent, -NoLog and -DoNotCreateALog. By adding -SuppressLog to the
command launching Get-ExifData, the CSV-file (exif_log.csv) is not created, touched
nor updated.

.PARAMETER Open
If the -Open parameter is used in the command launching Get-ExifData and new
EXIF data is found, the CSV-file destination folder (which is defined with
the -Output parameter) is opened in the default File Manager.

.PARAMETER Force
The -Force parameter affects the behaviour of Get-ExifData in two ways. If the
-Force parameter is used with the...


    1.  -Output parameter, the CSV-file destination folder (defined with the
        -Output parameter) is created, without asking any further confirmations
        from the end-user. The new folder is created with the command
        New-Item "$Output" -ItemType Directory -Force which may not be powerfull
        enough to create a new folder inside an arbitrary (system) folder.
        The Get-ExifData may gain additional rights, if it's run in an elevated
        PowerShell window (but for the most cases that is not needed at all).
    2.  -Open parameter, the CSV-file destination folder (defined with the
        -Output parameter) is opened in the default File Manager regardless 
        whether any new EXIF data was found or not.

.PARAMETER Audio
If the -Audio parameter is used in the command launching Get-ExifData and new
EXIF data is found, an audible beep will occur.

.OUTPUTS
Displays a summary of the actions in console. Displays a reduced EXIF data list
in a pop-up window (Out-GridView). Writes or updates a CSV log file at the path
defined with the -Output parameter, if the -SuppressLog parameter is not used.


    Default values:

    Path                                        Parameter       Type

    "$($env:USERPROFILE)\Pictures\exif_log.csv" : -Output       : CSV log file containing EXIF data
                                                  (concerning
                                                  the folder)
    "$($env:USERPROFILE)\Pictures"              : -Path         : The folder where the image files 
                                                                  are searched for their EXIF data,
                                                                  if no -Path or -File parameter
                                                                  is used (a non-recursive search).


.NOTES
Please note that all the parameters can be used in one get EXIF data command, and
that each of the parameters can be "tab completed" before typing them fully (by
pressing the [tab] key).

The MakerNote EXIF tag is not included, since that's camera maker specific 
information. For more information, please see for instance the Tag ID 0x927c 
at https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html

    Homepage:           https://github.com/auberginehill/get-exif-data
    Short URL:          http://tinyurl.com/ycbhtpba
    Version:            1.0

.EXAMPLE
./Get-ExifData.ps1
Runs the script. Please notice to insert ./ or .\ before the script name. Tries to
read the EXIF data from image files that reside in the
"$($env:USERPROFILE)\Pictures\" folder, since no values for the -Path or -File 
parameters were defined. Saves or updates the CSV log file (exif_log.csv) at the
default -Output folder ("$($env:USERPROFILE)\Pictures\exif_log.csv") - a file that 
contains all the gathered EXIF info columns/data types. A pop-up window listing
a partial list of the EXIF info will open, if image files were read. The console
will show rudimentary stats about the EXIF data gathering procedure.

.EXAMPLE
help ./Get-ExifData -Full
Displays the help file.

.EXAMPLE
.\Get-ExifData.ps1 -Path "C:\Users\Dropbox\" -Output "C:\Users\Dropbox\dc01" -Audio -Open -Recurse -Force
Runs the script and tries to recursively search for image files at
"C:\Users\Dropbox\" and read the EXIF info of the found image files and either
(1) update the CSV log file (exif_log.csv) at the C:\Users\Dropbox\dc01 folder if 
the folder exists or (2) create the C:\Users\Dropbox\dc01 folder (and exif_log.csv) 
without asking any further questions, if the -Output directory doesn't exist (since
the -Force was used). Also, since the -Force and -Open parameters were used, the
default File Manager will be opened at C:\Users\Dropbox\dc01 regardless whether any
image files were read or not. Furthermore, if new image files were indeed read,
an audible beep will occur. Also, a pop-up window listing a partial list of the EXIF
info will open, if image files were read, and the console will show rudimentary stats
about the EXIF data gathering procedure.

.EXAMPLE
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
This command is altering the Windows PowerShell rights to enable script execution
in the default (LocalMachine) scope, and defines the conditions under which Windows
PowerShell loads configuration files and runs scripts in general. In Windows Vista
and later versions of Windows, for running commands that change the execution policy
of the LocalMachine scope, Windows PowerShell has to be run with elevated rights
(Run as Administrator). The default policy of the default (LocalMachine) scope is
"Restricted", and a command "Set-ExecutionPolicy Restricted" will "undo" the changes
made with the original example above (had the policy not been changed before...).
Execution policies for the local computer (LocalMachine) and for the current user
(CurrentUser) are stored in the registry (at for instance the
HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ExecutionPolicy key), and remain
effective until they are changed again. The execution policy for a particular session
(Process) is stored only in memory, and is discarded when the session is closed.


    Parameters:

    Restricted      Does not load configuration files or run scripts, but permits
                    individual commands. Restricted is the default execution policy.

    AllSigned       Scripts can run. Requires that all scripts and configuration
                    files be signed by a trusted publisher, including the scripts
                    that have been written on the local computer. Risks running
                    signed, but malicious, scripts.

    RemoteSigned    Requires a digital signature from a trusted publisher on scripts
                    and configuration files that are downloaded from the Internet
                    (including e-mail and instant messaging programs). Does not
                    require digital signatures on scripts that have been written on
                    the local computer. Permits running unsigned scripts that are
                    downloaded from the Internet, if the scripts are unblocked by
                    using the Unblock-File cmdlet. Risks running unsigned scripts
                    from sources other than the Internet and signed, but malicious,
                    scripts.

    Unrestricted    Loads all configuration files and runs all scripts.
                    Warns the user before running scripts and configuration files
                    that are downloaded from the Internet. Not only risks, but
                    actually permits, eventually, running any unsigned scripts from
                    any source. Risks running malicious scripts.

    Bypass          Nothing is blocked and there are no warnings or prompts.
                    Not only risks, but actually permits running any unsigned scripts
                    from any source. Risks running malicious scripts.

    Undefined       Removes the currently assigned execution policy from the current
                    scope. If the execution policy in all scopes is set to Undefined,
                    the effective execution policy is Restricted, which is the
                    default execution policy. This parameter will not alter or
                    remove the ("master") execution policy that is set with a Group
                    Policy setting.
    __________
    Notes: 	      - Please note that the Group Policy setting "Turn on Script Execution"
                    overrides the execution policies set in Windows PowerShell in all
                    scopes. To find this ("master") setting, please, for example, open
                    the Local Group Policy Editor (gpedit.msc) and navigate to
                    Computer Configuration > Administrative Templates >
                    Windows Components > Windows PowerShell.

                  - The Local Group Policy Editor (gpedit.msc) is not available in any
                    Home or Starter edition of Windows.

                  - Group Policy setting "Turn on Script Execution":

               	    Not configured                                          : No effect, the default
                                                                              value of this setting
                    Disabled                                                : Restricted
                    Enabled - Allow only signed scripts                     : AllSigned
                    Enabled - Allow local scripts and remote signed scripts : RemoteSigned
                    Enabled - Allow all scripts                             : Unrestricted


For more information, please type "Get-ExecutionPolicy -List" -or "help Set-ExecutionPolicy -Full",
"help about_Execution_Policies" or visit https://technet.microsoft.com/en-us/library/hh849812.aspx
or http://go.microsoft.com/fwlink/?LinkID=135170.

.EXAMPLE
New-Item -ItemType File -Path C:\Temp\Get-ExifData.ps1
Creates an empty ps1-file to the C:\Temp directory. The New-Item cmdlet has an inherent
-NoClobber mode built into it, so that the procedure will halt, if overwriting (replacing
the contents) of an existing file is about to happen. Overwriting a file with the New-Item
cmdlet requires using the Force. If the path name and/or the filename includes space
characters, please enclose the whole -Path parameter value in quotation marks (single or
double):

    New-Item -ItemType File -Path "C:\Folder Name\Get-ExifData.ps1"

For more information, please type "help New-Item -Full".

.LINK
http://powershell.com/cs/media/p/7476.aspx
http://franckrichard.blogspot.com/2011/04/2011-scripting-games-advanced-event-8.html
http://powershell.com/cs/forums/t/9685.aspx
https://community.spiceworks.com/scripts/show/2263-get-the-sha1-sha256-sha384-sha512-md5-or-ripemd160-hash-of-a-file
http://stackoverflow.com/questions/8711564/unable-to-read-an-open-file-with-binary-reader
https://social.technet.microsoft.com/Forums/scriptcenter/en-US/76ae6430-4993-4422-aa97-8f8ec3ca4e87/selectobject-where?forum=winserverpowershell
http://stackoverflow.com/questions/27175137/powershellv2-remove-last-x-characters-from-a-string
http://nicholasarmstrong.com/2010/02/exif-quick-reference/
http://msdn.microsoft.com/en-us/library/ms630826%28v=vs.85%29.aspx
https://sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
https://stackoverflow.com/questions/7076958/read-exif-and-determine-if-the-flash-has-fired#7100717
https://technet.microsoft.com/en-us/library/ff730939.aspx
https://technet.microsoft.com/en-us/library/ee692804.aspx
http://kb.winzip.com/kb/entry/207/
https://msdn.microsoft.com/en-us/library/windows/desktop/ms630506(v=vs.85).aspx
https://blogs.msdn.microsoft.com/powershell/2009/03/30/image-manipulation-in-powershell/
http://stackoverflow.com/questions/4304821/get-startup-type-of-windows-service-using-powershell
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.management/get-wmiobject
https://social.microsoft.com/Forums/en-US/4dfe4eec-2b9b-4e6e-a49e-96f5a108c1c8/using-powershell-as-a-photoshop-replacement?forum=Offtopic
https://msdn.microsoft.com/en-us/library/ms630826(VS.85).aspx#SharedSample012
https://msdn.microsoft.com/en-us/powershell/reference/5.1/microsoft.powershell.utility/get-filehash
http://stackoverflow.com/questions/21252824/how-do-i-get-powershell-4-cmdlets-such-as-test-netconnection-to-work-on-windows
https://msdn.microsoft.com/en-us/library/system.security.cryptography.sha256cryptoserviceprovider(v=vs.110).aspx
https://www.experts-exchange.com/questions/25100459/I-need-to-send-the-details-of-a-jpg-file-to-an-array-any-windows-api-to-do-this-or-get-me-started.html
https://social.technet.microsoft.com/Forums/windowsserver/en-US/16124c53-4c7f-41f2-9a56-7808198e102a/attribute-seems-to-give-byte-array-how-to-convert-to-string?forum=winserverpowershell
http://compgroups.net/comp.databases.ms-access/handy-routine-for-getting-file-metad/1484921
http://www.exiv2.org/tags.html
https://sno.phy.queensu.ca/~phil/exiftool/TagNames/GPS.html
https://blogs.technet.microsoft.com/heyscriptingguy/2013/09/21/powertip-use-powershell-to-send-beep-to-console/
http://stackoverflow.com/questions/21048650/how-can-i-append-files-using-export-csv-for-powershell-2
https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/

#>
