 <#
.SYNOPSIS
    A script that sorts through a directory and pulls out all of the media files
    and places them in a \year\month\day directory structure based off the file's 
    creation date.
.DESCRIPTION
    The script recurses through a directory, looking for files that end in common
    media extensions. Depending on the argument passed in, it either copies or moves
    the files to a newly created directory structure based on the YEAR\MONTH\DAY of
    the file's creation date. The script also has the option of splitting out image
    files from video files and placing each in its own YEAR\MONTH\DAY directory scructure.

    Run with the following to redirect all system output to a log file:
    3>&1 2>&1 > redirection.log
.NOTES
    File Name      : photosort.ps1
    Author         : Warren S. Taylor (ws.taylor@gmail.com)
    Prerequisite   : PowerShell V2+
    Copyright 2018 - Warren S. Taylor
.LINK
    Script posted over:
    http://war2d2.com
.EXAMPLE
    Example 1
.EXAMPLE
    Example 2
#>
#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

### Load assembly for getting "Date Taken" metadata from images
[Reflection.Assembly]::LoadFile('C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll') | Out-Null

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------
<#
    .SYNOPSIS
    Get the Date Modified metadata from an MP4 or MOV file. Returns NULL if failed.

    .PARAMETER Filename
    Pass in the fully pathed filename of the MP3 or MOV file

    .EXAMPLE 
    VidGetExifDateModified C:\files\MyMovie.MP4

    .NOTES
    Got the main bits from https://geekeefy.wordpress.com/2016/10/15/powershell-get-mp3mp4-files-metadata-and-how-to-use-it-to-make-you-life-easy/
    His code is a little more short-cutty, and he was trying to do something different, 
    where all I wanted was the Date Modified. I had to poke around in the metadata array
    to get the right entry.
#>
function VidGetExifDateTaken($Filename) 
{
    $file = Get-Item $Filename
    $retDate = $null

    $shell = New-Object -ComObject "Shell.Application"
    $fileDir = $shell.NameSpace($file.DirectoryName)
    $Vid = ( $fileDir.Items() | Where-Object { $_.path -like $Filename } )

    # this is the index that you want from the video metadata array. Some other values of note are:
    # 0 = filename
    # 3 = Date modified
    # 208 = Media created
    $item = 208

    # If($fileDir.GetDetailsOf($Filename, $item)) #To avoid empty values
    If($fileDir.GetDetailsOf($Vid, $item)) #To avoid empty values
    {
        # $objkey = $fileDir.GetDetailsOf($Filename, $item)
        $objval = $fileDir.GetDetailsOf($Vid, $item)

        # This regex sanitizes the date returned from the COM object.
        # It was including several hidden characters, which would break date parsing.
        # The regex uses the "^" character to negate the statement after it, which basically
        # matches all Unicode letters "\p{L}" and numbers "\p{Nd}" as well as the ":" and "/"
        # characters, and blank spaces. It replaces everything NOT in that set with nothing ('')
        $objval = $objval -replace '[^\p{L}\p{Nd}/\//:/ ]', ''

        $retDate = Get-Date $objval
    }
    else 
    {
        $retDate = $null
    }

    return $retDate
}

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------
<#
    .SYNOPSIS
    Get the Date Modified metadata from an image file.

    .DESCRIPTION
    Get the Date Modified metadata from an image file. 
    Supports Jpeg, HEIC, and any other file that supports EXIF data.

    .PARAMETER image
    Pass in the fully pathed filename of the image file

    .EXAMPLE 
    GetExifDateTaken C:\files\MyImage.jpg

    .NOTES
    It appears to work fine for Jpeg and HEIC files from a variety of phones and cameras.
    A nice reference for EXIF property item values is here:
    https://nicholasarmstrong.com/2010/02/exif-quick-reference/
#>
function GetExifDateTaken($Filename) 
{
    $FileDetail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Filename 
    
    try
    {
        # $takenData = GetTakenData($FileDetail)
        if( $takenData = $FileDetail.GetPropertyItem(36867).Value )
        {
            $takenValue = [System.Text.Encoding]::Default.GetString($takenData, 0, $takenData.Length - 1)
            $taken = [DateTime]::ParseExact($takenValue, 'yyyy:MM:dd HH:mm:ss', $null)
            return $taken
        }
        else 
        {
            return $null
        }
    }
    catch
    {
        return $null
    }
    finally
    {
        $FileDetail.Dispose()
    }
}

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------
<#
    .SYNOPSIS
    Write a comment out to an established log file.

    .DESCRIPTION
    Write a comment out to an established log file.
    The log file is represented by the global $logfile variable
    
    .PARAMETER comment
    String variable passed in to be written to the log file.

    .EXAMPLE 
    WriteLog "This file has been written: $filename"

    .NOTES
    This is really hacky, and I'd like to make it more portable and/or useable. 
    For now I just want to be able to write to the logfile with a minimum of typing.
#>
function WriteLog($comment)
{
    Add-Content -Path $logFile -Value $comment #-PassThru
}

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------
#                                       M   A   I   N
#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

$srcDir = Join-Path $PSScriptRoot "\Source"
$destDir = Join-Path $PSScriptRoot "\Destination"
$photoDir = Join-Path $destDir "\Photos"
$vidDir = Join-Path $destDir "\HomeMovies"
$logDir = Join-Path $PSScriptRoot "\Logs"

$dt = Get-Date -Format "FileDate"
$logFile = Join-Path $logDir ( ("\logfile", $dt, ".txt") -join '' ) 

if( (Test-Path $logDir) -eq $false ) 
{ 
    mkdir $logDir
    WriteLog -comment "Created $logDir"
}

WriteLog -comment (Get-Date -Format o)
WriteLog -comment "Begin processing..."

if( (Test-Path $destDir) -eq $false ) 
{ 
    WriteLog -comment "Creating $destDir"
    mkdir $destDir
}

if( (Test-Path $photoDir) -eq $false ) 
{ 
    WriteLog -comment "Creating $photoDir"
    mkdir $photoDir
}

if( (Test-Path $vidDir) -eq $false ) 
{ 
    WriteLog -comment "Creating $vidDir"
    mkdir $vidDir
}

$MyFiles = get-childitem -path $srcDir\* -Recurse -include *.png, *.jpeg, *.gif, *.jpg, *.psd, *.bmp, *.heic, *.mov, *.mp4

ForEach($File in $MyFiles)
{
    # reset the destination directory each loop
    if( ($File.Extension).toupper() -EQ ".MOV" -or ($File.Extension).toupper() -EQ ".MP4")
    {
        $destPath = $vidDir # reset the destination directory each loop; set to Video
        $dateVar = VidGetExifDateTaken -Filename $File.FullName
    }
    else 
    {
        $destPath = $photoDir # reset the destination directory each loop; set to Photo
        $dateVar = GetExifDateTaken -Filename $File.FullName
    }
    
    if( $null -ne $dateVar )
    {
        $fYear = $dateVar.Year
        $fMonth = $dateVar.Month
        $fDay = $dateVar.Day
    }
    else 
    {
        # Get the YEAR\MONTH\DAY 
        $fYear = $File.LastWriteTime.Year
        $fMonth = "{0:00}" -f $File.LastWriteTime.Month
        $fDay = "{0:00}" -f $File.LastWriteTime.Day
    }

    $destPath = ($destPath, $fYear, $fMonth, $fDay) -join '\'
    if( (Test-Path $destPath) -eq $false ) 
    { 
        mkdir $destPath #| out-null 
    }

    # Get the full path of the destination + filename 
    $destFilePath = Join-Path $destPath $File.Name

    # Test to see if the file already exists in the destination directory
    if( (Test-Path $destFilePath) -eq $true ) 
    { 
        #use Get-FileHash to compare $File with $fYear\$fMonth\$fDay\$File
        $srcHash = Get-FileHash $File.FullName 

        # $destFile = ($destDir, $fYear, $fMonth, $fDay, $File.Name) -join '\'
        $destHash = Get-FileHash $destFilePath

        if($srcHash.Hash -ne $destHash.Hash)
        {
            # Files are not the same, but have the same name
            # Create "Dupes" directory in existing directory and save file there
            $destDirPath = Join-Path $destPath "Dupes" 

            if( (Test-Path $destDirPath) -eq $false ) 
            { 
                mkdir $destDirPath 
            }

            $destFilePath = Join-Path $destDirPath $File.Name

            # If file already exists in Dupes, increment name
            if( (Test-Path $destFilePath) -eq $true ) 
            {
                $comment = $File.Name + " exists in */Dupes, renaming file."
                write-host $comment -ForegroundColor Cyan 
                WriteLog -comment $comment

                $namecount = 1
                while( (Test-Path $destFilePath) -eq $true )
                {
                    $newFileName = ($File.BaseName, $namecount, $File.Extension )  -join '_'
                    $destFilePath = Join-Path $destDirPath $newFileName
                    $namecount++
                }
            }
            
            copy-item $File $destFilePath

            $comment = "Copying " + $File.BaseName + " to $destFilePath"
            write-host $comment -ForegroundColor Green 
            WriteLog -comment $comment
        }
        else 
        {
            # File already exists, write full path to log   
            $comment = "Exists: $File in $destFilePath"
            write-host $comment -ForegroundColor Green 
            WriteLog -comment $comment 
        }
    }
    else 
    {
        # If file doesn't exist, create the path and save the file
        # use New-Item to create path:
        # New-Item -ItemType Directory -Force -Path C:\Path\That\May\Or\May\Not\Exist
        copy-item $File $destPath
    }
} #end ForEach
