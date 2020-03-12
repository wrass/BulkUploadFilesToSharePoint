# import files to SharePoint online document library
# Ti Marner - 2020-02-27
# Parameters:
#   siteUrl - url for the SharePoint site to upload to
#   csvFile - .csv containing a list of files to upload
#	logFile - optional filename to use for logging messages
#---------------------------------------------
#
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)] [string] $siteUrl,
    [Parameter(Mandatory = $true)] [string] $csvFile,
    [Parameter()] [string] $logFile
)

# main ImportFile function
function ImportFile {

    # log function
    function Logmsg {
        Param($msg)
        Write-Host $msg
        if ($null -ne $logFile) {
            Add-Content -Value $msg -Path $logFile
        }
    }

    Clear-Host

    $results = Test-Path $csvFile
    if ($results -eq $false) {
        throw "File $csvFile does not exist"
        exit 1
    }

    $culture = Get-Culture
    if ($culture.Name -ne "en-GB") {
        throw "Culture is not en-GB, stopping"
        exit 1
    }

    # define logfile
    try {
        $logFileCreate = New-Item -Path . -Name $logFile -ItemType "file" -Force
    } catch {
        throw "Can't write to $logFile"
        exit 1
    }
    Logmsg -msg "File import started at $(Get-Date)"
    Logmsg -msg "CSV file [$csvFile]"
    Logmsg -msg "SharePoint site: [$siteUrl]"
    if ($null -ne $logFile) {
        Logmsg -msg "Logging messages to [$logFile]"
    }

    # get columns and CSV data
    $csvData = Import-Csv -Path $csvFile -Delimiter ";"
    # get column names
    $csvColumnNames = (Get-Content -Path $csvFile | Select-Object -First 1) -split ";"
    
    # connect to SharePoint online
    Logmsg -msg "Connecting to SharePoint Online"
    try {
        Connect-PnPOnline -Url $siteUrl
    } catch {
        Write-Host "Unable to connect to $siteUrl" -ForegroundColor Black -BackgroundColor Red
        exit 1
    }

    # read files from CSV
    $i = 1
    Logmsg -msg "Starting upload"
    Logmsg -msg "-----"
    foreach($csvRow in $csvData) {
        #Logmsg -msg "$(Get-Date): uploading file $($i): $($csvRow.sourcePath)"

        # test if source file exists
        $results = Test-Path $csvRow.sourcePath
        if ($results -eq $false) {
            Logmsg -msg "$(Get-Date): ERROR - file not found, Path: $($csvRow.sourcePath)"
        } else {
            # source exists, test if document library exists
            $docLibExists = Get-PnPList $csvRow.destLibrary

            if ($null -eq $docLibExists) {
                Logmsg -msg "$(Get-Date): ERROR - library [$($csvRow.destLibrary)] not found in site [$siteUrl]"
            } else {
                # doc library is good, get columns
                $libraryColumns = Get-PnPField -List $csvRow.destLibrary.trim()
                
                # metadata hash table
                $values = @{}
                for ($j = 5; $j -lt $csvColumnNames.length; $j++) {
                    $addColumnToValues = $true
                    $csvColumnName = $csvColumnNames[$j]
                    $columnValue = $csvRow.$csvColumnName.trim()
                    $libraryColumn = $libraryColumns | Where-Object {$_.Title -eq $csvColumnName}

                    if ($null -ne $libraryColumn) {
                        # deal with lookup types
                        if ($libraryColumn.TypeAsString -eq "Lookup") {
                            if ($null -eq $columnValue) {
                                $addColumnToValues = $false
                            } else {
                                $listId = $libraryColumn.LookupList
                                $listItem = (Get-PnPListItem -List $listId).FieldValues | Where-Object {$_.Title -eq $columnValue.trim()}
                                if ($null -eq $listItem) {
                                    Logmsg -msg "$(Get-Date): ERROR - store [$columnValue] does not exist in the list of stores"
                                    $addColumnToValues = $false
                                } else {
                                    $columnValue = $listItem.ID
                                    $addColumnToValues = $true
                                }
                            }
                        }

                        # format date columns to correct culture
                        if ($libraryColumn.TypeAsString -eq "DateTime") {
                            if (($null -ne $columnValue.trim()) -and ($columnValue.trim() -ne "")) {
                                # try to convert
                                $dt = Get-Date($columnValue) -format $culture.DateTimeFormat.ShortDatePattern
                                $columnValue = [datetime]::ParseExact($dt, $culture.DateTimeFormat.ShortDatePattern, $null)
                                $addColumnToValues = $true
                            } else {
                                $addColumnToValues = $false
                            }
                        }

                        # add column to collection if it's valid
                        # use the internal column name in this library
                        if ($addColumnToValues -eq $true) {
                            $values.Add($libraryColumn.InternalName, $columnValue)
                        }
                    }
                }

                # do file upload
                try {
                    # ensure destination folder exists
                    $destPath = $csvRow.destLibrary + "/" + $csvRow.destFolder
                    $siteFolder = Resolve-PnPFolder -SiteRelativePath $destPath

                    # add file
                    if ($null -eq $siteFolder) {
                        Logmsg -msg "$(Get-Date): ERROR - cannot upload file to SharePoint site [$siteUrl], document Library [$($csvRow.destLibrary)], file [$($csvRow.sourcePath)]. Unable to create folder [$destPath]"
                    } else {
                        # set content type for file
                        #$values.Add("ContentTypeId", $???)
                        $null = Add-PnPFile -Path $csvRow.sourcePath -Folder $destPath -NewFileName $csvRow.fileName -Values $values #-ContentType $csvRow.contentType
                        Logmsg -msg "$(Get-Date): file $($i) uploaded - [$($csvRow.sourcePath)]"
                    }
                } catch {
                    Logmsg -msg "$(Get-Date): ERROR - cannot upload file to SharePoint site [$siteUrl], document Library [$($csvRow.destLibrary)], file [$($csvRow.sourcePath)]. Exception: $($_.Exception.Message)"
                }
            }
        }

        $i += 1
    }

    Logmsg -msg "--- Import finished at $(Get-Date)"
    Disconnect-PnPOnline
}

#
### Import files to SharePoint site ###
# Main #
ImportFile
