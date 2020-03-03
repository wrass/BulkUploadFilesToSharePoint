# import files to SharePoint online document library
# Ti Marner - 2020-02-27
# Parameters:
#   siteurl
#   csvfile
#---------------------------------------------
#
[CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$siteUrl,
        [Parameter(Mandatory = $true)]
        [string]$csvFile
    )
# Function
function ImportFile {
    function Logmsg {
        Param($msg)
        Write-Host $msg
        Add-Content -Value $msg -Path $logfilename
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
    $logfilename = "importFilesLog.log"
    try {
        $logFile = New-Item -Path . -Name $logfilename -ItemType "file" -Force
    } catch {
        throw "Can't write to $logfilename"
        exit 1
    }
    Logmsg -msg "File import started at $(Get-Date)"
    Logmsg -msg "CSV file: $csvFile"
    Logmsg -msg "SharePoint site: $siteUrl"

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
    $i = 0
    foreach($csvRow in $csvData) {
        Logmsg -msg "Uploading file $($i): $($csvRow.sourcePath)"
        #Logmsg -msg $csvRow

        # test if source file exists
        $results = Test-Path $csvRow.sourcePath
        if ($results -eq $false) {
            Logmsg -msg "$(Get-Date -Format "yyyy-MM-dd") - Error - file not found, Path: $($csvRow.sourcePath)"
        } else {
            # source exists, test if document library exists
            $docLibExists = Get-PnPList $csvRow.destLibrary

            if ($null -eq $docLibExists) {
                Logmsg -msg "$(Get-Date -Format "yyyy-MM-dd") - Error - library $($csvRow.destLibrary) does not exist in site: $siteUrl"
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
                                    Logmsg -msg "$(Get-Date -Format "yyyy-MM-dd") - Error - store: $columnValue does not exist in the list of stores"
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
                                #Logmsg -msg "DateTime value is $columnValue"
                                $addColumnToValues = $true
                            } else {
                                $addColumnToValues = $false
                            }
                        }

                        #Logmsg -msg "column $csvColumnName, TypeAsString is $($libraryColumn.TypeAsString), InternalName is $($libraryColumn.InternalName)"

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
                        Logmsg -msg "$(Get-Date -Format "s") - Error uploading file to SharePoint site: $siteUrl, Document Library: $($csvRow.destLibrary), file: $($csvRow.sourcePath), Unable to create folder: $destPath"
                    } else {
                        # set content type for file
                        #$values.Add("ContentTypeId", $???)
                        #Write-Host "-Path "+$csvRow.sourcePath+" -Folder "+$destPath+" -NewFileName "+$csvRow.fileName+" -Values "+$values
                        $null = Add-PnPFile -Path $csvRow.sourcePath -Folder $destPath -NewFileName $csvRow.fileName -Values $values #-ContentType $csvRow.contentType
                    }
                } catch {
                    Logmsg -msg "$(Get-Date -Format "s") - Error uploading file to SharePoint site: $siteUrl, Document Library: $($csvRow.destLibrary), file: $($csvRow.sourcePath), Error: $($_.Exception.Message)"
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
