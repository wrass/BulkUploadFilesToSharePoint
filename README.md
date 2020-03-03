This is a Powershell script that allows files to be uploaded to a SharePoint Online document library together with accompanying metadata from a CSV file.  The [SharePoint PnP library][1] is used for operations against SharePoint Online, ensure this is installed on the machine you are running the script from.

## Setup

Note that this script will not create the document library or content types for files that are uploaded, these should be set up manually beforehand.

CSV format is:
* column 0 (sourcePath): Path to the file that is to be uploaded, relative to the place you are running the PowerShell script from
* column 1 (contentType): Content Type of the file in the destination library (must exist)
* column 2 (destLibrary): Name of the Document Library in the given SharePoint site
* column 3 (destFolder): Destination folder path for the uploaded file (will be created if it does not exist)
* column 4 (fileName): Filename that the uploaded file will be given in SharePoint
* column 5 onwards: Any additional metadata columns that are configured for the Content Type

For Metadata Columns the values must be in format "termgroup|termset|term", for lookup Columns the values must be the value of the reference column defined on the Lookup Column.

Warning: You should ensure that your metadata does not contain semi-colons, as these are used by the script to split the lines of your .csv file.  Either this, or change the file delimiter used by the script.

## Usage
1. Open a Command Prompt and cd to the location of this script
2. Run: `PowerShell ./UploadFiles.ps1 -siteUrl https://yourtenancy.sharepoint.com/path/to/sharepoint/site -csvFile FilesToUpload.csv`

Credit to [João Mendes][2] for the original script that this is based on.

[1]: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps 
[2]: https://github.com/joaojmendes/
