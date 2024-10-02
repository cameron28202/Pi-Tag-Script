# Title:        Tag Mapping Script
#
# Description: This script reads from an excel file to create attributes on the template level, then map a tag and label to each attribute.
#
# Params:
#   excelPath - path to excel file holding tag information
#
# Example Usage:
#   .\script.ps1 -excelPath "C:\Path\To\Your\Excel\File"

param([System.String] $excelPath)

$uomMapping = @{
    "0C" = "C"
    "kpa" = "kPa"
    "Kpad" = "kPad"
    "e3m3/day" = "e3m3/d"
    "mole %" = "mol %"
    "m3/hr" = "m3/h"
    "E3M3/D" = "e3m3/d"
    "ppmw" = "ppmwt"
}

function Read-ExcelFile($FilePath) {
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($FilePath)

        Write-Host "Successfully opened the workbook."
        $sheet = $workbook.WorkSheets.item(1)

        if ($null -eq $sheet) {
            throw "Sheet not found in the workbook."
        }

        Write-Host "Successfully accessed sheet: $($sheet.Name)"

        # Read configuration data from Row 2
        $config = @{
            'AFServerName' = $sheet.Cells.Item(2, 1).Text
            'DatabaseName' = $sheet.Cells.Item(2, 2).Text
            'AFSDKPath' = $sheet.Cells.Item(2, 3).Text
        }

        $rowCount = ($sheet.UsedRange.Rows).Count
        $colCount = ($sheet.UsedRange.Columns).Count

        $data = @()

        # Read headers from row 3
        $headers = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $headerText = $sheet.Cells.Item(3, $col).Text
            if ([string]::IsNullOrWhiteSpace($headerText)) {
                $headerText = "Column$col"
            }
            $headers += $headerText
        }

        Write-Host "Headers: $($headers -join ', ')"

        # Read data from row 4 onwards
        for ($row = 4; $row -le $rowCount; $row++) {
            $rowData = [ordered]@{}
            for ($col = 1; $col -le $colCount; $col++) {
                $cellValue = $sheet.Cells.Item($row, $col).Text
                $rowData[$headers[$col - 1]] = $cellValue
            }
            $data += [PSCustomObject]$rowData
        }
        return @{
            'Config' = $config
            'Data' = $data
        }
    }
    catch {
        Write-Error "An error occurred: $_"
        return $null
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

function Add-Attribute ($element, $attributeName, $categoryName, $uom, $type, $enumerationSet, $tagBase){

    if ($null -eq $element) {
        Write-Host "Element is null"
        return
    }
    if ($null -eq $element.Template) {
        Write-Host "Element Template is null"
        return
    }
    if ($null -eq $element.Template.AttributeTemplates) {
        Write-Host "AttributeTemplates is null"
        return
    }

    # for UOMS that don't directly translate from pi data archive to AF
    if ($uomMapping.ContainsKey($uom)) {
        $uom = $uomMapping[$uom]
    }

    # add temp, flow, pressure, control valve to the attribute name to add more context, based on UOM or tag base
    if ($uom -like "*C*" -and $attributeName -notmatch "Temp") {
        $attributeName += " Temp"
    }
    elseif (($uom -eq "e3m3/d" -or $uom -eq "m3/d" -or $uom -eq "m3/h") -and $attributeName -notmatch "Flow") {
        $attributeName += " Flow"
    }
    elseif ($uom -eq "kPa" -and $attributeName -notmatch "Press"){
        $attributeName += " Pressure"
    }

    if ($tagBase -like "*/OUT.CV" -and $attributeName -notmatch "Valve") {
        $attributeName += " Control Valve"
    }


    $existingAttribute = $element.Template.AttributeTemplates[$attributeName]
    if($null -ne $existingAttribute){
        Write-Host "Attribute '$attributeName' already exists. Skipping addition, but mapping tag/label."
        return $existingAttribute
    }


    # add attribute to template
    $template = $element.Template
    $attribute = $template.AttributeTemplates.Add($attributeName)
    $attribute.DisplayDigits = 1
    $attribute.DataReferencePlugIn = $afServer.DataReferencePlugIns["PI Point"]
    $attribute.ConfigString = "\\%@\Pembina|PIServerName%\%@%Attribute%|Tag Base%"

    # Add enumeration set, if it exists
    if ($enumerationSet -ne "" -and $null -ne $DB.EnumerationSets[$enumerationSet]){
        $enumSet = $DB.EnumerationSets[$enumerationSet]
        $attribute.TypeQualifier = $enumSet
    } 
    else {
        if ($uom -ne ""){
            $attribute.Type = $type
            $attribute.DefaultUOM = $afServer.UOMDatabase.UOMS[$uom]
        }
    }

    # add category, if it exists
    if ($categoryName -ne ""){
        $category = $DB.AttributeCategories[$categoryName]
        if($null -ne $category){
            $attribute.Categories.Add($category)
        } 
        else {
            Write-Host "Category '$categoryName' not found. Skipping category assignment."
        }
    }

    # Add Tag Base child attribute
    $tagBaseAttr = $attribute.AttributeTemplates.Add("Tag Base")
    $attribute.DisplayDigits = 1
    $tagBaseAttr.DataReferencePlugIn = $afServer.DataReferencePlugIns["String Builder"]
    $tagBaseAttr.Type = "String"

    # Add Label child attribute
    $labelAttr = $attribute.AttributeTemplates.Add("Label")
    $attribute.DisplayDigits = 1
    $labelAttr.DataReferencePlugIn = $afServer.DataReferencePlugIns["String Builder"]
    $labelAttr.Type = "String"

    
    return $attribute
}

function Get-AFElement($afDatabase, $elementPath) {
    # access the element based on path passed into excel file

    $pathParts = $elementPath -split '\\'
    $currentElement = $afDatabase.Elements
    foreach ($part in $pathParts) {
        $element = $currentElement[$part]
        if ($null -eq $element) {
            Write-Host "Element not found: $part in path $elementPath"
            return $null
        }
        $currentElement = $element.Elements
    }
    return $element
}

function Add-Data ($afDatabase, $rowData) {
    # call the add-attribute function that actually creates each attribute, then
    # map label/tag
    if(Is-Empty -rowData $rowData){
        return
    }

    $elementPath = $rowData.'Element Path'
    $elementName = $rowData.Name
    $element = Get-AFElement -afDatabase $afDatabase -elementPath $elementPath

    if ($null -ne $element) {
        Write-Host "Processing element: $elementName"
        if($rowData.Name -ne ""){
            $attribute = Add-Attribute -element $element `
                        -attributeName $rowData.Name `
                        -categoryName $rowData.Category `
                        -uom $rowData.UOM `
                        -type $rowData.Type `
                        -enumerationSet $rowData.'Enumeration Set' `
                        -tagBase $rowData.'Tag Base'

            $attribute.Attributes["Label"].ConfigString = $rowData.Label
            $attribute.Attributes["Tag Base"].ConfigString = $rowData.'Tag Base'
        }
        Write-Host "Processed element: $elementName"
    } 
    else {
        Write-Host "Element not found: $elementName"
    }
}

function Is-Empty($rowData) {
    # function to check if an excel cell is empty to determine if we want to use the row
    return ($null -eq $rowData -or 
            [string]::IsNullOrWhiteSpace($rowData.'Element Path') -or 
            [string]::IsNullOrWhiteSpace($rowData.Name) -or
            [string]::IsNullOrWhiteSpace($rowData.Selected))
}

# read from excel file
$excelData = Read-ExcelFile -FilePath $excelPath
$config = $excelData.Config
$data = $excelData.Data

# add path to the AFSDK
Add-Type -Path $config.AFSDKPath

# access AF server
$afServers = New-Object OSIsoft.AF.PISystems
$afServer = $afServers[$config.AFServerName]
Write-host("AFServer Name: {0}" -f $afServer.Name)

# access AF database
$DB = $afServer.Databases[$config.DatabaseName]
Write-host("DataBase Name: {0}" -f $DB.Name)

foreach($object in $data){
    Add-Data -afDatabase $DB -rowData $object
}
try {
    # check in changes!
    $DB.CheckIn()
    Write-Host "Successfully checked in changes to the database."
}
catch {
    Write-Error "Failed to check in changes: $_"
}

