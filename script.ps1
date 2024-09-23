$excelPath = "C:\Users\SHaw\OneDrive - Pembina Pipeline Corporation\Desktop\GC_Tag_Builder.xlsx"

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

function Add-Attribute ($element, $attributeName, $categoryName, $uom, $type, $enumerationSet){

    $existingAttribute = $element.Template.AttributeTemplates[$attributeName]
    if ($null -ne $existingAttribute) {
        Write-Host "Attribute '$attributeName' already exists. Skipping addition."
        return $existingAttribute
    }

    # Add attribute
    $template = $element.Template
    $attribute = $template.AttributeTemplates.Add($attributeName)
    $attribute.DisplayDigits = 1
    $attribute.DataReferencePlugIn = $afServer.DataReferencePlugIns["PI Point"]
    $attribute.ConfigString = "\\%@\Pembina|PIServerName%\%@%Attribute%|Tag Base%"

    # Add enumeration set, if it exists
    if ($enumerationSet -ne ""){
        $enumSet = $DB.EnumerationSets[$enumerationSet]
        if ($null -eq $enumSet) {
            Write-Host "Enumeration set '$enumerationSet' not found. Skipping enumeration set assignment."
            $attribute.Type = $type
            $attribute.DefaultUOM = $afServer.UOMDatabase.UOMS[$uom]
        } 
        else {
            $attribute.TypeQualifier = $enumSet
        }
    } 
    else {
        if ($uom -ne ""){
            $attribute.Type = $type
            $attribute.DefaultUOM = $afServer.UOMDatabase.UOMS[$uom]
        }
    }


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
                        -enumerationSet $rowData.'Enumeration Set'

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
    return ($null -eq $rowData -or 
            [string]::IsNullOrWhiteSpace($rowData.'Element Path') -or 
            [string]::IsNullOrWhiteSpace($rowData.Name) -or
            [string]::IsNullOrWhiteSpace($rowData.Selected))
}

$excelData = Read-ExcelFile -FilePath $excelPath
$config = $excelData.Config
$data = $excelData.Data

Add-Type -Path $config.AFSDKPath

$afServers = New-Object OSIsoft.AF.PISystems
$afServer = $afServers[$config.AFServerName]
Write-host("AFServer Name: {0}" -f $afServer.Name)

# AFDatabase
$DB = $afServer.Databases[$config.DatabaseName]
Write-host("DataBase Name: {0}" -f $DB.Name)

foreach($object in $data){
    Add-Data -afDatabase $DB -rowData $object
}
$DB.CheckIn()