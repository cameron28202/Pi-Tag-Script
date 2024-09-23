Add-Type -Path "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\OSIsoft.AFSDK\v4.0_4.0.0.0__6238be57836698e6\OSIsoft.AFSDK.dll"

$afServers = New-Object OSIsoft.AF.PISystems  
$afServer = $afServers["PEMBINA_AF-DEV"]
Write-host("AFServer Name: {0}" -f $afServers.Name)

# AFDatabase
$DB = $afServer.Databases["PI_Modernization_DEV"]
Write-host("DataBase Name: {0}" -f $DB.Name)

function Add-Attribute ($element, $attributeName, $categoryName, $uom, $type, $description) {

    # Add attribute
    $template = $element.Template
    $attribute = $template.AttributeTemplates.Add($attributeName)
    $attribute.DisplayDigits = 1
    $attribute.Type = $type
    $attribute.DataReferencePlugIn = $afServer.DataReferencePlugIns["PI Point"]
    $attribute.DefaultUOM = $afServer.UOMDatabase.UOMS[$uom]
    $category = $DB.AttributeCategories[$categoryName]
    $attribute.Categories.Add($category)
    $attribute.ConfigString = "\\%@\Pembina|PIServerName%\%@%Attribute%|Tag Base%"
    $attribute.Description = $description

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

function Read-ExcelFile($FilePath) {

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($FilePath)

        Write-Host "Successfully opened the workbook."
        Write-Host "Available sheets:"
        foreach ($sheet in $workbook.Sheets) {
            Write-Host "- $($sheet.Name)"
        }

        $sheet = $workbook.WorkSheets.item(1)

        if ($sheet -eq $null) {
            throw "Sheet '$SheetName' not found in the workbook."
        }

        Write-Host "Successfully accessed sheet: $($sheet.Name)"

        $rowCount = ($sheet.UsedRange.Rows).Count
        $colCount = ($sheet.UsedRange.Columns).Count

        Write-Host "Row count: $rowCount, Column count: $colCount"

        $data = @()

        # Assuming the first row contains headers
        $headers = @()
        for ($col = 1; $col -le $colCount; $col++) {
            $headerText = $sheet.Cells.Item(2, $col).Text
            if ([string]::IsNullOrWhiteSpace($headerText)) {
                $headerText = "Column$col"
            }
            $headers += $headerText
        }

        Write-Host "Headers: $($headers -join ', ')"

        # Read data from row 3 onwards (assuming row 2 is headers)
        for ($row = 3; $row -le $rowCount; $row++) {
            $rowData = [ordered]@{}
            for ($col = 1; $col -le $colCount; $col++) {
                $cellValue = $sheet.Cells.Item($row, $col).Text
                $rowData[$headers[$col - 1]] = $cellValue
            }
            $data += [PSCustomObject]$rowData
        }
        return $data
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
                        -description $rowData.Description

            $attribute.Attributes["Label"].ConfigString = $rowData.Label
            $attribute.Attributes["Tag Base"].ConfigString = $rowData.'Tag Base'
        }

        Write-Host "Processed element: $elementName"
    } else {
        Write-Host "Element not found: $elementName"
    }
}

function Is-Empty($rowData) {
    return ($null -eq $rowData -or 
            [string]::IsNullOrWhiteSpace($rowData.'Element Path') -or 
            [string]::IsNullOrWhiteSpace($rowData.Name))
}

$excelPath = "C:\Users\SHaw\OneDrive - Pembina Pipeline Corporation\Desktop\GC_Tag_Builder.xlsx"

$data = Read-ExcelFile -FilePath $excelPath

foreach($object in $data){
    Add-Data -afDatabase $DB -rowData $object
}
$DB.CheckIn()
