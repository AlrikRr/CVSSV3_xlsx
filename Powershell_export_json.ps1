<#
RedFou52
JSON parser V1 for Serpico Project Import
Get every JSON output foreach findings
and generate a JSON file ready for the SPERPICO import
Use the file CVSS3_Serpico_Output.xlsx
#>

#Init Excel Application 
$xl = New-Object -ComObject Excel.Application
$sheets = $xl.Workbooks.Open("[PATH_to_the_xlsx]")
$sheet = $sheets.Worksheets.Item(1)

#Init variable
$debut = "["
$fin = "]"
$espace = ","
$filedate= Get-Date -Format "yyyyMMdd"
$filename= $filedate+".json"
$rowCounter = 8 #default, the first findings on the xlsx file is the on the 8th line 

#Création du fichier JSON qui porte en nom la date actuelle
New-Item -Path . -Name $filename -ItemType "file"

#Count each non-empty line (Start at cell(8,88) This is the first address of the JSON output
while($sheet.Cells.Item($rowCounter,88).text -ne ""){
    $rowCounter++
}

#Add [ at the beginning of the file

Add-Content "[JSON_PATH_export]\$filename" $debut

#Second loop that traverses the non-empty lines to retrieve the JSON code in each cell and adds it to the JSON script file

for($intRow=8;$intRow -le $rowCounter-1; $intRow++)
{
    #Get the txt value of the cell
    $value = $sheet.Cells.Item($intRow,87)
    $value.text

    #Append the txt to the file
    Add-Content "U:\Test_Powershell\$filename" $value.text

    #CONDITION: Avoid adding in the script the last comma of the last line
    if( $rowCounter-1 -ne $intRow){
        #Add a comma between each JSON element in the rows
        Add-Content "[JSON_PATH_export]\$filename" $espace
    }

   
}

#Close the JSON script with ]
Add-Content "[JSON_PATH_export]\$filename" $fin

#Close xlsx file
$sheet.close()
$xl.Quit()
