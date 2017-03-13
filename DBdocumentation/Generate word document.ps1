# This function will get the comments on objects 
# MS calls these MS_Descriptionn when you add them through SSMS 
function GetDescriptionExtendedProperty 
{ 
    param ($item); 
    $description = ""; 
    foreach($property in $item.ExtendedProperties) 
    { 
        if($property.Name -eq "MS_Description") 
        { 
            $description = $property.Value; 
        } 
    } 
    return $description; 
} 

function CreateColumnsTable 
{ 
    param ($table, $selection); 

    $selection.TypeParagraph();
    $selection.Style = 'Subtitle with line';
    $text = "Fields"
    $selection.TypeText($text);
    $selection.TypeParagraph();
    $selection.Style = 'No Spacing';

    $columns = $table.Columns; 
    $rows = $columns.Count;

    $columnsTable = $selection.tables.add(
        $selection.Range,($rows + 1),7,
        [Microsoft.Office.Interop.word.WdDefaulttableBehavior]::wdword9tableBehavior,
        [Microsoft.Office.Interop.word.WdAutoFitBehavior]::wdAutoFitContent
    );

    $columnsTable.Style = "List Table 3 - Accent 3";
    $columnsTable.Range.Font.Size = 10;
    $columnsTable.Range.Font.Name = "Times New Roman" ;

    ## Header
    $columnsTable.cell(1,1).range.Bold=1
    $columnsTable.cell(1,1).range.text = "Name";
    $columnsTable.cell(1,2).range.Bold=1;
    $columnsTable.cell(1,2).range.text = "Data Type";
    $columnsTable.cell(1,3).range.Bold=1;
    $columnsTable.cell(1,3).range.text = "Max Length (Bytes)";
    $columnsTable.cell(1,4).range.Bold=1;
    $columnsTable.cell(1,4).range.text = "Allow Nulls";
    $columnsTable.cell(1,5).range.Bold=1;
    $columnsTable.cell(1,5).range.text = "Identity";
    $columnsTable.cell(1,6).range.Bold=1;
    $columnsTable.cell(1,6).range.text = "Default";
    $columnsTable.cell(1,7).range.Bold=1;
    $columnsTable.cell(1,7).range.text = "Descripcion";
    
    $i=2;

    foreach($column in $columns) 
    { 
        $description = getDescriptionExtendedProperty $column; 

        $columnsTable.Style = "List Table 3 - Accent 3";
        $columnsTable.Range.Font.Size = 10;
        $columnsTable.Range.Font.Name = "Times New Roman" ;

        $columnsTable.cell($i,1).range.Bold = 0;
        $columnsTable.cell($i,1).range.text = $column.Name;
        $columnsTable.cell($i,2).range.Bold = 0;
        $columnsTable.cell($i,2).range.text = $column.DataType.Name;
        $columnsTable.cell($i,3).range.Bold = 0;
        $columnsTable.cell($i,3).range.text = $column.DataType.MaximumLength;
        $columnsTable.cell($i,4).range.Bold = 0;
        $columnsTable.cell($i,4).range.text = $column.Nullable;
        $columnsTable.cell($i,5).range.Bold = 0;
        $columnsTable.cell($i,5).range.text = $column.Identity;
        $columnsTable.cell($i,6).range.Bold = 0;
        $columnsTable.cell($i,6).range.text = $column.Default;
        $columnsTable.cell($i,7).range.Bold = 0;
        $columnsTable.cell($i,7).range.text = $description;
          
        $i++;
    } 

    $selection.EndOf(15);
    $selection.MoveDown();

}

function CreateFkTable 
{ 
    param ($table, $selection); 

    $rows = $table.ForeignKeys.Count;

    if ($rows -gt 0){
        $nl = [Environment]::NewLine

        $selection.Style = 'Subtitle with line';
        $text = "Foreign Keys";
        $selection.TypeText($text);
        $selection.TypeParagraph();
        $selection.Style = 'No Spacing';

        $fkTable = $selection.tables.add(
            $selection.Range,($rows + 1),2,
            [Microsoft.Office.Interop.word.WdDefaulttableBehavior]::wdword9tableBehavior,
            [Microsoft.Office.Interop.word.WdAutoFitBehavior]::wdAutoFitContent
        );

        $fkTable.Style = "List Table 3 - Accent 3";
        $fkTable.Range.Font.Size = 10;
        $fkTable.Range.Font.Name = "Times New Roman" ;

        ## Header
        $fkTable.cell(1,1).range.Bold=1
        $fkTable.cell(1,1).range.text = "Name";
        $fkTable.cell(1,2).range.Bold=1;
        $fkTable.cell(1,2).range.text = "Columns";

    
        $i=2;

        Foreach ($foreignKey in $table.ForeignKeys)
        {
            $fkColumns = "";
            $fkName = "";
            $referencedTable = "";
            $referencedTableSchema = "";

            $fkName = $foreignKey.Name; 
            $referencedTable = $foreignKey.ReferencedTable; 
            $referencedTableSchema = $foreignKey.ReferencedTableSchema; 

            Foreach ($column in $foreignKey.Columns)
            {
                if($fkColumns -ne ""){
                    $fkColumns += $nl;
                }
                        
                $fkColumns += $column.Name + " -> [" + $referencedTableSchema + "].[" + $referencedTable + "].[" + $column.ReferencedColumn + "]";   
                      
            }
            $fkTable.Style = "List Table 3 - Accent 3";
            $fkTable.Range.Font.Size = 10;
            $fkTable.Range.Font.Name = "Times New Roman" ;
            $fkTable.cell($i,1).range.Bold = 0;
            $fkTable.cell($i,1).range.text = $fkName;
            $fkTable.cell($i,2).range.Bold = 0;
            $fkTable.cell($i,2).range.text = $fkColumns;

            $i++;
        }
    }

    $selection.EndOf(15);
    $selection.MoveDown();

}

function CreatePkIndexTable 
{ 
    param ($table, $selection); 

    $nl = [Environment]::NewLine
    $rows = $table.Indexes.Count;

    if ($rows -gt 0){

        $selection.Style = 'Subtitle with line';
        $text = "Prinary Key / Indexes";
        $selection.TypeText($text);
        $selection.TypeParagraph();
        $selection.Style = 'No Spacing';

        $indexesTable = $selection.tables.add(
            $selection.Range,($rows + 1),6,
            [Microsoft.Office.Interop.word.WdDefaulttableBehavior]::wdword9tableBehavior,
            [Microsoft.Office.Interop.word.WdAutoFitBehavior]::wdAutoFitContent
        );

        $indexesTable.Style = "List Table 3 - Accent 3";
        $indexesTable.Range.Font.Size = 10;
        $indexesTable.Range.Font.Name = "Times New Roman" ;

        ## Header
        $indexesTable.cell(1,1).range.Bold=1
        $indexesTable.cell(1,1).range.text = "Key";
        $indexesTable.cell(1,2).range.Bold=1
        $indexesTable.cell(1,2).range.text = "Name";
        $indexesTable.cell(1,3).range.Bold=1;
        $indexesTable.cell(1,3).range.text = "Key Columns";
        $indexesTable.cell(1,4).range.Bold=1;
        $indexesTable.cell(1,4).range.text = "Unique";
        $indexesTable.cell(1,5).range.Bold=1;
        $indexesTable.cell(1,5).range.text = "Is Clustered";
        $indexesTable.cell(1,6).range.Bold=1;
        $indexesTable.cell(1,6).range.text = "Fill Factor";

        $i=2;

        #Foreach($index in $table.Indexes | Where-Object { $_.IndexKeyType -ne "DriPrimaryKey" }){
        Foreach($index in $table.Indexes){
            $indexesTable.Style = "List Table 3 - Accent 3";
            $indexesTable.Range.Font.Size = 10;
            $indexesTable.Range.Font.Name = "Times New Roman" ;

            If($index.IndexKeyType -eq "DriPrimaryKey"){
                $indexesTable.cell($i,1).range.Bold = 0;
                $indexesTable.cell($i,1).range.text = "Pk";
            }
            else {
                $indexesTable.cell($i,1).range.Bold = 0;
                $indexesTable.cell($i,1).range.text = "";
            }
            $indexesTable.cell($i,2).range.Bold = 0;
            $indexesTable.cell($i,2).range.text = $index.Name;
            $indexesTable.cell($i,3).range.Bold = 0;
            $indexesTable.cell($i,3).range.text = ($index.IndexedColumns | select -ExpandProperty Name) -join " ,";
            $indexesTable.cell($i,4).range.Bold = 0;
            $indexesTable.cell($i,4).range.text = $index.IsUnique;
            $indexesTable.cell($i,5).range.Bold = 0;
            $indexesTable.cell($i,5).range.text = $index.IsClustered;
            $indexesTable.cell($i,6).range.Bold = 0;
            $indexesTable.cell($i,6).range.text = $index.FillFactor;

            $i++;
        }
    }

    $selection.EndOf(15);
    $selection.MoveDown();

}

#----------------------------------------------------------------------------------------------------------------------
#Begin Process

clear;

# Load required assemblies
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO")| Out-Null;
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMOExtended")| Out-Null;

$DatabaseName = "App2_V7_1_0";

$server = New-Object Microsoft.SqlServer.Management.Smo.Server $env:COMPUTERNAME;

$tablesNotToProcess = "Appointments_E11", "BI_Licensee", "BI_Licensee_FirstLoad", "BPE_IMPORT", "CCANY_Import", "Client_question_responses_E", "dtproperties",
    "DocumentationTables", "DocumentationColumns", "DWMaxID", "PNC_Import", "PNC_UsersLogin", "PNCDeletedRes", "PResSchedule", 
    "quest_data_import", "SHPSDataExport", "sysdiagrams", "TZOffSet", "TZOffSet", "Truliant_ResourceCharacteristic", "TraceData",
    "TraceDataEvents", "YPAppts", "ApptEngineTest", "ApptQueueGroup", "DailyResourceAvailability", "lock_moniter", "LocationImport", 
    "PartnerBusinessTemplate", "RefreshInterval", "SingleSignInTest", "SOAPService", "SOAPServiceType", "UserQR", "BlockedQueries", "Fkeys",
    "HDLDataExport"

$tablesToProcess = "LicenseeTimeZone", "LicenseeToReport", "LicenseeTransformers", "UserGroupExternalMapping"

$tables = $server.Databases[$DatabaseName].tables | Where { $tablesToProcess -contains $_.Name };
#$tables = $server.Databases[$DatabaseName].tables | Where { $tablesNotToProcess -notcontains $_.Name };
#$tables = $server.Databases[$DatabaseName].tables.Item("DependencyRules", "dbo");

$fileName = "Testdocumentation";
$savePath="C:\Temp\$fileName.docx";

$word = New-Object -ComObject Word.Application;
$word.Visible = $False;
$document = $word.Documents.Add("C:\Users\leandro.gomez\documents\Custom Office Templates\TimeTradedocument.dotx")         

$selection = $word.Selection;

$selection.Style = 'Heading 1';
$selection.TypeText("Data tables");
$selection.TypeParagraph();

Foreach ($table in $tables)
{
    $description = GetDescriptionExtendedProperty $table;

    $selection.Style = 'Heading 2';
    $selection.TypeText($table.Name);
    $selection.TypeParagraph();
    $selection.TypeParagraph();
    $selection.Style = 'No Spacing';
    $text = $description
    $selection.TypeText($text);
    $selection.TypeParagraph();    

    CreateColumnsTable $table $selection;
       
    $selection.Style = 'No Spacing';
    $selection.TypeParagraph();    

    CreatePkIndexTable $table $selection;

    $selection.Style = 'No Spacing';
    $selection.TypeParagraph();    
  
    CreateFkTable $table $selection;

    $selection.Style = 'No Spacing';

}


$document.SaveAs([ref]$savePath,[ref]$SaveFormat::wdFormatdocument);
$document.Close();
$word.Quit();
;
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect();
[gc]::WaitForPendingFinalizers();
Remove-Variable word;
