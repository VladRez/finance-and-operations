#Combine .csv files within a directory into a single .csv 


$INPUT_DIR = "C:\PATH\TO\INPUT\DIRECTORY\"
$OUTPUT_DIR = "C:\PATH\TO\OUTPUT\DIRECTORY\"
$OUTPUT_FILE = "OUTPUTNAME.CSV"
$FULLPATH = ("{0}{1}"-f $OUTPUT_DIR,$OUTPUT_FILE)

$getFirstLine = $true
Write-Output $FULLPATH


get-childItem $INPUT_DIR -File | foreach {
    $filePath = ("{0}{1}"-f $INPUT_DIR,$_)

    $lines =  $lines = Get-Content $filePath  
    $linesToWrite = switch($getFirstLine) {
			#skip first line of subsequent file
           $true  {$lines}
           $false {$lines | Select -Skip 1}

    }

    $getFirstLine = $false
    Add-Content $FULLPATH $linesToWrite
    } 
