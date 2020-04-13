#############################################################################################################################################
### Footer Script
#############################################################################################################################################

if ($Error.Count) {
    Write-Host "There were errors while executing the script." -ForegroundColor Red
    Write-Host "Please review the console output." -ForegroundColor Red
}
else {
    Write-Host "Script execution completed sucessfully." -ForegroundColor Green
    Write-Host "Please regenerate the report with the latest config exports to confirm." -ForegroundColor Green
    Write-Host "Once confirmed, it is recommended that you run Full Synchronization run profile all connectors." -ForegroundColor Green
}
