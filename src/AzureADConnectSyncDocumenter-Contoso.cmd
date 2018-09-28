AzureADConnectSyncDocumenterCmd.exe "Contoso\Pilot" "Contoso\Production"

ECHO OFF

IF %ERRORLEVEL% EQU 9009 (
	ECHO ****************************************************************************************************
	ECHO It seems you may have downloaded the source code instead of a release package. Please download the latest release from https://github.com/Microsoft/AADConnectConfigDocumenter/releases...
	ECHO ****************************************************************************************************
) ELSE (
	ECHO ****************************************************************************************************
	ECHO Execution complete. Please check any errors or warnings in the AADConnectSyncDocumenter-Error.log...
	ECHO ****************************************************************************************************
)

@pause