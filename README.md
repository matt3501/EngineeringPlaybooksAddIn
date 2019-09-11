# Engineering Playbooks Addin README

## COMMAND PROMPT INSTALL 
1. Assuming the EngineeringPlaybooksAddin.zip is in your downloads and you have 7zip installed, run the following commands in a (NON ADMIN) command prompt.

  ```batch
	REM start Commands
		cd %HOMEPATH%\Downloads\
		"C:\Program Files\7-Zip\7z.exe" e EngineeringPlaybooksAddin.zip *.* -o%HOMEPATH%\VisioAddins\EngineeringPlaybooksAddin\ -r
		explorer %HOMEPATH%\VisioAddins\EngineeringPlaybooksAddin\EngineeringPlaybooksAddIn.vsto
	REM end commands
  ```

## MANUAL INSTALLATION STEPS
1. Unzip this folder to your user directory 
	1. (E.G.) EngineeringPlaybooksAddIn.zip > C:\Users\<your user\VisioAddins\EngineeringPlaybooksAddin\
2. Double click EngineeringPlaybooksAddin\EngineeringPlaybooksAddIn.vsto
