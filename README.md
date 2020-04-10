# Engineering Playbooks Addin README

## How to use
1. Install the plugin
1. Open Visio
1. Create a new Visio Blank Drawing
1. Choose 'US Units'
1. In the Ribbon, go to Add-ins and select 'Draw Playbook'
1. Choose a playbook JSON file, example located  [here](https://github.com/matt3501/EngineeringPlaybooksAddIn/blob/master/EngineeringPlaybooksAddIn/samples/knowledge_WorkflowsBusinessAnalyst.json)

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
1. Double click EngineeringPlaybooksAddin\EngineeringPlaybooksAddIn.vsto
