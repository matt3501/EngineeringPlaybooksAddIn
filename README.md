# Engineering Playbooks Addin README

## How to use
1. Install the plugin
1. Open Visio
1. Create a new Visio Blank Drawing
  1. ![image](https://github.jci.com/storage/user/636/files/6bcc8e80-d49e-11e9-88f2-1e8895f78cb3)
1. Choose 'US Units'
  1. ![image](https://github.jci.com/storage/user/636/files/6e2ee880-d49e-11e9-8cdd-26cdc5a41b5a)
1. In the Ribbon, go to Add-ins and select 'Draw Playbook'
  1. ![image](https://github.jci.com/storage/user/636/files/6ff8ac00-d49e-11e9-99be-2d76eb895972)
1. Choose a playbook JSON file, example located  [here](https://github.jci.com/cplankm/EngineeringPlaybooksAddIn/blob/master/EngineeringPlaybooksAddIn/samples/knowledge_WorkflowsBusinessAnalyst.json)
  1. ![image](https://github.jci.com/storage/user/636/files/71c26f80-d49e-11e9-99ca-31b098e0542c)

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
