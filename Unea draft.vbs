
'Stats Gathering===========================================================
name_of_script ="stat, unea"
start_timer = timer
Stats_counter = 1
Stats_manualtime = 50
Stats_denomination = "C"
'End of Stats Block==========================================================



'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


'Check Maxis status to see if timed out Maxis case number
Call check_for_Maxis(true)
call Maxis_case_number_finder(case_number)


BeginDialog Unea_Panel, 0, 0, 166, 175, "Unea Panel"
  EditBox 70, 5, 50, 15, Client_name
  DropListBox 85, 25, 60, 45, "Select One"+chr(9)+"Direct Spousal"+chr(9)+"Direct Child Support", Income_Type
  DropListBox 70, 40, 60, 45, "Select One"+chr(9)+"Copy of Checks"+chr(9)+"Award Letters"+chr(9)+"System Initiated Verification"+chr(9)+"Coltrl Stmt"+chr(9)+"Pend Out State Verification"+chr(9)+"Other Document"+chr(9)+"Worker Initiated"+chr(9)+"RI Stubs"+chr(9)+"No Ver Prvd", List5
  EditBox 80, 60, 50, 15, claim_number
  EditBox 75, 80, 50, 15, Inc_Start_date
  EditBox 75, 100, 50, 15, Edit7
  EditBox 70, 120, 50, 15, Unea_Frequency
  ButtonGroup ButtonPress
    OkButton 5, 145, 50, 15
    CancelButton 65, 145, 50, 15
  Text 10, 85, 45, 10, "Inc Start date"
  Text 10, 25, 45, 10, "Income Type"
  Text 10, 105, 40, 10, "Prospective "
  Text 10, 10, 40, 10, "Client name"
  Text 10, 65, 50, 10, "Claim Number"
  Text 10, 125, 50, 10, "Frequency"
  Text 10, 45, 25, 10, "Verified"
EndDialog


EMConnect ""


'Read case number
call Maxis_case_number_finder(case_number)

'Navigate from self window to stat unea panel
Call navigate_to_Maxis_screen("stat", "unea")

'Create new unea panel
EmWriteScreen "nn", 20, 79

'Transmit to take new maxis screen into edit mode
Transmit

'Enter information on unea panel insert msgbox to fill info. Could be expanded to include all types of une, HH membs.
Dialog Unea_Panel
If buttonPressed = 0 then StopScript

' err msg added so mandatory info can't be blank
(err_msg = "")


'converting variable for Maxis Income Type
If Income_Type = "Direct Spousal" then Income_Type = "29" 
If Income_Type = "Direct Child Support" then Income_Type = "08"

'writing the information Income Type
EmWriteScreen Income_Type , 5,37

'Err msg added so mandatory info an't be left blank
(err_msg = "")


'converting variable for Maxis verification
If List5 = "Copy of Checks" then List5 = "1"
If List5 = "Award Letters" then List5 = "2"
If List5 = "System Initiated Verification" then List5 = "3"
If List5 = "Coltrl Stmt" then List5 = "4"
If List5 = "Pend Out State Verification" then List5 = "5"
If List5 = "Other Document" then List5 = "6"
If List5 = "Worker Initiated" then List5 = "7"
If List5 = "RI Stubs" then List5 = "8"
If List5 = "No Ver Prvd" then List5 = "N"



'writing variable maxis verification
EmWriteScreen List5, 5, 65


' no err msg written in as it could be left blank


'convert variable claim number for Maxis
EmWriteScreen claim_number, 6,37

'Write variable for 
EmWriteScreen Inc_Start_date



'Enter information on SNAP pop-up window insert a msgbox to fill info



'Transmit to exit the SNAP pop-up
Transmit

'Transmit
Transmit


Call check_for_Maxis (True)

'Enter case note with msgbox information 
Call write_bullet_and_variable_in_CASE_NOTE ("Bullet", variable)
Call write_variable_in_case_note(worker_signature)

'Stop script

Script_end_procedure ("")


