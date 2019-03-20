#include-once
#include <File.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GuiComboBox.au3>
#include <EditConstants.au3>
#include <GUIListBox.au3>
#include <GuiEdit.au3>
#include "CSV.au3"








;$rr = ShellExecuteWait('diff.exe', '"new 5.txt" "new 6.txt"', @ScriptDir, "", @SW_HIDE)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $rr = ' & $rr & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit


 ;Local $iPID = Run(@ScriptDir & '\diff.exe "new 5.txt" "new 6.txt"', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ; If you want to search with files that contains unicode characters, then use the /U commandline parameter.

    ; Wait until the process has closed using the PID returned by Run.
  ;  ProcessWaitClose($iPID)

    ; Read the Stdout stream of the PID returned by Run. This can also be done in a while loop. Look at the example for StderrRead.
   ; Local $sOutput = StdoutRead($iPID)
	;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sOutput = ' & $sOutput & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit



Global $run_ids
Global $html
Global $app_name = "Datapool Manager"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"

Global $main_gui = GUICreate($app_name, 680, 600)

GUICtrlCreateGroup("Project", 10, 10, 420, 50)
Global $project_select_button = GUICtrlCreateButton("Select", 30, 30, 80, 20)
Global $project_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "projectpath", "R:\"), 120, 30, 300, 20, $ES_READONLY)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Datapool Names", 10, 70, 320, 410)
Global $datapools_list = GUICtrlCreateList("", 20, 90, 300, 300, BitOR($LBS_SORT, $WS_BORDER, $WS_VSCROLL, $LBS_EXTENDEDSEL))
Global $sort_data_in_selected_datapools_button = GUICtrlCreateButton("Sort Data in Selected Datapools", 30, 400, 200, 20)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $sort_data_in_all_datapools_button = GUICtrlCreateButton("Sort Data in All Datapools", 30, 420, 200, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Test Case Names", 340, 70, 320, 410)
Global $test_cases_list = GUICtrlCreateList("", 350, 90, 300, 300)
GUICtrlCreateLabel("Copy Selected As ", 350, 405, 120, 20)
Global $new_test_case_edit = GUICtrlCreateInput("", 480, 400, 170, 20)
Global $copy_test_case_data_button = GUICtrlCreateButton("Copy && Sort Data in All Datapools", 450, 420, 200, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)


Global $status_input = GUICtrlCreateInput("Enter the ""Epic Key"" and click ""Start""", 10, 600 - 25, 320, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $status_progress = GUICtrlCreateProgress(340, 600 - 25, 320, 20)

GUISetState(@SW_SHOW, $main_gui)


$datapool_filename = _FileListToArrayRec("R:\RMS\local RMS Drop 4 Se project\Data\Pools", "*.csv", 1, 1, 1, 1)
_ArrayDelete($datapool_filename, 0)
$str = _ArrayToString($datapool_filename)
GUICtrlSetData($datapools_list, $str)


_CSV_Initialise()



GUICtrlSetData($status_input, "")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

; Loop until the user exits.
While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_CLOSE

;			IniWrite($ini_filename, "main", "testrailusername", GUICtrlRead($testrail_username_input))
;			IniWrite($ini_filename, "main", "testrailproject", GUICtrlRead($testrail_project_combo))
;			IniWrite($ini_filename, "main", "testrailplan", GUICtrlRead($testrail_plan_combo))
;			IniWrite($ini_filename, "main", "jirausername", GUICtrlRead($jira_username_input))
;			IniWrite($ini_filename, "main", "epickeys", GUICtrlRead($epic_key_input))

;			$testrail_encrypted_password = _Crypt_EncryptData(GUICtrlRead($testrail_password_input), "applesauce", $CALG_AES_256)
;			IniWrite($ini_filename, "main", "testrailpassword", $testrail_encrypted_password)

;			$jira_encrypted_password = _Crypt_EncryptData(GUICtrlRead($jira_password_input), "applesauce", $CALG_AES_256)
;			IniWrite($ini_filename, "main", "jirapassword", $jira_encrypted_password)


			ExitLoop

		Case $project_select_button

			Local $project_path = FileSelectFolder("Select the project folder", "R:\", 0, "", $main_gui)

			if StringLen($project_path) > 0 Then

				GUICtrlSetData($project_input, $project_path)
				IniWrite($ini_filename, "main", "projectpath", $project_path)

			EndIf


		Case $sort_data_in_all_datapools_button

			$ans = MsgBox(1 + 32 + 8192 + 262144, $app_name, "Continuing will update all datapools.  Continue?")

			if $ans = 1 Then

				; backup the pools folder

				if FileExists(@ScriptDir & "\backup") = True Then

					GUICtrlSetData($status_input, "Removing old backup folder ...")
					DirRemove(@ScriptDir & "\backup", 1)
				EndIf

				GUICtrlSetData($status_input, "Creating new backup folder ...")
				DirCreate(@ScriptDir & "\backup")
				GUICtrlSetData($status_input, "Copying Pools folder to backup folder ...")
				Local $result = DirCopy(GUICtrlRead($project_input) & "\Data\Pools", @ScriptDir & "\backup\Pools", 1)

				if $result = 0 Then

					GUICtrlSetData($status_input, "Failed to backup the existing Pools folder.  Aborted.")
				Else

					GUICtrlSetData($status_input, "")
				EndIf

				Local $datapool_updated[0]

				for $i = 0 to (_GUICtrlListBox_GetCount($datapools_list) - 1)

					Local $relative_datapool_path = _GUICtrlListBox_GetText($datapools_list, $i)
					GUICtrlSetData($status_input, UBound($datapool_updated) & " Datapools updated. Checking " & $relative_datapool_path)

					if StringInStr($relative_datapool_path, " - Parameters.csv", 1) = 0 And StringInStr($relative_datapool_path, "Release Parameters.csv", 1) = 0 Then

						Local $full_datapool_path = GUICtrlRead($project_input) & "\Data\Pools\" & $relative_datapool_path
						Local $csv_handle = _CSV_Open($full_datapool_path)
						Local $csv_result = ""

						; if records with the source 'Assigned to' are found

						Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
						_PathSplit($full_datapool_path, $sDrive, $sDir, $sFileName, $sExtension)

						Local $temp_csv_file = @TempDir & "\" & $sFileName & $sExtension
						_CSV_SaveAs($csv_handle, $temp_csv_file, "select * from csv order by `Assigned to`, `Comment 1`;")

						if FileGetSize($temp_csv_file) = 0 Then

							ConsoleWrite("Datapool " & $temp_csv_file & " is zero size - skipping" & @CRLF)
						Else

							Local $diff_result = ShellExecuteWait('diff.exe', '"' & $temp_csv_file & '" "' & $full_datapool_path & '"', @ScriptDir, "", @SW_HIDE)

							if $diff_result = 0 Then

	;								ConsoleWrite("Datapool " & $datapool_file[$i] & " is the same" & @CRLF)
							Else

								ConsoleWrite("Datapool " & $relative_datapool_path & " is different - updating ..." & @CRLF)
								Local $result = FileCopy($temp_csv_file, $full_datapool_path, 1)

								if $result = 0 Then

									ConsoleWrite("Updating " & $relative_datapool_path & " failed." & @CRLF)
								Else

									if StringInStr($relative_datapool_path, "Merged\", 1) > 0 Then

										_ArrayAdd($datapool_updated, $relative_datapool_path)
									EndIf
								EndIf
							EndIf
						EndIf
					EndIf

					GUICtrlSetData($status_progress, $i / (_GUICtrlListBox_GetCount($datapools_list) - 1) * 100)
				Next

				GUICtrlSetData($status_input, "")
				GUICtrlSetData($status_progress, 0)

				if UBound($datapool_updated) > 0 Then

					Local $str = "The following Merged datapools were updated:" & @CRLF & @CRLF
					$str = $str & _ArrayToString($datapool_updated, @CRLF) & @CRLF & @CRLF
					$str = $str & "Right click these in VS and ""Check out for Edit..."""

					MsgBox(64 + 262144, $app_name, $str, 0, $main_gui)
				EndIf

			EndIf



		Case $copy_test_case_data_button

			Local $source_test_case_name = GUICtrlRead($test_cases_list)
			Local $target_test_case_name = GUICtrlRead($new_test_case_edit)

			$ans = MsgBox(1 + 32 + 8192 + 262144, $app_name, "Continuing will update all datapools.  Continue?")

			if $ans = 1 Then

				; backup the pools folder

				if FileExists(@ScriptDir & "\backup") = True Then

					GUICtrlSetData($status_input, "Removing old backup folder ...")
					DirRemove(@ScriptDir & "\backup", 1)
				EndIf

				GUICtrlSetData($status_input, "Creating new backup folder ...")
				DirCreate(@ScriptDir & "\backup")
				GUICtrlSetData($status_input, "Copying Pools folder to backup folder ...")
				Local $result = DirCopy(GUICtrlRead($project_input) & "\Data\Pools", @ScriptDir & "\backup\Pools", 1)

				if $result = 0 Then

					GUICtrlSetData($status_input, "Failed to backup the existing Pools folder.  Aborted.")
				Else

					GUICtrlSetData($status_input, "")
				EndIf

				Local $datapool_updated[0]
				Local $num_datapools_updated = 0

				for $i = 0 to (_GUICtrlListBox_GetCount($datapools_list) - 1)

					Local $relative_datapool_path = _GUICtrlListBox_GetText($datapools_list, $i)
					GUICtrlSetData($status_input, $num_datapools_updated & " Datapools updated. Checking " & $relative_datapool_path)

					if StringInStr($relative_datapool_path, " - Parameters.csv", 1) = 0 And StringInStr($relative_datapool_path, "Release Parameters.csv", 1) = 0 Then

						Local $full_datapool_path = GUICtrlRead($project_input) & "\Data\Pools\" & $relative_datapool_path
						Local $csv_handle = _CSV_Open($full_datapool_path)
						Local $csv_result = ""

						; if records with the source 'Assigned to' are found

						Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
						_PathSplit($full_datapool_path, $sDrive, $sDir, $sFileName, $sExtension)

						_CSV_Exec($csv_handle, "delete from csv where `Assigned to` = '" & $target_test_case_name & "';")
						Local $largest_rowid = _CSV_GetLargestRowID($csv_handle)
						_CSV_Exec($csv_handle, "insert into csv select * from csv where `Assigned to` = '" & $source_test_case_name & "';")
						_CSV_Exec($csv_handle, "update csv set `Assigned to` = '" & $target_test_case_name & "' where rowid > " & $largest_rowid & ";")

						Local $temp_csv_file = @TempDir & "\" & $sFileName & $sExtension
						_CSV_SaveAs($csv_handle, $temp_csv_file, "select * from csv order by `Assigned to`, `Comment 1`;")

						if FileGetSize($temp_csv_file) = 0 Then

							ConsoleWrite("Datapool " & $temp_csv_file & " is zero size - skipping" & @CRLF)
						Else

							Local $diff_result = ShellExecuteWait('diff.exe', '"' & $temp_csv_file & '" "' & $full_datapool_path & '"', @ScriptDir, "", @SW_HIDE)

							if $diff_result = 0 Then

	;								ConsoleWrite("Datapool " & $datapool_file[$i] & " is the same" & @CRLF)
							Else

								ConsoleWrite("Datapool " & $relative_datapool_path & " is different - updating ..." & @CRLF)
								Local $result = FileCopy($temp_csv_file, $full_datapool_path, 1)
								$num_datapools_updated = $num_datapools_updated + 1

								if $result = 0 Then

									ConsoleWrite("Updating " & $relative_datapool_path & " failed." & @CRLF)
								Else

									if StringInStr($relative_datapool_path, "Merged\", 1) > 0 Then

										_ArrayAdd($datapool_updated, $relative_datapool_path)
									EndIf
								EndIf
							EndIf
						EndIf
					EndIf

					GUICtrlSetData($status_progress, $i / (_GUICtrlListBox_GetCount($datapools_list) - 1) * 100)
				Next

				GUICtrlSetData($status_input, "")
				GUICtrlSetData($status_progress, 0)

				if UBound($datapool_updated) > 0 Then

					Local $str = "The following Merged datapools were updated:" & @CRLF & @CRLF
					$str = $str & _ArrayToString($datapool_updated, @CRLF) & @CRLF & @CRLF
					$str = $str & "Right click these in VS and ""Check out for Edit..."""

					MsgBox(64 + 262144, $app_name, $str, 0, $main_gui)
				EndIf

			EndIf







#cs

		Case $copy_test_case_data_button

			Local $source_test_case_name = GUICtrlRead($test_cases_list)
			Local $target_test_case_name = GUICtrlRead($new_test_case_edit)

			$ans = MsgBox(1 + 32 + 256 + 8192 + 262144, $app_name, "Continuing will update all datapools.  Continue?")

			if $ans = 1 Then

				; backup the pools folder

				if FileExists(@ScriptDir & "\backup") = True Then

					GUICtrlSetData($status_input, "Removing old backup folder ...")
					DirRemove(@ScriptDir & "\backup", 1)
				EndIf

				GUICtrlSetData($status_input, "Creating new backup folder ...")
				DirCreate(@ScriptDir & "\backup")
				GUICtrlSetData($status_input, "Copying Pools folder to backup folder ...")
				Local $result = DirCopy("R:\RMS\local RMS Drop 4 Se project\Data\Pools", @ScriptDir & "\backup\Pools", 1)

				if $result = 0 Then

					GUICtrlSetData($status_input, "Failed to backup the existing Pools folder.  Aborted.")
				Else

					GUICtrlSetData($status_input, "")
				EndIf

				Local $datapool_file = _FileListToArrayRec("R:\RMS\local RMS Drop 4 Se project\Data\Pools", "*.csv", 1, 1, 0, 2)

				for $i = 1 to $datapool_file[0]

;					ConsoleWrite("Datapool " & $i & " of " & $datapool_file[0] & " - " & $datapool_file[$i] & @CRLF)
					GUICtrlSetData($status_input, "Datapool " & $i & " of " & $datapool_file[0] & " - " & $datapool_file[$i])

					if StringInStr($datapool_file[$i], " - Parameters.csv", 1) = 0 And StringInStr($datapool_file[$i], "Release Parameters.csv", 1) = 0 Then

						Local $csv_handle = _CSV_Open($datapool_file[$i])
						Local $csv_result = ""

						if StringLen($target_test_case_name) > 0 Then

							$csv_result = _CSV_GetRecordArray($csv_handle, "select * from csv where `Assigned to` = '" & $source_test_case_name & "';", False)
						Else

							$csv_result = _CSV_GetRecordArray($csv_handle, "select * from csv;", False)
						EndIf

							ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : UBound($csv_result) = ' & UBound($csv_result) & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

						; if records with the source 'Assigned to' are found

						if UBound($csv_result) = 0 Then

							ConsoleWrite("Datapool " & $datapool_file[$i] & " can't be read - maybe too many columns" & @CRLF)
						Else

							Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
							_PathSplit($datapool_file[$i], $sDrive, $sDir, $sFileName, $sExtension)
	;						ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $datapool_file[$i] = ' & $datapool_file[$i] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

							if StringLen($target_test_case_name) > 0 Then

								Local $number_of_records = _CSV_GetRecordCount($csv_handle)
								_CSV_Exec($csv_handle, "insert into csv select * from csv where `Assigned to` = '" & $source_test_case_name & "';")
								_CSV_Exec($csv_handle, "update csv set `Assigned to` = '" & $target_test_case_name & "' where rowid > " & $number_of_records & ";")
							EndIf

							_CSV_SaveAs($csv_handle, @TempDir & "\" & $sFileName & $sExtension, "select * from csv order by `Assigned to`, `Comment 1`;")
							Local $diff_result = ShellExecuteWait('diff.exe', '"' & @TempDir & "\" & $sFileName & $sExtension & '" "' & $datapool_file[$i] & '"', @ScriptDir, "", @SW_HIDE)

							if $diff_result = 0 Then

;								ConsoleWrite("Datapool " & $datapool_file[$i] & " is the same" & @CRLF)
							Else

								ConsoleWrite("Datapool " & $datapool_file[$i] & " is different - updating ..." & @CRLF)
								_CSV_SaveAs($csv_handle, $datapool_file[$i], "select * from csv order by `Assigned to`, `Comment 1`;")
							EndIf

						EndIf
					EndIf
				Next

				GUICtrlSetData($status_input, "")
			EndIf
#ce

	EndSwitch

WEnd

GUIDelete($main_gui)
_CSV_Cleanup()




Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom, $iIDFrom, $iCode
    $hWndFrom = $lParam
    $iIDFrom = BitAND($wParam, 0xFFFF) ; Low Word
    $iCode = BitShift($wParam, 16) ; Hi Word


	Switch $hWndFrom

        Case GUICtrlGetHandle($datapools_list)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box


					Local $csv_handle = _CSV_Open("R:\RMS\local RMS Drop 4 Se project\Data\Pools\" & GUICtrlRead($datapools_list))
					Local $csv_result = _CSV_GetRecordArray($csv_handle, "select distinct `Assigned to` from csv order by `Assigned to`;", False)
					$str = _ArrayToString($csv_result, "|", -1, -1, "|")
					GUICtrlSetData($test_cases_list, "")
					GUICtrlSetData($test_cases_list, $str)


            EndSwitch

        Case GUICtrlGetHandle($test_cases_list)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

					GUICtrlSetData($new_test_case_edit, GUICtrlRead($test_cases_list) & " copy")
					ControlFocus($main_gui, "", $new_test_case_edit)
					_GUICtrlEdit_SetSel($new_test_case_edit, StringLen(GUICtrlRead($test_cases_list)), StringLen(GUICtrlRead($test_cases_list)) + StringLen(" copy"))

            EndSwitch
	EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND
