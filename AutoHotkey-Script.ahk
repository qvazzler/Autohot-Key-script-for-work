#SingleInstance force
#InstallKeybdHook
SetTitleMatchMode 2
SetTitleMatchMode Fast

; >>>>>>>>>>>>>>>>>>>>>>>>
; >>> Editing function >>>
; >>>>>>>>>>>>>>>>>>>>>>>>
#IfWinActive AutoHotkey.ahk * SciTE4AutoHotkey
^s::
WinActivate AutoHotkey.ahk * SciTE4AutoHotkey
SendInput {Control Down}s{Control Up}
WinWait, AutoHotkey.ahk - SciTE4AutoHotkey, , 3
if ErrorLevel
{
  MsgBox, File did not get saved in 3 seconds - Not reloading the script now
}
else
{
  run "C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe" C:\Users\olesena\Documents\AutoHotkey.ahk
}
return
#IfWinActive
; <<<<<<<<<<<<<<<<<<<<<<<<
; <<< Editing function <<<
; <<<<<<<<<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>
; >>> Functions >>>
; >>>>>>>>>>>>>>>>>
CaseNumber(){
  DetectHiddenText, off                                       ; Make it so we can spot hidden text
  WinGetText, SRDWinText, SR-Dash                             ; Grab all windows test from SR-Dash
  CaseNPos := InStr(SRDWinText, "Case Notes: ", false, 1, 1)  ; Find the position of "Case Notes: " in the grabbed SR-Dash text
  CaseNumber := SubStr(SRDWinText,CaseNPos+12,10)             ; Get only the case number out of the grabbed text, using the position
  return %CaseNumber%                                         ; Send back the case number
}

URLDownloadToVar( URL ) {     
 WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")     ; Create a new HTTP Request object
 WebRequest.Open("GET", URL)                                  ; Open the new object and add the functions URL
 WebRequest.Send()                                            ; Send the request with the customer URL
 return WebRequest.ResponseText                               ; Send back the result of the HTTP Request based on the URL in the function call
}

NextFollowUpDateTime( Amountofdays := 1 ) {
  if Amountofdays between 1 and 6     
  {
    while Amountofdays > 0 
    {
      EvaluationDay += 1, days
      FormatTime, Weekday, %EvaluationDay%, WDay
      if Weekday between 2 and 6 
        Amountofdays--
    }
    FormatTime, NFDT_Parsed, %EvaluationDay% L1033, dddd dd-MM-yy
    NFDT_Final = Next Follow-up Date/Time: %NFDT_Parsed% 13:00 GMT
    return NFDT_Final
  } 
  else if Amountofdays = 0
  {
    FormatTime, NFDT_Parsed, %EvaluationDay% L1033, dddd dd-MM-yyyy
    NFDT_Final = Next Follow-up Date/Time: %NFDT_Parsed% 13:00 GMT
    return NFDT_Final
  }
  else 
    MsgBox , , Error in function call, NFDT Function used incorrectly`nIt can only accept integer values between 0 and 6`n`nPlease reprogram the function call.
}

GetHPSupportPassword() {
  WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
  WebRequest.Open("GET", "http://hpcrlib1.us.rdlabs.hpecorp.net/customercare/Technicalsupport/diagnostics/Ltt/pw_gen/LTT2.7-2day_password.txt")
  WebRequest.Send()
  BulkHTMLDump := WebRequest.ResponseText
  GeneratedPassword := SubStr(BulkHTMLDump,InStr(BulkHTMLDump, "FAYEFB", false, 1, 1)-8,14)
  return GeneratedPassword
}
 
PrintHPSupportPassword( TextInFront := "" ) {   ; Initialize function with Received text
 HPSupportPassword := GetHPSupportPassword()    ; Get HP Support password for today
 SetKeyDelay, 50                                ; Set the key to key delay to 25 milliseconds
 Send %TextInFront%%HPSupportPassword%          ; Write out the HP Support Password (Optionally with text in front
 SetKeyDelay, 0                                 ; Remove the key to key delay
 return                                         ; End function
}

SearchMarkedText( SearchURL := "Not entered" ) {
  ;Check if the variables have been filled in right
  IfInString, SearchURL, http
  { 
    if (StrLen(SearchURL) > 10)
    {
      ClipSaved := ClipboardAll     ; Save the current clipboard
      Clipboard =                   ; Start off empty to allow ClipWait to detect when the text has arrived.
      Send ^c                       ; Send control - C (Copy)
      ClipWait,5                    ; Wait for 5 seconds or until the clipboard contains text.
      if ErrorLevel{                ; If it fails
        MsgBox , , Function Error: 004, No data was made available in clipboard`nScript waited for 5 seconds got no data`nRemember to mark the text you need search.  
                                    ; ^^ Display messagebox
        Clipboard := ClipSaved      ; Restore saved clipboard
        return                      ; Stop Script
      }
      Run %SearchURL%%Clipboard%
      Clipboard := ClipSaved
      ClipSaved =
    return
    } else {
      ; Fail condition - Has http in it, but is not long enough to be a proper URL 
      MsgBox , , 'Function Error: 001', URL for the search function is too short`n`nPlease re-write the code executing this function
    }
  } else {
    ; Fail condition - does not contain the string "http"
    if (SearchURL = "Not entered")
    {
      ; Fail condition - There was no Search URL entered at all
      MsgBox , , Function Error: 003, Search URL has not been entered at all`n`nPlease re-write the code executing this funciton
    } else {
      ; Fail condition - Something was entered in the SearchURL variable, but it did not contain HTTP
      MsgBox , , Function Error: 002, Search URL is malformed`nContains: [%SearchURL%]`n`nPlease re-write the code executing this function  
    }
  }
}
; <<<<<<<<<<<<<<<<<
; <<< Functions <<<
; <<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> HP Support StoreOnce Shortcuts >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
::hpsupkey::
  PrintHPSupportPassword("hp ")
return

::chkfs::du -ch --max-depth=1

::clrbad::
  FormatTime, TimeStamp, , dd_MM_yyyy_HH_mm
  SendInput mv s.bad_integrity Cleared_bad_integrity_%TimeStamp%_Anders_Olesen
return
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<< HP Support StoreOnce Shortcuts <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

#n:: Run notepad.exe

#x::  ; Press Win+G to show the current settings of the window under the mouse.
  MouseGetPos,,, MouseWin
  WinGetTitle, WinTitle, ahk_id %MouseWin%
  WinGetClass, WinClass, ahk_id %MouseWin%
  ToolTip Title:`t`t`t%WinTitle%`nAHK Class:`t`t%WinClass%
  SetTimer, RemoveToolTip, 3000
return

RemoveToolTip:
 SetTimer, RemoveToolTip, Off
 ToolTip
return

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> Second Opinion helper start >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#IfWinActive ahk_class WindowsForms10.Window.8.app.0.218f99c
#2::
 ClipSaved := ClipboardAll                ; Save the clipboard as it is now
 Clipboard =                              ; Start off empty to allow ClipWait to detect when the text has arrived.
 CaseID := CaseNumber()                   ; Get the case number from the current case
 WinActivate NER HPSD L1 Nearline         ; Bring the Chat room to the forefront
 WinWait, NER HPSD L1 Nearline            ; Wait for it to be active
 SendInput %CaseID% - SO Needed -{Space}  ; Put in CaseID and standard text
 Clipboard := ClipSaved                   ; Restore saved clipboard
 ClipSaved =                              ; Empty savedclipboard variable
return

^2::
ClipSaved := ClipboardAll         ; Save the clipboard as it is now
Clipboard =                       ; Start off empty to allow ClipWait to detect when the text has arrived.
CaseID := CaseNumber()            ; Get the case number from the current case
WinActivate NER HPSD L1 Nearline  ; Bring the Chat room to the forefront
WinWait, NER HPSD L1 Nearline     ; Wait for it to be active
SendInput %CaseID% - Done         ; Put in CaseID and standard end text
Clipboard := ClipSaved            ; Restore saved clipboard
ClipSaved =                       ; Empty savedclipboard variable
return
#IfWinActive

#IfWinActive NER HPSD L1 Nearline
^2::
ClipSaved := ClipboardAll   ; Save the clipboard as it is now
Clipboard =                 ; Start off empty to allow ClipWait to detect when the text has arrived.
SendInput ^c                ; Press Control C (Copy)
ClipWait,5                  ; Wait 5 for Clipboard to fill up
if ErrorLevel               ; If it fails
{
  MsgBox, No data avaialble for the clipboard, remember to mark the text you need.  ; Display messagebox
  Clipboard := ClipSaved                                                            ; Restore saved clipboard
  return                                                                            ; Stop Script
}
else
{
  Send {Shift Down}{Tab 1}{Shift Up}^a^v{Space}-{Space}Checking...  ; Put in text to chat room
  SendInput {Enter}                                                 ; Send the text
  WinActivate ahk_class WindowsForms10.Window.8.app.0.218f99c       ; Activate the SR-Dash window
  WinWait, ahk_class WindowsForms10.Window.8.app.0.218f99c          ; Wait for the SR-Dash window to be active
  SendInput {Alt Down}a{Alt Up}                                     ; Alt + A
  return
}
#IfWinActive
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<< Second Opinion helper end <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> Search helpers start >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>
#g::  SearchMarkedText("https://www.google.co.uk/search?q=")
#^g:: SearchMarkedText("http://lmgtfy.com/?q=")
#c::  SearchMarkedText("http://lego-web-pro.houston.hp.com/ccat/caselookup.aspx?caseid=")
#s::  SearchMarkedText("http://lego-web-pro.houston.hp.com/tracking/dosearch.asp?SerialNumber=")
#z::  SearchMarkedText("http://h41302.www4.hp.com/km/saw/search.do?go=Go&query=")
#p::  SearchMarkedText("http://partsurfer.hpe.com/Search.aspx?searchText=")
#b::  SearchMarkedText("https://bugzilla.houston.hp.com:1181/bugzilla/buglist.cgi?bug_status=UNCONFIRMED&bug_status=NEW&bug_status=ASSIGNED&bug_status=REOPENED&bug_status=RESOLVED&bug_status=VERIFIED&bug_status=CLOSED&classification=StoreAll%20Classic&classification=StoreOnce&list_id=861040&query_format=advanced&short_desc_type=allwordssubstr&short_desc=")
; <<<<<<<<<<<<<<<<<<<<<<<<<
; <<< Search helper end <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<

^#c::
CaseID := CaseNumber()  ; Get case number
SendInput %CaseID%      ; Type case number
return

^#n::
ClipSaved := ClipboardAll         ; Save current clipboard
Clipboard =                       ; Start off empty to allow ClipWait to detect when the text has arrived.
Send ^c                           ; Press Control C (Copy)
ClipWait,1                        ; Wait for 5 seconds or until the clipboard contains text.
if ErrorLevel                     ; If it fails
{
  Clipboard := ClipSaved          ; Restore saved clipboard
  return                          ; Stop Script
}
Run notepad.exe                   ; Open notepad
WinWait, Untitled - Notepad, , 1  ; Wait until notepad is active
if ErrorLevel
{
  Clipboard := ClipSaved          ; Restore saved clipboard
  return                          ; Stop Script
}
 IfWinExist, Untitled - Notepad   ; If there is an untitled notepad
    WinActivate                   ; Then use that instead
Send ^v                           ; Paste in stuff from the clipboard
Clipboard := ClipSaved            ; Restore saved clipboard to current clipboard
ClipSaved =                       ; Empty the savedclipboard variable
return

#d::                        ; When Winkey + d is pressed
ClipSaved := ClipboardAll   ; Save current clipboard to variable
Clipboard =                 ; Empty cliboard
Send ^c                     ; Send Control+C (Copy)
ClipWait,5                  ; Wait for 5 seconds or until the clipboard contains text.
if ErrorLevel               ; If it fails
{
  MsgBox, No data avaialble for the clipboard, remember to mark the text you need.  ; Display messagebox
  Clipboard := ClipSaved                                                            ; Restore saved clipboard
  return                                                                            ; Stop Script
}
Link = %Clipboard%                                                                  ; Put clipboard into link variable
Clipboard := ClipSaved                                                              ; Restore saved clipboard
ClipSaved =                                                                         ; Clear saved clipboard
DFilename =                                                                         ; Make variable for filename only 
SplitPath, Link, DFilename                                                          ; Get filename from link
Dpath = C:\Users\olesena\Documents\Reports\AutoHotkey Logs\                         ; Download path
MsgBox, 0, Auto-open file, File is being downloaded and opened..., 1                ; Display messagebox
UrlDownloadToFile, %Link%, %Dpath%%DFilename%                                       ; Download the file
run "%Dpath%%DFilename%"                                                            ; Open the file
return                                                                              ; Stop script

; >>>>>>>>>>>>>>>>>>
; >>> Note maker >>>
; >>>>>>>>>>>>>>>>>>
#IfWinActive Notepad++
::cn::
  loggingTimeStamp = 
  FormatTime, loggingTimeStamp, , H:mm dd/MM-yy
  SendInput {Control Down}s{Control Up}
  SendInput {Control Down}{End}{Control Up}
  SendInput {Enter 2}-- %loggingTimeStamp% -----------------------{Enter}
  SendInput {Enter}
  SendInput HP CASE ID{Space 4}:{Space}{Enter}
  SendInput Customer Name{Space}:{Space}{Enter}
  SendInput Model Name{Space 4}:{Space}{Enter}
  SendInput Serial Number{Space}:{Space}{Enter}
  SendInput Telephone{Space 5}:{Space}{Enter}
  SendInput E-Mail{Space 8}:{Space}{Enter}
  SendInput {Space 3}<{Space}Notes:{Space}>{Enter}
  SendInput >--------------<{Enter 2}
  SendInput >--------------<
  SendInput {Home}{BACKSPACE 3}{UP}{DEL 3}{UP}{DEL 3}
  SendInput {Up 6}{End}
  SendInput {Control Down}s{Control Up} 
return
#IfWinActive
; <<<<<<<<<<<<<<<<<<
; <<< Note maker <<<
; <<<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> Simple Text macros that work ONLY in SR-Dash >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#IfWinActive ahk_class WindowsForms10.Window.8.app.0.218f99c
::critcase::
SendInput >-----------------< Current status >------------------<{Enter}
SendInput Reported problem:{Enter}{Space}->{Space}{Enter 2}
SendInput Overall severity of current issue:{Enter}{Space}->{Space}{Enter 2}
SendInput Issue relates to software or hardware?:{Enter}{Space}->{Space}{Enter 2}
SendInput Next action needed:{Enter}{Space}->{Space}{Enter}
SendInput >-------------------------------------<{Up 10}{End}
return
^Enter::SendInput {Enter}{Space}->{Space}
::cip::Checked in Partsurfer
:o:tcou::TCO Callback Unsuccessful:{Enter}{Space}->{Space}
:o:tcounvm::TCO Callback Unsuccessful:{Enter}{Space}->{Space}Did not reach anyone - There was no option for voicemail{Enter}{Space}->{Space}
:o:tcouvm::TCO Callback Unsuccessful:{Enter}{Space}->{Space}Went to voicemail - Left voicemail - Callback Failed{Enter}{Space}->{Space}
::mailf::Sending mail to follow up
::mailp::
 NFDT := NextFollowUpDateTime()
 SendInput Mail sent:{Enter}{Space}-> Pending for reply{Enter}------------{Enter}%NFDT%
 return
::maill::Mail sent:{Enter}{Space}->{Space}Pending for reply{Enter}{Space}->{Space}I will try to call later today to see if I can reach the customer then.
:o:tcos::TCO Callback Successful:{Enter}{Space}->{Space}Reached{Space}
::tryvr::Trying to start Virtual Room session
::vrtry::Trying to start Virtual Room session
::vrstart::Virtual Room session has been successfully started
::startvr::Virtual Room session has been successfully started
::trymy::Trying to start HP My Room session
::mytry::Trying to start HP My Room session
::mystart::HP My Room session has been successfully started
::startmy::HP My Room session has been successfully started
::tryml::Trying to start MS Lync remote screen sharing session
::mltry::Trying to start MS Lync remote screen sharing session
::mlstart::MS Lync remote screen sharing session has been successfully started
::startml::MS Lync remote screen sharing session has been successfully started
:o:vrses::VR-Session:{Enter}{Space}->{Space}
:o:myses::HP MyRoom-Session:{Enter}{Space}->{Space}
:o:mlses::MS Lync remote screen sharing session:{Enter}{Space}->{Space}
::rcall::
 ClipSaved := ClipboardAll								;Save current clipboard to variable
 Clipboard =  										;Empty cliboard
 Sleep, 25
 SendInput Calling{SPACE}{TAB 13}
 SendInput {UP 70}{DOWN 2}{SHIFTDOWN}{DOWN 5}{SHIFTUP}
 SendInput {CTRLDOWN}c{CTRLUP}{ALTDOWN}n{ALTUP}{TAB 3}
 SendInput {CTRLDOWN}v{CTRLUP}{UP 6}{END}{SPACE}
 SendInput {DEL}{END}{SPACE}{DEL}{SHIFTDOWN}{DOWN 3}{HOME}{SHIFTUP}
 SendInput {DEL}{END}{SPACE}to follow up
 Clipboard := ClipSaved									;Restore Clipboard after script is done
return
::dat::Drive Assessment Test
::datf::Drive Assessment Test - Failed - Drive is no longer recommended for use.
::dats::Drive Assessment Test - Passed - Drive is GOOD.
::av::Address has been verified with the customer.
::csra::CSR delivery of the part (part delivered without technician for the installation) has been accepted by 
::ceneed::Certified Engineer for the the onsite installation of this part has been reqeuested by 
::secop::
SendInput >-------------<{Space}SECOND{Space}OPINION{Space}>-------------<{Enter}
SendInput Checked troubleshooting: Yes{Enter}
SendInput Checked part number: Yes{Enter}
SendInput Checked RD/T: Yes{Enter}
SendInput Checked QA codes: Yes{Enter}
SendInput SBD / NBD checked: Yes{Enter}
SendInput EQT used: Yes{Enter}
SendInput CDS Compliant: Yes{Enter}
SendInput KAIZEN Compliant: Yes{Enter}
SendInput >-------------<{Space}SECOND{Space}OPINION{Space}>-------------<
return

::moncase::
 SendInput Case will be monitored for the next few days{Enter}-----------------{Enter}
 NFDT := NextFollowUpDateTime(2)
 SendInput %NFDT%
return

::nfdt::
 NFDT := NextFollowUpDateTime()
 SendInput %NFDT%
return

::nfdt2::
 NFDT := NextFollowUpDateTime(2)
 SendInput %NFDT%
return

::nfdt3::
 NFDT := NextFollowUpDateTime(3)
 SendInput %NFDT%
return

::nfdt4::
 NFDT := NextFollowUpDateTime(4)
 SendInput %NFDT%
return

::nfdt5::
 NFDT := NextFollowUpDateTime(5)
 SendInput %NFDT%
return

::nfdt6::
 NFDT := NextFollowUpDateTime(6)
 SendInput %NFDT%
return
#IfWinActive
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<< Simple text macros that work ONLY in SR-Dash <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> Text based macros that work in Telegram >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#IfWinActive ahk_exe Telegram.exe
:::):::-)
:::(:::-(
:::/:::-/
:::p:::-p
:::s:::-s
:::d:::-D
#IfWinActive
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<< Text based macros that work in Telegram <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
; >>> Simple text macros that work everywhere >>>
; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
::l1e::Level 1 Engineer
::l2e::Level 2 Engineer
::l3e::Level 3 Engineer
::l4e::Level 4 Engineer
::lalain::Level 3 Engineer Alain Vercammen
:o:qt::>-----------------<>------------------<{ENTER 2}>-------------------------------------<{UP 2}{RIGHT 19} 

return 
::HPTT::http://www.hp.com/support/tapetools
:C1:ltts::HP Library and Tape Tools Support ticket
::ty::Thank you
::tu::Thank you
::yw::You're welcome
::thx::Thanks
::brb::be right back
::afk::I am not available from now on
::ilm::i lige måde
::np::That is not a problem
::nw::No worries at all
::jsut::just
::ans::and
::fro::for
::eod::end of day
::doign::doing
::netwrok::network
:C1:k::Okay
:C1:ok::Okay
:C1:oki::Okay
:C1:vis::via
:C1:ms lync::MS Lync
:C1:hp l&tt::HP Library and Tape Tools
::customre::customer
::seraching::searching
::afaik::as far as I know
::atm::at the moment
::ccl::Case is okay to close.
::incall::I'm busy with a call at the moment, please write to me for now, and I will answer later.
::scr1::You will learn that quite a lot of things I say are scripted.. :-)
:C1:fw::Firmware
:C1:sw::Software
:C1:pw::Password
:C1:hw::Hardware
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<< Simple text macros that work everywhere <<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
