#NoEnv
#Warn
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%

;--------------Setup and train the script---------------------
CoordMode, Mouse

;Search Bar
Hotkey, LButton, Off
MsgBox Setting of infoodle, Please press on Search bar
Send {Home}
sleep(200)
Hotkey, LButton, On
Sleep 2000
searchBarX := xpos
searchBarY := ypos
selectionX := searchBarX
selectionY := searchBarY + 40
Hotkey, LButton, Off

;Note button
MsgBox Set coordinates to %searchBarX% and %searchBarY%, Please press the  Note now.
Send {Home}
sleep 200
Hotkey, LButton, On
Sleep 2000
noteButtonX := xpos
noteButtonY := ypos
Hotkey, LButton, Off
MsgBox Set coordinates to %noteButtonX% and %noteButtonY%, Please press the colour code now.

;Colour Code
Send {Home}
Sleep 200
Click, %noteButtonX% %noteButtonY%
Hotkey, LButton, On
Sleep 2000
colourCodeX := xpos
colourCodeY := ypos
Hotkey, LButton, Off
MsgBox Set coordinates to %colourCodeX% and %colourCodeY%, please press the add account button

;Add people account
Send {Home}
Sleep 200
Hotkey, LButton, On
Sleep 2000
addButtonX := xpos
addButtonY := ypos
Hotkey, LButton, Off
MsgBox Set coordinates to %addButtonX% and %addButtonY%. All setup and ready to go


;------------Functions-----------------------
checkForAccount()
{
  global
  Send {Home}
  Send ^c
  sleep()
  email := clipboard
  Send ^{Tab}
  sleep()
  Send {Home}
  sleep(200)
  Click %searchBarX% %searchBarY%
  sleep()
  Send ^v
  sleep()
  MsgBox, 36, Is there a record?, Please yes for record and no for unfound
  IfMsgBox Yes
  { ;Go to note fill out most of the interaction
    Click %selectionX% %selectionY%
    Sleep 3000
    fillOutNote()
  }
  else
  { ;Go to add a new person
    ;Head back to workbook and get information
    createInfoodleAccount()
  }
}

/**
*This should be invoked on the date column of the final attempt.
*/
finalAttemptInteraction()
{
  global
  Send ^c
  Sleep 200
  date := clipboard
  Send {Home}
  Send ^c
  Sleep 100
  Send ^{Tab}
  Sleep 100
  Send {Home}
  Sleep 200
  Click %searchBarX% %searchBarY%
  Sleep 100
  Send ^v

  MsgBox, 36, Is there a record?, Please yes for record and no for unfound
  IfMsgBox Yes
  {
    Click %selectionX% %selectionY%
    Sleep 3000
    tempNote = TF %date%, no answer on 3rd attempt
    fillOutNote(tempNote, true)
    ;Not yet implemented
    ;~ Sleep 2000 ;Waiting for the form to be submitted
    ;~ Send ^+{Tab}
    ;~ Sleep 400
    ;~ Send {Right}
    ;~ Sleep 100
    ;~ Send {Right}
    ;~ Sleep 100
    ;~ Send {Enter}
    ;~ Sleep 100
    ;~ Send ^{Backspace}
    ;~ Sleep 100
    ;~ Send True
    ;~ Sleep 100
    ;~ Send {Tab}
    ;~ Sleep 100
    ;~ Send {Space}
    ;~ Loop 9
    ;~ {
      ;~ Send {Right}
      ;~ Sleep 100
    ;~ }
  }
  else
  {
    createInfoodleAccount()
    ;Not implemented yet. Where it will automatically create account and say so in the workbook.
    ;~ Sleep 100
    ;~ Send {Backspace}
    ;~ Sleep 100
    ;~ Send ^+{Tab}
    ;~ Sleep 400
    ;~ Loop 2
    ;~ {
      ;~ Send {Right}
      ;~ Sleep 100
    ;~ }
    ;~ Send {Enter}
    ;~ Sleep 100
    ;~ Send ^{Backspace}
    ;~ Sleep 100
    ;~ Send False
    ;~ Sleep 100
    ;~ Send {Tab}
    ;~ Sleep 100
    ;~ Send {Right}
    ;~ Sleep 100
    ;~ Send {Enter}
    ;~ Sleep 100
    ;~ Send Checked for final call attempt interaction
    ;~ Sleep 100
    ;~ Send {Tab}
    ;~ Loop 7
    ;~ {
      ;~ Send {Right}
      ;~ Sleep 100
    ;~ }
  }

}

/**
*Should be called from infoodle homepage
*/
fillOutNote(note := "TF,", autoFinish := false)
{
  Global
  CoordMode, Mouse
  Send {Home}
  Click, %noteButtonX% %noteButtonY%
  Sleep 100
  tab(2)
  ;Type of note
  Send {Enter}
  Sleep 100
  Send {End}
  Sleep 100
  Send {Up}
  Sleep 100
  Send {Enter}
  Sleep 100
  ;Click colour code
  Click, %colourCodeX% %colourCodeY%
  Sleep 100
  tab(2)

  ;Who is note for
  Send {Space}
  Sleep 100
  tab()
  ;Note
  Send  %note%
  Sleep 300
  tab(3, 300)
  ;Lock to top
  Send {Space}
  Sleep 300
  tab(,300)
  Send {Space}
  Sleep 300

  if (autoFinish) {
    Send {Tab}
    Sleep 300
    ;Save note
    Send {Enter}
  } else
  {
    loop 3
    {
      Send +{Tab}
      sleep()
    }
  }
}

/**
*Should be called from infoodle home page.
*/
createInfoodleAccount()
{
  global
  Send ^+{Tab}
  sleep()
  Send {Home}
  tab(2)
  Send ^c
  sleep()
  firstName := Clipboard
  tab()
  Send ^c
  sleep()
  lastName := Clipboard
  tab(3)
  Send ^c
  sleep()
  phoneNumber := Clipboard

  ;Go to infoodle and add person
  Send ^{Tab}
  sleep()
  click %addButtonX% %addButtonY%
  sleep(2000)
  tab()
  Send %firstName%
  tab()
  Send %lastName%
  tab(3)
  Send %email%
  tab(3)
  Send %phoneNumber%
  tab()
  sleep(200)
  Send {Space}
  tab(5)
  Send Voted on MGC
  tab(,2200)
  tab()
  Send Email Subscribers
  tab()
}

tab(number := 1, sleepTime := 100) {
  loop %number%
  {
    Send {Tab}
    sleep(sleepTime)
  }
}

sleep(sleepTime := 100) {
  Sleep %sleepTime%
}

;--------Hotkeys---------------
<!L::finalAttemptInteraction()
Alt::checkForAccount()

LButton:: MouseGetPos, xpos, ypos