#NoEnv
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%

version := "v0.2.0"
programName := "NZAVS Automation Script "version


;--------------Setup and train the script---------------------
loop
{
  if !FileExist("config.ini")
  { ;----First time setup
    MsgBox, 64, %programName%, Setting up Automation script for first use.
    IniWrite, %version%, config.ini, Info, version
    setupMousePositions()
    changeDefaultNote()
  }
  else
  { ;----Normal startup
    IniRead, versionNumber, config.ini, Info, version
    if (version != versionNumber) {
      FileDelete, config.ini
      continue
    }
    MsgBox, 35, %programName%, Do you want to keep current settings?
    IfMsgBox, Cancel
      ExitApp
    IfMsgBox No
    {  
      MsgBox, 36, Settings Change, Do you want to change mouse position or default note (yes for mouse and no for default note)?
      IfMsgBox Yes
        setupMousePositions()
      IfMsgBox No
        changeDefaultNote()
    }
    IniRead, searchBarX, config.ini, Button Locations, searchBarX
    IniRead, searchBarY, config.ini, Button Locations, searchBarY
    IniRead, selectionX, config.ini, Button Locations, selectionX
    IniRead, selectionY, config.ini, Button Locations, selectionY
    IniRead, noteButtonX, config.ini, Button Locations, noteButtonX
    IniRead, noteButtonY, config.ini, Button Locations, noteButtonY
    IniRead, colourCodeX, config.ini, Button Locations, colourCodeX
    IniRead, colourCodeY, config.ini, Button Locations, colourCodeY
    IniRead, addButtonX, config.ini, Button Locations, addButtonX
    IniRead, addButtonY, config.ini, Button Locations, addButtonY
    IniRead, noteTemplate, config.ini, Settings, defaultNote
    break
  }
}



;------------Functions-----------------------
setupMousePositions()
{
  global
  CoordMode, Mouse
  ;Search Bar
  MsgBox Setting of infoodle, Please press on Search bar
  MouseDisabling := !MouseDisabling
  Send {Home}
  sleep(200)
  KeyWait, LButton, D
  MouseGetPos, xpos, ypos
  searchBarX := xpos
  searchBarY := ypos
  selectionX := searchBarX
  selectionY := searchBarY + 50

  ;Note button
  MouseDisabling := !MouseDisabling
  MsgBox Set coordinates to %searchBarX% and %searchBarY%, Please press the  Note now.
  MouseDisabling := !MouseDisabling
  Send {Home}
  sleep 200
  KeyWait, LButton, D
  MouseGetPos, xpos, ypos
  noteButtonX := xpos
  noteButtonY := ypos
  MouseDisabling := !MouseDisabling
  MsgBox Set coordinates to %noteButtonX% and %noteButtonY%, Please press the colour code now.
  MouseDisabling := !MouseDisabling

  ;Colour Code
  Click, %noteButtonX% %noteButtonY%
  KeyWait, LButton, D
  MouseGetPos, xpos, ypos
  colourCodeX := xpos
  colourCodeY := ypos
  MouseDisabling := !MouseDisabling
  MsgBox Set coordinates to %colourCodeX% and %colourCodeY%, please press the add account button
  MouseDisabling := !MouseDisabling

  ;Add people account
  KeyWait, LButton, D
  MouseGetPos, xpos, ypos
  addButtonX := xpos
  addButtonY := ypos
  MouseDisabling := !MouseDisabling
  MsgBox Set coordinates to %addButtonX% and %addButtonY%. All setup and ready to go.

  IniWrite, %searchBarX%, config.ini, Button Locations, searchBarX
  IniWrite, %searchBarY%, config.ini, Button Locations, searchBarY
  IniWrite, %selectionX%, config.ini, Button Locations, selectionX
  IniWrite, %selectionY%, config.ini, Button Locations, selectionY
  IniWrite, %noteButtonX%, config.ini, Button Locations, noteButtonX
  IniWrite, %noteButtonY%, config.ini, Button Locations, noteButtonY
  IniWrite, %colourCodeX%, config.ini, Button Locations, colourCodeX
  IniWrite, %colourCodeY%, config.ini, Button Locations, colourCodeY
  IniWrite, %addButtonX%, config.ini, Button Locations, addButtonX
  IniWrite, %addButtonY%, config.ini, Button Locations, addButtonY
}

changeDefaultNote()
{
  IniRead, currentNote, config.ini, Settings, defaultNote
  if (currentNote == "ERROR") { ;Create the note
    InputBox, newNote, %programName%, What do you want the note to be?
  } else { ;Change the note
    InputBox, newNote, %programName%, The current note template is "%currentNote%" enter new note if you want
    if ErrorLevel
      return
  }
  IniWrite %newNote%, config.ini, Settings, defaultNote
}

generalInteraction() {
  global
  infoodleAccountExists := checkForAccount()
  if (infoodleAccountExists) {
    fillOutNote(noteTemplate)
  } else {
    createInfoodleAccount()
  }
}

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
  Send ^a
  sleep()
  Send {BackSpace}
  sleep()
  Send ^v
  sleep()
  MsgBox, 36, Is there a record?, Please yes for record and no for unfound
  IfMsgBox Yes
  {
    Click %selectionX% %selectionY%
    MsgBox, 0, %programName%, Has the account loaded yet, 7
    Return true
  }
  else
  {
    Return false
  }
}

/**
*This should be invoked on the date column of the final attempt.
*/
finalAttemptInteraction()
{
  global
  Send ^c
  sleep(200)
  date := clipboard
  Send +{Tab}
  Send ^c
  sleep(200)
  callerName := clipboard

  infoodleAccountExits := checkForAccount()
  if (infoodleAccountExits)
  {
    finalAttemptNote = Telefund - MGC - %date%,`n Contact added by %callerName%.`n No contact on 3rd attempt
    fillOutNote(finalAttemptNote, true)

    MsgBox, 0, %programName%, Has the form submitted yet?, 3 ;Waiting for the form to be submitted
    Send ^+{Tab}
    sleep(200)
    tab(8)
    Send {Enter}
    sleep()
    Send ^{Backspace}
    sleep()
    Send True
    sleep()
    tab()
    Send {Enter}
    sleep()
    Send ^{Backspace}
    sleep()
    Send True
    tab(9)
  }
  else
  {
    createInfoodleAccount()
  }
}

/**
*Should be called from infoodle homepage
*/
fillOutNote(note := "", autoFinish := false)
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

  if (autoFinish) 
  {
    Send {Tab}
    Sleep 300
    ;Save note
    Send {Enter}
  } 
  else
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
  firstName := clipboard
  tab()
  Send ^c
  sleep()
  lastName := clipboard
  tab(3)
  Send ^c
  sleep()
  phoneNumber := clipboard

  ;Go to infoodle and add person
  Send ^{Tab}
  sleep()
  click %addButtonX% %addButtonY%
  MsgBox, 0, %programName%, Has the add account form loaded?, 4
  tab()
  Send %firstName%
  tab(2)
  Send %lastName%
  tab(2)
  Send %email%
  tab(3)
  Send %phoneNumber%
  tab()
  sleep(200)
  Send {Space}
  tab(5)
  Send Voted on MGC
  tab(,2500)
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

<!c::
if (!checkForAccount()) {
  createInfoodleAccount()
}
Return

Alt::generalInteraction()


;----Disable the mousepress

MouseDisabling := 0

#if  MouseDisabling
LButton::return
#if 
