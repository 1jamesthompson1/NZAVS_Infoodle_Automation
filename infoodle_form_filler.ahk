#NoEnv
#Warn
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%

;--------------Setup---------------------
CoordMode, Mouse
Hotkey, LButton, Off
MsgBox Setting of infoodle, Please press on Search bar
Hotkey, LButton, On
Sleep 2000
searchBarX := xpos
searchBarY := ypos
selectionX := searchBarX
selectionY := searchBarY + 40
Hotkey, LButton, Off
MsgBox Set coordinates to %searchBarX% and %searchBarY%, Please press the  Note now.
Hotkey, LButton, On
Sleep 2000
noteButtonX := xpos
noteButtonY := ypos
Hotkey, LButton, Off
MsgBox Set coordinates to %noteButtonX% and %noteButtonY%, Please press the colour code now.

Send {Home}
Sleep 200
Click, %noteButtonX% %noteButtonY%
Hotkey, LButton, On
Sleep 2000
colourCodeX := xpos
colourCodeY := ypos
Hotkey, LButton, Off
MsgBox Set coordinates to %colourCodeX% and %colourCodeY%. All setup and ready to go



;------------Functions-----------------------
fillOutInteraction()
{
  global
  Send ^c
  Sleep 200
  date := clipboard
  loop 12
  {
    Send {Left}
    Sleep 100
  }
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
      fillOutForm()
      Sleep 2000 ;Waiting for the form to be submitted
      Send ^+{Tab}
      Sleep 400
      Send {Right}
      Sleep 100
      Send {Right}
      Sleep 100
      Send {Enter}
      Sleep 100
      Send ^{Backspace}
      Sleep 100
      Send True
      Sleep 100
      Send {Tab}
      Sleep 100
      Send {Space}
      Loop 9
      {
        Send {Right}
        Sleep 100
      }
  }
  else
  {
      Click %searchBarX% %searchBarY%
      Sleep 10
      Click %searchBarX% %searchBarY%
      Sleep 10
      Click %searchBarX% %searchBarY%
      Sleep 100
      Send {Backspace}
      Sleep 100
      Send ^+{Tab}
      Sleep 400
      Loop 2
      {
        Send {Right}
        Sleep 100
      }
      Send {Enter}
      Sleep 100
      Send ^{Backspace}
      Sleep 100
      Send False
      Sleep 100
      Send {Tab}
      Sleep 100
      Send {Right}
      Sleep 100
      Send {Enter}
      Sleep 100
      Send Checked for final call attempt interaction
      Sleep 100
      Send {Tab}
      Loop 7
      {
        Send {Right}
        Sleep 100
      }
  }

}

fillOutForm()
{
  Global
  CoordMode, Mouse
  Send {Home}
  Sleep 700
  ;Open up note form
  Click, %noteButtonX% %noteButtonY%
  Sleep 100
  Send {Tab}
  Sleep 100
  Send {Tab}
  Sleep 100
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
  Send {Tab}
  Sleep 100
  Send {Tab}
  Sleep 100

  ;Who is note for
  Send {Space}
  Sleep 100
  Send {Tab}
  Sleep 100
  ;Note
  Send  TF %date%, no answer on 3rd attempt
  Sleep 300
  Send {Tab}
  Sleep 300
  Send {Tab}
  Sleep 300
  Send {Tab}
  Sleep 300
  ;Lock to top
  Send {Space}
  Sleep 300
  Send {Tab}
  Sleep 300
  Send {Space}
  Sleep 300
  Send {Tab}
  Sleep 300
  ;Save note
  Send {Enter}
}


;--------Hotkeys---------------

Ins::fillOutInteraction()
+Ins::fillOutForm()

LButton:: MouseGetPos, xpos, ypos