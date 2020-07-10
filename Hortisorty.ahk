;Hortisorty
;Made by Skullfurious and nou!
;if you run into issues delete your config in documents/autohotkey/hortisorty

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.#SingleInstance, force
#SingleInstance force
#Persistent

Menu,Tray,Icon,favicon.ico

if !FileExist(A_MyDocuments "\AutoHotKey\Hortisorty")
{
  FileCreateDir, %A_MyDocuments%\AutoHotKey\Hortisorty
  MsgBox, Creating Directory Documents\AutoHotkey\Hortisorty
}

if !FileExist(A_MyDocuments "\AutoHotKey\Hortisorty\config.ini")
{
  initext =
  (
[Prices]
DenoteCurrency=false
ConvertToExalted=false
ExaltedValue=175

[Augment]
Attack=60
Caster=60
Chaos=60
Cold=50
Critical=110
Defence=60
Fire=50
Influence=60
Life=60
Lightning=50
Physical=60
Random=60
Speed=100

[AugmentLucky]
Attack=90
Caster=90
Chaos=90
Cold=90
Critical=200
Defence=90
Fire=90
Influence=90
Life=90
Lightning=90
Physical=90
Random=90
Speed=180

[Remove]
Attack=50
Caster=50
Chaos=50
Cold=50
Critical=80
Defence=50
Fire=50
Influence=80
Life=50
Lightning=50
Physical=50
Random=50
Speed=70

[RemoveAdd]
Attack=80
Caster=80
Chaos=80
Cold=80
Critical=120
Defence=80
Fire=80
Influence=80
Life=80
Lightning=80
Physical=80
Speed=100

[RemoveNon]
Attack=60
Caster=60
Chaos=60
Cold=60
Critical=100
Defence=60
Fire=60
Influence=60
Life=60
Lightning=60
Physical=60
Speed=80

[Randomise]
Attack=40
Caster=40
Chaos=40
Cold=40
Critical=80
Defence=40
Fire=40
Influence=60
Life=50
Lightning=40
Physical=40
Speed=60

[Socket]
Five=30
Six=450
  )
  FileAppend, %initext%, %A_MyDocuments%\AutoHotkey\Hortisorty\config.ini
}

OnClipboardChange("CheckClipData")

;Init Gui's and whatnot
global listeningwarned = false
global listening := false
global craftcontainer := []
global DataArray := [[],[]]
global CharName := ""
global SheetName := ""

Gui, menu:New, -MaximizeBox -MinimizeBox -SysMenu
Gui, menu:Add, Text, , Spreadsheet Name:
;Gui, menu:Add, Text, , Account Name:
Gui, menu:Add, Edit, vSheetName w140 ym, MySheet
;Gui, menu:Add, Edit, vCharName w140, Tom Nook
Gui, menu:Add, Button, Default w60 gclosemenu, OK
Gui, menu:Show,, HortiSorty
return

;Reload
F5::
MsgBox, Reloaded!
Reload
Return

;End Search
F1::
MsgBox, Alright! Lets make that spreadsheet!
sortdata()
return

;Functions
CloseMenu()
{

  Gui, menu:Submit

  if FileExist(A_MyDocuments "\AutoHotKey\Hortisorty\" Sheetname ".csv")
  {
    MsgBox, 52, Overwrite Warning, A spreadsheet with the name %SheetName% already exists. Are you sure you want to continue? Data will be lost.
    IfMsgBox Yes
    {
      listening = true
      FileDelete, %A_MyDocuments%\AutoHotkey\Hortisorty\%SheetName%.csv
      MsgBox, 32,HortiSorty, HortiSorty is now listening to your clipboard changes. `n`nIf you press CTRL + C on a horticraft station it will add it to the resulting spreadsheet.`n`nStop this at any time by pressing F1 ending the process. `n`nTHIS WILL NOT CHECK FOR DUPLICATES. BE CAREFUL!
      return
    }
    IfMsgBox No
    {
      MsgBox Alright, we'll discontinue... for now!
      ExitApp
    }
    return
  }
  
  listening = true
  MsgBox, 32,HortiSorty, HortiSorty is now listening to your clipboard changes. `n`nIf you press CTRL + C on a horticraft station it will add it to the resulting spreadsheet.`n`nStop this at any time by pressing F1 ending the process. `n`nTHIS WILL NOT CHECK FOR DUPLICATES. BE CAREFUL!

  return
}

;Check Clip Data
CheckClipData()
{
  clip := Clipboard
  global lines := strSplit(clip, "`n")

  ;Warn user we aren't listening once.
  if !listening and !listeningwarned
  {
    MsgBox, HortiSorty is not currently listening for Horticraft Stations. `n`nIf you are aware and don't care, continue, if you are confused, close the program and open it again.
    listeningwarned = true
    return
  }

  ;If we aren't listening for changes, do nothing.
  if !listening
  {
    return
  }

  if listening 
  {
    ;If Format doesnt match the horti stations, we stop.
    if !(lines[2] ~= "Horticrafting Station")
      return
    
    RegExMatch(lines.5, "\d", craftamount)

    while(craftamount > 0)
    {
      position := 7 + craftamount
      craftcontainer.push(lines[position])
      craftamount := craftamount - 1
    }

  }

}

SortData()
{
  MsgBox, Sorting data!

  global augment := []
  global remove := [] 
  global removeadd := []
  global removenon := []
  global randomise := [] 

  for k, v in craftcontainer
  {
    RegExMatch(v, "(?i)(remove|randomise|augment)", type)
    RegExMatch(v, "(?i)(Physical|Defence|Fire|Lightning|Cold|Caster|Attack|Life|Chaos|Critical|Influence|Speed|Sockets)", tag)
    non := RegExMatch(v, "(?i)(non)")
    lucky := RegExMatch(v, "(?i)(lucky)")
    add := RegExMatch(v, "(?i)(add )")
    adding := RegExMatch(v, "(?i)(adding)") ; filters out "Upgrade a Magic item to a Rare item, adding fo..." mods
    RegExMatch(v, "\d\d", level)

    if (type == "Augment")
    {

      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
      augment.push([tag,level,price])
    }

    if (type == "Remove")
    {
      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
      remove.push([tag,level,price])
    }
    
    if (type == "Randomise")
    {
      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
      randomise.push([tag,level,price])
    }

    if (non > 0)
    {
      type := "RemoveNon"
      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
      removenon.push([tag,level,price])
    }

    if (lucky > 0) and (type == "Augment")
    {
      type := "AugmentLucky"
     IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
     augment.push([tag . " Lucky",level,price])
    }

    if (lucky > 0) and !(type == "Augment")
    {
      type := "Reroll"
      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
    }

    if (add > 0) and (non < 1)
    {
      type := "RemoveAdd"
      offset := RemoveAddOffset * 4
      IniRead, price, %A_MyDocuments%\AutoHotKey\Hortisorty\config.ini, % type, % tag
      removeadd.push([tag,level,price])
    }
    
  }

  vMax := Max(augment.count(), remove.count(), removeadd.count(), randomise.count(), removenon.count())

  template = 
  (
Augment Modifier,,,,Remove Modifier,,,,Remove Add Modifier,,,,Remove Non Modifier,,,,Randomize Modifier Values,,,
Type,Tag,Item Level,Price,Type,Tag,Item Level,Price,Type,Tag,Item Level,Price,Type,Tag,Item Level,Price,Type,Tag,Item Level,Price`n
  )

  MsgBox, % SheetName
  FileAppend, %template%, %A_MyDocuments%\AutoHotkey\Hortisorty\%SheetName%.csv

  global str := ""

  Loop, % vMax 
  {
    ; aug category
    if (augment[A_Index].count() > 0)
    {
      str .= "Augment,"
      . augment[A_Index][1] "," 
      . augment[A_Index][2] "," 
      . augment[A_Index][3] ","
    }
    else
    {
      str .= ","
      . augment[A_Index][1] "," 
      . augment[A_Index][2] "," 
      . augment[A_Index][3] ","
    }

    ; remove category 
    if (remove[A_Index].count() > 0)
    {
      str .= "Remove,"
      . remove[A_Index][1] "," 
      . remove[A_Index][2] "," 
      . remove[A_Index][3] ","
    }
    else
    {
      str .= ","
      . remove[A_Index][1] "," 
      . remove[A_Index][2] "," 
      . remove[A_Index][3] ","
    }

    ;remove add category
    if (removeadd[A_Index].count() > 0)
    {
      str .= "Remove Add,"
      . removeadd[A_Index][1] "," 
      . removeadd[A_Index][2] "," 
      . removeadd[A_Index][3] ","
    }
    else
    {
      str .= ","
      . removeadd[A_Index][1] "," 
      . removeadd[A_Index][2] "," 
      . removeadd[A_Index][3] ","
    }

    ;removenon category
    if (removenon[A_Index].count() > 0)
    {
      str .= "Remove Non,"
      . removenon[A_Index][1] "," 
      . removenon[A_Index][2] "," 
      . removenon[A_Index][3] ","
    }
    else
    {
      str .= ","
      . removenon[A_Index][1] "," 
      . removenon[A_Index][2] "," 
      . removenon[A_Index][3] ","
    }

    ;randomise category
    if (randomise[A_Index].count() > 0)
    {
      str .= "Randomise,"
      . randomise[A_Index][1] "," 
      . randomise[A_Index][2] "," 
      . randomise[A_Index][3] "`n"
    }
    else
    {
      str .= ","
      . randomise[A_Index][1] "," 
      . randomise[A_Index][2] "," 
      . randomise[A_Index][3] "`n"
    }
  }

  FIleAppend, %str%, %A_MyDocuments%\AutoHotkey\Hortisorty\%SheetName%.csv

  return
}