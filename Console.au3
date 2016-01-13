#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <GuiListView.au3>
#include <GuiButton.au3>

Opt("TrayIconHide", 1)

Global $item[1000]
Global $rffile[1000]
Global $x
Global $i
Global $rd2
Global $rd3
Global $rd4
Global $wd
Global $dbpath

Func mdm()
   $dbpath = "\\172.16.190.14\maildb\users.dbm"
EndFunc

Func mssu()
   $dbpath = "\\172.16.225.5\maildb\users.dbm"
EndFunc

Func nnkk()
   $dbpath = "\\172.16.166.5\maildb\users.dbm"
EndFunc

Func krkk()
   $dbpath = "\\172.16.162.4\maildb\users.dbm"
EndFunc

Func ufa()
   $dbpath = "\\172.16.153.3\maildb\users.dbm"
EndFunc

$form0 = GUICreate("Выбор филиала",260,320,100,100)
GUISetState(@SW_SHOW)
$but1 = GUICtrlCreateButton("МДМ",30,15,200,50)
$but2 =GUICtrlCreateButton("Сущевка",30,75,200,50)
$but3 =GUICtrlCreateButton("Н-Новгород",30,135,200,50)
$but4 =GUICtrlCreateButton("Красноярск",30,195,200,50)
$but5 =GUICtrlCreateButton("УФА",30,255,200,50)

While 1
   $nMsg = GUIGetMsg()
	  If $nMsg = $GUI_EVENT_CLOSE Then 	Exit

	  If $nMsg = $but1 Then
		 mdm()
		 ExitLoop
	  EndIf

	  If $nMsg = $but2 Then
		 mssu()
		 ExitLoop
	  EndIf

	  If $nMsg = $but3 Then
		 nnkk()
		 ExitLoop
	  EndIf

	  If $nMsg = $but4 Then
		 krkk()
		 ExitLoop
	  EndIf

	  If $nMsg = $but5 Then
		 ufa()
		 ExitLoop
	  EndIf
WEnd

GUIDelete($form0)

Func find_i()
   $findi = GUICtrlRead($edit1)
   $iI = _GUICtrlListView_FindInText($listview1,$findi)
   _GUICtrlListView_ClickItem($listview1,$iI,"left")
EndFunc

Func write_db()
   $file = FileOpen($dbpath,1)
   FileWrite($file,@CRLF)
   FileWrite($file,$rd2 &";"& $rd3 & ";" & $rd4)
   FileClose($file)
   GUICtrlCreateListViewItem($rd2 & "|" & $rd3 & "|" & $rd4  , $listview1)
EndFunc

Func del_lv()
   $a = _GUICtrlListView_GetItemTextString($listview1,-1)
   $a = StringReplace($a,"|",";")
   If _ReplaceStringInFile($dbpath, $a, " ; ; ") = -1 Then MsgBox (4096,"","Чето не так !")
   _GUICtrlListView_DeleteItemsSelected($listview1)
EndFunc

Func add_lv()
   $rd2 = GUICtrlRead($edit2)
   $rd3 = GUICtrlRead($edit3)
   $rd4 = GUICtrlRead($edit4)
   If $rd2 <> "" and $rd3 <> "" and $rd4 <> ""  Then
	  write_db()
   Else
	  MsgBox(16,"Внимание !","Для добавления необходимо заполнить все поля")
   EndIf
   GUICtrlSetData($edit2,"")
   GUICtrlSetData($edit3,"")
   GUICtrlSetData($edit4,"")
   MsgBox(4096,"console","Пользователь "& $rd2 & " успешно добавлен !")
EndFunc

$cl = _FileCountLines($dbpath)
$file = FileOpen($dbpath,0)
;MsgBox(0, $cl, $file)
For $j = 1 To $cl
   $rffile[$j] = FileReadLine($file,$j)
Next
FileClose($file)
$Form1 = GUICreate("mail console", 615, 438, 192, 124)
$ListView1 = GUICtrlCreateListView("AD Login                       |MailAccount | MailPass    ", 16, 8, 313, 417)
For $i = 1 To $cl
   $item = StringSplit($rffile[$i],";")
   If $item[1]<>" " Then GUICtrlCreateListViewItem($item[1] & "|" & $item[2] & "|" & $item[3] , $listview1)
Next

$edit1 = GUICtrlCreateInput("",370,24,99,25)
$Button1 = GUICtrlCreateButton("Найти", 476, 24, 99, 25)
$edit2 = GUICtrlCreateInput("",420,80,110,25)
$edit3 = GUICtrlCreateInput("",420,110,110,25)
$edit4 = GUICtrlCreateInput("",420,140,110,25)
$Button2 = GUICtrlCreateButton("<=", 540, 80, 50, 85)
GUICtrlCreateLabel("AD",350,85,40,25)
GUICtrlCreateLabel("Mailbox",350,115,40,25)
GUICtrlCreateLabel("Password",350,145,45,25)
$Button3 = GUICtrlCreateButton("Удалить", 350, 180, 100, 25)
$Button4 = GUICtrlCreateButton("Создание почтового ящика", 455, 180, 150, 25)

GUISetState(@SW_SHOW)

While 1
   $nMsg = GUIGetMsg()
   If $nMsg = $GUI_EVENT_CLOSE Then Exit
   If $nMsg = $button2 Then add_lv()
   If $nMsg = $button3 Then del_lv()
   If $nMsg = $button1 Then find_i()
   If $nMsg = $button4 Then ShellExecute("iexplore.exe", "http://172.16.210.201/admin/")
WEnd
