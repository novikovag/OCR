#include <File.au3>
#include <Array.au3>
#include <Word.au3>
#include <MsgBoxConstants.au3>

$iPath = "C:\X-Files\Text\mac\docx\"
$oPath = "C:\X-Files\Text\mac\post\stage1\txt\odt.out\"

$word = _Word_Create()

For $folder IN _FileListToArray($iPath, Default, 2)
    If Not FileExists($iPath & "\" & $folder & "\wcount.txt") Then
        MsgBox($MB_SYSTEMMODAL, "", "The wcount.txt doesn't exist!!!")
        Exit -1
    EndIf

    DirCreate($oPath & "\" & $folder);

    FileCopy($iPath & "\" & $folder & "\wcount.txt", $oPath & "\" & $folder & "\wcount.txt", 1)

    For $file IN _FileListToArray($iPath & "\" & $folder)
       If StringRegExp($file, ".docx", 0) ==  1 Then
          ;ConsoleWrite($file & @CRLF)
          $doc = _Word_DocOpen($word, $iPath & "\" & $folder & "\" & $file)


          $file = StringReplace ($file, ".docx", "")
          _Word_DocSaveAs($doc, $oPath & "\" & $folder & "\" & $file & ".odt", 23)

          _Word_DocClose($doc)
       EndIf
    Next
Next

