#include <IE.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#Include <Array.au3>


    ; Open the file for reading and store the handle to a variable.
    Local $f_input = FileOpen(@ScriptDir & '\work.txt', $FO_READ)
    If $f_input = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
        Return False
    EndIf

;$f_input = FileRead(@ScriptDir & '\work.txt')

$s_zakaz=FileReadLine($f_input, 1)

;Sleep(5000)
$o_object = _IEAttach("svparfum.ru | CMS.S3")

Func _IETagClassClick($o_Obj, $s_TagName, $s_ClassName, $s_Innertext = '')
    Local $o_Tags
    If Not IsObj($o_Obj) Then Return SetError(1)
    If (Not $s_TagName Or Not $s_ClassName) Then Return SetError(1)
    $o_Tags = _IETagNameGetCollection($o_Obj, $s_TagName)
    If @error Then Return SetError(1)
    For $o_Tag In $o_Tags
        If $o_Tag.ClassName == $s_ClassName Then
            If $s_Innertext Then
                If $o_Tag.innertext == $s_Innertext Then
                    _IEAction($o_Tag, 'click')
                    If @error Then Return SetError(1)
                    _IELoadWait($o_Obj)
                    If @error Then Return SetError(1)
                    Return SetError(0)
                EndIf
            Else
                _IEAction($o_Tag, 'click')
                If @error Then Return SetError(1)
                _IELoadWait($o_Obj)
                If @error Then Return SetError(1)
                Return SetError(0)
            EndIf
        EndIf
    Next
    Return SetError(2)
 EndFunc   ;==>_IETagClassClick

;MsgBox(0,"",$s_zakaz)
_IETagClassClick($o_object, 'span', 'objectAction', $s_zakaz)

$sHTML=_IEPropertyGet($o_object,"innerhtml")
$hFile = FileOpen(@ScriptDir & "\raw\"&$s_zakaz&".txt", 2)
FileWrite($hFile,$sHTML)
FileFlush ($hFile)
FileClose($hFile)
;Sleep(5000)
_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')


Exit