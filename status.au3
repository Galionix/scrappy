#include <IE.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#Include <Array.au3>


    ; Open the file for reading and store the handle to a variable.
    Local $f_input = FileOpen(@ScriptDir & '\work.txt', $FO_READ)
	    Local $f_comment = FileOpen(@ScriptDir & '\comment.txt', $FO_READ)
		    Local $f_dop = FileOpen(@ScriptDir & '\dop.txt', $FO_READ)
    If $f_input = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when reading the file.")
        Return False
    EndIf

;$f_input = FileRead(@ScriptDir & '\work.txt')
$s_status=FileReadLine($f_input, 2)
$s_zakaz=FileReadLine($f_input, 1)
FileClose($s_zakaz)

$s_comment=""
For $i = 1 to _FileCountLines($f_comment)
$s_comment&=FileReadLine($f_comment, $i)&@CRLF
Next
FileClose($f_comment)

$s_dop=""
For $i = 1 to _FileCountLines($f_dop)
$s_dop&=FileReadLine($f_dop, $i)&@CRLF
Next
FileClose($f_dop)


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


 Func _IETagClassClickN($o_Obj, $s_TagName, $s_ClassName, $i_count=0)
    Local $o_Tag
    If Not IsObj($o_Obj) Then Return SetError(1)
    If (Not $s_TagName Or Not $s_ClassName) Then Return SetError(1)
    $o_Tag = _IETagNameGetCollection($o_Obj, $s_TagName,$i_count)
    If @error Then Return SetError(1)

        If $o_Tag.ClassName == $s_ClassName Then
			   MsgBox(0,"","ya")
                _IEAction($o_Tag, 'click')
                If @error Then Return SetError(1)
                _IELoadWait($o_Obj)
                If @error Then Return SetError(1)
                Return SetError(0)

        EndIf

    Return SetError(2)
 EndFunc   ;==>_IETagClassClick
_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')
_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')
_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')
;MsgBox(0,"",$s_zakaz)
_IETagClassClick($o_object, 'span', 'objectAction', $s_zakaz)

_IETagClassClick($o_object, 'button', 's3-btn without-margin w-ico-first', "Изменить")

$o_Comment=_IEGetObjByName($o_object, "comment")
_IEFormElementSetValue($o_Comment,$s_comment)

$o_dop=_IEGetObjByName($o_object, "note")
_IEFormElementSetValue($o_dop,$s_dop)


;Sleep(5000)
_IETagClassClick($o_object, 'span', 's3-btn s3-btn-save')

_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')
    $o_Tags = _IETagNameGetCollection($o_object, "select")
    For $o_Tag In $o_Tags
        If $o_Tag.ClassName == "s3-select white without-search order_status_style chzn-done" Then



					 ;MsgBox(0,"",$s_zakaz)
					_IEFormElementOptionselect ($o_Tag, $s_status, 1, "byText",1)
                    _IELoadWait($o_object)
					 _IETagClassClick($o_object, 'button', 's3-btn without-margin green',"Сохранить изменения")

			 EndIf
    Next


;$o_stat=_IEGetObjByName($o_object,"status_id")
;_IEFormElementOptionselect ($o_stat, $s_status, 1, "byText",1)


$sHTML=_IEPropertyGet($o_object,"innerhtml")
$hFile = FileOpen(@ScriptDir & "\raw\"&$s_zakaz&".txt", 2)
FileWrite($hFile,$sHTML)
FileFlush ($hFile)
FileClose($hFile)


_IETagClassClick($o_object, 'span', 's3-ico winclose-thin')


   $f_output = FileOpen(@ScriptDir & "\"&$s_zakaz&".txt", 2)
   FileWriteLine($f_output,"complete")
Exit