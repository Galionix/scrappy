#include <IE.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#Include <Array.au3>
#include <StringConstants.au3>

;1-url
;2-delay
;3-

$f_input = FileOpen(@ScriptDir & "\init.txt")

;$f_output = FileOpen(@ScriptDir & "\nums.txt",2)
$s_url=""
$i_delay=0
$s_url=FileReadLine($f_input, 1)
$i_delay=Number( FileReadLine($f_input, 2) )


$oIE = _IEAttach("svparfum.ru | CMS.S3")
If @error = $_IEStatus_NoMatch Then

$oIE = _IECreate($s_url)
Sleep($i_delay)
EndIf
;MsgBox($MB_SYSTEMMODAL, "s_url", $s_url)
;MsgBox($MB_SYSTEMMODAL, "$i_delay", $i_delay)







;_IELinkClickByText($oIE,'23822&nbsp;(277875006)')
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
;_IETagClassClick($oIE, 'span', 'objectAction', '23822 (277875006)')


;'\D\d\d\d\d\d.......\d\d\d\d\d\d\d\d\d'
$sHTML=_IEPropertyGet($oIE,"innerhtml")
;$hFile = FileOpen(@ScriptDir & "\1check.txt", 2)
;FileWrite($hFile,$sHTML)
;FileFlush ($hFile)
;FileClose($hFile)




Local $aArray = 0, _
        $iOffset = 1
$s_tmp=""

$f_output = FileOpen(@ScriptDir & "\nums.txt", 2)


While 1
    $aArray = StringRegExp($sHTML, '\D\d\d\d\d\d.......\d\d\d\d\d\d\d\d\d', $STR_REGEXPARRAYMATCH, $iOffset)
    If @error Then ExitLoop
    $iOffset = @extended
    For $i = 0 To UBound($aArray) - 1
	    $s_tmp=$aArray[$i]

		$s_tmp=StringMid($s_tmp,2,5)&" "&StringMid($s_tmp,13)&')'
		FileWrite($f_output,$s_tmp&@CRLF)
        ;MsgBox($MB_SYSTEMMODAL, "RegExp Test with Option 1 - " & $i, $s_tmp)
    Next
 WEnd

;_FileWriteFromArray($f_output, $aArray)


Exit

;Sleep($i_delay)
;_IETagClassClick($oIE, 'span', 's3-ico winclose-thin')





