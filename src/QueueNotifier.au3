#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_Outfile=QueueNotifier.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include "WinHttp.au3"

Opt("MustDeclareVars", 1)

Const $sIniPath = 'config.ini'

Global $_iRaisID =  IniRead($sIniPath, "General", "RaisID", 35) ;35 - Сызранский МФЦ

Global $_sSmtpServer = IniRead($sIniPath, "General", "SmtpServer", "smtp.mail.ru")
Global $_iPort = IniRead($sIniPath, "General", "Port", 465)

Global $_sFromName = IniRead($sIniPath, "General", "FromName", "noreply")
Global $_sFromAddress = IniRead($sIniPath, "General", "FromAddress", "from@mail.com")

Global $_sToAddress = IniRead($sIniPath, "General", "ToAddress", "to@mail.com")

Global $_sLogin = IniRead($sIniPath, "General", "Login", "login")
Global $_sPassword = IniRead($sIniPath, "General", "Password", "password")

Global $_sAdminMail = IniRead($sIniPath, "General", "AdminMail", "sergey.2bite@gmail.com")

Global $_iCheckTimerMin = IniRead($sIniPath, "Timers", "Delay", 60)+0 ;60минут
Global $_iCheckTimerStart = IniRead($sIniPath, "Timers", "DayStart", 8)+0 ;8 часов
Global $_iCheckTimerEnd = IniRead($sIniPath, "Timers", "DayEnd", 18)+0 ;18 часов

_CheckQueueAndReport() ;для теста отправляем 1 раз при запуске
_Main()

Func _Main()
	Local $iTimer = TimerInit()
	While 1
		If 	(TimerDiff($iTimer) > 60 * 60 * 1000) And _
			(@HOUR >= $_iCheckTimerStart or @HOUR <= $_iCheckTimerEnd) Then

			_CheckQueueAndReport()
			$iTimer = TimerInit()
		EndIf
		Sleep(100)
	WEnd
EndFunc

Func _CheckQueueAndReport()
	Local $aTable = _GetQueueTable()
	If Not IsArray($aTable) Or UBound($aTable) < 1 Then _SendTable($_sToAddress, '', True)

	If UBound($aTable) > 1 Then
		Local $sHtmlTable = _ArrayToHtml2D($aTable, 3)
		_SendTable($_sToAddress, $sHtmlTable)
	EndIf
EndFunc


Func _SendTable($sToAddress, $sHtmlTable = '', $bError = False)
	Local $sSubject = "Уведомление об электронной записи"

	Local $asBody[4]
	If $bError Then
		$asBody[0] = 'При попытке получить список заявок с сайта mfc63.ru произошла непридвиденная ошибка. <br />'
	Else
		$asBody[0] = 'На сайте mfc63.ru есть необработанные заявки на прием. <br /><br />' & $sHtmlTable
	EndIf
	$asBody[1] = 'Для редактрирования/подтверждения/отмены заявок перейдите по адресу: http://mfc63.samregion.ru/user <br /><br />'
	$asBody[2] = '--  <br />'
	$asBody[3] = '<i> Это письмо сформировано автоматически. Пожалуйста, не отвечайте на него. <br /> Если у Вас есть вопросы, Вы можете обратиться по электронной почте ' & $_sAdminMail &' <i> '
	Local $sBody = _ArrayToString($asBody, @CRLF)
	Local $sTempFile = _TempFile()
	FileWrite($sTempFile, $sBody)

	Local $sArgs = 	" -t " & $sToAddress & ' -sub "' & $sSubject & '"' & _
					' -attach "' & $sTempFile & '",text/html,i' & _
					" -smtp " & $_sSmtpServer & " -port " & $_iPort & " " & _
					" -f " & $_sFromAddress & " -name " & $_sFromName & _
					" -ssl -auth-login -user " & $_sLogin & " -pass " & $_sPassword & " -q"
	Run("mailsend.exe" & $sArgs, @ScriptDir, @SW_HIDE)
EndFunc

Func _GetQueueTable()
	Local $sHtml = _GetHtmlData()
	If @error  Then Return SetError(1)

	Local $aTable = _ParseTable($sHtml)
	If @error  Then Return SetError(2)

	Return $aTable
EndFunc

Func _ParseTable($sHtml)

	Local $iWantedTable = 1
	Local $aTable = ExtractTable($sHtml, $iWantedTable)
	If @error  Then Return SetError(1)

	Return $aTable
EndFunc

Func _ArrayToHtml2D(Const ByRef $avArray, $iColumnLimit = 0, $attrib = 'border="1"', $iStart = 0, $iEnd = 0)
	 If Not IsArray($avArray) Then Return SetError(1, 0, "")
	 If UBound($avArray, 0) <> 2 Then Return SetError(2, 0, "")

	 Local $sResult, $iUBound = UBound($avArray) - 1
	 Local $row, $sDelimCol = "</td>" & @CRLF, $sDelimRow = '</tr>' & @CRLF

	 ; Bounds checking
	 If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
	 If $iStart < 0 Then $iStart = 0
	 If $iStart > $iEnd Then Return SetError(3, 0, "")

	 $sResult = '<table ' & $attrib & '>' & @CRLF

	 ; Combine
	 For $i = $iStart To $iEnd ; rows
		$row = '<tr>' & @CRLF
		For $j = 0 To UBound($avArray,2) - 1 - $iColumnLimit ; columns
			$row &= '<td>' & $avArray[$i][$j] & $sDelimCol
		Next
		$sResult &= $row & $sDelimRow
	 Next

	 Return $sResult & '</table>' & @CRLF
EndFunc

Func _GetHtmlData()

	; Open needed handles
	Local $hOpen = _WinHttpOpen()
	Local $hConnect = _WinHttpConnect($hOpen, "mfc63.samregion.ru")
	; Specify the reguest:
	Local $hRequest = _WinHttpOpenRequest($hConnect, Default, "my/enroll_requests_not_accepted.php?id=0&rais=" & $_iRaisID)

	; Send request
	_WinHttpSendRequest($hRequest)

	; Wait for the response
	_WinHttpReceiveResponse($hRequest)

	Local $sHeader = _WinHttpQueryHeaders($hRequest) ; ...get full header
	Local $sData = _WinHttpReadData($hRequest)

	; Clean
	_WinHttpCloseHandle($hRequest)
	_WinHttpCloseHandle($hConnect)
	_WinHttpCloseHandle($hOpen)

	$sData = _ANSIToUTF8($sData)

	;~ Local $sFile = _CreateHTML($sData)
	;~ ConsoleWrite($sFile &@CRLF)

	Return $sData
EndFunc

Func ExtractTable($sHtml, $iWantedTable, $bExtractTH = True, $bFillSpan = False)
    ;
    ; $sHtml:           the raw HTML of that contains the table(s)
    ; $iWantedTable:    the nr. of the table to extract from the HTML
    ; $bExtractTH:      if Headers should be extracted as well
    ; $bFillSpaw:       if the whole spaw areas should be filled  (not yet implemented)
    ;
    ; This will find all tables in the HTML page (even if nested)
    Local $aTables = ParseTags($sHtml, "<table", "</table>")
    If @error Then Return SetError(@error, 0, "")
    If $iWantedTable > $aTables[0] Then Return SetError(4, $aTables[0], "") ; in @extended nr. of tables in this HTML
    ; _ArrayDisplay($aTables, "ParseTags <table ... </table>")
    ;
    If $bExtractTH Then ;extract also TableHeaders as normal data?
        $aTables[$iWantedTable] = StringReplace(StringReplace($aTables[$iWantedTable], "<th", "<td"), "</th>", "</td>") ; th becomes td
    EndIf
    ;
    ; rows of the wanted table
    Local $aRows = ParseTags($aTables[$iWantedTable], "<tr", "</tr>") ; $aRows[0] = nr. of rows
    If @error Then Return SetError(@error, 0, "")

    Local $aCols[$aRows[0] + 1], $aTemp
    For $i = 1 To $aRows[0]
        $aTemp = ParseTags($aRows[$i], "<td", "</td>")
        If $aCols[0] < $aTemp[0] Then $aCols[0] = $aTemp[0] ; $aTemp[0] = max nr. of columns in table
        $aCols[$i] = $aTemp
    Next

    Local $aResult[$aRows[0]][$aCols[0]], $iStart, $iEnd, $aRowspan, $aColspan, $iSpanY, $iSpanX, $iSpanRow, $iSpanCol, $iMarkerCode, $iY, $iX, $iX1, $iSpawnCol ;  = 1
    Local $aMirror = $aResult
    For $i = 1 To $aRows[0] ;      scan all rows in this table
        $aTemp = $aCols[$i] ; <td ..> xx </td> .....
        For $ii = 1 To $aTemp[0] ; scan all cells in this row
            $iSpanY = 0
            $iSpanX = 0
            $iY = $i - 1 ; zero base index for vertical ref
            $iX = $ii - 1 ; zero based indexes for horizontal ref
            $aRowspan = StringRegExp($aTemp[$ii], "(?i)rowspan\s*=\s*[""']?\s*(\d+)", 1) ; check presence of rowspan
            If IsArray($aRowspan) Then
                $iSpanY = $aRowspan[0] - 1
                If $iSpanY + $iY > $aRows[0] Then
                    $iSpanY -= $iSpanY + $iY - $aRows[0] + 1
                EndIf
            EndIf
            ;
            $aColspan = StringRegExp($aTemp[$ii], "(?i)colspan\s*=\s*[""']?\s*(\d+)", 1) ; check presence of colspan
            If IsArray($aColspan) Then $iSpanX = $aColspan[0] - 1
            ;
            $iMarkerCode += 1 ; code to mark this span area or single cell
            If $iSpanY Or $iSpanX Then
                $iX1 = $iX
                For $iSpY = 0 To $iSpanY
                    For $iSpX = 0 To $iSpanX
                        $iSpanRow = $iY + $iSpY
                        If $iSpanRow > UBound($aMirror, 1) - 1 Then
                            $iSpanRow = UBound($aMirror, 1) - 1
                        EndIf
                        $iSpanCol = $iX1 + $iSpX
                        If $iSpanCol > UBound($aMirror, 2) - 1 Then
                            ReDim $aResult[$aRows[0]][UBound($aResult, 2) + 1]
                            ReDim $aMirror[$aRows[0]][UBound($aMirror, 2) + 1]
                        EndIf
                        ;
                        While $aMirror[$iSpanRow][$iX1 + $iSpX] ; search first free column
                            $iX1 += 1 ; $iSpanCol += 1
                            If $iX1 + $iSpX > UBound($aMirror, 2) - 1 Then
                                ReDim $aResult[$aRows[0]][UBound($aResult, 2) + 1]
                                ReDim $aMirror[$aRows[0]][UBound($aMirror, 2) + 1]
                            EndIf
                        WEnd
                    Next
                Next
            EndIf
            $iX1 = $iX
            For $iSpX = 0 To $iSpanX
                For $iSpY = 0 To $iSpanY
                    $iSpanRow = $iY + $iSpY
                    If $iSpanRow > UBound($aMirror, 1) - 1 Then
                        $iSpanRow = UBound($aMirror, 1) - 1
                    EndIf
                    $iSpawnCol = $iX1 + $iSpX
                    While $aMirror[$iSpanRow][$iX1 + $iSpX]
                        $iX1 += 1
                        If $iX1 + $iSpX > UBound($aMirror, 2) - 1 Then
                            ReDim $aResult[$aRows[0]][$iX1 + $iSpX + 1]
                            ReDim $aMirror[$aRows[0]][$iX1 + $iSpX + 1]
                        EndIf
                    WEnd
                    $aMirror[$iSpanRow][$iX1 + $iSpX] = $iMarkerCode ; 1
                    $aResult[$iY][$iX1] = StringRegExpReplace($aTemp[$ii], '<[^>]+>', "") ; "(?U)\<.*\>", "")
                Next
            Next
        Next
    Next
    ; _ArrayDisplay($aMirror)
    Return $aResult
EndFunc   ;==>ExtractTable
;
; -----------------------------------------------------------------------------------------
; returns an array containing a collection of <tag ...... </tag> lines. one in each element
; even if are nested
; -----------------------------------------------------------------------------------------
Func ParseTags($sHtml, $sOpening, $sClosing) ; example: $sOpening = '<table', $sClosing = '</table>'
    ; it finds how many of such tags are on the HTML page
    StringReplace($sHtml, $sOpening, $sOpening) ; in @xtended nr. of occurences
    Local $iNrOfThisTag = @extended
    ; I assume that opening <tag and closing </tag> tags are balanced (as should be)
    ; (so NO check is made to see if they are actually balanced)
    If $iNrOfThisTag Then ; if there is at least one of this tag
        ; $aThisTagsPositions array will contain the positions of the
        ; starting <tag and ending </tag> tags within the HTML
        Local $aThisTagsPositions[$iNrOfThisTag * 2 + 1][3] ; 1 based (make room for all open and close tags)
        ; 2) find in the HTML the positions of the $sOpening <tag and $sClosing </tag> tags
        For $i = 1 To $iNrOfThisTag
            $aThisTagsPositions[$i][0] = StringInStr($sHtml, $sOpening, 0, $i) ; start position of $i occurrence of <tag opening tag
            $aThisTagsPositions[$i][1] = $sOpening ; it marks which kind of tag is this
            $aThisTagsPositions[$i][2] = $i ; nr of this tag
            $aThisTagsPositions[$iNrOfThisTag + $i][0] = StringInStr($sHtml, $sClosing, 0, $i) + StringLen($sClosing) - 1 ; end position of $i^ occurrence of </tag> closing tag
            $aThisTagsPositions[$iNrOfThisTag + $i][1] = $sClosing ; it marks which kind of tag is this
        Next
        _ArraySort($aThisTagsPositions, 0, 1) ; now all opening and closing tags are in the same sequence as them appears in the HTML
        Local $aStack[UBound($aThisTagsPositions)][2]
        Local $aTags[Ceiling(UBound($aThisTagsPositions) / 2)] ; will contains the collection of <tag ..... </tag> from the html
        For $i = 1 To UBound($aThisTagsPositions) - 1
            If $aThisTagsPositions[$i][1] = $sOpening Then ; opening <tag
                $aStack[0][0] += 1 ; nr of tags in html
                $aStack[$aStack[0][0]][0] = $sOpening
                $aStack[$aStack[0][0]][1] = $i
            ElseIf $aThisTagsPositions[$i][1] = $sClosing Then ; a closing </tag> was found
                If Not $aStack[0][0] Or Not ($aStack[$aStack[0][0]][0] = $sOpening And $aThisTagsPositions[$i][1] = $sClosing) Then
                    Return SetError(3, 0, 0) ; Open/Close mismatch error
                Else ; pair detected (the reciprocal tag)
                    ; now get coordinates of the 2 tags
                    ; 1) extract this tag <tag ..... </tag> from the html to the array
                    $aTags[$aThisTagsPositions[$aStack[$aStack[0][0]][1]][2]] = StringMid($sHtml, $aThisTagsPositions[$aStack[$aStack[0][0]][1]][0], 1 + $aThisTagsPositions[$i][0] - $aThisTagsPositions[$aStack[$aStack[0][0]][1]][0])
                    ; 2) remove that tag <tag ..... </tag> from the html
                    $sHtml = StringLeft($sHtml, $aThisTagsPositions[$aStack[$aStack[0][0]][1]][0] - 1) & StringMid($sHtml, $aThisTagsPositions[$i][0] + 1)
                    ; 3) adjust the references to the new positions of remaining tags
                    For $ii = $i To UBound($aThisTagsPositions) - 1
                        $aThisTagsPositions[$ii][0] -= StringLen($aTags[$aThisTagsPositions[$aStack[$aStack[0][0]][1]][2]])
                    Next
                    $aStack[0][0] -= 1 ; nr of tags still in html
                EndIf
            EndIf
        Next
        If Not $aStack[0][0] Then ; all tags has been parsed correctly
            $aTags[0] = $iNrOfThisTag
            Return $aTags ; OK
        Else
            Return SetError(2, 0, 0) ; opening and closing tags are not balanced
        EndIf
    Else
        Return SetError(1, 0, 0) ; there are no of such tags on this HTML page
    EndIf
EndFunc   ;==>ParseTags

Func _ANSIToUTF8($sString)
	Return BinaryToString(StringToBinary($sString), 4)
EndFunc

Func _CreateHTML($sData)
	Local $sTempFile = _TempFile(Default, Default, ".html")
	FileWrite($sTempFile, $sData)
	Return $sTempFile
EndFunc