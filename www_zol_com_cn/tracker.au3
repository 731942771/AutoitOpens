#Region ;**** �� AccAu3Wrapper_GUI ����ָ�� ****
#AccAu3Wrapper_Icon=favicon.ico
#AccAu3Wrapper_OutFile=�йش�����-ͶӰ�����ݲɼ�.exe
#AccAu3Wrapper_Compression=4
#AccAu3Wrapper_Res_Comment=�����鲩��
#AccAu3Wrapper_Res_Description=www.cuiweiyou.com
#AccAu3Wrapper_Res_Fileversion=8.8.8.8
#AccAu3Wrapper_Res_ProductVersion=9.9.9.9
#AccAu3Wrapper_Res_LegalCopyright=vigiles
#AccAu3Wrapper_Res_Language=2052
#AccAu3Wrapper_Res_requestedExecutionLevel=None
#AccAu3Wrapper_Res_Field=OriginalFilename|�лս���-��ά��
#AccAu3Wrapper_Res_Field=ProductName|�лս���-��ά��
#AccAu3Wrapper_Res_Field=ProductVersion|V1.0
#AccAu3Wrapper_Res_Field=InternalName|�лս���-��ά��
#AccAu3Wrapper_Res_Field=FileDescription|�лս���-��ά��
#AccAu3Wrapper_Res_Field=Comments|�лս���-��ά��
#AccAu3Wrapper_Res_Field=LegalTrademarks|cuiweiyou.com
#AccAu3Wrapper_Res_Field=CompanyName|cuiweiyou.com
#AccAu3Wrapper_Run_AU3Check=n
#Tidy_Parameters=/sfc/rel
#AccAu3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** �� AccAu3Wrapper_GUI ����ָ�� ****

;_ExcelBookNew($fVisible)
	;�����µĹ��������ض���ʵ�����Ų�ִ�����Ƿ���ʾExcel���� 0=��ִ̨�в���ʾ, 1=��ʾ��
;_ExcelBookOpen($sFilePath[, $fVisible = 1[, $fReadOnly = False[, $sPassword = ""[, $sWritePassword = "")
	;��һ�����еģ��Ѵ��ڵģ�������������������ʶ�����ļ�·��������0/��ʾ1-Ĭ�� ������ֻ��-FalseĬ�ϣ��Ķ�����-noneĬ�ϣ��޸�����-noneĬ�ϣ�
;_ExcelBookAttach
;_ExcelBookSave
;_ExcelBookSaveAs
;_ExcelBookClose
;_ExcelWriteCell
;_ExcelWriteFormula
;_ExcelWriteArray
;_ExcelWriteSheetFromArray
;_ExcelHyperlinkInsert
;_ExcelNumberFormat
;_ExcelReadCell
;_ExcelReadArray
;_ExcelReadSheetToArray
;_ExcelRowDelete
;_ExcelColumnDelete
;_ExcelRowInsert
;_ExcelColumnInsert
;_ExcelSheetAddNew($oExcel[, $sName = ""])
;	�����µĹ�����(_ExcelBookNew()�򿪵�һ�� Excel object������������)
;_ExcelSheetDelete($oExcel, $vSheet, $fAlerts = False)
;	ɾ��ָ���Ĺ�����(_ExcelBookNew�õ����ļ������������ƣ��Ƿ񵯳�����-Ĭ��False)
;_ExcelSheetNameGet
;_ExcelSheetNameSet
;_ExcelSheetList
;	�õ�ȫ���Ĺ��������Ƶ�һ������
;_ExcelSheetActivate
;_ExcelSheetMove
;_ExcelHorizontalAlignSet
;_ExcelFontSetProperties
;_ExcelNumberFormat

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#include-once

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Excel2.au3> ;�Լ��޸Ĺ��ĵ��ӱ��
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

;------------------------------------------------------------------------------------------------
Opt("GUIOnEventMode", 1) ; �л��� OnEvent ģʽ

;------------------------------------------ Ԫ���� ----------------------------------------------
Global $btnSheets, $btnRows, $btnCollections, $btnCollectionsToXls, $progressState
Global $sRootPath = @ScriptDir & "\temp", $sFileBrands = @ScriptDir & "\Barand.txt", $sFileBrandArgs = @ScriptDir & "\BarandShowArg.txt", $sFileCollectArgs = @ScriptDir & "\Collection.txt", $sXlsPath = @DesktopDir & "\Results.xls"
If FileExists ( $sRootPath ) = 0 Then
	DirCreate ( $sRootPath )
EndIf
Global $nTmpStepProess = 0, $nTmpCountProess = 0	; ģ����߳��õ�����ֵ�������������½�����

;------------------------------------------ ��ʼ�� ----------------------------------------------
Local $hGUI = GUICreate("�лս���|������|�йش�����-ͶӰ�����ݲɼ�ϵͳ", 600, 300)
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE")

	Local $btnBrands = GUICtrlCreateButton("Ʒ���ܼ�", 5, 5, 125, 30)
		GUICtrlSetOnEvent($btnBrands, "_ShowBrands")
	
	$btnSheets = GUICtrlCreateButton("    1.����Ʒ�ƴ���������", 135, 5, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnSheets, "_CreateSheets")
		
	
	Local $btnRowArgs = GUICtrlCreateButton("չʾ����", 5, 40, 125, 30)
		GUICtrlSetOnEvent($btnRowArgs, "_ShowArgs")

	$btnRows = GUICtrlCreateButton("    2.��дƷ�Ʋ�����������", 135, 40, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnRows, "_CreateRows")
		
	
	Local $btnCollectArgs = GUICtrlCreateButton("�ɼ�����", 5, 75, 125, 30)
		GUICtrlSetOnEvent($btnCollectArgs, "_ShowCollections")

	$btnCollections = GUICtrlCreateButton("    3.�ɼ�Ʒ�����ݲ�����", 135, 75, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnCollections, "_CreateCollections")
		

	Local $btnCollectArgs = GUICtrlCreateButton("�鿴������", 5, 110, 125, 30)
		GUICtrlSetOnEvent($btnCollectArgs, "_ShowCache")

	$btnCollectionsToXls = GUICtrlCreateButton("    4.���滺�����ݵ�xls�ļ�", 135, 110, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnCollectionsToXls, "_WriteCaches")
		
	$progressState = GUICtrlCreateProgress ( 5, 150, 330, 10 )
		GUICtrlSetState($progressState, $GUI_HIDE )

GUISetState(@SW_SHOW, $hGUI)

GUICtrlSetState($btnRows, $GUI_DISABLE)
GUICtrlSetState($btnCollections, $GUI_DISABLE)
GUICtrlSetState($btnCollectionsToXls, $GUI_DISABLE)

;------------------------------------------ ����ά�� ---------------------------------------------
While 1
    Sleep(200)
WEnd

Func GUI_EVENT_CLOSE()
    Exit
EndFunc

;----------------------------------------- 1.���������� --------------------------------------------
Func _CreateSheets()
	
	GUICtrlSetState($btnSheets, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	Local $arrayBrand ; Ʒ�Ƽ��ϡ���ʽ ������ ��ÿ��һ��Ʒ������
	_FileReadToArray ( $sFileBrands, $arrayBrand) ; ��ȡ"Ʒ��"�ļ����ݵ�����
	_ArrayDelete ( $arrayBrand, 0 ) ; ɾ�������ָ��Ԫ�أ��Թ���������ʹ�á�0����������ЧԪ�ظ���
	_ArraySort ( $arrayBrand, 1 ) ; ʹ�� dualpivotsort(˫֧������)/quicksort(��������)/insertionsort(��������) �㷨��һ��һά���߶�ά������������
	
	$nTmpCountProess = UBound($arrayBrand)

	Local $oExcel = _ExcelBookNew(0) ; �����¹�����������������ʶ��

	For $i = 0 To UBound($arrayBrand) - 1
		_ExcelSheetAddNew($oExcel, $arrayBrand[$i]) ; ����µĹ������ļ�, ���������ǵ����ơ�������ǰ
		
		$nTmpStepProess = $i
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next

	_ExcelSheetDelete($oExcel, "Sheet1")	; ɾ��Ĭ�ϴ����Ĺ�����
	_ExcelSheetDelete($oExcel, "Sheet2")
	_ExcelSheetDelete($oExcel, "Sheet3")

	_ExcelBookSaveAs($oExcel, $sXlsPath, "xls", 0, 1)        ; ����-���Ϊ ������

	_ExcelBookClose($oExcel) ; ���, �ر��ļ�
	
	GUICtrlSetState($btnRows, $GUI_ENABLE)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 2.��д������ --------------------------------------------
Func _CreateRows()
	GUICtrlSetState($btnRows, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	Local $arrBrand ; Ʒ�Ƽ���
	_FileReadToArray ( $sFileBrands, $arrBrand) ; ��ȡ�ļ����ݵ�����
	_ArrayDelete ( $arrBrand, 0 ) ; ɾ�������ָ��Ԫ�أ��Թ���������ʹ��
	_ArraySort ( $arrBrand, 1 ) ; quicksort(��������)
	
	$nTmpCountProess = UBound($arrBrand)
	
	Local $arrBrandArg ; Ʒ�Ʋ������� ����ʽ ��ʾоƬ ��ÿ��һ����������
	_FileReadToArray ( $sFileBrandArgs, $arrBrandArg)
	_ArrayDelete ( $arrBrandArg, 0 )
	
	Local $oExcel = _ExcelBookOpen($sXlsPath, 0) ; ��һ�����ڵ�xls�ļ�

	For $i = 0 To UBound($arrBrand) - 1
		_ExcelSheetActivate($oExcel, $arrBrand[$i]) ; ѡ��ĳ��������
		_ExcelWriteArray($oExcel, 1, 1, $arrBrandArg, 0) ; д��һ��(�ļ����У��У����飬���� 0-ˮƽ-��/1-��ֱ-��)
		
		;�������ԡ���ָ����������Ч���ļ���  ��ʼ�У���ʼ�У������У������У�   ���壬     ���壬 б�壬  �ֺţ��»��ߣ���ɫ���иߣ��п�,   ���룩
		_ExcelFontSetProperties($oExcel, 1,    1,    1,   50, "΢���ź�", True, False, 13,  False,  5,  18,  10,  "center")
		
		$nTmpStepProess = $i
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next

	_ExcelBookClose($oExcel)
	
	GUICtrlSetState($btnCollections, $GUI_ENABLE)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 3.��д�ɼ����� --------------------------------------------
Func _CreateCollections()
	GUICtrlSetState($btnCollections, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	;;;---------------------------------------------- һ���ɼ�����
	
	Local $aBrandArgArr ; Ʒ�Ʋ�������
	_FileReadToArray ( $sFileBrandArgs, $aBrandArgArr)
	_ArrayDelete ( $aBrandArgArr, 0 )
	
	Local $aCollectArr	; Ʒ�� �ɼ��������ϡ�������ʽ http://detail.zol.com.cn/projector/benq/|���� ��ÿ��һ������
	_FileReadToArray ( $sFileCollectArgs, $aCollectArr)
	_ArrayDelete ( $aCollectArr, 0 )
	
	$nTmpCountProess = UBound($aCollectArr)
	
	Local $xmlhttp = ObjCreate("MSXML2.XMLHTTP.3.0")
	
	For $i = 0 To UBound($aCollectArr) - 1
		Local $aCollectUrlAndName = StringSplit($aCollectArr[$i], "|") ; ��ֳ�һһ��Ӧ�� URL��Ʒ�� ��0����������ЧԪ������
		ConsoleWrite("A,Ʒ�ƺ����ӣ�" & $aCollectUrlAndName[1] & " - " & $aCollectUrlAndName[2] & @CRLF)
		
		$xmlhttp.open("GET", $aCollectUrlAndName[1], False)	                                 ; 1.��URL���õ�Ʒ�Ƶ�ȫ���ͺ��б�ҳ��
		$xmlhttp.send()
		
		Local $body = BinaryToString($xmlhttp.responseBody, 1)			                     ; 2.��������ҳ��

		Local $more = StringRegExp($body, 'a class="more" href="(.*?)" target="_blank"', 3)	 ; 3.��ȡ ȫ���ͺŵ�"�������"ҳ��URL
		ConsoleWrite(@TAB & "B,��ϸҳ��������" & UBound($more) - 1 & @CRLF)
		
		For $j = 0 To UBound($more) - 1	                                                     ; 4.���� ÿ���ͺŵ���ϸҳ��URL
			Local $morep = "http://detail.zol.com.cn" & $more[$j] ; ƴ��·��
			ConsoleWrite(@TAB & @TAB & "C,��ϸҳ�棺" & $morep & @CRLF)
			
			$xmlhttp.open("GET", $morep, False)	                                             ; 5.���뵥���ͺ�ҳ��
			$xmlhttp.send()
			
			$body = BinaryToString($xmlhttp.responseBody, 1)
			$body = StringStripCR($body)
			$body = StringStripWS($body, 8)
			
			Local $title = StringRegExp($body, 'varproName="(.*?)";', 3)                              ; 1> ��ҳ������ȡ���ͺš�
			If IsArray($title) Then
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], "�ͺ�", $title[0] )    ; 2> ���浽ini�ļ����ļ�����̬��һ��Ʒ�ƶ�Ӧһ���ļ�
				ConsoleWrite(@TAB & @TAB & @TAB & "D���ͺţ�" & $title[0] & @CRLF)
			Else
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], "�ͺ�", "" )
				ConsoleWrite(@TAB & @TAB & @TAB & "D���ͺţ�" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
			EndIf
			
			Local $target = StringRegExp($body, '<bclass="price-type">(.*?)<', 3)                     ; 3> ��ҳ������ȡ���۸�
			If IsArray($target) Then
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], "�ο���", $target[0] ) ; ���ini�ļ��������򴴽����Ѵ������д
				ConsoleWrite(@TAB & @TAB & @TAB & "D���ο��ۣ�" & $target[0] & @CRLF)
			Else
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], "�ο���", "" )
				ConsoleWrite(@TAB & @TAB & @TAB & "D���ο��ۣ�" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
			EndIf
			
			;-----------------------------�����ɼ�
			For $k = 2 To UBound($aBrandArgArr)-1                                                     ; 4> ����Ĳ�������ҳ�л�����ʽ��ͬ��������ȡ
				;$target = StringRegExp($body, $aArgArray[$k] & "<[\s\S]+?>([^<]+?)</span", 3)
				$target = StringRegExp($body, '(?s)' & $aBrandArgArr[$k] & '<.+?">(.+?)</', 1) ; ����ƥ��ÿ���ɼ�����~~ ������������������� afan �ṩ
				If IsArray($target) Then
					Local $str = StringRegExpReplace($target[0], '<br/>', '|') ; �����滻�����ַ�
					$str = StringRegExpReplace($str, '<.*?>', '')
					$str = StringReplace($str, "&nbsp;", "")
					IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], $aBrandArgArr[$k], $str )
					;ConsoleWrite(@TAB & @TAB & @TAB & "D��" & $aArgArray[$k] & "��" & $str & @CRLF)
				Else
					IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "�ɼ�����.ini", $title[0], $aBrandArgArr[$k], "" )
					;ConsoleWrite(@TAB & @TAB & @TAB & "D��" & $aArgArray[$k] & "��" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
				EndIf
			Next
		Next
		
		
		$nTmpStepProess = $i
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next
	
	$xmlhttp.abort()
	
	GUICtrlSetState($btnCollectionsToXls, $GUI_ENABLE)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 4.���滺������-�ɼ���� --------------------------------------------
Func _WriteCaches()
	GUICtrlSetState($btnCollectionsToXls, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	;;;---------------------------------------------- �����洢����
	; �Ȼ��浽 ini ��Ե�ʣ�
	; ֱ�Ӵ���ҳ�ɼ�����д�� xls �Ļ���
	; ����Ҫȷ����������Ҫƥ���У�ƥ����ʱֻ�ܸ����е�����λ��
	; ���������硢�ļ�ռ�õ��ٶȻ�Ƚ���
	; ���˸о��ɼ���ini�����һ��ִ��д��xls�������ɿ��Ժã�Ҳ��ȽϿ� :)
	
	Local $arrBrand ; Ʒ�Ƽ���
	_FileReadToArray ( $sFileBrands, $arrBrand) ; ��ȡ�ļ����ݵ�����
	_ArrayDelete ( $arrBrand, 0 ) ; ɾ�������ָ��Ԫ�أ��Թ���������ʹ��
	_ArraySort ( $arrBrand, 1 ) ; quicksort(��������)
	
	$nTmpCountProess = UBound($arrBrand)
	
	Local $oExcel = _ExcelBookOpen($sXlsPath, 0) ; ��һ�����ڵ�xls�ļ�
	
	For $m = 0 To UBound($arrBrand) - 1 ; ����Ʒ��
		_ExcelSheetActivate( $oExcel, $arrBrand[$m] )                                                    ; 1.����Ʒ�ƣ�ѡ��ĳ��������
		ConsoleWrite("A��Ʒ�ƣ�" & $arrBrand[$m] & @CRLF)
		
		Local $arrModelNams = IniReadSectionNames( $sRootPath & "\" & $arrBrand[$m] & "�ɼ�����.ini" )   ; 2.��ȡһ��Ʒ��ȫ�����ͺ��ֶ�
		If @error = 0 Then
			ConsoleWrite(@TAB & "B���ͺ�������" & $arrModelNams[0] & @CRLF)
			
			For $n = 1 To $arrModelNams[0] ; �����ͺ��ֶΡ�0�����������Ԫ����������ЧԪ�ش�1��ʼ
				Local $arrModelKV = IniReadSection ( $sRootPath & "\" & $arrBrand[$m] & "�ɼ�����.ini", $arrModelNams[$n] ) ; 3.��ȡһ���ֶ��£�ȫ���ļ�ֵ�ԣ�����=���ݣ���[0][0]Ϊ��Ч��ֵ������
				If @error = 0 Then
					;ConsoleWrite(@TAB & @TAB & "C������������" & $arrModelKV[0][0] & @CRLF)
					
					Local $aCacheArray[ $arrModelKV[0][0] ] ; �������飬����Ϊ��Ч��ֵ������
					For $p = 1 To $arrModelKV[0][0]     ; ����ÿ�� ��ֵ��
						$aCacheArray[$p - 1] = $arrModelKV[$p][1]                                         ; 4.��ȡһһ��Ӧ�Ĳɼ�������浽���顣[n][0]Ϊ�ؼ��֣�[n][1]Ϊֵ
						;ConsoleWrite(@TAB & @TAB & @TAB & "D��������" & $arrModelKV[$p][0] & "=" & $arrModelKV[$p][1] & @CRLF)
					Next
					
					_ExcelWriteArray($oExcel, $n + 1, 1, $aCacheArray)                                    ; 5.дһά���鵽xls��һ��
				EndIf
			Next
		EndIf
		
		$nTmpStepProess = $m
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next
	
	For $q = 0 To UBound($arrBrand) - 1 ; ����Ʒ��
		$oExcel.worksheets($q + 1).Cells.EntireColumn.AutoFit()
	Next

	_ExcelBookClose($oExcel)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- �鿴ȫ��Ʒ�� --------------------------------------------
Func _ShowBrands()

	Local $aRetArray
	_FileReadToArray ( $sFileBrands, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "ͶӰ��Ʒ��", Default, 32 )

EndFunc

;----------------------------------------- �鿴Ʒ�Ʋ��� --------------------------------------------
Func _ShowArgs()

	Local $aRetArray
	_FileReadToArray ( $sFileBrandArgs, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "Ʒ�Ʋ���", Default, 32 )

EndFunc

;----------------------------------------- �鿴�ɼ����� --------------------------------------------
Func _ShowCollections()

	Local $aRetArray
	_FileReadToArray ( $sFileCollectArgs, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "�ɼ�����", Default, 32 )

EndFunc

;----------------------------------------- �鿴�ɼ���� --------------------------------------------
Func _ShowCache()
	ShellExecute('explorer', $sRootPath )
EndFunc
