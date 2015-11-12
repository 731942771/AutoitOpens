#Region ;**** 由 AccAu3Wrapper_GUI 创建指令 ****
#AccAu3Wrapper_Icon=favicon.ico
#AccAu3Wrapper_OutFile=中关村在线-投影机数据采集.exe
#AccAu3Wrapper_Compression=4
#AccAu3Wrapper_Res_Comment=威格灵博客
#AccAu3Wrapper_Res_Description=www.cuiweiyou.com
#AccAu3Wrapper_Res_Fileversion=8.8.8.8
#AccAu3Wrapper_Res_ProductVersion=9.9.9.9
#AccAu3Wrapper_Res_LegalCopyright=vigiles
#AccAu3Wrapper_Res_Language=2052
#AccAu3Wrapper_Res_requestedExecutionLevel=None
#AccAu3Wrapper_Res_Field=OriginalFilename|崔维友
#AccAu3Wrapper_Res_Field=ProductName|崔维友
#AccAu3Wrapper_Res_Field=ProductVersion|V1.0
#AccAu3Wrapper_Res_Field=InternalName|崔维友
#AccAu3Wrapper_Res_Field=FileDescription|崔维友
#AccAu3Wrapper_Res_Field=Comments|崔维友
#AccAu3Wrapper_Res_Field=LegalTrademarks|cuiweiyou.com
#AccAu3Wrapper_Res_Field=CompanyName|cuiweiyou.com
#AccAu3Wrapper_Run_AU3Check=n
#Tidy_Parameters=/sfc/rel
#AccAu3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** 由 AccAu3Wrapper_GUI 创建指令 ****

;_ExcelBookNew($fVisible)
	;创建新的工作表并返回对象实例（脚步执行中是否显示Excel程序 0=后台执行不显示, 1=显示）
;_ExcelBookOpen($sFilePath[, $fVisible = 1[, $fReadOnly = False[, $sPassword = ""[, $sWritePassword = "")
	;打开一个现有的（已存在的）工作簿并返回其对象标识符（文件路径，隐藏0/显示1-默认 操作，只读-False默认，阅读密码-none默认，修改密码-none默认）
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
;	创建新的工作表(_ExcelBookNew()打开的一个 Excel object，工作表名称)
;_ExcelSheetDelete($oExcel, $vSheet, $fAlerts = False)
;	删除指定的工作表(_ExcelBookNew得到的文件，工作表名称，是否弹出警告-默认False)
;_ExcelSheetNameGet
;_ExcelSheetNameSet
;_ExcelSheetList
;	得到全部的工作表名称到一个数组
;_ExcelSheetActivate
;_ExcelSheetMove
;_ExcelHorizontalAlignSet
;_ExcelFontSetProperties
;_ExcelNumberFormat

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#include-once

#include <Array.au3>
#include <ButtonConstants.au3>
#include <Excel2.au3> ;自己修改过的电子表格
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

;------------------------------------------------------------------------------------------------
Opt("GUIOnEventMode", 1) ; 切换到 OnEvent 模式

;------------------------------------------ 元数据 ----------------------------------------------
Global $btnSheets, $btnRows, $btnCollections, $btnCollectionsToXls, $progressState
Global $sRootPath = @ScriptDir & "\temp", $sFileBrands = @ScriptDir & "\Barand.txt", $sFileBrandArgs = @ScriptDir & "\BarandShowArg.txt", $sFileCollectArgs = @ScriptDir & "\Collection.txt", $sXlsPath = @DesktopDir & "\Results.xls"
If FileExists ( $sRootPath ) = 0 Then
	DirCreate ( $sRootPath )
EndIf
Global $nTmpStepProess = 0, $nTmpCountProess = 0	; 模拟多线程用到的数值变量，用来更新进度条

;------------------------------------------ 初始化 ----------------------------------------------
Local $hGUI = GUICreate("威格灵|中关村在线-投影机数据采集系统", 600, 300)
	GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE")

	Local $btnBrands = GUICtrlCreateButton("品牌总集", 5, 5, 125, 30)
		GUICtrlSetOnEvent($btnBrands, "_ShowBrands")
	
	$btnSheets = GUICtrlCreateButton("    1.根据品牌创建工作表", 135, 5, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnSheets, "_CreateSheets")
		
	
	Local $btnRowArgs = GUICtrlCreateButton("展示参数", 5, 40, 125, 30)
		GUICtrlSetOnEvent($btnRowArgs, "_ShowArgs")

	$btnRows = GUICtrlCreateButton("    2.填写品牌参数到工作表", 135, 40, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnRows, "_CreateRows")
		
	
	Local $btnCollectArgs = GUICtrlCreateButton("采集参数", 5, 75, 125, 30)
		GUICtrlSetOnEvent($btnCollectArgs, "_ShowCollections")

	$btnCollections = GUICtrlCreateButton("    3.采集品牌数据并缓存", 135, 75, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnCollections, "_CreateCollections")
		

	Local $btnCollectArgs = GUICtrlCreateButton("查看缓存结果", 5, 110, 125, 30)
		GUICtrlSetOnEvent($btnCollectArgs, "_ShowCache")

	$btnCollectionsToXls = GUICtrlCreateButton("    4.保存缓存数据到xls文件", 135, 110, 200, 30, $BS_LEFT )
		GUICtrlSetOnEvent($btnCollectionsToXls, "_WriteCaches")
		
	$progressState = GUICtrlCreateProgress ( 5, 150, 330, 10 )
		GUICtrlSetState($progressState, $GUI_HIDE )

GUISetState(@SW_SHOW, $hGUI)

GUICtrlSetState($btnRows, $GUI_DISABLE)
GUICtrlSetState($btnCollections, $GUI_DISABLE)
GUICtrlSetState($btnCollectionsToXls, $GUI_DISABLE)

;------------------------------------------ 窗体维持 ---------------------------------------------
While 1
    Sleep(200)
WEnd

Func GUI_EVENT_CLOSE()
    Exit
EndFunc

;----------------------------------------- 1.创建工作表 --------------------------------------------
Func _CreateSheets()
	
	GUICtrlSetState($btnSheets, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	Local $arrayBrand ; 品牌集合。格式 爱普生 。每行一个品牌名词
	_FileReadToArray ( $sFileBrands, $arrayBrand) ; 读取"品牌"文件内容到数组
	_ArrayDelete ( $arrayBrand, 0 ) ; 删除数组的指定元素，以供下面排序使用。0索引保存有效元素个数
	_ArraySort ( $arrayBrand, 1 ) ; 使用 dualpivotsort(双支点排序)/quicksort(快速排序)/insertionsort(插入排序) 算法对一个一维或者二维数组索引排序
	
	$nTmpCountProess = UBound($arrayBrand)

	Local $oExcel = _ExcelBookNew(0) ; 创建新工作簿并返回其对象标识符

	For $i = 0 To UBound($arrayBrand) - 1
		_ExcelSheetAddNew($oExcel, $arrayBrand[$i]) ; 添加新的工作表到文件, 并设置它们的名称。后来居前
		
		$nTmpStepProess = $i
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next

	_ExcelSheetDelete($oExcel, "Sheet1")	; 删除默认创建的工作表
	_ExcelSheetDelete($oExcel, "Sheet2")
	_ExcelSheetDelete($oExcel, "Sheet3")

	_ExcelBookSaveAs($oExcel, $sXlsPath, "xls", 0, 1)        ; 保存-另存为 到桌面

	_ExcelBookClose($oExcel) ; 最后, 关闭文件
	
	GUICtrlSetState($btnRows, $GUI_ENABLE)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 2.填写工作表 --------------------------------------------
Func _CreateRows()
	GUICtrlSetState($btnRows, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	Local $arrBrand ; 品牌集合
	_FileReadToArray ( $sFileBrands, $arrBrand) ; 读取文件内容到数组
	_ArrayDelete ( $arrBrand, 0 ) ; 删除数组的指定元素，以供下面排序使用
	_ArraySort ( $arrBrand, 1 ) ; quicksort(快速排序)
	
	$nTmpCountProess = UBound($arrBrand)
	
	Local $arrBrandArg ; 品牌参数集合 。格式 显示芯片 。每行一个参数名词
	_FileReadToArray ( $sFileBrandArgs, $arrBrandArg)
	_ArrayDelete ( $arrBrandArg, 0 )
	
	Local $oExcel = _ExcelBookOpen($sXlsPath, 0) ; 打开一个存在的xls文件

	For $i = 0 To UBound($arrBrand) - 1
		_ExcelSheetActivate($oExcel, $arrBrand[$i]) ; 选择某个工作表
		_ExcelWriteArray($oExcel, 1, 1, $arrBrandArg, 0) ; 写入一行(文件，行，列，数组，方向 0-水平-行/1-垂直-列)
		
		;设置属性。对指定的行列生效（文件，  起始行，起始列，结束行，结束列，   字体，     粗体， 斜体，  字号，下划线，颜色，行高，列宽,   对齐）
		_ExcelFontSetProperties($oExcel, 1,    1,    1,   50, "微软雅黑", True, False, 13,  False,  5,  18,  10,  "center")
		
		$nTmpStepProess = $i
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next

	_ExcelBookClose($oExcel)
	
	GUICtrlSetState($btnCollections, $GUI_ENABLE)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 3.填写采集数据 --------------------------------------------
Func _CreateCollections()
	GUICtrlSetState($btnCollections, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	;;;---------------------------------------------- 一。采集过程
	
	Local $aBrandArgArr ; 品牌参数集合
	_FileReadToArray ( $sFileBrandArgs, $aBrandArgArr)
	_ArrayDelete ( $aBrandArgArr, 0 )
	
	Local $aCollectArr	; 品牌 采集参数集合。参数格式 http://detail.zol.com.cn/projector/benq/|明基 ，每行一条参数
	_FileReadToArray ( $sFileCollectArgs, $aCollectArr)
	_ArrayDelete ( $aCollectArr, 0 )
	
	$nTmpCountProess = UBound($aCollectArr)
	
	Local $xmlhttp = ObjCreate("MSXML2.XMLHTTP.3.0")
	
	For $i = 0 To UBound($aCollectArr) - 1
		Local $aCollectUrlAndName = StringSplit($aCollectArr[$i], "|") ; 拆分出一一对应的 URL和品牌 。0索引保存有效元素数量
		ConsoleWrite("A,品牌和链接：" & $aCollectUrlAndName[1] & " - " & $aCollectUrlAndName[2] & @CRLF)
		
		$xmlhttp.open("GET", $aCollectUrlAndName[1], False)	                                 ; 1.打开URL，得到品牌的全部型号列表页面
		$xmlhttp.send()
		
		Local $body = BinaryToString($xmlhttp.responseBody, 1)			                     ; 2.缓存整个页面

		Local $more = StringRegExp($body, 'a class="more" href="(.*?)" target="_blank"', 3)	 ; 3.提取 全部型号的"更多参数"页面URL
		ConsoleWrite(@TAB & "B,详细页面数量：" & UBound($more) - 1 & @CRLF)
		
		For $j = 0 To UBound($more) - 1	                                                     ; 4.遍历 每个型号的详细页面URL
			Local $morep = "http://detail.zol.com.cn" & $more[$j] ; 拼接路径
			ConsoleWrite(@TAB & @TAB & "C,详细页面：" & $morep & @CRLF)
			
			$xmlhttp.open("GET", $morep, False)	                                             ; 5.进入单个型号页面
			$xmlhttp.send()
			
			$body = BinaryToString($xmlhttp.responseBody, 1)
			$body = StringStripCR($body)
			$body = StringStripWS($body, 8)
			
			Local $title = StringRegExp($body, 'varproName="(.*?)";', 3)                              ; 1> 从页面中提取“型号”
			If IsArray($title) Then
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], "型号", $title[0] )    ; 2> 保存到ini文件。文件名动态，一个品牌对应一个文件
				ConsoleWrite(@TAB & @TAB & @TAB & "D，型号：" & $title[0] & @CRLF)
			Else
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], "型号", "" )
				ConsoleWrite(@TAB & @TAB & @TAB & "D，型号：" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
			EndIf
			
			Local $target = StringRegExp($body, '<bclass="price-type">(.*?)<', 3)                     ; 3> 从页面中提取“价格”
			If IsArray($target) Then
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], "参考价", $target[0] ) ; 如果ini文件不存在则创建，已存在则读写
				ConsoleWrite(@TAB & @TAB & @TAB & "D，参考价：" & $target[0] & @CRLF)
			Else
				IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], "参考价", "" )
				ConsoleWrite(@TAB & @TAB & @TAB & "D，参考价：" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
			EndIf
			
			;-----------------------------遍历采集
			For $k = 2 To UBound($aBrandArgArr)-1                                                     ; 4> 后面的参数在网页中基本格式相同，批量提取
				;$target = StringRegExp($body, $aArgArray[$k] & "<[\s\S]+?>([^<]+?)</span", 3)
				$target = StringRegExp($body, '(?s)' & $aBrandArgArr[$k] & '<.+?">(.+?)</', 1) ; 正则匹配每个采集参数~~ 本程序中正则基本都由 afan 提供
				If IsArray($target) Then
					Local $str = StringRegExpReplace($target[0], '<br/>', '|') ; 正则替换无用字符
					$str = StringRegExpReplace($str, '<.*?>', '')
					$str = StringReplace($str, "&nbsp;", "")
					IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], $aBrandArgArr[$k], $str )
					;ConsoleWrite(@TAB & @TAB & @TAB & "D，" & $aArgArray[$k] & "：" & $str & @CRLF)
				Else
					IniWrite( $sRootPath & "\" & $aCollectUrlAndName[2] & "采集参数.ini", $title[0], $aBrandArgArr[$k], "" )
					;ConsoleWrite(@TAB & @TAB & @TAB & "D，" & $aArgArray[$k] & "：" & "~~~~~~~~~~~~~~~~~~~~~" & @CRLF)
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

;----------------------------------------- 4.保存缓存数据-采集结果 --------------------------------------------
Func _WriteCaches()
	GUICtrlSetState($btnCollectionsToXls, $GUI_DISABLE)
	
	GUICtrlSetState($progressState, $GUI_SHOW )
	GUICtrlSetData($progressState, 1)
	
	;;;---------------------------------------------- 二。存储过程
	; 先缓存到 ini 的缘故：
	; 直接从网页采集到后写入 xls 的话，
	; 首先要确定工作表，还要匹配列，匹配列时只能根据列的索引位置
	; 受制于网络、文件占用等速度会比较慢
	; 个人感觉采集到ini，最后一次执行写入xls。这样可控性好，也会比较快 :)
	
	Local $arrBrand ; 品牌集合
	_FileReadToArray ( $sFileBrands, $arrBrand) ; 读取文件内容到数组
	_ArrayDelete ( $arrBrand, 0 ) ; 删除数组的指定元素，以供下面排序使用
	_ArraySort ( $arrBrand, 1 ) ; quicksort(快速排序)
	
	$nTmpCountProess = UBound($arrBrand)
	
	Local $oExcel = _ExcelBookOpen($sXlsPath, 0) ; 打开一个存在的xls文件
	
	For $m = 0 To UBound($arrBrand) - 1 ; 遍历品牌
		_ExcelSheetActivate( $oExcel, $arrBrand[$m] )                                                    ; 1.根据品牌，选择某个工作表
		ConsoleWrite("A。品牌：" & $arrBrand[$m] & @CRLF)
		
		Local $arrModelNams = IniReadSectionNames( $sRootPath & "\" & $arrBrand[$m] & "采集参数.ini" )   ; 2.读取一个品牌全部的型号字段
		If @error = 0 Then
			ConsoleWrite(@TAB & "B。型号数量：" & $arrModelNams[0] & @CRLF)
			
			For $n = 1 To $arrModelNams[0] ; 遍历型号字段。0索引保存可用元素数量。有效元素从1开始
				Local $arrModelKV = IniReadSection ( $sRootPath & "\" & $arrBrand[$m] & "采集参数.ini", $arrModelNams[$n] ) ; 3.读取一个字段下，全部的键值对（参数=数据）。[0][0]为有效键值对数量
				If @error = 0 Then
					;ConsoleWrite(@TAB & @TAB & "C。参数数量：" & $arrModelKV[0][0] & @CRLF)
					
					Local $aCacheArray[ $arrModelKV[0][0] ] ; 缓存数组，长度为有效键值对数量
					For $p = 1 To $arrModelKV[0][0]     ; 遍历每行 键值对
						$aCacheArray[$p - 1] = $arrModelKV[$p][1]                                         ; 4.读取一一对应的采集结果缓存到数组。[n][0]为关键字，[n][1]为值
						;ConsoleWrite(@TAB & @TAB & @TAB & "D。参数：" & $arrModelKV[$p][0] & "=" & $arrModelKV[$p][1] & @CRLF)
					Next
					
					_ExcelWriteArray($oExcel, $n + 1, 1, $aCacheArray)                                    ; 5.写一维数组到xls的一行
				EndIf
			Next
		EndIf
		
		$nTmpStepProess = $m
		GUICtrlSetData($progressState, $nTmpStepProess/$nTmpCountProess*100)
	Next
	
	For $q = 0 To UBound($arrBrand) - 1 ; 遍历品牌
		$oExcel.worksheets($q + 1).Cells.EntireColumn.AutoFit()
	Next

	_ExcelBookClose($oExcel)
	
	GUICtrlSetState($progressState, $GUI_HIDE )
	$nTmpStepProess = 0
	$nTmpCountProess = 0
EndFunc

;----------------------------------------- 查看全部品牌 --------------------------------------------
Func _ShowBrands()

	Local $aRetArray
	_FileReadToArray ( $sFileBrands, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "投影仪品牌", Default, 32 )

EndFunc

;----------------------------------------- 查看品牌参数 --------------------------------------------
Func _ShowArgs()

	Local $aRetArray
	_FileReadToArray ( $sFileBrandArgs, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "品牌参数", Default, 32 )

EndFunc

;----------------------------------------- 查看采集参数 --------------------------------------------
Func _ShowCollections()

	Local $aRetArray
	_FileReadToArray ( $sFileCollectArgs, $aRetArray)
	_ArrayDelete ( $aRetArray, 0 )
	_ArraySort ( $aRetArray, 1 )
	_ArrayDisplay ( $aRetArray, "采集参数", Default, 32 )

EndFunc

;----------------------------------------- 查看采集结果 --------------------------------------------
Func _ShowCache()
	ShellExecute('explorer', $sRootPath )
EndFunc
