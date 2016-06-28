'Option Explicit	
Private oDlg, oDlgSAD, oDoc, oLib As Object
Private aAddress(1 to 5) As String
Private aComBox(1 to 5) As String
Private aConfNameParam(1 to 7) As String
Private aConfAddrParam(1 to 7) As String
Private aConfTitleParam(1 to 7) As String
Private nCount As Integer
Private sSheet , sStartCell, sFieldName, sFileName As String
Private bCloseFloodField, bStartAnalysis, bStartFloodField As Boolean

'Начальный блок запускающий компоненты для решения комплекса задач
Sub Main
	Dim sUrl, sLine As String
	'грузим библиотеки и общую информацию 
	If not isLibraryLoaded() Then Exit Sub
	oLib = DialogLibraries.GetByName("SensitivityAnalysis")
	oDlg = CreateUnoDialog(oLib.GetByName("DialogSA"))
	oDoc = ThisComponent.Sheets
	'получаем адрес файла
	sUrl = ThisComponent.getURL()
	sFileName = Left(sUrl, Len(sUrl) - 4) & ".conf"
	'читаем файл настроек если есть
	If (FileExists(sFileName)) Then 
		fRead(sFileName)
	End If
	StartAnalysisDialog("") 'запускаем основной диалог
End Sub

'читаем настройки из файла
Function fRead(Optional sFileName As String)
	Dim aTempConfParam(1 to 3) As String
	iNumField = 1
	iNumber = Freefile
	Open sFileName For Input As iNumber
	While Not eof(iNumber)
		Line Input #iNumber, sLine
		If sLine <> sFileName and sLine <> "" Then
			aTempConfParam = split(sLine, ";")
			aConfAddrParam(7) = 1
			If (aTempConfParam(1) <> "1" and aTempConfParam(0) = "CheckBox1") Then
				oDlg.getControl("CheckBox1").State = False
				aConfNameParam(7) = "CheckBox1"
				aConfAddrParam(7) = 0
			elseIf (aTempConfParam(0) <> "CheckBox1" and aTempConfParam(0) <> "Fac") Then
				oDlg.getControl("ComboBox" & iNumField).setText(aTempConfParam(0))
				aConfNameParam(iNumField) = aTempConfParam(0)
				EnableUpAvto("ComboBox" & iNumField)
				oDlg.getControl("TextField" & iNumField).setText(aTempConfParam(1))
				aConfAddrParam(iNumField) = aTempConfParam(1)
				If (aTempConfParam(1) = "Собственный") Then
					oDlg.getControl("TextField" & (iNumField + 5 )).setText(aTempConfParam(2))
					aConfTitleParam(iNumField) = aTempConfParam(2)
				end If
				iNumField = iNumField + 1
			else
				oDlg.getControl("TextField11").setText(aTempConfParam(1))
				aConfAddrParam(iNumField) = aTempConfParam(1)
            end if
        End If
   	Wend
    Close #iNumber
End Function

Function fSave(sAddress As String)
	iNumber = Freefile
	Open sFileName For Output As #iNumber
		iNumField = 1
		Print #iNumber, sFileName
		Print #iNumber, "CheckBox1" & ";" & oDlg.GetControl("CheckBox1").getState()
			While aComBox(iNumField) <> ""
				Print #iNumber, aComBox(iNumField) & ";" & Right(aAddress(iNumField), (Len(aAddress(iNumField)) - 1)) & ";"
				iNumField = iNumField + 1
			Wend
		Print #iNumber, "Fac" & ";" & Right(sAddress, (Len(sAddress) - 1))
		Close #iNumber
End Function

'Облегчает управление основным диалоговым окном
Sub StartAnalysisDialog(sAddress As String)
	Dim oDlgModel As Object
	oDlg.setVisible(True)
	If (Not IsEmpty(sFieldName)) Then oDlg.getControl(sFieldName).setText(sAddress)
	bStartFloodField = False
	bStartAnalysis = False
	'ждём нажатия кнопки
	Do
		If bStartAnalysis Then
		'начинаем расчеты
			oDlg.setVisible(False)
			StartAnalysis()
			exit Do
		elseif bStartFloodField Then
		'открывает вспомогательное окно для адреса
			oDlg.setVisible(False)
			StartAddressDialog()
			bStartFloodField = False
			exit Do
		End If
		wait (100)
	Loop
	Stop
End Sub

'облегчает управление диалоговым окнов для ввода адреса
Sub StartAddressDialog
	Dim Controls(), oDlgSADModel, Doc, TextFieldModel As Object
	oDlg.setVisible(False)
	oLib = DialogLibraries.GetByName("SensitivityAnalysis")
	oDlgSAD = CreateUnoDialog(oLib.GetByName("Address"))
	oDlgSADModel = oDlgSAD.Model
	Doc = ThisComponent
	oDlgSAD.setVisible(True)
	bCloseFloodField = false
	'ждём нажатия кнопки
	Do
		Controls() = oDlgSADModel.getControlModels
		TextFieldModel = Controls(0)
		TextFieldModel.Text = Doc.CurrentSelection.AbsoluteName
		If bCloseFloodField then
		'пересылаем данные в основное окно
			oDlgSAD.setVisible(False)
			StartAnalysisDialog(TextFieldModel.Text)
			exit Do
		End If
		wait (100)
	Loop
End Sub

'сравнение массивов, если массивы не одинаковы, возвращаем Истину
Function ChangesCheck() As Boolean
	Dim iNumField As Integer
	Dim sChangesCheck As Boolean
	iNumField = 1
	sChangesCheck = False
	If (oDlg.GetControl("CheckBox1").getState() <> aConfAddrParam(7)) Then
		sChangesCheck = True
	End If
	While iNumField <> 7 or sChangesCheck <> True
		On Error GoTo SkipStap
		If (oDlg.GetControl("ComboBox" & iNumField).Text <> aConfNameParam(iNumField)) Then
			sChangesCheck = True
		End If
		If (oDlg.GetControl("TextField" & iNumField).Text <> aConfAddrParam(iNumField)) Then
			sChangesCheck = True
		End If
		iNumField = iNumField + 1
	Wend
	ChangesCheck = sChangesCheck
	Exit Function
	SkipStap:
		ChangesCheck = True	
End Function

'Запуск основного блока
Sub StartAnalysis
	Dim iNumField, iNumber, StartTable As Integer
	Dim valCof As Double
	Dim sAddress, sRang, sNameSheet As String
	Dim oSheet, oWorkSheet, oCellRange, oColumns, oColumn, oController As Object
	Dim aTempAddres (0 to 4) As String
	oDlg.setVisible(False)
	'проверка полей
	If (FieldTest() <> 0 ) Then	
		sAddress = "=" & oDlg.GetControl("TextField11").getModel().Text
		If (sAddress <> "=") Then
			On Error GoTo ErrorAddress
			aTempAddres = Split(sAddress, "$")
			sRang = aTempAddres(2) & aTempAddres(3) & aTempAddres(4) & aTempAddres(5)
			sStartCell = aTempAddres(2) & aTempAddres(3)
			sStartCell = Left(sStartCell, Len(sStartCell) - 1)
			aTempAddres = Split(aTempAddres(1) & aTempAddres(2), "'")
			aTempAddres = Split(aTempAddres(0), ".")
		    sSheet = aTempAddres(0)
		else
			MsgBox "Поле Диапазон факторов пусто!"
			Stop
		End If
		'проверяем бфли ли изменены данные в диалоге
		If (ChangesCheck()) Then
			'если да, записываем новые данные в файл
			fSave(sAddress)
			iNumField = 1
			'удаляем лишние листы
			If (oDoc.hasByName("Interim calculation")) Then
				oDoc.removeByName("Interim calculation")
			End If
			sNameSheet = aComBox(iNumField)				
			While iNumField <> 5
				If (oDoc.hasByName(sNameSheet)) Then
					oDoc.removeByName(sNameSheet)
				End If
				iNumField = iNumField + 1
				sNameSheet = aComBox(iNumField)
			Wend
		End If
		'Получаем предварительные сведения для создания документов
		oSheet = oDoc.GetByName(sSheet)
	    oCellRange = oSheet.getCellRangeByName(sRang)
	    nCount = getCountNonEmpt(oCellRange)
	    oController = ThisComponent.getcurrentController
	    oColumns = oSheet.getColumns()
	    oColumn = oColumns.getByIndex(sRang)
	    StartTable = 1
	    iNumField = 1
	    'создаём табличные и графические формы
		If (oDlg.GetControl("CheckBox1").getState()) Then
			oDoc.insertNewByName("Interim calculation", iNumField, nCount)
			oWorkSheet = oDoc.GetByName("Interim calculation")
			oController.setActiveSheet(oWorkSheet)
			While aComBox(iNumField) <> ""
				CreateTableForm(StartTable, iNumField)
				StartTable = StartTable + nCount + 18
				iNumField = iNumField + 1
			Wend
		else
			While aComBox(iNumField) <> ""
				oDoc.insertNewByName(aComBox(iNumField), iNumField)
				oWorkSheet = oDoc.GetByName(aComBox(iNumField))
				oController.setActiveSheet(oWorkSheet)
				CreateTableForm(StartTable, iNumField)
				StartTable = 1
				iNumField = iNumField + 1
			Wend
		End If
		'запускаем расчет анализа чувсвительности
		SensitivityAnlysis(nCount, sStartCell)
	else
		MsgBox "Вы не указали ни одного поля."
		Stop
	End If
	Exit Sub
	ErrorAddress:
		MsgBox err & " Адрес для факторов указан не верно!"
		Stop
End Sub

'Расчет анализа чувствительности
Function SensitivityAnlysis(nCount As Integer, sStartCell As String)
	Dim iNumField As Integer
	Dim nCfVol As Double
	'настраиваем переменные
	oSheetSource = oDoc.getByName(sSheet)
	CellRowStart = 2
	CellColumn = 11
	iNumField = 1
	nCfVol = 1.5
	nCfCellColumn = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Column
	nCfCellRow = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Row
	'цыкл пока не просмотрим все факторы
	While CellRowStart <> (nCount + 2)
		CellRow = CellRowStart
		If (oDlg.GetControl("CheckBox1").getState()) Then
		'сичитаем, если документ один (тестоые расчеты)
			oSheet = oDoc.getByName("Interim calculation")
			While CellColumn <> 0
				oSheetSource.getCellByPosition(nCfCellColumn, nCfCellRow).Value = nCfVol
				While aComBox(iNumField) <> ""
					oCellCopy = oSheet.getCellByPosition(6, CellRow)
            		oCellPast = oSheet.getCellByPosition(CellColumn, CellRow)
            		oCellPast.DataArray = oCellCopy.DataArray
            		oCellPast.NumberFormat = oCellCopy.NumberFormat
            		CellRow = nCount + CellRow + 25
            		iNumField = iNumField + 1
            	Wend
            	CellColumn = CellColumn - 1
            	If (CellColumn = 6) Then
            		CellColumn = CellColumn - 1
            		nCfVol = nCfVol - 0.1
            	End If
            	nCfVol = nCfVol - 0.1
            	CellRow = CellRowStart
            	iNumField = 1
			Wend
		else
		'считаем, если документы раздельные (финальные расчет)
			While CellColumn <> 0
				oSheetSource.getCellByPosition(nCfCellColumn, nCfCellRow).Value = nCfVol
				While aComBox(iNumField) <> ""
					oSheet = oDoc.getByName(aComBox(iNumField))
					oCellCopy = oSheet.getCellByPosition(6, CellRow)
            		oCellPast = oSheet.getCellByPosition(CellColumn, CellRow)
            		oCellPast.DataArray = oCellCopy.DataArray
            		oCellPast.NumberFormat = oCellCopy.NumberFormat
            		iNumField = iNumField + 1
            	Wend
            	CellColumn = CellColumn - 1
            	if (CellColumn = 6) Then
            		CellColumn = CellColumn - 1
            		nCfVol = nCfVol - 0.1
            	end if
            	nCfVol = nCfVol - 0.1
            	iNumField = 1
			Wend
		End If
		oSheetSource.getCellByPosition(nCfCellColumn, nCfCellRow).Value = 1
		CellColumn = 11
		nCfCellRow = nCfCellRow + 1
		CellRowStart = CellRowStart + 1
		nCfVol = 1.5
	Wend
End Function

'Проверка полей
Function FieldTest() As Integer
	Dim aTempAddres (0 to 4) As String
	Dim sComBox, sAddress As String
	Dim iNumField, iArrayIndex As Integer
	iNumField = 1
	iArrayIndex = 0
	While iNumField <> 5
		sAddress = "=" & oDlg.GetControl("TextField" & iNumField).getModel().Text
		sComBox = oDlg.GetControl("ComboBox" & iNumField).getModel().Text
		If (sComBox <> "Не использовать") Then
			If (sAddress <> "=") Then
				aTempAddres = Split(sAddress, "$")
				aTempAddres = Split(aTempAddres(1)&aTempAddres(2)&aTempAddres(3), "'")
				aTempAddres = Split(aTempAddres(1)&aTempAddres(2), ".")
				iArrayIndex = iArrayIndex + 1
				aAddress(iArrayIndex) = sAddress
				aComBox(iArrayIndex) = sComBox
			else
				MsgBox "Поле " & sComBox & " пусто!"
				Stop
			End If
		End If
		iNumField = iNumField + 1
	Wend
	FieldTest = iArrayIndex
End Function

'блок создания табличных форм
Function CreateTableForm (StartTable as Integer, iTitleTable As Integer )
	Dim oDocement , dispatcher, oSheetSource, oSheet As object
	Dim StartTableChart, numS, ind, nCfCellColumn, nCfCellRow, CountRow As Integer
	Dim sABC (1 to 11) As String
	StartTableChart = StartTable
	CreateChart(iTitleTable, StartTableChart) 'создание графических форм
	oDocement = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	'делаем заголовок таблиной формы
	Dim args(0) as new com.sun.star.beans.PropertyValue
	args(0).Name = "ToPoint"
	args(0).Value = "$A$" & StartTable & ":$L$" & StartTable
	dispatcher.executeDispatch(oDocement, ".uno:GoToCell", "", 0, args())
	dispatcher.executeDispatch(oDocement, ".uno:MergeCells", "", 0, Array())
	args(0).Name = "HorizontalAlignment"
	args(0).Value = com.sun.star.table.CellHoriJustify.CENTER
	dispatcher.executeDispatch(oDocement, ".uno:HorizontalAlignment", "", 0, args())
	args(0).Name = "Bold"
	args(0).Value = true
	dispatcher.executeDispatch(oDocement, ".uno:Bold", "", 0, args())
	args(0).Name = "BackgroundColor"
	args(0).Value = 14540253
	dispatcher.executeDispatch(oDocement, ".uno:BackgroundColor", "", 0, args())
	args(0).Name = "StringName"
	args(0).Value = aComBox(iTitleTable)
	dispatcher.executeDispatch(oDocement, ".uno:EnterString", "", 0, args())
	dispatcher.executeDispatch(oDocement, ".uno:JumpToNextCell", "", 0, Array())
	sABC(1) = "B"
	sABC(2) = "C"
	sABC(3) = "D"
	sABC(4) = "E"
	sABC(5) = "F"
	sABC(6) = "G"
	sABC(7) = "H"
	sABC(8) = "I"
	sABC(9) = "J"
	sABC(10) = "K"
	sABC(11) = "L"
	numS = -50
	ind = 1
	StartTable = StartTable + 1
	'делаем подзаголовок табличной формы
	While ind < 12
		args(0).Name = "ToPoint"
		args(0).Value =  "$" & sABC(ind) & "$" & StartTable
		dispatcher.executeDispatch(oDocement, ".uno:GoToCell", "", 0, args())
		args(0).Name = "StringName"
		args(0).Value = numS & "%"
		dispatcher.executeDispatch(oDocement, ".uno:EnterString", "", 0, args())
		numS = numS + 10
		ind = ind + 1
	Wend
	args(0).Name = "ToPoint"
	args(0).Value = "$B$" & StartTable & ":$L$" & StartTable
	dispatcher.executeDispatch(oDocement, ".uno:GoToCell", "", 0, args())
	args(0).Name = "BackgroundColor"
	args(0).Value = 15658734
	dispatcher.executeDispatch(oDocement, ".uno:BackgroundColor", "", 0, args())
	args(0).Name = "NumberFormatValue"
	args(0).Value = 10
	dispatcher.executeDispatch(oDocement, ".uno:NumberFormatValue", "", 0, args())
	args(0).Name = "HorizontalAlignment"
	args(0).Value = com.sun.star.table.CellHoriJustify.CENTER
	dispatcher.executeDispatch(oDocement, ".uno:HorizontalAlignment", "", 0, args())
	StartTable = StartTable + 1
	CountRow = nCount + ind
	'копируем ссылки в центральный столбец
	oSheetSource = oDoc.getByName(sSheet)
	If (oDlg.GetControl("CheckBox1").getState()) Then
		oSheet = oDoc.getByName("Interim calculation")
	else
		oSheet = oDoc.getByName(aComBox(iTitleTable))
	End If
	nCfCellColumn = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Column
	nCfCellRow = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Row
	While ind < CountRow
		oCellCopy = oSheetSource.getCellByPosition((nCfCellColumn - 1), nCfCellRow)
		oCellPast = oSheet.getCellByPosition(0, (StartTable - 1))
		oCellPast.DataArray = oCellCopy.DataArray
		args(0).Name = "ToPoint"
		args(0).Value =  "$G" & "$" & StartTable
		dispatcher.executeDispatch(oDocement, ".uno:GoToCell", "", 0, args())
		args(0).Name = "StringName"
		args(0).Value = aAddress(iTitleTable)
		dispatcher.executeDispatch(oDocement, ".uno:EnterString", "", 0, args())
		nCfCellRow = nCfCellRow + 1
		StartTable = StartTable + 1
		ind = ind + 1
	Wend
end Function

'событие, если пользователь выбрал из списка показатель
Sub EnabledUp(NameCall)
	Dim cName As String
	cName = NameCall.Source.getModel().Name
	EnableUpAvto(cName)
End Sub

'Активация списка и полей при изменении предыдущего
Function EnableUpAvto(cName As String)
	Dim iNumField, iNextNumField As Integer
	iNumField = CDbl(right(cName, 1))
	If (oDlg.GetControl("ComboBox" & iNumField).getModel().Text <> "Не использовать" ) Then
		iNextNumField = iNumField + 1
		oDlg.GetControl("ComboBox" & iNextNumField).getModel().Enabled = True
		oDlg.GetControl("TextField" & iNextNumField).getModel().Enabled = True
		oDlg.GetControl("CommandButton" & iNextNumField).getModel().Enabled = True
		iNumField = CDbl(right(cName, 1))
	End If
	If (oDlg.GetControl("ComboBox" & iNumField).getModel().Text = "Собственный") Then
		iNextNumField = iNumField + 5
		oDlg.GetControl("TextField" & iNextNumField).getModel().Enabled = True
		oDlg.GetControl("CommandButton" & iNextNumField).getModel().Enabled = True
	End If
End Function

'Создание графических форм
Sub CreateChart(iTitleTable As Integer, StartTableChart As Integer)
  Dim oSheet , oRect, oCharts, oChart, oChartDoc, oTitle, oDiagram  As Object
  Dim sName, sDataRng As String
  Dim RangeAddress(0) As New com.sun.star.table.CellRangeAddress
  oRect = createObject("com.sun.star.awt.Rectangle")
	    oRect.X = 1000
	    oRect.Y = 4000
  If (oDlg.GetControl("CheckBox1").getState()) Then
		oSheet = oDoc.getByName("Interim calculation")
		If (iTitleTable > 1)  Then
	    	oRect.Y = 4500 + (4500 + 9000) * ( iTitleTable - 1 )
	    End If
	else
		oSheet = oDoc.getByName(aComBox(iTitleTable))
	End If
	oRect.width = 20000
	oRect.Height= 9000
    sName = "Chart" & iTitleTable
  	RangeAddress(0).Sheet = oSheet.getRangeAddress().Sheet
	RangeAddress(0).StartColumn = 0 
	RangeAddress(0).StartRow = (StartTableChart)
	RangeAddress(0).EndColumn = 11
	RangeAddress(0).EndRow = (StartTableChart + nCount)
	oCharts = oSheet.getCharts()
 	oCharts.addNewByName(sName, oRect, RangeAddress(), True, True)
 	oChart = oCharts.getByName(sName)
 	oChart.setRanges(RangeAddress())
 	oChartDoc = oChart.getEmbeddedObject()
 	oDiagram = oChartDoc.createInstance("com.sun.star.chart.LineDiagram")
 	oChartDoc.setDiagram(oDiagram)
 	oChartDoc.HasLegend = True 
 	oTitle = oChartDoc.getTitle()
 	oTitle.String = aComBox(iTitleTable)
 	oDiagram = oChartDoc.getDiagram()
 	oDiagram.DataRowSource = Rows 
End Sub

'событие "нажата кнопка Ok" вспомогательного диалога
Sub onBtnOKPressed(oEvent)
	bCloseFloodField = True
End Sub

'событие "нажата кнопка выбора адреса" лоя вызова вспомогательного окна
Sub onBtnStartFloodField(oEvent)
	bStartFloodField = True
	sFieldName = oEvent.Source.Model.Name
	sFieldName = "TextField" & Right(sFieldName, Len(sFieldName) - 13)
End Sub

'событие "нажата кнопка Ok" основного диалога
Sub onBtnStartAnalysis(oEvent)
	bStartAnalysis = True
End Sub

'возвращает полное имя файла с путем до него в синтаксисе ОС
function GetConfFileName(sUrl As String)
	GetConfFileName = Left(sUrl, Len(sUrl) - 4) & ".conf"
End Function

'Закрытие без сохранения
Sub CloseDialog
    Stop
End Sub

'Подсчет ячеек
Function getCountNonEmpt(oRange As Variant)
	Dim oQry, oCells, oEnum, iCountCells As Variant
    oQry = oRange.queryContentCells(com.sun.star.sheet.CellFlags.VALUE)
    oEnum =  oQry.getCells().createEnumeration()
    iCountCells = 0
    Do while oEnum.hasMoreElements()
        iCountCells = iCountCells + 1
        oEnum.nextElement()
    Loop
    getCountNonEmpt = iCountCells
End Function

REM AM:Listing 5.78: Using the ApplicationScriptLibraryContainer.
REM Modified: JohnSUN
REM Check whether is loaded the macro-library with the specified name
REM (by default "Tools"). If not already loaded, it tries to load.
Function isLibraryLoaded(Optional LibName As String) As Boolean
Dim oLibs As Object
	If IsMissing(LibName) Then LibName="Tools"
	oLibs = GlobalScope.BasicLibraries
	If oLibs.HasByName (LibName) Then
		If (Not oLibs.isLibraryLoaded(LibName)) Then
		oLibs.LoadLibrary(LibName)
		End If
		If (Not oLibs.isLibraryLoaded(LibName)) Then
			MsgBox("The library named """ + LibName + """ is not loaded!" + Chr(13) + _
			"Be sure to install the office properly." + Chr(13) + _
			"Some of the macros will not be available!", 48, "Attention!")
			isLibraryLoaded = False
		Else
			isLibraryLoaded = True
		End If
	Else
		MsgBox("The library named """ + LibName + """ not found!" + Chr(13) + _
		"Check the office's Libraries." + Chr(13) + _
		"Some of the macros will not be available!", 48, "Oops!")
		isLibraryLoaded = False
	End If
End Function
