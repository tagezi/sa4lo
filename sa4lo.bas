'Option Explicit
Private oLib As Object						'библиотека	
Private oDlg As Object						'основное диалоговое окно
Private oDlgSAD As Object					'диалоговое окно адреса
Private oDoc As Object						'документ
Private oSheet As Object					'рабочий лист
Private aAddress(1 to 5) As String			'массив для адресов ячеек показателей и факторам
Private aComBox(1 to 5) As String			'массив для ярлыков показателей
Private aAddrName (1 to 5) As String
Private aConfNameParam(1 to 7) As String	'массив для ярлыков показателей и факторов из файла настроек
Private aConfAddrParam(1 to 7) As String	'массив для адресов ячеек показателей и факторам из файла настроек
Private aConfTitleParam(1 to 7) As String	'массив для адресов названий показателей и факторам из файла настроек
Private nCount As Integer					'для подстчета ячеек (можно избавиться)
Private sSheet As String					'название листа
Private sStartCell As String				'тепвая ячейка таблицы
Private sFieldName As String				'имя поля параметра
Private sFileName As String					'имя файла настроек
Private bCloseFloodField As Boolean			'переменная для закрытия/открытия диалога для выбора адресов
Private bStartAnalysis As Boolean			'переменная для начала анализа
Private bStartFloodField As Boolean			'переменная для закрытия/открытия основного диалога

'Начальный блок запускающий компоненты для решения комплекса задач
Sub Main
	Dim sUrl As String	'строка адреса файла
	'грузим библиотеки и общую информацию 
	oLib = GlobalScope.BasicLibraries
	oLib.LoadLibrary("Tools")
	oLib = DialogLibraries.GetByName("SensitivityAnalysis")
	DialogLibraries.loadLibrary("SensitivityAnalysis")
	oDlg = CreateUnoDialog(oLib.GetByName("DialogSA"))
	'получаем информациюо документе
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

'читаем настройки из файла, заполняем форму и массив для дальнейшего
'определения изменений в форме
Function fRead(Optional sFileName As String)
	Dim aTempConfParam(1 to 3) As String	'массив для временного хранения подстрок для заполнения формы
	iNumField = 1
	iNumber = Freefile
	Open sFileName For Input As iNumber		'открываем файл на чтение
	While Not eof(iNumber)
		Line Input #iNumber, sLine
		If sLine <> sFileName and sLine <> "" Then
			'0 элемент массива всегда хранит метку поля;
			'1 элемент адрес ячейки показателя/фактора или значение флажка;
			'2 элемент хранит адрес названия
			aTempConfParam = split(sLine, ";")
			'проверяем флажок, по умолчанию в форме он установлен в 1
			If (aTempConfParam(0) = "CheckBox1") Then
				aConfNameParam(7) = "CheckBox1"
				If (aTempConfParam(1) <> "1") Then
					oDlg.getControl("CheckBox1").State = False
					aConfAddrParam(7) = 0
				else
					aConfAddrParam(7) = 1
				End If
			'Fac - метка где хранится адрес диапазона факторов
			elseIF (aTempConfParam(0) = "Fac") Then
				oDlg.getControl("TextField11").setText(aTempConfParam(1))
				aConfAddrParam(iNumField) = aTempConfParam(1)
			else
			'иначе остались только поля показателей
				'значение выподающего списка с показателями
				oDlg.getControl("ComboBox" & iNumField).setText(aTempConfParam(0))
				aConfNameParam(iNumField) = aTempConfParam(0)
				'делаем нижние поля доступными
				EnableUpAvto("ComboBox" & iNumField)		
				'адрес ячейки с расчитаным показателем
				oDlg.getControl("TextField" & iNumField).setText(aTempConfParam(1))
				aConfAddrParam(iNumField) = aTempConfParam(1)
				'адрес названия
				oDlg.getControl("TextField" & (iNumField + 5 )).setText(aTempConfParam(2))
				aConfTitleParam(iNumField) = aTempConfParam(2)
				iNumField = iNumField + 1
			End If
        End If
   	Wend
    Close #iNumber
End Function

'пишем настройки в файл
Function fSave(sAddress As String)
	iNumber = Freefile
	Open sFileName For Output As #iNumber
		iNumField = 1
		'первая строка полный путь до файла с именем
		Print #iNumber, sFileName
		'вторая строка значение флажка
		Print #iNumber, "CheckBox1" & ";" & oDlg.GetControl("CheckBox1").getState()
			'в цикле пишем показатели
			While aComBox(iNumField) <> ""
				Print #iNumber, aComBox(iNumField) & ";" & _
				Right(aAddress(iNumField), (Len(aAddress(iNumField)) - 1)) & ";" & _
				Right(aAddrName(iNumField), (Len(aAddrName(iNumField)) - 1)) & ";"
				iNumField = iNumField + 1
			Wend
		'Fac - адрес диапазона факторов
		Print #iNumber, "Fac" & ";" & Right(sAddress, (Len(sAddress) - 1))
		Close #iNumber
End Function

'Облегчает управление основным диалоговым окном
Sub StartAnalysisDialog(sAddress As String)
	Dim oDlgModel As Object
	oDlg.setVisible(True)
	If (sFieldName <> "") Then
		oDlg.getControl(sFieldName).setText(sAddress)
	End If
	bStartFloodField = False
	bStartAnalysis = False
	'ждём нажатия кнопки
	Do
		If bStartAnalysis Then
		'начинаем расчеты
			oDlg.setVisible(False)
			StartAnalysis()
			Exit Do
		elseif bStartFloodField Then
		'открывает вспомогательное окно для адреса
			oDlg.setVisible(False)
			StartAddressDialog()
			bStartFloodField = False
			exit Do
		End If
		wait (100)
	Loop
End Sub

'облегчает управление диалоговым окнов для ввода адреса
Sub StartAddressDialog
	Dim Controls() As Object, oDlgSADModel As Object, Doc As Object, TextFieldModel As Object
	oDlg.setVisible(False)
'	oLib = DialogLibraries.GetByName("SensitivityAnalysis")
	oDlgSAD = CreateUnoDialog(oLib.GetByName("Address"))
	oDlgSADModel = oDlgSAD.Model
	Doc = ThisComponent
	oDlgSAD.setVisible(True)
	bCloseFloodField = False
	'ждём нажатия кнопки
	Do
		Controls() = oDlgSADModel.getControlModels
		TextFieldModel = Controls(0)
		TextFieldModel.Text = Doc.CurrentSelection.AbsoluteName
		If bCloseFloodField then
		'пересылаем данные в основное окно
			oDlgSAD.setVisible(False)
			StartAnalysisDialog(TextFieldModel.Text)
			Exit Do
		End If
		wait (100)
	Loop
End Sub

'сравнение массава данных из файла настроек и значений полей
'если они не совпадают, нужно переделывать листы
'НЕРАБОТАЕТ, пересоздание листов идёт в любом случае
Function ChangesCheck() As Boolean
	Dim iNumField As Integer
	iNumField = 1
	'если значение флажка изменено, значит нужно делать формы по новой
	If (oDlg.GetControl("CheckBox1").State() <> aConfAddrParam(7)) Then
		ChangesCheck = True
		Exit Function
	End If
	Do
		If (aConfNameParam(iNumField) = "") Then
			If (oDlg.GetControl("ComboBox" & iNumField).Text <> "Не использовать") Then
				ChangesCheck = True
				Exit Function
			End If
		else
			If (oDlg.GetControl("ComboBox" & iNumField).Text <> aConfNameParam(iNumField)) Then
				ChangesCheck = True
				Exit Function
			elseif (oDlg.GetControl("TextField" & iNumField).Text <> aConfAddrParam(iNumField)) Then
				str1 = oDlg.GetControl("TextField" & iNumField).Text
				ChangesCheck = True
				Exit Function
			End If
			'если значение адреса заголовка в масиве не равно нулю, то	
			If (aConfTitleParam(iNumField) <> "") Then
				'если поле адреса заголовка и значение в масиве не равны
				If (oDlg.GetControl("TextField" & (iNumField + 5)).Text <> aConfTitleParam(iNumField)) Then
					ChangesCheck = True
					Exit Function
				End If
			'если поле адреса не равно нулю
			elseif (oDlg.GetControl("TextField" & (iNumField + 5)).Text <> "") Then
				ChangesCheck = True
				Exit Function
			'во всех остальных случаях они равны
			End If
		End If
		iNumField = iNumField + 1
	Loop Until iNumField = 6
	ChangesCheck = False
End Function

'Удаление листов
Sub RemoveSheets
	Dim iNumField As Integer
	Dim sNameSheet As String
	iNumField = 1
	If (oDoc.hasByName("Interim calculation")) Then
		oDoc.removeByName("Interim calculation")
	End If
	sNameSheet = aConfNameParam(iNumField)				
	Do
		If (oDoc.hasByName(sNameSheet)) Then
			oDoc.removeByName(sNameSheet)
		End If
		iNumField = iNumField + 1
		sNameSheet = aConfNameParam(iNumField)
	Loop Until sNameSheet = ""
End Sub

'Создание листов
Sub CreateSheets
	Dim iNumField As Integer, StartTable As Integer
	Dim oController As Object
	iNumField = 1
	iStartTable = 0
	oController = ThisComponent.getCurrentController
	'создаём табличные и графические формы
	If (oDlg.GetControl("CheckBox1").getState()) Then
		If (Not oDoc.hasByName("Interim calculation")) Then
			oDoc.insertNewByName("Interim calculation", iNumField, nCount)
			oController.setActiveSheet(oDoc.GetByName("Interim calculation"))
			Do
				CreateTableForm(iStartTable, iNumField)
				iStartTable = iStartTable + nCount + 25
				iNumField = iNumField + 1
			Loop Until aComBox(iNumField) = ""
		End If
	else
		Do 
			If (Not oDoc.hasByName(aComBox(iNumField))) Then
				oDoc.insertNewByName(aComBox(iNumField), iNumField)
				oController.setActiveSheet(oDoc.GetByName(aComBox(iNumField)))
				CreateTableForm(iStartTable, iNumField)
			End If
			iStartTable = 0
			iNumField = iNumField + 1
		Loop Until aComBox(iNumField) = ""
	End If
End Sub

'Запуск основного блока
Sub StartAnalysis
	Dim iNumField As Integer, iNumber As Integer, StartTable As Integer
	Dim valCof As Double
	Dim sAddress As String, sRang As String, sNameSheet As String
	Dim oColumn As Object, oCellRange As Object
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
		'проверяем были ли изменены данные в диалоге		
		If (ChangesCheck()) Then
			RemoveSheets()		'удаляем старые листы
			fSave(sAddress)		'записываем новые данные в файл
		End If
		'Получаем предварительные сведения для создания новых листов
		oSheet = oDoc.GetByName(sSheet)
		nCount = oSheet.getCellRangeByName(sRang).getRows().getCount()
		oColumn = oSheet.getColumns().getByIndex(sRang)
		'проверяем наличие и создаём табличные и графические формы
		CreateSheets()
		'запускаем расчет анализа чувсвительности
		SensitivityAnlysis(nCount, sStartCell)
	else
		MsgBox "Вы не указали ни одного поля."
	End If
	Stop
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
		'считаем, если документ один (тестовые расчеты)
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
            	If (CellColumn = 6) Then
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
	Dim aTempAddres(0 to 4) As String
	Dim sComBox As String, sAddress As String, sAddrName As String
	Dim iNumField As Integer
	iNumField = 1
	While iNumField <> 5
		sAddress = "=" & oDlg.GetControl("TextField" & iNumField).getModel().Text
		sAddrName = "=" & oDlg.GetControl("TextField" & (iNumField + 5)).getModel().Text
		sComBox = oDlg.GetControl("ComboBox" & iNumField).getModel().Text
		If (sComBox <> "Не использовать") Then
			If (sAddress <> "=") Then
				aTempAddres = Split(sAddress, "$")
				aTempAddres = Split(aTempAddres(1)&aTempAddres(2)&aTempAddres(3), "'")
				aTempAddres = Split(aTempAddres(1)&aTempAddres(2), ".")
				aAddress(iNumField) = sAddress
				If (sAddrName <> "=") Then
					aTempAddres = Split(sAddrName, "$")
					aTempAddres = Split(aTempAddres(1)&aTempAddres(2)&aTempAddres(3), "'")
					aTempAddres = Split(aTempAddres(1)&aTempAddres(2), ".")
					aAddrName(iNumField) = sAddrName
				End If
			else
				MsgBox "Поле адреса показателя " & sComBox & " пусто!"
				Stop
			End If
			aComBox(iNumField) = sComBox
		End If
		iNumField = iNumField + 1
	Wend
	FieldTest = iNumField
End Function

'блок создания табличных форм
Sub CreateTableForm (StartTable as Integer, iTitleTable As Integer )
	Dim oSheetSource As Object, oRangCells As Object
	Dim numS As Integer, ind As Integer, _
	 nCfCellColumn As Integer, nCfCellRow As Integer, CountRow As Integer
	CreateChart(iTitleTable, StartTable) 'создание графических форм
	'делаем заголовок таблиной формы
	oSheetSource = oDoc.getByName(sSheet)
	If (oDlg.GetControl("CheckBox1").getState()) Then
		oSheet = oDoc.getByName("Interim calculation")
	else
		oSheet = oDoc.getByName(aComBox(iTitleTable))
	End If
	oRangCells = oSheet.getCellRangeByPosition(0, StartTable, 11, StartTable)
	oRangCells.Merge(True)							'объединяем ячейки
	oRangCells.CharWeight = 150						'шрифт жирный
	oRangCells.HoriJustify = 2						'по центру
	oRangCells.CellBackColor = RGB(151, 151, 151)	'цвет фона
	oSheet.getCellByPosition(0,StartTable).String = aComBox(iTitleTable)	
	numS = -50
	ind = 1
	StartTable = StartTable + 1
	'делаем подзаголовок табличной формы
	While ind < 12
		oSheet.getCellByPosition(ind, StartTable).String = numS & "%"
		numS = numS + 10
		ind = ind + 1
	Wend
	oRangCells = oSheet.getCellRangeByPosition(0,StartTable, 11, StartTable)
	oRangCells.HoriJustify = 2						'по центру
	oRangCells.CellBackColor = RGB(188, 188, 188)	'цвет фона
	StartTable = StartTable + 1
	CountRow = nCount + ind
	'копируем ссылки в центральный столбец
	nCfCellColumn = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Column
	nCfCellRow = oSheetSource.getCellRangeByName(sStartCell).getCellAddress.Row
	While ind < CountRow
		oCellCopy = oSheetSource.getCellByPosition((nCfCellColumn - 1), nCfCellRow)
		oCellPast = oSheet.getCellByPosition(0, StartTable)
		oCellPast.DataArray = oCellCopy.DataArray
		oSheet.getCellByPosition(6, StartTable).Formula = aAddress(iTitleTable)
		nCfCellRow = nCfCellRow + 1
		StartTable = StartTable + 1
		ind = ind + 1
	Wend
end Sub

'Создание графических форм
Sub CreateChart(iTitleTable As Integer, StartTableChart As Integer)
  Dim oSheet , oRect, oCharts, oChart, oChartDoc  As Object
  Dim sName, sDataRng As String
  Dim aTempAddres(0 to 4) As String
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
	RangeAddress(0).StartRow = (StartTableChart + 1)
	RangeAddress(0).EndColumn = 11
	RangeAddress(0).EndRow = (StartTableChart + nCount + 1)
	oCharts = oSheet.getCharts()
 	oCharts.addNewByName(sName, oRect, RangeAddress(), True, True)
 	oChart = oCharts.getByName(sName)
 	oChart.setRanges(RangeAddress())
 	oChartDoc = oChart.getEmbeddedObject()
 	oChartDoc.setDiagram(oChartDoc.createInstance("com.sun.star.chart.LineDiagram"))
 	oChartDoc.HasLegend = True
 	If (oDlg.GetControl("TextField" & (iTitleTable + 5)).Text = "") Then 
 		oChartDoc.getTitle().String = aComBox(iTitleTable)
 	else
 		aTempAddres = Split(oDlg.GetControl("TextField" & (iTitleTable + 5)).Text, "$")
		aTempAddres = Split(aTempAddres(1)&aTempAddres(2)&aTempAddres(3), "'")
		aTempAddres = Split(aTempAddres(1)&aTempAddres(2), "."
		oCellCopy = oDoc.getByName(aTempAddres(0)).getCellRangeByName(aTempAddres(1))
		'oCellCopy = oDoc.getSheets().getCellRangesByName(oDlg.GetControl("TextField" & (iTitleTable + 5)))
		oChartDoc.getTitle.String = oCellCopy.String
 	End If
 	oChartDoc.getDiagram().DataRowSource = Rows 
End Sub

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
		iNextNumField = iNumField + 5
		oDlg.GetControl("TextField" & iNextNumField).getModel().Enabled = True
		oDlg.GetControl("CommandButton" & iNextNumField).getModel().Enabled = True
	End If
End Function

'событие "нажата кнопка Ok" основного диалога
Sub onBtnStartAnalysis(oEvent)
	bStartAnalysis = True
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

'Закрытие без сохранения
Sub CloseDialog
    Stop
End Sub

'возвращает полное имя файла с путем до него в синтаксисе ОС
function GetConfFileName(sUrl As String)
	GetConfFileName = Left(sUrl, Len(sUrl) - 4) & ".conf"
End Function
