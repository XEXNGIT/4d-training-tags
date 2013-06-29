Function GETENV(variableName)
	
	Set objShell = CreateObject("WScript.Shell")
	Set theVariable = objShell.Environment("PROCESS")
	GETENV = theVariable(variableName)
	Set objShell = Nothing

end Function

Function GETAPP(applicationName)

	On Error Resume Next
		Set GETAPP = GetObject(, applicationName)
		If Err.Number <> 0 Then
			Set GETAPP = CreateObject(applicationName) 
		End If
	On Error GoTo 0

end Function

Set objExcelApplication	= GETAPP("Excel.Application")

theXmlPath = GETENV("XML_DOCUMENT_PATH")
theXlsPath = GETENV("XLS_DOCUMENT_PATH")

objExcelApplication.Visible = True

objExcelApplication.ScreenUpdating = False
objExcelApplication.DisplayAlerts = False

Set theWorkbooks = objExcelApplication.Workbooks

For Each theWorkbook In theWorkbooks
	If (theWorkbook.FullName = theXlsPath) Or (theWorkbook.Name = theXlsPath) Then
		theWorkbook.Close (False)
	End If
Next

Set theWorkbook = objExcelApplication.Workbooks.Open(theXmlPath)

theWorkbook.SaveAs theXlsPath, 56

objExcelApplication.ScreenUpdating = True

CreateObject("WScript.Shell").AppActivate objExcelApplication.Caption
