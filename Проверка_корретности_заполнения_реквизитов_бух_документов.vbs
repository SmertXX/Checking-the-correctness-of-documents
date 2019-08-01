SourceFolder1     = "\\s28175\XML\BOO\"
SourceFolder2     = "\\s28175\XML\BOO\BOO_XML\Загрузки\"
SourceFolder22    = "\\s28175\XML\BOO\Архив\"
SourceFolder31    = "\\s28175\XML\BOO\Ошибки\НетВБазе\"
SourceFolder35    = "\\s28175\XML\BOO\Ошибки\ПериодГод\"
SourceFolder6     = "\\s28175\XML\BOO\Ошибки\ОКПОИНН\"

SourceSQLServerGS = "BASE-F2"
SourceSQLBaseGS   = "BOO_2018"
SourceSQLTableGS  = "gs_tabl"

MyGod             = "2018"

Set objFSO           = WScript.CreateObject("Scripting.FileSystemObject")
Set objConnectionGS  = WScript.CreateObject("ADODB.Connection")

objConnectionGS.Open("Provider=SQLOLEDB.1;Data Source=" & SourceSQLServerGS & ";Initial Catalog=" & SourceSQLBaseGS & ";User ID=texnol;Password=20171903")

If NOT objFSO.FolderExists(SourceFolder2) Then
    Set objFolder = objFSO.CreateFolder(SourceFolder2)
End if

If NOT objFSO.FolderExists(SourceFolder22) Then
    Set objFolder = objFSO.CreateFolder(SourceFolder22)
End if

If NOT objFSO.FolderExists(SourceFolder31) Then
    Set objFolder = objFSO.CreateFolder(SourceFolder31)
End if

If NOT objFSO.FolderExists(SourceFolder35) Then
    Set objFolder = objFSO.CreateFolder(SourceFolder35)
End if

If NOT objFSO.FolderExists(SourceFolder6) Then
    Set objFolder = objFSO.CreateFolder(SourceFolder6)
End if


Set objFolder = Nothing

Sub sMySub()
    	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
   	xmlDoc.load(SourceFolder1 & FileName)

    	If objFSO.FileExists(SourceFolder22 & FileName) Then
        	objFSO.DeleteFile SourceFolder22 & FileName
    	End If
    	objFSO.CopyFile SourceFolder1 & FileName, SourceFolder22 & FileName


    	Set objNode = xmlDoc.documentElement.selectSingleNode("//Файл/Документ")
    	Period = objNode.getAttribute("Период")
    	OtchGod = objNode.getAttribute("ОтчетГод")
 
    	If (OtchGod <> MyGod) AND ((Period <> 34) OR (Period <> 90)) Then
        	If objFSO.FileExists(SourceFolder35 & FileName) Then
            		objFSO.DeleteFile SourceFolder35 & FileName
        	End If
            		objFSO.MoveFile SourceFolder1 & FileName, SourceFolder35 & FileName
        	Exit Sub
    	End If

    	INN = objNode.getAttribute("ИНН")
    	OKPO = objNode.getAttribute("ОКПО")

    	If (OKPO) Then
        	If NOT IsNumeric(OKPO) Then
			OKPO = 0
		End If
    	Else
        	OKPO = 0
    	End If

    	If (OKPO > 2147483646) Then
        	objFSO.MoveFile SourceFolder1 & FileName, SourceFolder6 & FileName
        	Exit Sub
    	End If


    	If (INN) Then
        	''''''
    	Else
        	INN = 0
    	End If

	If (OKPO=0) And (INN=0) Then
        	If objFSO.FileExists(SourceFolder6 & FileName) Then
            		objFSO.DeleteFile SourceFolder6 & FileName
        	End If
            		objFSO.MoveFile SourceFolder1 & FileName, SourceFolder6 & FileName
        	Exit Sub
    	End If

    	If (OKPO>0) And (INN>0) Then
		Set RS = Nothing
        	SQLstr = "Declare @a integer " _
            		& "SELECT @a=COUNT(okpo) FROM " & SourceSQLTableGS & " GROUP BY inn, okpo HAVING (inn = '" & INN & "') AND (okpo = '" & OKPO & "') " _
            		& "IF (@a IS NULL) SET @a=0 "_
            		& "SELECT @a;"
                Set RS = objConnectionGS.Execute(SQLstr)
                CountOKPOINN = RS(0)                
    	Else
        	CountOKPOINN = 0
    	End If

    	If (CountOKPOINN > 0) Then
		If objFSO.FileExists(SourceFolder22 & FileName) Then
          		objFSO.DeleteFile SourceFolder22 & FileName
        	End If
            		objFSO.CopyFile SourceFolder1 & FileName, SourceFolder22 & FileName

        	If objFSO.FileExists(SourceFolder2 & FileName) Then
            		objFSO.DeleteFile SourceFolder2 & FileName
        	End If
        	objFSO.MoveFile SourceFolder1 & FileName, SourceFolder2 & FileName
	Exit Sub
	End If

	CountINN = 0
    	CountOKPO = 0

    	If(CountOKPOINN=0) Then
        	If (INN>0) Then
            		SQLstr = "SELECT COUNT(inn) FROM " & SourceSQLTableGS & " WHERE (((inn)='" & INN & "'));"
                        Set RS = objConnectionGS.Execute(SQLstr)
                        CountINN = RS(0)
        	Else
            		CountINN = 0
        	End If

        	If (OKPO>0) Then
            		SQLstr = "SELECT COUNT(okpo) FROM " & SourceSQLTableGS & " WHERE (((okpo)='" & OKPO & "'));"
                        Set RS = objConnectionGS.Execute(SQLstr)
                        CountOKPO = RS(0)
        	Else
            		CountOKPO = 0
        	End If

        	If (CountOKPO=0 And CountINN=0) Then
            		If objFSO.FileExists(SourceFolder31 & FileName) Then
               			objFSO.DeleteFile SourceFolder31 & FileName
            		End If
            		objFSO.MoveFile SourceFolder1 & FileName, SourceFolder31 & FileName
            		Exit Sub
        	End If

        	If objFSO.FileExists(SourceFolder6 & FileName) Then
            		objFSO.DeleteFile SourceFolder6 & FileName
        	End If
            	objFSO.MoveFile SourceFolder1 & FileName, SourceFolder6 & FileName
        	Exit Sub
                        
    	End if
End Sub
