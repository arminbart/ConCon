Public MustInherit Class clContactFile

	Private mstrFile As String
	Private moContacts As clContactList = Nothing
	Protected moLogger As Logger

	Public Sub New(strFile As String, oLogger As Logger)
		mstrFile = strFile
		moLogger = oLogger
	End Sub

	Public Sub New(oContacts As clContactList)
		moContacts = oContacts
		moLogger = Nothing
	End Sub

	Public ReadOnly Property FileName As String
		Get
			Return mstrFile
		End Get
	End Property

	Public ReadOnly Property IsValidFile() As Boolean
		Get
			Try
				Return Not GetContacts().IsEmpty
			Catch ex As Exception
				LogException(moLogger, ex, FileName)
				Return False
			End Try
		End Get
	End Property

	Public Function GetContacts() As clContactList

		If moContacts Is Nothing Then moContacts = ReadFile()

		Return moContacts
	End Function


	Protected MustOverride Function ReadFile() As clContactList

	Public MustOverride Sub WriteFile(strFile As String)


	Public Shared Function ReadFile(strFile As String, oLogger As Logger) As clContactList
		Try
			Dim oFile As clContactFile = Nothing

			If strFile.EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase) Then
				oFile = New clPSRawContactFile(strFile, oLogger)
			ElseIf strFile.EndsWith(".vcf", StringComparison.CurrentCultureIgnoreCase) Then
				oFile = New clVCardFile(strFile, oLogger)
			End If

			If oFile IsNot Nothing AndAlso oFile.IsValidFile Then Return oFile.GetContacts()
		Catch ex As Exception
			LogException(oLogger, ex, strFile)
		End Try

		Return New clContactList()
	End Function

	Public Shared Function ReadFilesFromFolder(strPath As String, oLogger As Logger) As clContactList
		Dim oContacts As New clContactList()

		Try
			Dim arrFiles As String() = System.IO.Directory.GetFiles(strPath)

			For Each strFile As String In arrFiles
				For Each oContact As clContact In clContactFile.ReadFile(strFile, oLogger)
					oContacts.AddLast(oContact)
				Next
			Next
		Catch ex As Exception
			LogException(oLogger, ex, strPath)
		End Try

		Return oContacts
	End Function

	Public Shared Sub LogException(oLogger As Logger, ex As Exception, strFile As String)
		oLogger.Log("Error: " & ex.Message & vbCrLf & "   from " & strFile)
	End Sub

End Class
