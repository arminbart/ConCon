Public MustInherit Class clContactFile

	Private mstrFile As String
	Private moContacts As clContactList = Nothing
	Private moLogger As Logger

	Public Sub New(strFile As String, oLogger As Logger)
		mstrFile = strFile
		moLogger = oLogger
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
				moLogger.Log("Error: " & ex.Message)
				Return False
			End Try
		End Get
	End Property

	Public Function GetContacts() As clContactList

		If moContacts Is Nothing Then moContacts = ReadFile()

		Return moContacts
	End Function


	Protected MustOverride Function ReadFile() As clContactList


	Public Shared Function ReadFile(strFile As String, oLogger As Logger) As clContactList
		Try
			Dim oFile As clContactFile = Nothing

			oLogger.Log("Read " & strFile)
			If strFile.EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase) Then
				oFile = New clPSRawContactFile(strFile, oLogger)
				If Not oFile.IsValidFile() Then oFile = Nothing
			ElseIf strFile.EndsWith(".vcf", StringComparison.CurrentCultureIgnoreCase) Then
				Debug.Assert(False, ".vcf yet to implement!")
				Throw New NotImplementedException()
			End If

			If oFile IsNot Nothing Then Return oFile.GetContacts()
		Catch ex As Exception
			oLogger.Log("Error: " & ex.Message)
		End Try

		Return New clContactList
	End Function

End Class
