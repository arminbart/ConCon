Imports System.IO

Public Class clVCardFile
	Inherits clContactFile

	Private Class clVCardItem
		Public Name As String
		Public Value As String
		Public Params As Generic.List(Of String)

		Public Sub New(strName As String, strValue As String, arrKey As String())
			Name = strName
			Value = strValue

			' Ignore arrKey(0) because it contains strName
			For i As Integer = 1 To arrKey.Length - 1
				AddParam(arrKey(i))
			Next
		End Sub

		Public Sub AddParam(strParam As String)
			If Params Is Nothing Then Params = New Generic.List(Of String)()
			Params.Add(strParam)
		End Sub

		Public ReadOnly Property ValueArray As String()
			Get
				Return Value.Split({";"c}, StringSplitOptions.None)
			End Get
		End Property
	End Class

	Public Sub New(strFile As String, oLogger As Logger)
		MyBase.New(strFile, oLogger)
	End Sub

	Public Sub New(oContacts As clContactList)
		MyBase.New(oContacts)
	End Sub

	Protected Overrides Function ReadFile() As clContactList
		Dim oContacts As New clContactList()
		Dim oReader As StreamReader = File.OpenText(FileName)

		While oReader.Peek() <> -1
			Dim strLine As String = oReader.ReadLine().Trim()

			If IsAnyOf(strLine, "BEGIN:VCARD") Then
				oContacts.Add(ReadContact(oReader))
			ElseIf IsNotEmpty(strLine) Then
				Throw New Exception("File is not a valid vCard")
			End If
		End While

		Return oContacts
	End Function

	Public Overrides Sub WriteFile(strFile As String)
		Dim oWriter As New StreamWriter(strFile)

		For Each oContact As clContact In GetContacts()
			WriteContact(oWriter, oContact)
		Next

		oWriter.Close()
	End Sub

	Private Function ReadContact(oReader As StreamReader) As clContact
		Dim oContact As New clContact(FileName)
		Dim oItems As New Generic.Dictionary(Of String, clVCardItem)()
		Dim oItem As clVCardItem = Nothing
		Dim nItem As Integer = 0

		While oReader.Peek() <> -1
			Dim strLine As String = oReader.ReadLine().TrimEnd()
			Dim arrLine As String() = strLine.Split({":"c}, 2)
			Dim strKey As String = arrLine(0).Trim().ToUpper()

			If strLine.Trim().Length = 0 Then Continue While

			If strKey = "VERSION" Then Continue While ' We don't care (only know version 2.1 and 3.0 and hope the others are compatible ;-)
			If strKey = "END" Then Exit While

			Dim arrKey As String() = strKey.Split({";"c}) : strKey = arrKey(0)
			Dim strValue As String = If(arrLine.Length > 1, arrLine(1).Trim(), "")

			If strKey.StartsWith("ITEM") AndAlso strKey.Contains(".") Then ' Items with "ITEMx." (multi line items)
				Dim strItem As String = strKey.Split("."c)(0) : strKey = strKey.Split("."c)(1)

				If oItems.ContainsKey(strItem) Then
					oItem = oItems(strItem)
					oItem.AddParam(strKey & "=" & strValue.ToUpper())
				Else
					oItem = New clVCardItem(strKey, strValue, arrKey)
					oItems.Add(strItem, oItem)
				End If
			Else ' Items without "ITEMx." (single line items)
				oItem = New clVCardItem(strKey, strValue, arrKey)
				oItems.Add("_" & nItem, oItem) : nItem += 1
			End If
		End While

		For Each oItem In oItems.Values
			ReadItem(oItem, oContact)
		Next

		Return oContact
	End Function

	Private Sub ReadItem(oItem As clVCardItem, oContact As clContact)
		Select Case oItem.Name
			Case "N" : ReadName(oItem, oContact)
			Case "FN" : oContact.DisplayName = oItem.Value
			Case "ORG" : oContact.Organization = oItem.Value.Trim(New Char() {vbTab, " "c, ";"c})
			Case "TITLE" : oContact.Title = oItem.Value
			Case "TEL", "EMAIL" : ReadChannel(oItem, oContact)
			Case Else : Throw New Exception("clVCardFile.ReadContact(): Unsupported element '" & oItem.Name & "'!")
		End Select
	End Sub

	Private Sub ReadName(oItem As clVCardItem, oContact As clContact)
		Dim arrValue As String() = oItem.ValueArray

		If arrValue.Length > 0 Then oContact.LastName = arrValue(0).Trim()
		If arrValue.Length > 1 Then oContact.FirstName = arrValue(1).Trim()
		If arrValue.Length > 2 Then oContact.MiddleName = arrValue(2).Trim()
		If arrValue.Length > 3 Then oContact.Prefix = arrValue(3).Trim()
		If arrValue.Length > 4 Then oContact.Suffix = arrValue(4).Trim()

	End Sub

	Private Function GetChannelTagFromType(nType As clContact.enType, bPrimary As Boolean) As String
		Dim strTag As String = Nothing

		Select Case nType
			Case clContact.enType.MobilePrivate : strTag = "TEL;CELL"
			Case clContact.enType.PhonePrivate : strTag = "TEL;HOME"
			Case clContact.enType.FaxPrivate : strTag = "TEL;HOME;FAX"
			Case clContact.enType.MobileBusiness : strTag = "TEL;CELL" ' Don't know if there are sub types of CELL
			Case clContact.enType.PhoneBusiness : strTag = "TEL;WORK"
			Case clContact.enType.FaxBusiness : strTag = "TEL;WORK;FAX"
			Case clContact.enType.MobileOther : strTag = "TEL;CELL" ' Don't know if there are sub types of CELL
			Case clContact.enType.PhoneOther : strTag = "TEL"
			Case clContact.enType.FaxOther : strTag = "TEL;FAX"
			Case clContact.enType.PhoneMain : strTag = "TEL;MAIN"

			Case clContact.enType.EMailPrivate : strTag = "EMAIL;INTERNET;HOME"
			Case clContact.enType.EMailBusiness : strTag = "EMAIL;INTERNET;WORK"
			Case clContact.enType.EMailOther : strTag = "EMAIL;INTERNET"

			Case Else
				Throw New Exception("clVCardFile.GetChannelTagFromType(): Unsupported channel type '" & nType & "'!")
		End Select

		Return If(bPrimary, strTag & ";PREF", strTag)
	End Function

	Private Function GetChannelType(oItem As clVCardItem, ByRef bPrimary As Boolean) As clContact.enType
		Dim nType As clContact.enType
		Dim nLine As clContact.enType = clContact.enType.PhoneOther

		Select Case oItem.Name
			Case "TEL" : nType = clContact.enType.PhoneOther
			Case "EMAIL" : nType = clContact.enType.EMailOther
			Case Else : Throw New Exception("clVCardFile.GetChannelType(): Unsupported channel type '" & oItem.Name & "'!")
		End Select

		If oItem.Params IsNot Nothing Then
			For Each strParam As String In oItem.Params
				If strParam.Contains("=") Then strParam = strParam.Split("="c)(1)

				Select Case strParam
					Case "PREF" : bPrimary = True
					Case "CELL", "MOBIL" : nLine = clContact.enType.MobileOther
					Case "FAX" : nLine = clContact.enType.FaxOther
					Case "HOME", "PRIVAT"
						If nType = clContact.enType.EMailOther Then Return clContact.enType.EMailPrivate
						nType = clContact.enType.PhonePrivate
					Case "WORK", "ARBEIT"
						If nType = clContact.enType.EMailOther Then Return clContact.enType.EMailBusiness
						nType = clContact.enType.PhoneBusiness
					Case "MAIN"
						If nType = clContact.enType.PhoneOther Then Return clContact.enType.PhoneMain
					Case Else
						' Ignore other parameters
				End Select
			Next

			Select Case nType
				Case clContact.enType.PhonePrivate
					If nLine = clContact.enType.MobileOther Then nType = clContact.enType.MobilePrivate
					If nLine = clContact.enType.FaxOther Then nType = clContact.enType.FaxPrivate
				Case clContact.enType.PhoneBusiness
					If nLine = clContact.enType.MobileOther Then nType = clContact.enType.MobileBusiness
					If nLine = clContact.enType.FaxOther Then nType = clContact.enType.FaxBusiness
				Case clContact.enType.PhoneOther
					nType = nLine
			End Select
		End If

		Return nType
	End Function

	Private Sub ReadChannel(oItem As clVCardItem, oContact As clContact)
		Dim bPrimary As Boolean = False
		Dim nType As clContact.enType = GetChannelType(oItem, bPrimary)

		oContact.AddChannel(New clContact.clChannel(oItem.Value, nType, bPrimary))

	End Sub

	Private Sub WriteContact(oWriter As StreamWriter, oContact As clContact)

		Write(oWriter, "BEGIN", "VCARD")
		Write(oWriter, "VERSION", "2.1")
		Write(oWriter, "N", oContact.LastName, oContact.FirstName, oContact.MiddleName, oContact.Prefix, oContact.Suffix)
		Write(oWriter, "FN", oContact.DisplayName)

		For Each oChannel As clContact.clChannel In oContact.Channels
			For Each strContact In oChannel.Contact.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
				Write(oWriter, GetChannelTagFromType(oChannel.Type, oChannel.Primary), strContact)
			Next
		Next

		Write(oWriter, "ORG", oContact.Organization)
		Write(oWriter, "TITLE", oContact.Title)
		Write(oWriter, "END", "VCARD")

	End Sub

	Private Sub Write(oWriter As StreamWriter, strKey As String, strValue As String)
		Write(oWriter, strKey, New String() {strValue})
	End Sub

	Private Sub Write(oWriter As StreamWriter, strKey As String, ParamArray arrValue As String())
		If IsEmpty(arrValue(0)) AndAlso (arrValue.Length < 2 OrElse IsEmpty(arrValue(1))) Then Exit Sub

		Dim strValue As New System.Text.StringBuilder(arrValue(0))

		For i As Integer = 1 To arrValue.Length - 1
			strValue.Append(";").Append(If(arrValue(i) Is Nothing, "", arrValue(i).Trim()))
		Next

		oWriter.WriteLine(strKey & ":" & strValue.ToString())
	End Sub

End Class
