Imports System.IO

Public Class clVCardFile
	Inherits clContactFile

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

		While oReader.Peek() <> -1
			Dim strLine As String = oReader.ReadLine().Trim()
			Dim arrLine As String() = strLine.Split({":"c}, 2)
			Dim strKey As String = arrLine(0).Trim().ToUpper()
			Dim strValue As String = If(arrLine.Length > 1, arrLine(1).Trim(), "")

			Select Case strKey
				Case "VERSION" ' We don't care (only know version 2.1 and hope the others are compatible ;-)
				Case "END" : Exit While
				Case "N" : ReadName(strValue, oContact)
				Case "FN" : oContact.DisplayName = strValue
				Case "ORG" : oContact.Organization = strValue
				Case "TITLE" : oContact.Title = strValue
				Case "URL" : Debug.Assert(False, "URL tag not supported yet in vCards")
				Case Else
					Dim arrKey As String() = strKey.Split({";"c}, 2)

					If arrKey.Length > 1 AndAlso IsAnyOf(arrKey(0), "ADR") Then
						ReadAddress(arrKey, strValue, oContact)
					ElseIf arrKey.Length > 1 AndAlso IsAnyOf(arrKey(0), "TEL", "EMAIL") Then
						ReadChannel(arrKey, strValue, oContact)
					Else
						Throw New Exception("clVCardFile.ReadContact(): Unsupported element '" & strLine & "'!")
					End If
			End Select
		End While

		Return oContact
	End Function

	Private Sub ReadName(strValue As String, oContact As clContact)
		Dim arrValue As String() = strValue.Split({";"c}, StringSplitOptions.None)

		If arrValue.Length > 0 Then oContact.LastName = arrValue(0).Trim()
		If arrValue.Length > 1 Then oContact.FirstName = arrValue(1).Trim()
		If arrValue.Length > 2 Then oContact.MiddleName = arrValue(2).Trim()
		If arrValue.Length > 3 Then oContact.Prefix = arrValue(3).Trim()
		If arrValue.Length > 4 Then oContact.Suffix = arrValue(4).Trim()

	End Sub

	Private Sub ReadAddress(arrKey As String(), strValue As String, oContact As clContact)
		Debug.Assert(False, "ADR tag not supported yet in vCards")
	End Sub

	Private Function GetChannelTagFromType(nType As clContact.enType) As String
		Select Case nType
			Case clContact.enType.MobilePrivate : Return "TEL;CELL"
			Case clContact.enType.PhonePrivate : Return "TEL;HOME"
			Case clContact.enType.FaxPrivate : Return "TEL;HOME;FAX"
			Case clContact.enType.MobileBusiness : Return "TEL;CELL" ' Don't know if there are sub types of CELL
			Case clContact.enType.PhoneBusiness : Return "TEL;WORK"
			Case clContact.enType.FaxBusiness : Return "TEL;WORK;FAX"
			Case clContact.enType.MobileOther : Return "TEL;CELL" ' Don't know if there are sub types of CELL

			Case clContact.enType.EMailPrivate : Return "EMAIL;HOME"
			Case clContact.enType.EMailBusiness : Return "EMAIL;WORK"
			Case clContact.enType.EMailOther : Return "EMAIL;X-INTERNET"

			Case Else
				Throw New Exception("clVCardFile.GetChannelTagFromType(): Unsupported channel type '" & nType & "'!")
		End Select
	End Function

	Private Function GetChannelType(strType As String, strSubType As String) As clContact.enType

		If IsAnyOf(strType, "TEL") Then
			Select Case strSubType
				Case "CELL" : Return clContact.enType.MobilePrivate
				Case "HOME" : Return clContact.enType.PhonePrivate
				Case "HOME;FAX" : Return clContact.enType.FaxPrivate
				Case "WORK" : Return clContact.enType.PhoneBusiness
				Case "WORK;FAX" : Return clContact.enType.FaxBusiness
				Case Else : Return clContact.enType.PhoneOther
			End Select
		ElseIf IsAnyOf(strType, "EMAIL") Then
			Select Case strSubType
				Case "HOME" : Return clContact.enType.EMailPrivate
				Case "WORK" : Return clContact.enType.EMailBusiness
				Case "X-INTERNET" : Return clContact.enType.EMailOther
				Case Else : Return clContact.enType.EMailOther
			End Select
		End If

		Throw New Exception("clVCardFile.GetChannelType(): Unsupported channel type '" & strType & ";" & strSubType & "'!")
	End Function

	Private Sub ReadChannel(arrKey As String(), strValue As String, oContact As clContact)
		Dim strType = arrKey(0)
		Dim strSubType As String = arrKey(1).Replace(";PREF", "").Replace(";VOICE", "")

		oContact.AddChannel(New clContact.clChannel(strValue, GetChannelType(strType, strSubType)))
	End Sub

	Private Sub WriteContact(oWriter As StreamWriter, oContact As clContact)

		Write(oWriter, "BEGIN", "VCARD")
		Write(oWriter, "VERSION", "2.1")
		Write(oWriter, "N", oContact.LastName, oContact.FirstName, oContact.MiddleName, oContact.Prefix, oContact.Suffix)
		Write(oWriter, "FN", oContact.DisplayName)

		For Each oChannel As clContact.clChannel In oContact.Channels
			For Each strContact In oChannel.Contact.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
				Write(oWriter, GetChannelTagFromType(oChannel.Type), strContact)
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
