Public Class clContact

	' Name
	Public DisplayName As String
	Public Prefix As String
	Public FirstName As String
	Public MiddleName As String
	Public LastName As String
	Public Suffix As String

	Public Channels As New Generic.LinkedList(Of clContact.clChannel)()

	Public FileName As String

	Public Sub New(strFile As String)
		FileName = strFile
	End Sub

	Public Enum enType
		Unknown = 0

		PHONE_BEGIN = 10
		MobilePrivate = 11
		PhonePrivate = 12
		FaxPrivate = 13
		MobileBusiness = 14
		PhoneBusiness = 15
		FaxBusiness = 16
		MobileOther = 17
		PhoneOther = 18
		FaxOther = 19
		PHONE_END = 20

		EMAIL_BEGIN = 20
		EMailPrivate = 21
		EMailBusiness = 22
		EMailOther = 22
		EMAIL_END = 23
	End Enum

	Public Shared Function GetTypeDesc(nType As enType) As String
		Select Case nType
			Case enType.MobilePrivate : Return "Mobile priv."
			Case enType.PhonePrivate : Return "Phone priv."
			Case enType.FaxPrivate : Return "Fax priv."
			Case enType.MobileBusiness : Return "Mobile biz."
			Case enType.PhoneBusiness : Return "Phone biz."
			Case enType.FaxBusiness : Return "Fax biz."
			Case enType.MobileOther : Return "Mobile oth."
			Case enType.PhoneOther : Return "Phone oth."
			Case enType.FaxOther : Return "Fax oth."

			Case enType.EMailPrivate : Return "E-Mail priv."
			Case enType.EMailBusiness : Return "E-Mail biz."
			Case enType.EMailOther : Return "E-Mail oth."

			Case Else : Return "Unknown"
		End Select
	End Function

	Public Shared Function GetTypeByDesc(strTypeDesc As String) As enType

		For nPhoneType As clContact.enType = clContact.enType.PHONE_BEGIN + 1 To clContact.enType.PHONE_END - 1
			If strTypeDesc.Equals(clContact.GetTypeDesc(nPhoneType), StringComparison.CurrentCultureIgnoreCase) Then Return nPhoneType
		Next

		For nEMailType As clContact.enType = clContact.enType.EMAIL_BEGIN + 1 To clContact.enType.EMAIL_END - 1
			If strTypeDesc.Equals(clContact.GetTypeDesc(nEMailType), StringComparison.CurrentCultureIgnoreCase) Then Return nEMailType
		Next

		Return enType.Unknown
	End Function

	Public Shared Function IsPhone(nType As enType)
		Return nType > enType.PHONE_BEGIN And nType < enType.PHONE_END
	End Function

	Public Shared Function IsEMail(nType As enType)
		Return nType > enType.EMAIL_BEGIN And nType < enType.EMAIL_END
	End Function

	Public Class clChannel
		Public Type As enType
		Public Contact As String

		Public Sub New(strContact As String, nType As enType)
			strContact = strContact.Trim().Replace(" ", "").ToLower()
			Contact = If(IsPhone(nType), FormatPhone(strContact), strContact)
			Type = nType
		End Sub

		Public Shared Function FormatPhone(strPhone As String)
			Dim strFormatted As New System.Text.StringBuilder()

			strPhone = strPhone.Trim()
			If strPhone.StartsWith("+") Then strFormatted.Append("+") : strPhone = strPhone.Substring(1)

			For i As Integer = 0 To strPhone.Length - 1
				If InStr(1, "0123456789", strPhone.Substring(i, 1)) Then strFormatted.Append(strPhone.Chars(i))
			Next

			Return strFormatted.ToString()
		End Function
	End Class

	Public Sub AddChannel(oChannel As clChannel)
		Channels.AddLast(oChannel)
	End Sub

	'Public Function GetPhones() As clChannel()
	'	Return (From c In Channels Where c.IsPhone Select c).ToArray()
	'End Function

	'Public Function GetEMails() As clChannel()
	'	Return (From c In Channels Where c.IsEMail Select c).ToArray()
	'End Function

End Class
