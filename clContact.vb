Public Class clContact

	' Name
	Public Property DisplayName As String
	Public Property Prefix As String
	Public Property FirstName As String
	Public Property MiddleName As String
	Public Property LastName As String
	Public Property Suffix As String

	Public Channels As New Generic.LinkedList(Of clContact.clChannel)()
	'Public Phones As New Generic.LinkedList(Of clContact.clPhone)()
	'Public EMails As New Generic.LinkedList(Of clContact.clEMail)()

	Public Enum enType
		Unknown = 0

		PHONE_BEGIN = 10
		MobilePrivate = 11
		PhonePrivate = 12
		FaxPrivate = 13
		MobileBusiness = 14
		PhoneBusiness = 15
		FaxBusiness = 16
		PHONE_END = 17

		EMAIL_BEGIN = 20
		EMailPrivate = 21
		EMailBusiness = 22
		EMAIL_END = 23
	End Enum

	Public Shared Function GetTypeDesc(nType As enType)
		Select Case nType
			Case enType.MobilePrivate : Return "Mobile priv."
			Case enType.PhonePrivate : Return "Phone priv."
			Case enType.FaxPrivate : Return "Fax priv."
			Case enType.MobileBusiness : Return "Mobile biz."
			Case enType.PhoneBusiness : Return "Phone biz."
			Case enType.FaxBusiness : Return "Fax biz."

			Case enType.EMailPrivate : Return "E-Mail priv."
			Case enType.EMailBusiness : Return "E-Mail biz."

			Case Else : Return "Unknown"
		End Select
	End Function

	Public Class clChannel
		Public Type As enType
		Public Contact As String

		Public Sub New(strContact As String, nType As enType)
			Contact = strContact.Trim().Replace("-", "").Replace(" ", "")
			Type = nType
		End Sub

		Public ReadOnly Property IsPhone() As Boolean
			Get
				Return Type > enType.PHONE_BEGIN And Type < enType.PHONE_END
			End Get
		End Property

		Public ReadOnly Property IsEMail As Boolean
			Get
				Return Type > enType.EMAIL_BEGIN And Type < enType.EMAIL_END
			End Get
		End Property
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
