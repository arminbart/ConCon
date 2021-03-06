﻿Public Class clContact

	Private mstrDisplayName As String

	Public Property Prefix As String
	Public Property FirstName As String
	Public Property MiddleName As String
	Public Property LastName As String
	Public Property Suffix As String

	Public Property Organization As String
	Public Property Title As String

	Public Property Channels As New Generic.LinkedList(Of clContact.clChannel)()

	Public Property FileName As String

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
		PhoneMain = 20 ' Main/Festnetz
		PHONE_END = 21

		EMAIL_BEGIN = 30
		EMailPrivate = 31
		EMailBusiness = 32
		EMailOther = 32
		EMAIL_END = 33

		ADDRESS_BEGIN = 50
		AddressPrivate = 51
		AddressBusiness = 52
		AddressOther = 53
		ADDRESS_END = 54
	End Enum

	Public Property DisplayName() As String
		Get
			If IsEmpty(mstrDisplayName) Then
				mstrDisplayName = If(IsNotEmpty(FirstName), FirstName, "")
				If IsNotEmpty(MiddleName) Then mstrDisplayName &= If(IsNotEmpty(mstrDisplayName), " ", "") & MiddleName
				If IsNotEmpty(LastName) Then mstrDisplayName &= If(IsNotEmpty(mstrDisplayName), " ", "") & LastName
				If IsNotEmpty(Suffix) Then mstrDisplayName &= If(IsNotEmpty(mstrDisplayName), ", ", "") & Suffix
				mstrDisplayName = mstrDisplayName.Trim()
			End If

			Return mstrDisplayName
		End Get
		Set(value As String)
			mstrDisplayName = value
		End Set
	End Property

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
			Case enType.PhoneMain : Return "Phone main"

			Case enType.EMailPrivate : Return "E-Mail priv."
			Case enType.EMailBusiness : Return "E-Mail biz."
			Case enType.EMailOther : Return "E-Mail oth."

			Case enType.AddressPrivate : Return "Address priv."
			Case enType.AddressBusiness : Return "Address biz."
			Case enType.AddressOther : Return "Address oth."

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

	Public Shared Function IsAddress(nType As enType)
		Return nType > enType.ADDRESS_BEGIN And nType < enType.ADDRESS_END
	End Function

	Public Class clChannel
		Public Type As enType
		Public Contact As String
		Public Primary As Boolean

		Public Sub New(strContact As String, nType As enType, bPrimary As Boolean)
			strContact = strContact.Trim().Replace(" ", "").ToLower()
			Contact &= If(IsPhone(nType), FormatPhone(strContact), strContact)
			Type = nType
			Primary = bPrimary
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

	Public Class clAddress
		Inherits clChannel

		Public Street As String
		Public Zip As String
		Public City As String
		Public Country As String

		Public Sub New(strContact As String, nType As enType)
			MyBase.New(strContact, nType, False)
		End Sub
	End Class

	Public Sub AddChannel(oChannel As clChannel)
		Channels.AddLast(oChannel)
	End Sub

End Class
