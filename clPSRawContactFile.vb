Imports System.IO
Imports System.Xml

Public Class clPSRawContactFile
	Inherits clContactFile

	Public Sub New(strFile As String, oLogger As Logger)
		MyBase.New(strFile, oLogger)
	End Sub

	Protected Overrides Function ReadFile() As clContactList
		Dim oContacts As New clContactList()
		Dim oStream As FileStream = File.OpenRead(FileName)
		Dim oReader As XmlReader = XmlReader.Create(oStream)

		While oReader.Read()
			Select Case oReader.NodeType
				Case XmlNodeType.Element
					If IsAnyOf(oReader.Name, "ws:PSRawContact") Then
						oContacts.Add(ReadContact(oReader))
					End If
			End Select
		End While

		Return oContacts
	End Function

	Private Function ReadContact(oReader As XmlReader) As clContact
		Dim oContact As New clContact(FileName)
		Dim nElementStack As Integer = 1
		Dim strCurrentNode As String = ""

		While oReader.Read()
			Select Case oReader.NodeType
				Case XmlNodeType.Element
					nElementStack += 1
					If IsAnyOf(strCurrentNode, "ws:metadata") Then
						' The metadata is boring -> Skip sub nodes
					ElseIf IsAnyOf(strCurrentNode, "ws:phones", "ws:emails") Then
						If IsAnyOf(oReader.Name, "ws:PSPhone", "ws:PSEmail") Then
							oContact.AddChannel(ReadChannel(oReader, oReader.Name))
							nElementStack -= 1
						End If
					Else
						strCurrentNode = oReader.Name.ToLower()

						Select Case strCurrentNode
							Case "ws:metadata" ' The metadata is boring -> Skip node
							Case "ws:phones" ' The phones are indeed very interesting -> Parse sub nodes
							Case "ws:emails" ' Same for e-mails
							Case "ws:structuredname"
								ReadName(oReader, oContact)
								nElementStack -= 1
							Case Else
								Throw New Exception("clPSRawContactFile.ReadContact(): Unsupported node '" & strCurrentNode & "'!")
						End Select
					End If
				Case XmlNodeType.EndElement
					nElementStack -= 1
					If nElementStack = 1 Then
						strCurrentNode = ""
					ElseIf nElementStack = 0 Then
						If IsAnyOf(oReader.Name, "ws:PSRawContact") Then
							Exit While
						Else
							Throw New Exception("clPSRawContactFile.ReadContact(): Ending node tag </ws:PSRawContact> expected!")
						End If
					End If
			End Select
		End While

		Return oContact
	End Function

	Private Function GetChannelType(strNode As String, strType As String) As clContact.enType

		If IsNumeric(strType) AndAlso IsAnyOf(strNode, "ws:itype") Then
			Select Case CInt(strType)
				Case 1 : Return clContact.enType.PhonePrivate
				Case 2 : Return clContact.enType.MobilePrivate
				Case 3 : Return clContact.enType.PhoneBusiness
			End Select
		ElseIf IsNumeric(strType) AndAlso IsAnyOf(strNode, "ws:iemailtype") Then
			Select Case CInt(strType)
				Case 0 : Return clContact.enType.EMailPrivate
				Case 1 : Return clContact.enType.EMailOther
				Case 2 : Return clContact.enType.EMailBusiness
			End Select
		End If

		Throw New Exception("clPSRawContactFile.GetChannelType(): Unsupported type '" & strNode & "/" & strType & "'!")
	End Function

	Private Function ReadChannel(oReader As XmlReader, strNode As String) As clContact.clChannel
		Dim strContact As String = ""
		Dim nType As clContact.enType = clContact.enType.Unknown
		Dim nElementStack As Integer = 1
		Dim strCurrentNode As String = ""

		While oReader.Read()
			Select Case oReader.NodeType
				Case XmlNodeType.Element
					nElementStack += 1
					strCurrentNode = oReader.Name.ToLower()
				Case XmlNodeType.Text
					If IsAnyOf(strCurrentNode, "ws:itype", "ws:iemailtype") Then
						nType = GetChannelType(strCurrentNode, oReader.Value.Trim())
					ElseIf IsAnyOf(strCurrentNode, "ws:snumber", "ws:semailaddr") Then
						strContact = oReader.Value.Trim()
					End If
				Case XmlNodeType.EndElement
					nElementStack -= 1
					If nElementStack = 0 Then
						If IsAnyOf(oReader.Name, "ws:PSPhone", "ws:PSEmail") Then
							Exit While
						Else
							Throw New Exception("clPSRawContactFile.ReadPhone(): Ending node tag </ws:PSPhone> expected!")
						End If
					End If
			End Select
		End While

		Return New clContact.clChannel(strContact, nType)
	End Function

	Private Sub ReadName(oReader As XmlReader, oContact As clContact)
		Dim nElementStack As Integer = 1
		Dim strCurrentNode As String = ""

		While oReader.Read()
			Select Case oReader.NodeType
				Case XmlNodeType.Element
					nElementStack += 1
					strCurrentNode = oReader.Name.ToLower()
				Case XmlNodeType.Text
					Select Case strCurrentNode
						Case "ws:sdisplayname" : oContact.DisplayName = oReader.Value.Trim()
						Case "ws:sgivenname" : oContact.FirstName = oReader.Value.Trim()
						Case "ws:smiddlename" : oContact.MiddleName = oReader.Value.Trim()
						Case "ws:sfamilyname" : oContact.LastName = oReader.Value.Trim()
						Case "ws:sdisplayname" : oContact.DisplayName = oReader.Value.Trim()
						Case "ws:_id", "ws:idataversion", "ws:iisprimary", "ws:iissuperprimary", "ws:lrawcontactid"
							' Boring
						Case Else
							Throw New Exception("clPSRawContactFile.ReadName(): Unsupported node '" & strCurrentNode & "'!")
					End Select
				Case XmlNodeType.EndElement
					nElementStack -= 1
					If nElementStack = 0 Then
						If IsAnyOf(oReader.Name, "ws:structuredName") Then
							Exit While
						Else
							Throw New Exception("clPSRawContactFile.ReadName(): Ending node tag </ws:structuredName> expected!")
						End If
					End If
			End Select
		End While

	End Sub

End Class
