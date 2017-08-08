Public Class frmContactList

	Private mnChannelOffset As Integer = 0
	Private moChannels As New Generic.List(Of clContact.enType)

	Public Sub New(oContacts As clContactList)
		MyBase.New()

		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		dgvContacts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
		dgvContacts.MultiSelect = False
		dgvContacts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
		dgvContacts.EditMode = DataGridViewEditMode.EditOnEnter

		For Each oContact In oContacts
			AddContact(oContact)
		Next

		dgvContacts.Refresh()

	End Sub

	Private Sub AddContact(oContact As clContact)

		If mnChannelOffset = 0 Then
			dgvContacts.Columns.Add("DisplayName", "Display Name")
			dgvContacts.Columns.Add("Prefix", "Prefix")
			dgvContacts.Columns.Add("FirstName", "First Name")
			dgvContacts.Columns.Add("MiddleName", "Middle Name")
			dgvContacts.Columns.Add("LastName", "Last Name")
			mnChannelOffset = dgvContacts.ColumnCount
		End If

		dgvContacts.Rows.Add(oContact.DisplayName, oContact.Prefix, oContact.FirstName, oContact.MiddleName, oContact.LastName)

		For i As Integer = 0 To oContact.Channels.Count - 1
			Dim oChannel As clContact.clChannel = oContact.Channels.ElementAt(i)
			Dim nCol As Integer = moChannels.IndexOf(oChannel.Type)

			If nCol < 0 Then nCol = InsertColumn(oChannel.Type)

			dgvContacts.Rows(dgvContacts.Rows.Count - 2).Cells(mnChannelOffset + nCol).Value = oChannel.Contact
		Next
	End Sub

	Private Function InsertColumn(nType As clContact.enType) As Integer
		Dim oCol As New DataGridViewColumn(dgvContacts.Rows(0).Cells(0))
		Dim nCol As Integer = dgvContacts.ColumnCount - mnChannelOffset

		oCol.Name = "Channel_" & CInt(nType)
		oCol.HeaderText = clContact.GetTypeDesc(nType)

		For i As Integer = 0 To moChannels.Count - 1
			If nType < moChannels.Item(i) Then
				nCol = i
				Exit For
			End If
		Next

		moChannels.Insert(nCol, nType)
		dgvContacts.Columns.Insert(mnChannelOffset + nCol, oCol)

		Return nCol
	End Function

End Class
