Public Class frmContactList

	Private mnChannelOffset As Integer = 0
	Private moChannels As New List(Of clContact.enType)
	Private moDuplicates As New Dictionary(Of String, Integer)

	Private COLOR_HIGHLIGHT As System.Drawing.Color = Color.Red
	Private COLOR_DOUBLEHIGHLIGHT As System.Drawing.Color = Color.Orange

	Public Sub New(oContacts As clContactList)
		MyBase.New()

		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		For nPhoneType As clContact.enType = clContact.enType.PHONE_BEGIN + 1 To clContact.enType.PHONE_END - 1
			cboPhoneTypes.Items.Add(clContact.GetTypeDesc(nPhoneType))
		Next
		If cboPhoneTypes.Items.Count > 0 Then cboPhoneTypes.SelectedIndex = 0

		For nEMailType As clContact.enType = clContact.enType.EMAIL_BEGIN + 1 To clContact.enType.EMAIL_END - 1
			cboEMailTypes.Items.Add(clContact.GetTypeDesc(nEMailType))
		Next
		If cboEMailTypes.Items.Count > 0 Then cboEMailTypes.SelectedIndex = 0

		dgvContacts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
		dgvContacts.MultiSelect = False
		dgvContacts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells 'DataGridViewAutoSizeColumnsMode.Fill
		dgvContacts.EditMode = DataGridViewEditMode.EditOnEnter

		' Respect column order at all places!!!
		dgvContacts.Columns.Add("DisplayName", "Display Name")
		dgvContacts.Columns.Add("Prefix", "Prefix")
		dgvContacts.Columns.Add("FirstName", "First Name")
		dgvContacts.Columns.Add("MiddleName", "Middle Name")
		dgvContacts.Columns.Add("LastName", "Last Name")
		dgvContacts.Columns.Add("Suffix", "Suffix")
		dgvContacts.Columns.Add("Orga", "Company")
		dgvContacts.Columns.Add("Title", "Title")
		mnChannelOffset = dgvContacts.ColumnCount
		dgvContacts.Columns.Add("File", "File Name")

		If oContacts IsNot Nothing Then
			For Each oContact In oContacts
				AddContact(oContact)
			Next
		End If

		dgvContacts.Refresh()

	End Sub

	Private Shared Function StyleWithColor(oCell As DataGridViewCell, nColor As System.Drawing.Color, bSetBackColor As Boolean) As DataGridViewCellStyle
		Dim oStyle As DataGridViewCellStyle = oCell.Style

		If bSetBackColor Then
			oStyle.BackColor = nColor
		Else
			oStyle.ForeColor = nColor
		End If

		Return oStyle
	End Function

	Private Sub AddContact(oContact As clContact)

		' Respect column order at all places!!!
		Dim nRow As Integer = dgvContacts.Rows.Add(oContact.DisplayName, oContact.Prefix, oContact.FirstName, oContact.MiddleName, oContact.LastName, oContact.Suffix, oContact.Organization, oContact.Title)

		For i As Integer = 0 To oContact.Channels.Count - 1
			Dim oChannel As clContact.clChannel = oContact.Channels.ElementAt(i)
			Dim nCol As Integer = moChannels.IndexOf(oChannel.Type) : If nCol < 0 Then nCol = InsertColumn(oChannel.Type)
			Dim oCell As DataGridViewCell = dgvContacts.Rows(nRow).Cells(mnChannelOffset + nCol)

			If String.IsNullOrEmpty(oCell.Value) Then
				oCell.Value = oChannel.Contact
			Else
				oCell.Style = StyleWithColor(oCell, COLOR_HIGHLIGHT, True)
				oCell.Value = oCell.Value & "," & oChannel.Contact
				chkFilterMultiUsage.Enabled = True
				chkFilterMultiUsage.BackColor = COLOR_HIGHLIGHT
			End If
		Next

		Dim oFileCell As DataGridViewCell = dgvContacts.Rows(nRow).Cells(mnChannelOffset + moChannels.Count)
		oFileCell.Value = System.IO.Path.GetFileName(oContact.FileName)
		oFileCell.Style = StyleWithColor(oFileCell, Color.Gray, False)

		HighlightIfDuplicate(oContact, nRow)

	End Sub

	Private Sub HighlightIfDuplicate(oContact As clContact, nRow As Integer)
		Dim oRow As DataGridViewRow = dgvContacts.Rows.Item(nRow)

		If Not String.IsNullOrEmpty(oContact.LastName) Then HighlightIfDuplicateHelper(oContact.LastName & "_" & oContact.FirstName, oRow)

		HighlightIfDuplicateHelper(oContact.DisplayName, oRow)

		For Each oChannel As clContact.clChannel In oContact.Channels
			HighlightIfDuplicateHelper(oChannel.Contact, oRow)
		Next

	End Sub

	Private Sub HighlightIfDuplicateHelper(strKey As String, oRow As DataGridViewRow)

		strKey = strKey.Trim().ToLower()
		If Not String.IsNullOrEmpty(strKey) AndAlso moDuplicates.ContainsKey(strKey) Then
			HighlightRow(oRow)
			HighlightRow(dgvContacts.Rows.Item(moDuplicates.Item(strKey)))
			chkFilterDuplicates.Enabled = True
			chkFilterDuplicates.ForeColor = COLOR_HIGHLIGHT
		ElseIf Not String.IsNullOrEmpty(strKey) Then
			moDuplicates.Add(strKey, oRow.Index)
		End If

	End Sub

	Private Sub HighlightRow(oRow As DataGridViewRow)

		For Each oCell As DataGridViewCell In oRow.Cells
			If oCell.Style.BackColor = COLOR_HIGHLIGHT Then
				oCell.Style = StyleWithColor(oCell, COLOR_DOUBLEHIGHLIGHT, False)
			Else
				oCell.Style = StyleWithColor(oCell, COLOR_HIGHLIGHT, False)
			End If
		Next

	End Sub

	Private Sub FilterRows()

		For Each oRow As DataGridViewRow In dgvContacts.Rows
			Dim bVisible As Boolean = Not chkFilterMultiUsage.Checked AndAlso Not chkFilterDuplicates.Checked

			If chkFilterMultiUsage.Checked And Not bVisible Then
				For i As Integer = mnChannelOffset To mnChannelOffset + moChannels.Count - 1
					If oRow.Cells(i).Style.BackColor = COLOR_HIGHLIGHT Then bVisible = True : Exit For
				Next
			End If

			If chkFilterDuplicates.Checked And Not bVisible Then
				If oRow.Cells(0).Style.ForeColor = COLOR_HIGHLIGHT Then bVisible = True
			End If

			If Not oRow.IsNewRow Then oRow.Visible = bVisible
		Next

	End Sub

	Private Function InsertColumn(nType As clContact.enType) As Integer
		Dim oCol As New DataGridViewColumn(dgvContacts.Rows(0).Cells(0))
		Dim nCol As Integer = moChannels.Count ' dgvContacts.ColumnCount - mnChannelOffset

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

		If clContact.IsPhone(nType) Then
			CheckAllChannelsAdded(nType, cboPhoneTypes, cmdAddPhone)
		ElseIf clContact.IsEMail(nType) Then
			CheckAllChannelsAdded(nType, cboEMailTypes, cmdAddEMail)
		End If

		Return nCol
	End Function

	Private Sub CheckAllChannelsAdded(nType As clContact.enType, cboTypes As ComboBox, cmdAdd As Button)

		For Each oItem As Object In cboTypes.Items
			If oItem.ToString().Equals(clContact.GetTypeDesc(nType)) Then cboTypes.Items.Remove(oItem) : Exit For
		Next

		If cboTypes.Items.Count > 0 Then
			cboTypes.SelectedIndex = 0
		Else
			cboTypes.SelectedItem = Nothing
			cmdAdd.Enabled = False
		End If

	End Sub

	Private Sub cmdAddPhone_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddPhone.Click
		If Not String.IsNullOrEmpty(cboPhoneTypes.SelectedItem.ToString()) Then InsertColumn(clContact.GetTypeByDesc(cboPhoneTypes.SelectedItem.ToString()))
	End Sub

	Private Sub cmdAddEMail_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddEMail.Click
		If Not String.IsNullOrEmpty(cboEMailTypes.SelectedItem.ToString()) Then InsertColumn(clContact.GetTypeByDesc(cboEMailTypes.SelectedItem.ToString()))
	End Sub

	Private Sub dgvContacts_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dgvContacts.RowPostPaint
		Using oBrush As SolidBrush = New SolidBrush(dgvContacts.RowHeadersDefaultCellStyle.ForeColor)
			If e.RowIndex < dgvContacts.Rows.Count - 1 Then
				e.Graphics.DrawString((e.RowIndex + 1).ToString(), dgvContacts.DefaultCellStyle.Font, oBrush, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4)
			End If
		End Using
	End Sub

	Private Sub chkFilterMultiUsage_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFilterMultiUsage.CheckedChanged
		FilterRows()
	End Sub

	Private Sub chkFilterDuplicates_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFilterDuplicates.CheckedChanged
		FilterRows()
	End Sub

	Private Sub cmdRemove_Click(sender As System.Object, e As System.EventArgs) Handles cmdRemove.Click

		For Each oRow As DataGridViewRow In dgvContacts.SelectedRows
			Dim oOldDuplicates As Dictionary(Of String, Integer) = moDuplicates

			If oRow.IsNewRow Then Continue For

			moDuplicates = New Dictionary(Of String, Integer)
			For Each kvp As KeyValuePair(Of String, Integer) In oOldDuplicates
				moDuplicates.Add(kvp.Key, If(kvp.Value > oRow.Index, kvp.Value - 1, kvp.Value))
			Next
			dgvContacts.Rows.Remove(oRow)
		Next

	End Sub

	Private Sub cmdUnlightMultiUsage_Click(sender As System.Object, e As System.EventArgs) Handles cmdUnlightMultiUsage.Click
		Dim nColor As System.Drawing.Color = dgvContacts.Rows.Item(dgvContacts.Rows.Count - 1).Cells(0).Style.BackColor

		For Each oRow As DataGridViewRow In dgvContacts.SelectedRows
			For i As Integer = mnChannelOffset To mnChannelOffset + moChannels.Count - 1
				oRow.Cells(i).Style = StyleWithColor(oRow.Cells(i), nColor, True)
			Next
		Next

	End Sub

	Private Sub cmdUnlightDuplicate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUnlightDuplicate.Click
		Dim nColor As System.Drawing.Color = dgvContacts.Rows.Item(dgvContacts.Rows.Count - 1).Cells(0).Style.ForeColor

		For Each oRow As DataGridViewRow In dgvContacts.SelectedRows
			For Each oCell In oRow.Cells
				oCell.Style = StyleWithColor(oCell, nColor, False)
			Next
		Next

	End Sub

	Private Sub dgvContacts_SelectionChanged(sender As Object, e As System.EventArgs) Handles dgvContacts.SelectionChanged
		Dim bAllowUnlightMultiUsage As Boolean = False
		Dim bAllowUnlightDuplicate As Boolean = False

		For Each oRow As DataGridViewRow In dgvContacts.SelectedRows
			For i As Integer = mnChannelOffset To mnChannelOffset + moChannels.Count - 1
				If oRow.Cells(i).Style.BackColor = COLOR_HIGHLIGHT Then bAllowUnlightMultiUsage = True : Exit For
			Next

			If oRow.Cells(0).Style.ForeColor = COLOR_HIGHLIGHT Then bAllowUnlightDuplicate = True
		Next

		cmdUnlightMultiUsage.Enabled = bAllowUnlightMultiUsage
		cmdUnlightDuplicate.Enabled = bAllowUnlightDuplicate
	End Sub

End Class
