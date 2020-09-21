Public Class frmContactList

	Private Shared snListNr As Integer = 0

	Private mnChannelOffset As Integer = 0
	Private moChannels As New List(Of clContact.enType)
	Private moDuplicates As New Dictionary(Of String, Integer)
	Private moLogger As Logger

	Private COLOR_HIGHLIGHT As System.Drawing.Color = Color.Red
	Private COLOR_DOUBLEHIGHLIGHT As System.Drawing.Color = Color.Orange

	Public Sub New(oContacts As clContactList, oLogger As Logger)
		MyBase.New()

		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		snListNr += 1
		Text = Text & " " & snListNr

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
		dgvContacts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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

		moLogger = oLogger
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

			If IsEmpty(oCell.Value) Then
				oCell.Value = oChannel.Contact
			Else
				oCell.Style = StyleWithColor(oCell, COLOR_HIGHLIGHT, True)
				oCell.Value = oCell.Value & "," & oChannel.Contact
				chkFilterMultiUsage.Enabled = True
				chkFilterMultiUsage.BackColor = COLOR_HIGHLIGHT
			End If

			If oChannel.Primary Then oCell.Style.Font = New Font(DataGridView.DefaultFont, FontStyle.Bold)
		Next

		Dim oFileCell As DataGridViewCell = dgvContacts.Rows(nRow).Cells(mnChannelOffset + moChannels.Count)
		oFileCell.Value = System.IO.Path.GetFileName(oContact.FileName)
		oFileCell.Style = StyleWithColor(oFileCell, Color.Gray, False)

		HighlightIfDuplicate(oContact, nRow)

	End Sub

	Private Sub HighlightIfDuplicate(oContact As clContact, nRow As Integer)
		Dim oRow As DataGridViewRow = dgvContacts.Rows.Item(nRow)

		If IsNotEmpty(oContact.LastName) Then HighlightIfDuplicateHelper(oContact.LastName & "_" & oContact.FirstName, oRow)

		HighlightIfDuplicateHelper(oContact.DisplayName, oRow)

		For Each oChannel As clContact.clChannel In oContact.Channels
			HighlightIfDuplicateHelper(oChannel.Contact, oRow)
		Next

	End Sub

	Private Sub HighlightIfDuplicateHelper(strKey As String, oRow As DataGridViewRow)

		strKey = strKey.Trim().ToLower()
		If IsNotEmpty(strKey) AndAlso moDuplicates.ContainsKey(strKey) Then
			HighlightRow(oRow)
			HighlightRow(dgvContacts.Rows.Item(moDuplicates.Item(strKey)))
			chkFilterDuplicates.Enabled = True
			chkFilterDuplicates.ForeColor = COLOR_HIGHLIGHT
		ElseIf IsNotEmpty(strKey) Then
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
		Dim strFilter As String = txtSearch.Text.Trim().ToLower()

		fraSearch.Enabled = False

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

			If bVisible AndAlso IsNotEmpty(strFilter) Then
				bVisible = False
				For Each oCell As DataGridViewCell In oRow.Cells
					For Each strSearch As String In strFilter.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
						If oCell.Value IsNot Nothing AndAlso oCell.Value.ToString().ToLower().Contains(strSearch) Then bVisible = True : Exit For
					Next
				Next
			End If

			cmdSearch.Text = (oRow.Index + 1) & " / " & dgvContacts.RowCount
			cmdSearch.Refresh()
			If Not oRow.IsNewRow Then oRow.Visible = bVisible
		Next

		cmdSearch.Text = "&Search"
		fraSearch.Enabled = True

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
		If IsNotEmpty(cboPhoneTypes.SelectedItem.ToString()) Then InsertColumn(clContact.GetTypeByDesc(cboPhoneTypes.SelectedItem.ToString()))
	End Sub

	Private Sub cmdAddEMail_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddEMail.Click
		If IsNotEmpty(cboEMailTypes.SelectedItem.ToString()) Then InsertColumn(clContact.GetTypeByDesc(cboEMailTypes.SelectedItem.ToString()))
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

	Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdSearch.Click
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

	Private Sub cmdExportCSV_Click(sender As Object, e As EventArgs) Handles cmdExportCSV.Click
		Dim frm As New SaveFileDialog()

		If frm.ShowDialog(Me) = DialogResult.OK AndAlso IsNotEmpty(frm.FileName) Then
			Try
				Dim oWriter As New System.IO.StreamWriter(frm.FileName)

				For Each oRow As DataGridViewRow In dgvContacts.Rows
					If oRow.Visible AndAlso Not oRow.IsNewRow Then
						Dim strLine As New System.Text.StringBuilder(If(oRow.Cells(0).Value, "").ToString().Trim())

						For i As Integer = 1 To dgvContacts.ColumnCount - 1
							Dim oCell As DataGridViewCell = oRow.Cells(i)

							strLine.Append(";"c)
							strLine.Append(If(oRow.Cells(i).Value, "").ToString().Trim().Replace(";"c, ","))
						Next

						oWriter.WriteLine(strLine.ToString())
					End If
				Next

				oWriter.Close()
			Catch ex As Exception
				clContactFile.LogException(moLogger, ex, frm.FileName)
			End Try
		End If

	End Sub

	Private Sub cmdExportVCF_Click(sender As Object, e As EventArgs) Handles cmdExportVCF.Click
		Dim frm As New FolderBrowserDialog()

		If frm.ShowDialog(Me) = DialogResult.OK AndAlso IsNotEmpty(frm.SelectedPath) Then
			For Each oRow As DataGridViewRow In dgvContacts.Rows
				If oRow.Visible AndAlso Not oRow.IsNewRow Then
					Dim oContact As clContact = GetContactFromRow(oRow)
					Dim oContacts As New clContactList() : oContacts.Add(oContact)
					Dim oFile As New clVCardFile(oContacts)
					Dim strFile As String = If(IsEmpty(oContact.FileName), oContact.DisplayName, System.IO.Path.GetFileNameWithoutExtension(oContact.FileName)) & ".vcf"

					Try
						oFile.WriteFile(frm.SelectedPath & System.IO.Path.DirectorySeparatorChar & strFile)
					Catch ex As Exception
						clContactFile.LogException(moLogger, ex, strFile)
					End Try
				End If
			Next
		End If

	End Sub

	Private Function GetContactFromRow(oRow As DataGridViewRow) As clContact
		Dim oContact As New clContact(oRow.Cells(oRow.Cells.Count - 1).Value)

		' Respect column order at all places!!!
		If oRow.Cells(0).Value IsNot Nothing Then oContact.DisplayName = oRow.Cells(0).Value.ToString().Trim()
		If oRow.Cells(1).Value IsNot Nothing Then oContact.Prefix = oRow.Cells(1).Value.ToString().Trim()
		If oRow.Cells(2).Value IsNot Nothing Then oContact.FirstName = oRow.Cells(2).Value.ToString().Trim()
		If oRow.Cells(3).Value IsNot Nothing Then oContact.MiddleName = oRow.Cells(3).Value.ToString().Trim()
		If oRow.Cells(4).Value IsNot Nothing Then oContact.LastName = oRow.Cells(4).Value.ToString().Trim()
		If oRow.Cells(5).Value IsNot Nothing Then oContact.Suffix = oRow.Cells(5).Value.ToString().Trim()
		If oRow.Cells(6).Value IsNot Nothing Then oContact.Organization = oRow.Cells(6).Value.ToString().Trim()
		If oRow.Cells(7).Value IsNot Nothing Then oContact.Title = oRow.Cells(7).Value.ToString().Trim()

		For i As Integer = 0 To moChannels.Count - 1
			Dim oCell As DataGridViewCell = oRow.Cells(mnChannelOffset + i)

			If IsNotEmpty(oCell.Value) Then
				Dim strChannel As String = oCell.Value.ToString().Trim()
				Dim nType As clContact.enType = moChannels(i)
				Dim bPrimary As Boolean = oCell.Style.Font IsNot Nothing AndAlso oCell.Style.Font.Bold

				For Each strContact As String In strChannel.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
					strContact = strContact.Trim()

					If clContact.IsPhone(nType) AndAlso strContact.Length > 2 AndAlso strContact.StartsWith("0") AndAlso strContact.Chars(1) <> "0"c Then
						strContact = "+49" & strContact.Substring(1)
					End If

					oContact.AddChannel(New clContact.clChannel(strContact, nType, bPrimary))
				Next
			End If
		Next

		Return oContact
	End Function

End Class
