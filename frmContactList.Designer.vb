<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmContactList
	Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.dgvContacts = New System.Windows.Forms.DataGridView()
		Me.cmdAddPhone = New System.Windows.Forms.Button()
		Me.cmdAddEMail = New System.Windows.Forms.Button()
		Me.cboPhoneTypes = New System.Windows.Forms.ComboBox()
		Me.cboEMailTypes = New System.Windows.Forms.ComboBox()
		Me.chkFilterDuplicates = New System.Windows.Forms.CheckBox()
		Me.chkFilterMultiUsage = New System.Windows.Forms.CheckBox()
		Me.cmdRemove = New System.Windows.Forms.Button()
		Me.cmdUnlightMultiUsage = New System.Windows.Forms.Button()
		Me.cmdUnlightDuplicate = New System.Windows.Forms.Button()
		Me.txtSearch = New System.Windows.Forms.TextBox()
		Me.cmdSearch = New System.Windows.Forms.Button()
		Me.cmdExportCSV = New System.Windows.Forms.Button()
		Me.cmdExportVCF = New System.Windows.Forms.Button()
		Me.fraSearch = New System.Windows.Forms.GroupBox()
		Me.fraContacts = New System.Windows.Forms.GroupBox()
		Me.fraChannels = New System.Windows.Forms.GroupBox()
		CType(Me.dgvContacts, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.fraSearch.SuspendLayout()
		Me.fraContacts.SuspendLayout()
		Me.fraChannels.SuspendLayout()
		Me.SuspendLayout()
		'
		'dgvContacts
		'
		Me.dgvContacts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.dgvContacts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgvContacts.Location = New System.Drawing.Point(12, 91)
		Me.dgvContacts.Name = "dgvContacts"
		Me.dgvContacts.Size = New System.Drawing.Size(1160, 580)
		Me.dgvContacts.TabIndex = 0
		'
		'cmdAddPhone
		'
		Me.cmdAddPhone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdAddPhone.Location = New System.Drawing.Point(8, 42)
		Me.cmdAddPhone.Name = "cmdAddPhone"
		Me.cmdAddPhone.Size = New System.Drawing.Size(135, 23)
		Me.cmdAddPhone.TabIndex = 1
		Me.cmdAddPhone.Text = "Add &Phone Column"
		Me.cmdAddPhone.UseVisualStyleBackColor = True
		'
		'cmdAddEMail
		'
		Me.cmdAddEMail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdAddEMail.Location = New System.Drawing.Point(149, 42)
		Me.cmdAddEMail.Name = "cmdAddEMail"
		Me.cmdAddEMail.Size = New System.Drawing.Size(135, 23)
		Me.cmdAddEMail.TabIndex = 2
		Me.cmdAddEMail.Text = "Add E-&Mail Column"
		Me.cmdAddEMail.UseVisualStyleBackColor = True
		'
		'cboPhoneTypes
		'
		Me.cboPhoneTypes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cboPhoneTypes.FormattingEnabled = True
		Me.cboPhoneTypes.Location = New System.Drawing.Point(8, 17)
		Me.cboPhoneTypes.Name = "cboPhoneTypes"
		Me.cboPhoneTypes.Size = New System.Drawing.Size(135, 21)
		Me.cboPhoneTypes.TabIndex = 3
		'
		'cboEMailTypes
		'
		Me.cboEMailTypes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cboEMailTypes.FormattingEnabled = True
		Me.cboEMailTypes.Location = New System.Drawing.Point(149, 17)
		Me.cboEMailTypes.Name = "cboEMailTypes"
		Me.cboEMailTypes.Size = New System.Drawing.Size(135, 21)
		Me.cboEMailTypes.TabIndex = 4
		'
		'chkFilterDuplicates
		'
		Me.chkFilterDuplicates.AutoSize = True
		Me.chkFilterDuplicates.Enabled = False
		Me.chkFilterDuplicates.Location = New System.Drawing.Point(184, 48)
		Me.chkFilterDuplicates.Name = "chkFilterDuplicates"
		Me.chkFilterDuplicates.Size = New System.Drawing.Size(128, 17)
		Me.chkFilterDuplicates.TabIndex = 5
		Me.chkFilterDuplicates.Text = "Show only &Duplicates"
		Me.chkFilterDuplicates.UseVisualStyleBackColor = True
		'
		'chkFilterMultiUsage
		'
		Me.chkFilterMultiUsage.AutoSize = True
		Me.chkFilterMultiUsage.Enabled = False
		Me.chkFilterMultiUsage.Location = New System.Drawing.Point(184, 21)
		Me.chkFilterMultiUsage.Name = "chkFilterMultiUsage"
		Me.chkFilterMultiUsage.Size = New System.Drawing.Size(173, 17)
		Me.chkFilterMultiUsage.TabIndex = 6
		Me.chkFilterMultiUsage.Text = "Show only &Multi Usage of fields"
		Me.chkFilterMultiUsage.UseVisualStyleBackColor = True
		'
		'cmdRemove
		'
		Me.cmdRemove.Location = New System.Drawing.Point(6, 42)
		Me.cmdRemove.Name = "cmdRemove"
		Me.cmdRemove.Size = New System.Drawing.Size(135, 23)
		Me.cmdRemove.TabIndex = 7
		Me.cmdRemove.Text = "&Remove Contact"
		Me.cmdRemove.UseVisualStyleBackColor = True
		'
		'cmdUnlightMultiUsage
		'
		Me.cmdUnlightMultiUsage.Location = New System.Drawing.Point(363, 17)
		Me.cmdUnlightMultiUsage.Name = "cmdUnlightMultiUsage"
		Me.cmdUnlightMultiUsage.Size = New System.Drawing.Size(229, 23)
		Me.cmdUnlightMultiUsage.TabIndex = 8
		Me.cmdUnlightMultiUsage.Text = "Unhighlight Multi Usage of selected Contact"
		Me.cmdUnlightMultiUsage.UseVisualStyleBackColor = True
		'
		'cmdUnlightDuplicate
		'
		Me.cmdUnlightDuplicate.Location = New System.Drawing.Point(363, 44)
		Me.cmdUnlightDuplicate.Name = "cmdUnlightDuplicate"
		Me.cmdUnlightDuplicate.Size = New System.Drawing.Size(229, 23)
		Me.cmdUnlightDuplicate.TabIndex = 9
		Me.cmdUnlightDuplicate.Text = "Unhighlight selected Duplicate Contact"
		Me.cmdUnlightDuplicate.UseVisualStyleBackColor = True
		'
		'txtSearch
		'
		Me.txtSearch.Location = New System.Drawing.Point(6, 19)
		Me.txtSearch.Name = "txtSearch"
		Me.txtSearch.Size = New System.Drawing.Size(135, 20)
		Me.txtSearch.TabIndex = 10
		'
		'cmdSearch
		'
		Me.cmdSearch.Location = New System.Drawing.Point(6, 44)
		Me.cmdSearch.Name = "cmdSearch"
		Me.cmdSearch.Size = New System.Drawing.Size(135, 23)
		Me.cmdSearch.TabIndex = 11
		Me.cmdSearch.Text = "&Search"
		Me.cmdSearch.UseVisualStyleBackColor = True
		'
		'cmdExportCSV
		'
		Me.cmdExportCSV.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdExportCSV.Enabled = False
		Me.cmdExportCSV.Location = New System.Drawing.Point(1037, 677)
		Me.cmdExportCSV.Name = "cmdExportCSV"
		Me.cmdExportCSV.Size = New System.Drawing.Size(135, 23)
		Me.cmdExportCSV.TabIndex = 12
		Me.cmdExportCSV.Text = "Export &CSV"
		Me.cmdExportCSV.UseVisualStyleBackColor = True
		'
		'cmdExportVCF
		'
		Me.cmdExportVCF.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cmdExportVCF.Location = New System.Drawing.Point(896, 677)
		Me.cmdExportVCF.Name = "cmdExportVCF"
		Me.cmdExportVCF.Size = New System.Drawing.Size(135, 23)
		Me.cmdExportVCF.TabIndex = 13
		Me.cmdExportVCF.Text = "Export &vCard"
		Me.cmdExportVCF.UseVisualStyleBackColor = True
		'
		'fraSearch
		'
		Me.fraSearch.Controls.Add(Me.txtSearch)
		Me.fraSearch.Controls.Add(Me.cmdSearch)
		Me.fraSearch.Controls.Add(Me.cmdUnlightDuplicate)
		Me.fraSearch.Controls.Add(Me.chkFilterDuplicates)
		Me.fraSearch.Controls.Add(Me.cmdUnlightMultiUsage)
		Me.fraSearch.Controls.Add(Me.chkFilterMultiUsage)
		Me.fraSearch.Location = New System.Drawing.Point(12, 12)
		Me.fraSearch.Name = "fraSearch"
		Me.fraSearch.Size = New System.Drawing.Size(601, 73)
		Me.fraSearch.TabIndex = 15
		Me.fraSearch.TabStop = False
		Me.fraSearch.Text = "Search and Filter"
		'
		'fraContacts
		'
		Me.fraContacts.Controls.Add(Me.cmdRemove)
		Me.fraContacts.Location = New System.Drawing.Point(619, 14)
		Me.fraContacts.Name = "fraContacts"
		Me.fraContacts.Size = New System.Drawing.Size(151, 71)
		Me.fraContacts.TabIndex = 16
		Me.fraContacts.TabStop = False
		Me.fraContacts.Text = "Add/Remove Contacts"
		'
		'fraChannels
		'
		Me.fraChannels.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.fraChannels.Controls.Add(Me.cmdAddPhone)
		Me.fraChannels.Controls.Add(Me.cmdAddEMail)
		Me.fraChannels.Controls.Add(Me.cboPhoneTypes)
		Me.fraChannels.Controls.Add(Me.cboEMailTypes)
		Me.fraChannels.Location = New System.Drawing.Point(882, 14)
		Me.fraChannels.Name = "fraChannels"
		Me.fraChannels.Size = New System.Drawing.Size(290, 71)
		Me.fraChannels.TabIndex = 17
		Me.fraChannels.TabStop = False
		Me.fraChannels.Text = "Add Communication Channels"
		'
		'frmContactList
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1184, 712)
		Me.Controls.Add(Me.fraChannels)
		Me.Controls.Add(Me.fraContacts)
		Me.Controls.Add(Me.fraSearch)
		Me.Controls.Add(Me.cmdExportVCF)
		Me.Controls.Add(Me.cmdExportCSV)
		Me.Controls.Add(Me.dgvContacts)
		Me.MinimumSize = New System.Drawing.Size(1200, 750)
		Me.Name = "frmContactList"
		Me.Text = "frmContactList"
		CType(Me.dgvContacts, System.ComponentModel.ISupportInitialize).EndInit()
		Me.fraSearch.ResumeLayout(False)
		Me.fraSearch.PerformLayout()
		Me.fraContacts.ResumeLayout(False)
		Me.fraChannels.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents dgvContacts As DataGridView
	Friend WithEvents cmdAddPhone As System.Windows.Forms.Button
	Friend WithEvents cmdAddEMail As System.Windows.Forms.Button
	Friend WithEvents cboPhoneTypes As System.Windows.Forms.ComboBox
	Friend WithEvents cboEMailTypes As System.Windows.Forms.ComboBox
	Friend WithEvents chkFilterDuplicates As System.Windows.Forms.CheckBox
	Friend WithEvents chkFilterMultiUsage As System.Windows.Forms.CheckBox
	Friend WithEvents cmdRemove As System.Windows.Forms.Button
	Friend WithEvents cmdUnlightMultiUsage As System.Windows.Forms.Button
	Friend WithEvents cmdUnlightDuplicate As System.Windows.Forms.Button
	Friend WithEvents txtSearch As TextBox
	Friend WithEvents cmdSearch As Button
	Friend WithEvents cmdExportCSV As Button
	Friend WithEvents cmdExportVCF As Button
	Friend WithEvents fraSearch As GroupBox
	Friend WithEvents fraContacts As GroupBox
	Friend WithEvents fraChannels As GroupBox
End Class
