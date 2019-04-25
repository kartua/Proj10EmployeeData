<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.lblExtension = New System.Windows.Forms.Label()
        Me.lblTelephone = New System.Windows.Forms.Label()
        Me.lblDept = New System.Windows.Forms.Label()
        Me.lblEmployeeNum = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblMiddleName = New System.Windows.Forms.Label()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblRecordNum = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(234, 390)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(107, 49)
        Me.btnClear.TabIndex = 12
        Me.btnClear.Text = "C&lear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(65, 390)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(107, 49)
        Me.btnNext.TabIndex = 11
        Me.btnNext.Text = "&Next Record"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(396, 390)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 49)
        Me.btnExit.TabIndex = 13
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'lblEmail
        '
        Me.lblEmail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEmail.Location = New System.Drawing.Point(169, 250)
        Me.lblEmail.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(311, 25)
        Me.lblEmail.TabIndex = 22
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblExtension
        '
        Me.lblExtension.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblExtension.Location = New System.Drawing.Point(169, 218)
        Me.lblExtension.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.Size = New System.Drawing.Size(311, 25)
        Me.lblExtension.TabIndex = 21
        Me.lblExtension.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTelephone
        '
        Me.lblTelephone.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelephone.Location = New System.Drawing.Point(169, 186)
        Me.lblTelephone.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTelephone.Name = "lblTelephone"
        Me.lblTelephone.Size = New System.Drawing.Size(311, 25)
        Me.lblTelephone.TabIndex = 20
        Me.lblTelephone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDept
        '
        Me.lblDept.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDept.Location = New System.Drawing.Point(169, 154)
        Me.lblDept.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDept.Name = "lblDept"
        Me.lblDept.Size = New System.Drawing.Size(311, 25)
        Me.lblDept.TabIndex = 19
        Me.lblDept.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmployeeNum
        '
        Me.lblEmployeeNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEmployeeNum.Location = New System.Drawing.Point(169, 122)
        Me.lblEmployeeNum.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEmployeeNum.Name = "lblEmployeeNum"
        Me.lblEmployeeNum.Size = New System.Drawing.Size(311, 25)
        Me.lblEmployeeNum.TabIndex = 18
        Me.lblEmployeeNum.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLastName
        '
        Me.lblLastName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLastName.Location = New System.Drawing.Point(169, 90)
        Me.lblLastName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(311, 25)
        Me.lblLastName.TabIndex = 17
        Me.lblLastName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblMiddleName
        '
        Me.lblMiddleName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiddleName.Location = New System.Drawing.Point(169, 58)
        Me.lblMiddleName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(311, 25)
        Me.lblMiddleName.TabIndex = 16
        Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFirstName
        '
        Me.lblFirstName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFirstName.Location = New System.Drawing.Point(169, 26)
        Me.lblFirstName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(311, 25)
        Me.lblFirstName.TabIndex = 15
        Me.lblFirstName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(56, 255)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(107, 17)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "E-mail Address:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(87, 223)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(73, 17)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Extension:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(80, 191)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 17)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Telephone:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(75, 159)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 17)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Department:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(33, 127)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 17)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Employee Number:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(80, 95)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Last Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(65, 63)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Middle Name:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(81, 31)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(80, 17)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "First Name:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblEmail)
        Me.GroupBox1.Controls.Add(Me.lblExtension)
        Me.GroupBox1.Controls.Add(Me.lblTelephone)
        Me.GroupBox1.Controls.Add(Me.lblDept)
        Me.GroupBox1.Controls.Add(Me.lblEmployeeNum)
        Me.GroupBox1.Controls.Add(Me.lblLastName)
        Me.GroupBox1.Controls.Add(Me.lblMiddleName)
        Me.GroupBox1.Controls.Add(Me.lblFirstName)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Location = New System.Drawing.Point(44, 64)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(506, 295)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter Employee Data"
        '
        'lblRecordNum
        '
        Me.lblRecordNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRecordNum.Location = New System.Drawing.Point(325, 20)
        Me.lblRecordNum.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblRecordNum.Name = "lblRecordNum"
        Me.lblRecordNum.Size = New System.Drawing.Size(59, 25)
        Me.lblRecordNum.TabIndex = 9
        Me.lblRecordNum.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(204, 25)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Record Number:"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ReadDataForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(582, 491)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblRecordNum)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ReadDataForm"
        Me.Text = "Employee Data"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnClear As Button
    Friend WithEvents btnNext As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents lblEmail As Label
    Friend WithEvents lblExtension As Label
    Friend WithEvents lblTelephone As Label
    Friend WithEvents lblDept As Label
    Friend WithEvents lblEmployeeNum As Label
    Friend WithEvents lblLastName As Label
    Friend WithEvents lblMiddleName As Label
    Friend WithEvents lblFirstName As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblRecordNum As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
End Class
