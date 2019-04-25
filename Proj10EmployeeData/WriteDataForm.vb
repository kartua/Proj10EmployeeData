Imports System.IO

Public Class WriteDataForm
    Dim strFilePath As String = MainForm.strFilePath
    Dim ResponseSaveDialogResult As DialogResult
    Dim employeeFile As StreamWriter



    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        MainForm.inputFile = File.OpenText(strFilePath)
        Me.Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clear the foem
        txtFirstName.Text = String.Empty
        txtMiddleName.Text = String.Empty
        txtLastName.Text = String.Empty
        txtEmployeeNumber.Text = String.Empty
        cboDepartment.Text = String.Empty
        txtTelephone.Text = String.Empty
        txtExtension.Text = String.Empty
        txtEmail.Text = String.Empty

        ' Reset the focus
        txtFirstName.Focus()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try


            employeeFile = File.AppendText(strFilePath)
            employeeFile.WriteLine(txtFirstName.Text)
            employeeFile.WriteLine(txtMiddleName.Text)
            employeeFile.WriteLine(txtLastName.Text)
            employeeFile.WriteLine(txtEmployeeNumber.Text)
            employeeFile.WriteLine(cboDepartment.Text)
            employeeFile.WriteLine(txtTelephone.Text)
            employeeFile.WriteLine(txtExtension.Text)
            employeeFile.WriteLine(txtEmail.Text)
            employeeFile.Close()
            MessageBox.Show("The record has been saved.")
        Catch ex As Exception

        End Try

    End Sub


End Class
