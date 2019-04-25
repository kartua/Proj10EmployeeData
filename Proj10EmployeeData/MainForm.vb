Imports System.IO

Public Class MainForm
    Public strFilePath As String = ""               ' To hold the filename
    Public inputFile As StreamReader


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        inputFile.Close()
        Me.Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clear the form.
        lblFirstName.Text = String.Empty
        lblMiddleName.Text = String.Empty
        lblLastName.Text = String.Empty
        lblEmployeeNum.Text = String.Empty
        lblDept.Text = String.Empty
        lblTelephone.Text = String.Empty
        lblExtension.Text = String.Empty
        lblEmail.Text = String.Empty

        ' Reset the focus
        lblFirstName.Focus()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Read the next record.
        If inputFile.Peek <> -1 Then

            lblFirstName.Text = inputFile.ReadLine
            lblMiddleName.Text = inputFile.ReadLine
            lblLastName.Text = inputFile.ReadLine
            lblEmployeeNum.Text = inputFile.ReadLine
            lblDept.Text = inputFile.ReadLine
            lblTelephone.Text = inputFile.ReadLine
            lblExtension.Text = inputFile.ReadLine
            lblEmail.Text = inputFile.ReadLine

        Else
            MessageBox.Show("End of file.")
        End If

    End Sub

    Private Sub mnuFileOpen_Click(sender As Object, e As EventArgs) Handles mnuFileOpen.Click
        Try
            Dim respondOpenDialog As DialogResult
            With OpenFileDialog1
                ' Display the current directory in the window
                .InitialDirectory = Directory.GetCurrentDirectory
                ' Display a file name in the filename box
                .FileName = "employee.txt"
                ' Display a title for the OPEN dialog window
                .Title = "Select File"
                ' Save the user response - OPEN or CANCEL
                respondOpenDialog = .ShowDialog()
            End With
            If respondOpenDialog <> DialogResult.Cancel Then
                strFilePath = OpenFileDialog1.FileName
                inputFile = File.OpenText(strFilePath)
            Else
                Me.Close()
            End If

        Catch ex As Exception
            MessageBox.Show("File cannot be opened")
            Me.Close()
        End Try
    End Sub

    Private Sub AddRecordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddRecordToolStripMenuItem.Click
        If strFilePath = "" Then
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                strFilePath = SaveFileDialog1.FileName
                inputFile = File.OpenText(strFilePath)
            End If
        End If
        inputFile.Close()
        Dim writeForm = New WriteDataForm
        writeForm.ShowDialog()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' FORM ENTRIES REPORT

        ' Declare the Variables for the Fonts
        Dim HeadingFont As New Font("Arial", 14, FontStyle.Bold)
        Dim PrintFont As New Font("Arial", 12)

        ' Declare the Variable for the Height of the Characters
        Dim LineHeightSingle As Single = PrintFont.GetHeight + 2

        'Declare the X and Y Coordinate Variables
        Dim X_CoordinateSingle As Single = e.MarginBounds.Left
        Dim Y_CoordinateSingle As Single = e.MarginBounds.Top

        ' Declare a Variable for storing a horizontal line of print on the paper.
        Dim PrintLineString As String

        PrintLineString = "Employee Data"
        e.Graphics.DrawString(PrintLineString, HeadingFont, Brushes.Black, X_CoordinateSingle, Y_CoordinateSingle)

        ' Move the Cursor down the page vertically - increase the y-value
        Y_CoordinateSingle += LineHeightSingle
        Y_CoordinateSingle += LineHeightSingle

        PrintLineString = "First Name: " & lblFirstName.Text & vbCrLf
        PrintLineString += "Middle Name: " & lblMiddleName.Text & vbCrLf
        PrintLineString += "Last Name: " & lblLastName.Text & vbCrLf
        PrintLineString += "Employee Number: " & lblEmployeeNum.Text & vbCrLf
        PrintLineString += "Department: " & lblDept.Text & vbCrLf
        PrintLineString += "Telephone: " & lblTelephone.Text & vbCrLf
        PrintLineString += "Estension: " & lblExtension.Text & vbCrLf
        PrintLineString += "E-mail Address: " & lblEmail.Text & vbCrLf

        e.Graphics.DrawString(PrintLineString, PrintFont, Brushes.Black, X_CoordinateSingle, Y_CoordinateSingle)
    End Sub
End Class
