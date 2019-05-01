Imports System.IO

Public Class MainForm
    Public strFilePath As String = ""               ' To hold the filename
    Public inputFile As StreamReader
    Dim maxRecord As Int16 = 9
    Dim numberRecord As Int16 = 0
    Public emplyoeeData(maxRecord) As EmployeeData

    Structure EmployeeData
        Public firstName As String
        Public midName As String
        Public lastName As String
        Public employeeNumber As Long
        Public department As String
        Public telephone As String
        Public ext As String
        Public email As String
    End Structure


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Try
            'close the Stream Reader
            inputFile.Close()
            Me.Close()
        Catch ex As Exception
            Me.Close()
        End Try

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
        Try

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
        Catch ex As Exception
            MessageBox.Show("Cannot read a file")
        End Try


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
                'call populateDataToArray precedure
                populateDataToArray()
            Else
                Me.Close()
            End If

        Catch ex As Exception
            MessageBox.Show("File cannot be opened")
        Me.Close()
        End Try
    End Sub

    Private Sub AddRecordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddRecordToolStripMenuItem.Click
        'Check if the file has not been opened yet
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

        'add data to PrintLineString
        PrintLineString = "First Name: " & lblFirstName.Text & vbCrLf
        PrintLineString += "Middle Name: " & lblMiddleName.Text & vbCrLf
        PrintLineString += "Last Name: " & lblLastName.Text & vbCrLf
        PrintLineString += "Employee Number: " & lblEmployeeNum.Text & vbCrLf
        PrintLineString += "Department: " & lblDept.Text & vbCrLf
        PrintLineString += "Telephone: " & lblTelephone.Text & vbCrLf
        PrintLineString += "Estension: " & lblExtension.Text & vbCrLf
        PrintLineString += "E-mail Address: " & lblEmail.Text & vbCrLf

        'put PrintLineString to the file
        e.Graphics.DrawString(PrintLineString, PrintFont, Brushes.Black, X_CoordinateSingle, Y_CoordinateSingle)
    End Sub

    Public Sub populateDataToArray()
        Try
            inputFile = File.OpenText(strFilePath)
            numberRecord = 0
            While Not (inputFile.EndOfStream)
                'Check if the number of record exceed the array capacity or not
                If numberRecord > maxRecord Then
                    increaseArraySize()
                End If
                emplyoeeData(numberRecord).firstName = inputFile.ReadLine
                emplyoeeData(numberRecord).midName = inputFile.ReadLine
                emplyoeeData(numberRecord).lastName = inputFile.ReadLine
                emplyoeeData(numberRecord).employeeNumber = inputFile.ReadLine
                emplyoeeData(numberRecord).department = inputFile.ReadLine
                emplyoeeData(numberRecord).telephone = inputFile.ReadLine
                emplyoeeData(numberRecord).ext = inputFile.ReadLine
                emplyoeeData(numberRecord).email = inputFile.ReadLine

                'Tracking number of record
                numberRecord += 1

            End While
            inputFile.Close()
            inputFile = File.OpenText(strFilePath)
        Catch ex As Exception
            MessageBox.Show("Cannot read a file")
        End Try

    End Sub

    Private Sub increaseArraySize()
        'Expand array size of employeeData
        maxRecord += 10
        ReDim Preserve emplyoeeData(maxRecord)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i As Int16 = 0 To numberRecord - 1
            Dim strOutput As String = ""
            strOutput += "record#" & i
            strOutput += emplyoeeData(i).firstName & " "
            strOutput += emplyoeeData(i).midName & " "
            strOutput += emplyoeeData(i).lastName & " "
            strOutput += emplyoeeData(i).employeeNumber & " "
            MessageBox.Show(strOutput)
        Next




    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        Dim searchID = InputBox("Input Employee ID")
        Dim isFound As Boolean = False 'Variable to indicate that the record is found or not

        'Iterate trhough all record in the array
        For i As Int16 = 0 To numberRecord - 1
            If emplyoeeData(i).employeeNumber = searchID Then
                lblFirstName.Text = emplyoeeData(i).firstName
                lblMiddleName.Text = emplyoeeData(i).midName
                lblLastName.Text = emplyoeeData(i).lastName
                lblEmployeeNum.Text = emplyoeeData(i).employeeNumber
                lblDept.Text = emplyoeeData(i).department
                lblTelephone.Text = emplyoeeData(i).telephone
                lblExtension.Text = emplyoeeData(i).ext
                lblEmail.Text = emplyoeeData(i).email
                isFound = True
            End If
        Next

        If Not isFound Then
            MessageBox.Show("No record is found")
        End If

    End Sub
End Class
