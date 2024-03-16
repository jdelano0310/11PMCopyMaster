Imports Microsoft.Office.Interop
Imports System.IO
Public Class Form1
    Dim _xlApp As Excel.Application
    Dim _xlMasterFile As Excel.Workbook
    Dim _functionSuccess As Boolean
    Dim _showLogInListBox As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' un comment this to have it run when the app is launched
        ' call each of the functions that are needed to copy the data

        Dim args As String() = Environment.GetCommandLineArgs() ' pull in any passed arguments

        ' if there is at least 2 argument then automatically run the code
        ' the exe is passed a DLL name while running in VS
        If args.Length > 1 Then

            If Not OpenMasterFile() Then End

            If Not CopyMasterData() Then End

            CloseMasterFile()

            End
        End If

    End Sub

    Private Function OpenMasterFile() As Boolean

        ' open the excel workbook deemed the master file
        Dim masterFile As String = "F:\Temp\region1.xlsx"

        _functionSuccess = True

        Try
            WriteToLog($"Opening Master File {masterFile}")
            _xlApp = New Excel.Application
            _xlMasterFile = _xlApp.Workbooks.Open(masterFile)
            WriteToLog("Opened")

        Catch ex As Exception
            WriteToLog($"Error Opening Master File {masterFile} {ex.Message}")
            _functionSuccess = False

            If Not _xlApp Is Nothing Then _xlApp.Quit() ' close Excel if it is open
        End Try

        Return _functionSuccess

    End Function

    Private Sub WriteToLog(writeText As String)
        ' this sub simply opens a text file for appending
        ' adds a time stamps and whatever the write text is to a log file

        Dim logFile As New IO.StreamWriter(Application.StartupPath & "11pmCopyLog.txt", True)

        logFile.WriteLine($"{Now.ToShortDateString} {Now.ToShortTimeString}: {writeText}")

        If _showLogInListBox Then lbTestResults.Items.Add($"{Now.ToShortDateString}: {writeText}")
        logFile.Close()

    End Sub

    Private Function CopyMasterData() As Boolean

        ' copy the values of the formulas to a new sheet in the master file

        ' you could move these values to a text file so if
        ' the master file format changes you don't have to
        ' modify the code directly
        Dim masterSheetName As String = "Master"
        Dim masterCellForNewName As String = "D1"
        Dim masterCopyFromRange As String = "D3:D5"
        Dim newSheetCopyToRange As String = "A2:A4"

        Dim xlMasterSheet As Excel.Worksheet = _xlMasterFile.Worksheets(masterSheetName)
        Dim xlNewSheet As Excel.Worksheet = Nothing

        Dim currentCodeAttempt As String = ""

        _functionSuccess = True

        Try
            WriteToLog($"Copying Master Sheet Data")

            ' create a new sheet for the data - after the master sheet
            currentCodeAttempt = "Adding new sheet"
            _xlMasterFile.Worksheets.Add(, xlMasterSheet)
            xlNewSheet = _xlMasterFile.Worksheets(_xlMasterFile.Worksheets.Count)
            xlNewSheet.Name = xlMasterSheet.Range(masterCellForNewName).Value

            ' add a title row to the new sheet
            currentCodeAttempt = "Adding title to new sheet"
            xlNewSheet.Range("A1").Value = "Totals"   ' ********* need to modify where the row heading(s) go

            ' copy the values to the new sheet
            currentCodeAttempt = "Copying data to the new sheet"
            xlMasterSheet.Range(masterCopyFromRange).Copy()
            xlNewSheet.Range(newSheetCopyToRange).PasteSpecial(Excel.XlPasteType.xlPasteValues)

            ' save the master file with the new sheet
            currentCodeAttempt = "Saving the workbook"
            _xlMasterFile.Save()

            WriteToLog("All Data Copied")

        Catch ex As Exception
            WriteToLog($"Error Copying Master Sheet Data during {currentCodeAttempt} {ex.Message}")
            _functionSuccess = False

        End Try

        Return _functionSuccess

    End Function

    Private Function CloseMasterFile() As Boolean

        ' close the workbook
        _functionSuccess = True

        Try
            WriteToLog("Closing Master File")
            _xlMasterFile.Close()
            _xlApp.Quit()
            WriteToLog("Closed")
        Catch ex As Exception
            WriteToLog($"Error Closing Master File {ex.Message}")
            _functionSuccess = False
        End Try

        Return _functionSuccess
    End Function

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click

        _showLogInListBox = True
        lbTestResults.Items.Clear()

        ' call each of the functions that are needed to copy the data
        If Not OpenMasterFile() Then End

        If Not CopyMasterData() Then End

        CloseMasterFile()

    End Sub
End Class
