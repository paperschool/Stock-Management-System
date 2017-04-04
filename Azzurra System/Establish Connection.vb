'importing the system database control module
Imports System.Data.OleDb

Module Establish_Connection

    'this creates a file dialog for locating a missing database
    Dim OpenFileDlg As New OpenFileDialog


    Dim filepath As String

    'this defines conna as a public database connection variable
    Public conn As New OleDbConnection

    'this subroutine will define the connection string to the database
    Public Sub connect()

        'this try will run the connection code with the potential to fail if the database file is missing or corrupt
        Try

            'this defines the relative location of the database object
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|datadirectory|\AzzuraDatabase.accdb;"

            'this checks if the connection state is open or closed
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            'this will catch any errors found in the above code and proceed with the alternative code
        Catch ex As Exception

            'this message box will alert the user that the error has occured
            MsgBox("Database Not Located, User Location required.")

            'this will run the database search subroutine
            finddatabase()

        End Try
    End Sub

    'this subroutine will open a dialog box and will further allow the user to select a database
    Sub finddatabase()

        'these filedialog properties will set the type of file the database opens, and will allow for some other settings
        OpenFileDlg.FileName = "" ' Default file name
        OpenFileDlg.DefaultExt = ".accdb" ' Default file extension
        OpenFileDlg.Filter = "Access Documents (*.ACCDB)|*.ACCDB"
        OpenFileDlg.Multiselect = True
        OpenFileDlg.RestoreDirectory = True

        'Show open file dialog box
        Dim result? As Boolean = OpenFileDlg.ShowDialog()

        'Process open file dialog box results
        For Each path In OpenFileDlg.FileNames
            filepath = path
        Next

        'this will check to ensure a file was chosen if not the program is closed
        If filepath = "" Then
            MsgBox("No Database selected")
            Login.Close()
        End If

        'experimental code to move database from located directory to main directory 
        Dim FileToMove As String = filepath
        Dim MoveLocation As String = My.Application.Info.DirectoryPath

        MsgBox(My.Application.Info.DirectoryPath & " | " & FileToMove)

        If System.IO.File.Exists(FileToMove) = True Then

            System.IO.File.Move(FileToMove, MoveLocation)

        End If

        'this formats a new connection string for the database location with the selected database
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|datadirectory|\AzzuraDatabase.accdb;"

        'this checks if the connection state is open or closed
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

    End Sub

End Module
