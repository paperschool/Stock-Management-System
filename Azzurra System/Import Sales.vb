'this imports the database connection module
Imports System.Data.OleDb

Module Import_Sales

    'this defines the connection variable
    Dim cn As OleDbConnection

    'this defines the data adapter variable
    Dim da As OleDbDataAdapter

    'this defines the datatable variable
    Dim dt As DataTable

    'this defines the dataset variable
    Dim ds As DataSet = New DataSet

    'this defines a sales dataset
    Dim dssales As DataSet = New DataSet

    'this declares a filepath string
    Dim file_path As String

    'this defines a file dialog 
    Dim importfiledialog As New OpenFileDialog

    'this defines a sales id string variable
    Dim SalesID As String

    'this defines a database connection
    Dim tbl As System.Data.OleDb.OleDbConnection

    'this defines an excel worksheet datatable
    Dim ExcelTables As DataTable = Nothing

    'this import file subroutine manages the imported file data
    Sub Import_file()

        'this resets the dataset for the next file import
        dssales.Clear()

        'this subroutine will manage the file to be imported and its location
        open_File()

        'this lays out the file location of the spreadsheet and provides some excel properties
        cn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & file_path & "; Extended properties=Excel 8.0;")

        'this sets the table name to null
        Dim tablename As String = ""

        'this try statement will attempt to located specific database scheme information
        Try
            'this declares a database connection string
            Dim sheetname As OleDbConnection
            sheetname = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & file_path & "; Extended properties=Excel 8.0;")
            sheetname.Open()
            ExcelTables = sheetname.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLES"})
            tablename = (sheetname.GetSchema("TABLES").Rows(0)("TABLE_NAME"))
        Catch ex As Exception
        Finally
        End Try

        'this attempts to reopen the connection and if not exits the subroutine
        Try
            cn.Open()
        Catch
            Exit Sub
        End Try

        'this runs the datagridview fill subrotuine, porting the table name of the worksheet
        FillSalesDataGrid("select * from [" & tablename & "]")

    End Sub

    'this subroutine manages the file to be imported
    Sub open_File()

        'this clears the dataset
        ds.Clear()

        'these filedialog properties will set the type of file the database opens, and will allow for some other settings
        importfiledialog.FileName = "" ' Default file name
        importfiledialog.DefaultExt = ".xls" ' Default file extension
        importfiledialog.Filter = "Excel Documents (*.XLS)|*.XLS"
        importfiledialog.Multiselect = True
        importfiledialog.RestoreDirectory = True

        'Show open file dialog box
        Dim result? As Boolean = importfiledialog.ShowDialog()

        'Process open file dialog box results
        For Each path In importfiledialog.FileNames
            file_path = path
        Next

    End Sub

    'this subroutine will use file path data from the previous 
    Private Sub FillSalesDataGrid(ByVal Query As String)

        'this redeclares the data adapter with data collected about the location of the 
        da = New OleDbDataAdapter(Query, cn)
        da.Fill(dssales, "ImportSales")
        dt = dssales.Tables("ImportSales")

        'this will loop through all the data in the dataset belonging to the worksheet
        For Each row As DataRow In dt.Rows
            'this condition checks if the current row has no data in it, signifying an empty row
            If IsDBNull(row.Item(0)) = True And IsDBNull(row.Item(1)) = True And IsDBNull(row.Item(2)) = True And IsDBNull(row.Item(3)) = True And IsDBNull(row.Item(4)) = True Then
                'if so, the row is deleted 
                row.Delete()
            End If
        Next

        'this line will save any changes made to the data set
        dssales.AcceptChanges()

        Try
            'this will fill a datagrid with all the data from the datasource 
            With Main.dgvExcel
                .DataSource = dssales
                .DataMember = "ImportSales"
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Main.dgvExcel.AutoResizeRows()
        Main.dgvExcel.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'the purpose of this try is to catch any mistakes that could occur in this very complicated 
        Try
            collectdata()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    'this subroutine will generate a new sale id
    Sub Generate_Sales_ID()

        'SQL Select Statement Searching for the highest UserID assciated with that user type'
        SalesID = "SELECT MAX(SaleID) FROM tblSales WHERE SaleID LIKE '" & "SA" & "%%%%%' "
        'Data adapter defined'
        da = New OleDbDataAdapter(SalesID, conn)
        'Data adapter told to fill the da   taset
        da.Fill(ds, "tblSales")
        'Defining the datatable
        dt = ds.Tables("tblSales")
        'Associating User ID to the first item on the first row of the datatabl whilst converting the data to a string'
        SalesID = dt.Rows(0).Item(0).ToString

        'For defining the first student in the class;
        If SalesID = "" Then
            'this is some "Fake Data" that is used to calculate the first User ID'
            SalesID = "SA00000"
            Exit Sub
        End If


    End Sub

    'this subroutine will collate all the data that has been imported and then fill various arrays that can be cycled through added to the database as records
    Sub collectdata()

        'this defines the length of each array as the number of rows in the data grid
        Dim salearraylength As Integer = Main.dgvExcel.RowCount

        'these arrays will contain each column of data from the spread sheet
        Dim SaleIDArray(salearraylength) As String
        Dim StockID(salearraylength) As String
        Dim Colour(salearraylength) As String
        Dim StockStyle(salearraylength) As String
        Dim Size(salearraylength) As String
        Dim Price(salearraylength) As String
        Dim SaleDate(salearraylength) As String

        'this subroutine is used to find the highest stock id in the database
        Generate_Sales_ID()

        'this loops through each of the rows in the datagridview and asigns a value to each array corresponding to the grid in the datagridview and the row count
        For i As Integer = 0 To (salearraylength - 2)
            'the same series of arrays declared earlier
            StockID(i) = dssales.Tables("ImportSales").Rows(i).Item(0).ToString
            Colour(i) = dssales.Tables("ImportSales").Rows(i).Item(1).ToString
            StockStyle(i) = dssales.Tables("ImportSales").Rows(i).Item(2).ToString
            Size(i) = dssales.Tables("ImportSales").Rows(i).Item(3).ToString
            Price(i) = dssales.Tables("ImportSales").Rows(i).Item(4).ToString
            SaleDate(i) = dssales.Tables("ImportSales").Rows(i).Item(5).ToString
            Console.WriteLine("count: " & i)
        Next

        For i As Integer = 0 To (Main.dgvExcel.RowCount - 1)
            Console.WriteLine("count: " & i)
            'This strips off the User Type Identified'
            SalesID = SalesID.Substring(2, 5)
            'This removes the leading zero's from the stripped string'
            SalesID = SalesID.TrimStart("0"c)

            'This catches the User ID string if it is something like "S0000", after all the stripping and formating, would leave nothing, this code repairs it'
            If SalesID = "" Then
                'Re-Defines the SalesID to 0'
                SalesID = 0
            End If

            'Converts the split string into an integer'
            SalesID = CType(SalesID, Integer) + 1

            'This block is responsible for adding the appropriate value onto the end of the user ID'
            'This converts the User ID to a string'
            SalesID = CType(SalesID, String)
            'Length check's used to identify the different tiers of the user id'
            If SalesID.Length = 1 Then
                'This combines the User ID's componants into a full id'
                SalesID = "SA" & "0000" & SalesID
                'Length check's used to identify the different tiers of the user id'
            ElseIf SalesID.Length = 2 Then
                'This combines the User ID's componants into a full id'
                SalesID = "SA" & "000" & SalesID
                'Length check's used to identify the different tiers of the user id'
            ElseIf SalesID.Length = 3 Then
                'This combines the User ID's componants into a full id'
                SalesID = "SA" & "00" & SalesID
            ElseIf SalesID.Length = 4 Then
                SalesID = "SA" & "0" & SalesID
            Else
                'This combines the User ID's componants into a full id'
                SalesID = "SA" & SalesID
            End If

            SaleIDArray(i) = SalesID
        Next


        'this loop will cycle through each row in the datagridview 
        For i As Integer = 0 To Main.dgvExcel.RowCount - 2
            Console.WriteLine("count: " & i)

            If i = 45 Then
                Console.WriteLine("")
            End If

            'this string variable contains the insert statement which will create a record for each sale record complete with all the parameters
            Dim Salesinsert As String = "INSERT INTO tblSales ([SaleID],[StockID],[Colour],[StockStyle],[StockSize],[Price],[SaleDate]) VALUES " & _
                                        "(@SaleIDArray,@StockID,@Colour,@StockStyle,@Size,@price,@Date)"
            Dim SalesinsertCommand As New OleDbCommand
            With SalesinsertCommand
                .CommandText = Salesinsert

                'this block contains all the referenced variables that are to be inserted into the database
                .Parameters.AddWithValue("@SaleID", SaleIDArray(i))
                .Parameters.AddWithValue("@StockID", StockID(i))
                .Parameters.AddWithValue("@Colour", Colour(i))
                .Parameters.AddWithValue("@StockStyle", StockStyle(i))
                .Parameters.AddWithValue("@Size", Size(i))
                .Parameters.AddWithValue("@Price", Price(i))
                .Parameters.AddWithValue("@Date", SaleDate(i))

                'this will redeclar the connection 
                .Connection = conn

                'this will execute the database expression
                .ExecuteNonQuery()
            End With


            'these two lines provide me with more specific data about the database, in this case it will find all the field headers in the stocktable in my database
            Dim filterValues = {Nothing, Nothing, "tblStock", Nothing}
            Dim columns = conn.GetSchema("Columns", filterValues)

            'this defines the stock number variable that is used to update the stock size level
            Dim newstocknumber As Integer

            'this counter is used to keep track of which row in the record is currently being viewed
            Dim counter As Integer = 0

            'this string variable that contains the select statement from the stock table
            Dim stockdeduct As String = "SELECT * FROM tblStock WHERE StockID = '" & StockID(i) & "' "
            'Data adapter defined'
            da = New OleDbDataAdapter(stockdeduct, conn)
            'Data adapter told to fill the dataset
            da.Fill(ds, "stockdeduct")
            'Defining the datatable
            dt = ds.Tables("stockdeduct")

            'this will end the subroutine if no data matches the current stock id
            If ds.Tables("stockdeduct").Rows.Count = 0 Then
                Continue For
            End If

            'this loop will cycle through each row in columns variable until it find the table header that matches the size record
            For Each row As DataRow In columns.Rows

                'this condition only runs when the table header of the stock table matches the current size
                If row("column_name").ToString = Size(i) Then

                    'this defines the stock number as the current value of the stock size
                    newstocknumber = CInt(ds.Tables("stockdeduct").Rows(0).Item(counter + 6).ToString)

                    'this condition checks if the current stock value is at 0
                    If newstocknumber = 0 Then
                    Else
                        'if the stock value is not at zero it deducts one from the current value deducting a stock item from the database
                        newstocknumber = newstocknumber - 1
                    End If

                    'this update string will insert data where the stock id matches the one currently being looked at and will replace the current
                    'stock value with the newly calculated one
                    Dim updatestocklevel As String = "UPDATE tblStock SET " & row("column_name") & "= @Stockval WHERE StockID = @StockID"

                    Dim updatestocklevelCommand As New OleDbCommand
                    With updatestocklevelCommand
                        .CommandText = updatestocklevel

                        'these parameters reference the variable value and where its being sent in the sql statement
                        .Parameters.AddWithValue("@Stockval", newstocknumber)
                        .Parameters.AddWithValue("@StockID", StockID(i))

                        'this redeclares the connection string and executes the sql statement
                        .Connection = conn
                        .ExecuteNonQuery()
                    End With

                    'this clears both the dataset and datatable to ready the program for a new input
                    ds.Clear()
                    dt.Clear()

                End If

                'this counter keeps track of the current column number
                counter = counter + 1
            Next

            'this resets both the counter and the new stock number
            counter = 0
            newstocknumber = 0
        Next

        'this is called to repopulate the chart data
        populate_chart()

        populate_sales_history_all()

    End Sub

End Module
