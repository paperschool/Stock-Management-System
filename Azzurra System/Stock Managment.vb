'This imports the system database libaries
Imports System.Data.OleDb

'defining the stock management module
Module Stock_Managment

    'these variable are for pulling manipulating and storing data from the database
    Dim stockmanagementda As OleDbDataAdapter
    Dim stockmanagementdt As DataTable
    Dim stockmanagementds As DataSet = New DataSet

    'this subroutine is used to populate the style navigation window
    Sub populate_navigation_management()

        'this clears the dataset from previous selections
        stockmanagementds.Clear()

        'this sql select statement is used to select every record from the presets table that is a style 
        Dim fill_navigation_styles As String = "SELECT * FROM tblPresets WHERE PresetCategory = '" & "Style" & "' "
        stockmanagementda = New OleDbDataAdapter(fill_navigation_styles, conn)
        stockmanagementda.Fill(stockmanagementds, "styles")
        stockmanagementdt = stockmanagementds.Tables("styles")

        'using data in the datatable, the stock style navigation datagridview is populated
        With Main.dgvStock_Man_Style
            .AutoGenerateColumns = True
            'this identifies the dataset as the datasoure
            .DataSource = stockmanagementds
            .DataMember = "styles"
        End With

        'this formats all rows and columns of the datagridview and also hides unused columms
        Main.dgvStock_Man_Style.AutoResizeRows()
        Main.dgvStock_Man_Style.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Main.dgvStock_Man_Style.Columns(0).Visible = False
        Main.dgvStock_Man_Style.Columns(2).Visible = False

    End Sub


    'this sub is used to select stock IDs for editing
    Sub populate_navigation_names(i)

        'this variable is used to reference the style selected in the other navigation window
        Dim style As String = ""
        'this defines a new dataset
        Dim stockmanagementds2 As DataSet = New DataSet

        'this sets the style as the selected value in the other naviagtion window
        style = Main.dgvStock_Man_Style.Item(1, i).Value()

        'this catches a null selection of the navigation window
        If style = "" Then
            'this exits the sub
            Exit Sub
        End If

        'this clears the dataset from previous selections
        stockmanagementds2.Clear()

        'this sql statement selects all records from the stock table that share the stock style selected earlier
        Dim fill_navigation_names As String = "SELECT * FROM tblStock WHERE StockStyle = '" & style & "' "
        stockmanagementda = New OleDbDataAdapter(fill_navigation_names, conn)
        stockmanagementda.Fill(stockmanagementds2, "names")
        stockmanagementdt = stockmanagementds2.Tables("names")

        'using data in the datatable, the stock id navigation datagridview is populated
        With Main.dgvStock_Man_Name
            .AutoGenerateColumns = True
            .DataSource = stockmanagementds2
            .DataMember = "names"
        End With

        'this loop is used to hide unneeded columsn in the datagrid view
        For x As Integer = 0 To stockmanagementds2.Tables("names").Columns.Count - 1
            'a condition checking if x equals the columns 0, 1, or 5
            If x = 1 Or x = 0 Or x = 5 Then
            Else
                'otherwise all columns are hidden
                Main.dgvStock_Man_Name.Columns(x).Visible = False
            End If
        Next

        'this formats the rows and columsn in the datagridview
        Main.dgvStock_Man_Name.AutoResizeRows()
        Main.dgvStock_Man_Name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'this for loop, loops through each row in the datatable
        For x As Integer = 0 To stockmanagementds2.Tables("names").Rows.Count - 1
            'this variable is used to store the current stock level
            Dim Stocklevel As String
            'this identifies the stock level as the stock id of the current row in the datatable being observed
            Stocklevel = Main.dgvStock_Man_Name.Item(0, x).Value()
            'this runs the check sizes subroutine passing the stock level 
            check_sizes(Stocklevel)
        Next

    End Sub

    'this subroutine is used to check all the stock totals for all the stock items to check whats low and what isnt
    Sub check_sizes(stocklevel)

        'this sets the stock count to 0
        Dim stockcount As Integer = 0

        'this redefines the dataset
        Dim stockmanagementds As DataSet = New DataSet

        'this defines an sql statement that seletcs all records that match that stock id (should be only 1)
        Dim checkstocklevel As String = "SELECT * FROM tblStock WHERE StockID = '" & stocklevel & "' "
        stockmanagementda = New OleDbDataAdapter(checkstocklevel, conn)
        stockmanagementda.Fill(stockmanagementds, "stocklevel")
        stockmanagementdt = stockmanagementds.Tables("stocklevel")

        'this loops through all the size columns in the stock record 
        For i As Integer = 6 To stockmanagementds.Tables("stocklevel").Columns.Count - 3
            'each time totalling the current total with the content of the size count producing a total count
            stockcount = stockcount + CInt(stockmanagementds.Tables("stocklevel").Rows(0).Item(i).ToString)
        Next

        'this condition checks if the value is lower then or equal to 20% of the original total
        If stockcount <= (stockmanagementds.Tables("stocklevel").Rows(0).Item(27) * 0.2) Then
            'this loops through each row in the datagridview
            For Each row As DataGridViewRow In Main.dgvStock_Man_Name.Rows
                'this checks if the current rows stock id matches the one being checked
                If row.Cells("stockid").Value = stocklevel Then
                    'this changes the row colour to red highlighting the stock level as low
                    row.DefaultCellStyle.BackColor = Color.Salmon
                End If
            Next
        End If

        'the stock count is then reset to 0
        stockcount = 0

    End Sub

    'this subroutine is used to take a stock item selected and to populate a stock profile to be edited
    Sub populate_Stock_information(i, a)

        'this defines stock id
        Dim stockid As String = ""
        'this defines a new dataset
        Dim stockmanagementds3 As DataSet = New DataSet

        'this prevents the subroutine from referencing null values when populating searched stock data
        If Not IsDBNull(Main.dgvStock_Man_Name.Item(0, i).Value()) Then
        ElseIf Not IsDBNull(Main.dgvStock_Man_Name.Item(0, i).Value()) Then
        Else
            Exit Sub
        End If

        'this sets the value of the stock id to the stock id column regardless of which cell is selected
        If a = 0 Or a = 1 Or a = 6 Then
            stockid = Main.dgvStock_Man_Name.Item(0, i).Value()
        End If

        'this defines an sql statement that will select that stock record containing the stock id
        Dim fill_stock_information As String = "SELECT * FROM tblStock WHERE StockID = '" & stockid & "' "
        stockmanagementda = New OleDbDataAdapter(fill_stock_information, conn)
        stockmanagementda.Fill(stockmanagementds3, "stockdata")
        stockmanagementdt = stockmanagementds3.Tables("stockdata")

        'this large block of text inserts data from each column in the stock record into the appropriate controls on the edit stock page
        'this is the main information block
        Main.txtStockID.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(0).ToString
        Main.txtStocknameedit.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(1).ToString
        Main.cbStockstyleedit.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(2).ToString
        Main.cbstocksupplieredit.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(3).ToString
        Main.txtstockpriceedit.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(4).ToString
        Main.cbstockcolouredit.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(5).ToString

        'this is the size information block
        Main.txtedit34.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(6).ToString
        Main.txtedit345.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(7).ToString
        Main.txtedit35.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(8).ToString
        Main.txtedit355.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(9).ToString
        Main.txtedit36.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(10).ToString
        Main.txtedit365.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(11).ToString
        Main.txtedit37.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(12).ToString
        Main.txtedit375.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(13).ToString
        Main.txtedit38.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(14).ToString
        Main.txtedit385.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(15).ToString
        Main.txtedit39.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(16).ToString
        Main.txtedit395.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(17).ToString
        Main.txtedit40.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(18).ToString
        Main.txtedit405.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(19).ToString
        Main.txtedit41.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(20).ToString
        Main.txtedit415.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(21).ToString
        Main.txtedit42.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(22).ToString
        Main.txtedit425.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(23).ToString
        Main.txtedit43.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(24).ToString
        Main.txtedit435.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(25).ToString
        Main.txtedit44.Text = stockmanagementds3.Tables("stockdata").Rows(0).Item(26).ToString

        'this will contain the control ids
        Dim l As Control

        'this loops through each of the size controls in the stock edit page
        For Each l In Main.gbStock_man_sizes.Controls
            'this checks if the control is a textbox
            If (TypeOf l Is TextBox) Then
                'if so, it changes the background of the control to green
                l.BackColor = Color.YellowGreen
            End If
        Next

        'this loops through each of the main controls in the stock edit page
        For Each l In Main.gbStock_man_info.Controls
            'this condition checks if the control is a textbox
            If (TypeOf l Is TextBox) Then
                'if so, it changes the background of the control to green
                l.BackColor = Color.YellowGreen
                'this condition checks if the control is a combobox
            ElseIf (TypeOf l Is ComboBox) Then
                'if so, it changes the background of the control to green
                l.BackColor = Color.YellowGreen
            End If
        Next

    End Sub

    'this subroutine is used to save the stock once enterted
    Sub save_edited_stock()

        'this sets I as a control id 
        Dim i As Control

        'this loops through each of the size controls
        For Each i In Main.gbStock_man_sizes.Controls
            'this condition checks if the control is a textbox
            If (TypeOf i Is TextBox) Then
                'this condition checks if the back colour doesnt equal green
                If i.BackColor <> Color.YellowGreen Then
                    'this message box points out the erroneous entry and prompts the user to fix it whilst exiting the subroutine
                    MsgBox("Value: " & i.Text & " is not valid")
                    Exit Sub
                End If
            End If
        Next

        'this loops through each of the main controls
        For Each i In Main.gbStock_man_info.Controls
            'this condition checks if the control is a textbox
            If (TypeOf i Is TextBox) Then
                'this condition checks if the back colour doesnt equal green
                If i.BackColor <> Color.YellowGreen Then
                    'this message box points out the erroneous entry and prompts the user to fix it whilst exiting the subroutine
                    MsgBox("Value: " & i.Text & " is not correct!")
                    Exit Sub
                End If
                'this condition checks if the control is a combobox
            ElseIf (TypeOf i Is ComboBox) Then
                'this condition checks if the back colour doesnt equal green
                If i.BackColor <> Color.YellowGreen Then
                    'this message box points out the erroneous entry and prompts the user to fix it whilst exiting the subroutine
                    MsgBox("Value: " & i.Text & " is not correct!")
                    Exit Sub
                End If
            End If
        Next

        'this message box displays assuming all validation passed successfully, asking if the user wants to overwrite current data
        Dim result As Integer = MessageBox.Show("Are you sure you want to overwrite current data?", "Warning", MessageBoxButtons.YesNo)
        'if yes is selected, the operation is ended exiting the subroutine
        If result = DialogResult.No Then
            Exit Sub
            'if yes is selected, the subroutine is continues
        ElseIf result = DialogResult.Yes Then
        End If

        'this stock insert statement is used to reinsert all the data new or not into the old record updating it
        Dim Stockinsert As String = "UPDATE tblStock SET StockName =?, StockStyle=?, StockSupplier=?, StockPrice=?, StockColour=?, 34=?, 345=?, 35=?, 355=?, 36=?, 365=?, 37=?, 375=?,  " & _
            "38=?, 385=?, 39=?, 395=?, 40=?, 405=?, 41=?, 415=?, 42=?, 425=?, 43=?, 435=?, 44=? WHERE StockID = @StockID"
        Dim StockinsertCommand As New OleDbCommand
        With StockinsertCommand
            .CommandText = Stockinsert
            .Parameters.AddWithValue("@p1", Main.txtStocknameedit.Text)
            .Parameters.AddWithValue("@p2", Main.cbStockstyleedit.Text)
            .Parameters.AddWithValue("@p3", Main.cbstocksupplieredit.Text)
            .Parameters.AddWithValue("@p4", Main.txtstockpriceedit.Text)
            .Parameters.AddWithValue("@p5", Main.cbstockcolouredit.Text)
            .Parameters.AddWithValue("@p6", Main.txtedit34.Text)
            .Parameters.AddWithValue("@p7", Main.txtedit345.Text)
            .Parameters.AddWithValue("@p8", Main.txtedit35.Text)
            .Parameters.AddWithValue("@p9", Main.txtedit355.Text)
            .Parameters.AddWithValue("@p10", Main.txtedit36.Text)
            .Parameters.AddWithValue("@p11", Main.txtedit365.Text)
            .Parameters.AddWithValue("@p12", Main.txtedit37.Text)
            .Parameters.AddWithValue("@p13", Main.txtedit375.Text)
            .Parameters.AddWithValue("@p14", Main.txtedit38.Text)
            .Parameters.AddWithValue("@p15", Main.txtedit385.Text)
            .Parameters.AddWithValue("@p16", Main.txtedit39.Text)
            .Parameters.AddWithValue("@p17", Main.txtedit395.Text)
            .Parameters.AddWithValue("@p18", Main.txtedit40.Text)
            .Parameters.AddWithValue("@p19", Main.txtedit405.Text)
            .Parameters.AddWithValue("@p20", Main.txtedit41.Text)
            .Parameters.AddWithValue("@p21", Main.txtedit415.Text)
            .Parameters.AddWithValue("@p22", Main.txtedit42.Text)
            .Parameters.AddWithValue("@p23", Main.txtedit425.Text)
            .Parameters.AddWithValue("@p24", Main.txtedit43.Text)
            .Parameters.AddWithValue("@p25", Main.txtedit435.Text)
            .Parameters.AddWithValue("@p26", Main.txtedit44.Text)
            .Parameters.AddWithValue("@StockID", Main.txtStockID.Text)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        'successful stock edit message
        MsgBox("Stock Edited Successfuly")

        'repopulates stock navigation menu
        populate_navigation_management()

    End Sub

    'this subroutine runs when the cancel button is pressed
    Sub Cancel_edited_stock()

        'this defines i as a control id
        Dim i As Control

        'this loops through each size control
        For Each i In Main.gbStock_man_sizes.Controls
            'this condition checks if the control is a textbox
            If (TypeOf i Is TextBox) Then
                'this resets the colour and content of the control
                i.BackColor = Color.White
                i.Text = ""
            End If
        Next

        'this loops through each main control
        For Each i In Main.gbStock_man_info.Controls
            'this condition checks if the control is a textbox
            If (TypeOf i Is TextBox) Then
                'this resets the colour and content of the control
                i.BackColor = Color.White
                i.Text = ""
                'this condition checks if the control is a combobox
            ElseIf (TypeOf i Is ComboBox) Then
                'this resets the colour and content of the control
                i.BackColor = Color.White
                i.Text = ""
            End If
        Next

    End Sub

    'this subroutine is used to validate newely input data in the stock profile
    Sub validate_edited_information()

        'this condition checks if the name of the active control = txtstocknameedit
        If Main.ActiveControl.Name = "txtStocknameedit" Then
            'this loops through each character in the current active controls content
            For Each i As Char In Main.ActiveControl.Text
                'this conidtion checks if the current character is a letter
                If Char.IsLetter(i) Then
                    'this changes the back colour of the textbox to red
                    Main.ActiveControl.BackColor = Color.YellowGreen
                    'this condition checks if the current character being checked is a null value
                ElseIf String.IsNullOrEmpty(i) Then
                    'this changes the back colour to white
                    Main.ActiveControl.BackColor = Color.White
                    Exit Sub
                Else
                    'otherwise the back colour is changed to red
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                End If
            Next
        End If

        'validation for price input string
        If Main.ActiveControl.Name = "txtstockpriceedit" Then
            'this loop will iterate through each character in the content of the currently active control
            For Each i As Char In Main.ActiveControl.Text
                'this condition checks if the current character is a number
                If Char.IsNumber(i) Then
                    'this will change the back colour to green
                    Main.ActiveControl.BackColor = Color.YellowGreen
                    'this condition checks if the current character is a letter
                ElseIf Char.IsLetter(i) Then
                    'this will change the back colour to red and end the subroutine
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                    'this condition checks if the character is a symbol
                ElseIf Char.IsSymbol(i) = True Then
                    'this changes the back colour to red and exits the subroutine
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                    'this condition checks if the character is a punctuation mark
                ElseIf Char.IsPunctuation(i) = True Then
                    'user may use a decimal in the price so this punctuation must be removed from search'
                    'this condition will catch the use of a full stop
                    If i = "." Then
                        'this condition checks if the 3 character from the end is a full stop, if not 
                        If Main.ActiveControl.Text.Substring(Main.ActiveControl.Text.Length - 3, 1) <> "." Then
                            'the back colour is changed to blue signifying that the input is not wrong but conventionally wrong, and will exit the subroutine
                            Main.ActiveControl.BackColor = Color.Aquamarine
                            Exit Sub
                        End If
                        'this condition checks if the last character is not numeric or if the second to last character is not numeric if so
                        If IsNumeric(Main.ActiveControl.Text.Substring(Main.ActiveControl.Text.Length - 1, 1)) = False Or _
                            IsNumeric(Main.ActiveControl.Text.Substring(Main.ActiveControl.Text.Length - 2, 1)) = False Then
                            'the back colour is changed to red and the subroutine is exited
                            Main.ActiveControl.BackColor = Color.Salmon
                            Exit Sub
                        End If
                        Exit Sub
                    End If
                    'this will automatically change the back colour to red if the above condition is not true as any other punctuation is incorrect
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                    'this condition checks if the entry is null 
                ElseIf Main.ActiveControl.Text = "" Then
                    'if so it changes the back colour to white
                    Main.ActiveControl.BackColor = Color.White
                    Exit Sub
                    'this condition checks if the current character is a space, if so
                ElseIf String.IsNullOrWhiteSpace(i) Then
                    'the back colour is changed to red
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                End If
            Next
        End If

    End Sub

    'this subroutine is used to validate each size 
    Sub validate_edited_sizes()

        'this defines entry as a string
        Dim Entry As String = ""

        'this sets the backcolour automatically to red
        Main.ActiveControl.BackColor = Color.Red

        'this condition checks if the content of the textbox exceed 2 characters long
        If Main.ActiveControl.Text.Length < 2 Then
            'this sets entry to equal the content of the active control text
            Entry = Main.ActiveControl.Text
            'this checks if the character in the textbox ix a number
            If Char.IsNumber(Entry) Then
                'if so it changes the back colour to green
                Main.ActiveControl.BackColor = Color.YellowGreen
                Exit Sub
            Else
                'if its not a number it changes the backcolour to red
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            End If
            'if the content is nothing
        ElseIf Main.ActiveControl.Text = "" Then
            'then the back colour is changed to white
            Main.ActiveControl.BackColor = Color.White
        End If

    End Sub

End Module
