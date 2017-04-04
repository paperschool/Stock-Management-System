'imports data base connectivity module
Imports System.Data.OleDb

Class Stock_Input

    'this will bring all object referencing over from the main form
    Inherits Main

    'defines database connection and data collection variables
    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim ds As DataSet = New DataSet

    'a stock id variable used to store both sql statements and stock id information
    Public StockID As String

    'an array storing a generic run of stock values
    Dim linearrun As New List(Of Integer) From {1, 1, 1, 1, 2, 3, 3, 4, 4, 5, 5, 5, 4, 4, 3, 3, 2, 1, 1, 1, 1}

    'unused variables for a custom calculated linear generator
    Dim curve_value As Integer
    Dim ar_value As Integer = 0

    'a subroutine used to populate all stock ammounts in each variable
    Sub Linear_Run()

        Dim i As Control

        curve_value = 0
        ar_value = 0

        For Each i In Main.gbSize_stock_information.Controls
            If (TypeOf i Is TextBox) Then
                i.BackColor = Color.YellowGreen
                i.Text = linearrun(ar_value)
                ar_value = ar_value + 1
            End If
        Next

    End Sub

    Sub Fill_sizes()

        Dim Entry As String = ""

        Main.ActiveControl.BackColor = Color.Salmon

        If Main.ActiveControl.Text.Length < 2 Then
            Entry = Main.ActiveControl.Text
            If Char.IsLetter(Entry) Then
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            ElseIf Char.IsSymbol(Entry) = True Then
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            ElseIf Char.IsPunctuation(Entry) = True Then
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            ElseIf Entry = "" Then
                Main.ActiveControl.BackColor = Color.White
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(Entry) Then
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            ElseIf Main.ActiveControl.Text.Length > 2 Then
                Main.ActiveControl.BackColor = Color.Salmon
                Exit Sub
            Else
                Main.ActiveControl.BackColor = Color.YellowGreen
            End If
        ElseIf Main.ActiveControl.Text = "" Then
            Main.ActiveControl.BackColor = Color.DodgerBlue
        End If


    End Sub

    Sub Main_information_validation()

        Main.ActiveControl.BackColor = Color.Salmon

        'validation for stock name string, this will check each character of the entry ensuring it is of correct datatype
        If Main.ActiveControl.Name <> "txtStockPrice" Then
            If Main.ActiveControl.Text = "" Then
                Main.ActiveControl.BackColor = Color.White
                Exit Sub
            End If
            For Each i As Char In Main.ActiveControl.Text
                If Char.IsLetter(i) Then
                    'this is an example of a correct input, this changes the textbox background to green
                    Main.ActiveControl.BackColor = Color.YellowGreen
                Else
                    Main.ActiveControl.BackColor = Color.Salmon
                End If
                If Main.ActiveControl.Text = "" Then
                    'this is an example of an incorrect entry showing the result as red
                    Main.ActiveControl.BackColor = Color.White
                    Exit Sub
                End If
                If String.IsNullOrWhiteSpace(i) Then
                    'this is an example of an incorrect entry showing the result as red
                    Main.ActiveControl.BackColor = Color.YellowGreen
                    Exit Sub
                End If
            Next
        End If

        'validation for price input string, this will check each character of the entry ensuring it is of correct datatype
        If Main.ActiveControl.Name = "txtStockPrice" Then
            For Each i As Char In Main.ActiveControl.Text
                If Char.IsNumber(i) Then
                    Main.ActiveControl.BackColor = Color.YellowGreen

                End If
                If Char.IsLetter(i) Then
                    Main.ActiveControl.BackColor = Color.Salmon

                End If
                If Char.IsSymbol(i) = True Then
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub

                End If
                If Char.IsPunctuation(i) = True Then
                    'user may use a decimal in the price so this punctuation must be removed from search'
                    If i = "." Then
                        If Main.ActiveControl.Text.Substring(Main.ActiveControl.Text.Length - 3, 1) <> "." Then
                            Main.ActiveControl.BackColor = Color.Aquamarine
                            Exit Sub
                        End If
                    Else
                        Main.ActiveControl.BackColor = Color.Salmon
                        Exit Sub
                    End If

                End If
                If Main.ActiveControl.Text = "" Then
                    Main.ActiveControl.BackColor = Color.White
                    Exit Sub

                End If
                If String.IsNullOrWhiteSpace(i) Then
                    Main.ActiveControl.BackColor = Color.Salmon
                    Exit Sub
                End If
            Next
            
        End If

    End Sub

    Sub Clear_stock_inputs()

        Dim result As Integer = MessageBox.Show("Are you sure you want to empty the stock profile?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
        End If

        'this loop will cycle through each  control in the Size info groupbox that is either a text box or a combobox 
        'and will clear its content and reset is color formatting
        Dim i As Control
        For Each i In Main.gbSize_stock_information.Controls
            If (TypeOf i Is TextBox) Then
                i.BackColor = Color.White
                i.Text = ""
            ElseIf (TypeOf i Is ComboBox) Then
                i.BackColor = Color.White
                i.Text = ""
            End If
        Next

        'this loop will cycle through each  control in the main info groupbox that is either a text box or a combobox 
        'and will clear its content and reset is color formatting
        For Each i In Main.gbMain_stock_information.Controls
            If (TypeOf i Is TextBox) Then
                i.BackColor = Color.White
                i.Text = ""
            ElseIf (TypeOf i Is ComboBox) Then
                i.BackColor = Color.White
                i.Text = ""
            End If
        Next

    End Sub

    Sub Generate_StockID()

        'SQL Select Statement Searching for the highest UserID assciated with that user type'
        StockID = "SELECT MAX(StockID) FROM tblStock WHERE StockID LIKE '" & "ST" & "%%%%%' "
        'Data adapter defined'
        da = New OleDbDataAdapter(StockID, conn)
        'Data adapter told to fill the da   taset
        da.Fill(ds, "tblStock")
        'Defining the datatable
        dt = ds.Tables("tblStock")
        'Associating User ID to the first item on the first row of the datatabl whilst converting the data to a string'
        StockID = dt.Rows(0).Item(0).ToString

        'For defining the first student in the class;
        If StockID = "" Then
            'this is some "Fake Data" that is used to calculate the first User ID'
            StockID = "ST00000"
            MsgBox(StockID)
            Exit Sub
        End If

        'This strips off the User Type Identified'
        StockID = StockID.Substring(2, 5)
        'This removes the leading zero's from the stripped string'
        StockID = StockID.TrimStart("0"c)

        'This catches the User ID string if it is something like "S0000", after all the stripping and formating, would leave nothing, this code repairs it'
        If StockID = "" Then
            'Re-Defines the StockID to 0'
            StockID = 0
        End If

        'Converts the split string into an integer'
        StockID = CType(StockID, Integer) + 1

        'This block is responsible for adding the appropriate value onto the end of the user ID'
        'This converts the User ID to a string'
        StockID = CType(StockID, String)
        'Length check's used to identify the different tiers of the user id'
        If StockID.Length = 1 Then
            'This combines the User ID's componants into a full id'
            StockID = "ST" & "0000" & StockID
            'Length check's used to identify the different tiers of the user id'
        ElseIf StockID.Length = 2 Then
            'This combines the User ID's componants into a full id'
            StockID = "ST" & "000" & StockID
            'Length check's used to identify the different tiers of the user id'
        ElseIf StockID.Length = 3 Then
            'This combines the User ID's componants into a full id'
            StockID = "ST" & "00" & StockID
        ElseIf StockID.Length = 4 Then
            StockID = "ST" & "0" & StockID
        Else
            'This combines the User ID's componants into a full id'
            StockID = "ST" & StockID
        End If

        ds.Clear()

    End Sub

    Sub Save_stock_profile()

        Dim sizetotal As Integer = 0

        'this loop will cycle through each  control in the Size info groupbox that is either a text box or a combobox 
        'and will clear its content and reset is color formatting
        Dim i As Control
        For Each i In Main.gbSize_stock_information.Controls
            If (TypeOf i Is TextBox) Then
                If i.BackColor = Color.White Or i.BackColor = Color.Salmon Then
                    i.BackColor = Color.CornflowerBlue
                    MsgBox("Please Complete Profile")
                    Exit Sub
                Else
                    sizetotal = sizetotal + CInt(i.Text)
                End If
            End If
        Next

        'this loop will cycle through each  control in the main info groupbox that is either a text box or a combobox 
        'and will clear its content and reset is color formatting
        For Each i In Main.gbMain_stock_information.Controls
            If (TypeOf i Is TextBox) Then
                If i.BackColor = Color.White Or i.BackColor = Color.Salmon Then
                    i.BackColor = Color.CornflowerBlue
                    MsgBox("Please Complete Profile")
                    Exit Sub
                End If
            ElseIf (TypeOf i Is ComboBox) Then
                If i.BackColor = Color.White Or i.BackColor = Color.Salmon Then
                    i.BackColor = Color.CornflowerBlue
                    MsgBox("Please Complete Profile")
                    Exit Sub
                End If
            End If
        Next

        Generate_StockID()

        'this defines the insert statement for inserting stock data into the database
        Dim Stockinsert As String = "INSERT INTO tblStock ([StockID],[StockName],[StockStyle],[StockSupplier],[StockPrice],[StockColour],[34],[345],[35],[355],[36],[365],[37],[375],[38],[385],[39],[395],[40],[405],[41],[415],[42],[425],[43],[435],[44],[initial_total]) VALUES (@StockID,@Name,@Style,@Supplier,@price,@Colour,@34,@345,@35,@355,@36,@365,@37,@375,@38,@385,@39,@395,@40,@405,@41,@415,@42,@425,@43,@435,@44,@sizetotal)"
        Dim StockinsertCommand As New OleDbCommand
        With StockinsertCommand
            .CommandText = Stockinsert
            'this block of parameters matches the stock id, main information and size information with the variables in the sql statement
            .Parameters.AddWithValue("@StockID", StockID)
            .Parameters.AddWithValue("@Name", Main.txtStockName.Text)
            .Parameters.AddWithValue("@Style", Main.cbStyle2.Text)
            .Parameters.AddWithValue("@Supplier", Main.cbSupplier2.Text)
            .Parameters.AddWithValue("@Price", Main.txtStockPrice.Text)
            .Parameters.AddWithValue("@Colour", Main.cbColour2.Text)
            .Parameters.AddWithValue("@34", Main.txt34.Text)
            .Parameters.AddWithValue("@345", Main.txt345.Text)
            .Parameters.AddWithValue("@35", Main.txt35.Text)
            .Parameters.AddWithValue("@355", Main.txt355.Text)
            .Parameters.AddWithValue("@36", Main.txt36.Text)
            .Parameters.AddWithValue("@365", Main.txt365.Text)
            .Parameters.AddWithValue("@37", Main.txt37.Text)
            .Parameters.AddWithValue("@375", Main.txt375.Text)
            .Parameters.AddWithValue("@38", Main.txt38.Text)
            .Parameters.AddWithValue("@358", Main.txt385.Text)
            .Parameters.AddWithValue("@39", Main.txt39.Text)
            .Parameters.AddWithValue("@395", Main.txt395.Text)
            .Parameters.AddWithValue("@40", Main.txt40.Text)
            .Parameters.AddWithValue("@405", Main.txt405.Text)
            .Parameters.AddWithValue("@41", Main.txt41.Text)
            .Parameters.AddWithValue("@415", Main.txt415.Text)
            .Parameters.AddWithValue("@42", Main.txt42.Text)
            .Parameters.AddWithValue("@425", Main.txt425.Text)
            .Parameters.AddWithValue("@43", Main.txt43.Text)
            .Parameters.AddWithValue("@435", Main.txt435.Text)
            .Parameters.AddWithValue("@44", Main.txt44.Text)
            .Parameters.AddWithValue("@sizetotal", sizetotal)
            'this references the database connection string and executes the nonquery
            .Connection = conn
            .ExecuteNonQuery()
        End With

        'this notifies the user that the stock was added correctly
        MsgBox("Stock Item: " & Main.txtStockName.Text & " Added Successfully as " & StockID & " !")

        'this enables the generate label button
        Main.btnGenLabel.Enabled = True

    End Sub




End Class