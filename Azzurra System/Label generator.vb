
Module Label_generator

    'this declares a counting variable
    Public ammountcounter As Integer = 0

    'this creates 7 column 2d array with as many rows as their are individuals items of stock
    Public grid(ammountcounter, 6) As String


    'this subroutine is used to populate the label datagridview with individual label records
    Sub populate_Grid()

        'this bloc declares all the variables for stock price etc...
        Dim stockname As String = Main.txtStockName.Text
        Dim stockprice As String = Main.txtStockPrice.Text
        Dim stockStyle As String = Main.cbStyle2.Text
        Dim stockColour As String = Main.cbColour2.Text
        Dim stocksupplier As String = Main.cbSupplier2.Text

        ammountcounter = 0

        'this declares a row
        Dim row As Integer

        'this declares a counting variable
        Dim counter As Integer = 0

        'this declares i as a control
        Dim i As Control

        'this loop will loop through each size control on the information control
        For Each i In Main.gbSize_stock_information.Controls
            'this conditionw will check if the control is a textbox
            If (TypeOf i Is TextBox) Then
                'this adds one to the counter
                counter = counter + 1
                'this condition checks if the control content is empty
                If i.Text = "" Then
                    'if the backcolour doesnt equal the colour yellow green
                ElseIf i.BackColor <> Color.YellowGreen Then
                    'this prompts the user that the stock profile contains incorrect data
                    MsgBox("Error Value in stock profile")
                    Exit Sub
                Else
                    'this adds the content value of the current textbox to the current total ammount of labels
                    ammountcounter = CInt(i.Text) + ammountcounter
                End If
            End If
        Next

        'this declares a 2 dimensional array
        Dim sizes(20, 1) As String

        'this block fills each item of the first column of the sizes array with each size in the range
        sizes(0, 1) = "34"
        sizes(1, 1) = "34.5"
        sizes(2, 1) = "35"
        sizes(3, 1) = "35.5"
        sizes(4, 1) = "36"
        sizes(5, 1) = "36.5"
        sizes(6, 1) = "37"
        sizes(7, 1) = "37.5"
        sizes(8, 1) = "38"
        sizes(9, 1) = "38.5"
        sizes(10, 1) = "39"
        sizes(11, 1) = "39.5"
        sizes(12, 1) = "40"
        sizes(13, 1) = "40.5"
        sizes(14, 1) = "41"
        sizes(15, 1) = "41.5"
        sizes(16, 1) = "42"
        sizes(17, 1) = "42.5"
        sizes(18, 1) = "43"
        sizes(19, 1) = "43.5"
        sizes(20, 1) = "44"

        'this algorithm determines how many pairs of shoes per size there are in each run of a shoe but locating textboxes in the form, passing its name through a substring 
        'that will format the string in to a value that matches the shoe size itself, then locating a pertient record in the sizes array and combining its value with its relative 
        'size'
        Dim sizesfinal(ammountcounter)
        counter = 0

        'this defines j as a control variable
        Dim j As Control

        'this loops through each size control
        For Each j In Main.gbSize_stock_information.Controls
            'this condition checks if the current control is a textbox
            If (TypeOf j Is TextBox) Then
                'this defines name as a null string
                Dim name As String = ""
                'this condition checks the length of the name of the current control is its 5
                If j.Name.Length = 5 Then
                    'the name string is given a substringed version of the control name stripping off the txt part
                    name = j.Name.ToString.Substring(3, 2)
                    'this condition checks the length of the name of the current control is its 6
                ElseIf j.Name.Length = 6 Then
                    'the name string is given a substringed version of the control name stripping off the txt part
                    name = j.Name.ToString.Substring(3, 3)
                    'this inserts a full stop in the decimal of the size
                    name = name.Insert(2, ".")
                End If
                'this runs a loop 21 times
                For k As Integer = 0 To 20
                    'this condition checks to see if the name matches the current row in the 2d sizes array
                    If name = sizes(k, 1) Then
                        'this condition resets the control text to nothing
                        If j.Text = "" Then
                        Else
                            'this loops runs for the value of the current control minus 1
                            For l As Integer = 0 To CInt(j.Text) - 1
                                'this array will contain the final label sizes in their extracted tallys
                                sizesfinal(counter) = name

                                'this adds one to the counter variable
                                counter = counter + 1
                            Next
                        End If
                    End If
                Next
            End If
        Next

        ReDim grid(ammountcounter, 6)
        
        'this loops for as many times as there are items of stock minus 1
        For row = 0 To ammountcounter - 1
            'this sets the current rows column as the stock name variable
            grid(row, 0) = stockname
            'this sets the current rows column as the stock price variable
            grid(row, 1) = stockprice
            'this sets the current rows column as the stock style variable
            grid(row, 2) = stockStyle
            'this sets the current rows column as the stock colour variable
            grid(row, 3) = stockColour
            'this sets the current rows column as the stock supplier variable
            grid(row, 4) = stocksupplier
            'this sets the current rows column as the stock size variable
            grid(row, 5) = sizesfinal(row)
        Next

        'this resets the counter and row variable to 0
        row = 0
        counter = 0

        'this declares a new label table to fill all the data from the 2d array
        Dim Labeltable As New DataTable

        'this block will add each column of the 2d array to the datatable column
        Labeltable.Columns.Add("Stock Name")
        Labeltable.Columns.Add("Stock Price")
        Labeltable.Columns.Add("Stock Style")
        Labeltable.Columns.Add("Stock Colour")
        Labeltable.Columns.Add("Stock Supplier")
        Labeltable.Columns.Add("Stock Size")

        'this loops for every individual item of stock
        For outerIndex As Integer = 0 To ammountcounter
            'this declares a variable as new datarow pulling data from the labeltable and using it in a new row
            Dim newRow As DataRow = Labeltable.NewRow()
            'this loops through each item in the label array
            For innerIndex As Integer = 0 To 5
                'this declares a new row number as 0 - 5 where pulling data from every item in the row and for a specific row
                newRow(innerIndex) = grid(outerIndex, innerIndex)
            Next
            'this then adds the new row data to the datatable
            Labeltable.Rows.Add(newRow)
        Next

        'this inserts the labeltable data into the datagridview
        Main.dgvLabel_gen.DataSource = Labeltable
        Main.dgvLabel_gen.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


        'this resets the ammount counter
        ammountcounter = 0
    End Sub

End Module
