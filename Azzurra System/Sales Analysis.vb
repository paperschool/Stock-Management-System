'these imports handle the database connectivity and also handle the chart printing
Imports System.Data.OleDb
Imports System.Drawing

Module Sales_Analysis

    'these variables are used when data is being sourced from the database
    Dim salesanalysisda As OleDbDataAdapter
    Dim salesanalysisdt As DataTable
    Dim salesanalysisds As DataSet = New DataSet

    'this subroutine populates the navigation datagridview for product styles
    Sub populate_navigation_analysis()

        'this empties the dataset
        salesanalysisds.Clear()

        'this sql string is used to pull any presets whose category is product style
        Dim fill_navigation_styles As String = "SELECT * FROM tblPresets WHERE PresetCategory = '" & "Style" & "' "

        'this defines uses the sql query and connection string
        salesanalysisda = New OleDbDataAdapter(fill_navigation_styles, conn)
        'this datadapter is filled with the dataset declared earlier
        salesanalysisda.Fill(salesanalysisds, "styles")
        'this datatable is filled with data pulled from the dataset
        salesanalysisdt = salesanalysisds.Tables("styles")

        'this block is used to fill the navigation datagrid view with the data in the dataset
        With Main.dgvStock_Ana_Style
            .AutoGenerateColumns = True
            .DataSource = salesanalysisds
            .DataMember = "styles"
        End With

        'this sets some properties for the style navigation datagridview, such as allowing all the rows to auto size themselves, hiding unused columns or auto sizing columns
        Main.dgvStock_Ana_Style.AutoResizeRows()
        Main.dgvStock_Ana_Style.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Main.dgvStock_Ana_Style.Columns(0).Visible = False
        Main.dgvStock_Ana_Style.Columns(2).Visible = False

    End Sub

    'this subroutine populates the stock sale data grid view with any sales belonging to that product category
    Sub populate_sales_grid_style_history(style)

        'this variable is used to generate the new dataset
        Dim salesanalysisds5 As DataSet = New DataSet

        'this sql query is used to select any sales where the stock style equals the style string
        Dim fill_stock_information As String = "SELECT * FROM tblSales WHERE StockStyle ='" & style & "' "
        'this defines uses the sql query and connection string
        salesanalysisda = New OleDbDataAdapter(fill_stock_information, conn)
        'this datadapter is filled with the dataset declared earlier
        salesanalysisda.Fill(salesanalysisds5, "stocksales")
        'this datatable is filled with data pulled from the dataset
        salesanalysisdt = salesanalysisds5.Tables("stocksales")

        'this with statement uses the stock sale datagridview to fill it with data
        With Main.dgvSelected_stock_sales
            .AutoGenerateColumns = True
            .DataSource = salesanalysisds5
            .DataMember = "stocksales"
        End With

        'this formats the row size and column size automatically
        Main.dgvSelected_stock_sales.AutoResizeRows()
        Main.dgvSelected_stock_sales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    'this subroutine is used to populate the names in the analysis navigation
    Sub populate_navigation_names_ana(i)

        'this defines style as a string
        Dim style As String
        'this defines the dataset 
        Dim salesanalysisds2 As DataSet = New DataSet

        'this defines style as the value of the selected cell in the datagrid view
        style = Main.dgvStock_Ana_Style.Item(1, i).Value()

        'this runs the populate style sales history grid
        populate_sales_grid_style_history(style)

        'this catches a null selection in the style grid and exits the subroutine 
        If style = "" Then
            Exit Sub
        End If

        'this clears the sales analysis dataset
        salesanalysisds2.Clear()

        'this selection statement selects all stock that matches the currently selected stock style 
        Dim fill_navigation_names As String = "SELECT * FROM tblStock WHERE StockStyle = '" & style & "' "
        'this defines the sales analysis dataadapter 
        salesanalysisda = New OleDbDataAdapter(fill_navigation_names, conn)
        'this fills the data set uaing the dataadapter
        salesanalysisda.Fill(salesanalysisds2, "names")
        'this defines the datatable 
        salesanalysisdt = salesanalysisds2.Tables("names")

        'this uses the stock name datagridview
        With Main.dgvStock_Ana_Name
            .AutoGenerateColumns = True
            'this sources the dataset to be used as data
            .DataSource = salesanalysisds2
            .DataMember = "names"
        End With

        'this loops through every column in the datatable
        For x As Integer = 0 To salesanalysisds2.Tables("names").Columns.Count - 1
            'this condition checks whether x = 1, 0 or 5
            If x = 1 Or x = 0 Or x = 5 Then
            Else
                'this hides any column that wasnt caught in the previous condition
                Main.dgvStock_Ana_Name.Columns(x).Visible = False
            End If
        Next

        'this formats the width of the fields and rows in the dgv
        Main.dgvStock_Ana_Name.AutoResizeRows()
        Main.dgvStock_Ana_Name.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


        'For x As Integer = 0 To salesanalysisds2.Tables("names").Rows.Count - 1
        '    Dim Stocklevel As String
        '    Stocklevel = Main.dgvStock_Ana_Name.Item(0, x).Value()
        '    check_sizes(Stocklevel)
        'Next

    End Sub

    'check size subroutine potentially used in the next version of the system
    'Sub check_sizes(stocklevel)

    '    Dim stockcount As Integer = 0
    '    Dim salesanalysisds As DataSet = New DataSet
    '    Dim checkstocklevel As String = "SELECT * FROM tblStock WHERE StockID = '" & stocklevel & "' "
    '    salesanalysisda = New OleDbDataAdapter(checkstocklevel, conn)
    '    salesanalysisda.Fill(salesanalysisds, "stocklevel")
    '    salesanalysisdt = salesanalysisds.Tables("stocklevel")

    '    For x As Integer = 0 To salesanalysisds.Tables("stocklevel").Rows.Count - 1
    '        For i As Integer = 6 To salesanalysisds.Tables("stocklevel").Columns.Count - 1
    '            stockcount = stockcount + CInt(salesanalysisds.Tables("stocklevel").Rows(x).Item(i).ToString)
    '        Next

    '        If stockcount < (salesanalysisds.Tables("stocklevel").Rows(x).Item(27) * 0.2) Then
    '            Main.dgvStock_Man_Name.Rows(x).DefaultCellStyle.BackColor = Color.Salmon
    '        End If

    '        stockcount = 0
    '    Next

    'End Sub

    'this subroutine fills the stock sales data grid 
    Sub populate_Stock_sales_Grid(i, a)

        'this defines the stock id and sales data set
        Dim stockid As String = ""
        Dim salesanalysisds3 As DataSet = New DataSet

        'this prevents the subroutine from referencing null values when populating searched stock data
        If Not IsDBNull(Main.dgvStock_Ana_Name.Item(0, i).Value()) Then
        ElseIf Not IsDBNull(Main.dgvStock_Ana_Name.Item(0, i).Value()) Then
        Else
            Exit Sub
        End If

        'this condition checks which cell is selected in the data grid
        If a = 0 Or a = 1 Or a = 6 Then
            stockid = Main.dgvStock_Ana_Name.Item(0, i).Value()
        End If

        'this defines the stock id as a cell from the datagrid view
        'stockid = Main.dgvStock_Ana_Name.Item(0, i).Value()

        'this selection statement selects all stock sales records that matches the stock id
        Dim fill_stock_information As String = "SELECT * FROM tblSales WHERE StockID = '" & stockid & "' "
        'this defines the sales analysis dataadapter 
        salesanalysisda = New OleDbDataAdapter(fill_stock_information, conn)
        'this fills the data set uaing the dataadapter
        salesanalysisda.Fill(salesanalysisds3, "stocksales")
        'this defines the datatable 
        salesanalysisdt = salesanalysisds3.Tables("stocksales")


        'this uses the sales dgv to populate sales data
        With Main.dgvSelected_stock_sales
            .AutoGenerateColumns = True
            'this data adapter is used to populate the grid
            .DataSource = salesanalysisds3
            .DataMember = "stocksales"
        End With

        'this formats the sizes of the rows and columns
        Main.dgvSelected_stock_sales.AutoResizeRows()
        Main.dgvSelected_stock_sales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    'this subroutine is not used in this system
    Sub populate_sales_history_all()

        'dataset used when populating a sales table for all sales in the system
        Dim salesanalysisds4 As DataSet = New DataSet

        'this selection statement selects all stock sales records
        Dim fill_stock_information As String = "SELECT * FROM tblSales"
        'this defines the sales analysis dataadapter 
        salesanalysisda = New OleDbDataAdapter(fill_stock_information, conn)
        'this fills the data set using the data adapter
        salesanalysisda.Fill(salesanalysisds4, "stocksales")
        'this defines the datatable 
        salesanalysisdt = salesanalysisds4.Tables("stocksales")

        'this ustilises the sales datagrid view
        With Main.dgvSelected_stock_sales
            .AutoGenerateColumns = True
            'this data adapter is used to populate the grid
            .DataSource = salesanalysisds4
            'the data source used
            .DataMember = "stocksales"
        End With

        'this formats the sizes of the rows and columns
        Main.dgvSelected_stock_sales.AutoResizeRows()
        Main.dgvSelected_stock_sales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'this defines two variables, a substringged now variable holding only the month and year &
        'a now variable holding only the day
        Dim searchabledate As String = Now.ToString.Substring(2, 8)
        Dim searchableday As Integer = Now.ToString.Substring(0, 2)

        'this subtracts 1 from the searchable day variable
        searchableday = searchableday - 1

        'this defines the searchable date as the day and the month and year variable combined
        searchabledate = searchableday.ToString + searchabledate.ToString

        If searchabledate.Length < 10 Then
            searchabledate = "0" + searchabledate
        End If

        'this loops through each row in the datatable
        For Each row As DataRow In salesanalysisdt.Rows
            'variable defined to contain the date of sale on record
            Dim todayval As String = row.Item(6)

            'condtition used to identify if the sale date matches yesterdays date
            If todayval = searchabledate Then
                Console.Write("asd")
            Else
                'if it doesnt, it deletes the row from the datatable
                row.Delete()
            End If
        Next

        'this line will save any changes made to the data set
        'salesanalysisds4.AcceptChanges()

        'the edited data is then populated into the datagridview
        With Main.dgvHome_Sales
            .AutoGenerateColumns = True
            .DataSource = salesanalysisds4
            .DataMember = "stocksales"
        End With

        'this formats the sizes of the rows and columns
        Main.dgvHome_Sales.AutoResizeRows()
        Main.dgvHome_Sales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    'this sets a numeric value to the graph interval 
    Public graphinterval As Integer = 2
    'this subroutine contains all the settings for the chart population
    Sub graph_settings()

        'this checks if the interval textbox is empty
        If Main.txtgraphInterval.Text = "" Then
            Exit Sub
        End If

        'this try defines the interval with the text from the interval textbox and changes the colour of the textbox
        Try
            graphinterval = CInt(Main.txtgraphInterval.Text)
            Main.ActiveControl.BackColor = Color.White
            'this runs the subroutine that will populate chart data
            populate_chart()
        Catch ex As Exception
            'this catches an exception and converts the colour of the textbox to red
            Main.ActiveControl.BackColor = Color.Salmon
            Exit Sub
        End Try
    End Sub


    'this defines a bar count integer
    Public barcount As Integer = 4
    'this defines the modifier subroutine
    Sub barcountmodifier()

        'this checks for content in the textbox 
        If Main.txtBarcount.Text = "" Then
            Exit Sub
        End If

        'this try attempts to define the barcount value
        Try
            barcount = CInt(Main.txtBarcount.Text)
        Catch ex As Exception
            'this catches an error and changes the textbox colour
            Main.ActiveControl.BackColor = Color.Salmon
            Exit Sub
        End Try

        'this defines the barcount as its value minus 1
        barcount = barcount - 1

        'this condition checks if the value is not under 4
        If CInt(Main.txtBarcount.Text) < 4 Then
            'resets the value and chanegs the colour to red
            Main.ActiveControl.BackColor = Color.Salmon
            barcount = 4
        Else
            Main.ActiveControl.BackColor = Color.White
        End If

        'populates the charts
        populate_chart()

    End Sub

    'this defines the current date as the now modules time and date
    Public curdate As String = Now

    'this populates the charts
    Sub populate_chart()
        'this clears the points on both charts and home chart
        Main.chSales_best.Series(0).Points.Clear()
        Main.chSales_worst.Series(0).Points.Clear()
        Main.chHome.Series(0).Points.Clear()
        Main.chHome.Series(1).Points.Clear()

        'this try will attempt to substring the current data
        Try
            curdate = curdate.Substring(2, 8)
        Catch ex As Exception

        End Try

        'this runs the month subroutine
        month_cycle()

        'this sets the textbox content to the current date
        Main.txtSalesmonth.Text = curdate

        'this sets all the graphs intervals to the previously defined values
        Main.chHome.ChartAreas(0).AxisY.Interval = graphinterval
        Main.chSales_best.ChartAreas(0).AxisY.Interval = graphinterval
        Main.chSales_worst.ChartAreas(0).AxisY.Interval = graphinterval

        'this attempt to create the charts
        Try
            Chartpopulation(curdate)
        Catch
        End Try

        'this then repopulates the navigation box for the stock styles
        populate_navigation_analysis()

    End Sub

    'this subroutine is used to chart the best selling and worst selling items stock items
    Sub Chartpopulation(curdate)

        'this sql statement is used to select all sale information from the sales table
        Dim stocksales As String = "SELECT * FROM tblSales "
        salesanalysisda = New OleDbDataAdapter(stocksales, conn)
        salesanalysisda.Fill(salesanalysisds, "stocksales")
        salesanalysisdt = salesanalysisds.Tables("stocksales")

        'validation that ends the subroutine if there is no data in the system
        If salesanalysisds.Tables("stocksales").Rows.Count = 0 Then
            'ends the subroutine
            Exit Sub
        End If

        'this loop, will loop through each row in the datatable
        For Each row As DataRow In salesanalysisdt.Rows
            'this defines a variable as the 7th column item of the current sale record being looked at
            Dim timestrip As String = row.Item(6)
            'this substrings that variable selecting only the month and year
            timestrip = timestrip.Substring(2, 8)
            'this then compares it to the curdate variable
            If timestrip = curdate Then
            Else
                'if it doesnt equal it then it deletes that row
                row.Delete()
            End If
        Next

        'this comits those changes made earlier
        salesanalysisds.AcceptChanges()

        'this defines a new array equalling the lenth of all the remaining rows 
        Dim sales(salesanalysisds.Tables("stocksales").Rows.Count - 1, 1) As String
        'this defines a new array equalling the lenth of all the remaining rows 
        Dim sales2(salesanalysisds.Tables("stocksales").Rows.Count - 1, 1) As String
        'this defines a boolean used to identify duplicates
        Dim located As Boolean = False
        'this is a counter used to count repeitions
        Dim locatedcounter = 0

        'this is a loop used to loop through each row in the sales table
        For i As Integer = 0 To salesanalysisds.Tables("stocksales").Rows.Count - 1
            'this resets the current boolean to false
            located = False
            'this loops through the sales table again
            For j As Integer = 0 To salesanalysisds.Tables("stocksales").Rows.Count - 1
                'this condition checks if the currently observed item matches the item being looked at in the 2d array
                If salesanalysisds.Tables("stocksales").Rows(i).Item(1).ToString = sales(j, 1) Then
                    'this changes the located boolean to true
                    located = True
                    'this exits the nested loop
                    Exit For
                End If
            Next
            'this checks whether the located boolean is false
            If located = False Then
                'if false the row item is added to a new index in the array
                sales(locatedcounter, 1) = salesanalysisds.Tables("stocksales").Rows(i).Item(1).ToString()
                sales2(locatedcounter, 1) = salesanalysisds.Tables("stocksales").Rows(i).Item(1).ToString()
                'one is then added to the located counter
                locatedcounter = locatedcounter + 1
            End If
        Next

        'this loops through the sales table
        For i As Integer = 0 To salesanalysisds.Tables("stocksales").Rows.Count - 1
            'this then loops through the sales table again
            For j As Integer = 0 To salesanalysisds.Tables("stocksales").Rows.Count - 1
                'this condition checks if the currently being observed value equals one in the 2d array
                If salesanalysisds.Tables("stocksales").Rows(i).Item(1).ToString = sales(j, 1) Then
                    'if yes, it adds one to the 0th column but same row in the 2d array
                    sales(j, 0) = sales(j, 0) + 1
                    sales2(j, 0) = sales2(j, 0) + 1
                End If
            Next
        Next

        'this defines a current value and current index variable
        Dim curval As Integer = 0
        Dim curindex As Integer = 0

        'this defines two arrays that are as long as the barcount
        Dim top5(barcount) As String
        Dim label(barcount) As String

        'this sets the curval equalling the 0th column and row of the sales array
        curval = sales(0, 0)

        'this loops round as many times as the barcount value equals 
        For x As Integer = 0 To barcount
            'this loops round half the length of the 2d array (half as it counts each row twice due to multiple columns)
            For i As Integer = 0 To (sales.Length / 2) - 1

                If sales2(i, 0) Is Nothing Then
                ElseIf sales2(i, 0) = 0 Then
                Else
                    'this compares the curval seeing if its greater then or equal to the row i, 0th column of the sales array
                    If curval >= sales(i, 0) Then
                    Else
                        'if not, it sets the current value to that array item
                        curval = sales(i, 0)
                        'this then sets the current index to that row
                        curindex = i
                    End If
                End If
            Next
            'this sets the first item of this array to that current value
            top5(x) = curval
            'this sets the first item of this array to the name associated with the current value
            label(x) = sales(curindex, 1).ToString
            'this sets that item of the sale array to 0
            sales(curindex, 0) = Nothing
            'this resets the current index and resets the value of the current value to the first item of the sales array
            curindex = 0
            curval = sales(0, 0)
        Next

        'this checks if the useredited boolean has been edited
        If Main.useredited = False Then
            'this condition checks to see if the highest value exceeds a max value, if so, sets the current interval to a higher number to account for this
            If top5(0) >= 80 Then
                'this sets the graph interval to 10
                graphinterval = 10
                'this then applies that property to all the graphs in question
                Main.chHome.ChartAreas(0).AxisY.Interval = graphinterval
                Main.chSales_best.ChartAreas(0).AxisY.Interval = graphinterval
                Main.chSales_worst.ChartAreas(0).AxisY.Interval = graphinterval
            End If
        End If

        'note this is simply the same as the algorithm above except has been reverse in order to produce a bottom five array

        'this defines a current value and current index variable
        Dim curval2 As Integer = 0
        Dim curindex2 As Integer = 0

        'this defines two arrays that are as long as the barcount
        Dim bot5(barcount) As String
        Dim label2(barcount) As String

        'this sets the curval equalling the 0th column and row of the sales array
        curval2 = sales2(0, 0)

        'this loops round as many times as the barcount value equals 
        For x As Integer = 0 To barcount
            'this loops round half the length of the 2d array (half as it counts each row twice due to multiple columns)
            For i As Integer = 0 To (sales2.Length / 2) - 1

                If sales2(i, 0) Is Nothing Then
                ElseIf sales2(i, 0) = 0 Then
                Else
                    'this compares the curval seeing if its less then or equal to the row i, 0th column of the sales array
                    If curval2 <= sales2(i, 0) Then
                    Else
                        'if not, it sets the current value to that array item
                        curval2 = sales2(i, 0)
                        'this then sets the current index to that row
                        curindex2 = i
                    End If
                End If
            Next
            'this sets the first item of this array to that current value
            bot5(x) = curval2
            'this sets the first item of this array to the name associated with the current value
            label2(x) = sales2(curindex2, 1).ToString
            'this loop will remove any ranked data from the array once it has been used
            For i As Integer = 0 To (sales2.Length / 2) - 1
                'this condition checks if the currently being viewed index of the array matches the current label in the label array
                If sales2(i, 1) = label2(x) Then
                    'these two lines reset the value of the array 
                    sales2(i, 0) = Nothing
                    sales2(i, 1) = Nothing
                    Exit For
                End If
            Next
            'sales2(curindex2, 1) = Nothing
            curindex2 = 0
            'this resets the current index and resets the value of the current value to the first item of the sales array
            curval2 = sales2(0, 0)
            'this loop is designed to counter act the problem of locating a remaining number that isnt 0 or the first item of the array which may n
            'now be a null value too
            For i As Integer = 0 To (sales2.Length / 2) - 1
                'this condition checks if the current value selected is nothing, if so
                'the value is asigned to the next item in the list
                If curval2 = Nothing Then
                    curval2 = sales2(i, 0)
                    'this also sets the current index to the same value as the newly appened current value
                    curindex2 = i
                End If
            Next
        Next

        'this checks if the useredited boolean has been edited
        If Main.useredited = False Then
            'this condition checks to see if the highest value exceeds a max value, if so, sets the current interval to a higher number to account for this
            If bot5(0) >= 80 Then
                'this sets the graph interval to 10
                graphinterval = 10
                'this then applies that property to all the graphs in question
                Main.chHome.ChartAreas(0).AxisY.Interval = graphinterval
                Main.chSales_best.ChartAreas(0).AxisY.Interval = graphinterval
                Main.chSales_worst.ChartAreas(0).AxisY.Interval = graphinterval
            End If
        End If

        'this then loops for the value of the barcount
        For i As Integer = 0 To barcount
            'this applies the label and value to all charts affected to actually show the data found
            Main.chHome.Series("Best Selling").Points.AddXY(label(i), top5(i))
            Main.chHome.Series("Worst Selling").Points.AddXY(label2(i), bot5(i))
            Main.chSales_best.Series("Best Selling").Points.AddXY(label(i), top5(i))
            Main.chSales_worst.Series("Worst Selling").Points.AddXY(label2(i), bot5(i))
        Next

        'this reruns the navigation analysis subroutine
        populate_navigation_analysis()

    End Sub

    'this subroutine will at form load, and when the foward or back buttons are pressed, this formats the month and year string correctly ensuring it will find data correctly
    Sub month_cycle()

        'this defines two variables that will account for the months and the years
        Dim monthedit As Integer = 0
        Dim yearedit As Integer = 0

        'this redefines the current curdate with a substring module
        curdate = curdate.Substring(1, curdate.Length - 1)

        'this determines which button was pressed, forward or backwards
        If Main.direction = "forward" Then
            'this substrings the year edit value as the curdate value grabbing only the year
            yearedit = curdate.Substring(curdate.Length - 4, 4)
            'this substrings the monthedit value as the curfate value grabbing the day and month
            monthedit = curdate.Substring(0, curdate.Length - 5)
            'this condition checks whether the month december
            If monthedit = 12 Then
                'this then sets the month to 1
                monthedit = 1
                'and adds one to the year edit
                yearedit = yearedit + 1
                'then formats the currentdate variable combining all the new data
                curdate = "/" + "0" + CStr(monthedit) + "/" + CStr(yearedit)
                'if the month is september
            ElseIf monthedit = 9 Then
                'it then adds one to the month
                monthedit = monthedit + 1
                'and formats the currentdate variable combining all the new data
                curdate = "/" + CStr(monthedit) + "/" + CStr(yearedit)
                'this checks if the month is under october
            ElseIf monthedit < 10 Then
                'this adds one to the month
                monthedit = monthedit + 1
                'and formats the current date variable combining all the new data
                curdate = "/" + "0" + CStr(monthedit) + "/" + CStr(yearedit)
            Else
                'this increments the month value
                monthedit = monthedit + 1
                'and formats the current date string
                curdate = "/" + CStr(monthedit) + "/" + CStr(yearedit)
            End If
            'this resets the direction value
            Main.direction = ""
            'and populates the textbox with the current date value
            Main.txtSalesmonth.Text = curdate
            'this checks if the back direction button is pressed
        ElseIf Main.direction = "back" Then
            'this defines the year and month edit variables
            yearedit = curdate.Substring(curdate.Length - 4, 4)
            monthedit = curdate.Substring(0, curdate.Length - 5)

            'this checks if the month is january
            If monthedit = 1 Then
                'this defines the month as 12
                monthedit = 12
                'this defines the year as the year minus 1
                yearedit = yearedit - 1
                'this formats the currentdate variable
                curdate = "/" + CStr(monthedit) + "/" + CStr(yearedit)
                'this condition checks if the month is lower then october
            ElseIf monthedit < 10 Then
                'this increments the month
                monthedit = monthedit - 1
                'and formats the current date variable
                curdate = "/" + "0" + CStr(monthedit) + "/" + CStr(yearedit)
                'this condition checks if the month equals october
            ElseIf monthedit = 10 Then
                'this increments the month
                monthedit = monthedit - 1
                'and formats the current date variable
                curdate = "/" + "0" + CStr(monthedit) + "/" + CStr(yearedit)
            Else
                'this increments the month
                monthedit = monthedit - 1
                'and formats the current date string
                curdate = "/" + CStr(monthedit) + "/" + CStr(yearedit)
            End If
            'this resets the direction 
            Main.direction = ""
            'and inserts the currentdate value into the textbox
            Main.txtSalesmonth.Text = curdate
            'if no direction is selected
        ElseIf Main.direction = "" Then
            'the current date is formated
            curdate = "/" + curdate
            'and is inserted into the textbox
            Main.txtSalesmonth.Text = curdate
        End If

    End Sub

    'this subroutine will run when the print charts button is pressed 
    Sub print_chart_data()

        'this condition checks if a directory exists on the computer 
        If (Not System.IO.Directory.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\")) Then
            'if it doesnt, it creates the new directory
            System.IO.Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\")
        End If

        'these variables store the locations of all the three different charts to be outputted
        Dim print_date_best As String = ""
        Dim print_date_worst As String = ""
        Dim print_date_Home As String = ""

        'this loop replaces all "/" with "-", this means it wont confusion the directory property
        For Each i As Char In curdate
            'this condition checks if the current iteration's character is a "/"
            If i = "/" Then
                'this replaces the current character with the safe version
                curdate = curdate.Replace(i, "-")
            End If
        Next

        'these lines format each of the strings for the three different images to be saved and stores that in the three variables
        print_date_best = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\" + curdate + "_Best" + ".png"
        print_date_worst = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\" + curdate + "_Worst" + ".png"
        print_date_Home = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\" + curdate + "_Main" + ".png"

        'these lines run the save image module and save the control as a png in the location described above
        Main.chSales_best.SaveImage(print_date_best, System.Drawing.Imaging.ImageFormat.Png)
        Main.chSales_worst.SaveImage(print_date_worst, System.Drawing.Imaging.ImageFormat.Png)
        Main.chHome.SaveImage(print_date_Home, System.Drawing.Imaging.ImageFormat.Png)

        'this process start will open an explorer window at the save location
        Process.Start(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\charts\")

    End Sub

End Module
