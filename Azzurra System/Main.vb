'importing the system database control module and timer module
Imports System.Data.OleDb
Imports System.Timers

Public Class Main

    'this is the load subroutine, its function is to run what ever code is within the subroutine when the form loads
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'this subroutine will calculate chart output
        populate_chart()

        'this subroutine will dynamically layout the form elements on all the pages based on form size
        home_arrange()

        'this will fill all the preset comboboxes with current stock constants
        populate_Presets()

        'this will enable the timer object
        Timer1.Enabled = True

        'this will make the stockid textfield read only
        txtStockID.Enabled = False

        'this will disable the generate label button
        btnGenLabel.Enabled = False

        'this engages the autocomplete subroutine for stock entry
        autocomplete()

        'these subroutines will fill datagridviews with data pulled from the database
        'about yesterdays sales, and stock style types
        populate_navigation_management()
        populate_sales_history_all()
        populate_navigation_analysis()


    End Sub

    'this subroutine will run a size check when the form size is altered ensuring its minium functioning size is met
    Private Sub MyButton1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        'this condition will check to ensure the width and height of the form are larger then the minimum requirement
        If Me.Width < 1290 Or Me.Height < 800 Then
            Me.Width = 1320
            Me.Height = 820
            Me.CenterToScreen()
        End If

        'this calls the dynamic arrange subroutine
        home_arrange()

    End Sub


    'this subroutine will dynamically layout the form elements on all the pages based on form size
    'most form elements have an absolute start point where as others will change dependent on the size of the form and how far it is stretched
    'other things like width and height of pannels is also controlled here
    Sub home_arrange()

        'Console.WriteLine(Me.Width & " " & Me.Height)

        If Me.tc_main.SelectedIndex = 0 Then
            '----------Home Page Dynamic controls----------'

            home_content_1.Location = New Point(Me.tc_main.Width / 2, 5)
            home_content_1.Width = Me.tc_main.Width / 2 - 13
            home_content_1.Height = Me.tc_main.Height / 2 - 31

            home_content_2.Location = New Point(Me.tc_main.Width / 2, Me.tc_main.Height / 2 - 21)
            home_content_2.Width = Me.tc_main.Width / 2 - 13
            home_content_2.Height = Me.tc_main.Height / 2 - 31

            home_content_4.Location = New Point(5, 5)
            home_content_4.Width = (Me.tc_main.Width / 2) - 10
            home_content_4.Height = (Me.tc_main.Height) - 57

            dgvHome_Sales.Location = New Point(0, 0)
            dgvHome_Sales.Width = (Me.tc_main.Width / 2) - 10
            dgvHome_Sales.Height = (Me.tc_main.Height) - 57

        End If

        '----------Stock Entry Dynamic controls----------'

        Stock_content_1.Location = New Point(5, 5)
        Stock_content_1.Width = Me.tc_main.Width / 2 - 10
        Stock_content_1.Height = Me.tc_main.Height - 57

        Stock_content_2.Location = New Point(Me.tc_main.Width / 2, 5)
        Stock_content_2.Width = Me.tc_main.Width / 2 - 13
        Stock_content_2.Height = Me.tc_main.Height - 57

        dgvLabel_gen.Location = New Point(5, 51)
        dgvLabel_gen.Width = Me.tc_main.Width / 2 - 23
        dgvLabel_gen.Height = Me.Stock_content_2.Height - 105

        gbMain_stock_information.Location = New Point(5, 5)
        gbMain_stock_information.Width = Me.tc_main.Width / 2 - 20
        gbMain_stock_information.Height = 220
        'gbMain_stock_information.BackColor = Color.FromArgb(19, 166, 221)


        gbSize_stock_information.Location = New Point(5, 225)
        gbSize_stock_information.Width = Me.tc_main.Width / 2 - 20
        gbSize_stock_information.Height = 400

        btnGenLabel.Location = New Point(5, 5)
        btnPrint_label.Location = New Point(5, Me.Stock_content_2.Height - 49)

        '----------Stock Levels dynamic controls----------'

        dgvStock_Man_Style.Location = New Point(5, 5)
        dgvStock_Man_Style.Width = Me.tc_main.Width / 6
        dgvStock_Man_Style.Height = Me.tc_main.Height / 3

        dgvStock_Man_Name.Location = New Point(5, Me.tc_main.Height / 3 + 7)
        dgvStock_Man_Name.Width = Me.tc_main.Width / 6
        dgvStock_Man_Name.Height = Me.tc_main.Height / 3 * 2 - 60

        Stock_level_content_1.Location = New Point(Me.tc_main.Width / 6 + 9, 5)
        Stock_level_content_1.Width = (Me.tc_main.Width / 48 * 21)
        Stock_level_content_1.Height = Me.tc_main.Height - 58

        Stock_level_content_2.Location = New Point((Me.tc_main.Width / 48 * 21) + Me.tc_main.Width / 6 + 14, 5)
        Stock_level_content_2.Width = (Me.tc_main.Width / 24) * 9
        Stock_level_content_2.Height = Me.tc_main.Height - 58

        dgvExcel.Location = New Point(5, 42)
        dgvExcel.Width = Stock_level_content_2.Width - 10
        dgvExcel.Height = Stock_level_content_2.Height - 48

        

        btnOpen_excel.Location = New Point(5, 5)

        '----------Stock Analysis dynamic controls----------'

        dgvStock_Ana_Style.Location = New Point(5, 5)
        dgvStock_Ana_Style.Width = CInt(Me.tc_main.Width / 6)
        dgvStock_Ana_Style.Height = CInt(Me.tc_main.Height / 3)

        dgvStock_Ana_Name.Location = New Point(5, CInt(Me.tc_main.Height / 3 + 7))
        dgvStock_Ana_Name.Width = CInt(Me.tc_main.Width / 6)
        dgvStock_Ana_Name.Height = CInt(Me.tc_main.Height / 3 * 2 - 60)

        dgvSelected_stock_sales.Location = New Point(CInt(Me.tc_main.Width / 6 + 9), 5)
        dgvSelected_stock_sales.Width = CInt((Me.tc_main.Width / 12 * 5))
        dgvSelected_stock_sales.Height = Me.tc_main.Height - 58

        Stock_Analysis_Content_2.Location = New Point(Me.tc_main.Width / 12 * 7 + 14, 5)
        Stock_Analysis_Content_2.Width = Me.tc_main.Width / 12 * 5 - 27
        Stock_Analysis_Content_2.Height = Me.tc_main.Height - 58

        chSales_best.Location = New Point(0, 30)
        chSales_best.Width = Me.tc_main.Width / 12 * 5 - 27
        chSales_best.Height = Stock_Analysis_Content_2.Height / 2 - 15

        chSales_worst.Location = New Point(0, Stock_Analysis_Content_2.Height / 2)
        chSales_worst.Width = Me.tc_main.Width / 12 * 5 - 27
        chSales_worst.Height = Stock_Analysis_Content_2.Height / 2

    End Sub

    'this subroutine is an event that fire every time the timer object ticks, this populates the timer labels giving
    'a second by second look at the current time
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'now being a build in function that refers to the current computer time
        lblTime.Text = CStr(Now)
        lblTime2.Text = CStr(Now)

    End Sub

    'this text changed event will fire if any of the stock size fields are edited'
    Public Sub Size_TextChanged(sender As Object, e As EventArgs) Handles txt34.TextChanged, txt345.TextChanged, txt35.TextChanged, txt355.TextChanged, txt36.TextChanged, _
        txt365.TextChanged, txt37.TextChanged, txt375.TextChanged, txt38.TextChanged, txt385.TextChanged, txt39.TextChanged, txt395.TextChanged, txt40.TextChanged, txt405.TextChanged, txt41.TextChanged, _
        txt415.TextChanged, txt42.TextChanged, txt425.TextChanged, txt43.TextChanged, txt435.TextChanged, txt44.TextChanged

        'this will run the stock input validation on any of the active textboxes within the size range
        Stock_Input.Fill_sizes()

    End Sub

    'this text changes event fires when any of the stock edit size fields are edited
    Public Sub editSize_TextChanged(sender As Object, e As EventArgs) Handles txtedit34.TextChanged, txtedit345.TextChanged, txtedit35.TextChanged, txtedit355.TextChanged, txtedit36.TextChanged, _
        txtedit365.TextChanged, txtedit37.TextChanged, txtedit375.TextChanged, txtedit38.TextChanged, txtedit385.TextChanged, txtedit39.TextChanged, txtedit395.TextChanged, txtedit40.TextChanged, txtedit405.TextChanged, txtedit41.TextChanged, _
        txtedit415.TextChanged, txtedit42.TextChanged, txtedit425.TextChanged, txtedit43.TextChanged, txtedit435.TextChanged, txtedit44.TextChanged

        'this will run the stock edi input validation on any of the active textboxes within the size range
        validate_edited_sizes()

    End Sub

    'this event fires when the text inside the edited main information objects changes
    Private Sub Mainedit_textchanged(sender As Object, ByVal e As EventArgs) Handles cbstockcolouredit.TextChanged, cbStockstyleedit.TextChanged, cbstocksupplieredit.TextChanged, _
        txtStocknameedit.TextChanged, txtstockpriceedit.TextChanged

        'this subroutine will validate any of the text within those active textboxes
        validate_edited_information()

    End Sub

    'this event fires when the text inside the main information objects changes
    Private Sub Main_textchanged(sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbColour.KeyDown, cbColour2.KeyDown, cbStyle.KeyDown, cbStyle2.KeyDown, _
        cbSupplier.KeyDown, cbSupplier2.KeyDown

        'this subroutine will check if the correct key press was entered, if so it will add presets to the database
        main_information_updater(e)

    End Sub

    'this event fires when the text inside the main information objects changes
    Private Sub cbStyle_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cbStyle.TextChanged, cbStyle2.TextChanged, cbColour.TextChanged, cbColour2.TextChanged, _
        cbSupplier.TextChanged, cbSupplier2.TextChanged, txtStockName.TextChanged, txtStockPrice.TextChanged

        'this subroutine will validate any information in the main information controls
        Stock_Input.Main_information_validation()
        'autocomplete()

    End Sub

    'this subroutine will clear any stock inputs in the stock add page, when the button is fired
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        'this subroutine will cycle through any controls on the stock input page and clear the stock input controls
        Stock_Input.Clear_stock_inputs()

    End Sub

    'this event will fire when the linear run button is clicked
    Private Sub btnLinear_run_Click(sender As Object, e As EventArgs) Handles btnLinear_run.Click

        'this subroutine will fill all the sizes in the size page in a linear order fashion
        Stock_Input.Linear_Run()

    End Sub

    'this event will fire when the save button is pressed
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        'this subroutine will lead on to the saveing of the stock data to the database
        Stock_Input.Save_stock_profile()

        'this subroutine will repopulate the datagridview control for stock styles on the navigation page
        populate_navigation_management()

    End Sub

    'this event will fire when the generate labels button is pressed
    Private Sub btnGenLabel_Click(sender As Object, e As EventArgs) Handles btnGenLabel.Click

        populate_Grid()

    End Sub

    'this event will fire when the selection on the stock style datagridview control is changed
    Private Sub dgvStock_Man_style_selectedindexchanged(sender As Object, e As EventArgs) Handles dgvStock_Man_Style.Click

        'this defined i as an integer
        Dim i As Integer

        'this sets i to be the current row number
        i = dgvStock_Man_Style.CurrentRow.Index

        'this will populate stock name datagridview with any data that matches that stock style criteria
        populate_navigation_names(i)

    End Sub

    'this subroutine will fire when a stock item is selected from the stock name
    Private Sub dgvStock_Man_name_selectedindexchanged(sender As Object, e As EventArgs) Handles dgvStock_Man_Name.Click

        'this defines i and a as an integer
        Dim i As Integer
        Dim a As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvStock_Man_Name.CurrentRow.Index
        a = dgvStock_Man_Name.CurrentCell.ColumnIndex

        'this passes i and a to the subroutine that will fill all the stock data into the form for editing
        populate_Stock_information(i, a)

    End Sub

    'this event will fire when the selection on the stock style datagridview control is changed
    Private Sub dgvStock_Ana_Style_Click(sender As Object, e As EventArgs) Handles dgvStock_Ana_Style.Click

        Try
            'this defined i as an integer
            Dim i As Integer

            'this sets i to be the current row number
            i = dgvStock_Ana_Style.CurrentRow.Index

            'this will populate stock name datagridview with any data that matches that stock style criteria
            populate_navigation_names_ana(i)
        Catch
        End Try

    End Sub

    'this event will fire when the selection on the stock style datagridview control is changed
    Private Sub dgvStock_Ana_Name_Click(sender As Object, e As EventArgs) Handles dgvStock_Ana_Name.Click

        'this defines i and a as an integer
        Dim i As Integer
        Dim a As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvStock_Ana_Name.CurrentRow.Index
        a = dgvStock_Ana_Name.CurrentCell.ColumnIndex

        'this passes i and a to the subroutine that will fill a datagridview with all sales for that current item of stock
        populate_Stock_sales_Grid(i, a)

    End Sub

    'this event will fire when the import sales document button is pressed
    Private Sub btnOpen_excel_Click_1(sender As Object, e As EventArgs) Handles btnOpen_excel.Click

        'this will run the import file subroutine
        Import_file()

    End Sub

    'this sub launches when the event calls it
    Sub logout()

        Dim result As Integer = MessageBox.Show("Are you sure you want to exit the system?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
        ElseIf result = DialogResult.Yes Then
            Me.Close()

            Login.txtpassword.Enabled = True
            Login.txtUsername.Enabled = True

            Login.center_position()
        End If

    End Sub

    'this event fires when the logout button is pressed
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click

        logout()

    End Sub

    'this event fires when the save edited stock button is pressed
    Private Sub FlatButton3_Click(sender As Object, e As EventArgs) Handles btnSaveEdit.Click

        'this will save the edited stock informatin to its original stock record
        save_edited_stock()

    End Sub

    'this event will fire when the clear edited information button is pressed
    Private Sub FlatButton4_Click(sender As Object, e As EventArgs) Handles btnCancelEdit.Click

        'this will cancel the edit and will clear all the stock controls
        Cancel_edited_stock()

    End Sub

    'this event fires when the tab control changes tab
    Private Sub tc_main_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tc_main.SelectedIndexChanged

        'runs this dynamic control sub
        home_arrange()

    End Sub

    'this variable is used to determine which way the calender is moving when cycling through the sales months
    Public direction As String = ""

    'this event fires when the back button is pressed
    Private Sub FlatButton2_Click(sender As Object, e As EventArgs) Handles btnBackMonth.Click

        direction = "back"
        populate_chart()

    End Sub

    'this event fires when the forward button is pressed
    Private Sub FlatButton5_Click(sender As Object, e As EventArgs) Handles btnForwardMonth.Click

        direction = "forward"
        populate_chart()

    End Sub

    'this keydown event will fire when the sale month textbox is changed
    Private Sub txtSalesmonth_Keydown(sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSalesmonth.KeyDown

        'this defines a boolean that will act as some validation for the date input box
        Dim datecheck As Boolean = False

        'this clears all the datapoints on the graph readying it for new data
        chSales_best.Series(0).Points.Clear()
        chSales_worst.Series(0).Points.Clear()
        chHome.Series(0).Points.Clear()
        chHome.Series(1).Points.Clear()

        'this defines the current date variable to the textbox data just entered by the user
        Sales_Analysis.curdate = txtSalesmonth.Text

        'this condition will check to make sure the length of the string entered is of a reasonable length
        If Sales_Analysis.curdate.Length = 8 Or Sales_Analysis.curdate.Length = 7 Then

            'this character loop will check the date entered to ensure only numbers and the "/" symbol is used
            For Each i As Char In Sales_Analysis.curdate
                If Char.IsNumber(i) Then
                ElseIf i = "/" Then
                Else
                    'changes the boolean value to state that an error was made
                    datecheck = True
                    'provides a visual cue to the user that an error was made
                    Me.ActiveControl.BackColor = Color.Salmon
                    'this prompts the user that a mistake was made
                    'MsgBox(i & " is not a valid date character")
                    Exit Sub
                End If
            Next

            Me.ActiveControl.BackColor = Color.White

            'this condition is only true if the date entered is correct and properly formated
            If datecheck = False Then
                Try
                    Chartpopulation(curdate)
                Catch ex As Exception
                End Try
            End If
        Else

        End If

    End Sub

    'this variable is used to determine if the user has already edited the chart interval value
    Public useredited As Boolean = False

    'this subroutine will run when the interval textbox is edited
    Private Sub txtgraphInterval_TextChanged(sender As Object, e As EventArgs) Handles txtgraphInterval.TextChanged

        'once fired the useredited boolean state is set to true
        useredited = True
        graph_settings()

    End Sub

    'this subroutine fires when the bar number is changed
    Private Sub txtBarcount_TextChanged(sender As Object, e As EventArgs) Handles txtBarcount.TextChanged

        'this will proceed to change the maximum number of ranked values the chart diplays
        barcountmodifier()

    End Sub

    Dim currentstate As Boolean = True
    Private Sub btnMaxNorm_Click(sender As Object, e As EventArgs) Handles btnMaxNorm.Click

        If currentstate = True Then
            Me.WindowState = FormWindowState.Maximized
            currentstate = False
        Else
            Me.WindowState = FormWindowState.Normal
            Me.CenterToScreen()
            currentstate = True
        End If

    End Sub


    'this section of code is used to capture mouse coordinates in order to allow the user to drag the form around from the base control
    'in this case this is the tab control

    'this variable is a boolean that is used to capture whether the form is being dragged
    Private IsFormBeingDragged As Boolean = False

    'this variables capture the location the mouse has began at
    Private MouseDownX As Integer
    Private MouseDownY As Integer

    'this sub is used to capture the intial location of the mouse down
    Private Sub Main_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseDown

        'this ensure the mouse click is a left click
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    'this sub is used to reset the boolean as the mouse is no longer dragging the form
    Private Sub Main_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseUp

        'this resets the forms currently dragging boolean
        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    'this subroutine is used to capture the mouse move from point xy to point xy
    Private Sub Main_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseMove

        'this checks the boolean to determine if the form is actually moving
        If IsFormBeingDragged Then
            'this variable is used to act the second point
            Dim temp As Point = New Point()

            'this defines two x and y points as the current location + what ever the move on either axis has been
            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)

            'this code will maximize the window then it is dragged to the top of the screen
            If temp.Y = 0 Then
                currentstate = False
                Me.WindowState = FormWindowState.Maximized
            End If

            'this then moves the form to match that new location giving the illusion the user has moved the form rather then it following the mouse
            Me.Location = temp
            'this resets the temp variable
            temp = Nothing
        End If
    End Sub

    'this event is ran when the print button is pressed
    Private Sub btnPrintChart_Click(sender As Object, e As EventArgs) Handles btnPrintChart.Click

        'this subroutine acts to print out the chart as an image
        print_chart_data()

    End Sub

   
    Private Sub btnPrint_label_Click(sender As Object, e As EventArgs) Handles btnPrint_label.Click

        Label_Output.printlabel(Label_generator.grid, Label_generator.ammountcounter)

    End Sub
End Class