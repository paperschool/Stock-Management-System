'This imports the database connectivity module
Imports System.Data.OleDb

Module Presets

    'these variables are used when data is being sourced from the database
    Dim presetda As OleDbDataAdapter
    Dim presetdt As DataTable
    Dim presetds As DataSet = New DataSet

    'this boolean is used when detecing duplications
    Dim duplicatebool As Boolean = False

    'Auto Complete'
    Dim colorlst As New List(Of String)
    Dim stylelst As New List(Of String)
    Dim supplierlst As New List(Of String)
    Dim ColourSource As New AutoCompleteStringCollection()
    Dim StyleSource As New AutoCompleteStringCollection()
    Dim SupplierSource As New AutoCompleteStringCollection()

    'unued auto complete subroutine
    Sub autocomplete()

        ColourSource.AddRange(colorlst.ToArray)
        Main.cbColour2.AutoCompleteCustomSource = ColourSource
        Main.cbColour2.AutoCompleteMode = AutoCompleteMode.None
        Main.cbColour2.AutoCompleteSource = AutoCompleteSource.CustomSource

        StyleSource.AddRange(stylelst.ToArray)
        Main.cbStyle2.AutoCompleteCustomSource = StyleSource
        Main.cbStyle2.AutoCompleteMode = AutoCompleteMode.None
        Main.cbStyle2.AutoCompleteSource = AutoCompleteSource.CustomSource

        SupplierSource.AddRange(supplierlst.ToArray)
        Main.cbSupplier2.AutoCompleteCustomSource = SupplierSource
        Main.cbSupplier2.AutoCompleteMode = AutoCompleteMode.None
        Main.cbSupplier2.AutoCompleteSource = AutoCompleteSource.CustomSource

    End Sub

    'this subroutine is used to insert presets into the database
    Sub main_information_updater(e)

        'this condition is used to catch a "enter" key entry
        If e.keycode = Keys.Enter Then
            'this send keys function is used to virtually insert the tab key to act as a next line short cut
            SendKeys.Send("{TAB}")
        End If

        'this condition is used to catch a combination of keys "enter" and "ctrl"
        If e.KeyCode = Keys.Enter AndAlso e.control = True Then
            'this condition checks if any controls are being used
            If Main.ActiveControl Is Nothing Then
                'this ends the subroutine if so
                Exit Sub
            End If
            If Main.ActiveControl.BackColor = Color.Salmon Then
                MsgBox("Preset Entered: " & Main.ActiveControl.Text & " Is not valid.")
                Exit Sub
            End If
            'this condition checks the name of the active control to ensure only controls that are capable of entering presets are used
            If Main.ActiveControl.Name.ToString = "cbStyle" Or Main.ActiveControl.Name.ToString = "cbStyle2" Or Main.ActiveControl.Name.ToString = "cbColour" Or _
                Main.ActiveControl.Name.ToString = "cbColour2" Or Main.ActiveControl.Name.ToString = "cbSupplier" Or Main.ActiveControl.Name.ToString = "cbSupplier2" Then

                'this line changes the back colour from its current colour to green
                Main.ActiveControl.BackColor = Color.YellowGreen

                'this runs the save preset subroutine
                save_presets()

                'this empties the textbox after entering the data
                Main.ActiveControl.Text = ""
            End If
        End If

        'this condition is used to catch a combination of keys "delete" and "ctrl"
        If e.KeyCode = Keys.Delete AndAlso e.control = True Then

            'this condition checks if any controls are being used
            If Main.ActiveControl Is Nothing Then
                'this ends the subroutine if so
                Exit Sub
            End If

            'this condition checks the name of the active control to ensure only controls that are capable of deleting presets are used
            If Main.ActiveControl.Name.ToString = "cbStyle" Or Main.ActiveControl.Name.ToString = "cbStyle2" Or Main.ActiveControl.Name.ToString = "cbColour" Or _
                Main.ActiveControl.Name.ToString = "cbColour2" Or Main.ActiveControl.Name.ToString = "cbSupplier" Or Main.ActiveControl.Name.ToString = "cbSupplier2" Then
                'this line changes the back colour from its current colour to red
                Main.ActiveControl.BackColor = Color.Salmon

                'this runs the delete preset subroutine
                Delete_presets()

                'this empties the textbox after entering the data
                Main.ActiveControl.Text = ""
            End If
        End If

    End Sub

    'this subroutine is used to fill the comboboxes with all the current presets
    Sub populate_Presets()

        'this sql string grabs all data from the preset table
        Dim CategorySearch As String = "SELECT * FROM tblPresets" 'WHERE PresetCategory = '" & "Category" & "' "

        'this block empties all the current combobox contents making way for new content
        Main.cbStyle.Items.Clear()
        Main.cbStyle2.Items.Clear()
        Main.cbColour.Items.Clear()
        Main.cbColour2.Items.Clear()
        Main.cbSupplier.Items.Clear()
        Main.cbSupplier2.Items.Clear()

        'this data adapter is used to transfer data from the database using the connection string and sql query
        presetda = New OleDbDataAdapter(CategorySearch, conn)
        'Data adapter told to fill the dataset
        presetda.Fill(presetds, "Presets")
        'Defining the datatable
        presetdt = presetds.Tables("Presets")

        'this loops round for as many rows as there are in the datatable minus 1
        For i As Integer = 0 To presetds.Tables("Presets").Rows.Count - 1
            'this condiiton checks if the preset category column of the current row = "colour"
            If presetds.Tables("Presets").Rows(i).Item(2).ToString = "Colour" Then
                'the add function adds the information in the preset name column of the current row to all the pertinent comboboxes
                Main.cbColour.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbColour2.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbstockcolouredit.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbColour2.AutoCompleteCustomSource.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                'this condiiton checks if the preset category column of the current row = "Style"
            ElseIf presetds.Tables("Presets").Rows(i).Item(2).ToString = "Style" Then
                'the add function adds the information in the preset name column of the current row to all the pertinent comboboxes
                Main.cbStyle.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbStyle2.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbStockstyleedit.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbStyle2.AutoCompleteCustomSource.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                'this condiiton checks if the preset category column of the current row = "supplier"
            ElseIf presetds.Tables("Presets").Rows(i).Item(2).ToString = "Supplier" Then
                'the add function adds the information in the preset name column of the current row to all the pertinent comboboxes
                Main.cbSupplier.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbSupplier2.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbstocksupplieredit.Items.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
                Main.cbSupplier2.AutoCompleteCustomSource.Add(presetds.Tables("Presets").Rows(i).Item(1).ToString)
            End If

        Next

    End Sub

    'this subroutine is used to remove any duplicates found in the database upon attempting to save a preset
    Sub Duplicate_check(presetcategory)

        'this sql query pulls all data from the preset table
        Dim duplicatecheck As String = "SELECT * FROM tblPresets"

        'this data adapter is used to transfer data from the database using the connection string and sql query
        presetda = New OleDbDataAdapter(duplicatecheck, conn)
        'Data adapter told to fill the dataset
        presetda.Fill(presetds, "duplicatecheck")
        'Defining the datatable
        presetdt = presetds.Tables("duplicatecheck")

        'this loop runs for as many iterations as there are rows in the in the datatable minus 1
        For i As Integer = 0 To presetds.Tables("duplicatecheck").Rows.Count - 1
            'this condition checks to see if the current category matches the current one in the current row and the value within the current active control matches the current preset name 
            If presetds.Tables("duplicatecheck").Rows(i).Item(1).ToString = Main.ActiveControl.Text And presetds.Tables("duplicatecheck").Rows(i).Item(2).ToString = presetcategory Then
                'this variable is used to define the optional message box
                Dim result As Integer = MessageBox.Show("Preset: " & Main.ActiveControl.Text & " Alredy Exists, Are you sure you want to add it again?", "Warning", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    'this sets a boolean to true
                    duplicatebool = True
                ElseIf result = DialogResult.Yes Then
                End If
            End If
        Next

    End Sub

    'this subroutine is used to save a preset to the database
    Sub save_presets()

        'this query string is used to grab the highest primary key in the preset table by subtringing the primary key grabbing only the number value, then converting them
        'to an integer then identifying the maximum number in the database and converting that to a string with a p before it
        Dim MAXID As String = "SELECT 'P' & cstr(MAX(CINT(MID(PresetID,2,5)))) FROM tblPresets WHERE PresetID LIKE '" & "P" & "%%%%%%' "

        'this data adapter is used to transfer data from the database using the connection string and sql query
        presetda = New OleDbDataAdapter(MAXID, conn)
        'Data adapter told to fill the dataset
        presetda.Fill(presetds, "PresetID")
        'Defining the datatable
        presetdt = presetds.Tables("PresetID")

        'this sets the max id value to the first row as this will be the only row
        MAXID = presetds.Tables("PresetID").Rows(0).Item(0)

        'this subtrings the current max id grabbing only the number value
        MAXID = MAXID.Substring(1, MAXID.Length - 1)
        'this condition checks to see if the last value of the preset = 9, this is to counter a logic error that seems to skip 0 in the sql statement
        If MAXID.Substring(MAXID.Length - 1, 1).ToString = 9 Then
            'this adds two to the current preset id
            MAXID = MAXID + 2
        Else
            'this adds one to the current preset id
            MAXID = MAXID + 1
        End If

        'this inserts "p" into the maxid
        MAXID = "P" & MAXID

        'this locates the preset category name
        Dim presetcategory As String = Main.ActiveControl.Name.ToString()
        'this locates the preset name based off of the content in the combobox
        Dim preset As String = Main.ActiveControl.Text

        'this strips off the identifier on the combobox control name eg "cb"
        presetcategory = presetcategory.Substring(2, presetcategory.Length - 2)

        'this condition checks to make sure that combobox is not the secondary combobox found on the stock creation page
        If presetcategory.Substring(presetcategory.Length - 1, 1) = "2" Then
            'this strips off the 2 at the end of the preset category control name
            presetcategory = presetcategory.Substring(0, presetcategory.Length - 1)
        End If

        'check for duplicate values before proceeding with the entry in the database
        Duplicate_check(presetcategory)

        'this boolean is refered to in the duplication check algorithm, if there is duplication, it resets the boolean and exits the sub
        If duplicatebool = True Then
            duplicatebool = False
            Exit Sub
        End If

        'this string variable defines the insert statement, which will insert the three points of data into the presets table
        Dim presetinsert As String = "INSERT INTO tblPresets (PresetID, Information, PresetCategory) VALUES (@PresetID, @Information, @PresetCategory)"

        Dim SQL_Command As New OleDbCommand
        With SQL_Command
            .CommandText = presetinsert
            'these parameters match the swl string variables with the data points in the code
            .Parameters.AddWithValue("@PresetID", MAXID)
            .Parameters.AddWithValue("@Information", preset)
            .Parameters.AddWithValue("@PresetCategory", presetcategory)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        'this clears the preset data set
        presetds.Clear()

        'this repopulates all the dropdowns
        populate_Presets()

        'succesful add prompt
        MsgBox("Preset: " & preset & " Successfully Added")

    End Sub

    'this subroutine is used to delete the presets if desired
    Sub Delete_presets()

        'this yes no message box is used ran to ensure the user wants to perform a deletion 
        Dim result As Integer = MessageBox.Show("You are about to Delete the Preset: " & Main.ActiveControl.Text & ", Are you sure you want to do this?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            'this exists the sub if the user selects no
            Exit Sub
        ElseIf result = DialogResult.Yes Then
        End If

        'this defines the preset category to be deleted
        Dim presetcategory As String = Main.ActiveControl.Name.ToString()

        'this defines the preset itself to be deleted
        Dim preset As String = Main.ActiveControl.Text

        'this substrings off the identifier on the preset control leaving only the preset name eg "supplier"
        presetcategory = presetcategory.Substring(2, presetcategory.Length - 2)

        'as there are two comboboxes the 2 at the end needs to be deleted, this condition substrings everythign but the last character
        'this will leave either a letter or the number 2
        If presetcategory.Substring(presetcategory.Length - 1, 1) = "2" Then
            'this strips off the 2 at the end of the control name
            presetcategory = presetcategory.Substring(0, presetcategory.Length - 1)
        End If

        'this defines the delete preset string, deleting the record where the preset name and preset category match those variables defined earlier
        Dim deletepreset As String = "DELETE FROM tblpresets WHERE Information = @Information AND PresetCategory = @PresetCategory"

        Dim SQL_Command As New OleDbCommand
        With SQL_Command
            .CommandText = deletepreset
            'these add parameters feed the two variables into the sql deletion string
            .Parameters.AddWithValue("@Information", preset)
            .Parameters.AddWithValue("@PresetCategory", presetcategory)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        'this clears the preset dataset
        presetds.Clear()

        'this repopulates the preset comboboxes
        populate_Presets()

    End Sub


End Module
