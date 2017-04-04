'this imports the database manipulationa and connectivity libraries
Imports System.Data.OleDb

'this class contains the login form 
Public Class Login

    'this declares the username and password variables as strings
    Public Username As String
    Public Password As String

    'this is a text keyup event that runs when the keypress is disengaged within the username textbox
    Private Sub txtUsername_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUsername.KeyUp

        'this defines the username variable as the content of the username textbox
        Username = txtUsername.Text
        'this runs the checkuser subroutine
        CheckUser()
        'Console.WriteLine("Username: " & Username)

    End Sub

    'this is a text keyup event that runs when the keypress is disengaged within the password textbox
    Private Sub txtPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpassword.KeyUp

        'this defines the password variable as the content of the password textbox
        Password = txtpassword.Text

        'this runs the check user subroutine
        CheckUser()
        'Console.WriteLine("Password: " & Password)

    End Sub

    'this subroutine checks the content of the username and password against the databases user table
    Sub CheckUser()

        'Console.WriteLine("Username: " & Username)
        'Console.WriteLine("Password: " & Password)

        'this automatically sets the background colour of the username and password textbox to red
        txtusername.BackColor = Color.Salmon
        txtpassword.BackColor = Color.Salmon


        Username = txtusername.Text
        Password = txtpassword.Text

        'this condition checks if there is any data in the textbox
        If Username = "" Then
            'this resets the colour of the textbox to white
            txtUsername.BackColor = Color.White
        End If

        'this condition checks if there is any data in the textbox
        If Password = "" Then
            'this resets the colour of the textbox to white
            txtpassword.BackColor = Color.White
        End If


        'this sql selection statement selects everything from the usertable
        Dim User_Info_Check As String = "SELECT * FROM tblUsers"
        'this defines a data adapter for the login form
        Dim loginda As OleDbDataAdapter = New OleDbDataAdapter(User_Info_Check, conn)
        'this defines the dataset
        Dim loginds As DataSet = New DataSet
        'this fills data from the dataadapter into the data set users
        loginda.Fill(loginds, "Users")
        'this defines the datatable
        Dim logindt As DataTable = loginds.Tables("Users")

        'this loop will loop through each record in the datatable
        For Each row As DataRow In logindt.Rows
            'this condition checks each row first item in the iterated loop with the username variable containing username entered data
            'from the username textbox, if a match is found
            If row.Item(1) = Username Then
                'the username textbox colour is changed to green
                txtUsername.BackColor = Color.LawnGreen
                'this condition checks each second item iteration of the loop with the content of the password variable 
                If row.Item(2) = Password Then

                    'this changes the colour of the textbox to green 
                    txtpassword.BackColor = Color.LawnGreen
                    'this resets the location of the form to the top left hand corner
                    Me.Location = New Point(0, 0)

                    'this message box displays when a successful login attempt is made
                    MsgBox("Login Successful")

                    'this displays the main form after a successful login
                    Main.Show()

                    'this resets the usenamd and password textboxes
                    txtUsername.Text = ""
                    txtpassword.Text = ""

                    'this changes the textbox colours from current colour to white
                    txtUsername.BackColor = Color.White
                    txtpassword.BackColor = Color.White

                    'this disables the use of the these textboxes
                    txtUsername.Enabled = False
                    txtpassword.Enabled = False

                End If
            End If
        Next

    End Sub

    'this defines the centre position of the screen where the form is placed at the start of the program
    Sub center_position()

        'this defines the middle of the horizontal line in the screen
        Dim x As Integer = (My.Computer.Screen.Bounds.Width / 2) - (Me.Size.Width / 2)
        'this defines the middle of the vertical line in the screen
        Dim y As Integer = (My.Computer.Screen.Bounds.Height / 2) - (Me.Size.Height / 2)
        'this sets the current form location to the middle of the screen
        Me.Location = New Point(x, y)

    End Sub

    'this is the form load event
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'this resets the content of the username and password textbox 
        txtpassword.Text = ""
        txtUsername.Text = ""
        'this places the form at the centre using the centre form subroutine
        center_position()
        'this runs the connection module
        connect()

    End Sub

    Private Sub FlatClose1_Click(sender As Object, e As EventArgs) Handles btnCloseLogin.Click

        Main.Close()
        'this closes the form when the close button is pressed
        Me.Close()

    End Sub

    Private Sub FlatMini1_Click(sender As Object, e As EventArgs) Handles btnMinLogin.Click

        'this minimizes the form when the minimize button is pressed
        Me.MinimizeBox() = True

    End Sub

End Class