<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.FlatLabel1 = New Azzurra_System_UI.FlatLabel()
        Me.lblUsername = New Azzurra_System_UI.FlatLabel()
        Me.lblPassword = New Azzurra_System_UI.FlatLabel()
        Me.btnCloseLogin = New Azzurra_System_UI.FlatClose()
        Me.btnMinLogin = New Azzurra_System_UI.FlatMini()
        Me.txtpassword = New System.Windows.Forms.TextBox()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'FlatLabel1
        '
        Me.FlatLabel1.AutoSize = True
        Me.FlatLabel1.BackColor = System.Drawing.Color.Transparent
        Me.FlatLabel1.Font = New System.Drawing.Font("Segoe UI", 14.0!)
        Me.FlatLabel1.ForeColor = System.Drawing.Color.White
        Me.FlatLabel1.Location = New System.Drawing.Point(38, 18)
        Me.FlatLabel1.Name = "FlatLabel1"
        Me.FlatLabel1.Size = New System.Drawing.Size(91, 25)
        Me.FlatLabel1.TabIndex = 12
        Me.FlatLabel1.Text = "Welcome"
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.BackColor = System.Drawing.Color.Transparent
        Me.lblUsername.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblUsername.ForeColor = System.Drawing.Color.White
        Me.lblUsername.Location = New System.Drawing.Point(40, 60)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(60, 15)
        Me.lblUsername.TabIndex = 10
        Me.lblUsername.Text = "Username"
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.BackColor = System.Drawing.Color.Transparent
        Me.lblPassword.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblPassword.ForeColor = System.Drawing.Color.White
        Me.lblPassword.Location = New System.Drawing.Point(40, 113)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(57, 15)
        Me.lblPassword.TabIndex = 11
        Me.lblPassword.Text = "Password"
        '
        'btnCloseLogin
        '
        Me.btnCloseLogin.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCloseLogin.BackColor = System.Drawing.Color.White
        Me.btnCloseLogin.BaseColor = System.Drawing.Color.FromArgb(CType(CType(168, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(35, Byte), Integer))
        Me.btnCloseLogin.Font = New System.Drawing.Font("Marlett", 10.0!)
        Me.btnCloseLogin.Location = New System.Drawing.Point(225, 12)
        Me.btnCloseLogin.Name = "btnCloseLogin"
        Me.btnCloseLogin.Size = New System.Drawing.Size(18, 18)
        Me.btnCloseLogin.TabIndex = 14
        Me.btnCloseLogin.Text = "FlatClose1"
        Me.btnCloseLogin.TextColor = System.Drawing.Color.FromArgb(CType(CType(243, Byte), Integer), CType(CType(243, Byte), Integer), CType(CType(243, Byte), Integer))
        '
        'btnMinLogin
        '
        Me.btnMinLogin.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMinLogin.BackColor = System.Drawing.Color.White
        Me.btnMinLogin.BaseColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btnMinLogin.Font = New System.Drawing.Font("Marlett", 12.0!)
        Me.btnMinLogin.Location = New System.Drawing.Point(201, 12)
        Me.btnMinLogin.Name = "btnMinLogin"
        Me.btnMinLogin.Size = New System.Drawing.Size(18, 18)
        Me.btnMinLogin.TabIndex = 13
        Me.btnMinLogin.Text = "FlatMini1"
        Me.btnMinLogin.TextColor = System.Drawing.Color.FromArgb(CType(CType(243, Byte), Integer), CType(CType(243, Byte), Integer), CType(CType(243, Byte), Integer))
        '
        'txtpassword
        '
        Me.txtpassword.Location = New System.Drawing.Point(43, 134)
        Me.txtpassword.Name = "txtpassword"
        Me.txtpassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(8226)
        Me.txtpassword.Size = New System.Drawing.Size(155, 20)
        Me.txtpassword.TabIndex = 1
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(43, 82)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(155, 20)
        Me.txtUsername.TabIndex = 0
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(39, Byte), Integer), CType(CType(39, Byte), Integer), CType(CType(39, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(255, 191)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.txtpassword)
        Me.Controls.Add(Me.btnMinLogin)
        Me.Controls.Add(Me.btnCloseLogin)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.FlatLabel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Login"
        Me.Text = "Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FlatLabel1 As Azzurra_System_UI.FlatLabel
    Friend WithEvents lblUsername As Azzurra_System_UI.FlatLabel
    Friend WithEvents lblPassword As Azzurra_System_UI.FlatLabel
    Friend WithEvents btnCloseLogin As Azzurra_System_UI.FlatClose
    Friend WithEvents btnMinLogin As Azzurra_System_UI.FlatMini
    Friend WithEvents txtpassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
End Class
