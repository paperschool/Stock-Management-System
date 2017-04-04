Imports System.Drawing

Public Class Label_Output

    Private Sub Button1_Click(sender As Object, e As EventArgs)
    End Sub

    Sub printlabel(grid, ammountcounter)



        Dim directory As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Labels\" + Stock_Input.StockID + "\"

        If (Not System.IO.Directory.Exists(directory)) Then
            'if it doesnt, it creates the new directory
            System.IO.Directory.CreateDirectory(directory)
        End If

        ammountcounter = Main.dgvLabel_gen.Rows.Count

        For i As Integer = 0 To ammountcounter - 1

            Me.Hide()

            Me.lblStockID.Text = Stock_Input.StockID + Environment.NewLine + grid(i, 0) + Environment.NewLine + grid(i, 2) + Environment.NewLine + grid(i, 3) + Environment.NewLine + grid(i, 4) + Environment.NewLine + "Size: " + grid(i, 5)


            Dim Label_name As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Labels\" + Stock_Input.StockID + "\" + i.ToString + "_" + Stock_Input.StockID + ".png"

            Console.WriteLine(Stock_Input.StockID & " | " & grid(i, 0) & " | " & grid(i, 2) & " | " & grid(i, 3) & " | " & grid(i, 4) & " | " & grid(i, 5))

            Me.Show()

            LabelPrintDoc.PrinterSettings.PrintToFile = True
            LabelPrintDoc.PrinterSettings.PrintFileName = Label_name
            LabelPrintDoc.Print()

            'Me.lblStockID.SaveImage(Label_name, System.Drawing.Imaging.ImageFormat.Png)

        Next

    End Sub


End Class