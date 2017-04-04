Public Class Label_Output

    Private Sub Button1_Click(sender As Object, e As EventArgs)
    End Sub

    Sub printlabel(grid, ammountcounter)

        ammountcounter = Main.dgvLabel_gen.Rows.Count

        For i As Integer = 0 To ammountcounter - 1
            For j As Integer = 0 To 5

                Me.lblStockName.Text = grid(i, 0)

                Me.lblStockID.Text = Stock_Input.StockID

                Me.lblStockCategory.Text = grid(i, 2)

                Me.lblStockColour.Text = grid(i, 3)

                Me.lblStockSupplier.Text = grid(i, 4)

                Me.lblStockSize.Text = grid(i, 5)


            Next

            Label_Output.pn_label.SaveImage(print_date_best, System.Drawing.Imaging.ImageFormat.Png)
            Console.WriteLine(Stock_Input.StockID & " | " & grid(i, 0) & " | " & grid(i, 2) & " | " & grid(i, 3) & " | " & grid(i, 4) & " | " & grid(i, 5))
        Next

    End Sub

End Class