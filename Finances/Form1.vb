Public Class Form1
    Private Sub TransactionsBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles TransactionsBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.TransactionsBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.TailoringBusinessDataSet)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'TailoringBusinessDataSet.Transactions' table. You can move, or remove it, as needed.
        Me.TransactionsTableAdapter.Fill(Me.TailoringBusinessDataSet.Transactions)

    End Sub

    Private Sub BtnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click
        Dim filter As String = Nothing
        If (txtBoxDate.Text <> Nothing) Then
            If (filter <> Nothing) Then
                filter += " And "
            End If
            filter += "[Date] = '" & txtBoxDate.Text & "'"
        End If

        Me.TransactionsBindingSource.Filter = filter
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtBoxDate.Text = Nothing
        Me.TransactionsBindingSource.RemoveFilter()
    End Sub

    Private Sub BtnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click


        Dim strSql As String = "SELECT * FROM Transactions"

        Dim strPath As String = "Provider=Microsoft.ACE.OLEDB.12.0 ;" & "Data Source=C:\Users\ryanb\source\repos\Finances\Finances\TailoringBusiness.accdb"
        Dim odaItems As New OleDb.OleDbDataAdapter(strSql, strPath)
        Dim datValue As New DataTable
        Dim intCount As Integer

        Dim decTotalIncomes As Decimal = 0D
        Dim decTotalExpenses As Decimal = 0D
        Dim decTotalCommissions As Decimal = 0D

        odaItems.Fill(datValue)
        odaItems.Dispose()

        For Each row As DataRow In TailoringBusinessDataSet.Transactions.Rows
            If (txtBoxDate.Text <> Nothing) Then
                If (String.Compare(row("Date"), txtBoxDate.Text) = 0) Then
                    If (Convert.ToDecimal(row("Total")) <= 0) Then
                        decTotalExpenses += Convert.ToDecimal(row("Total"))
                    ElseIf (Convert.ToDecimal(row("Total")) > 0) Then
                        decTotalIncomes += Convert.ToDecimal(row("Total"))
                    End If
                    decTotalCommissions += Convert.ToDecimal(row("Commission"))
                End If
            Else
                If (Convert.ToDecimal(row("Total")) <= 0) Then
                    decTotalExpenses += Convert.ToDecimal(row("Total"))
                ElseIf (Convert.ToDecimal(row("Total")) > 0) Then
                    decTotalIncomes += Convert.ToDecimal(row("Total"))
                End If
                decTotalCommissions += Convert.ToDecimal(row("Commission"))
            End If
        Next
        Dim FILE_NAME As String = "FinancesReport-" & Format(Now, "dddd, d MMM yyyy").ToString() & ".txt"
        MsgBox(FILE_NAME)
        If System.IO.File.Exists(FILE_NAME) = False Then
            System.IO.File.Create(FILE_NAME).Dispose()
        End If
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
        objWriter.WriteLine("Finances Report " & Format(Now, "dddd, d MMM yyyy").ToString())
        objWriter.WriteLine("Date..........:  " & txtBoxDate.Text)
        objWriter.WriteLine("Incomes.......: $" & decTotalIncomes.ToString())
        objWriter.WriteLine("Expenses......: $" & decTotalExpenses.ToString())
        objWriter.WriteLine("Commissions...: $" & decTotalCommissions.ToString())
        objWriter.WriteLine("-----------------------------------------------------")
        objWriter.WriteLine("Profits.......: $" & ((decTotalIncomes + decTotalExpenses) - decTotalCommissions).ToString())
        objWriter.Close()
    End Sub
End Class
