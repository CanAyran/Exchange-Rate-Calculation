Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Add("Euro") ' You must write the name of the currency you are going to add        
        ComboBox1.Text = "The currency you want to convert."
        WebBrowser1.Navigate("") ' The site where you will get exchange rate data
        WebBrowser1.ScriptErrorsSuppressed = True

    End Sub

    Function Daily(Choice As String)

        Dim Data As HtmlElementCollection = WebBrowser1.Document.All
        For Each element As HtmlElement In Data
            If element.GetAttribute("classname").Contains("") Then 'Contains is a name of the class in which the data is located
                ListBox1.Items.Add(Replace(element.InnerText, ".", ",") & Environment.NewLine)
            End If
        Next

        Dim A(ListBox1.Items.Count - 1) As Object
        ListBox1.Items.CopyTo(A, 0) ' listbox to array / Use new case when you added to new currency unit.
        Select Case Choice
            Case "Euro"
                Return A(0) 'index of exchange's rate
        End Select

    End Function

    Public Function Calc(Choice As Double)

        Calc = Choice * TextBox1.Text
        If Calc < 0 Then
            Calc = 0 - Calc
            MsgBox("A negative value has been entered.")
        Else
            Return Calc
        End If

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        result_txt.Text = Calc(Daily(ComboBox1.Text))
    End Sub
End Class
