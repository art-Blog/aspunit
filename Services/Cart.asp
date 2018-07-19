<%
Class Cart 
    Private Sub Class_Initialize()
    End Sub 
    Private Sub Class_Terminate()
    End Sub

    Public Default function Init()
        Set Init = Me
    End function

    public function ShouldPay(products, userLevel)
        dim result : result = 0
        dim prod : for each prod in products
            result = result + ( prod.Price * prod.Qty )
        next
        ShouldPay = result
    end function

End Class
%>