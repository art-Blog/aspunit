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

        if (ucase(userLevel)="VIP") then 
            result = result * 0.8
        end if

        ShouldPay = result
    end function

End Class
%>