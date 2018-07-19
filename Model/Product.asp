<%
Class Product 
    public Name     '名稱
    public Price    '單價
    public Qty      '數量

    Private Sub Class_Initialize()
    End Sub 
    Private Sub Class_Terminate()
    End Sub

    Public Default function Init(productName, productPrice, productQty)
        Name = productName
        Price = productPrice
        Qty = productQty
        Set Init = Me
    End function
End Class
%>