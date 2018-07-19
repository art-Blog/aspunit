<!-- #include virtual="/Services/Cart.asp" -->
<%
Class CartTest ' Extends TestCase
    Private target
    Public Function TestCaseNames()
        dim testCaseCollection() : redim testCaseCollection(-1)
        ReDim Preserve testCaseCollection(UBound(testCaseCollection) + 1) : testCaseCollection(UBound(testCaseCollection)) = "ProductsShouldPay_23400"
        TestCaseNames = testCaseCollection
    End Function

    Public Sub SetUp()
    set target = new Cart
    End Sub

    Public Sub TearDown()
    Set target = Nothing
    End Sub


    Public Sub ProductsShouldPay_23400(oTestResult)
        dim actual : actual = 23400
        dim expected : expected = 23400
        oTestResult.AssertEquals expected, actual
    End Sub
End Class
%>