<!-- #include virtual="/Services/Cart.asp" -->
<!-- #include virtual="/Model/Product.asp" -->
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
        dim productA : set productA = (new Product)("任天堂 Nintendo Switch 藍紅手把組", 9780, 1)
        dim productB : set productB = (new Product)("OSIM按摩椅", 12800, 1)
        dim productC : set productC = (new Product)("柔韌潔淨抽取衛生紙100抽(8包x8串/箱)", 820, 1)
        dim products : products = Array(productA, productB, productC)

        dim actual : actual = 23400
        dim expected : expected = target.ShouldPay(products)
        oTestResult.AssertEquals expected, actual
    End Sub
End Class
%>