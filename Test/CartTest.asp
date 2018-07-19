<!-- #include virtual="/Services/Cart.asp" -->
<!-- #include virtual="/Model/Product.asp" -->
<%
Class CartTest ' Extends TestCase
    Private target
    Public Function TestCaseNames()
        dim testCaseCollection() : redim testCaseCollection(-1)
        ReDim Preserve testCaseCollection(UBound(testCaseCollection) + 1) : testCaseCollection(UBound(testCaseCollection)) = "NormalMember_ProductsShouldPay_23400"
        ReDim Preserve testCaseCollection(UBound(testCaseCollection) + 1) : testCaseCollection(UBound(testCaseCollection)) = "VIP_ProductsShouldPay_18720"
        TestCaseNames = testCaseCollection
    End Function

    Public Sub SetUp()
    set target = new Cart
    End Sub

    Public Sub TearDown()
    Set target = Nothing
    End Sub

    Public Sub NormalMember_ProductsShouldPay_23400(oTestResult)
        dim userLevel : userLevel = ""
        dim products : products = GetWantedProducts()
        dim actual : actual = 23400
        dim expected : expected = target.ShouldPay(products, userLevel)
        oTestResult.AssertEquals expected, actual
    End Sub

    Public Sub VIP_ProductsShouldPay_18720(oTestResult)
        dim userLevel : userLevel = "VIP"
        dim products : products = GetWantedProducts()
        dim actual : actual = 18720
        dim expected : expected = target.ShouldPay(products, userLevel)
        oTestResult.AssertEquals expected, actual
    End Sub

    private function GetWantedProducts()
        dim result() : redim result(-1)
        ReDim Preserve result(UBound(result) + 1) : set result(UBound(result)) = (new Product)("任天堂 Nintendo Switch 藍紅手把組", 9780, 1)
        ReDim Preserve result(UBound(result) + 1) : set result(UBound(result)) = (new Product)("OSIM按摩椅", 12800, 1)
        ReDim Preserve result(UBound(result) + 1) : set result(UBound(result)) = (new Product)("柔韌潔淨抽取衛生紙100抽(8包x8串/箱)", 820, 1)
        GetWantedProducts = result
    end function
End Class
%>