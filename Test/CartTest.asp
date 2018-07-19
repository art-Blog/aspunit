<!-- #include virtual="/Services/Cart.asp" -->
<%
Class CartTest ' Extends TestCase
    Private target
    Public Function TestCaseNames()

    End Function

    Public Sub SetUp()
    set target = new Cart
    End Sub

    Public Sub TearDown()
    Set target = Nothing
    End Sub


    Public Sub AddNumberTest(oTestResult)
    'oTestResult.AssertEquals expected, actual
    End Sub
End Class
%>