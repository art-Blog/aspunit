<!-- #include virtual="/Services/Sample.asp" -->
<%
Class SampleTest ' Extends TestCase

  Private target

  Public Function TestCaseNames()
    dim Something() : redim Something(-1)
    ReDim Preserve Something(UBound(Something) + 1) : Something(UBound(Something)) = "AddNumberTest"
    TestCaseNames = something
  End Function
  
  Public Sub SetUp()
    set target = new Calculator
  End Sub
  
  Public Sub TearDown()
    Set target = Nothing
  End Sub
  

  Public Sub AddNumberTest(oTestResult)
    'arrange'
    dim number1 : number1 = 50
    dim number2 : number2 = 100
    dim expected : expected = 150
    
    'act'
    dim actual : actual = target.Add(number1, number2)

    'assert'
    oTestResult.AssertEquals expected, actual
  End Sub


End Class
%>