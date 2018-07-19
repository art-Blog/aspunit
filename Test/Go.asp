<%
Option Explicit
%>
<!-- #include virtual="/include/ASPUnitRunner.asp"-->
<!-- #include file="SampleTest.asp"-->
<%
Dim oRunner
Set oRunner = New UnitRunner
oRunner.AddTestContainer New SampleTest
oRunner.Display()

%>
