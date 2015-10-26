Attribute VB_Name = "TestFileTools"
Option Explicit
Option Private Module

'@TestModule
Private Assert As New Rubberduck.AssertClass

'@ModuleInitialize
Public Sub ModuleInitialize()
    ' No initialisation necessary
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    ' No cleanup
End Sub

'@TestInitialize
Public Sub TestInitialize()
    ' No test init required
End Sub

'@TestCleanup
Public Sub TestCleanup()
    ' No test pull-down required
End Sub

'@TestMethod
Public Sub TestGetFileTitle()
    On Error GoTo TestFail
    
    Assert.AreEqual GetFileTitle("qaz.txt"), "qaz", "File path"
    Assert.AreEqual GetFileTitle("C:\temp\asda-2k.txt"), "asda-2k", "Full file path"
    'TODO: Is this really right??
    Assert.AreEqual GetFileTitle("C:\temp\jqw_2h.wav.txt"), "jqw_2h.wav", "Double extension"
    Assert.AreEqual GetFileTitle("\\server\temp\ABcD.txt"), "ABcD", "Network address"
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestSecondsToHoursAndMinutes()
    On Error GoTo TestFail
    
    Assert.AreEqual SecondsToHoursAndMinutes(29), "0m", "Rounding-down of minutes"
    Assert.AreEqual SecondsToHoursAndMinutes(31), "1m", "Rounding-up of minutes"
    Assert.AreEqual SecondsToHoursAndMinutes(59 * 60 + 29), "59m", "Rounding-down of hours"
    Assert.AreEqual SecondsToHoursAndMinutes(60.5 * 60), "1h 0m", "Rounding-down of hours and minutes"
    Assert.AreEqual SecondsToHoursAndMinutes(60.5 * 60, "hr"), "1hr 0m", "Hour marker"
    Assert.AreEqual SecondsToHoursAndMinutes(60.5 * 60, , "min"), "1h 0min", "Minute marker"
    Assert.AreEqual SecondsToHoursAndMinutes(60.5 * 60, " hour", " min"), "1 hour 0 min", "Both markers"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

