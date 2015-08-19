Option Explicit
Option Private Module

' 1 team, 16 bells, 1 change
Public TestArray(1, 16, 1) As Strike
Public NumBells(1) As Integer

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
End Sub

'@TestCleanup
Public Sub TestCleanup()
    ' No test pull-down required
End Sub

'@TestMethod
Public Sub Test_outputUnchangedWhenStrikesAlreadyInOrder_withOffset()
    On Error GoTo TestFail
    
    'Set strike data for four bells
    TestArray(1, 1, 1).stroke = "H"
    TestArray(1, 2, 1).stroke = "H"
    TestArray(1, 3, 1).stroke = "H"
    TestArray(1, 4, 1).stroke = "H"
    TestArray(1, 1, 1).bell = "1"
    TestArray(1, 2, 1).bell = "2"
    TestArray(1, 3, 1).bell = "3"
    TestArray(1, 4, 1).bell = "4"
    TestArray(1, 1, 1).time = "100"
    TestArray(1, 2, 1).time = "200"
    TestArray(1, 3, 1).time = "300"
    TestArray(1, 4, 1).time = "400"
        
    NumBells(1) = 4
    
    ' This function creates "TimeOrder"
    Call RodModel2.order(TestArray, NumBells, 1, 1)
    
    ' Assert strike times are added in order, with indices 1 to 4
    Assert.AreEqual TimeOrder(1), 100&, "First strike"
    Assert.AreEqual TimeOrder(2), 200&, "Second strike"
    Assert.AreEqual TimeOrder(3), 300&, "Third strike"
    Assert.AreEqual TimeOrder(4), 400&, "Fourth strike"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Test_outputUnchangedWhenStrikesAlreadyInOrder_withoutOffset()
    On Error GoTo TestFail
    
    'Set strike data for four bells
    TestArray(1, 1, 1).stroke = "H"
    TestArray(1, 2, 1).stroke = "H"
    TestArray(1, 3, 1).stroke = "H"
    TestArray(1, 4, 1).stroke = "H"
    TestArray(1, 1, 1).bell = "1"
    TestArray(1, 2, 1).bell = "2"
    TestArray(1, 3, 1).bell = "3"
    TestArray(1, 4, 1).bell = "4"
    TestArray(1, 1, 1).time = "0"
    TestArray(1, 2, 1).time = "100"
    TestArray(1, 3, 1).time = "200"
    TestArray(1, 4, 1).time = "300"
        
    ' This function creates "TimeOrder"
    Call RodModel2.order(TestArray, NumBells, 1, 1)
    
    ' Assert strike times are added in order, with indices 1 to 4
    Assert.AreEqual TimeOrder(1), 0&, "First strike"
    Assert.AreEqual TimeOrder(2), 100&, "Second strike"
    Assert.AreEqual TimeOrder(3), 200&, "Third strike"
    Assert.AreEqual TimeOrder(4), 300&, "Fourth strike"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Test_outputChangedWhenStrikesNotAlreadyInOrder_withOffset()
    On Error GoTo TestFail
    
    'Set strike data for four bells
    TestArray(1, 1, 1).stroke = "H"
    TestArray(1, 2, 1).stroke = "H"
    TestArray(1, 3, 1).stroke = "H"
    TestArray(1, 4, 1).stroke = "H"
    TestArray(1, 1, 1).bell = "1"
    TestArray(1, 2, 1).bell = "2"
    TestArray(1, 3, 1).bell = "3"
    TestArray(1, 4, 1).bell = "4"
    TestArray(1, 1, 1).time = "200"
    TestArray(1, 2, 1).time = "100"
    TestArray(1, 3, 1).time = "400"
    TestArray(1, 4, 1).time = "300"
        
    ' This function creates "TimeOrder"
    Call RodModel2.order(TestArray, NumBells, 1, 1)
    
    ' Assert strike times are added in order, with indices 1 to 4
    Assert.AreEqual TimeOrder(1), 100&, "First strike"
    Assert.AreEqual TimeOrder(2), 200&, "Second strike"
    Assert.AreEqual TimeOrder(3), 300&, "Third strike"
    Assert.AreEqual TimeOrder(4), 400&, "Fourth strike"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub Test_outputChangedWhenStrikesNotAlreadyInOrder_withoutOffset()
    On Error GoTo TestFail
    
    'Set strike data for four bells
    TestArray(1, 1, 1).stroke = "H"
    TestArray(1, 2, 1).stroke = "H"
    TestArray(1, 3, 1).stroke = "H"
    TestArray(1, 4, 1).stroke = "H"
    TestArray(1, 1, 1).bell = "1"
    TestArray(1, 2, 1).bell = "2"
    TestArray(1, 3, 1).bell = "3"
    TestArray(1, 4, 1).bell = "4"
    TestArray(1, 1, 1).time = "100"
    TestArray(1, 2, 1).time = "0"
    TestArray(1, 3, 1).time = "300"
    TestArray(1, 4, 1).time = "200"
        
    ' This function creates "TimeOrder"
    Call RodModel2.order(TestArray, NumBells, 1, 1)
    
    ' Assert strike times are added in order, with indices 1 to 4
    Assert.AreEqual TimeOrder(1), 0&, "First strike"
    Assert.AreEqual TimeOrder(2), 100&, "Second strike"
    Assert.AreEqual TimeOrder(3), 200&, "Third strike"
    Assert.AreEqual TimeOrder(4), 300&, "Fourth strike"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


