Attribute VB_Name = "Module1"
Sub MATTEVATOREmigration()
Dim floors As Integer
floors = 0
floors = Range("D9").Value
Elevators = Range("D11").Value
Worksheets("MATTEVATOR EMIGRATION").Activate
Worksheets("MATTEVATOR EMIGRATION").Range("A2:B101").ClearContents
Dim FloorIterator As Integer
FloorIterator = 2
For FloorIterator = 2 To (floors + 1)
Range("A" & FloorIterator).Value = "Floor " & (FloorIterator - 1)
Range("B" & FloorIterator).Value = Application.InputBox(prompt:="Number of travelers from Floor " & (FloorIterator - 1), Type:=1)
Next FloorIterator
Call MATTEVATORResults
End Sub
Sub MATTEVATORLearn()
Worksheets("MATTEVATOR LEARN").Activate
End Sub
Sub MATTEVATORResults()
Worksheets("MATTEVATOR RESULTS").Range("A1:Z101").ClearContents
Worksheets("MATTEVATOR RESULTS").Range("A2:A101").Value = Worksheets("MATTEVATOR EMIGRATION").Range("A2:A101").Value
Dim elevatorcount As Integer
elevatorcount = Worksheets("MATTEVATOR HOME").Range("D11").Value
Dim ElevatorIterator As Integer
ElevatorIterator = 1
For ElevatorIterator = 1 To elevatorcount
Worksheets("MATTEVATOR RESULTS").Cells(1, ElevatorIterator + 1).Value = "Elevator " & (ElevatorIterator)
Next ElevatorIterator
Dim EmigratorIterator As Integer
Dim elevatorplacement As Integer
elevatorplacement = 2
EmigratorIterator = 2
floors = Worksheets("MATTEVATOR HOME").Range("D9").Value
For floormaxloop = 1 To 101
For EmigratorIterator = 2 To (floors + 1)
If elevatorcount > 0 Then
If Range("B" & EmigratorIterator).Value = Application.WorksheetFunction.Max(Range("B2:B101")) Then
Worksheets("MATTEVATOR RESULTS").Cells(EmigratorIterator, elevatorplacement).Value = "X"
elevatorplacement = elevatorplacement + 1
Range("B" & EmigratorIterator).Value = 0
elevatorcount = elevatorcount - 1
End If
End If
Next EmigratorIterator
Next floormaxloop
Worksheets("MATTEVATOR RESULTS").Activate
elevatorcount = Worksheets("MATTEVATOR HOME").Range("D11").Value
MsgBox "Your elevator configuration is most similar to the configuration of the building at the address of " & Application.WorksheetFunction.VLookup(elevatorcount, Worksheets("Elevators by location").Range("A:B"), 2, 0) & " which according to regulatory data has " & Application.WorksheetFunction.VLookup(elevatorcount, Worksheets("Elevators by location").Range("A:B"), 1, 0) & " elevators."
End Sub
Sub MATTEVATORHome()
Worksheets("MATTEVATOR HOME").Activate
End Sub
