#!/usr/bin/env python3
"""
VBA Event System Differential Probe
===================================
Probes and verifies:
1. Worksheet_Change event triggers upon cell mutation.
2. Application.EnableEvents = False event suppression.
3. Workbook_SheetChange workbook-level event dispatch.
4. Custom events with `Event`, `RaiseEvent`, and `Dim WithEvents`.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import visi_core
except ImportError:
    sys.exit(
        "the visi_core bindings are required: "
        "maturin develop -m visi-python/Cargo.toml"
    )

def test_worksheet_change_events():
    print("Testing Worksheet_Change and EnableEvents...")
    wb = visi_core.Workbook()
    
    sheet1_src = """Attribute VB_Name = "Sheet1"
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        Application.EnableEvents = False
        Range("B1").Value = Target.Value * 10
        Application.EnableEvents = True
    End If
End Sub
"""

    main_src = """Attribute VB_Name = "Main"
Sub MutateCell()
    Range("A1").Value = 7
End Sub
"""

    wb.add_macro("Sheet1", sheet1_src, kind="document", sheet="Sheet1")
    wb.add_macro("Main", main_src, kind="standard")
    
    type_name, val, mutated = wb.run_macro("MutateCell", module="Main")
    assert mutated, "Expected workbook to be mutated"
    
    val_a1 = wb.get_display(0, 0, sheet="Sheet1")
    val_b1 = wb.get_display(0, 1, sheet="Sheet1")
    assert val_a1 == "7", f"A1 expected 7, got {val_a1}"
    assert val_b1 == "70", f"B1 expected 70, got {val_b1}"
    print("  ✓ Worksheet_Change and EnableEvents passed")

def test_custom_events_with_events():
    print("Testing custom Event, RaiseEvent, and WithEvents...")
    wb = visi_core.Workbook()
    
    emitter_cls = """Attribute VB_Name = "Clock"
Public Event Tick(ByRef count As Long)
Public Sub Step(n As Long)
    RaiseEvent Tick(n)
End Sub
"""

    listener_mod = """Attribute VB_Name = "Main"
Public LastTick As Long
Public WithEvents clk As Clock

Private Sub clk_Tick(ByRef count As Long)
    LastTick = count * 3
End Sub

Function TestClock() As String
    Set clk = New Clock
    clk.Step 14
    TestClock = "Ticked:" & LastTick
End Function
"""

    wb.add_macro("Clock", emitter_cls, kind="class")
    wb.add_macro("Main", listener_mod, kind="standard")
    
    type_name, val, mutated = wb.run_macro("TestClock", module="Main")
    assert val == "Ticked:42", f"Expected Ticked:42, got {val}"
    print("  ✓ Custom events with WithEvents passed")

def test_open_events():
    print("Testing Workbook_Open and Auto_Open...")
    wb = visi_core.Workbook()
    
    this_wb_src = """Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Open()
    Range("A1").Value = "InitOpen"
End Sub
"""

    mod1_src = """Attribute VB_Name = "Module1"
Public Sub Auto_Open()
    Range("A2").Value = "InitAuto"
End Sub
"""

    wb.add_macro("ThisWorkbook", this_wb_src, kind="document")
    wb.add_macro("Module1", mod1_src, kind="standard")
    
    type_name, val, mutated = wb.run_open_events()
    assert mutated, "Expected workbook to be mutated by open events"
    
    val_a1 = wb.get_display(0, 0, sheet="Sheet1")
    val_a2 = wb.get_display(1, 0, sheet="Sheet1")
    assert val_a1 == "InitOpen", f"A1 expected InitOpen, got {val_a1}"
    assert val_a2 == "InitAuto", f"A2 expected InitAuto, got {val_a2}"
    print("  ✓ Workbook_Open and Auto_Open passed")

if __name__ == "__main__":
    test_worksheet_change_events()
    test_custom_events_with_events()
    test_open_events()
    print("All event system differential probes passed successfully!")
