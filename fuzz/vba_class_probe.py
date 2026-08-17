#!/usr/bin/env python3
"""
VBA Class Module and Object Model Differential Probe
===================================================
Probes and verifies:
1. Class module property Get / Let / Set procedures.
2. Default member dispatch (Attribute Item.VB_UserMemId = 0).
3. Class_Initialize and Class_Terminate lifecycle via refcounting.
4. Lazy instantiation with `Dim As New`.
5. `TypeOf ... Is` and `TypeName(...)` verification for user-defined classes.
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

def test_class_properties_and_methods():
    print("Testing Class properties and methods...")
    wb = visi_core.Workbook()
    
    person_cls = """Attribute VB_Name = "Person"
Private m_name As String
Private m_age As Long

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(val As String)
    m_name = val
End Property

Public Property Get Age() As Long
    Age = m_age
End Property

Public Property Let Age(val As Long)
    m_age = val
End Property

Public Function Describe() As String
    Describe = Me.Name & " is " & Me.Age
End Function
"""

    main_mod = """Attribute VB_Name = "Main"
Function RunTest() As String
    Dim p As Person
    Set p = New Person
    p.Name = "Bob"
    p.Age = 30
    RunTest = p.Describe() & "|" & TypeName(p) & "|" & (TypeOf p Is Person) & "|" & (TypeOf p Is Object)
End Function
"""

    wb.add_macro("Person", person_cls, kind="class")
    wb.add_macro("Main", main_mod, kind="standard")
    
    type_name, val, mutated = wb.run_macro("RunTest", module="Main")
    assert val == "Bob is 30|Person|True|True", f"Unexpected result: {val}"
    print("  ✓ Class properties and methods passed")

def test_auto_new_and_lifecycle():
    print("Testing Dim As New and Class_Initialize / Terminate...")
    wb = visi_core.Workbook()
    
    counter_cls = """Attribute VB_Name = "Counter"
Public Count As Long
Private Sub Class_Initialize()
    Count = 10
End Sub
Private Sub Class_Terminate()
    Count = 0
End Sub
"""

    main_mod = """Attribute VB_Name = "Main"
Function TestLifecycle() As String
    Dim c As New Counter
    Dim a As Long, b As Long
    a = c.Count
    c.Count = 50
    Set c = Nothing
    b = c.Count
    TestLifecycle = a & "|" & b
End Function
"""

    wb.add_macro("Counter", counter_cls, kind="class")
    wb.add_macro("Main", main_mod, kind="standard")
    
    type_name, val, mutated = wb.run_macro("TestLifecycle", module="Main")
    assert val == "10|10", f"Unexpected result: {val}"
    print("  ✓ Dim As New and lifecycle passed")

def test_default_member():
    print("Testing default member dispatch (VB_UserMemId = 0)...")
    wb = visi_core.Workbook()
    
    vector_cls = """Attribute VB_Name = "Vector"
Private m_items(10) As Long

Public Property Get Item(idx As Long) As Long
    Attribute Item.VB_UserMemId = 0
    Item = idx * 10
End Property
"""

    main_mod = """Attribute VB_Name = "Main"
Function TestDefault() As Long
    Dim v As Vector
    Set v = New Vector
    TestDefault = v(4)
End Function
"""

    wb.add_macro("Vector", vector_cls, kind="class")
    wb.add_macro("Main", main_mod, kind="standard")
    
    type_name, val, mutated = wb.run_macro("TestDefault", module="Main")
    assert val == "40", f"Unexpected result: {val}"
    print("  ✓ Default member dispatch passed")

if __name__ == "__main__":
    test_class_properties_and_methods()
    test_auto_new_and_lifecycle()
    test_default_member()
    print("All class module differential probes passed successfully!")
