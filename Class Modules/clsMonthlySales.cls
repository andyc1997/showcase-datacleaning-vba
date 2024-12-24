Option Explicit
Private month_year As String
Private store_types() As String
Private internal_names() As String
Private internal_values() As Double
' Initilization: reconfigure the lengths of arrays
Sub Init(n As Integer)
    ReDim store_types(n)
    ReDim internal_names(n)
    ReDim internal_values(n)
End Sub
' Iterative updates
Sub UpdateArray(n As Integer, store_type As String, name As String, value As Double)
        store_types(n) = store_type
        internal_names(n) = name
        internal_values(n) = value
End Sub
Property Let MonthYear(this_month As String)
    month_year = this_month
End Property
Property Get MonthYear() As String
    MonthYear = month_year
End Property
Property Get StoreType(n As Integer) As String
    StoreType = store_types(n)
End Property
Property Get InternalName(n As Integer) As String
    InternalName = internal_names(n)
End Property
Property Get InternalValue(n As Integer) As Double
    InternalValue = internal_values(n)
End Property
