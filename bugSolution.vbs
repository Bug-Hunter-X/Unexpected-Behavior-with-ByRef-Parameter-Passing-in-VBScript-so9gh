Function f(ByRef a)
  a = a + 1
end function

x = 10
f x
MsgBox x ' Displays 10, not 11

' Corrected version:
Function g(ByRef a)
  Set a = CreateObject("Scripting.Dictionary")
  a.Add "Value", a.Item("Value") + 1
end function

x = 10
g x
MsgBox x 'Displays 10

'Another better solution
Function h(ByVal a)
  a = a + 1
  h = a
end function

x = 10
y = h(x)
MsgBox x ' Displays 10
MsgBox y ' Displays 11