# super-fiesta
Dim oDomain As IADs Dim oLargeInt As LargeInteger  Set oDomain = GetObject("LDAP://DC=fabrikam,DC=com") Set oLargeInt = oDomain.Get("creationTime")  Debug.Print oLargeInt.HighPart Debug.Print oLargeInt.LowPart  strTemp = "&amp;H" + CStr(Hex(oLargeInt.HighPart)) + _      CStr(Hex(oLargeInt.LowPart)) Debug.Print strTemp
