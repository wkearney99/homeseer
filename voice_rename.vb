Sub Main(ByVal Param As Object)

	Dim theCount
	theCount = 0

	Dim Debug As Boolean = False
	Dim logName As String = "Voice_command: "

	Dim dv As Scheduler.Classes.DeviceClass
	Dim EN As Scheduler.Classes.clsDeviceEnumeration
	EN = hs.GetDeviceEnumerator

	If EN Is Nothing Then
		hs.WriteLog(logName, "Error getting Enumerator")
		Exit Sub
	End If

	Do
		dv = EN.GetNext
		If dv Is Nothing Then Continue Do

		Dim s
		Dim vc
		Dim av

		s = dv.name(Nothing)
		vc = dv.voicecommand(Nothing)
		av = dv.MISC_Check(Nothing, Enums.dvMISC.AUTO_VOICE_COMMAND)

		If (av) Then
			If (vc = "") Then
				hs.writelog(logName, "Name:  " & s)
				hs.writelog(logName, "Voice: " & vc)
				hs.writelog(logName, "Auto:  " & av)

				dv.voicecommand(Nothing) = dv.name(Nothing)
				hs.writelog(logName, "Voice command copied for: " & s)

				theCount += 1
			End If
		End If
	Loop Until EN.Finished

	hs.writelog(logName, theCount & " voice devices detected without command.")
End Sub