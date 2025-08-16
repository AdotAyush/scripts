Set oLocator = CreateObject("WbemScripting.SWbemLocator")
Set oServices = oLocator.ConnectServer(".", "root\cimv2")

Set objShell = CreateObject("WScript.Shell")

While (True)
    Set oResults = oServices.ExecQuery("Select * from Win32_Battery")
    For Each oResult in oResults
        iPercent = oResult.EstimatedChargeRemaining
        bCharging = oResult.BatteryStatus = 6 ' 6 = Charging
    Next

    If bCharging And (iPercent >= 99) Then
        popupHTML = "mshta.exe ""about:<html><head>" & _
            "<style>" & _
            "body{font-family:Segoe UI;background:#1a1a1a;color:#f2f2f2;text-align:center;padding:20px;}" & _
            ".alert{background:#2e7d32;border-radius:12px;padding:20px;box-shadow:0 4px 12px rgba(0,0,0,.5);}" & _
            "h2{margin:0;font-size:22px;}" & _
            "</style></head><body>" & _
            "<div class='alert'><h2>🔋 Battery at 99%</h2><p>Please unplug the charger.</p></div>" & _
            "<script>" & _
            "var snd=new Audio('https://actions.google.com/sounds/v1/alarms/alarm_clock.ogg');" & _
            "snd.play();" & _
            "setTimeout(function(){window.close();},15000);" & _
            "</script></body></html>"""

        objShell.Run popupHTML, 1, False
    End If

    WScript.Sleep 300000 ' 5 minutes
Wend
