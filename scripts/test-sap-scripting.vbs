' Minimal SAP GUI Scripting test — run while logged into a transaction (not just Logon Pad).
On Error Resume Next

Dim SapGuiAuto, application, connection, session

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If application Is Nothing Then
  WScript.Echo "FAIL: GetScriptingEngine is NULL"
  WScript.Echo "Log off SAP completely, log back in, open a transaction, run again."
  WScript.Echo "Also check RZ11: sapgui/user_scripting_per_user (may need S_SCR user auth)."
  WScript.Quit 2
End If

If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

If Not IsObject(session) Then
  WScript.Echo "WARN: No session — open a transaction (e.g. VA01, ME2N)"
  WScript.Quit 0
End If

WScript.Echo "OK: Scripting engine connected"

Set info = session.Info
WScript.Echo "Transaction: " & info.Transaction
WScript.Echo "Program: " & info.Program
WScript.Echo "Screen: " & info.ScreenNumber

Set fld = session.FindById("wnd[0]/usr/ctxtVBAK-AUART")
If fld Is Nothing Then
  WScript.Echo "Order Type field not on this screen (expected unless on VA01 initial screen)"
Else
  WScript.Echo "Order Type: " & fld.Text
End If
