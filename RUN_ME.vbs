Option Explicit

Dim xlApp, xlBook, WshShell, CurDir

'SET CURRENT DIRECTORY
Set WshShell = CreateObject("WScript.Shell")
CurDir = WshShell.CurrentDirectory

'SET EXCEL FEATURES
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(CurDir & "\ARCHIVO.xlsm", 0, True)
xlApp.Run "MAIN"
xlBook.Close
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing

WScript.Quit
