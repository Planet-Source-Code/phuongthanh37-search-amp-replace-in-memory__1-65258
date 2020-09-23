VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Search & Replace in Memory
'Coded by PhuongThanh37 in VB5.0
'http://dasaco.net
'http://donganhol.com
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 _
    As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const WM_CLOSE = &H10
Private Const WM_SETTEXT = &HC
    
Private Function RepInMem(hwnd As Long, Address As Long, Optional strReplaceWith As String) As Boolean
    On Error Resume Next
    Dim pId As Long
    Dim pHandle As Long
    Dim bytValue As Long
    Dim i As Long
   
    GetWindowThreadProcessId hwnd, pId
    pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pId)

    If pHandle = 0 Then MsgBox "Unable to open process!", vbCritical: Exit Function
    If Address = 0 Then Exit Function
    
    If LenB(strReplaceWith) <> 0 Then
        'Coded by PhuongThanh37
        RepInMem = WriteProcessMemory(pHandle, Address, StrPtr(strReplaceWith), LenB(strReplaceWith), 0&)
    End If
    CloseHandle pHandle
End Function

Private Function SearchInMem(hwnd As Long, AdrStart As Long, AdrStop As Long, strFind As Variant, ByRef AdrFound As Long) As Boolean
On Error Resume Next
    Dim pId As Long
    Dim pHandle As Long
    Dim bfLg As Long
    Dim i As Long, Founds As Long
    Dim LenSrc As Integer, tmpCheck
    
    GetWindowThreadProcessId hwnd, pId
    pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pId)
    If pHandle = 0 Then MsgBox "Unable to open process!", vbCritical: Exit Function
    If AdrStop <= AdrStart Then Exit Function
    DoEvents
    
For LenSrc = 0 To 255
On Error Resume Next
    tmpCheck = strFind(LenSrc)
    If Err.Number <> 0 Then Exit For
Next
i = AdrStart
Do
  ReadProcessMemory pHandle, i, bfLg, 1, 0&
  If bfLg = strFind(Founds) Then
    Founds = Founds + 1
    If Founds = LenSrc Then
        i = i - Founds + 1
        AdrFound = i
        'Coded by PhuongThanh37
        SearchInMem = True
        Exit Do
    End If
  Else
    If i > 0 Then i = i - Founds
    Founds = 0
  End If
  i = i + 1
Loop Until i >= AdrStop

End Function

Private Sub Form_Load()
Dim hNote As Long
Dim Adr1 As Long
Dim stFix1 As Variant
Dim S1 As String
Const sName As String = "PhuongThanh37"

MsgBox "Step 1: Start Notepad app", vbInformation
If Shell("notepad", vbNormalNoFocus) = 0 Then Exit Sub
DoEvents
hNote = FindWindow("Notepad", vbNullString)
Sleep 1000

MsgBox "Step 2: Type text PhuongThanh37 to Notepad", vbInformation
Call SendMessage(FindWindowEx(hNote, 0&, "edit", vbNullString), WM_SETTEXT, 0&, _
    "Search & Replace in Memory" + vbCrLf + _
    "Coded by PhuongThanh37 in VB5.0" + vbCrLf + _
    "http://dasaco.net" + vbCrLf + _
    "http://donganhol.com")
Sleep 1500

MsgBox "Step 3: Search ""thanh37phuong"" on Notepad"
SendMessage hNote, &H111, 21, 1&
Dim hNoteFind As Long
hNoteFind = FindWindow("#32770", "Find")
Call SendMessage(FindWindowEx(hNoteFind, 0&, "edit", vbNullString), WM_SETTEXT, 0&, _
    "thanh37phuong")
Sleep 2000

MsgBox "Step 4: Click ""Find Next"" button" + vbCrLf + _
        "Return ""cannot find ""thanh37phuong"" """
SetForegroundWindow hNoteFind
DoEvents
Dim button As Long
button = FindWindowEx(hNoteFind, 0&, "button", "&Find Next")
Call SendMessage(button, &H100, &H20, 0&)
Call SendMessage(button, &H101, &H20, 0&)
Sleep 4000

TT1:
MsgBox "Close WINDOW Find in Notepad then " + vbLf + _
    "Click OK button to continue Step5"
If SetForegroundWindow(hNoteFind) Then GoTo TT1
DoEvents

MsgBox "Step 5: Search & Replace ""thanh37phuong""" + vbLf + _
    "to ""PhuongThanh37"" in Entry Memory" + vbLf + _
    "Click OK to WAIT..."
'(&H74, &H0, &H68, &H0, ..., &H6E, &H0, &H67, &H0) = "thanh37phuong"
stFix1 = Array(&H74, &H0, &H68, &H0, &H61, &H0, &H6E, &H0, &H68, &H0, &H33, &H0, &H37, &H0, _
&H70, &H0) ', &H68, &H0, &H75, &H0, &H6F, &H0, &H6E, &H0, &H67, &H0)
If SearchInMem(hNote, &H10087E0, &H100B9E0, stFix1, Adr1) Then If RepInMem(hNote, Adr1, sName) Then _
    MsgBox "Step 6: Find Next in Notepad" + vbLf + _
        "Click menu ""Edit"", click ""Find Next""" + vbLf + _
        "ohh Notepad FOUND ""PhuongThanh37"""
SendMessage hNote, &H111, 22, 1&
DoEvents
SetForegroundWindow hNote
End
End Sub
