'
' DESCRIPTION:
' This model contain some usefull methods
' Contributed by Emma                                                                                                                                               's super handsome boyfriend                         :)
'
Option Explicit

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Public Const HH = vbCrLf

#If Mac Then
    ' ignore
#Else
    #If VBA7 Then
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    #Else
        Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
        Declare Function CloseClipboard Lib "User32" () As Long
        Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Declare Function EmptyClipboard Lib "User32" () As Long
        Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
        Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

    #End If

#End If


' ******************************************************
' *                                                    *
' *                 function blow                      *
' *                                                    *
' ******************************************************
'

'
' DESCRIPTION:
' session.findById
' 调用 例子
' HUGA = ESessionCheck(session, "wnd[2]/usr/txtMESSTXT2")
'      等价于 HUGA = session.findById( "wnd[2]/usr/txtMESSTXT2").Text
' 如果报错则返回 空字符  ""
Function ESessionCheck(session As Variant, title As String) As String
On Error GoTo errline
ESessionCheck = session.findById(title).Text

If False Then
errline:
ESessionCheck = ""
End If

End Function

' ******************************************************
' *                                                    *
' *                  sub blow                          *
' *                                                    *
' ******************************************************
'

'
' DESCRIPTION:
' copy string to clipbpard
'
Sub ECopyNew(MyString As String)
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory   As LongPtr
            Dim lpGlobalMemory    As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory   As Long
            Dim lpGlobalMemory    As Long
        #End If
        Dim x   As Long
        
        hGlobalMemory = GlobalAlloc(GHND, LenB(StrConv(MyString, vbFromUnicode)) + 1)
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Could not unlock memory location. Copy aborted."
            GoTo OutOfHere2
        End If

        If OpenClipboard(0&) = 0 Then
            MsgBox "Could not open the Clipboard. Copy aborted."
            Exit Sub
        End If

        x = EmptyClipboard()

        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
OutOfHere2:
        If CloseClipboard() = 0 Then
            MsgBox "Could not close Clipboard."
        End If
    #End If
End Sub


'
' DESCRIPTION:
' this sub can be ignored. it is a system method,  and it may have a bug in win 10
'
Sub ECopy(value As String)
    Dim MyDataObj As New DataObject
    
    'clear current clipboard
    With MyDataObj
    .SetText ""
    .PutInClipboard
    End With
    
    'put value into clipboard
    With MyDataObj
    .SetText value
    .PutInClipboard
    End With
    
End Sub



