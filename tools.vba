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
Public Const CP_ACP = 0 ' default to ANSI code page
Public Const CP_UTF8 = 65001 ' default to UTF-8 code page

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
                                                                
#If Win64 And VBA7 Then
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStrPtr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    Private Declare  Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare  Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If                                                                

' ******************************************************
' *                                                    *
' *                 function blow                      *
' *                                                    *
' ******************************************************
'

'string to UTF8
Public Function EncodeToBytes(ByVal sData As String) As Byte() ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), -1, 0, 0, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize, 0, 0
    EncodeToBytes = aRetn
    Erase aRetn
End Function

' UTF8 to string
Public Function DecodeToBytes(ByVal sData As String) As Byte() ' Note: Len(sData) > 0
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sData), -1, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To 2 * nSize - 1) As Byte
    MultiByteToWideChar CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize
    DecodeToBytes = aRetn
    Erase aRetn
End Function

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
        lpGlobalMemory = lstrcpy(lpGlobalMemory, EncodeToBytes(MyString))

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



