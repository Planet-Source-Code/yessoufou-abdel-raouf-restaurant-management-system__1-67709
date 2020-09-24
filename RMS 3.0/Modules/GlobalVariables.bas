Attribute VB_Name = "GlobalVariables"
Option Explicit

Public MaxBooks As Integer
Public Fines As Single
Public LibFName As String
Public LibMname As String
Public LibLname As String
Public LibPass As String
Public LibInsti As String
Public LibUser As String
Public OverDue As Integer
Public DueBooks As Integer
Public BorrowedBooks As Integer
'**************************************************************
'Title: RestauSys                                             *
'application domain: Restaurant Management System             *
'Description:                                                 *
'Version: 1.0                                                 *
'Author: YESSOUFOU Abdel Raouf                                *
'Copyright: All right reserved by the author Abdel Raouf.     *
'   No part of this source code shall be copied or reused     *
'   without the author's notice.2005-2006                     *
'**************************************************************
Public TotalBooks As Integer
Public LastInLib As String
Public TitlesNum As Integer

Public Const HiLyt = "{HOME}+{END}"

Public Main_On As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long


'Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Dim tgtButton As lvButtons_H
'
'CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
'Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))  ' fire the button's event
'CopyMemory tgtButton, 0&, &H4                                    ' erase this instance
'End Function
'
'Public Sub CenterFrm(ByVal Parentfrm As MDIForm, ByVal Childfrm As Form) 'used for the frmInsignia
'
'    Childfrm.Left = (Parentfrm.Width \ 2) - (Childfrm.Width \ 2)
'    Childfrm.Top = (Parentfrm.ScaleHeight \ 2) - (Childfrm.Height \ 2)
'
'End Sub
