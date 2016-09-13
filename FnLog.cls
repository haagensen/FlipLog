VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FnLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbModoDebug As Boolean
Private msLogFile As String

Public Property Let Ativar(ByVal v As Boolean)
    mbModoDebug = v
End Property
Public Property Get Ativar() As Boolean
    Ativar = mbModoDebug
End Property

Public Property Let LogFile(ByVal v As String)
    msLogFile = v
End Property
Public Property Get LogFile() As String
    LogFile = msLogFile
End Property

Public Sub DebugLog(ByVal sTexto As String)

    Dim f As Integer

    On Error Resume Next

    If Not mbModoDebug Then Exit Sub

    f = FreeFile
    Open msLogFile For Append As #f

    If sTexto <> "" Then
        Print #f, Format$(Now, "yy-mm-dd hh:nn:ss") & " - " & sTexto
    Else
        Print #f, ""
    End If

    Close #f

End Sub

