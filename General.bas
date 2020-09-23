Attribute VB_Name = "General"
Public Conn As ADODB.Connection
Public Tpx  As ADODB.Recordset

' For put a windows in the middle of the screen
' FrmChild  = Windows to center
' FrmParent = MDI Windows (Optional)
Public Sub CenterForm(FrmChild As Form, Optional FrmParent As Variant)
    Dim iTop As Integer, iLeft As Integer
    If Not IsMissing(FrmParent) Then
        iTop = (FrmParent.ScaleHeight - FrmChild.Height) \ 2
        iLeft = (FrmParent.ScaleWidth - FrmChild.Width) \ 2
    Else
        iTop = (Screen.Height - FrmChild.Height) \ 2
        iLeft = (Screen.Width - FrmChild.Width) \ 2
    End If
    If iTop And iLeft Then
        FrmChild.Move iLeft, iTop
    End If
End Sub

Function NoNulo(Vrx As Variant) As String
    If IsNull(Vrx) Then
        NoNulo = ""
    Else
        NoNulo = Vrx
    End If
End Function
