VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Judul Form Rata Tengah"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldsize As Long
Private Sub Form_Resize()
    If Me.Width = oldsize Then 'Jika lebar form berubah
       Exit Sub 'tidak perlu mengubah letak captionnya.
    Else
        CenterC Me
        oldsize = Me.Width
    End If
End Sub
  
Private Sub Form_Load()
    CenterC Me
    oldsize = Me.Width
End Sub

