VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   10425
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2280
      Width           =   8775
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1S = False
Me.Hide
Cancel = 1
End Sub
