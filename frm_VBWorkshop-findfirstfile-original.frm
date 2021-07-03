VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Text            =   "*.txt"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "c:\windows"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCantidad 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblArchivo 
      Caption         =   "Label2"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arks() As String
Dim ruta As String
Private Sub Command1_Click()
Dim textoBusqueda As String
ruta = Text1
textoBusqueda = Text2
VBWorkshop.FindFilesAPI ruta, textoBusqueda, 0, arks()
MsgBox UBound(arks, 1)

End Sub

