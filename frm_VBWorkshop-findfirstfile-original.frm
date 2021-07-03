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
      Text            =   "*.bmp"
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCantidad 
      Caption         =   "0"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblArchivo 
      Caption         =   "Archivo"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Directorio"
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
List1.Clear
Label1 = VBWorkshop.FindFilesAPI(ruta, textoBusqueda, 0, arks())
MsgBox "Hallados " & (UBound(arks, 2) + 1) & " Archivos"
ReDim arks(1, 0)


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload Form2
End


End Sub
