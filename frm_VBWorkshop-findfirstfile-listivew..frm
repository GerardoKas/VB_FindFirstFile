VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   11033
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "c:\windows"
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "*.bmp"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDirectorio 
      Caption         =   "Directorio"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblArchivo 
      Caption         =   "Archivo"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblCantidad 
      Caption         =   "0"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arks() As String
Dim ruta As String
Private Sub Command1_Click()
Dim textoBusqueda As String
Dim total As Long, i As Long
Dim Item As ListItem
ruta = Text1: textoBusqueda = Text2

Me.Caption = "Leyendo ..."
total = VBWorkshop.FindFilesAPI(ruta, textoBusqueda, 0, arks())
MsgBox "Hallados " & (UBound(arks, 2) + 1) & " Archivos"
Me.Caption = "Hecho "
Me.Caption = "Convirtiendo ... "
For i = 0 To UBound(arks, 2)
  '  lblCantidad = arks(0, i)
    Set Item = ListView1.ListItems.Add(, , arks(0, i))
    Item.ListSubItems.Add , , arks(1, i)
    
    
Next
Me.Caption = "Hecho Lectura "

'MsgBox UBound(arks, 1)
'ReDim arks(1, 0)


End Sub


Private Sub Form_Load()
ListView1.ColumnHeaders.Add , , "Directorio"
ListView1.ColumnHeaders.Add , , "Archivo"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload Form1
End

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'msgbox ListView1.ListItems.Item
Dim f As String
Dim d As String
d = Item.Text
f = Item.ListSubItems(1)
'MsgBox d & " | " & f

Shell "explorer " & d & f, vbNormalFocus

End Sub
