VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form grafico 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GRAFICO POLIGONO"
   ClientHeight    =   9375
   ClientLeft      =   945
   ClientTop       =   795
   ClientWidth     =   9375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9360
      Left            =   0
      ScaleHeight     =   9300
      ScaleWidth      =   9300
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      Begin MSComDlg.CommonDialog dialogo5 
         Left            =   1575
         Top             =   -210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu bmp 
      Caption         =   "GUARDAR BMP"
   End
End
Attribute VB_Name = "grafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bmp_Click()
dialogo5.DefaultExt = "bmp"
If Form1.esp.Value = True Then
    dialogo5.Filter = "Mapa de bits (*.bmp) | *.bmp"
End If
If Form1.ing.Value = True Then
    dialogo5.Filter = "Bitmap image(*.bmp) | *.bmp"
End If
dialogo5.InitDir = App.Path
dialogo5.ShowSave
Call SavePicture(grafico.Picture1.Image, dialogo5.FileName)
End Sub


