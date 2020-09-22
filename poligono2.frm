VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCULO DE POLIGONO CERRADO  - TOPOGRAFIA"
   ClientHeight    =   6960
   ClientLeft      =   2910
   ClientTop       =   1080
   ClientWidth     =   9255
   ForeColor       =   &H00FFFFFF&
   Icon            =   "poligono2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9255
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12250
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1) Entrada de datos"
      TabPicture(0)   =   "poligono2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "resultado"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "perimetro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tabla"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "azimutcalcular"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "texto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "borrar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dialogo1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "coordenadaguardar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "coordenadacargar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "tablanombre"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "2) Angulos internos"
      TabPicture(1)   =   "poligono2.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabla3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "3) Azimut de Lados"
      TabPicture(2)   =   "poligono2.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabla2"
      Tab(2).Control(1)=   "Image1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "4) Grafico del poligono"
      TabPicture(3)   =   "poligono2.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "izquierda"
      Tab(3).Control(1)=   "derecha"
      Tab(3).Control(2)=   "abajo"
      Tab(3).Control(3)=   "arriba"
      Tab(3).Control(4)=   "Command3"
      Tab(3).Control(5)=   "Command2"
      Tab(3).Control(6)=   "grafy"
      Tab(3).Control(7)=   "grafx"
      Tab(3).Control(8)=   "Command1"
      Tab(3).Control(9)=   "guardar"
      Tab(3).Control(10)=   "Picture1"
      Tab(3).Control(11)=   "Line4"
      Tab(3).Control(12)=   "Line3(1)"
      Tab(3).Control(13)=   "Line3(0)"
      Tab(3).Control(14)=   "Label11"
      Tab(3).Control(15)=   "Label10"
      Tab(3).Control(16)=   "Label9"
      Tab(3).Control(17)=   "Label6"
      Tab(3).Control(18)=   "Label5"
      Tab(3).ControlCount=   19
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   6600
         TabIndex        =   36
         Top             =   600
         Width           =   2295
         Begin VB.OptionButton ing 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "English"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   1200
            Picture         =   "poligono2.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton esp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Español"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   120
            Picture         =   "poligono2.frx":0684
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton izquierda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67800
         Picture         =   "poligono2.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton derecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66720
         Picture         =   "poligono2.frx":0AD8
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton abajo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67200
         Picture         =   "poligono2.frx":0C22
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton arriba 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67200
         Picture         =   "poligono2.frx":0D6C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         Picture         =   "poligono2.frx":0EB6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67800
         Picture         =   "poligono2.frx":1000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox grafy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67320
         TabIndex        =   24
         Top             =   1560
         Width           =   1065
      End
      Begin VB.TextBox grafx 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67320
         TabIndex        =   23
         Top             =   1080
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid tablanombre 
         Height          =   4950
         Left            =   3600
         TabIndex        =   21
         ToolTipText     =   "Ingrese el nombre de puntos (opcional)"
         Top             =   1800
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   8731
         _Version        =   393216
         ScrollBars      =   2
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Abrir en otra  ventana"
         Height          =   540
         Left            =   -67920
         TabIndex        =   20
         Top             =   6000
         Width           =   1785
      End
      Begin VB.CommandButton coordenadacargar 
         Caption         =   "Cargar coordenadas"
         Height          =   915
         Left            =   7560
         Picture         =   "poligono2.frx":114A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton coordenadaguardar 
         Caption         =   "Guardar coordenadas"
         Height          =   915
         Left            =   6120
         Picture         =   "poligono2.frx":1294
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4200
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog dialogo1 
         Left            =   6240
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.cor"
      End
      Begin VB.CommandButton borrar 
         Caption         =   "Borrar datos y grafico"
         Height          =   915
         Left            =   6120
         Picture         =   "poligono2.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton guardar 
         Caption         =   "Guardar como .bmp"
         Height          =   675
         Left            =   -67965
         Picture         =   "poligono2.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5160
         Width           =   1785
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   -74760
         MouseIcon       =   "poligono2.frx":168A
         MousePointer    =   2  'Cross
         ScaleHeight     =   6345
         ScaleWidth      =   6345
         TabIndex        =   15
         Top             =   420
         Width           =   6375
         Begin VB.Line Line2 
            BorderStyle     =   3  'Dot
            X1              =   4800
            X2              =   3240
            Y1              =   1680
            Y2              =   3720
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   840
            X2              =   1920
            Y1              =   1320
            Y2              =   2760
         End
      End
      Begin VB.CommandButton texto 
         Caption         =   "Guardar resultados"
         Height          =   915
         Left            =   7560
         Picture         =   "poligono2.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4200
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid tabla2 
         Height          =   6210
         Left            =   -74760
         TabIndex        =   9
         ToolTipText     =   "Datos de cada lado"
         Top             =   525
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   10954
         _Version        =   393216
         ScrollBars      =   2
      End
      Begin VB.CommandButton azimutcalcular 
         Caption         =   "Calcular"
         Height          =   795
         Left            =   6840
         Picture         =   "poligono2.frx":1ADE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid tabla3 
         Height          =   6105
         Left            =   -72480
         TabIndex        =   7
         ToolTipText     =   "Angulos internos en cada vertice"
         Top             =   525
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   10769
         _Version        =   393216
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid tabla 
         Height          =   4950
         Left            =   210
         TabIndex        =   5
         ToolTipText     =   "Ingrese las coordenadas de los puntos"
         Top             =   1785
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   8731
         _Version        =   393216
         ScrollBars      =   2
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "INGRESE EL NUMERO DE VERTICES"
         Top             =   480
         Width           =   1170
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   -67440
         Picture         =   "poligono2.frx":1C28
         Top             =   2280
         Width           =   1470
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6120
         X2              =   8880
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   -68040
         X2              =   -66000
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -68040
         X2              =   -66000
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   -68040
         X2              =   -66120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Mover"
         Height          =   255
         Left            =   -67320
         TabIndex        =   35
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67320
         TabIndex        =   28
         Top             =   4400
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenadas del cursor"
         Height          =   645
         Left            =   -67800
         TabIndex        =   27
         Top             =   480
         Width           =   1590
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Y ="
         Height          =   330
         Left            =   -67680
         TabIndex        =   26
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "X ="
         Height          =   435
         Left            =   -67680
         TabIndex        =   25
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE DE LOS VERTICES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   3840
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label perimetro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6600
         TabIndex        =   14
         Top             =   6360
         Width           =   2115
      End
      Begin VB.Label resultado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6600
         TabIndex        =   13
         Top             =   5640
         Width           =   2115
      End
      Begin VB.Label Label8 
         Caption         =   "PERIMETRO DEL POLIGONO"
         Height          =   330
         Left            =   6480
         TabIndex        =   12
         Top             =   6120
         Width           =   2325
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "AREA DEL POLIGONO"
         Height          =   225
         Left            =   6720
         TabIndex        =   11
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "( MAXIMO 500 PUNTOS )"
         Height          =   225
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X : NORTE     Y : ESTE"
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   1560
         Width           =   2010
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "COORDENADAS DE LOS VERTICES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO DE LADOS / VERTICES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   3270
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declara las variables
Public areapos As Single, areaneg As Single, areatotal As Single
Public xmin As Single, xmax As Single, ymin As Single, ymax As Single
Public deltax As Single, deltay As Single
Public xcentro, ycentro As Single
Public limxmin, limxmax As Single
Public limymin, limymax As Single
Public movido As Boolean
Public deltaxzoom, deltayzoom As Single
Dim nombres(1 To 500) As String
Sub buscarlimites()
'busca los minimos y maximos de las coordenadas
'para calcular los limites del picturebox del poligono
xmin = tabla.TextMatrix(1, 1)
xmax = tabla.TextMatrix(1, 1)
ymin = tabla.TextMatrix(1, 2)
ymax = tabla.TextMatrix(1, 2)
For i = 1 To (tabla.Rows) - 1
    If tabla.TextMatrix(i, 1) < xmin Then xmin = tabla.TextMatrix(i, 1)
    If tabla.TextMatrix(i, 1) > xmax Then xmax = tabla.TextMatrix(i, 1)
    If tabla.TextMatrix(i, 2) < ymin Then ymin = tabla.TextMatrix(i, 2)
    If tabla.TextMatrix(i, 2) > ymax Then ymax = tabla.TextMatrix(i, 2)
Next i
End Sub
Sub dibujarpuntos()
If movido = False Then    'no se uso zoom ni se movio
    'los limites del picturebox segun las coordenadas
    'de los puntos del poligono
    deltay = (ymax - ymin)
    deltax = (xmax - xmin)
    ycentro = (ymin + ymax) / 2
    xcentro = (xmin + xmax) / 2
    'si es mas largo segun Y
    If deltay > deltax Then
        limymin = ycentro - deltay * 0.5 * 1.2
        limymax = ycentro + deltay * 0.5 * 1.2
        limxmin = xcentro - deltax * 0.5 * 1.2 * deltay / deltax
        limxmax = xcentro + deltax * 0.5 * 1.2 * deltay / deltax
    End If
    'si es mas largo segun X
    If deltax > deltay Then
        limxmin = xcentro - deltax * 0.5 * 1.2
        limxmax = xcentro + deltax * 0.5 * 1.2
        limymin = ycentro - deltay * 0.5 * 1.2 * deltax / deltay
        limymax = ycentro + deltay * 0.5 * 1.2 * deltax / deltay
    End If
    'si tiene las mismas proporciones
    If deltay = deltax Then
        limymin = ycentro - deltay * 0.5 * 1.2
        limymax = ycentro + deltay * 0.5 * 1.2
        limxmin = xcentro - deltax * 0.5 * 1.2
        limxmax = xcentro + deltax * 0.5 * 1.2
    End If
End If
Picture1.Height = 6375
Picture1.Width = 6375
Picture1.Scale (limymin, limxmax)-(limymax, limxmin)
'grafica los ejes
Picture1.Line (0, limxmax)-(0, limxmin), RGB(250, 0, 0)
Picture1.Line (limymin, 0)-(limymax, 0), RGB(250, 0, 0)
'grafica los lados y nombra los puntos
For i = 1 To (tabla.Rows) - 2
    Picture1.Line (tabla.TextMatrix(i, 2), tabla.TextMatrix(i, 1))-(tabla.TextMatrix(i + 1, 2), tabla.TextMatrix(i + 1, 1)), RGB(0, 0, 200)
    Picture1.ForeColor = RGB(255, 0, 0)
    Picture1.Print "  " & nombres(i + 1)
Next i
Picture1.Line (tabla.TextMatrix((tabla.Rows) - 1, 2), tabla.TextMatrix((tabla.Rows) - 1, 1))-(tabla.TextMatrix(1, 2), tabla.TextMatrix(1, 1)), RGB(0, 0, 200)
Picture1.ForeColor = RGB(255, 0, 0)
Picture1.Print "  " & nombres(1)
movido = False
End Sub
Private Sub abajo_Click()
movido = True: xmaxtemp = limxmax
xcentro = xcentro + (limxmax - limxmin) / 10
limxmax = limxmax + (limxmax - limxmin) / 10
limxmin = limxmin + (xmaxtemp - limxmin) / 10
Form1.Picture1.Cls: dibujarpuntos
End Sub
Private Sub arriba_Click()
movido = True: xmaxtemp = limxmax
xcentro = xcentro - (limxmax - limxmin) / 10
limxmax = limxmax - (limxmax - limxmin) / 10
limxmin = limxmin - (xmaxtemp - limxmin) / 10
Form1.Picture1.Cls: dibujarpuntos
End Sub
Private Sub azimutcalcular_Click()
perimetro = 0
'el calculo del area con el metodo de los trapecios
areapos = 0
areaneg = 0
areatotal = 0
For i = 1 To (tabla.Rows) - 2
    areapos = areapos + tabla.TextMatrix(i, 1) * tabla.TextMatrix(i + 1, 2)
Next i
areapos = areapos + tabla.TextMatrix((tabla.Rows) - 1, 1) * tabla.TextMatrix(1, 2)
For i = 1 To (tabla.Rows) - 2
    areaneg = areaneg - tabla.TextMatrix(i, 2) * tabla.TextMatrix(i + 1, 1)
Next i
areaneg = areaneg - tabla.TextMatrix((tabla.Rows) - 1, 2) * tabla.TextMatrix(1, 1)
areatotal = Abs((areapos + areaneg) / 2)
resultado.Caption = areatotal
buscarlimites
'el calculo de los azimut de los lados
Dim azimut(1 To 500) As Single
Dim dy(1 To 500) As Single, dx(1 To 500) As Single, rumbo(1 To 500) As Single
Dim grados(1 To 500) As Integer: Dim minutos(1 To 500) As Single
Dim segundos(1 To 500) As Single: Dim longuitud(1 To 500) As Single
Const pi = 3.141592654
For i = 1 To (tabla.Rows) - 2
    dy(i) = tabla.TextMatrix(i + 1, 2) - tabla.TextMatrix(i, 2)
    dx(i) = tabla.TextMatrix(i + 1, 1) - tabla.TextMatrix(i, 1)
    longuitud(i) = Sqr(dx(i) * dx(i) + dy(i) * dy(i))
    perimetro = perimetro + longuitud(i)
    If dx(i) = 0 Then rumbo(i) = pi / 2
    If dx(i) <> 0 Then rumbo(i) = Atn(dy(i) / dx(i))
    If dx(i) > 0 And dy(i) > 0 Then azimut(i) = (Abs(rumbo(i))) * 180 / pi
    If dx(i) < 0 And dy(i) > 0 Then azimut(i) = (pi - Abs(rumbo(i))) * 180 / pi
    If dx(i) < 0 And dy(i) < 0 Then azimut(i) = (pi + Abs(rumbo(i))) * 180 / pi
    If dx(i) > 0 And dy(i) < 0 Then azimut(i) = ((2 * pi) - Abs(rumbo(i))) * 180 / pi
    grados(i) = Fix(azimut(i))
    minutos(i) = Fix(60 * (azimut(i) - grados(i)))
    segundos(i) = 60 * (60 * (azimut(i) - grados(i)) - minutos(i))
    If segundos(i) < 0 Then
        minutos(i) = minutos(i) - 1
        segundos(i) = 60 - Abs(segundos(i))
    End If
Next i
'azimut del ultimo lado
i = (tabla.Rows) - 1
dy(i) = tabla.TextMatrix(1, 2) - tabla.TextMatrix(i, 2)
dx(i) = tabla.TextMatrix(1, 1) - tabla.TextMatrix(i, 1)
longuitud(i) = Sqr(dx(i) * dx(i) + dy(i) * dy(i))
perimetro = perimetro + longuitud(i)
If dx(i) = 0 Then rumbo(i) = pi / 2
If dx(i) <> 0 Then rumbo(i) = Atn(dy(i) / dx(i))
If dx(i) > 0 And dy(i) > 0 Then azimut(i) = (Abs(rumbo(i))) * 180 / pi
If dx(i) < 0 And dy(i) > 0 Then azimut(i) = (pi - Abs(rumbo(i))) * 180 / pi
If dx(i) < 0 And dy(i) < 0 Then azimut(i) = (pi + Abs(rumbo(i))) * 180 / pi
If dx(i) > 0 And dy(i) < 0 Then azimut(i) = ((2 * pi) - Abs(rumbo(i))) * 180 / pi
grados(i) = Fix(azimut(i))
minutos(i) = Fix(60 * (azimut(i) - grados(i)))
segundos(i) = 60 * (60 * (azimut(i) - grados(i)) - minutos(i))
If segundos(i) < 0 Then
    minutos(i) = minutos(i) - 1
    segundos(i) = 60 - Abs(segundos(i))
End If
'muestra el resultado de los lados en la tabla
For i = 1 To (tabla.Rows) - 1
    tabla2.TextMatrix(i, 1) = Format(dx(i), "######.###")
    tabla2.TextMatrix(i, 2) = Format(dy(i), "######.###")
    tabla2.TextMatrix(i, 3) = Format(longuitud(i), "######.###")
    tabla2.TextMatrix(i, 4) = grados(i)
    tabla2.TextMatrix(i, 5) = Fix(minutos(i))
    tabla2.TextMatrix(i, 6) = Format(segundos(i), "##.#")
Next i
'calculo de los angulos internos por diferencia de azimut de lados
Dim angulo(1 To 500) As Single, angulogrado(1 To 500) As Integer
Dim angulominuto(1 To 500) As Integer, angulosegundo(1 To 500) As Single
If azimut((tabla.Rows - 1)) > 180 Then angulo(1) = azimut((tabla.Rows - 1)) - azimut(1) - 180
If azimut((tabla.Rows - 1)) < 180 Then angulo(1) = azimut((tabla.Rows - 1)) - azimut(1) + 180
If angulo(1) < 0 Then angulo(1) = angulo(1) + 360
angulogrado(1) = Fix(angulo(1))
angulominuto(1) = (60 * (angulo(1) - angulogrado(1)))
angulosegundo(1) = 60 * (60 * (angulo(1) - angulogrado(1)) - angulominuto(1))
If angulosegundo(1) < 0 Then
    angulominuto(1) = angulominuto(1) - 1
    angulosegundo(1) = 60 - Abs(angulosegundo(1))
End If
For i = 2 To (tabla.Rows - 1)
    If azimut(i - 1) > 180 Then angulo(i) = azimut(i - 1) - azimut(i) - 180
    If azimut(i - 1) < 180 Then angulo(i) = azimut(i - 1) - azimut(i) + 180
    If angulo(i) < 0 Then angulo(i) = angulo(i) + 360
    angulogrado(i) = Fix(angulo(i))
    angulominuto(i) = Fix(60 * (angulo(i) - angulogrado(i)))
    angulosegundo(i) = 60 * (60 * (angulo(i) - angulogrado(i)) - angulominuto(i))
    If angulosegundo(i) < 0 Then
        angulominuto(i) = angulominuto(i) - 1
        angulosegundo(i) = 60 - Abs(angulosegundo(i))
    End If
Next i
'muestra los angulos interiores en la tabla
For i = 1 To (tabla.Rows - 1)
    tabla3.TextMatrix(i, 1) = angulogrado(i)
    tabla3.TextMatrix(i, 2) = Fix(angulominuto(i))
    tabla3.TextMatrix(i, 3) = Format(angulosegundo(i), "##.#")
Next i
perimetro.Caption = Format(perimetro, "#.###")
texto.Enabled = True
guardar.Enabled = True
Command1.Enabled = True
'pasa la tabla de nombres a un vector
For i = 1 To tablanombre.Rows - 1
    nombres(i) = tablanombre.TextMatrix(i, 1)
Next i
'borrar graficos
Form1.Picture1.Cls
grafico.Picture1.Cls
dibujarpuntos    'grafica poligono
End Sub
Private Sub borrar_Click()
'limpia los datos de las tablas, el resultado del area y el grafico del poligono
For i = 1 To tabla.Rows - 1
    For j = 1 To 2: tabla.TextMatrix(i, j) = "": Next j
    For k = 1 To 6: tabla2.TextMatrix(i, k) = "": Next k
    For L = 1 To 3: tabla3.TextMatrix(i, L) = "": Next L
    tablanombre.TextMatrix(i, 1) = ""
Next i
resultado.Caption = ""
perimetro.Caption = ""
Form1.Picture1.Cls
grafico.Picture1.Cls
guardar.Enabled = False
Command1.Enabled = False
grafico.Hide
End Sub
Private Sub Command1_Click()
grafico.Height = 9750
grafico.Width = 9465
grafico.Show
'los limites del picturebox segun las coordenadas
'de los puntos del poligono
deltay = (ymax - ymin)
deltax = (xmax - xmin)
ycentro = (ymin + ymax) / 2
xcentro = (xmin + xmax) / 2
'si es mas largo segun Y
If deltay > deltax Then
    limymin = ycentro - deltay * 0.5 * 1.2
    limymax = ycentro + deltay * 0.5 * 1.2
    limxmin = xcentro - deltax * 0.5 * 1.2 * deltay / deltax
    limxmax = xcentro + deltax * 0.5 * 1.2 * deltay / deltax
End If
'si es mas largo segun X
If deltax > deltay Then
    limxmin = xcentro - deltax * 0.5 * 1.2
    limxmax = xcentro + deltax * 0.5 * 1.2
    limymin = ycentro - deltay * 0.5 * 1.2 * deltax / deltay
    limymax = ycentro + deltay * 0.5 * 1.2 * deltax / deltay
End If
'si tiene las mismas proporciones
If deltay = deltax Then
    limymin = ycentro - deltay * 0.5 * 1.2
    limymax = ycentro + deltay * 0.5 * 1.2
    limxmin = xcentro - deltax * 0.5 * 1.2
    limxmax = xcentro + deltax * 0.5 * 1.2
End If
grafico.Picture1.Height = 9360: grafico.Picture1.Width = 9360
grafico.Picture1.Scale (limymin, limxmax)-(limymax, limxmin)
'grafica los ejes
grafico.Picture1.Line (0, limxmax)-(0, limxmin), RGB(250, 0, 0)
grafico.Picture1.Line (limymin, 0)-(limymax, 0), RGB(250, 0, 0)
'grafica los lados y nombra los puntos
For i = 1 To (tabla.Rows) - 2
    grafico.Picture1.Line (tabla.TextMatrix(i, 2), tabla.TextMatrix(i, 1))-(tabla.TextMatrix(i + 1, 2), tabla.TextMatrix(i + 1, 1)), RGB(0, 0, 200)
    grafico.Picture1.ForeColor = RGB(255, 0, 0)
    grafico.Picture1.Print "  " & nombres(i + 1)
Next i
grafico.Picture1.Line (tabla.TextMatrix((tabla.Rows) - 1, 2), tabla.TextMatrix((tabla.Rows) - 1, 1))-(tabla.TextMatrix(1, 2), tabla.TextMatrix(1, 1)), RGB(0, 0, 200)
grafico.Picture1.ForeColor = RGB(255, 0, 0)
grafico.Picture1.Print "  " & nombres(1)
End Sub
Private Sub Command2_Click()  'zoom achica
movido = True
deltaxzoom = (limxmax - limxmin) * 1.2
deltayzoom = (limymax - limymin) * 1.2
limxmin = xcentro - deltaxzoom / 2
limxmax = xcentro + deltaxzoom / 2
limymin = ycentro - deltayzoom / 2
limymax = ycentro + deltayzoom / 2
Form1.Picture1.Cls
dibujarpuntos
End Sub
Private Sub Command3_Click()   'zoom agranda
movido = True
deltaxzoom = (limxmax - limxmin) / 1.2
deltayzoom = (limymax - limymin) / 1.2
limxmin = xcentro - deltaxzoom / 2
limxmax = xcentro + deltaxzoom / 2
limymin = ycentro - deltayzoom / 2
limymax = ycentro + deltayzoom / 2
Form1.Picture1.Cls
dibujarpuntos
End Sub
Private Sub coordenadacargar_Click()
'limpia los datos de las tablas, el resultado del area y el grafico del poligono
For i = 1 To tabla.Rows - 1
    For j = 1 To 2: tabla.TextMatrix(i, j) = "": Next j
    For k = 1 To 6: tabla2.TextMatrix(i, k) = "": Next k
    For L = 1 To 3: tabla3.TextMatrix(i, L) = "": Next L
    tablanombre.TextMatrix(i, 1) = ""
Next i
resultado.Caption = ""
perimetro.Caption = ""
Form1.Picture1.Cls
grafico.Picture1.Cls
guardar.Enabled = False
Command1.Enabled = False
grafico.Hide
'carga coordenadas
Dim xtemporal As Single: Dim ytemporal As Single: Dim nomtemporal As String
dialogo1.DefaultExt = "cor"
If esp.Value = True Then
    dialogo1.Filter = "coordenadas de puntos (*.cor) | *.cor"
End If
If ing.Value = True Then
    dialogo1.Filter = "point's coordinates (*.cor) | *.cor"
End If
dialogo1.InitDir = App.Path
dialogo1.ShowOpen
arch = FreeFile
punto = 0
Open dialogo1.FileName For Input As #arch
Do While Not EOF(arch)      ' hasta que se termina el archivo
    punto = punto + 1    'cuenta los puntos antes de cargarlos
    Input #arch, xtemporal, ytemporal, nomtemporal
Loop
Text1.Text = punto
Close #arch
Open dialogo1.FileName For Input As #arch     ' lo abre para cargar
punto = 0
Do While Not EOF(arch)
    punto = punto + 1
    Input #arch, xtemporal, ytemporal, nomtemporal
    tabla.TextMatrix(punto, 1) = xtemporal
    tabla.TextMatrix(punto, 2) = ytemporal
    tablanombre.TextMatrix(punto, 1) = nomtemporal
Loop
Close #arch
End Sub
Private Sub coordenadaguardar_Click()
Dim xtemporal As Single: Dim ytemporal As Single: Dim nomtemporal As String
dialogo1.DefaultExt = "cor"
If esp.Value = True Then
    dialogo1.Filter = "coordenadas de puntos (*.cor) | *.cor"
End If
If ing.Value = True Then
    dialogo1.Filter = "point's coordinates (*.cor) | *.cor"
End If
dialogo1.InitDir = App.Path
dialogo1.ShowSave
arch = FreeFile
Open dialogo1.FileName For Output As #arch
For i = 1 To tabla.Rows - 1
    xtemporal = tabla.TextMatrix(i, 1)
    ytemporal = tabla.TextMatrix(i, 2)
    nomtemporal = tablanombre.TextMatrix(i, 1)
    Write #arch, xtemporal
    Write #arch, ytemporal
    Write #arch, nomtemporal
Next i
Close #arch
End Sub
Private Sub derecha_Click()
movido = True
ymaxtemp = limymax
ycentro = ycentro - (limymax - limymin) / 10
limymax = limymax - (limymax - limymin) / 10
limymin = limymin - (ymaxtemp - limymin) / 10
Form1.Picture1.Cls
dibujarpuntos
End Sub

Private Sub esp_Click()
esp.Value = True
Form1.Caption = "CALCULO DE POLIGONO CERRADO - TOPOGRAFIA"
' en pestaña 1===========================================
tabla.ToolTipText = "Ingrese las coordenadas de los puntos"
tablanombre.ToolTipText = "Ingrese el nombre de puntos (opcional)"
SSTab1.TabCaption(0) = "1) Entrada de datos"
SSTab1.TabCaption(1) = "2) Angulos internos"
SSTab1.TabCaption(2) = "3) Azimut de lados"
SSTab1.TabCaption(3) = "4) Grafico del poligono"
Label1.Caption = "NUMERO DE LADOS / VERTICES"
Label4.Caption = "( MAXIMO 500 PUNTOS )"
Label2(0).Caption = "COORDENADAS DE LOS VERTICES"
Label3.Caption = "X : NORTE     Y : ESTE"
Label2(1).Caption = "NOMBRE DE LOS VERTICES"
borrar.Caption = "Borrar datos y grafico"
coordenadacargar.Caption = "Cargar coordenadas"
azimutcalcular.Caption = "Calcular"
coordenadaguardar.Caption = "Guardar coordenadas"
texto.Caption = "Guardar resultados"
Label7.Caption = "AREA DEL POLIGONO"
Label8.Caption = "PERIMETRO DEL POLIGONO"
tabla.TextMatrix(0, 0) = "    Punto"
tablanombre.TextMatrix(0, 0) = "    Punto"
tabla.TextMatrix(0, 1) = "        X"
tabla.TextMatrix(0, 2) = "        Y"
tablanombre.TextMatrix(0, 1) = "  Nombre"
' en pestaña 2============================================
tabla3.ToolTipText = "Angulos internos en cada vertice"
tabla3.TextMatrix(0, 0) = "Ang. interno"
tabla3.TextMatrix(0, 1) = "   grados"
tabla3.TextMatrix(0, 2) = "  minutos"
tabla3.TextMatrix(0, 3) = " segundos"
' en pestaña 3=================================
tabla2.ToolTipText = "Datos de cada lado"
tabla2.TextMatrix(0, 0) = "     Lado"
tabla2.TextMatrix(0, 1) = "   Delta X"
tabla2.TextMatrix(0, 2) = "   Delta Y"
tabla2.TextMatrix(0, 3) = "   Long."
tabla2.TextMatrix(0, 4) = "Az grados"
tabla2.TextMatrix(0, 5) = "Az minutos"
tabla2.TextMatrix(0, 6) = "Az segundos"
' en pestaña 4===================================
Label9.Caption = "Coordenadas del cursor"
Label11.Caption = "Mover"
guardar.Caption = "Guardar como .bmp"
Command1.Caption = "Abrir en otra ventana"
' en form grafico==============================
grafico.bmp.Caption = "GUARDAR BMP"
End Sub

Private Sub Form_Load()
'no se pueden usar los botones hasta resolver
texto.Enabled = False: guardar.Enabled = False
Command1.Enabled = False: Text1.Text = 20
'dimensiones iniciales de las tablas
tabla.Rows = 21: tabla2.Rows = 21: tabla3.Rows = 21: tablanombre.Rows = 21
tabla.Cols = 3: tabla2.Cols = 7: tabla3.Cols = 4: tablanombre.Cols = 2
'encabezados de las tablas
tabla.TextMatrix(0, 0) = "    Punto"
tablanombre.TextMatrix(0, 0) = "    Punto"
tabla.TextMatrix(0, 1) = "        X"
tabla.TextMatrix(0, 2) = "        Y"
tablanombre.TextMatrix(0, 1) = "  Nombre"
tabla3.TextMatrix(0, 0) = "Ang. interno"
tabla3.TextMatrix(0, 1) = "   grados"
tabla3.TextMatrix(0, 2) = "  minutos"
tabla3.TextMatrix(0, 3) = " segundos"
tabla2.TextMatrix(0, 0) = "     Lado"
tabla2.TextMatrix(0, 1) = "   Delta X"
tabla2.TextMatrix(0, 2) = "   Delta Y"
tabla2.TextMatrix(0, 3) = "   Long."
tabla2.TextMatrix(0, 4) = "Az grados"
tabla2.TextMatrix(0, 5) = "Az minutos"
tabla2.TextMatrix(0, 6) = "Az segundos"
For i = 1 To (tabla.Rows - 1)
    tabla.TextMatrix(i, 0) = "        " & i
    tabla3.TextMatrix(i, 0) = "        " & i
    tablanombre.Col = 0: tablanombre.Row = i: tablanombre.Text = "        " & i
    If i < (tabla.Rows - 1) Then tabla2.TextMatrix(i, 0) = i & " - " & i + 1
    If i = (tabla.Rows - 1) Then tabla2.TextMatrix(i, 0) = i & " - " & "1"
Next i
areapos = 0: areaneg = 0: areatotal = 0
movido = False
Form1.Show
'prepara lineas de grafico de poligono
Line1.X1 = 0: Line1.X2 = 0: Line1.Y1 = 0: Line1.Y2 = 0
Line2.X1 = 0: Line2.X2 = 0: Line2.Y1 = 0: Line2.Y2 = 0
Line1.BorderColor = RGB(50, 100, 100)
Line2.BorderColor = RGB(50, 100, 100)
esp.Value = True  'arranca en español
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub guardar_Click()
dialogo1.DefaultExt = "bmp"
If esp.Value = True Then
    dialogo1.Filter = "Mapa de bits (*.bmp) | *.bmp"
End If
If ing.Value = True Then
    dialogo1.Filter = "Bitmap image (*.bmp) | *.bmp"
End If
dialogo1.InitDir = App.Path
dialogo1.ShowSave
Call SavePicture(Picture1.Image, dialogo1.FileName)
End Sub

Private Sub ing_Click()
ing.Value = True
Form1.Caption = "SOLVE A CLOSED POLIGON - TOPOGRAPHY"
' en pestaña 1======================================
tabla.ToolTipText = "Enter the point's coordinates"
tablanombre.ToolTipText = "Enter the name of the points (optional)"
SSTab1.TabCaption(0) = "1) Enter the data"
SSTab1.TabCaption(1) = "2) Internal angles"
SSTab1.TabCaption(2) = "3) Azimuth of sides"
SSTab1.TabCaption(3) = "4) Graphic of the poligon"
Label1.Caption = "NUMBER OF POINTS / SIDES"
Label4.Caption = "( MAX. 500 POINTS )"
Label2(0).Caption = "COORDINATES OF VERTICES"
Label3.Caption = "X : NORTH     Y : EAST"
Label2(1).Caption = "POINT'S NAMES"
borrar.Caption = "Erase data and graphic"
coordenadacargar.Caption = "Load coordinates"
azimutcalcular.Caption = "Solve"
coordenadaguardar.Caption = "Save coordinates"
texto.Caption = "Save results"
Label7.Caption = "AREA OF POLIGON"
Label8.Caption = "PERIMETER OF POLIGON"
tabla.TextMatrix(0, 0) = "    Point"
tablanombre.TextMatrix(0, 0) = "    Point"
tabla.TextMatrix(0, 1) = "        X"
tabla.TextMatrix(0, 2) = "        Y"
tablanombre.TextMatrix(0, 1) = "  Name"
' en pestaña 2================================
tabla3.ToolTipText = "Angle in each point"
tabla3.TextMatrix(0, 0) = "Angle in ..."
tabla3.TextMatrix(0, 1) = "   degrees"
tabla3.TextMatrix(0, 2) = "  minutes"
tabla3.TextMatrix(0, 3) = " seconds"
' en pestaña 3====================================
tabla2.ToolTipText = "Data of each side"
tabla2.TextMatrix(0, 0) = "     Side"
tabla2.TextMatrix(0, 1) = "   Delta X"
tabla2.TextMatrix(0, 2) = "   Delta Y"
tabla2.TextMatrix(0, 3) = "   Length"
tabla2.TextMatrix(0, 4) = "Az degrrees"
tabla2.TextMatrix(0, 5) = "Az minutes"
tabla2.TextMatrix(0, 6) = "Az seconds"
' en pestaña 4======================================
Label9.Caption = "Coordinates of the cursor"
Label11.Caption = "Move"
guardar.Caption = "Save as .bmp"
Command1.Caption = "Open in other window"
' en form grafico=================================
grafico.bmp.Caption = "SAVE BMP"
End Sub

Private Sub izquierda_Click()
movido = True
ymaxtemp = limymax
ycentro = ycentro + (limymax - limymin) / 10
limymax = limymax + (limymax - limymin) / 10
limymin = limymin + (ymaxtemp - limymin) / 10
Form1.Picture1.Cls
dibujarpuntos
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'coordenadas
grafx.Text = Format(Y, "###0.00")
grafy.Text = Format(X, "###0.00")
'lineas que siguen al mouse
Line1.X1 = limymin
Line1.X2 = limymax
Line1.Y1 = Y
Line1.Y2 = Y
Line2.Y1 = limxmin
Line2.Y2 = limxmax
Line2.X1 = X
Line2.X2 = X
End Sub

Private Sub tabla_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then           'si se aprieta backspace
        If Len(tabla.Text) > 0 Then           'si no esta vacia la celda
            tabla.Text = Left((tabla.Text), Len(tabla.Text) - 1)       'borra el ultimo caracter
        End If
Else
                 've si ya no hay un punto entrado (".")
            Dim npunto As Integer     'variable temporal cuenta los puntos
            Dim sinpunto As Boolean     'si es true  no hay puntos
            sinpunto = True             ' valor inicial
            
            If Len(tabla.Text) > 0 Then     'si hay algo en la celda
                For npunto = 1 To Len(tabla.Text)
                        If Mid(tabla.Text, Count, 1) = "." Then
                            sinpunto = False
                        End If
                Next npunto
            End If
        'acepta el menos solo si es el primer caracter
        If Len(tabla.Text) = 0 And (Chr$(KeyAscii) = "-") Then
            tabla.Text = tabla.Text + Chr$(KeyAscii)
        End If
        If (Len(tabla.Text) < 15) And IsNumeric(Chr$(KeyAscii)) Or (Chr$(KeyAscii) = ".") Then
            If sinpunto = True Or (Chr$(KeyAscii) <> ".") Then      'que no tenga dos puntos
                If Chr$(KeyAscii) = "." Then a$ = "," Else a$ = Chr$(KeyAscii)
                tabla.Text = tabla.Text + a$
            End If
           End If
End If
End Sub
Private Sub tabla2_Click()
If esp.Value = True Then
    MsgBox "EN ESTE SECTOR EL USUARIO NO DEBE INGREAR NINGUN DATO"
End If
If ing.Value = True Then
    MsgBox "THE USER DON'T HAVE TO ENTER DATA IN THIS SECTOR"
End If
End Sub
Private Sub tabla3_Click()
If esp.Value = True Then
    MsgBox "EN ESTE SECTOR EL USUARIO NO DEBE INGREAR NINGUN DATO"
End If
If ing.Value = True Then
    MsgBox "THE USER DON'T HAVE TO ENTER DATA IN THIS SECTOR"
End If
End Sub
Private Sub tablanombre_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then          'si se aprieta backspace
        If Len(tablanombre.Text) > 0 Then         'si no esta vacia la celda
            tablanombre.Text = Left((tablanombre.Text), Len(tablanombre.Text) - 1) 'borra el ultimo caracter
        End If
    Else
            tablanombre.Text = tablanombre.Text + Chr$(KeyAscii)
    End If
End Sub
Private Sub Text1_Change()
'si se cambia la cantidad de puntos
'se actualizan la cantidad de filas de las tablas
n = Val(Text1.Text)
tabla.Rows = n + 1
tabla2.Rows = n + 1
tabla3.Rows = n + 1
tablanombre.Rows = n + 1
'se vuelve a escribir los encabezados de tablas
For i = 1 To n
    tabla.TextMatrix(i, 0) = "        " & i
    tabla3.TextMatrix(i, 0) = "        " & i
    tablanombre.TextMatrix(i, 0) = "        " & i
    If i < (tabla.Rows - 1) Then tabla2.TextMatrix(i, 0) = i & " - " & i + 1
    If i = (tabla.Rows - 1) Then tabla2.TextMatrix(i, 0) = i & " - " & "1"
Next i
puntos = n
End Sub
Private Sub texto_Click()
' guarda resultados en español
If esp.Value = True Then
    dialogo1.DefaultExt = "txt"
    dialogo1.Filter = "resultados (*.txt) | *.txt"
    dialogo1.InitDir = App.Path
    dialogo1.ShowSave
    arch = FreeFile
    Open dialogo1.FileName For Output As #arch
    Print #arch, "================================================================="
    Print #arch, "COORDENADAS DE LOS PUNTOS ( X Norte   Y Este )"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Punto"; Tab(23); "X"; Tab(38); "Y"
    For i = 1 To tabla.Rows - 1
        Print #arch, tablanombre.TextMatrix(i, 1); Tab(20); tabla.TextMatrix(i, 1); Tab(35); tabla.TextMatrix(i, 2)
    Next i
    Print #arch, ""
    Print #arch, ""
    Print #arch, "ANGULOS INTERNOS EN LOS PUNTOS"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Punto"; Tab(20); "Grados"; Tab(29); "Minutos"; Tab(38); "Segundos"
    For i = 1 To tabla.Rows - 1
        Print #arch, tablanombre.TextMatrix(i, 1); Tab(23); tabla3.TextMatrix(i, 1); Tab(32); tabla3.TextMatrix(i, 2); Tab(41); tabla3.TextMatrix(i, 3)
    Next i
    Print #arch, ""
    Print #arch, ""
    Print #arch, "LONGUITUD DE LOS LADOS"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Lado"; Tab(22); "Delta x"; Tab(35); "Delta y"; Tab(47); "Longuitud"
    For i = 1 To tabla.Rows - 2
        Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(i + 1, 1); Tab(22); tabla2.TextMatrix(i, 1); Tab(36); tabla2.TextMatrix(i, 2); Tab(48); tabla2.TextMatrix(i, 3)
    Next i
    i = tabla.Rows - 1
    Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(1, 1); Tab(22); tabla2.TextMatrix(i, 1); Tab(36); tabla2.TextMatrix(i, 2); Tab(48); tabla2.TextMatrix(i, 3)
    Print #arch, ""
    Print #arch, ""
    Print #arch, "AZIMUT DE LOS LADOS"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Lado"; Tab(22); "Grados"; Tab(35); "Minutos"; Tab(47); "Segundos"
    For i = 1 To tabla.Rows - 2
        Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(i + 1, 1); Tab(22); tabla2.TextMatrix(i, 4); Tab(36); tabla2.TextMatrix(i, 5); Tab(48); tabla2.TextMatrix(i, 6)
    Next i
    i = tabla.Rows - 1
    Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(1, 1); Tab(22); tabla2.TextMatrix(i, 4); Tab(36); tabla2.TextMatrix(i, 5); Tab(48); tabla2.TextMatrix(i, 6)
    Print #arch, ""
    Print #arch, ""
    Print #arch, "PERIMETRO"
    Print #arch, "-------------"
    Print #arch, Format(perimetro, "#.###")
    Print #arch, ""
    Print #arch, "AREA"
    Print #arch, "-------------"
    Print #arch, resultado.Caption
    Print #arch, ""
    Print #arch, "==================FIN DE ARCHIVO==============================================="
    Close #arch
End If
' guarda resultados en ingles
If ing.Value = True Then
    dialogo1.DefaultExt = "txt"
    dialogo1.Filter = "results (*.txt) | *.txt"
    dialogo1.InitDir = App.Path
    dialogo1.ShowSave
    arch = FreeFile
    Open dialogo1.FileName For Output As #arch
    Print #arch, "================================================================="
    Print #arch, "COORDINATES OF THE POINTS  ( X North   Y East )"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Point"; Tab(23); "X"; Tab(38); "Y"
    For i = 1 To tabla.Rows - 1
        Print #arch, tablanombre.TextMatrix(i, 1); Tab(20); tabla.TextMatrix(i, 1); Tab(35); tabla.TextMatrix(i, 2)
    Next i
    Print #arch, ""
    Print #arch, ""
    Print #arch, "INTERNAL ANGLES IN THE POINTS"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Point"; Tab(20); "Degrees"; Tab(29); "Minutes"; Tab(38); "Seconds"
    For i = 1 To tabla.Rows - 1
        Print #arch, tablanombre.TextMatrix(i, 1); Tab(23); tabla3.TextMatrix(i, 1); Tab(32); tabla3.TextMatrix(i, 2); Tab(41); tabla3.TextMatrix(i, 3)
    Next i
    Print #arch, ""
    Print #arch, ""
    Print #arch, "LENGTH OF THE SIDES"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Side"; Tab(22); "Delta x"; Tab(35); "Delta y"; Tab(47); "Length"
    For i = 1 To tabla.Rows - 2
        Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(i + 1, 1); Tab(22); tabla2.TextMatrix(i, 1); Tab(36); tabla2.TextMatrix(i, 2); Tab(48); tabla2.TextMatrix(i, 3)
    Next i
    i = tabla.Rows - 1
    Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(1, 1); Tab(22); tabla2.TextMatrix(i, 1); Tab(36); tabla2.TextMatrix(i, 2); Tab(48); tabla2.TextMatrix(i, 3)
    Print #arch, ""
    Print #arch, ""
    Print #arch, "AZIMUT OF THE SIDES"
    Print #arch, "-----------------------------------------------------------------"
    Print #arch, "Side"; Tab(22); "Degrees"; Tab(35); "Minutes"; Tab(47); "Seconds"
    For i = 1 To tabla.Rows - 2
        Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(i + 1, 1); Tab(22); tabla2.TextMatrix(i, 4); Tab(36); tabla2.TextMatrix(i, 5); Tab(48); tabla2.TextMatrix(i, 6)
    Next i
    i = tabla.Rows - 1
    Print #arch, tablanombre.TextMatrix(i, 1) & "-"; tablanombre.TextMatrix(1, 1); Tab(22); tabla2.TextMatrix(i, 4); Tab(36); tabla2.TextMatrix(i, 5); Tab(48); tabla2.TextMatrix(i, 6)
    Print #arch, ""
    Print #arch, ""
    Print #arch, "PERIMETER"
    Print #arch, "-------------"
    Print #arch, Format(perimetro, "#.###")
    Print #arch, ""
    Print #arch, "AREA"
    Print #arch, "-------------"
    Print #arch, resultado.Caption
    Print #arch, ""
    Print #arch, "==================FEND OF FILE==============================================="
    Close #arch
End If
End Sub


