VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QueryBuilder"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Too 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   847
      ButtonWidth     =   1244
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Re-Connect to another BD"
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Execute Query"
            Object.Tag             =   "2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Cancel Query Execution"
            Object.Tag             =   "3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Set Query on Clipboard"
            Object.Tag             =   "4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Retrieve Query from Clipboard"
            Object.Tag             =   "5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Set Query on own ClipBoard"
            Object.Tag             =   "6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Retrieve query from own ClipBoard"
            Object.Tag             =   "7"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Create a variable on ClipBoard with the Query (Direct export to VB)"
            Object.Tag             =   "8"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Save Query"
            Object.Tag             =   "9"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Load Query"
            Object.Tag             =   "10"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   5700
      Top             =   2970
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic 
      Align           =   3  'Align Left
      Height          =   5910
      Index           =   1
      Left            =   4125
      ScaleHeight     =   5850
      ScaleWidth      =   7545
      TabIndex        =   3
      Top             =   480
      Width           =   7605
      Begin TabDlg.SSTab SSTab1 
         Height          =   5850
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10319
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Query"
         TabPicture(0)   =   "frm.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Results"
         TabPicture(1)   =   "frm.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Configure"
         TabPicture(2)   =   "frm.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame2"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame1"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000A&
            Caption         =   "Data Retrieval"
            Height          =   1035
            Left            =   -74940
            TabIndex        =   12
            Top             =   150
            Width           =   5085
            Begin VB.TextBox Text2 
               Height          =   330
               Left            =   1260
               MaxLength       =   3
               TabIndex        =   13
               Text            =   "10"
               Top             =   600
               Width           =   615
            End
            Begin ComctlLib.Slider Slider1 
               Height          =   210
               Left            =   120
               TabIndex        =   14
               Top             =   300
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   370
               _Version        =   327682
               LargeChange     =   10
               SmallChange     =   5
               Min             =   10
               Max             =   300
               SelStart        =   10
               TickFrequency   =   10
               Value           =   10
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Groups of:"
               Height          =   225
               Left            =   210
               TabIndex        =   16
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "rows every time"
               Height          =   225
               Left            =   1950
               TabIndex        =   15
               Top             =   660
               Width           =   1275
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5340
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   11
            Top             =   60
            Width           =   7155
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5325
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   10
            Top             =   90
            Width           =   7155
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000A&
            Caption         =   "TimeOut"
            Height          =   1035
            Left            =   -69750
            TabIndex        =   7
            Top             =   150
            Width           =   1965
            Begin VB.TextBox Text4 
               Height          =   330
               Left            =   1260
               MaxLength       =   3
               TabIndex        =   8
               Text            =   "10"
               Top             =   330
               Width           =   615
            End
            Begin ComctlLib.Slider Slider2 
               Height          =   210
               Left            =   60
               TabIndex        =   17
               Top             =   720
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   370
               _Version        =   327682
               LargeChange     =   10
               SmallChange     =   5
               Min             =   10
               Max             =   60
               SelStart        =   10
               TickFrequency   =   10
               Value           =   10
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cancel on:"
               Height          =   225
               Left            =   120
               TabIndex        =   9
               Top             =   390
               Width           =   885
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000A&
            Caption         =   "Own ClipBoard"
            Height          =   4215
            Left            =   -74910
            TabIndex        =   5
            Top             =   1200
            Width           =   7125
            Begin VB.TextBox Text5 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   3885
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   6
               Top             =   240
               Width           =   6915
            End
         End
      End
   End
   Begin VB.PictureBox Pic 
      Align           =   3  'Align Left
      Height          =   5910
      Index           =   0
      Left            =   0
      ScaleHeight     =   5850
      ScaleWidth      =   4065
      TabIndex        =   1
      Top             =   480
      Width           =   4125
      Begin ComctlLib.TreeView Tre 
         Height          =   5835
         Left            =   90
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   10292
         _Version        =   327682
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8430
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":1090
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":1802
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":1F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":26E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":2E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":35CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":3D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":44AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":4C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm.frx":5392
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu o 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu o 
         Caption         =   "Paste"
         Index           =   1
      End
      Begin VB.Menu o 
         Caption         =   "Erase"
         Index           =   2
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu o 
         Caption         =   "Basic Select"
         Index           =   4
      End
      Begin VB.Menu o 
         Caption         =   "Erase last comma"
         Index           =   5
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu o 
         Caption         =   "Insert"
         Index           =   7
      End
      Begin VB.Menu o 
         Caption         =   "Update"
         Index           =   8
      End
      Begin VB.Menu o 
         Caption         =   "Select"
         Index           =   9
      End
      Begin VB.Menu o 
         Caption         =   "Delete"
         Index           =   10
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu o 
         Caption         =   "where"
         Index           =   12
      End
      Begin VB.Menu o 
         Caption         =   "=''"
         Index           =   13
      End
      Begin VB.Menu o 
         Caption         =   "Enter"
         Index           =   14
      End
      Begin VB.Menu o 
         Caption         =   "Order by"
         Index           =   15
      End
      Begin VB.Menu o 
         Caption         =   "Group by"
         Index           =   16
      End
      Begin VB.Menu o 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu o 
         Caption         =   "Insert Columns"
         Index           =   18
      End
      Begin VB.Menu o 
         Caption         =   "Insert one Column"
         Index           =   19
      End
      Begin VB.Menu o 
         Caption         =   "SQL Query to create table"
         Index           =   20
      End
      Begin VB.Menu o 
         Caption         =   "Copy DD to Clipboard"
         Index           =   21
      End
      Begin VB.Menu o 
         Caption         =   "CR on commas (Querys from ClipBoard)"
         Index           =   22
      End
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public OriginalWidth As Integer ' THE DEFAULT WIDTH
Public MaxLength As Double

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'Private Const TVM_SETBKCOLOR = 4381&

Dim Seguir As Boolean

Private Sub Form_Load()
Dim nodX As Node, Contador As Integer, Conta As Integer, Enter As String
Dim hMenu As Long
Idioma
Const SC_SIZE = &HF000
Const MF_BYCOMMAND = &H0
OriginalWidth = 120
MaxLength = Me.Width
'Call SendMessage(Tre.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(192, 192, 255))
Enter = Chr(13) + Chr(10)
Text1.Text = ""
Me.Left = 0
SSTab1.Left = OriginalWidth
Tre.Width = Pic(0).Width - 2 * OriginalWidth
Tre.Left = OriginalWidth
SSTab1.Width = Pic(1).Width - 2 * OriginalWidth
End Sub

Private Sub Form_Activate()
Me.WindowState = 0
End Sub

Private Sub Form_Resize()
MaxLength = Me.Width
End Sub


'RECONECT
Private Sub Recon()
frmIngreso.Show
End Sub

'TO BUFFER
Private Sub ABuffer()
If MsgBox(IIf(Language = 1, "Reemplazar ClipBoard interno: ", "Replace Own CliBoard: ") + Text5.Text, vbYesNo, "QueryBuilder") = vbYes Then
    Text5.Text = Text1.Text
End If
End Sub

Private Sub Variable()
Dim valor As String
valor = Text1.Text
valor = StrTran(valor, Chr(13) + Chr(10), "@@@@@")
valor = StrTran(valor, "@@@@@", " " + Chr(34) + " + _" + Chr(13) + Chr(10) + Chr(34))
valor = "Cadena=" + Chr(34) + valor + Chr(34)
Clipboard.Clear
Clipboard.SetText valor
MsgBox IIf(Language = 1, "Variable creada: ", "Variable created: ") + valor, vbOKOnly, "QueryBuilder"
End Sub

Private Sub AClipBoard()
Select Case SSTab1.Tab
Case 0
    Clipboard.Clear
    Clipboard.SetText Text1.Text
Case 2
    Clipboard.Clear
    Clipboard.SetText Text5.Text
End Select
End Sub

Private Sub Ejecuta()
Dim RowBuf As Variant, Cuantos As Double
Dim RowsReturned As Integer
Dim Mycn As New rdoConnection
Dim qy As New rdoQuery
Dim rs As rdoResultset
Dim i As Integer, J As Integer, Buffer As String
Screen.MousePointer = vbHourglass
Seguir = True
Set Mycn = cn
Set qy = New rdoQuery
qy.Name = "GetRowsQuery"
If Text1.SelLength <> 0 Then
    qy.SQL = Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength + 1)
Else
    qy.SQL = Text1.Text
End If
qy.RowsetSize = 1
Set qy.ActiveConnection = cn
qy.QueryTimeout = Val(Text4.Text)
On Error GoTo MALSQL
Set rs = qy.OpenResultset(rdOpenStatic, rdConcurReadOnly, rdExecDirect)
On Error GoTo 0
Text3.Text = ""
On Error Resume Next
Cuantos = 0
SSTab1.Tab = 1
Do Until rs.EOF
     RowBuf = rs.GetRows(CLng(Text2.Text))
     RowsReturned = UBound(RowBuf, 2) + 1
     For i = 0 To RowsReturned - 1
          Cuantos = Cuantos + 1
          Buffer = ""
          For J = 0 To UBound(RowBuf, 1)
               Buffer = Buffer + CStr(RowBuf(J, i)) & Chr(9)
          Next
          Text3.Text = Text3.Text + Buffer + Chr(13) + Chr(10)
     Next i
     DoEvents
     If Not Seguir Then Exit Do
Loop
Text3.SetFocus
On Error GoTo 0
SALEMALSQL:
Screen.MousePointer = vbDefault
Set rs = Nothing
Set qy = Nothing
Set Mycn = Nothing
Exit Sub
MALSQL:
     MsgBox IIf(Language = 1, "Error en Sentencia! > ", "Error on Query!  >  ") + Err.Description, vbOKOnly, "QueryBuilder"
     GoTo SALEMALSQL
End Sub

Private Sub Detener()
Seguir = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
If MsgBox(IIf(Language = 1, "¿Desea Salir?", "Do you want to Quit?"), vbYesNo, "QueryBuilder") = vbYes Then
    Cancel = 0
Else
    Me.WindowState = 1
End If
End Sub

Private Sub Slider1_Change()
Text2.Text = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
Text2.Text = Slider1.Value
End Sub

Private Sub Slider2_Change()
Text4.Text = Slider2.Value
End Sub

Private Sub Slider2_Scroll()
Text4.Text = Slider2.Value
End Sub

Public Sub CargaTree()
Dim Contador As Integer, Conta As Integer, Buffer, Actual As String
Dim nodX As Node
Dim Arrtipo As Variant
Arrtipo = Array("Char", "Numeric", "DECIMAL", "INTEGER", "SmallInt", "Float", "Real", "Double", "Date", "Time", "TimeStamp", "VarChar", "LONGVARCHAR", "Binary", "VarBinary", "LONGVARBINARY", "BIGINT", "TINYINT", "Bit")
Buffer = UCase(cn.Connect)
If InStr(Buffer, "DATABASE=") <> 0 Then
    Buffer = Mid(Buffer, InStr(Buffer, "DATABASE=") + 9) + ".dbo."
Else
    Buffer = ""
End If
Screen.MousePointer = vbHourglass
Tre.LabelEdit = tvwManual
For Contador = 0 To cn.rdoTables.Count - 1
     Set nodX = Tre.Nodes.Add(, , Chr(64 + cnActual) + CStr(Offset) + CStr(Contador), Buffer + cn.rdoTables(Contador).Name)
     For Conta = 0 To cn.rdoTables(Contador).rdoColumns.Count - 1
        If cn.rdoTables(Contador).rdoColumns(Conta).Type >= 1 Then
            Set nodX = Tre.Nodes.Add(Chr(64 + cnActual) + CStr(Offset) + CStr(Contador), tvwChild, Chr(64 + cnActual) + Chr(64 + cnActual) + CStr(Offset) + CStr(Contador) + CStr(Conta), cn.rdoTables(Contador).rdoColumns(Conta).Name + " > " + Arrtipo(cn.rdoTables(Contador).rdoColumns(Conta).Type - 1) + " (" + CStr(cn.rdoTables(Contador).rdoColumns(Conta).Size) + ")")
        Else
            Set nodX = Tre.Nodes.Add(Chr(64 + cnActual) + CStr(Offset) + CStr(Contador), tvwChild, Chr(64 + cnActual) + Chr(64 + cnActual) + CStr(Offset) + CStr(Contador) + CStr(Conta), cn.rdoTables(Contador).rdoColumns(Conta).Name + " > **********"")")
        End If
     Next
     Offset = Offset + 1
Next
cnActual = cnActual + 1
Screen.MousePointer = vbDefault
End Sub

Private Sub DeClipBoard()
If MsgBox(IIf(Language = 1, "Desea recuperar del ClipBoard: ", "Do you want to retrieve from Clipboard : ") + Clipboard.GetText, vbYesNo, "QueryBuilder") = vbYes Then
    Select Case SSTab1.Tab
    Case 0
        Text1.Text = Text1.Text + Clipboard.GetText
    Case 2
        Text5.Text = Clipboard.GetText
    End Select
End If
End Sub

Private Sub DeBuffer()
If MsgBox(IIf(Language = 1, "Desea recuoperar del ClipBoard interno: ", "Do you want to retrieve from your own ClipBoard : ") + Text5.Text, vbYesNo, "QueryBuilder") = vbYes Then
    Text1.Text = Text5.Text
End If
End Sub

Private Sub Text1_DblClick()
Dim Desde As Integer
If Text1.SelLength > 2 Then
    Clipboard.Clear
    Clipboard.SetText Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength)
Else
    Text1.Text = Mid(Text1.Text, 1, Text1.SelStart) + Clipboard.GetText + Mid(Text1.Text, Text1.SelStart + 1)
End If
End Sub



Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 18 Then
    SSTab1.Tab = 0
End If
End Sub

'BUTTONS
Private Sub Too_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Val(Button.Tag)
Case 1
    Recon
Case 2
    Ejecuta
Case 3
    Detener
Case 4
    AClipBoard
Case 5
    DeClipBoard
Case 6
    ABuffer
Case 7
    DeBuffer
Case 8
    Variable
Case 9
    Graba
Case 10
    Recupera
End Select
End Sub

Private Sub Graba()
Dim Arr As Variant, Contador As Integer
Com.Filter = "MiSQL (*.TS)|*.TS"
Com.DialogTitle = IIf(Language = 1, "GRABAR COMO...", "SAVE AS...")
Com.ShowSave
If Com.filename <> "" Then
    Open (Com.filename) For Output As #1
    Print #1, Text1.Text + Chr(13) + Chr(10)
    Close #1
    Text5.Text = Text1.Text
    Text1.Text = ""
    Arr = tsleeini(Com.filename)
    For Contador = 0 To UBound(Arr, 1) - 1
        Text1.Text = Text1.Text + CStr(Arr(Contador)) + Chr(13) + Chr(10)
    Next
    MsgBox IIf(Language = 1, "Sentencia grabada y puesta en ClipBoard interno.", "Query saved and stored on own ClipBoard"), vbInformation, "ATENCION"
    Beep
End If
End Sub

Private Sub Recupera()
Dim Arr As Variant, Contador As Integer
Com.Filter = "MiSQL (*.TS)|*.TS"
Com.DialogTitle = IIf(Language = 1, "CARGAR SENTENCIA", "LOAD QUERY")
Com.ShowOpen
If File(Com.filename) Then
    Text1.Text = ""
    Arr = tsleeini(Com.filename)
    For Contador = 0 To UBound(Arr, 1) - 1
        Text1.Text = Text1.Text + CStr(Arr(Contador)) + Chr(13) + Chr(10)
    Next
    SSTab1.Tab = 0
    Beep
End If
End Sub

Private Sub o_Click(Index As Integer)
Dim Buffer As String, Veces As Integer, Contador As Integer, Buffer2 As String
Dim Parte1 As String, Parte2 As String, Posicion As Integer, Enter As String
Enter = Chr(13) + Chr(10)
Select Case Index
Case 0
    Clipboard.Clear
    Clipboard.SetText IIf(Text1.SelLength > 0, Mid(Text1.Text, Text1.SelStart + 1, Text1.SelLength), Text1.Text)
Case 1
    Text1.Text = Text1.Text + Enter + Clipboard.GetText
Case 2
    If MsgBox(IIf(Language = 1, "¿Desea borrar el texto?", "Do you want to erase the text?"), vbYesNo, "QueryBuilder") = vbYes Then
        If Text1.SelLength > 0 Then
            Text1.SetFocus
            SendKeys "{DEL}", True
        Else
            Text1.Text = ""
        End If
    End If
Case 4
    Text1.Text = "select" + Enter + Enter + "from " + Enter + Enter + "where " + Enter + Enter + "order by " + Enter + Enter + "group by " + Enter
Case 5
    Text1.SetFocus
    Posicion = Text1.SelStart
    Parte1 = Mid(Text1.Text, 1, Text1.SelStart)
    Parte2 = Mid(Text1.Text, Text1.SelStart + 1)
    For Contador = Len(Parte1) To 1 Step -1
        If Mid(Parte1, Contador, 1) = "," Then
            Parte1 = Mid(Parte1, 1, Contador - 1)
            Exit For
        End If
    Next
    Text1.Text = Parte1 + IIf(Parte2 = Enter, "", Parte2)
    If Posicion > 1 Then
        Text1.SelStart = Posicion - 1
    End If
Case 7
    If Tre.SelectedItem.Children > 0 Then
        Buffer = "insert into " + StrTran(Tre.SelectedItem.FullPath, "\", ".") + "(" + Enter
        Buffer2 = ""
        For Contador = Tre.SelectedItem.Index + 1 To Tre.Nodes.Count - 1
            If Tre.Nodes(Contador).Children > 0 Then
                Exit For
            End If
            Buffer = Buffer + X(StrTran(Tre.Nodes(Contador).Text, "\", ".")) + "," + Enter
            Buffer2 = Buffer2 + "'" + Chr(34) + "+ " + X(StrTran(Tre.Nodes(Contador).Text, "\", ".")) + " + " + Chr(34) + "', " + Enter
        Next
        Buffer = Mid(Buffer, 1, Len(Buffer) - 4) + Enter + ") values (" + Enter
        Buffer2 = Mid(Buffer2, 1, Len(Buffer2) - 4) + ")" + Chr(34) + " " + Enter
        Text1.Text = Buffer + Buffer2
    End If
Case 8
    If Tre.SelectedItem.Children > 0 Then
        Buffer = "update " + StrTran(Tre.SelectedItem.FullPath, "\", ".") + " set (" + Enter
        For Contador = Tre.SelectedItem.Index + 1 To Tre.Nodes.Count - 1
            If Tre.Nodes(Contador).Children > 0 Then
                Exit For
            End If
            Buffer = Buffer + X(StrTran(Tre.Nodes(Contador).Text, "\", ".")) + "='" + Chr(34) + "+" + Trim(X(StrTran(Tre.Nodes(Contador).Text, "\", "."))) + "+" + Chr(34) + "', " + Enter
        Next
        Buffer = Mid(Buffer, 1, Len(Buffer) - 4) + Enter + ")" + Enter
        Text1.Text = Buffer
    End If
Case 9
    If Tre.SelectedItem.Children > 0 Then
        Buffer = "select " + Enter
        For Contador = Tre.SelectedItem.Index + 1 To Tre.Nodes.Count - 1
            If Tre.Nodes(Contador).Children > 0 Then
                Exit For
            End If
            Buffer = Buffer + X(StrTran(Tre.Nodes(Contador).FullPath, "\", ".")) + ", " + Enter
        Next
        Buffer = Mid(Buffer, 1, Len(Buffer) - 4) + Enter
        Buffer = Buffer + "from " + StrTran(Tre.SelectedItem.FullPath, "\", ".")
        Text1.Text = Buffer
    End If
Case 10
    If Not Tre.SelectedItem.Child Is Nothing Then
        Text1.Text = "delete from " + StrTran(Tre.SelectedItem.FullPath, "\", ".") + Chr(13) + Chr(10) + "where    =''"
    End If
Case 12
    Text1.SetFocus
    SendKeys "where " + Chr(vbKeyReturn), True
Case 13
    Text1.SetFocus
    SendKeys "=''{LEFT}", True
Case 14
    Text1.SetFocus
    SendKeys "{ENTER}", True
Case 15
    Text1.SetFocus
    SendKeys "Order by " + Chr(13) + Chr(10), True
Case 16
    Text1.SetFocus
    SendKeys "Group by " + Chr(13) + Chr(10), True
Case 18
    If Tre.SelectedItem.Children = 0 Then
        Text1.SetFocus
        For Contador = Tre.SelectedItem.Index + 1 To Tre.Nodes.Count - 1
            If Tre.Nodes(Contador).Children > 0 Then
                Exit For
            End If
            SendKeys X(StrTran(Tre.Nodes(Contador).Text, "\", ".")) + ",", True
        Next
    End If
Case 19
    Text1.SetFocus
    SendKeys X(StrTran(Tre.SelectedItem.FullPath, "\", ".")) + ", " + Chr(vbKeyReturn), True
    Tre.SetFocus
    DoEvents
    SendKeys "{DOWN}", True
Case 20
    
    If Tre.SelectedItem.Children > 0 Then
        Buffer = "create table " + StrTran(Tre.SelectedItem.FullPath, "\", ".") + " (" + Enter
        For Contador = Tre.SelectedItem.Index + 1 To Tre.Nodes.Count - 1
            If Tre.Nodes(Contador).Children > 0 Then
                Exit For
            End If
            If InStr(Tre.Nodes(Contador).FullPath, " Numeric ") <> 0 Or InStr(Tre.Nodes(Contador).FullPath, " Double ") <> 0 Then
                Buffer = Buffer + StrTran(StrTran(Mid(Tre.Nodes(Contador).Text, 1, Len(Tre.Nodes(Contador).Text) - 1), "\", "."), ">", "") + ",2) NOT NULL, " + Enter
            Else
                Buffer = Buffer + StrTran(StrTran(Tre.Nodes(Contador).Text, "\", "."), ">", "") + " NOT NULL, " + Enter
            End If
        Next
        Buffer = Mid(Buffer, 1, Len(Buffer) - 4) + Enter + ")" + Enter
        Text1.Text = Buffer
    End If
Case 21
    Buffer = ""
    DoEvents
    For Contador = 1 To Tre.Nodes.Count - 1
        If InStr(Tre.Nodes.Item(Contador), " Numeric ") <> 0 Or InStr(Tre.Nodes.Item(Contador), " Double ") <> 0 Then
            Buffer = Buffer + Mid(Tre.Nodes.Item(Contador), 1, Len(Tre.Nodes.Item(Contador)) - 1) + Chr(13) + Chr(10)
        Else
            If Tre.Nodes.Item(Contador).Children > 0 Then
                Buffer = Buffer + Chr(13) + Chr(10)
            End If
            Buffer = Buffer + Tre.Nodes.Item(Contador) + Chr(13) + Chr(10)
        End If
    Next
    Buffer = StrTran(Buffer, " > ", " ")
    Buffer = StrTran(Buffer, "(", "")
    Buffer = StrTran(Buffer, ")", "")
    Buffer = StrTran(Buffer, " ", Chr(9))
    Clipboard.Clear
    Clipboard.SetText Buffer
Case 22
    Buffer = Text1.Text
    Text1.Text = ""
    For Contador = 1 To Len(Buffer) - 1
        If Mid(Buffer, Contador, 1) = "," Then
            Text1.Text = Text1.Text + "," + Enter
        Else
            Text1.Text = Text1.Text + Mid(Buffer, Contador, 1)
        End If
    Next
End Select
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
Source.Left = SSTab1.Left + Text1.Left + X
Tre.Width = Source.Left + Source.Width
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 5 Then
    Ejecuta
End If
End Sub

Private Function X(Cadena As String) As String
If InStr(Cadena, ">") <> 0 Then
    X = Mid(Cadena, 1, InStr(Cadena, ">") - 1)
Else
    X = Cadena
End If
End Function

'IF OPENED, CLOSE IT. IF CLOSED, MAXIMIZE IT
Private Sub Pic_DblClick(Index As Integer)
Dim Cont As Integer, LongTotal As Double
LongTotal = 0 'AMOUNT OF USED WIDTH
If Pic(Index).Width = OriginalWidth Then 'IF CLOSED...
    OpenDoor Index, False
Else 'IF OPENED...
    CloseDoor Index
End If
End Sub

'RESIZE THE DOORS OF TEH CLOSET
Private Sub Pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cont As Integer, LongTotal As Double
LongTotal = 0 'AMOUNT OF USED WIDTH
For Cont = 0 To Pic.Count - 1
    LongTotal = LongTotal + IIf(Index = Cont, X, Pic(Cont).Width) '+ OriginalWidth
Next
LongTotal = LongTotal + OriginalWidth
If LongTotal < MaxLength Or X < Pic(Index).Width Then 'IF IT SUPPORTS A WIDE DOOR OR IF YOU ARE CLOSENING THE DOOR
    If Button = vbLeftButton Then
        If Pic(Index).Width - X <= 5 * OriginalWidth Then 'IF YOU ARE DRAGING ON THE RIGHT SIDE OF THE DOOR
            Pic(Index).Width = IIf(X > OriginalWidth, X, OriginalWidth)
            DoEvents
        End If
    End If
End If
If Index = 0 Then
    If Pic(0).Width > 2 * OriginalWidth Then
        Tre.Width = Pic(0).Width - 2 * OriginalWidth
    End If
Else
    If Pic(1).Width > 2 * OriginalWidth Then
        SSTab1.Width = Pic(1).Width - 2 * OriginalWidth
    End If
End If
End Sub

'CLOSE THE DOR
Private Sub CloseDoor(Index As Integer)
Pic(Index).Width = OriginalWidth
End Sub

'OPEN A DOOR
Private Sub OpenDoor(Index As Integer, CloseRest As Boolean)
Dim Cont As Integer, LongTotal As Double
If CloseRest Then
    For Cont = 0 To Pic.Count - 1
        Pic(Cont).Width = OriginalWidth 'CLOSE ALL THE DOORS
    Next
End If
For Cont = 0 To Pic.Count - 1
    LongTotal = LongTotal + Pic(Cont).Width
Next
LongTotal = LongTotal + OriginalWidth
Pic(Index).Width = MaxLength - LongTotal 'THE MAX PERMITED WIDTH
End Sub

Private Sub OpenAll()
Dim Cont As Integer
For Cont = 0 To Pic.Count - 1
    Pic(Cont).Width = MaxLength / (Pic.Count)
Next
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
End If
If Button = vbRightButton Then
   LockWindowUpdate Text1.hWnd
   Text1.Enabled = False
   DoEvents
   Text1.Enabled = True
   PopupMenu Me.Opciones
   LockWindowUpdate 0&
End If
End Sub

Private Sub Tre_KeyPress(KeyAscii As Integer)
Dim Contador As Integer
Select Case KeyAscii
Case 32
    Text1.SetFocus
    SendKeys X(StrTran(Tre.SelectedItem.FullPath, "\", ".")) + "," + Chr(13), True
    Tre.SetFocus
    SendKeys "{DOWN}", True
End Select
End Sub

Private Sub Idioma()
If Language = 1 Then
    SSTab1.TabCaption(0) = "Sentencia"
    SSTab1.TabCaption(1) = "Resultado"
    SSTab1.TabCaption(2) = "Configuracion"
    Too.Buttons(1).ToolTipText = "Reconectar a una BD"
    Too.Buttons(3).ToolTipText = "Ejecutar sentencia SQL"
    Too.Buttons(4).ToolTipText = "Cancelar sentencia SQL en ejecucion"
    Too.Buttons(6).ToolTipText = "Pone sentencia en el ClipBoard"
    Too.Buttons(7).ToolTipText = "Carga de ClipBoard sentencia SQL"
    Too.Buttons(8).ToolTipText = "Pone sentencia en el ClipBoard del programa"
    Too.Buttons(9).ToolTipText = "Carga de ClipBoard del programa sentencia SQL"
    Too.Buttons(11).ToolTipText = "Crea variable con cadena en ClipBoard para VB"
    Too.Buttons(13).ToolTipText = "Graba sentencia en archivo"
    Too.Buttons(14).ToolTipText = "Recupera sentencia de archivo"
    Frame1.Caption = "Recuperacion de Informacion"
    Label2.Caption = "Grupos de"
    Label4.Caption = "por vez"
    Frame3.Caption = "ClipBoard Interno"
    Frame2.Caption = "Cancelar"
    Label3.Caption = "Cancelar en"
    o(0).Caption = "Copiar"
    o(1).Caption = "Pegar"
    o(2).Caption = "Eliminar"
    o(4).Caption = "SQL Base"
    o(5).Caption = "Quitar ultima coma"
    o(18).Caption = "Insertar campos"
    o(19).Caption = "Insertar un campo"
    o(20).Caption = "Sentencia SQL para crear tabla"
    o(21).Caption = "Poner DD en el ClipBoard"
    o(22).Caption = "Separar comas (Sentencias en ClipBoard)"
End If
End Sub
