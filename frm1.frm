VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONNECT TO"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Subscribe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5970
      TabIndex        =   2
      ToolTipText     =   "Utilize este boton para crear una nueva conexion a la lista"
      Top             =   3180
      Width           =   915
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "(Dbl Click load chain already subscribed) Write the connect chain to your BD"
      Top             =   3180
      Width           =   5685
   End
   Begin MSComDlg.CommonDialog cmm 
      Left            =   2580
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   90
      TabIndex        =   3
      ToolTipText     =   "Enter to begin"
      Top             =   3600
      Width           =   6825
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      ItemData        =   "frm1.frx":08CA
      Left            =   90
      List            =   "frm1.frx":08CC
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Press Del to remove the ODBC chain"
      Top             =   90
      Width           =   6795
   End
End
Attribute VB_Name = "frmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Arrini As Variant

'INITIAL LOAD
Private Sub Form_Load()
Language = 1 'SPANISH
Idioma
cnActual = 0
Set cn = Nothing
CargaRegistro
End Sub

'NEW ODBC STRING
Private Sub Command1_Click()
CreaRegistro
CargaRegistro
End Sub

'TRY TO CONNECT
Private Sub cmd_Click(Index As Integer)
Dim Contador As Integer, OK As Boolean
On Error GoTo HELL
OK = False
Screen.MousePointer = vbHourglass
For Contador = 0 To lst.ListCount - 1
    If lst.Selected(Contador) Then
        Me.Caption = IIf(Language = 1, "Conectando a : ", "Tryon : ") + lst.List(Contador)
        DoEvents
        OK = True
        Set cn = Nothing
        Set cn = New rdoConnection
        cn.CursorDriver = rdUseOdbc
        cn.Connect = lst.List(Contador)
        cn.EstablishConnection
        CreaRegistro
        frmQuery.CargaTree
    End If
Next
If Not OK Then
    MsgBox IIf(Language = 1, "Debe seleccionar alguna cadena de conexion!", "You must define some connection arragnment!"), vbOKOnly, "QueryBuilder"
Else
    Unload Me
    frmQuery.Show
End If
SIGUE:
On Error GoTo 0
Screen.MousePointer = vbDefault
Exit Sub
HELL:
    MsgBox IIf(Language = 1, "No se puede conectar a BD > ", "Can't connect to the BD > ") + Err.Description, vbOKOnly, "QueryBuilder"
    GoTo SIGUE
End Sub

'CREATE NEW ODBC CHAIN
Private Sub CreaRegistro()
Dim Contador As Integer, Actual As String, Existe As Boolean
Actual = txt.Text
Existe = False
For Contador = 0 To lst.ListCount
    If Actual = lst.List(Contador) Then
        Existe = True
    End If
Next
If Not Existe Then
    tsgraini App.Path + "\sele.ini", "f" + Format(Now, "yyyymmddhhmmss"), Actual
End If
End Sub

'CHARGE ODBC CHAINS
Private Sub CargaRegistro()
Dim Contador As Integer, Buffer As String
Arrini = tsleeini(App.Path + "\sele.ini")
lst.Clear
For Contador = 0 To UBound(Arrini, 1)
    Buffer = CStr(Arrini(Contador))
    Buffer = Mid(Buffer, InStr(Buffer, "=") + 1)
    If Len(Trim(Buffer)) > 0 Then
        lst.AddItem Buffer
    End If
Next
lst.ListIndex = 0
End Sub

'DELETE CHAIN
Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Contador As Integer, Parte2 As String, CU As Variant
If KeyCode = 46 Then
    For Contador = 0 To UBound(Arrini, 1)
        CU = tsstrarr(CStr(Arrini(Contador)), "=")
        Parte2 = Mid(CStr(Arrini(Contador)), InStr(CStr(Arrini(Contador)), "=") + 1)
        If Parte2 = lst.Text Then
            Arrini = tsgraini(App.Path + "\sele.ini", CStr(CU(0)), "")
            Exit For
        End If
    Next
    Arrini = tsleeini(App.Path + "\sele.ini")
    CargaRegistro
    lst.SetFocus
End If
End Sub

'SET ON TEXT THE ACTUAL STRING
Private Sub txt_DblClick()
If Len(Trim(txt.Text)) = 0 Then
    txt.Text = lst.Text
End If
End Sub

'SELECT LANGUAGE
Private Sub Idioma()
If Language = 1 Then 'SPANISH
    Me.Caption = "Conectarse a:"
    Command1.Caption = "Matricular"
    lst.ToolTipText = "Delete elimina el registro marcado. Doble Click pone el registro marcado para ediciony matricula"
    Command1.ToolTipText = "Anade una nueva cadena ODBC"
    cmd(0).ToolTipText = "Presione Enter para ingresar a la BD"
    txt.ToolTipText = "Defina la cadena de conexion al motor de BD"
End If
End Sub
