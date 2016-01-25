VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClasifica 
   Caption         =   "Generacion archivo clasificacion"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmClasifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameImportar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   4500
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   690
         Width           =   6735
      End
      Begin VB.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   1230
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   450
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   900
         Picture         =   "frmClasifica.frx":1782
         Top             =   420
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7260
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEscribir 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin ComctlLib.ProgressBar Pb1 
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1170
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   690
         Width           =   7155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fichero generado"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1530
         Picture         =   "frmClasifica.frx":1884
         Top             =   450
         Width           =   240
      End
   End
   Begin VB.Frame FrameConfig 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2790
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Text            =   "Text8"
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   2790
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   2790
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Máximo de Calidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1710
         Width           =   2145
      End
      Begin VB.Label Label7 
         Caption         =   "CLASIFICACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   1350
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmClasifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim SQL As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FecAlbaran As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Cajones As String
Dim Calidad(20) As String

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1



Private Sub cmdConfig_Click(Index As Integer)
Dim I As Integer

    If Index = 1 Then
        Unload Me
    Else
        SQL = ""
        For I = 0 To Text8.Count - 1
            If Text8(I).Text = "" Then SQL = SQL & "Campo: " & I & vbCrLf
        Next I
        If SQL <> "" Then
            SQL = "No pueden haber campos vacios: " & vbCrLf & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Text8(0).SetFocus
            Exit Sub
        End If
        
        mConfig.MaxCalidades = Text8(0).Text
        mConfig.SERVER = Text8(1).Text
        mConfig.User = Text8(2).Text
        mConfig.password = Text8(3).Text
        
        mConfig.Guardar
        
        vConfiguracion False
'        If varConfig.Grabar = 0 Then End
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        
    If Text2.Text <> "" Then
        If Dir(Text2.Text) <> "" Then
            MsgBox "Fichero ya existe", vbExclamation
            Exit Sub
        Else
            FileCopy App.Path & "\" & mConfig.Plantilla, Text2.Text
            NombreHoja = Text2.Text
        End If
    End If
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Si queremos que se vea descomentamos  esto
        MiXL.Application.visible = False
'        MiXL.Parent.Windows(1).Visible = False
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            Screen.MousePointer = vbHourglass
            
            'Vamos linea a linea
            Mens = "Error insertando en Excel"
            
            If EsImportaci = 3 Then
                If Not RecorremosLineasInformes(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
            Else ' esimportaci = 2
                If Not RecorremosLineas(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
            End If
            
            Screen.MousePointer = vbDefault
            
        End If
    
        'Cerramos el excel
        CerrarExcel
                
        MsgBox "Proceso finalizado", vbExclamation


    End If
    
    
End Sub

Private Sub Command2_Click()
Dim Rc As Byte
Dim I As Integer
Dim Rs1 As ADODB.Recordset
Dim KilosI As Long
Dim b As Boolean
Dim Notas As String

    'IMPORTAR
    If Text5.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
    If Dir(Text5.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    
    NombreHoja = Text5.Text
    'Abrimos excel
    
    If EsImportaci = 5 Then
        Screen.MousePointer = vbHourglass
        
        Label1(0).visible = True
        
        If AbrirConexion(BaseDatos) Then
            Conn.BeginTrans
        
            For I = 1 To 20
                ' abrimos la paginas de clasificacion
                If AbrirEXCELPag(I) = 0 Then
                    Exit For
                End If
            
                Label1(0).Caption = "Procesando Solapa " & I
                DoEvents
            
                If I = 1 Then
                    Notas = ExistenNotas
                    If Notas <> "" Then
                        MsgBox "No existen las siguientes notas. Revise." & vbCrLf & vbCrLf & Mid(Notas, 1, Len(Notas) - 2), vbExclamation
                        Command1_Click (1)
                        CerrarExcel
                        Conn.RollbackTrans
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                
                Else
                    Notas = ExistenNotas
                    If Notas <> "" Then
                        MsgBox "No existen las siguientes notas en la solapa " & I & ". Revise." & vbCrLf & vbCrLf & Mid(Notas, 1, Len(Notas) - 2), vbExclamation
                        Command1_Click (1)
                        CerrarExcel
                        Conn.RollbackTrans
                        Screen.MousePointer = vbDefault
                        Label1(0).visible = False
                        DoEvents
                        Exit Sub
                    End If
                    
                    b = CargarNombresCalidades(I)
                    If b Then
                        b = CargarClasificacion(I)
                        If Not b Then
                            Command1_Click (1)
                            CerrarExcel
                            Conn.RollbackTrans
                            Screen.MousePointer = vbDefault
                            Label1(0).visible = False
                            DoEvents
                            Exit Sub
                        End If
                    Else
                        Command1_Click (1)
                        CerrarExcel
                        Conn.RollbackTrans
                        Screen.MousePointer = vbDefault
                    Exit Sub
                    End If
                End If
                CerrarExcel
            Next I
            
            Label1(0).visible = False
            DoEvents
            
            If b Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                Command1_Click (1)
                Conn.CommitTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            
        End If
    
    
    Else
    
        Rc = AbrirEXCEL
        
        If Rc = 0 Then
        
            If EsImportaci = 4 Then
                If AbrirConexion(BaseDatos) Then
                
                    
                    'Vamos linea a linea, buscamos su trabajador
                    RecorremosLineasClasificacion
                    
                End If
            
                'Cerramos el excel
                CerrarExcel
            
            Else
        
                'Realizamos todos los datos
                'abrimos conexion
                If AbrirConexion(BaseDatos) Then
                
                    
                    'Vamos linea a linea, buscamos su trabajador
                    RecorremosLineasLiquidacion
                    
                End If
            
                'Cerramos el excel
                CerrarExcel
              
        
        
                Dim RS As ADODB.Recordset
                Dim C As Long
                Dim cad As String
                SQL = "Select * from tmpexcel WHERE situacion <> 0 and codusu = " & Usuario
        
        
                Set RS = New ADODB.Recordset
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                C = 0
                While Not RS.EOF
                    SQL = SQL & (RS!numalbar) & "        "
                    If (C Mod 6) = 5 Then SQL = SQL & vbCrLf
                    C = C + 1
                    RS.MoveNext
                Wend
                RS.Close
                If C > 0 Then
                    Set frmMens = New frmMensajes
                    
                    frmMens.Cadena = "select * from tmpexcel where situacion <> 0 and codusu = " & Usuario
                    frmMens.OpcionMensaje = 1
                    frmMens.Show vbModal
                    
        '            SQL = "Se han encontrado " & C & " registros con datos incorrectos en la BD: " & vbCrLf & SQL
        '            SQL = SQL & " ¿Desea continuar ?"
        '            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbNo Then Exit Sub
                End If
        
                'Abrimos los registros =0 k son los OK'
                SQL = "¿ Desea importar las clasificaciones correctas ?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        
                SQL = "Select * from tmpexcel WHERE situacion = 0 and codusu = " & Usuario
                
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                C = 0
                While Not RS.EOF
                    C = C + 1
                    
                    SQL = "delete from rhisfruta_clasif where numalbar = " & RS!numalbar
                    Conn.Execute SQL
                    
                    
                    For I = 1 To mConfig.MaxCalidades
                        Select Case I
                            Case 1
                                KilosI = RS!calidad1
                            Case 2
                                KilosI = RS!calidad2
                            Case 3
                                KilosI = RS!calidad3
                            Case 4
                                KilosI = RS!calidad4
                            Case 5
                                KilosI = RS!calidad5
                            Case 6
                                KilosI = RS!calidad6
                            Case 7
                                KilosI = RS!calidad7
                            Case 8
                                KilosI = RS!calidad8
                            Case 9
                                KilosI = RS!calidad9
                            Case 10
                                KilosI = RS!calidad10
                            Case 11
                                KilosI = RS!calidad11
                            Case 12
                                KilosI = RS!calidad12
                            Case 13
                                KilosI = RS!calidad13
                            Case 14
                                KilosI = RS!calidad14
                            Case 15
                                KilosI = RS!calidad15
                            Case 16
                                KilosI = RS!calidad16
                            Case 17
                                KilosI = RS!calidad17
                            Case 18
                                KilosI = RS!calidad18
                            Case 19
                                KilosI = RS!calidad19
                            Case 20
                                KilosI = RS!calidad20
                        End Select
                        
                        If KilosI <> 0 Then
                            SQL = "insert into rhisfruta_clasif (numalbar, codvarie, codcalid, kilosnet) "
                            SQL = SQL & " values (" & RS!numalbar & "," & RS!codvarie & ","
                            SQL = SQL & I & ","
                            SQL = SQL & KilosI & ")"
                        
                            Conn.Execute SQL
                        End If
                            
                    Next I
                    
                    RS.MoveNext
                Wend
                RS.Close
            
            End If
        End If
        
        MsgBox "FIN", vbInformation
        
    End If
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
'    Combo1.ListIndex = Month(Now) - 1
'    Text3.Text = Year(Now)
    FrameEscribir.visible = False
    FrameImportar.visible = False
    Me.FrameConfig.visible = False
    Limpiar
    Select Case EsImportaci
    Case 0
        Caption = "CONFIGURACION"
        FrameConfig.visible = True
'        vConfiguracion True
    Case 1
        Caption = "Cargar Clasificacion desde fichero excel"
        FrameImportar.visible = True
    Case 2
        Caption = "Crear fichero Clasificacion"
        FrameEscribir.visible = True
    Case 3
        Caption = "Crear fichero de Informes a Excel"
        FrameEscribir.visible = True
    Case 4, 5
        Caption = "Cargar Clasificacion desde fichero excel"
        FrameImportar.visible = True
        
    End Select
    
    
 
End Sub

Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub
Private Function TransformaComasPuntos(Cadena) As String
Dim cad As String
Dim J As Integer
    
    J = InStr(1, Cadena, ",")
    If J > 0 Then
        cad = Mid(Cadena, 1, J - 1) & "." & Mid(Cadena, J + 1)
    Else
        cad = Cadena
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
'    Text4.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    AbrirDialogo 0
End Sub

Private Sub Image2_Click()
    AbrirDialogo 1
End Sub


Private Sub AbrirDialogo(Opcion As Integer)

    On Error GoTo EA
    
    With Me.CommonDialog1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        .Filter = "EXCEL (*.xls)|*.xls"
        .CancelError = True
        If Opcion <> 1 Then
            .ShowOpen
            If Opcion = 0 Then
                Text2.Text = .FileName
            Else
                Text5.Text = .FileName
            End If
        Else
            .ShowSave
            Text2.Text = .FileName
        End If
        
        
        
    End With
EA:
End Sub

Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function RecorremosLineas(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer

    On Error GoTo eRecorremosLineas

    RecorremosLineas = False


    SQL = "select * from rhisfruta "
    Sql1 = "select count(*) from rhisfruta "
    
    '[Monica] 19/04/2010: añadida la condicion del sql en el fichero condicionsql.txt
    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then
    
        NFile = FreeFile
    
        Open App.Path & "\condicionsql.txt" For Input As #NFile
 
        If Not EOF(NFile) Then
            Line Input #NFile, Lin
    
            SQL = SQL & " where numalbar in (" & Lin & ")"
            Sql1 = Sql1 & " where numalbar in (" & Lin & ")"
        End If
    End If
    '[Monica] 19/04/2010
    

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    I = 1
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
    
        ExcelSheet.Cells(I, 1).Value = RT!numalbar ' numero de albaran
        ExcelSheet.Cells(I, 2).Value = Format(RT!fecalbar, "yyyy/mm/dd") ' fecha de albaran
        ExcelSheet.Cells(I, 3).Value = RT!codsocio ' codigo de socio
        
        SQL = "select nomsocio from rsocios where codsocio = " & RT!codsocio
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
            ExcelSheet.Cells(I, 4).Value = RS.Fields(0).Value ' nombre de socio
        Else
            ExcelSheet.Cells(I, 4).Value = "" ' nombre de socio
        End If
        
        Set RS = Nothing
        
        ExcelSheet.Cells(I, 5).Value = RT!codcampo ' codigo de campo
        ExcelSheet.Cells(I, 6).Value = RT!codvarie ' codigo de variedad
        
        SQL = "select nomvarie from variedades where codvarie = " & RT!codvarie
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
            ExcelSheet.Cells(I, 7).Value = RS.Fields(0).Value  ' nombre de variedad
        Else
            ExcelSheet.Cells(I, 7).Value = "" ' nombre de variedad
        End If
        
        Set RS = Nothing
        
        ExcelSheet.Cells(I, 8).Value = RT!TipoEntr ' tipo de entrada
        ExcelSheet.Cells(I, 9).Value = RT!KilosNet ' kilos netos
        
        ' cargamos las calidades
        SQL = "select * from rhisfruta_clasif where numalbar = " & RT!numalbar
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            Calidad = RS!codcalid
            
            ExcelSheet.Cells(I, Calidad + 9).Value = RS!KilosNet ' kilos netos
            
            RS.MoveNext
        Wend
        Set RS = Nothing
        
        ' si no hay kilos de algunas calidades las rellenamos a cero
        For JJ = 10 To 29
            If ExcelSheet.Cells(I, JJ).Value = "" Then ExcelSheet.Cells(I, JJ).Value = 0
        Next JJ
    
'        ExcelSheet.Cells(I, 23).Value = 0
'        ExcelSheet.Cells(I, 24).Value = 0
'
    
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineas = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function


Private Sub Image3_Click()
 AbrirDialogo 2
End Sub


Private Sub Image4_Click()
'    Set frmC = New frmCal
'    frmC.Fecha = Now
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then frmC.Fecha = CDate(Text4.Text)
'    End If
'    frmC.Show vbModal
'    Set frmC = Nothing
End Sub

Private Sub Image5_Click()
    MsgBox "Formato importe:   SOLO el punto decimal: 1.49", vbExclamation
End Sub

'Private Sub Text4_LostFocus()
'    Text4.Text = Trim(Text4.Text)
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then
'            Text4.Text = Format(Text4.Text, "dd/mm/yyyy")
'        Else
'            MsgBox "Fecha incorrecta", vbExclamation
'            Text4.Text = ""
'        End If
'    End If
'End Sub
'
'

'-------------------------------------
Private Function RecorremosLineasLiquidacion()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer

    'Desde la fila donde empieza los trabajadores
    'Hasta k este vacio
    'Iremos insertando en tmpHoras
    ' Con trbajador, importe, 0 , 1 ,2
    '             Existe, No existe, IMPORTE negativo
    '
    
    SQL = "DELETE FROM tmpExcel where codusu = " & Usuario
    Conn.Execute SQL
    FIN = False
    I = 2
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            LineasEnBlanco = 0
            If IsNumeric((ExcelSheet.Cells(I, 1).Value)) Then
                If Val(ExcelSheet.Cells(I, 1).Value) > 0 Then
                        'albaran
                        Albaran = Val(ExcelSheet.Cells(I, 1).Value)
                        
                        'Importe
                        FecAlbaran = Format(ExcelSheet.Cells(I, 2).Value, "yyyy/mm/dd")
                        Socio = ExcelSheet.Cells(I, 3).Value
                        Campo = ExcelSheet.Cells(I, 5).Value
                        Variedad = ExcelSheet.Cells(I, 6).Value
                        TipoEntr = ExcelSheet.Cells(I, 8).Value
                        KilosNet = ExcelSheet.Cells(I, 9).Value
                        
                        
                        For JJ = 1 To 20
                            Calidad(JJ) = Val(ExcelSheet.Cells(I, 9 + JJ).Value)
                        Next JJ
                        
                        'InsertartmpLiquida
                        InsertaTmpExcel
                    
                    End If
            End If
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
               
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function




Private Sub InsertaTmpExcel()
Dim vSql As String
Dim vSql2 As String
Dim RT As ADODB.Recordset
Dim RT1 As ADODB.Recordset
Dim RT2 As ADODB.Recordset
Dim Existe As Boolean
Dim ExisteCalidad As Boolean
Dim ExisteEnTemporal As Boolean
Dim TotalKilos As Long
Dim Cuadra As Boolean
Dim JJ As Integer

    On Error GoTo EInsertaTmpExcel
    
    vSql = "Select * from rhisfruta "
    vSql = vSql & " WHERE numalbar = " & Albaran
    vSql = vSql & " and fecalbar = '" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "'"
    vSql = vSql & " and codsocio = " & Socio
    vSql = vSql & " and codcampo = " & Campo
    vSql = vSql & " and codvarie = " & Variedad
    vSql = vSql & " and tipoentr = " & TipoEntr

    Set RT = New ADODB.Recordset
    RT.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RT.EOF Then
        Existe = False
    Else
        Existe = True
    End If
    
    ' si existe la entrada vemos si podemos actualizarla
    If Existe Then
        ExisteCalidad = True
        
        For JJ = 1 To mConfig.MaxCalidades
            If Calidad(JJ) <> 0 Then  ' solo si hay kilos
'                vSQL = "select * from rhisfruta_clasif where numalbar = " & Albaran
'                vSQL = vSQL & " and codvarie = " & Variedad
'                vSQL = vSQL & " and codcalid = " & JJ
'
'                Set RT1 = New ADODB.Recordset
'                RT1.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RT1.EOF Then
                    vSql2 = "select * from rcalidad where codvarie = " & Variedad
                    vSql2 = vSql2 & " and codcalid = " & JJ
                    
                    Set RT2 = New ADODB.Recordset
                    RT2.Open vSql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RT2.EOF Then
                        ExisteCalidad = False
                        Set RT2 = Nothing
                        Exit For
                    Else
                        ExisteCalidad = True
                        Set RT2 = Nothing
                    End If
'                End If
                
            End If
        Next JJ
    
    
        If ExisteCalidad Then ' comprobamos que la suma de calidades da kilosnetos
            TotalKilos = 0
            For JJ = 1 To 20
                TotalKilos = TotalKilos + Calidad(JJ)
            Next JJ
            If TotalKilos <> RT!KilosNet Then
                Cuadra = False
            Else
                Cuadra = True
            End If
        End If
    
    End If
    
    If Existe And ExisteCalidad And Cuadra Then
        
        ExisteEnTemporal = False
        vSql = "select * from tmpexcel where numalbar = " & Albaran & " and codusu = " & Usuario
    
        Set RT2 = New ADODB.Recordset
        RT2.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        If Not RT2.EOF Then
            ExisteEnTemporal = True
        End If
    
        SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        SQL = SQL & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        SQL = SQL & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17, "
        SQL = SQL & "calidad18, calidad19, calidad20, situacion) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & Albaran & ","
        SQL = SQL & "'" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        SQL = SQL & Socio & ","
        SQL = SQL & Campo & ","
        SQL = SQL & Variedad & ","
        SQL = SQL & TipoEntr & ","
        SQL = SQL & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            SQL = SQL & Calidad(JJ) & ","
        Next JJ
    
        If ExisteEnTemporal Then
            SQL = SQL & "2)"
        Else
            SQL = SQL & "0)"
        End If
        
    Else
        SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        SQL = SQL & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        SQL = SQL & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17,"
        SQL = SQL & "calidad18, calidad19, calidad20, situacion) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & Albaran & ",'"
        SQL = SQL & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        SQL = SQL & Socio & ","
        SQL = SQL & Campo & ","
        SQL = SQL & Variedad & ","
        SQL = SQL & TipoEntr & ","
        SQL = SQL & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            SQL = SQL & Calidad(JJ) & ","
        Next JJ
    
        If Not Existe Then
            SQL = SQL & "1)" ' no existe el albaran
        Else
            If Not ExisteCalidad Then ' no existe la calidad
                SQL = SQL & "11)"
            Else
                SQL = SQL & "12)" ' no cuadran kilos
            End If
        End If
        
    End If
    
    
    If SQL <> "" Then Conn.Execute SQL
        
    RT.Close
    
    Exit Sub
EInsertaTmpExcel:
    MsgBox Err.Description
End Sub



Private Sub vConfiguracion(Leer As Boolean)

'    With varConfig
'        If Leer Then
'            Text8(0).Text = .IniLinNomina
'            Text8(1).Text = .FinLinNominas
'            Text8(2).Text = .ColTrabajadorNom
'            Text8(3).Text = .hc
'            Text8(4).Text = .HPLUS
'            Text8(5).Text = .DIAST
'            Text8(6).Text = .Anticipos
'            Text8(7).Text = .ColTrabajadoresLIQ
'            Text8(8).Text = .ColumnaLiquidacion
'            Text8(9).Text = .FilaLIQ
'            Text8(10).Text = .HN
'        Else
'            .IniLinNomina = Val(Text8(0).Text)
'            .FinLinNominas = Val(Text8(1).Text)
'            .ColTrabajadorNom = Val(Text8(2).Text)
'            .hc = Val(Text8(3).Text)
'            .HPLUS = Val(Text8(4).Text)
'            .DIAST = Val(Text8(5).Text)
'            .Anticipos = Val(Text8(6).Text)
'            .ColTrabajadoresLIQ = Val(Text8(7).Text)
'            .ColumnaLiquidacion = Val(Text8(8).Text)
'            .FilaLIQ = Val(Text8(9).Text)
'            .HN = Val(Text8(10).Text)
'        End If
'    End With
End Sub

Private Sub Text8_GotFocus(Index As Integer)
    With Text8(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text8_LostFocus(Index As Integer)
    With Text8(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        Select Case Index
            Case 0 ' numero de calidades
                If Not IsNumeric(.Text) Then
                    MsgBox "Campo debe ser numérico", vbExclamation
                    .Text = ""
                    .SetFocus
                    Exit Sub
                End If
                .Text = Val(.Text)
            
            Case 2, 3 ' usuario y password deben de estar encriptados
            
            
        End Select
            
            
    End With
End Sub



Private Function RecorremosLineasInformes(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer
Dim AlbaranAnt As Long
Dim Primero As Boolean

    On Error GoTo eRecorremosLineas

    RecorremosLineasInformes = False

    Select Case TipoListado
        Case 1
            SQL = "select tmpinformes.*, variedades.nomvarie from tmpinformes INNER JOIN variedades ON tmpinformes.importe1 = variedades.codvarie where codusu = " & Usuario
            Sql1 = "select count(*) from tmpinformes where codusu = " & Usuario
            
            SQL = SQL & " order by campo1, codigo1, importe1, fecha1, importe2"
    End Select

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Select Case TipoListado
        Case 1 ' informe de comprobacion de venta fruta
            ExcelSheet.Cells(1, 1).Value = "Socio/Cliente"
            ExcelSheet.Cells(1, 2).Value = "Código"
            ExcelSheet.Cells(1, 3).Value = "Nombre"
            ExcelSheet.Cells(1, 4).Value = "Cod.Var"
            ExcelSheet.Cells(1, 5).Value = "Variedad"
            ExcelSheet.Cells(1, 6).Value = "Fecha"
            ExcelSheet.Cells(1, 7).Value = "Albarán"
            ExcelSheet.Cells(1, 8).Value = "Palot"
            ExcelSheet.Cells(1, 9).Value = "Tara Palot"
            ExcelSheet.Cells(1, 10).Value = "Calibre"
            ExcelSheet.Cells(1, 11).Value = "Cajas"
            ExcelSheet.Cells(1, 12).Value = "Palets"
            ExcelSheet.Cells(1, 13).Value = "Peso Neto"
            ExcelSheet.Cells(1, 14).Value = "Tipo Alb"
            
            For I = 1 To 15
                ExcelSheet.Cells(1, I + 14).Value = ""
            Next I
            
    End Select
            
    I = 1
    
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
    
        If RT!campo1 = 0 Then
            ExcelSheet.Cells(I, 1).Value = "Socio" 'tipo
        Else
            ExcelSheet.Cells(I, 1).Value = "Cliente" 'tipo
        End If
        
        ExcelSheet.Cells(I, 2).Value = RT!codigo1 ' codigo de socio o de cliente
        ExcelSheet.Cells(I, 3).Value = RT!nombre1 ' nombre de socio o de cliente
        ExcelSheet.Cells(I, 4).Value = RT!importe1 ' codigo de la variedad
        ExcelSheet.Cells(I, 5).Value = RT!nomvarie ' nombre de la variedad
        ExcelSheet.Cells(I, 6).Value = Format(RT!fecha1, "dd/mm/yyyy") ' fecha de albaran
        ExcelSheet.Cells(I, 7).Value = RT!importe2 ' numero del albaran
        ExcelSheet.Cells(I, 8).Value = RT!importeb1 ' numero de palots
        ExcelSheet.Cells(I, 9).Value = RT!importeb2 ' tara de palots
        ExcelSheet.Cells(I, 10).Value = RT!nombre2 ' calibre
        ExcelSheet.Cells(I, 11).Value = RT!importe3 ' numero de cajas
        ExcelSheet.Cells(I, 12).Value = RT!importe4 ' numero de palets
        ExcelSheet.Cells(I, 13).Value = RT!importe5 ' peso neto
    
        If RT!importeb3 = 0 Then
            ExcelSheet.Cells(I, 14).Value = "Vta.Fruta" 'tipo de albaran
        Else
            ExcelSheet.Cells(I, 14).Value = "Precalibrado" 'tipo de albaran
        End If
    
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineasInformes = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function



Private Function RecorremosLineasClasificacion()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
Dim vSql As String
Dim Cadena As String
Dim Posicion As Integer

    FIN = False
    I = 4 '17
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        '[Monica]31/11/2014: Cambio del fichero
        'If Trim(CStr(ExcelSheet.Cells(I, 2).Value)) <> "" Or Trim(CStr(ExcelSheet.Cells(I, 38).Value)) = "Kilos Entrados" Then
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "Total" Then
            LineasEnBlanco = 0
            
            '[Monica]21/11/2014: antes (I, 2)
            If IsNumeric((ExcelSheet.Cells(I, 1).Value)) Then
                If Val(ExcelSheet.Cells(I, 1).Value) > 0 Then
                        ' Nro.Entrada
                        Albaran = Val(ExcelSheet.Cells(I, 1).Value) '(I, 2)
                        
                        
                        '[Monica]25/01/2016: para el caso de anna si el albaran es de longitud > 7 recortamos a 7
                        If Len(CStr(Albaran)) > 7 Then
                            Albaran = Mid(CStr(Albaran), Len(CStr(Albaran)) - 6, 7)
                        End If
                        
                        
                        ' Fecha entrada
                        FecAlbaran = Format(ExcelSheet.Cells(I, 3).Value, "yyyy-mm-dd") '(I, 11)
                        ' socio
                        Socio = "0"
                        ' Campo siempre 0
                        Campo = "0"
                        ' Variedad
                        Variedad = ExcelSheet.Cells(I, 9).Value '(I, 35)
                        
                        SQL = "select codvarie from variedades where nomvarie = '" & Trim(Variedad) & "'"
                        Set RS = New ADODB.Recordset
                        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Variedad = "0"
                        If Not RS.EOF Then
                            Variedad = RS!codvarie
                        End If
                        Set RS = Nothing
                        
                        ' El socio me viene
                        Cadena = ExcelSheet.Cells(I, 4).Value '(I, 14)
                        Posicion = InStr(InStr(1, Cadena, "-"), Cadena, " ") + 1
                        Cadena = Mid(Cadena, Posicion, 20)
                        Socio = Mid(Cadena, 1, InStr(1, Cadena, " "))
                        
                        If Socio <> "" Then
                            SQL = "select codsocio from rsocios where codsocio = " & Trim(Socio)
                            Set RS = New ADODB.Recordset
                            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            Socio = "0"
                            If Not RS.EOF Then
                                Socio = RS!codsocio
                            End If
                            Set RS = Nothing
                        Else
                            Socio = "0"
                        End If
                        
                        SQL = "select codcampo from rcampos where codsocio = " & Trim(Socio) & " and codvarie = " & Variedad
                        Set RS = New ADODB.Recordset
                        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        Campo = "0"
                        If Not RS.EOF Then
                            Campo = RS!codcampo
                        End If
                        Set RS = Nothing
                        
                        ' Cajones
                        Cajones = ExcelSheet.Cells(I, 7).Value '(I, 25)
                        ' Kilos Netos
                        KilosNet = ExcelSheet.Cells(I, 10).Value '(I, 38)
                        
                        vSql = "insert ignore into rclasifica_imp (numnotac,fechaent,codvarie,codsocio,codcampo,numcajon,kilosnet,procesado) values ("
                        vSql = vSql & Albaran & ",'" & FecAlbaran & "'," & Val(Variedad) & "," & Socio & "," & Campo & "," & Val(Cajones) & "," & Val(KilosNet) & ",0)"
                        
                        Conn.Execute vSql
                        
                        If Trim(CStr(ExcelSheet.Cells(I + 1, 1).Value)) = "" Then
                            ' incrementamos la linea, todo lo demás ya me viene cargado
                            I = I + 1
                            
                            Albaran = Mid(Trim(Albaran), 1, Len(Trim(Albaran)) - 1)  ' quitamos el ultimo digito del nro de albaran
                            
                            ' Variedad
                            Variedad = ExcelSheet.Cells(I, 9).Value '(I, 35)
                            
                            SQL = "select codvarie from variedades where nomvarie = '" & Trim(Variedad) & "'"
                            Set RS = New ADODB.Recordset
                            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            Variedad = "0"
                            If Not RS.EOF Then
                                Variedad = RS!codvarie
                            End If
                            Set RS = Nothing
                            
                            ' El socio me viene de antes
                            If Socio <> "" Then
                                SQL = "select codsocio from rsocios where codsocio = " & Trim(Socio)
                                Set RS = New ADODB.Recordset
                                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                Socio = "0"
                                If Not RS.EOF Then
                                    Socio = RS!codsocio
                                End If
                                Set RS = Nothing
                            Else
                                Socio = "0"
                            End If
                            
                            SQL = "select codcampo from rcampos where codsocio = " & Trim(Socio) & " and codvarie = " & Variedad
                            Set RS = New ADODB.Recordset
                            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            Campo = "0"
                            If Not RS.EOF Then
                                Campo = RS!codcampo
                            End If
                            Set RS = Nothing
                            
                            ' Cajones
                            Cajones = ExcelSheet.Cells(I, 7).Value '(I, 25)
                            ' Kilos Netos
                            KilosNet = ExcelSheet.Cells(I, 10).Value '(I, 38)
                            
                            vSql = "insert ignore into rclasifica_imp (numnotac,fechaent,codvarie,codsocio,codcampo,numcajon,kilosnet,procesado) values ("
                            vSql = vSql & Albaran & ",'" & FecAlbaran & "'," & Val(Variedad) & "," & Socio & "," & Campo & "," & Val(Cajones) & "," & Val(KilosNet) & ",0)"
                            
                            Conn.Execute vSql
                            
                        
                        End If
                    End If
            End If
        Else
'            LineasEnBlanco = LineasEnBlanco + 1
'            If LineasEnBlanco < 30 Then
'               ' FIN = False
'            Else
'                FIN = True
'
'            End If
            FIN = True

        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function



'*******************************************************
'** IMPORTACION DE CLASIFICACION DE EXCEL PARA ALZIRA **
'*******************************************************
Private Function CargarNombresCalidades(Solapa As Integer) As Boolean
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim vSql As String
Dim Variedad As String
Dim NumNotac As String
Dim NomCalid As String

    On Error GoTo eCargarNombresCalidades

    CargarNombresCalidades = False

    NumNotac = Mid(ExcelSheet.Cells(2, 1).Value, InStr(1, ExcelSheet.Cells(2, 1).Value, "-") + 1, Len(ExcelSheet.Cells(2, 1).Value))
    
    If NumNotac = "" Then
        CargarNombresCalidades = True
        Exit Function
    End If
    
    SQL = "select codvarie from rclasifica where numnotac = " & NumNotac
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Variedad = "0"
    If Not RS.EOF Then
        Variedad = RS!codvarie
    End If
    Set RS = Nothing
    
    If Variedad = "0" Then
        SQL = "select codvarie from rhisfruta inner join rhisfruta_entradas on rhisfruta.numalbar = rhisfruta_entradas.numalbar  where rhisfruta_entradas.numnotac = " & NumNotac
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Variedad = "0"
        If Not RS.EOF Then
            Variedad = RS!codvarie
        End If
        Set RS = Nothing
    End If


    For JJ = 1 To 20
        Calidad(JJ) = 0
    Next JJ
    
    FIN = False
    I = 3
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(1, I).Value)) <> "" Then
            
            NomCalid = ExcelSheet.Cells(1, I).Value
            If NomCalid = "M.P." Then
                ' No hacemos nada
            Else
                Sql2 = "select codcalid from rcalidad where codvarie = " & Variedad & " and nomcalid = '" & Trim(NomCalid) & "'"
                Set RS2 = New ADODB.Recordset
                RS2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS2.EOF Then
                    Calidad(I - 2) = RS2!codcalid
                End If
                Set RS2 = Nothing
            End If
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
                
    CargarNombresCalidades = True
    Exit Function
    
eCargarNombresCalidades:
    MsgBox "Error Carga Nombres Calidades de la solapa " & Solapa & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function

Private Function ExistenNotas() As String
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim vSql As String
Dim Variedad As String
Dim NumNotac As String
Dim NomCalid As String
Dim cad As String

    ExistenNotas = ""
    
    FIN = False
    I = 2
    LineasEnBlanco = 0
    cad = ""
    While Not FIN
        NumNotac = Mid(ExcelSheet.Cells(I, 1).Value, InStr(1, ExcelSheet.Cells(I, 1).Value, "-") + 1, Len(ExcelSheet.Cells(I, 1).Value))
        If NumNotac = "" Then
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
            End If
        Else
            SQL = "select count(*) from rclasifica where numnotac = " & NumNotac
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.Fields(0).Value = 0 Then
            
                Sql2 = "select count(*) from rhisfruta_entradas where numnotac = " & NumNotac
                
                Set RS2 = New ADODB.Recordset
                RS2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS2.Fields(0).Value = 0 Then
                    cad = cad & NumNotac & ", "
                End If
                Set RS2 = Nothing
            
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
    ExistenNotas = cad
End Function


Private Function CargarClasificacion(Solapa As Integer) As Boolean
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim RS4 As ADODB.Recordset
Dim SQL As String
Dim Sql4 As String
Dim Sql5 As String
Dim Rs5 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim Sql3 As String
Dim vSql As String
Dim Variedad As String
Dim NumNotac As String
Dim Kilos As Long
Dim KilosNetos As String
Dim KilosClasifica As String

Dim KilosEntr As Long
Dim KilosClas As Long
Dim KilosNetosLinea As Long
Dim KilosT As Long
Dim UltimaCalidad As Integer


    On Error GoTo eCargarClasificacion
    
    CargarClasificacion = False
    
    FIN = False
    I = 2
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            NumNotac = Mid(ExcelSheet.Cells(I, 1).Value, InStr(1, ExcelSheet.Cells(I, 1).Value, "-") + 1, Len(ExcelSheet.Cells(I, 1).Value))
            
            Sql3 = "select count(*) from rclasifica where numnotac = " & NumNotac
            Set Rs3 = New ADODB.Recordset
            Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Rs3.Fields(0).Value <> 0 Then
                SQL = "select codvarie from rclasifica where numnotac = " & NumNotac
                
                Set RS = New ADODB.Recordset
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Variedad = "0"
                If Not RS.EOF Then
                    Variedad = RS!codvarie
                End If
                
                SQL = "select kilosnet from rclasifica where numnotac = " & NumNotac
                
                Set RS = New ADODB.Recordset
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                KilosEntr = "0"
                If Not RS.EOF Then
                    KilosEntr = RS!KilosNet
                End If
                
                KilosClas = 0
                For JJ = 1 To 20
                     If IsNumeric((ExcelSheet.Cells(I, JJ + 2).Value)) Then
                        If Val(ExcelSheet.Cells(I, JJ + 2).Value) > 0 Then
                            KilosClas = KilosClas + Val(ExcelSheet.Cells(I, JJ + 2).Value)
                        End If
                     End If
                Next JJ
                
                KilosT = 0
                For JJ = 1 To 20
                
                    If IsNumeric((ExcelSheet.Cells(I, JJ + 2).Value)) Then
                        If Val(ExcelSheet.Cells(I, JJ + 2).Value) > 0 Then
                            ' Kilos
                            Kilos = Val(ExcelSheet.Cells(I, JJ + 2).Value)
                
                            KilosNetosLinea = Round(KilosEntr * Kilos / KilosClas, 0)
                            KilosT = KilosT + KilosNetosLinea
                
                            ' si existe actualizamos
                            Sql5 = "select count(*) from rclasifica_clasif where numnotac = " & NumNotac & " and codvarie = " & Variedad & " and codcalid = " & Calidad(JJ)
                            Set Rs5 = New ADODB.Recordset
                            Rs5.Open Sql5, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            
                            If Rs5.Fields(0).Value <> 0 Then
                                ' montamos el sql de la actualizacion
                                SQL = "update rclasifica_clasif set muestra = " & Kilos & ", kilosnet = " & KilosNetosLinea
                                SQL = SQL & " where numnotac = " & NumNotac
                                SQL = SQL & " and codvarie = " & Variedad
                                SQL = SQL & " and codcalid = " & Calidad(JJ)
                            Else
                
                                ' Montamos el sql de insercion
                                SQL = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values ("
                                SQL = SQL & NumNotac & "," & Variedad & "," & Calidad(JJ) & "," & Kilos & "," & KilosNetosLinea & ")"
                            End If
                        
                            Conn.Execute SQL
                        
                            UltimaCalidad = Calidad(JJ)
                        End If
                    End If
                Next JJ
                
                If KilosT <> KilosEntr Then
                    SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & (KilosEntr - KilosT)
                    SQL = SQL & " where numnotac = " & NumNotac
                    SQL = SQL & " and codvarie = " & Variedad
                    SQL = SQL & " and codcalid = " & UltimaCalidad
                    
                    Conn.Execute SQL
                End If
                
             Set Rs3 = Nothing
                            
'            ' Una vez metida la clasificacion miramos si la suma de kilos no cuadra con los kilos netos, damos un aviso
'            ' y no hace nada
'            Sql4 = "select if(sum(kilosnet) is null, 0,sum(kilosnet)) from rclasifica_clasif where numnotac = " & NumNotac
'            Set RS4 = New ADODB.Recordset
'            RS4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            KilosClasifica = "0"
'            If Not RS.EOF Then
'                KilosClasifica = RS4.Fields(0).Value
'            End If
'            Set RS4 = Nothing
'
'            Sql4 = "select if(kilosnet is null, 0, kilosnet) from rclasifica where numnotac = " & NumNotac
'            Set RS4 = New ADODB.Recordset
'            RS4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            KilosNetos = "0"
'            If Not RS.EOF Then
'                KilosNetos = RS4.Fields(0).Value
'            End If
'            Set RS4 = Nothing
'
'            If KilosNetos <> KilosClasifica Then
'                MsgBox "Los kilos netos de la nota " & NumNotac & " no coinciden con los clasificados. Revise.", vbExclamation
'                CargarClasificacion = False
'                Exit Function
'            End If
            End If
            
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
            End If
        End If
        I = I + 1
    Wend

    CargarClasificacion = True
    Exit Function
    
eCargarClasificacion:
    MsgBox "Se ha producido un error Cargando Clasificacion de la solapa " & Solapa & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function




