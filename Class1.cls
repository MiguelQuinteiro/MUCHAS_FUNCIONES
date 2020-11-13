VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'****************************************************************************************
'* PROYECTO      : NOMBRE DEL PROYECTO
'* CONTENIDO     : CONTENIDO DEL PROGRAMA
'* VERSION       : # DE LA VERSIÓN
'* AUTORES       : NOMBRE DEL AUTOR O RESPONSABLE
'* INICIO        : FECHA DE CREACIÓN DEL CÓDIGO
'* ACTUALIZACION : FECHA DE LA ÚLTIMA ACTUALIZACIÓN
'****************************************************************************************

Option Explicit
'DacorLIB, versión 3.2  2/2009.
'Declara funciones API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpdefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private SEP_DEC As String
Private iSepara As Integer
Private aOrden() As Integer
Private aMenu() As String
Private SYS_VER As String
Private SYS_PC As String
Private Const TIT_EXCEL = "Exportación a Microsoft Excel"
Private Const DIA_EXCEL = "Recuperando información para el período seleccionado. Espere un momento por favor..."
Private Const LBL_EXCEL = "Enviando información a Microsoft Excel. Espere unos instantes..."
Private Const MSG_EXCEL = "El envío de información se realizó satisfactoriamente."
Private Const MSG_NIL = "No se encontró información para el período especificado."
Private aExcel() As String
Private exTitulos As Boolean
Private SEPDEC As String * 1
Private Const CAR_ESP = ",.-()%$/Ññ!°#$%&?¡@\_áéíóúÁÉÍÓÚ+"
Private Const CAR_RUT = "-kK"
Private Const CAR_FAX = ".-() "
Private Const BOX_NEW = 1
Private Const BOX_ADD = 2
Private Const BOX_MOD = 3
Private Const BOX_DEL = 4
Private Const COLOR_BLANCO = &H80000005
Private Const COLOR_PLOMO = &H8000000F
Private Const COLOR_NEGRO = &H80000007
Private Const COLOR_AMARILLO = &H80000018
Private Const COLOR_ROJO = &HFF&
Private hAviso As Boolean
Private hTexto As String
Private hIcono As Integer
Private hTitulo As String

'   *   **   ***   ****  FORM_LOAD   ****   ***   **   *

'   *   **   ***   ****  BOTONES   ****   ***   **   *

'   *   **   ***   ****  TEXTBOX   ****   ***   **   *

'   *   **   ***   ****  COMBOBOX   ****   ***   **   *

'   *   **   ***   ****  LISTBOX   ****   ***   **   *

'   *   **   ***   ****  LABEL   ****   ***   **   *

'   *   **   ***   ****  PROCEDIMIENTOS   ****   ***   **   *
' .:. Nombre del Procedimiento .:.
Public Sub Excel(ByVal sArchivo As String)
  Dim s As String
  Dim f As Integer
  Dim i As Integer
  Dim j As Integer
  Dim m As Integer
  Dim n As Integer
  s = NombreArchivo(sArchivo)
  If Left(Right(s, 4), 1) <> "." Then
    sArchivo = sArchivo & ".xls"
  Else
    If Right(s, 3) <> "xls" Then
      sArchivo = sArchivo & "xls"
    End If
  End If
  f = FreeFile
  Open sArchivo For Output As #f
  m = UBound(aExcel)
  n = UBound(aExcel, 2)
  For i = 0 To m - 1
    s = ""
    For j = 0 To n
      s = s & aExcel(i, j)
      If j < n Then s = s & vbTab
    Next j
    Print #f, s
  Next i
  Close #f
  s = MSG_EXCEL & vbCrLf & "'" & sArchivo & "'."
  MsgBox s, vbInformation, TIT_EXCEL
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MesEntero(ByRef sDesde As String, ByRef sHasta As String, Optional ByVal iMes As Integer, Optional ByVal iAño As Integer)
  If iMes = 0 Then iMes = Month(Now)
  If iAño = 0 Then iAño = Year(Now)
  iMes = iMes - 1
  If iMes = 0 Then
    iMes = 12
    iAño = iAño - 1
  End If
  sDesde = Dia("01/" & iMes & "/" & iAño)
  sHasta = DateAdd("m", 1, sDesde) - 1
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MesLabel(lblDesde As Object, lblHasta As Object)
  Dim dDia As Date
  Dim mMes As Integer
  Dim mAño As Integer
  mMes = Month(Now)
  mAño = Year(Now)
  If mMes = 12 Then
    mMes = 1
    mAño = mAño + 1
  Else
    mMes = mMes + 1
  End If
  dDia = Dia1(mMes, mAño) - 1
  FechaLabel lblHasta, dDia
  dDia = dDia - Day(dDia) + 1
  FechaLabel lblDesde, dDia
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub GrabaLog(ByVal sPath As String, ByVal sFileLog As String, ByVal sUser As String, ByVal iError As Integer, ByVal sError As String, ByVal sOperacion As String, ByVal sQuery As String, ByVal sTabla As String)
'
'Autor       : Patricio Dahmen G.
'Versión     : 02/2000
'Descripción : Graba archivo LOG con descripción del error ocurrido.
'Parámetros :
'     I        sPath, ruta archivo LOG.
'     I        sFileLog, nombre archivo LOG.
'     I        sUser, identificación usuario.
'     I        iError, número del error ocurrido.
'     I        sError, descripción del error ocurrido.
'     I        sOperacion, operación SQL que falló.
'     I        sQuery, query realizada.
'     I        sTabla, nombre tabla.
'
'Declara variables locales
  Dim f As Integer
  Const EXT = ".txt"
  'Crea path archivo
  If Right$(sPath, 1) <> "\" Then
    sFileLog = sPath & "\" & sFileLog & EXT
  Else
    sFileLog = sPath & sFileLog & EXT
  End If
  'Escribe registro
  f = FreeFile
  Open sFileLog For Append Access Write As #f
  Print #f, CStr(Now) & "; user=" & sUser & "; error=" & sError & " [" & iError & "]; operacion=" & sOperacion & "; sql=" & sQuery & "; tabla=" & sTabla
  Close #f
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Pantalla(iAlto As Integer, iAncho As Integer)
  Const SCR = 2 / 30
  iAlto = Screen.Height * SCR
  iAncho = Screen.Width * SCR
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Private Sub CarpetaVer(ByVal sDir As String)
  Dim sCarpeta As String
  Dim iTam As Integer
  On Error Resume Next
  iTam = Len(sDir)
  sCarpeta = sDir
  If Mid(sCarpeta, iTam, 1) <> "}" Then Exit Sub
  Name sCarpeta As Mid(sCarpeta, 1, iTam - 39)
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Private Sub CarpetaNoVer(ByVal sDir As String)
  Dim sCarpeta As String
  Dim iTam As Integer
  On Error Resume Next
  sCarpeta = sDir
  iTam = Len(sCarpeta)
  If Mid(sCarpeta, iTam, 1) = "}" Then Exit Sub
  If iTam = 3 Then Exit Sub
  Name sCarpeta As sCarpeta & ".{00021401-0000-0000-C000-000000000046}"
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Carpeta(ByVal sDir As String, ByVal bOcultar As Boolean)
  If bOcultar Then
    CarpetaNoVer sDir
  Else
    CarpetaVer sDir
  End If
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub UpdateStatus(pic As Object, ByVal sngPercent As Single, Optional ByVal fBorderCase)
  Dim strPercent As String
  Dim intX As Integer
  Dim intY As Integer
  Dim intWidth As Integer
  Dim intHeight As Integer
  If IsMissing(fBorderCase) Then fBorderCase = False
  Const colBackground = &HFFFFFF  ' blanco
  Const colForeground = &H800000  ' azul oscuro
  pic.ForeColor = colForeground
  pic.BackColor = colBackground
  Dim intPercent
  intPercent = Int(100 * sngPercent + 0.5)
  If intPercent = 0 Then
    If Not fBorderCase Then
      intPercent = 1
    End If
  ElseIf intPercent = 100 Then
    If Not fBorderCase Then
      intPercent = 99
    End If
  End If
  strPercent = Format$(intPercent) & "%"
  intWidth = pic.TextWidth(strPercent)
  intHeight = pic.TextHeight(strPercent)
  intX = pic.Width / 2 - intWidth / 2
  intY = 30  'pic.Height / 2 - intHeight / 2
  pic.DrawMode = 13  ' Copy Pen
  pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
  pic.CurrentX = intX
  pic.CurrentY = intY
  pic.Print strPercent
  pic.DrawMode = 10  ' Not XOR Pen
  If sngPercent > 0 Then
    pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
  Else
    pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
  End If
  pic.Refresh
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub qVista(lvLista As Object)
  With lvLista
    If .ListItems.Count = 0 Then Exit Sub
    .ListItems(1).Selected = True
    .ListItems(1).Selected = False
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Vacio(ByVal sNombreCampo As String)
  MsgBox "Deberá ingresar la información en el campo.", vbCritical, sNombreCampo
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub OrdenDate(lview As Object, ByVal iColumna As Integer, Optional ByVal iColumnaFecha As Integer, Optional ByVal iColumnaFechaOrden As Integer)
  Dim iOrden As Integer
  Dim iAnterior As Integer
  Dim iColOrden As Integer
  If iColumna <> iColumnaFecha Then
    iColumnaFecha = 0
  Else
    iColumnaFecha = iColumnaFechaOrden
  End If
  If iColumnaFecha = 0 Then
    iAnterior = 1
    iColumnaFecha = iColumna
  End If
  iColOrden = iColumnaFecha - iAnterior
  With lview
    iOrden = aOrden(iColumna)
    If iOrden = 0 Then
      iOrden = 1
    Else
      iOrden = 0
    End If
    If .SortKey <> iColOrden Then iOrden = 0
    .SortKey = iColOrden
    .SortOrder = iOrden
    aOrden(iColumna) = iOrden
    .Sorted = True
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MenuBotones(fm As Object, barToolbar As Object)
  Dim h As String
  Dim xMnu As Menu
  Dim b As Integer
  Dim v As Integer
  With barToolbar
    v = .Buttons.Count
    For b = 1 To v
      h = .Buttons(b).Tag
      If Len(h) > 0 Then
        Set xMnu = fm(h)
        xMnu.Enabled = .Buttons(b).Enabled
      End If
    Next b
  End With
  Set xMnu = Nothing
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Menus(ByVal iBotones As Integer, ByVal sNombreMenu As String)
  Dim cl As Integer
  If iBotones > 0 Then
    aMenu(iBotones) = sNombreMenu
    Exit Sub
  End If
  iBotones = Abs(iBotones)
  ReDim aMenu(iBotones)
  For cl = 1 To iBotones
    aMenu(cl) = ""
  Next cl
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub OrdenCol(ByVal iColumnas As Integer)
  Dim cl As Integer
  ReDim aOrden(iColumnas)
  For cl = 1 To iColumnas
    aOrden(cl) = 0
  Next cl
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Orden(lview As Object, ByVal iColumna As Integer)
  Dim iOrden As Integer
  Dim iColOrden As Integer
  iColOrden = iColumna - 1
  With lview
    iOrden = aOrden(iColumna)
    If iOrden = 0 Then
      iOrden = 1
    Else
      iOrden = 0
    End If
    If .SortKey <> iColOrden Then iOrden = 0
    .SortKey = iColOrden
    .SortOrder = iOrden
    aOrden(iColumna) = iOrden
    .Sorted = True
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub FechaLabel(lblFecha As Object, ByVal dFecha As Date)
  lblFecha = UCase(Format(dFecha, "dd/mmm/yyyy"))
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MaxNro(txtNro As Object, ByVal iCol As Integer, ByRef KeyAscii As Integer, Optional ByVal bDecimal As Boolean, Optional ByVal iMax As Integer, Optional ByVal iKol As Integer)
  Dim lmax As Integer
  Dim k As Integer
  Dim q As Integer
  If iCol = 0 Then
    lmax = 5
    If bDecimal Then lmax = 2
  Else
    lmax = 8
  End If
  If iMax = 0 Then iMax = 2
  If iKol > 0 Then lmax = iKol + 2
  q = lmax
  With txtNro
    If bDecimal Then
      k = InStr(1, .Text, SepDecimal, vbTextCompare)
      If Len(.Text) = iMax Then
        If k = 0 Then
          If KeyAscii <> Asc(SepDecimal) Then
            lmax = -1
          End If
        End If
      Else
        If Len(.Text) = 0 Then
          If KeyAscii = Asc(SepDecimal) Then lmax = -1
        Else
          If k > 0 Then lmax = k + iKol
        End If
      End If
    End If
    If Len(.Text) >= lmax Then
      If KeyAscii <> vbKeyBack Then
        If .SelLength <> q Then
          KeyAscii = 0
          Beep
        End If
      End If
    End If
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Oculto(lblEtiq As Object, Optional ByVal bDesarrollo As Boolean)
  With lblEtiq
    .Visible = bDesarrollo
    .ToolTipText = .Name
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Cursor(txtTextBox As Object, Optional ByVal bSeleccionar As Boolean)
  Dim i As Integer
  With txtTextBox
    i = Len(.Text)
    If bSeleccionar Then
      .SelStart = 0
      .SelLength = i
      Exit Sub
    End If
    .SelStart = i
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub InfoArchivo(ByVal sRutayArchivo As String, ByRef sRuta As String, ByRef sArchivo As String)
  Dim l As Integer
  Dim p As Integer
  l = 1
  Do While l > 0
    l = InStr(l, sRutayArchivo, "\", vbTextCompare)
    If l > 0 Then
      p = l
      l = l + 1
    Else
      Exit Do
    End If
  Loop
  sRuta = Mid(sRutayArchivo, 1, p)
  sArchivo = Mid(sRutayArchivo, p + 1)
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SetToolBarData(tbBarra As Object, aBarra() As Variant, imgImageList As Object)
'"1__" : 1=visible, 0=invisible
'"_1_" : 1=activo, 0=inactivo
'"__c" : b=limpiar, g=grabar, x=borrar, z=separador,
'        p=primero, a=anterior, s=siguiente, u=último,
'        d=detalle, i=imprimir
'        h=ayuda, e=salir
  Dim bt As Integer
  Dim sTip As String
  With tbBarra
    .ImageList = imgImageList
    For bt = 1 To UBound(aBarra)
      If Left(aBarra(bt), 1) = 0 Then .Buttons(bt).Visible = False
      If Mid(aBarra(bt), 2, 1) = 0 Then .Buttons(bt).Enabled = False
      .Buttons(bt).Image = bt
      Select Case Right(aBarra(bt), 1)
      Case "r"
        sTip = "Actualizar"
      Case "t"
        sTip = "Listar"
      Case "d"
        sTip = "Detalle registro"
      Case "b"
        sTip = "Nuevo registro"
      Case "g"
        sTip = "Grabar registro"
      Case "x"
        sTip = "Eliminar registro"
      Case "p"
        sTip = "Primer registro"
      Case "a"
        sTip = "Registro anterior"
      Case "s"
        sTip = "Siguiente registro"
      Case "u"
        sTip = "Último registro"
      Case "i"
        sTip = "Imprimir"
      Case "h"
        sTip = "Ayuda"
      Case "e"
        sTip = "Cerrar"
      Case "z"
        .Buttons(bt).Style = 3
      End Select
      .Buttons(bt).ToolTipText = sTip
    Next bt
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub UnformatHora(ByVal sHora As String, iHoras As Integer, iMinutos As Integer)
  If sHora = "" Then sHora = Reloj(0, 0)
  iHoras = Left(sHora, 2)
  iMinutos = Mid(sHora, 4)
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub FondoControles(frmForm As Object)
  Dim cnt As Control
  Dim sName As String
  With frmForm
    For Each cnt In frmForm
      sName = LCase(Left(cnt.Name, 3))
      Select Case sName
      Case "lab", "lbl"
        cnt.BackStyle = 0
        cnt.BackColor = COLOR_AMARILLO
        cnt.ForeColor = COLOR_NEGRO
      Case "txt"
        cnt.BackColor = COLOR_BLANCO
        cnt.ForeColor = COLOR_NEGRO
      Case "fra"
        cnt.BackColor = .BackColor
      Case "chk"
        cnt.BackColor = .BackColor
        cnt.ForeColor = COLOR_BLANCO
      End Select
    Next
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub ConfigINI(ByVal sAppPath As String, ByVal sysApp As String, ByVal sHead As String, ByVal sValue As String)
'declara variables locales
  Dim sItem As String
  Dim sRuta As String
  Dim b As Boolean
  'obtiene items
  sItem = LeeIni(sAppPath, sysApp, sHead, sValue, "nil")
  If Left(sItem, 2) = "__" Then
    sRuta = sAppPath & Mid(sItem, 3)
    b = EscribeIni(sAppPath, sysApp, sHead, sValue, sRuta)
  End If
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub CentrarFecha(frmFrame As Object, lblLabel As Object)
'centra fecha dentro del frame
  With frmFrame
    lblLabel.Move (.Width - lblLabel.Width) \ 2, (.Height - lblLabel.Height) \ 2
    lblLabel.BorderStyle = 0
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SetToolBar(tbBarra As Object, aBarra() As Variant)
'"1__" : 1=visible, 0=invisible
'"_1_" : 1=activo, 0=inactivo
'"__c" : b=limpiar, a=agregar, u=modificar, x=borrar, z=separador,
'        s=listar, i=imprimir, h=ayuda, e=salir
  Dim bt As Integer
  Dim sTip As String
  With tbBarra
    For bt = 1 To UBound(aBarra)
      If Left(aBarra(bt), 1) = 0 Then .Buttons(bt).Visible = False
      If Mid(aBarra(bt), 2, 1) = 0 Then .Buttons(bt).Enabled = False
      .Buttons(bt).Image = bt
      Select Case Right(aBarra(bt), 1)
      Case "b"
        sTip = "Limpiar cuadros"
      Case "a"
        sTip = "Nuevo registro"
      Case "u"
        sTip = "Modificar registro"
      Case "x"
        sTip = "Eliminar registro"
      Case "s"
        sTip = "Listar registros"
      Case "i"
        sTip = "Imprimir"
      Case "h"
        sTip = "Ayuda"
      Case "e"
        sTip = "Cerrar"
      Case "z"
        .Buttons(bt).Style = 3
      End Select
      .Buttons(bt).ToolTipText = sTip
    Next bt
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SetCampos(f As Object)
  Dim c As Integer
  Dim r As Integer
  Dim s As Integer
  c = f.txtFields.Count
  s = iSepara
  If s = 0 Then s = 480
  For r = 0 To c - 1
    If r > 0 Then f.txtFields(r).Top = f.txtFields(r - 1).Top + s
    f.lblLabels(r).Top = f.txtFields(r).Top + 30
    f.lblLabels(r).FontName = "Tahoma"
    f.lblLabels(r).FontSize = 9
    f.lblLabels(r).Caption = Extrae(f.lblLabels(r).Caption, ":")
    f.txtFields(r).TabIndex = r
  Next r
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Mouse(obj As Object, Optional ByVal bReloj As Boolean)
'cambia apariencia cursor
  If bReloj Then
    obj.MousePointer = vbHourglass
  Else
    obj.MousePointer = vbDefault
  End If
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub BotonesMenu(barToolbar As Object, Optional ByVal bInicializar As Boolean)
  Dim b As Integer
  Dim v As Integer
  With barToolbar
    v = .Buttons.Count
    If bInicializar Then
      Menus -v, ""
      Exit Sub
    End If
    For b = 1 To v
      .Buttons(b).Tag = aMenu(b)
    Next b
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub BarraMenu(barToolbar As Object, imlIconos As Object)
  Dim b As Integer
  Dim sTipText As String
  With barToolbar
    .ImageList = imlIconos
    For b = 1 To .Buttons.Count
      If .Buttons(b).Style <> 3 Then
        .Buttons(b).Image = .Buttons(b).Index
        sTipText = .Buttons(b).ToolTipText
        If sTipText = "" Then
          .Buttons(b).ToolTipText = Extrae(.Buttons(b).Caption, "&")
        Else
          .Buttons(b).ToolTipText = sTipText
        End If
      End If
    Next b
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SizeList(lvListView As Object, fraFrame As Object)
  With lvListView
    .Left = 120
    .Top = 240
    .Width = fraFrame.Width - 240
    .Height = fraFrame.Height - 420
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MensajesSQL(ByVal bFemenino As Boolean, ByVal sEntidad As String, ByVal sTitulo As String, Optional ByVal bBorrar As Boolean)
  Dim sTexto As String
  Dim sElLa As String
  Dim sUnUna As String
  Dim bDel As Boolean
  If bFemenino Then
    sElLa = "la "
    sUnUna = "una "
  Else
    sElLa = "el "
    sUnUna = "un "
  End If
  If bBorrar Then
    bDel = False
    sTexto = "¿Está seguro de eliminar "
    sTexto = sTexto & sElLa & sEntidad & "?"
    bDel = (MsgBox(sTexto, vbQuestion + vbYesNo, sTitulo) = vbYes)
    Exit Sub
  End If
  sTexto = "Deberá seleccionar "
  sTexto = sTexto & sUnUna & sEntidad & "."
  MsgBox sTexto, vbExclamation, sTitulo
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Mensaje(ByVal sTexto As String, ByVal iIcono As Integer, ByVal sTitulo As String)
  If sTexto <> "" Then
    hTexto = sTexto
    hIcono = iIcono
    hTitulo = sTitulo
    hAviso = True
    Exit Sub
  End If
  If Not hAviso Then Exit Sub
  MsgBox hTexto, hIcono, hTitulo
  hAviso = False
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub LookClave(detForm As Object, ByVal lID As Long, Optional ByVal iIndex As Integer)
  With detForm
    .campo(iIndex) = lID
    .campo(iIndex).Enabled = False
    .campo(iIndex).BackColor = COLOR_AMARILLO
    .campo(iIndex).FontBold = True
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Etiqueta(lblEtiqueta As Object, ByVal sCaption As String)
  With lblEtiqueta
    If sCaption = "" Then
      .Caption = ""
      Exit Sub
    End If
    Select Case .Alignment
    Case 0
      .Caption = " " & sCaption
    Case 1
      .Caption = sCaption & " "
    Case 2
      .Caption = sCaption
    End Select
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub ComboListIndex(cboCombo As Object, ByVal lValorIndice)
  Dim z As Integer
  'obtiene ítem
  With cboCombo
    For z = 0 To .ListCount - 1
      If lValorIndice = .ItemData(z) Then
        cboCombo.ListIndex = z
        Exit Sub
      End If
    Next z
    .ListIndex = -1
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SQLfile(ByVal sPath As String, ByVal sFileSQL As String, ByVal sQuery As String)
'
'Autor       : Patricio Dahmen G.
'Versión     : 02/2000
'Descripción : Graba archivo LOG con descripción del error ocurrido.
'Parámetros :
'     I        sPath, ruta archivo LOG.
'     I        sFileLog, nombre archivo LOG.
'     I        sUser, identificación usuario.
'     I        iError, número del error ocurrido.
'     I        sError, descripción del error ocurrido.
'     I        sOperacion, operación SQL que falló.
'     I        sQuery, query realizada.
'     I        sTabla, nombre tabla.
'
'Declara variables locales
  Dim f As Integer
  Const EXT = ".txt"
  'Crea path archivo
  If Right$(sPath, 1) <> "\" Then
    sFileSQL = sPath & "\" & sFileSQL & EXT
  Else
    sFileSQL = sPath & sFileSQL & EXT
  End If
  'Escribe registro
  f = FreeFile
  Open sFileSQL For Append Access Write As #f
  Print #f, sQuery
  Close #f
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Esperar(ByVal lSegundos As Long)
'
'Autor       : Patricio Dahmen G.
'Versión     : 03/2000
'Descripción : Detiene la ejecución de la aplicación por un período
'              de n segundos.
'Parámetros  :
'     I        lSegundos, segundos de espera.
'
'Declara variables locales
  Dim dSegIni As Date
  Dim dSegFin As Date
  'Verifica parámetro
  If lSegundos <= 0 Then Exit Sub
  'Cambia apariencia cursor
  Screen.MousePointer = vbHourglass
  'Obtiene segundos iniciales
  dSegIni = CDate(Format(Time, "hh:mm:ss"))
  'Ciclo
  Do While True
    'Obtiene segundos finales
    dSegFin = CDate(Format(Time, "hh:mm:ss"))
    'Compara diferencia
    If DateDiff("s", dSegIni, dSegFin) >= lSegundos Then
      Exit Do
    End If
  Loop
  'Restaura apariencia cursor
  Screen.MousePointer = vbDefault
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub CargarMeses(cboComboBox As Object)
'
'Autor       : Patricio Dahmen G.
'Versión     : 06/1999
'Descripción : Carga combobox con los meses del año.
'Parámetros  :
'     I        cboComboBox, control combobox.
'
'Declara variables locales
  Dim i As Integer
  'Carga meses
  For i = 1 To 12
    cboComboBox.AddItem Format("28/" & CStr(i) & "/1999", "mmmm")
  Next i
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Degradado(frmForm As Object, Optional ByVal bPlomo As Boolean, Optional ByVal iMaxColor As Integer)
'
'Autor       : Patricio Dahmen G.
'Versión     : 08/1999
'Descripción : Obtiene fondo formulario en degradé.
'Parámetros  :
'     I        frmForm, form a tratar.
'     I        bPlomo, color plomo. Opcional.
'     I        iMaxColor, color máximo degradé. Opcional.
'
'Declara variables locales
  Dim iDegree As Integer
  Dim iDegrado As Integer
  With frmForm
    .DrawStyle = vbInsideSolid
    .DrawMode = vbCopyPen
    .ScaleMode = vbPixels
    .DrawWidth = 2
    .ScaleHeight = 256
  End With
  If iMaxColor <= 0 Then iMaxColor = 255
  If iMaxColor > 255 Then iMaxColor = 255
  If bPlomo Then
    iDegrado = 0
    For iDegree = 0 To 255
      If iDegree < iMaxColor Then
        iDegrado = iDegree
      End If
      frmForm.Line (0, iDegree)-(Screen.Width, _
                                 iDegree - 1), RGB(255 - iDegrado, 255 - iDegrado, 255 - iDegree), B
    Next iDegree
    Exit Sub
  End If
  iDegrado = iMaxColor
  For iDegree = 0 To 255
    If iDegree > iMaxColor Then
      iDegrado = iDegree
    End If
    frmForm.Line (0, iDegree)-(Screen.Width, _
                               iDegree - 1), RGB(255 - iDegrado, 255 - iDegrado, 255 - iDegree), B
  Next iDegree
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub LimpiarControles(frmForm As Object, asControlName() As String)
'
'Autor       : Patricio Dahmen G.
'Versión     : 02/1998
'Descripción : Limpia controles contenidos en form.
'Parámetros  :
'     I        frmForm, formulario que contiene los controles.
'     I        asControlName(), array con el nombre de controles a no limpiar.
'
'Declara variables locales
  Dim iCtrl As Integer
  Dim q As Integer
  Dim Control
  'Recorre controles
  For Each Control In frmForm
    'Compara control
    q = LBound(asControlName)
    Do While q <= UBound(asControlName)
      If frmForm.Controls(iCtrl).Name <> asControlName(q) Then
        q = q + 1
      Else
        Exit Do
      End If
    Loop
    'Control limpiable
    If q > UBound(asControlName) Then
      'Limpia textbox
      If Left$(frmForm.Controls(iCtrl).Name, 3) = "txt" Then
        frmForm.Controls(iCtrl).Text = ""
      End If
      'Limpia label
      If Left$(frmForm.Controls(iCtrl).Name, 3) = "lbl" Then
        frmForm.Controls(iCtrl).Caption = ""
      End If
      'Limpia combobox
      If Left$(frmForm.Controls(iCtrl).Name, 3) = "cbo" Then
        frmForm.Controls(iCtrl).ListIndex = -1
      End If
      'Limpia checkbox
      If Left$(frmForm.Controls(iCtrl).Name, 3) = "chk" Then
        frmForm.Controls(iCtrl).Value = 0
      End If
    End If
    'Incrementa índice
    iCtrl = iCtrl + 1
  Next
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub CentrarForm(Formulario As Object)
'
'Autor       : Patricio Dahmen G.
'Versión     : 06/1999
'Descripción : Centra formulario.
'Parámetros  :
'     I        Formulario, form a centrar.
'
'Centra formulario dentro de la pantalla
  Formulario.Move (Screen.Width - Formulario.Width) \ 2, (Screen.Height - Formulario.Height) \ 2
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub IndexText(cboComboBox As Object, ByVal sTexto As String)
'
'Autor       : Patricio Dahmen G.
'Versión     : 02/1998
'Descripción : Parea índice combobox según texto.
'Parámetros  :
'     I        cboComboBox, control combobox.
'     I        sTexto, texto a comparar en combobox.
'
'Declara variables locales
  Dim ii As Integer
  With cboComboBox
    'Recorre combobox
    For ii = 0 To .ListCount - 1
      .ListIndex = ii
      'Compara cadenas
      If sTexto = .Text Then
        Exit For
      End If
    Next ii
    'Chequea si recorrió entero el control
    If ii = .ListCount Then
      .ListIndex = -1
    End If
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Mayus(ByRef KeyAscii As Integer, Optional ByVal bEnterNext As Boolean)
'
'Autor       : Patricio Dahmen G.
'Versión     : 02/1998
'Descripción : Cambia caracter a mayúscula.
'Parámetros  :
'    I/O       KeyAscii, código ASCII caracter.
'     I        bEnterNext, si es verdadero, ejecuta TAB. opcional.
'              código ASCII en mayúscula.
'
'Tecla enter
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If bEnterNext Then SendKeys "{TAB}"
    Exit Sub
  End If
  'Mayúscula
  KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub SelTexto(txtTextBox As Object)
  With txtTextBox
    .SelStart = Len(.Text)
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub MaxLenCombobox(cboComboBox As Object, ByVal iMaxLength As Integer, ByRef iAscii As Integer)
'evento KeyPress
  If Len(cboComboBox.Text) >= iMaxLength And _
     iAscii <> vbKeyBack Then
    iAscii = 0
  End If
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub UnirArchivos(ByVal sArchivo1 As String, ByVal sArchivo2 As String, ByVal sArchivoDestino As String)
  Dim Buffer() As Byte
  Open sArchivo1 For Binary Access Read As #1
  Open sArchivo2 For Binary Access Read As #2
  Open sArchivoDestino For Binary Access Write As #3
  ReDim Buffer(1 To LOF(1))
  Get #1, , Buffer
  Put #3, , Buffer
  ReDim Buffer(1 To LOF(2))
  Get #2, , Buffer
  Put #3, , Buffer
  Close #1, #2, #3
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub AjustarImagen(frmForm As Object, picPicture As Object, Optional ByVal bResize As Boolean)
'eventos Paint/Resize
  Dim i As Integer
  Dim j As Integer
  With picPicture
    If bResize Then
      .Left = -(.Width + 200)
      .Top = -(.Height + 200)
      Exit Sub
    End If
    For i = 0 To frmForm.ScaleHeight Step .Height
      For j = 0 To frmForm.ScaleWidth Step .Width
        frmForm.PaintPicture picPicture, j, i, .Width, .Height
      Next j
    Next i
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub NroID(lblNro As Object, ByVal lNroID As Long, Optional ByVal bRojo As Boolean)
  Dim lColor As Long
  Dim sNro As String
  If lNroID = 0 Then
    lColor = COLOR_NEGRO
    sNro = ""
    GoTo Sale
  End If
  If bRojo Then
    lColor = COLOR_ROJO
  Else
    lColor = COLOR_NEGRO
  End If
  sNro = lNroID
Sale:
  lblNro.ForeColor = lColor
  lblNro.Caption = sNro & " "
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Aceptar(fDetalle As Object, ByVal bModificar As Boolean, Optional ByVal sBotonAceptar As String)
'Declara variables locales
  Dim iCtrl As Integer
  Dim q As Integer
  Dim sControl As String
  Dim Control
  If sBotonAceptar = "" Then sBotonAceptar = "cmdAceptar"
  'Recorre controles
  For Each Control In fDetalle
    sControl = LCase(fDetalle.Controls(iCtrl).Name)
    If Left(sControl, 3) = "cmd" Then
      If sControl = LCase(sBotonAceptar) Then
        fDetalle.Controls(iCtrl).Enabled = bModificar
        q = -1
        Exit For
      End If
    End If
    iCtrl = iCtrl + 1
  Next
  If q = 0 Then MsgBox fDetalle.Name, vbCritical, "aceptar"
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub TotalItem(txtPrecio As Object, txtCantidad As Object, lblTotal As Object, Optional ByVal iDecimales As Integer)
  Dim zPrecio As Single
  Dim zCant As Single
  Dim zTotal As Single
  If IsNumeric(txtPrecio) Then zPrecio = MaxDecimal(txtPrecio)
  If IsNumeric(txtCantidad) Then zCant = MaxDecimal(txtCantidad)
  zTotal = zPrecio * zCant
  Etiqueta lblTotal, Cant(zTotal, iDecimales)
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub ConfigurarPantalla(frmForm As Object)
  With frmForm
    .Height = 7200
    .Width = 9600
    .ScaleMode = 3
  End With
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub VerArchivo(frmForma As Object, ByVal sFile As String)
  Dim xh As Long
  If sFile = "" Then Exit Sub
  Screen.MousePointer = 11
  xh = ShellExecute(frmForma.hwnd, "Open", sFile, "", "", 1)
  Screen.MousePointer = 0
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub Congelar(f As Object, Optional bCongela As Boolean)
  If bCongela Then
    LockWindowUpdate f.hwnd
  Else
    LockWindowUpdate 0&
  End If
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub titExcel(ByVal iColumna As Integer, ByVal sTitulo As String)
  If Not exTitulos Then Exit Sub
  aExcel(0, iColumna - 1) = sTitulo
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub addExcel(ByVal iFila As Integer, ByVal iColumna As Integer, ByVal sInfo As String)
  If exTitulos Then
    iFila = iFila + 1
  End If
  aExcel(iFila - 1, iColumna - 1) = sInfo
End Sub
' Fin del Procedimiento ******************************************

' .:. Nombre del Procedimiento .:.
Public Sub dimExcel(ByVal iFilas As Integer, ByVal iColumnas As Integer, Optional ByVal bTitulos As Boolean)
  exTitulos = False
  If bTitulos Then
    exTitulos = True
    iFilas = iFilas + 1
  End If
  ReDim aExcel(0 To iFilas - 1, 0 To iColumnas - 1)
End Sub
' Fin del Procedimiento ******************************************


'   *   **   ***   ****  FUNCIONES   ****   ***   **   *
' .:. Nombre de la Función .:.
Private Function RutBD(ByVal sRut As String) As String
  Dim rdv As String
  rdv = Right(sRut, 2)
  sRut = Left(sRut, Len(sRut) - 2)
  sRut = Val(sRut)
  sRut = sRut & rdv
  RutBD = sRut
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Cortar(ByVal sTexto As String, ByVal iMenos As Integer, Optional ByVal iInicio As Integer) As String
  Dim g As Integer
  Dim u As Integer
  If iInicio = 0 Then
    iInicio = 1
  Else
    sTexto = Mid(sTexto, iInicio)
  End If
  g = Len(sTexto)
  u = g - iMenos
  Cortar = Left(sTexto, u)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Cuenta(ByVal sTexto As String, ByVal sBuscar As String, Optional ByVal iInicio As Integer) As Long
  Dim z As Integer
  Dim g As Integer
  Dim u As Integer
  If iInicio = 0 Then iInicio = 1
  z = 1
  g = Len(sBuscar)
  u = 0
  Do While z > 0
    z = InStr(iInicio, sTexto, sBuscar, vbTextCompare)
    If z > 0 Then
      iInicio = z + g
      u = u + 1
    End If
  Loop
  Cuenta = u
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Reemplazar(ByVal sTexto As String, ByVal sBuscar As String, ByVal sReemplazar As String, Optional ByVal iInicio As Integer, Optional iContar As Integer, Optional iEspacios As Integer) As Variant
  Dim z As Integer
  Dim g As Integer
  Dim u As Integer
  Dim v As Integer
  Dim zTexto As String
  If iInicio = 0 Then iInicio = 1
  zTexto = sTexto
  v = Len(sReemplazar)
  g = Len(sBuscar)
  If v = 0 Then v = g
  v = v * Len(sTexto)
  z = 1
  Do While z > 0
    z = InStr(iInicio, zTexto, sBuscar, vbTextCompare)
    If z > 0 Then
      iInicio = z + g
      zTexto = Left(zTexto, z - 1) & sReemplazar & Mid(zTexto, iInicio)
      iInicio = iInicio + iEspacios
      u = u - 1
    End If
    If iContar = Abs(u) Then
      z = 0
    Else
      If Abs(u) > v Then
        zTexto = sTexto
        z = 0
      End If
    End If
  Loop
  Reemplazar = zTexto
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function nValor(ByVal sValor As String, Optional ByVal bDecimal As Boolean) As Double
  If Not IsNumeric(sValor) Then sValor = ""
  If sValor = "" Then sValor = 0
  If bDecimal Then
    sValor = Valor(sValor)
  Else
    sValor = Unformat(sValor)
  End If
  nValor = CDbl(sValor)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function CantLabel(ByVal sCantidad As String) As String
  Dim sValor As String
  sValor = RTrim(sCantidad)
  sValor = ValorSQL(sValor)
  sValor = MaxDecimal(sValor)
  CantLabel = sValor
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Folio(ByVal lNumero As Long, Optional ByVal iCeros As Integer) As String
  Dim sCeros As String
  If iCeros = 0 Then iCeros = 4
  sCeros = String(iCeros, "0")
  Folio = Format(lNumero, sCeros)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function NombreDOS(ByVal sFile As String) As String
  Dim sExt As String
  Dim sNombreCorto As String * 255
  Dim p As Integer
  Dim nDOS As String
  sExt = Right(sFile, 4)
  GetShortPathName sFile, sNombreCorto, 255
  nDOS = sNombreCorto
  p = InStr(1, nDOS, sExt, vbTextCompare)
  nDOS = Left(nDOS, p + Len(sExt) - 1)
  NombreDOS = nDOS
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function dv(ByVal lCarnet As Long) As String
'Declara variables locales
  Dim sCte As String
  Dim sRut As String
  Dim iSuma As Integer
  Dim iDigito As Integer
  Dim i As Integer
  'Controla Error
  On Error GoTo DV_Err
  'Crea string constante
  sCte = "32765432"
  'Extrae -DV
  sRut = lCarnet
  'Setea rut
  sRut = Format(sRut, "00000000")
  'Obtiene suma dígitos
  iSuma = 0
  For i = 1 To 8
    iSuma = iSuma + (Val(Mid$(sCte, i, 1)) * Val(Mid$(sRut, i, 1)))
  Next
  'Obtiene dígito
  iDigito = 11 - (iSuma Mod 11)
  Select Case iDigito
  Case 10
    dv = "K"
  Case 11
    dv = "0"
  Case Else
    dv = Trim$(CStr(iDigito))
  End Select
  dv = sRut & "-" & dv
  Exit Function
DV_Err:
  dv = "E"
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function EditRut(ByVal sRut As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 12/2000
'Descripción : Edita rut.
'Parámetros  :
'     I        sRut, rut formatado.
'     O        Rut sin formato y sin ceros a la izquierda.
'
'Declara variables locales
  Dim dv As String
  Dim sNro As String
  Dim lNro As Long
  'valida parámetro
  If Len(Trim(sRut)) = 0 Then Exit Function
  'quita formato
  sNro = Unformat(sRut)
  'dígito verificador
  dv = Right(sNro, 2)
  'valor rut
  lNro = Val(Left(sNro, Len(sNro) - 2))
  'retorna rut para edición
  EditRut = CStr(lNro) & dv
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Invertir(ByVal sCadena As String) As String
'Declara variables locales
  Dim i As Integer
  Dim sSalida As String
  Dim u As Integer
  sSalida = ""
  u = Len(sCadena)
  For i = u To 1 Step -1
    sSalida = sSalida & Mid(sCadena, i, 1)
  Next i
  'Retorna string
  Invertir = sSalida
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function InStrRev(ByVal sCadena As String, ByVal sBuscado As String, ByVal iInicio As Integer) As Long
'Declara variables locales
  Dim i As Integer
  Dim u As Integer
  i = Len(sCadena)
  sCadena = Invertir(sCadena)
  sBuscado = Invertir(sBuscado)
  u = InStr(1, sCadena, sBuscado, vbTextCompare)
  InStrRev = i - u + 1
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function DiaHabil(ByVal dFecha As Date, ByVal iDias As Integer) As Date
'Declara variables locales
  Dim iDia As Integer
  Dim di As Integer
  Dim iDia0 As Integer
  Dim dDia As Date
  di = 1
  Do While di <= iDias
    'fecha
    dDia = DateAdd("d", di, dFecha)
    iDia0 = Weekday(dDia, vbMonday)
    'obtiene día semana
    If iDia0 >= 6 Then
      iDia = iDia + 1
      iDias = iDias + 1
    End If
    di = di + 1
  Loop
  dDia = DateAdd("d", iDias, dFecha)
  iDia0 = Weekday(dDia, vbMonday)
  'obtiene día semana
  If iDia0 >= 6 Then
    If iDia0 = 6 Then
      iDia = 2
    Else
      iDia = 1
    End If
  Else
    iDia = 0
  End If
  'obtiene día
  DiaHabil = DateAdd("d", iDia, dDia)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function FechaLarga(ByVal dFecha As Date) As String
  FechaLarga = Format(dFecha, "dddddd", vbMonday)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function FechaSinSemana(ByVal dFecha As Date) As String
  Dim sFecha As String
  Dim p As Integer
  sFecha = FechaLarga(dFecha)
  p = InStr(1, sFecha, ",", vbTextCompare)
  FechaSinSemana = Mid(sFecha, p + 2)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Dia1(ByVal iMes As Integer, ByVal iAño As Integer) As Date
  If iMes = 0 Then iMes = Month(Now)
  If iAño = 0 Then iAño = Year(Now)
  Dia1 = CDate(Format("01/" & iMes & "/" & iAño, "dd/mm/yyyy"))
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ExtraeEOL(ByVal sCadena As String, Optional ByVal bInicio As Boolean) As String
'Declara variables locales
  Dim i As Integer
  Dim sSalida As String
  'Inicializa salida
  sSalida = ""
  'Obtiene caracteres cadena
  For i = 1 To Len(sCadena)
    'Verifica caracter
    If Not Mid(sCadena, i, 1) = Chr(13) Then
      sSalida = sSalida & Mid(sCadena, i, 1)
    Else
      If Not Mid(sCadena, i + 1, 1) = Chr(10) Then
        sSalida = sSalida & Mid(sCadena, i + 1, 1)
      Else
        If bInicio Then
          If i = 1 Then
            sSalida = Mid(sCadena, 3)
          Else
            sSalida = sCadena
          End If
          If Right(sSalida, 1) = Chr(10) Then sSalida = Left(sSalida, Len(sSalida) - 2)
          Exit For
        Else
          If i < Len(sCadena) - 1 Then
            If i > 1 Then sSalida = sSalida & " "
          End If
        End If
        i = i + 1
      End If
    End If
  Next i
  'Retorna string
  ExtraeEOL = LTrim(sSalida)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ExtraeEnter(ByVal sCadena As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Extrae todos los  enter en sCadena.
'Parámetros  :
'     I        sCadena, cadena a extaer.
'     O        cadena sin ocurrencias.
'
'Declara variables locales
  Dim i As Integer
  Dim sSalida As String
  'Inicializa salida
  sSalida = ""
  'Obtiene caracteres cadena
  For i = 1 To Len(sCadena)
    'Verifica caracter
    If Not Mid(sCadena, i, 1) = Chr(13) Then
      sSalida = sSalida & Mid(sCadena, i, 1)
    Else
      If Not Mid(sCadena, i + 1, 1) = Chr(10) Then
        sSalida = sSalida & Mid(sCadena, i + 1, 1)
      Else
        sSalida = sSalida & " "
        i = i + 1
      End If
    End If
  Next i
  'Retorna string
  ExtraeEnter = sSalida
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function CarpetaOculta(ByVal sDir As String) As Boolean
  CarpetaOculta = (Mid(sDir, Len(sDir), 1) = "}")
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Cero(ByVal vValorCampo As Variant, Optional ByVal bFormat As Boolean, Optional ByVal byDec As Byte) As String
  If vValorCampo <> 0 Then
    If Not bFormat Then
      Cero = vValorCampo
    Else
      Cero = Cant(vValorCampo, byDec)
    End If
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Primera(ByVal sCadena As String) As String
  sCadena = LCase(sCadena)
  Primera = UCase(Left(sCadena, 1)) & Mid(sCadena, 2)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Dia(ByVal vFecha As Variant) As String
  Dia = UCase(Format(vFecha, "dd/mmm/yyyy"))
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function RutaArchivo(ByVal sRuta As String, ByVal sNombreApp As String, Optional ByVal bUltimaBarra As Boolean) As String
  Dim p As Integer
  Dim na As String
  If Len(sRuta) = 0 Then Exit Function
  na = NombreArchivo(sRuta)
  p = InStr(1, sRuta, na, vbTextCompare)
  If p = 0 Then
    MsgBox "Error en directorio '" & sRuta & "'.", vbCritical, sNombreApp
    Exit Function
  End If
  If Not bUltimaBarra Then
    p = p - 1
  End If
  RutaArchivo = Mid(sRuta, 1, p - 1)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Foco(txtNro As Object) As Boolean
  With txtNro
    Foco = True
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ComaDec(ByVal lTecla As Long, ByVal sCaja As String) As Long
  Dim t As String
  Dim cOtro As String
  Dim bDec As Boolean
  Dim bHay As Boolean
  If SEP_DEC = "," Then
    cOtro = "."
  Else
    cOtro = ","
  End If
  bDec = (lTecla = Asc(SEP_DEC) Or lTecla = Asc(cOtro))
  bHay = (InStr(1, sCaja, SEP_DEC, vbTextCompare) > 0)
  If bDec And bHay Then Exit Function
  t = Chr(lTecla)
  Select Case t
  Case cOtro
    ComaDec = Asc(SEP_DEC)
  Case Else
    ComaDec = lTecla
  End Select
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Porcentaje(ByVal uCantidad As Double, ByVal zPonderacion As Single, ByVal iDecimales As Integer) As Single
  Dim zValor As Single
  If zPonderacion = 0 Then Exit Function
  zValor = Abs(uCantidad * (zPonderacion / 100))
  If iDecimales >= 0 Then
    Porcentaje = Redondear(zValor, iDecimales)
  Else
    Porcentaje = zValor
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ValorDec(ByVal sValor As String) As String
  Dim zMil As String
  Dim zDec As String
  Dim p As Integer
  Dim q As Integer
  If SEP_DEC = "," Then
    zMil = "."
    zDec = ","
  Else
    zMil = ","
    zDec = "."
  End If
  If sValor = "" Then sValor = 0
  p = InStr(1, sValor, zMil, vbTextCompare)
  If p > 0 Then
    sValor = Reemplazar(sValor, zMil, zDec)
  End If
  q = InStr(1, sValor, zDec, vbTextCompare)
  If q > 0 Then
    sValor = Reemplazar(sValor, zDec, SEP_DEC)
  End If
  ValorDec = sValor
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Valor(ByVal sValor As String) As String
  Dim p As Integer
  If sValor = "" Then sValor = 0
  p = InStr(1, sValor, ".", vbTextCompare)
  If p > 0 Then Mid(sValor, p) = ","
  Valor = sValor
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Valor0(txtCuadro As Object) As Long
  Dim tttt As String
  Dim qh As Integer
  tttt = txtCuadro.Text
  qh = Len(tttt)
  If qh = 0 Then
    Valor0 = 0
    Exit Function
  End If
  If Not IsNumeric(tttt) Then
    Valor0 = 0
  End If
  Valor0 = tttt
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function OrdenFecha(ByVal dFecha As Date) As String
  OrdenFecha = Format(Year(dFecha), "0000") & _
               Format(Month(dFecha), "00") & _
               Format(Day(dFecha), "00")
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function MesAño(ByVal iMes As Integer, ByVal iAño As Integer, Optional ByVal b3 As Boolean, Optional ByVal sSeparacion As String) As String
  Dim sMes As String
  If iAño = 0 Then iAño = Year(Now)
  If iMes = 0 Then
    MesAño = "AÑO " & iAño
    Exit Function
  End If
  If sSeparacion = "" Then sSeparacion = "."
  sMes = Mes(iMes, True)
  If b3 Then sMes = Left(sMes, 3)
  MesAño = sMes & sSeparacion & iAño
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Menor(ByVal vValor1 As Variant, ByVal vValor2 As Variant) As Double
  If IsNull(vValor1) Then vValor1 = 0
  If IsNull(vValor2) Then vValor2 = 0
  vValor1 = nValor(vValor1)
  vValor2 = nValor(vValor2)
  If vValor1 > vValor2 Then
    Menor = vValor2
  Else
    Menor = vValor1
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Mayor(ByVal vValor1 As Variant, ByVal vValor2 As Variant) As Double
  If IsNull(vValor1) Then vValor1 = 0
  If IsNull(vValor2) Then vValor2 = 0
  vValor1 = nValor(vValor1)
  vValor2 = nValor(vValor2)
  If vValor1 >= vValor2 Then
    Mayor = vValor1
  Else
    Mayor = vValor2
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function aMenor(ByVal aArray) As Double
  Dim i As Long
  Dim uMenor As Double
  Dim vValor As Variant
  uMenor = aArray(0)
  For i = 0 To UBound(aArray)
    vValor = aArray(i)
    If vValor <= uMenor Then
      uMenor = vValor
    End If
  Next i
  aMenor = uMenor
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function aMayor(ByVal aArray) As Double
  Dim i As Long
  Dim uMayor As Double
  Dim vValor As Variant
  For i = 0 To UBound(aArray)
    vValor = aArray(i)
    If vValor >= uMayor Then
      uMayor = vValor
    End If
  Next i
  aMayor = uMayor
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Minutos(ByVal sHoraMenor As String, ByVal sHoraMayor As String) As Long
  If sHoraMenor = "" Then sHoraMenor = Reloj(0, 0)
  If sHoraMayor = "" Then sHoraMayor = Reloj(0, 0)
  Minutos = Abs(DateDiff("n", sHoraMenor, sHoraMayor))
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Reloj(ByVal iHora As Integer, ByVal iMinutos As Integer) As String
  Reloj = Hora(iHora, iMinutos, 0)
  Reloj = Left(Reloj, 5)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function MaxDecimal(ByVal sValor As String, Optional ByVal iDecimales As Integer, Optional ByVal bValorDecimal As Boolean) As String
  Dim sNumero As String
  If iDecimales = 0 Then iDecimales = 2
  If bValorDecimal Then
    sNumero = ValorDec(sValor)
  Else
    sNumero = Valor(sValor)
  End If
  MaxDecimal = Redondear(sNumero, iDecimales)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Periodo(ByVal dDesde As Date, ByVal dHasta As Date) As String
'Declara variables locales
  Dim iMesInicial As Byte
  Dim iAñoInicial As Integer
  Dim iMesFinal As Byte
  Dim iAñoFinal As Integer
  'obtiene mes y año inicial
  iMesInicial = Month(dDesde)
  iAñoInicial = Year(dDesde)
  'obtiene mes y año final
  iMesFinal = Month(dHasta)
  iAñoFinal = Year(dHasta)
  'determina mismo año
  If iAñoInicial = iAñoFinal Then
    'determina mismo mes
    If iMesInicial = iMesFinal Then
      Periodo = Mes(iMesInicial, True) & " DE " & iAñoInicial
    Else
      Periodo = Mes(iMesInicial, True) & "-" & _
                Mes(iMesFinal, True) & " DE " & iAñoInicial
    End If
  Else
    Periodo = Mes(iMesInicial, True) & " DE " & iAñoInicial & _
              " - " & Mes(iMesFinal, True) & " DE " & iAñoFinal
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function RutOK(TextboxRut As Object, Optional ByVal sMensaje As String, Optional ByVal sTitulo As String) As Boolean
  Dim tttt As String
  Dim qh As Integer
  tttt = TextboxRut.Text
  qh = Len(tttt)
  If qh = 0 Then Exit Function
  If InStr(1, tttt, "-", vbTextCompare) = 0 Then
    tttt = Left(tttt, qh - 1) & "-" & Right(tttt, 1)
  End If
  tttt = RutBD(tttt)
  If Len(tttt) > 10 Then tttt = Mid(tttt, 2)
  'verifica rut
  If LCase(DigitoVerificador(tttt)) = LCase(Right(tttt, 1)) Then
    TextboxRut.Text = tttt
    RutOK = True
  Else
    If sMensaje <> "" Then
      MsgBox sMensaje, vbCritical, sTitulo
      TextboxRut.SetFocus
      SendKeys "{Home}+{End}"
    End If
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function WinExe(ByVal sFileExe As String) As Boolean
  Dim fh As Integer
  Dim t As String * 1
  fh = FreeFile
  Open sFileExe For Binary As #fh
  Get fh, 25, t
  Close #fh
  WinExe = (Asc(t) = &H40&)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Comas(ByVal sCadena As String, Optional bSimple As Boolean) As String
  Dim sComa As String
  If bSimple Then
    sComa = "'"
  Else
    sComa = Chr(34)
  End If
  Comas = sComa & sCadena & sComa
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Porciento(ByVal uParcial As Double, ByVal uTotal As Double) As Single
  If uTotal = 0 Then Exit Function
  Porciento = (uParcial / uTotal) * 100
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Hora(ByVal lHoras As Long, ByVal lMinutos As Long, ByVal lSegundos As Long) As String
  Dim xx As Long
  If lSegundos >= 60 Then
    xx = lSegundos \ 60
    lSegundos = lSegundos Mod 60
    lMinutos = lMinutos + xx
  End If
  If lMinutos >= 60 Then
    xx = lMinutos \ 60
    lMinutos = lMinutos Mod 60
    lHoras = lHoras + xx
  End If
  Hora = Format(lHoras, "00") & ":" & _
         Format(lMinutos, "00") & ":" & _
         Format(lSegundos, "00")
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Hoy(ByVal bMMM As Boolean, ByVal bMayus As Boolean) As String
  Dim sHoyDia As String
  'Mes de forma MMM
  If bMMM Then
    sHoyDia = Format(Now, "dd/mmm/yyyy")
  Else
    sHoyDia = Format(Now, "dd/mm/yyyy")
  End If
  'Mes en mayúscula
  If bMayus Then
    Hoy = UCase(sHoyDia)
  Else
    Hoy = sHoyDia
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function HoraDec(ByVal sHora As String) As Single
  HoraDec = Minutos("00:00", sHora) / 60
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Cant(ByVal uMonto As Double, Optional ByVal byDecimales As Byte) As String
'declara variables locales
  Dim sDecimal As String
  Dim sFormato As String
  'formato
  sFormato = "###,###,###,###,##0"
  If byDecimales <> 0 Then
    sFormato = sFormato & "." & String(byDecimales, "0")
  End If
  'retorna cantidad
  Cant = Format(uMonto, sFormato)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function TituloBox(ByVal iAccion As Integer, ByVal bFemenino As Boolean, ByVal sEntidad As String) As String
  Dim sTitBox As String
  Select Case iAccion
  Case BOX_NEW
    sTitBox = "Nuevo "
    If bFemenino Then sTitBox = "Nueva "
    sTitBox = sTitBox & sEntidad
  Case BOX_ADD
    sTitBox = "Agregar " & sEntidad
  Case BOX_MOD
    sTitBox = "Modificar " & sEntidad
  Case BOX_DEL
    sTitBox = "Eliminar " & sEntidad
  Case Else
    sTitBox = sEntidad
  End Select
  TituloBox = sTitBox
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Prefijo(ByVal sCadena As String, ByVal sPrefijo As String) As String
  Prefijo = sPrefijo & sCadena
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Redondear(ByVal uNro As Double, Optional iDec As Integer) As Double
  Dim uPot As Double
  Dim uF As Double
  If uNro < 0 Then
    uF = -0.5
  Else
    uF = 0.5
  End If
  uPot = 10 ^ iDec
  Redondear = Fix(uNro * uPot * (1 + 1E-16) + uF) / uPot
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ComboItemData(cboCombo As Object) As Long
'obtiene ítem
  With cboCombo
    If .ListIndex < 0 Then Exit Function
    ComboItemData = .ItemData(.ListIndex)
  End With
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function FormatoRut(ByVal sRut As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 06/1999
'Descripción : Formatea rut.
'Parámetros  :
'     I        sRut, rut de la forma 99999999-A.
'     O        Rut formateado de la forma 99.999.999-A.
'
'Declara variables locales
  Dim dv As String
  'Extrae DV
  dv = Right(sRut, 2)
  sRut = Left(sRut, Len(sRut) - 2)
  'Aplica formato
  sRut = Format(sRut, "00,000,000")
  'Rut formateado
  FormatoRut = sRut & dv
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function LeeIni(ByVal sPath As String, ByVal sArchivo As String, ByVal sCabecera As String, ByVal sItem As String, ByVal sDefecto As String) As Variant
'
'Autor       : Patricio Dahmen G.
'Versión     : 03/1998
'Descripción : Lee item archivo INI.
'Parámetros  :
'    I         sPath, ruta ubicación archivo INI.
'    I         sArchivo, nombre archivo Ini sin extensión
'    I         sCabecera, head del grupo de items.
'    I         sItem, item a leer.
'    I         sDefecto, retorno por defecto.
'    O         contenido del item. En caso de error retorna 'sDefecto'.
'
'Declara variables locales
  Dim sRetorno As String
  Dim sPathIni As String
  Dim iLeidos As Integer
  'Mensaje al ocurrir error
  On Error GoTo ErrorLeeIni
  'Inicializa string que devuelve variable
  sRetorno = String$(255, 32)
  'Crea ruta completa
  If Right$(sPath, 1) = "\" Then
    sPathIni = sPath & sArchivo & ".ini"
  Else
    sPathIni = sPath & "\" & sArchivo & ".ini"
  End If
  'Llama función del Kernel
  iLeidos = GetPrivateProfileString(sCabecera, sItem, "", sRetorno, 255, sPathIni)
  'Valida lectura
  If iLeidos = 0 Then
    LeeIni = sDefecto
  Else
    LeeIni = Left$(sRetorno, iLeidos)
  End If
  Exit Function
ErrorLeeIni:
  MsgBox "Imposible leer archivo de configuración.", vbCritical, "Error de lectura"
  Resume Next
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Mes(ByVal iMes As Integer, ByVal bMayus As Boolean) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 07/2000
'Descripción : Obtiene mes en palabras.
'Parámetros  :
'     I        iMes, número del mes.
'     I        bMayus, verdadero si mes es mayúscula.
'     O        Mes en palabras. En otro caso, retorna "".
'
'Valida parámetro
  If iMes < 1 And iMes > 12 Then Exit Function
  'Obtiene mes en palabras
  Mes = Format("18/" & CStr(iMes) & "/1990", "mmmm")
  'Mayus
  If bMayus Then
    Mes = UCase(Mes)
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function FullName(ByVal sPaterno As String, ByVal sMaterno As String, ByVal sNombres As String, Optional ByVal iForma As Integer) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 07/2000.
'Descripción : Concatena nombres y apellidos.
'Parámetros  :
'     I        sPaterno, apellido paterno.
'     I        sMaterno, apellido materno.
'     I        sNombres, nombres.
'     I        sForma, forma de concatenación (opcional):
'              1 -> paterno materno, nombres. Por defecto.
'              2 -> nombres paterno materno.
'              3 -> 1er_nombre paterno materno.
'              4 -> 1er_nombre paterno inicial_materno.
'              5 -> paterno inicial_materno, nombres.
'     O        Nombre completo.
'
'Declara variables locales
  Dim iBlanco As Integer
  Dim s1Name As String
  'Determina forma y devuelve nombre completo
  Select Case iForma
  Case 1
    FullName = sNombres & " " & sPaterno & " " & sMaterno
  Case 3, 4
    iBlanco = InStr(1, sNombres, " ", vbTextCompare)
    If iBlanco <> 0 Then
      s1Name = Left$(sNombres, iBlanco) & sPaterno
    Else
      s1Name = sNombres & " " & sPaterno
    End If
    If iForma = 4 Then
      FullName = s1Name & " " & Left$(sMaterno, 1) & "."
    Else
      FullName = s1Name & " " & sMaterno
    End If
  Case 5
    FullName = sPaterno & " " & Left$(sMaterno, 1) & ", " & sNombres
  Case Else
    FullName = sPaterno & " " & sMaterno & ", " & sNombres
  End Select
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function IndexButton(optButton As Object) As Integer
'
'Autor       : Patricio Dahmen G.
'Versión     : 12/1999
'Descripción : Obtiene el índice del option button seleccionado.
'Parámetros  :
'     I        Array de controles option button.
'     O        índice option button seleccionado. En cualquier otro
'              caso, retorna -1.
'
'Declara variables locales
  Dim ind As Integer
  'Inicializa retorno
  IndexButton = -1
  'Recorre array
  For ind = 0 To optButton.UBound
    If optButton(ind).Value Then
      IndexButton = ind
      Exit Function
    End If
  Next ind
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Segundos(ByVal lHoras As Long, ByVal lMinutos As Long, ByVal lSeconds As Long) As Long
'
'Autor       : Patricio Dahmen G.
'Versión     : 03/2000
'Descripción : Comvierte tiempo hh:mm:ss a segundos.
'Parámetros  :
'     I        lHoras, horas.
'     I        lMinutos, mintutos.
'     I        lSeconds, segundos.
'     O        tiempo convertido a segundos.
'
'Valida parámetros
  If lHoras < 0 Then lHoras = 0
  If lMinutos < 0 Then lMinutos = 0
  If lSeconds < 0 Then lSeconds = 0
  'Convierte horas
  lHoras = lHoras * 3600
  'Convierte minutos
  lMinutos = lMinutos * 60
  'Retorna segundos
  Segundos = lHoras + lMinutos + lSeconds
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ValidaTxt(ByRef iAscii As Integer, ByVal sTipo As String, ByVal bEnterNext As Boolean, Optional ByVal sEspeciales As String) As Integer
'
'Autor       : Patricio Dahmen G.
'Versión     : 12/1998
'Descripción : Valida caracteres en Textbox.
'Parámetros :
'     I        iAscii, código ascii caracter,
'     I        sTipo, tipo caracter permitido.
'              9, solo dígitos
'              A!, solo mayúsculas
'              A, solo caracteres
'              X, caracteres y dígitos
'               , cualquier caracter
'     I        bEnterNext, si es verdadero, ejecuta TAB.
'     I        sEspeciales, caracteres especiales permitidos. Opcional.
'     O        código ascii caracter permitido. En otro caso, 0.
'
'Declara variables locales
  Dim i As Integer
  'Tecla enter
  If iAscii = vbKeyReturn Then
    iAscii = 0
    If bEnterNext Then SendKeys "{TAB}"
    Exit Function
  End If
  'Convierte mayúscula
  sTipo = UCase(sTipo)
  'Valida caracter
  Select Case iAscii
  Case 48 To 57              'números
    If sTipo = "9" Or sTipo = "X" Or sTipo = "" Then
      ValidaTxt = iAscii
    Else
      ValidaTxt = 0
    End If
  Case 65 To 90              'mayúsculas
    If sTipo = "A" Or sTipo = "A!" Or sTipo = "X" Or sTipo = "" Then
      ValidaTxt = iAscii
    Else
      ValidaTxt = 0
    End If
  Case 97 To 122             'minúsculas
    If sTipo = "A" Or sTipo = "X" Or sTipo = "" Then
      ValidaTxt = iAscii
    Else
      ValidaTxt = 0
    End If
  Case 4, 5, 19, 24          'flechas
    ValidaTxt = iAscii
  Case 1, 3, 6, 7, 18, 22    'páginas
    ValidaTxt = iAscii
  Case 8, 9, 13, 27          'enter,esc,back,tab
    ValidaTxt = iAscii
  Case 32                    'space
    If sTipo = "A" Or sTipo = "A!" Or sTipo = "X" Or sTipo = "" Then
      ValidaTxt = iAscii
    Else
      ValidaTxt = 0
    End If
  Case 34, 39, 44, 96        'comillas
    ValidaTxt = 0
  Case Else
    ValidaTxt = 0
    For i = 1 To Len(sEspeciales)
      If iAscii = Asc(Mid$(sEspeciales, i, 1)) Then
        ValidaTxt = iAscii
        Exit For
      End If
    Next i
  End Select
  'Caracteres especiales (letras y nros.)
  If Len(sEspeciales) > 0 And ValidaTxt = 0 Then
    For i = 1 To Len(sEspeciales)
      If iAscii = Asc(Mid$(sEspeciales, i, 1)) Then
        ValidaTxt = iAscii
        Exit For
      End If
    Next i
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Unformat(ByVal sCant As String) As Variant
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : antónimo función Format().
'Parámetros  :
'     I        sCant, cantidad a quitar format.
'     O        cantidad pura.
'
'Valida parámetro
  If sCant = "" Then
    Unformat = 0
    Exit Function
  End If
  'Extrae formato
  Unformat = Extrae(Extrae(sCant, ","), ".")
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function SoloRut(ByVal sRutFormateado As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Extrae caracteres {. , - DV}.
'Parámetros  :
'     I        sRutFormateado, rut con formato.
'     O        rut sin formato.
'
'Declara variables locales
  Dim sRut As String
  'Quita separador de miles
  sRut = Extrae(Extrae(sRutFormateado, ","), ".")
  'Quita guión
  sRut = Extrae(sRut, "-")
  'Quita DV
  sRut = Left$(sRut, Len(sRut) - 1)
  'Retorna resultado
  SoloRut = CStr(Val(sRut))
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function NombrePC() As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene el nombre del equipo.
'Parámetros  :
'     O        Nombre del PC. "" en cualquier otro caso.
'
'Declara variables locales
  Dim sEquipo As String * 255
  Dim sName As String
  Dim p As Integer
  'Obtiene nombre pc
  GetComputerName sEquipo, 255
  p = InStr(1, sEquipo, Chr(0), vbTextCompare)
  If p > 0 Then
    sName = Left(sEquipo, p - 1)
  End If
  'Retorna resultado
  NombrePC = sName
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function NombreArchivo(ByVal sRuta As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene nombre archivo extrayendo path.
'Parámetros  :
'    I         sRuta, nombre archivo que incluye path.
'    O         nombre archivo. En otro caso retorna ""
'
'Declara variables locales
  Dim iCont As Integer
  Dim sCad As String
  'Inicializa retorno
  NombreArchivo = ""
  'Valida parámetro
  If sRuta = "" Then
    Exit Function
  End If
  'Copia path
  sCad = sRuta
  'Recorre path
  iCont = 0
  Do While True
    If Right$(sCad, 1) = "\" Then
      Exit Do
    Else
      sCad = Left$(sCad, Len(sCad) - 1)
      iCont = iCont + 1
    End If
  Loop
  'Retorna nombre archivo
  NombreArchivo = Right$(sRuta, iCont)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Encripta(ByVal sCadena As String, ByVal bEncripta As Boolean) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Encripta o desencripta cadena.
'Obs.        : Sólo permite caracteres alfabéticos sin acento.
'Parámetros  :
'     I        sCadena, cadena a encriptar o desencriptar.
'     I        bEncripta, acción a realizar. True=encripta, False=Desencripta.
'     O        cadena encriptada o desencriptada.
'
'Declara variables locales
  Dim sChar As String
  Dim sNewChar As String
  Dim iLargo As Integer
  Dim i As Integer
  Dim iAccion As Integer
  'Obtiene largo cadena
  iLargo = Len(Trim$(sCadena))
  'Determina acción
  If bEncripta Then
    iAccion = 1
  Else
    iAccion = -1
  End If
  'Recorre cadena
  For i = 1 To iLargo
    'Obtiene caracter
    sChar = Mid$(sCadena, i)
    'Encripta caracter
    sNewChar = sNewChar & Chr$(Asc(sChar) + (iLargo * iAccion))
  Next i
  'Retorna resultado
  Encripta = sNewChar
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function EdadActual(ByVal dFechaInicial As Date) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene edad actual.
'Parámetros  :
'     I        dFechaInicial, fecha partida.
'     O        edad actual de la forma "99 años, 99 meses".
'              Si es negativa retorna entre paréntesis.
'
'Declara variables locales
  Dim iAños As Integer
  Dim iMeses As Integer
  Dim iTotMes As Integer
  Dim bNeg As Boolean
  'Obtiene total meses
  iTotMes = DateDiff("m", dFechaInicial, Now)
  'Obtiene signo
  bNeg = (iTotMes < 0)
  'Obtiene años
  iAños = iTotMes \ 12
  'Obtiene meses
  iMeses = iTotMes Mod 12
  'Setea salida
  EdadActual = ""
  'Obtiene años
  If iAños <> 0 Then
    If Abs(iAños) = 1 Then
      EdadActual = CStr(Abs(iAños)) & " año, "
    Else
      EdadActual = CStr(Abs(iAños)) & " años, "
    End If
  End If
  'Obtiene meses
  If iMeses <> 0 Then
    If Abs(iMeses) = 1 Then
      EdadActual = EdadActual & CStr(Abs(iMeses)) & " mes"
    Else
      EdadActual = EdadActual & CStr(Abs(iMeses)) & " meses"
    End If
  End If
  'Extrae ", " si meses = 0
  If Right$(EdadActual, 2) = ", " Then
    EdadActual = Left$(EdadActual, Len(EdadActual) - 2)
  End If
  'Fecha negativa
  If bNeg Then
    EdadActual = "(" & EdadActual & ")"
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function EscribeIni(ByVal sPath As String, ByVal sArchivo As String, ByVal sCabecera As String, ByVal sItem As String, ByVal sDato As String) As Boolean
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Escribe item del archivo INI.
'Parámetros  :
'    I         sPath, ruta ubicación archivo INI.
'    I         sArchivo, nombre archivo Ini sin extensión.
'    I         sCabecera, head del grupo de items.
'    I         sItem, item a leer.
'    I         sDato, parámetro a escribir.
'    O         Verdadero si pudo realizar la escritura. Falso en cualquier otro caso.
'
'Declara variables locales
  Dim sPathIni As String
  Dim lEscritos As Long
  'Mensaje al ocurrir error
  On Error GoTo ErrorEscribeIni
  'Crea ruta completa
  If Right$(sPath, 1) = "\" Then
    sPathIni = sPath & sArchivo & ".ini"
  Else
    sPathIni = sPath & "\" & sArchivo & ".ini"
  End If
  'Llama función del Kernel
  lEscritos = WritePrivateProfileString(sCabecera, sItem, sDato, sPathIni)
  'Chequea escritura
  If lEscritos = 0 Then
    EscribeIni = False
  Else
    EscribeIni = True
  End If
  Exit Function
ErrorEscribeIni:
  MsgBox "Imposible actualizar archivo de configuración.", vbCritical, "Error de escritura"
  Resume Next
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function LeeHide(ByVal sPath As String, ByVal sArchivo As String, ByVal sCabecera As String, ByVal sItem As String, ByVal sDefecto As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 03/1998
'Descripción : Lee item archivo oculto.
'Parámetros  :
'    I         sPath, ruta ubicación archivo.
'    I         sArchivo, nombre archivo con extensión.
'    I         sCabecera, head del grupo de items.
'    I         sItem, item a leer.
'    I         sDefecto, retorno por defecto.
'    O         contenido del item. En caso de error retorna 'sDefecto'.
'
'Declara variables locales
  Dim sRetorno As String
  Dim sPathHide As String
  Dim iLeidos As Integer
  'Mensaje al ocurrir error
  On Error GoTo ErrorLeeHide
  'Inicializa string que devuelve variable
  sRetorno = String$(255, 32)
  'Crea ruta completa
  If Right$(sPath, 1) = "\" Then
    sPathHide = sPath & sArchivo
  Else
    sPathHide = sPath & "\" & sArchivo
  End If
  'Llama función del Kernel
  iLeidos = GetPrivateProfileString(sCabecera, sItem, "", sRetorno, 255, sPathHide)
  'Valida lectura
  If iLeidos = 0 Then
    LeeHide = sDefecto
  Else
    LeeHide = Left$(sRetorno, iLeidos)
  End If
  Exit Function
ErrorLeeHide:
  MsgBox "Imposible leer librería de funciones.", vbCritical, "Error de lectura"
  Resume Next
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Letras(ByVal lNro As Long, Optional ByVal sGenero As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Convierte una cantidad numérica a letras.
'Parámetros  :
'     I        lNro, cantidad comprendida entre 0 y 999.999.999.999.999 que se
'              desea convertir a letras.
'     I        sGenero, caracter que puede ser "M"asculino o "F"emenino
'     O        cantidad expresada en letras.
'
'Declaración de variables locales
  Dim sEnLetras As String
  Dim iGrupo As Integer
  Dim sNumStr As String
  Dim sUnidad As String
  Dim sDecena As String
  Dim sCentena As String
  Dim i As Integer
  'Declaración de arreglos locales
  Dim asGrupos(5) As String
  Dim asUnidad(10) As String
  Dim asDecena(10, 10) As String
  Dim asCentena(10) As String
  Dim asConector(5) As String
  'Declaración de constantes
  Const y = "y "
  Const m = "os "
  Const f = "as "
  Const MIL = "mil "
  Const MILLON = "millón "
  Const MILLONES = "millones "
  Const BILLON = "billón "
  Const BILLONES = "billones "
  'Control de parámetros opcionales
  If sGenero = "" Then
    sGenero = "M"
  Else
    sGenero = UCase$(sGenero)
  End If
  'Control del tipo de parámetros
  If VarType(lNro) <> vbLong Then
    Exit Function
  End If
  If VarType(sGenero) <> vbString Then
    Exit Function
  End If
  'Elementos constantes
  asDecena(1, 1) = ""
  asDecena(2, 2) = "once "
  asDecena(2, 3) = "doce "
  asDecena(2, 4) = "trece "
  asDecena(2, 5) = "catorce "
  asDecena(2, 6) = "quince "
  asDecena(2, 7) = "dieciseis "
  asDecena(2, 8) = "diecisiete "
  asDecena(2, 9) = "dieciocho "
  asDecena(2, 10) = "diecinueve "
  asCentena(1) = ""
  asConector(1) = ""
  'Convertir el nro. a caracteres justificando con 0 a la izq.
  sNumStr = Trim$(CStr(lNro))
  sNumStr = String$(15 - Len(sNumStr), "0") & sNumStr
  'Confección de los grupos
  For iGrupo = 1 To 5
    asGrupos(5 - iGrupo + 1) = Mid$(sNumStr, (iGrupo - 1) * 3 + 1, 3)
  Next iGrupo
  'Proceso
  sEnLetras = ""
  For iGrupo = 5 To 1 Step -1
    'Extraer c/u de las 3 cifras del grupo actual
    sUnidad = Right$(asGrupos(iGrupo), 1)
    sDecena = Mid$(asGrupos(iGrupo), 2, 1)
    sCentena = Left$(asGrupos(iGrupo), 1)
    'Cargar unidades
    If lNro = 0 And iGrupo = 1 Then
      asUnidad(1) = "cero"
    Else
      asUnidad(1) = ""
    End If
    If sDecena = "1" Then
      asUnidad(2) = asDecena(2, CInt(sUnidad) + 1)
    Else
      If asGrupos(iGrupo) = "001" And (iGrupo = 2 Or iGrupo = 4) Then
        asUnidad(2) = ""
      Else
        If iGrupo > 2 Then
          asUnidad(2) = "un "
        Else
          If sGenero = "M" Then
            asUnidad(2) = "un "
          Else
            asUnidad(2) = "una "
          End If
        End If
      End If
    End If
    If sDecena = "1" Then
      asUnidad(3) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(3) = "dos "
    End If
    If sDecena = "1" Then
      asUnidad(4) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(4) = "tres "
    End If
    If sDecena = "1" Then
      asUnidad(5) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(5) = "cuatro "
    End If
    If sDecena = "1" Then
      asUnidad(6) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(6) = "cinco "
    End If
    If sDecena = "1" Then
      asUnidad(7) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(7) = "seis "
    End If
    If sDecena = "1" Then
      asUnidad(8) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(8) = "siete "
    End If
    If sDecena = "1" Then
      asUnidad(9) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(9) = "ocho "
    End If
    If sDecena = "1" Then
      asUnidad(10) = asDecena(2, CInt(sUnidad) + 1)
    Else
      asUnidad(10) = "nueve "
    End If
    'Cargar decenas
    If sUnidad = "0" Then
      asDecena(2, 1) = "diez "
    Else
      asDecena(2, 1) = ""
    End If
    If sUnidad = "0" Then
      asDecena(3, 1) = "veinte "
    Else
      asDecena(3, 1) = "veinti"
    End If
    If sUnidad <> "0" Then
      asDecena(4, 1) = "treinta " & y
    Else
      asDecena(4, 1) = "treinta "
    End If
    If sUnidad <> "0" Then
      asDecena(5, 1) = "cuarenta " & y
    Else
      asDecena(5, 1) = "cuarenta "
    End If
    If sUnidad <> "0" Then
      asDecena(6, 1) = "cincuenta " & y
    Else
      asDecena(6, 1) = "cincuenta "
    End If
    If sUnidad <> "0" Then
      asDecena(7, 1) = "sesenta " & y
    Else
      asDecena(7, 1) = "sesenta "
    End If
    If sUnidad <> "0" Then
      asDecena(8, 1) = "setenta " & y
    Else
      asDecena(8, 1) = "setenta "
    End If
    If sUnidad <> "0" Then
      asDecena(9, 1) = "ochenta " & y
    Else
      asDecena(9, 1) = "ochenta "
    End If
    If sUnidad <> "0" Then
      asDecena(10, 1) = "noventa " & y
    Else
      asDecena(10, 1) = "noventa "
    End If
    'Cargar centenas
    If sDecena & sUnidad = "00" Then
      asCentena(2) = "cien "
    Else
      asCentena(2) = "ciento "
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(3) = "doscient" & m
    Else
      asCentena(3) = "doscient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(4) = "trescient" & m
    Else
      asCentena(4) = "trescient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(5) = "cuatrocient" & m
    Else
      asCentena(5) = "cuatrocient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(6) = "quinient" & m
    Else
      asCentena(6) = "quinient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(7) = "seiscient" & m
    Else
      asCentena(7) = "seiscient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(8) = "setecient" & m
    Else
      asCentena(8) = "setecient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(9) = "ochocient" & m
    Else
      asCentena(9) = "ochocient" & f
    End If
    If iGrupo < 3 And sGenero = "M" Then
      asCentena(10) = "novecient" & m
    Else
      asCentena(10) = "novecient" & f
    End If
    'Cargar conectores
    If CInt(asGrupos(2)) > CInt("000") Then
      asConector(2) = MIL
    Else
      asConector(2) = ""
    End If
    If CInt(asGrupos(3)) > CInt("000") Or CInt(asGrupos(4)) > CInt("000") Then
      If asGrupos(3) = "001" Then
        asConector(3) = MILLON
      Else
        asConector(3) = MILLONES
      End If
    Else
      asConector(3) = ""
    End If
    If CInt(asGrupos(4)) > CInt("000") Then
      asConector(4) = MIL
    Else
      asConector(4) = ""
    End If
    If CInt(asGrupos(5)) > CInt("000") Then
      If asGrupos(5) = "001" Then
        asConector(5) = BILLON
      Else
        asConector(5) = BILLONES
      End If
    Else
      asConector(5) = ""
    End If
    'Componer la cifra en letras del grupo actual
    sEnLetras = sEnLetras & asCentena(CInt(sCentena) + 1) & _
                asDecena(CInt(sDecena) + 1, 1) & _
                asUnidad(CInt(sUnidad) + 1) & asConector(iGrupo)
  Next iGrupo
  'Retorna resultado
  Letras = sEnLetras
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Extrae(ByVal sCadena As String, ByVal sCaracter As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Extrae todas las ocurrencias de sCaracter en sCadena.
'Parámetros  :
'     I        sCadena, cadena a extaer caracteres sCaracter.
'     I        sCaracter, caracter a extraer.
'     O        cadena sin ocurrencias de sCaracter.
'
'Declara variables locales
  Dim i As Integer
  Dim sSalida As String
  'Inicializa salida
  sSalida = ""
  'Obtiene caracteres cadena
  For i = 1 To Len(sCadena)
    'Verifica caracter
    If Not Mid(sCadena, i, 1) = sCaracter Then
      sSalida = sSalida & Mid(sCadena, i, 1)
    End If
  Next i
  'Retorna string
  Extrae = sSalida
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ExisteArchivo(ByVal sPath As String) As Boolean
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Determina si existe archivo.
'Parámetros  :
'    I        sPath, ruta completa archivo.
'    O        True si existe archivo. False en cualquier otro caso.
'
'Declara variables locales
  Dim x As Integer
  'En caso de error continua
  On Error Resume Next
  'Obtiene nro. archivo disponible
  x = FreeFile
  'Abre archivo
  Open sPath For Input As x
  'Valida tipo error
  If Err = 0 Then
    ExisteArchivo = True
  Else
    ExisteArchivo = False
  End If
  'Cierra archivo
  Close x
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function DigitoVerificador(ByVal sCarnet As String) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene dígito verificador.
'Parámetros  :
'     I        sCarnet, rut a obtener dígito verificador.
'     O        Dígito verificador. En otro caso, retorna "E".
'
'Declara variables locales
  Dim sCte As String
  Dim sRut As String
  Dim iSuma As Integer
  Dim iDigito As Integer
  Dim i As Integer
  'Controla Error
  On Error GoTo DigitoVerificador_Err
  'Crea string constante
  sCte = "32765432"
  'Extrae -DV
  sRut = Left$(sCarnet, InStr(1, sCarnet, "-", vbTextCompare) - 1)
  'Setea rut
  sRut = Format(sRut, "00000000")
  'Obtiene suma dígitos
  iSuma = 0
  For i = 1 To 8
    iSuma = iSuma + (Val(Mid$(sCte, i, 1)) * Val(Mid$(sRut, i, 1)))
  Next
  'Obtiene dígito
  iDigito = 11 - (iSuma Mod 11)
  Select Case iDigito
  Case 10
    DigitoVerificador = "K"
  Case 11
    DigitoVerificador = "0"
  Case Else
    DigitoVerificador = Trim$(CStr(iDigito))
  End Select
  Exit Function
DigitoVerificador_Err:
  DigitoVerificador = "E"
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function ValorSQL(ByVal vValor As Variant) As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 10/2000
'Descripción : Prepara cantidad númerica para instrucción SQL.
'Parámetros  :
'     I        vValor, valor numérico.
'     O        cadena con vValor "." como separador decimal.
'
'Declara variables locales
  Dim sValor As String
  Dim sVal As String
  Dim iLen As Integer
  Dim iCont As Integer
  Const COMODIN = "&"
  'Cambia tipo dato
  sValor = CStr(vValor)
  'Obtiene largo cadena
  iLen = Len(sValor)
  'Valida parámetro
  If iLen = 0 Then Exit Function
  'Recorre cadena
  For iCont = 1 To iLen
    If Mid$(sValor, iCont, 1) = "," Then
      Mid$(sValor, iCont) = "."
    ElseIf Mid$(sValor, iCont, 1) = "." Then
      Mid$(sValor, iCont) = COMODIN
    End If
  Next iCont
  'Extrae comodin
  For iCont = 1 To iLen
    If Not Mid$(sValor, iCont, 1) = COMODIN Then
      sVal = sVal & Mid$(sValor, iCont, 1)
    End If
  Next iCont
  'Retorna valor
  ValorSQL = sVal
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function WinTempDir() As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene directorio temporal de Windows.
'Parámetros  :
'     O        Path temporal de Windows. "C:\Windows\Temp\".
'
'Declara variables locales
  Dim sTemp As String
  Dim iRet As Integer
  'Directorio Temporal de Windows
  sTemp = Space(256)
  iRet = GetTempPath(256, sTemp)
  WinTempDir = Left(sTemp, iRet)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function WinDir() As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene directorio Windows.
'Parámetros  :
'     O        Path de Windows. "C:\Windows\".
'
'Declara variables locales
  Dim sWin As String
  Dim iRet As Integer
  'Directorio Windows
  sWin = Space(256)
  iRet = GetWindowsDirectory(sWin, 256)
  WinDir = Left(sWin, iRet)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function UserIdRed() As String
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Obtiene nombre usuario red.
'Parámetros  :
'     O        Nombre usuario red.
'
'Declara variables locales
  Dim sName As String
  Dim iRet As Integer
  'Perfil usuario red
  sName = Space(256)
  iRet = GetUserName(sName, 256)
  UserIdRed = Left$(sName, Len(Trim$(sName)) - 1)
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function MilesComa() As Boolean
'
'Autor       : Patricio Dahmen G.
'Versión     : 09/1999
'Descripción : Determina si separador de miles es coma.
'Parámetros  :
'    O         Verdadero si separador de miles es coma (","). Falso en cualquier otro caso.
'
'Declara variables locales
  Dim sPath As String
  Dim c As String
  'Obtiene directorio windows
  sPath = WinDir
  'Lee windows.ini
  If CStr(LeeIni(sPath, "win", "intl", "sThousand", "nil")) = "," Then
    MilesComa = True
  Else
    MilesComa = False
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function Cancelar() As Boolean
  If TypeOf Screen.ActiveControl Is CommandButton Then
    Cancelar = (Screen.ActiveControl.Name = "cmdCancelar")
  End If
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Private Function MsjExcel(ByVal iMensaje As Integer) As String
  Dim x As String
  Select Case iMensaje
  Case 0
    x = TIT_EXCEL
  Case 1
    x = DIA_EXCEL
  Case 2
    x = LBL_EXCEL
  Case 3
    x = MSG_EXCEL
  Case Is < 0
    x = MSG_NIL
  Case Else
    x = "Número de mensaje desconocido."
  End Select
  MsjExcel = x
End Function
' Fin de la Función **********************************************

' .:. Nombre de la Función .:.
Public Function SQLver(ByVal sCampos As String, ByVal sTablas As String, Optional ByVal sWhere As String, Optional ByVal sOrderBy As String, Optional ByVal bDistinct As Boolean, Optional ByVal sGroupBy As String, Optional ByVal bValues As Boolean, Optional ByVal sValores As String) As String
  Dim sTexto As String
  If bValues Then
    sTexto = "INSERT INTO " & sTablas & "(" & vbCrLf
    sTexto = sTexto & sCampos & ")" & vbCrLf
    sTexto = sTexto & "VALUES " & "(" & vbCrLf
    sTexto = sTexto & sValores & ")" & vbCrLf
    Exit Function
  End If
  If bDistinct Then
    sTexto = "SELECT DISTINCT " & vbCrLf
  Else
    sTexto = "SELECT " & vbCrLf
  End If
  sTexto = sTexto & sCampos & vbCrLf
  sTexto = sTexto & "FROM " & vbCrLf
  sTexto = sTexto & sTablas & vbCrLf
  If sWhere <> "" Then
    sTexto = sTexto & "WHERE " & vbCrLf
    sTexto = sTexto & sWhere & vbCrLf
  End If
  If sGroupBy <> "" Then
    sTexto = sTexto & "GROUP BY " & vbCrLf
    sTexto = sTexto & sGroupBy & vbCrLf
  End If
  If sOrderBy <> "" Then
    sTexto = sTexto & "ORDER BY " & vbCrLf
    sTexto = sTexto & sOrderBy & vbCrLf
  End If
  SQLver = sTexto
End Function
' Fin de la Función **********************************************


'   *   **   ***   ****  PROPIEDADES   ****   ***   **   *
' .:. Nombre de la Propiedad .:.
Public Property Get SepDecimal() As String
  SepDecimal = SEP_DEC
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Let SepDecimal(ByVal sNewValue As String)
  SEP_DEC = sNewValue
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Get MensajeExcel(ByVal iNro As Integer) As String
  MensajeExcel = MsjExcel(iNro)
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Get Separacion() As Integer
  If iSepara = 0 Then iSepara = 480
  Separacion = iSepara
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Let Separacion(ByVal iValor As Integer)
  iSepara = iValor
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Get Resolucion() As String
  Dim h As Integer
  Dim w As Integer
  Pantalla h, w
  Resolucion = CStr(h) & "X" & CStr(w)
End Property
' Fin de la Propiedad ********************************************

' .:. Nombre de la Propiedad .:.
Public Property Get PantallaXP() As Boolean
  Dim h As Integer
  Dim w As Integer
  Dim bXP As Boolean
  Pantalla h, w
  If h > 600 Then bXP = True
  If w > 800 Then bXP = True
  PantallaXP = bXP
End Property
' Fin de la Propiedad ********************************************




