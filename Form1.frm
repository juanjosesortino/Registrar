VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Componentes"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpieza 
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   525
      Left            =   4980
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Limpieza Total"
      Top             =   450
      Width           =   585
   End
   Begin VB.CommandButton cmdPaquete 
      CausesValidation=   0   'False
      Height          =   525
      Left            =   4320
      Picture         =   "Form1.frx":0F4C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cliente"
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdCliente 
      CausesValidation=   0   'False
      Height          =   525
      Left            =   3660
      Picture         =   "Form1.frx":14F4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cliente"
      Top             =   450
      Width           =   615
   End
   Begin VB.CommandButton cmdReRegistra 
      CausesValidation=   0   'False
      Height          =   525
      Left            =   3000
      Picture         =   "Form1.frx":1CE5
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Completo"
      Top             =   450
      Width           =   615
   End
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   990
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdProduccion 
      Caption         =   "ALG02 --> ALG01"
      Height          =   525
      Left            =   1530
      TabIndex        =   2
      Top             =   450
      Width           =   1425
   End
   Begin VB.CommandButton cmdTesting 
      Caption         =   "ALG01 --> ALG02"
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   1425
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   315
      Left            =   6390
      TabIndex        =   4
      Top             =   420
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   556
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6510
      TabIndex        =   5
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Gamma                                                Registrar Componentes"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6285
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const TESTING      As String = "ALG01"
Private Const PRODUCCION   As String = "PC5"

Private Const STB_PANEL1   As Integer = 1
Private Const STB_PANEL2   As Integer = 2

Dim x                      As Long
Dim ix                     As Long
Dim aParametros()          As String
Dim bConApplicationServer  As String

'LimpiarRegistro
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const Delete = &H10000
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private WithEvents cReg    As cRegSearch
Attribute cReg.VB_VarHelpID = -1
Dim objReg                 As Object
Dim ArrSubKeys()
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'LimpiarRegistro

Private Sub Form_Load()

   Dim parametros As String
   
   With ListView1
        .View = lvwReport
        .ColumnHeaders.Add , , "Encontrado en:"
        .ColumnHeaders.Add , , "RootKey"
        .ColumnHeaders.Add , , "SubKey"
        .ColumnHeaders.Add , , "Nivel"
        .ColumnHeaders.Add , , "Del Niv."
        .ColumnHeaders.Add , , "Value"
        .Visible = False
   End With

   GetApplicationServer
   
   cmdCliente.ToolTipText = "Actualizar Cliente " & Label1.Caption
   cmdPaquete.ToolTipText = "Actualizar Paquete " & Label1.Caption
   cmdReRegistra.ToolTipText = "Actualización Completa " & Label1.Caption
   
   cmdTesting.Caption = TESTING & " --> " & PRODUCCION
   cmdProduccion.Caption = PRODUCCION & " --> " & TESTING
   
   'Parametros
   parametros = Command$

   aParametros = Split(parametros, " ")
   For ix = LBound(aParametros) To UBound(aParametros)
      If ix = 1 Then
         If aParametros(ix) = "/?" Then
            MsgBox "<Registrar> registra los componentes al servidor elegido." & _
            " " & _
            "Uso: Registrar [Servidor] [-s silencioso] [/? Muetra Parametros]"
         End If
      End If
      If ix = 2 Then
         If aParametros(ix) <> "-s" Then
         End If
      End If
   Next ix
   
End Sub

Private Sub cmdProduccion_Click()
   cmdProduccion.Enabled = False
   Proceso PRODUCCION, TESTING
   GetApplicationServer
End Sub

Private Sub cmdTesting_Click()
   cmdTesting.Enabled = False
   Proceso TESTING, PRODUCCION
   GetApplicationServer
End Sub

Private Sub cmdReRegistra_Click()
   cmdTesting.Enabled = False
   cmdProduccion.Enabled = False
   cmdReRegistra.Enabled = False
   cmdPaquete.Enabled = False
   cmdCliente.Enabled = False
   cmdLimpieza.Enabled = False
   
   If Label1.Caption = PRODUCCION Then
      ReRegistra PRODUCCION
   Else
      ReRegistra TESTING
   End If

   GetApplicationServer
End Sub

Private Sub cmdLimpieza_Click()
   cmdTesting.Enabled = False
   cmdProduccion.Enabled = False
   cmdReRegistra.Enabled = False
   cmdPaquete.Enabled = False
   cmdCliente.Enabled = False
   cmdLimpieza.Enabled = False

   If Label1.Caption = PRODUCCION Then
      ReRegistra PRODUCCION, True
   Else
      ReRegistra TESTING, True
   End If

   GetApplicationServer
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Proceso Terminado"
   Screen.MousePointer = vbNormal
End Sub

Private Sub Proceso(strSaca As String, strPone As String)

   Dim strPid As String
   Dim objInstaler As Object
   
   On Error GoTo GestErr
   
   Screen.MousePointer = vbArrowHourglass

   BorrarProxy
   
   stb1.Panels(STB_PANEL1).Text = "Desregistrando Cliente " & strSaca
   Desregistrar_Cliente "BOGeneral.dll", strSaca
   Desregistrar_Cliente "BOFiscal.dll", strSaca
   Desregistrar_Cliente "BOCereales.dll", strSaca
   Desregistrar_Cliente "BOContabilidad.dll", strSaca
   Desregistrar_Cliente "BOGescom.dll", strSaca
   Desregistrar_Cliente "BOSeguridad.dll", strSaca
   Desregistrar_Cliente "Cereales.dll", strSaca
   Desregistrar_Cliente "Fiscal.dll", strSaca
   Desregistrar_Cliente "ReportsCereales.dll", strSaca
   Desregistrar_Cliente "ReportsCereales2.dll", strSaca
   Desregistrar_Cliente "ReportsGescom.dll", strSaca
'   If strSaca <> PRODUCCION Then
      Desregistrar_Cliente "ReportsGescom2.dll", strSaca
'   End If
   Desregistrar_Cliente "GestionComercial.dll", strSaca
   Desregistrar_Cliente "Contabilidad.dll", strSaca
   Desregistrar_Cliente "AdministradorGeneral.dll", strSaca
   Desregistrar_Cliente "Seguridad.dll", strSaca
   Desregistrar_Cliente "PowerMaskControl.ocx", strSaca
   Desregistrar_Cliente "ALGControls.ocx", strSaca
   Desregistrar_Cliente "AlgStdFunc.dll", strSaca
   Desregistrar_Cliente "BOProduccion.dll", strSaca
   Desregistrar_Cliente "Produccion.dll", strSaca

   LimpiarRegistro HKEY_CLASSES_ROOT
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Instalando Paquete Application Server de " & strPone
   WaitForShelledApp "msiexec /i " & """" & "\\" & strPone & "\d\Algoritmo\Paquete Application Server\Algoritmo.MSI" & """" & " /qn" '/qn: es la forma silenciosa
   
   stb1.Panels(STB_PANEL1).Text = "Registrando Cliente " & strPone
   Registrar_Cliente "BOGeneral.dll", strPone
   Registrar_Cliente "BOFiscal.dll", strPone
   Registrar_Cliente "BOCereales.dll", strPone
   Registrar_Cliente "BOContabilidad.dll", strPone
   Registrar_Cliente "BOGescom.dll", strPone
   Registrar_Cliente "BOSeguridad.dll", strPone
   Registrar_Cliente "Cereales.dll", strPone
   Registrar_Cliente "Fiscal.dll", strPone
   Registrar_Cliente "ReportsCereales.dll", strPone
   Registrar_Cliente "ReportsCereales2.dll", strPone
   Registrar_Cliente "ReportsGescom.dll", strPone
'   If strPone <> PRODUCCION Then
      Registrar_Cliente "ReportsGescom2.dll", strPone
'   End If
   Registrar_Cliente "GestionComercial.dll", strPone
   Registrar_Cliente "Contabilidad.dll", strPone
   Registrar_Cliente "AdministradorGeneral.dll", strPone
   Registrar_Cliente "Seguridad.dll", strPone
   Registrar_Cliente "PowerMaskControl.ocx", strPone
   Registrar_Cliente "ALGControls.ocx", strPone
   Registrar_Cliente "AlgStdFunc.dll", strPone
   Registrar_Cliente "BOProduccion.dll", strPone
   Registrar_Cliente "Produccion.dll", strPone
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Proceso Terminado"
   Screen.MousePointer = vbNormal
   
   Exit Sub
   
GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[Proceso]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub ReRegistra(strServer As String, Optional ByVal bLimpieza As Boolean)
   
   Dim strPid          As String
   Dim objInstaler     As Object
   Dim strIdAplicacion As String
   
   On Error GoTo GestErr
   
   Screen.MousePointer = vbArrowHourglass
   
   BorrarProxy
   
   stb1.Panels(STB_PANEL1).Text = "Desregistrando Cliente " & strServer
   Desregistrar_Cliente "BOGeneral.dll", strServer
   Desregistrar_Cliente "BOFiscal.dll", strServer
   Desregistrar_Cliente "BOCereales.dll", strServer
   Desregistrar_Cliente "BOContabilidad.dll", strServer
   Desregistrar_Cliente "BOGescom.dll", strServer
   Desregistrar_Cliente "BOSeguridad.dll", strServer
   Desregistrar_Cliente "BOProduccion.dll", strServer
   Desregistrar_Cliente "Cereales.dll", strServer
   Desregistrar_Cliente "Fiscal.dll", strServer
   Desregistrar_Cliente "ReportsCereales.dll", strServer
   Desregistrar_Cliente "ReportsCereales2.dll", strServer
   Desregistrar_Cliente "ReportsGescom.dll", strServer
   Desregistrar_Cliente "ReportsGescom2.dll", strServer
   Desregistrar_Cliente "GestionComercial.dll", strServer
   Desregistrar_Cliente "Contabilidad.dll", strServer
   Desregistrar_Cliente "AdministradorGeneral.dll", strServer
   Desregistrar_Cliente "Produccion.dll", strServer
   Desregistrar_Cliente "Seguridad.dll", strServer
   Desregistrar_Cliente "PowerMaskControl.ocx", strServer
   Desregistrar_Cliente "ALGControls.ocx", strServer
   Desregistrar_Cliente "AlgStdFunc.dll", strServer

   LimpiarRegistro HKEY_CLASSES_ROOT
   
   If bLimpieza Then Exit Sub
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Instalando Paquete Application Server de " & strServer
   WaitForShelledApp "msiexec /i " & """" & "\\" & strServer & "\d\Algoritmo\Paquete Application Server\Algoritmo.MSI" & """" & " /qn" '/qn: es la forma silenciosa
   
   stb1.Panels(STB_PANEL1).Text = "Registrando Cliente " & strServer
   Registrar_Cliente "BOGeneral.dll", strServer
   Registrar_Cliente "BOFiscal.dll", strServer
   Registrar_Cliente "BOCereales.dll", strServer
   Registrar_Cliente "BOContabilidad.dll", strServer
   Registrar_Cliente "BOGescom.dll", strServer
   Registrar_Cliente "BOSeguridad.dll", strServer
   Registrar_Cliente "BOProduccion.dll", strServer
   Registrar_Cliente "Cereales.dll", strServer
   Registrar_Cliente "Fiscal.dll", strServer
   Registrar_Cliente "ReportsCereales.dll", strServer
   Registrar_Cliente "ReportsCereales2.dll", strServer
   Registrar_Cliente "ReportsGescom.dll", strServer
   Registrar_Cliente "ReportsGescom2.dll", strServer
   Registrar_Cliente "GestionComercial.dll", strServer
   Registrar_Cliente "Contabilidad.dll", strServer
   Registrar_Cliente "AdministradorGeneral.dll", strServer
   Registrar_Cliente "Produccion.dll", strServer
   Registrar_Cliente "Seguridad.dll", strServer
   Registrar_Cliente "PowerMaskControl.ocx", strServer
   Registrar_Cliente "ALGControls.ocx", strServer
   Registrar_Cliente "AlgStdFunc.dll", strServer
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Proceso Terminado"
   Screen.MousePointer = vbNormal
   
   Exit Sub
   
GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[ReRegistra]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub WaitForShelledApp(strproceso As String)
Dim ProcessId As Long
Dim hProcess As Long
Dim exitCode As Long

   ProcessId = Shell(strproceso, vbHide)
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
   Do
      DoEvents
      Call GetExitCodeProcess(hProcess, exitCode)
   Loop While exitCode > 0
   CloseHandle hProcess
End Sub
Private Sub Desregistrar_Cliente(strDll As String, strSaca As String)
   stb1.Panels(STB_PANEL2).Text = strDll
   WaitForShelledApp "regsvr32 /u /s " & """" & "\\" & strSaca & "\D\Algoritmo\Componentes Client\" & strDll & """"
End Sub

Private Sub Registrar_Cliente(strDll As String, strPone As String)
   stb1.Panels(STB_PANEL2).Text = strDll
   WaitForShelledApp "regsvr32 /s " & """" & "\\" & strPone & "\D\Algoritmo\Componentes Client\" & strDll & """"
End Sub
Private Sub LimpiarRegistro(RootKeys As ROOT_KEYS)
         Dim itmX As ListItem
         Dim aKey() As String
         Dim sKeyToDel As String
         Dim Nivel As Integer
   
10       On Error GoTo GestErr
   
         'Busco y lleno el listview
20       Set cReg = New cRegSearch
30       cReg.RootKey = RootKeys
40       cReg.SubKey = "TypeLib"
50       cReg.SearchFlags = 3
60       cReg.SearchString = "Algoritmo"
70       stb1.Panels(STB_PANEL1).Text = "Buscando y Eliminando Claves del Registro"
         stb1.Panels(STB_PANEL2).Text = ""
80       cReg.DoSearch
90       Set cReg = Nothing
   
         'Selecciono nivel 2
100      For Each itmX In ListView1.ListItems
110         If itmX.SubItems(3) >= 2 Then ' Nivel 2
120            itmX.SubItems(4) = 2
130         End If
140         DoEvents
150      Next itmX

         'Borro las claves
160      Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Machine & "\root\default:StdRegProv")
170      For Each itmX In ListView1.ListItems
180         aKey = Split(itmX.SubItems(2), "\")
     
190         sKeyToDel = ""
200         If itmX.SubItems(4) <> "" Then
210            Nivel = Val(itmX.SubItems(4)) - 1
220         End If
230         If Nivel > 0 Then
240            For ix = LBound(aKey) To Nivel
250               sKeyToDel = sKeyToDel & aKey(ix) & "\"
260            Next ix
               'es la clave que se quiere borrar
270            sKeyToDels = Left(sKeyToDel, Len(sKeyToDel) - 1)

280            On Error Resume Next
290            objReg.EnumKey HKEY_CLASSES_ROOT, sKeyToDels, ArrSubKeys

300            If IsArrayEmpty(ArrSubKeys) Then
310               objReg.DeleteKey HKEY_CLASSES_ROOT, sKeyToDels
'320               ListView1.ListItems.Remove itmX.Index
330               stb1.Panels(STB_PANEL2).Text = itmX.SubItems(5)
340            Else
350               For Each SubKey In ArrSubKeys
360                  BorrarRama sKeyToDels & "\" & SubKey
370               Next SubKey
380               objReg.DeleteKey HKEY_CLASSES_ROOT, sKeyToDels
390               stb1.Panels(STB_PANEL2).Text = itmX.SubItems(5)
400            End If
410         End If
420         DoEvents
430      Next itmX
   
440      Exit Sub
   
GestErr:
450      Me.MousePointer = vbNormal
460      MsgBox "[ReRegistra]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
   Dim lvItm As ListItem
   Dim sTemp As String
   
   Select Case lFound
         Case FOUND_IN_KEY_NAME
              sTemp = "KEY_NAME"
         Case FOUND_IN_VALUE_NAME
              sTemp = "VALUE NAME"
         Case FOUND_IN_VALUE_VALUE
              sTemp = "VALUE VALUE"
   End Select
   
   With ListView1
       Set lvItm = .ListItems.Add(, , sTemp)
       lvItm.SubItems(1) = sRootKey
       lvItm.SubItems(2) = sKey
       lvItm.SubItems(3) = InStrCount(sKey, "\")
       lvItm.SubItems(5) = sValue
       '.Visible = True
   End With
   Set lvItm = Nothing
End Sub

Private Function InStrCount(ByVal Source As String, Search As String) As Long
   'devuelve el número de ocurrencias de un substring dentro de un string
   InStrCount = Len(Source) - Len(Replace(Source, Search, Mid(Search, 2)))
End Function
Private Function Machine() As String
   Dim s As String, c As Long
   c = 16: s = String$(16, 0)
   If GetComputerName(s, c) Then Machine = Left$(s, c)
End Function
Private Function IsArrayEmpty(vArray As Variant) As Boolean
   '  establece si el arreglo esta vacio o contiene almenos un elemento
   IsArrayEmpty = (ArraySize(vArray) = 0)
End Function
Private Sub BorrarRama(ByVal Key As String)
'Dim ArrSubKeys()

   objReg.EnumKey HKEY_CLASSES_ROOT, Key, ArrSubKeys
   
   If IsArrayEmpty(ArrSubKeys) Then
      objReg.DeleteKey HKEY_CLASSES_ROOT, Key
   Else
      For Each SubKey In ArrSubKeys
         BorrarRama Key & "\" & SubKey
      Next SubKey
      objReg.DeleteKey HKEY_CLASSES_ROOT, Key
   End If

End Sub

Private Function ArraySize(vArray As Variant, Optional ByVal Dimension As Long = 1) As Long
   ' Dimension es la dimension del arreglo del cual deseo obtener su tamaño. Si se omite asume 1
   ' Devuelve el numero de elementos de la dimension indicada
   
   On Error GoTo ArrayEmpty
   If Not IsArray(vArray) Then Err.Raise Err.Number, Err.Source, Err.Description

   ArraySize = 1 + UBound(vArray, Dimension) - LBound(vArray, Dimension)
ArrayEmpty:
End Function
Private Sub GetApplicationServer()
Dim oCatalog      As Object
Dim oApplications As Object
Dim oApplication  As Object
Dim ix            As Integer
Dim GetApplicationServer As String

   bConApplicationServer = False
    
   Set oCatalog = CreateObject("COMAdmin.COMAdminCatalog")
   Set oApplications = oCatalog.GetCollection("Applications")
  
   oApplications.Populate
   ix = 0
  
   Do Until ix = oApplications.Count
  
      Set oApplication = oApplications.Item(ix)
  
      If UCase(oApplication.Name) = "ALGORITMO" Then
  
         GetApplicationServer = oApplication.Value("ApplicationProxyServerName")
         Exit Do
      
      End If
      
      ix = ix + 1
     
   Loop
   
   If Len(GetApplicationServer) = 0 Then
      stb1.Panels(STB_PANEL1).Text = "Sin Application Server"
      bConApplicationServer = False
   Else
      bConApplicationServer = True
   End If
   
   Label1.Caption = IIf(Len(GetApplicationServer) = 0, TESTING, GetApplicationServer)
   cmdReRegistra.Caption = IIf(Len(GetApplicationServer) = 0, TESTING, GetApplicationServer)
   
   If Label1.Caption = TESTING Then
      cmdTesting.Enabled = True
      cmdProduccion.Enabled = False
   Else
      cmdTesting.Enabled = False
      cmdProduccion.Enabled = True
   End If
   cmdReRegistra.Enabled = True
   cmdCliente.Enabled = True
   cmdPaquete.Enabled = True
   cmdLimpieza.Enabled = True
   
End Sub
Private Sub BorrarProxy()

10       On Error GoTo GestErr
   
20       If bConApplicationServer Then
30          strIdAplicacion = IdAplicacion
      
40          Set objInstaler = CreateObject("WindowsInstaller.Installer")
50          For Each prod In objInstaler.ProductsEx("", "", 7)
60             If prod.InstallProperty("InstalledProductName") = "Algoritmo (Application Proxy)" Then
70               strPid = prod.ProductCode
80               Exit For
90            End If
100         Next
110         Set objInstaler = Nothing

120         stb1.Panels(STB_PANEL2).Text = ""
130         stb1.Panels(STB_PANEL1).Text = "Desinstalando Paquete Application Server"
      
140         WaitForShelledApp "msiexec /x " & strPid & " /qn" '/qn: es la forma silenciosa
      
150         BorrarCarpeta "C:\Archivos de programa\ComPlus Applications\" & strIdAplicacion
160      End If
   
170      Exit Sub
   
GestErr:
180      Me.MousePointer = vbNormal
190      MsgBox "[BorrarProxy]" & vbCrLf & Err.Description & Erl
End Sub

Private Function IdAplicacion() As String
   
         Dim objCatlog           As Object
         Dim objApplications     As Object
         Dim objApplication      As Object
         Dim objComponents       As Object
   
10       On Error GoTo GestErr
   
20       Set objCatlog = CreateObject("COMAdmin.COMAdminCatalog")
30       objCatlog.Connect ("")
40       Set objApplications = objCatlog.GetCollection("Applications")
50       objApplications.Populate
60       For Each objApplication In objApplications
70           Set objComponents = objApplications.GetCollection("Components", objApplication.Key)
80           objComponents.Populate

90           If objApplication.Name = "Algoritmo" Then
100            IdAplicacion = objApplication.Key
110          End If
120      Next
   
130      Set objCatlog = Nothing
140      Set objApplications = Nothing
150      Set objApplication = Nothing
160      Set objComponents = Nothing
   
170      Exit Function

GestErr:
180      Set objCatlog = Nothing
190      Set objApplications = Nothing
200      Set objApplication = Nothing
210      Set objComponents = Nothing
   
220      Me.MousePointer = vbNormal
230      MsgBox "[IdAplicacion]" & vbCrLf & Err.Description & Erl
End Function

Private Sub BorrarCarpeta(ByVal FullPath As String)
   
   On Error Resume Next
   
   Dim oFso As New Scripting.FileSystemObject

   If oFso.FolderExists(FullPath) Then
       'Setting the 2nd parameter to true forces deletion of read-only files
       oFso.DeleteFolder FullPath, True
   End If
   
   Set oFso = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub cmdCliente_Click()
   cmdTesting.Enabled = False
   cmdProduccion.Enabled = False
   cmdReRegistra.Enabled = False
   cmdCliente.Enabled = False
   cmdPaquete.Enabled = False
   cmdLimpieza.Enabled = False
   
   If Label1.Caption = PRODUCCION Then
      SoloCliente PRODUCCION
   Else
      SoloCliente TESTING
   End If

   GetApplicationServer
End Sub
Private Sub SoloCliente(strServer As String)
   
   Dim strPid          As String
   Dim objInstaler     As Object
   Dim strIdAplicacion As String
   
   On Error GoTo GestErr
   
   Screen.MousePointer = vbArrowHourglass
   
   stb1.Panels(STB_PANEL1).Text = "Registrando Cliente " & strServer
   Registrar_Cliente "BOGeneral.dll", strServer
   Registrar_Cliente "BOFiscal.dll", strServer
   Registrar_Cliente "BOCereales.dll", strServer
   Registrar_Cliente "BOContabilidad.dll", strServer
   Registrar_Cliente "BOGescom.dll", strServer
   Registrar_Cliente "BOSeguridad.dll", strServer
   Registrar_Cliente "BOProduccion.dll", strServer
   Registrar_Cliente "Cereales.dll", strServer
   Registrar_Cliente "Fiscal.dll", strServer
   Registrar_Cliente "ReportsCereales.dll", strServer
   Registrar_Cliente "ReportsCereales2.dll", strServer
   Registrar_Cliente "ReportsGescom.dll", strServer
   Registrar_Cliente "ReportsGescom2.dll", strServer
   Registrar_Cliente "GestionComercial.dll", strServer
   Registrar_Cliente "Contabilidad.dll", strServer
   Registrar_Cliente "AdministradorGeneral.dll", strServer
   Registrar_Cliente "Produccion.dll", strServer
   Registrar_Cliente "Seguridad.dll", strServer
   Registrar_Cliente "PowerMaskControl.ocx", strServer
   Registrar_Cliente "ALGControls.ocx", strServer
   Registrar_Cliente "AlgStdFunc.dll", strServer
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Proceso Terminado"
   Screen.MousePointer = vbNormal
   
   Exit Sub
   
GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[SoloCliente]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub cmdPaquete_Click()
   cmdTesting.Enabled = False
   cmdProduccion.Enabled = False
   cmdReRegistra.Enabled = False
   cmdCliente.Enabled = False
   cmdPaquete.Enabled = False
   cmdLimpieza.Enabled = False
   
   If Label1.Caption = PRODUCCION Then
      SoloPaquete PRODUCCION
   Else
      SoloPaquete TESTING
   End If

   GetApplicationServer
End Sub
Private Sub SoloPaquete(strServer As String)
   
   Dim strPid          As String
   Dim objInstaler     As Object
   Dim strIdAplicacion As String
   
   On Error GoTo GestErr
   
   Screen.MousePointer = vbArrowHourglass
   
   BorrarProxy
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Instalando Paquete Application Server de " & strServer
   WaitForShelledApp "msiexec /i " & """" & "\\" & strServer & "\d\Algoritmo\Paquete Application Server\Algoritmo.MSI" & """" & " /qn" '/qn: es la forma silenciosa
   
   stb1.Panels(STB_PANEL2).Text = ""
   stb1.Panels(STB_PANEL1).Text = "Proceso Terminado"
   Screen.MousePointer = vbNormal
   
   Exit Sub
   
GestErr:
   Me.MousePointer = vbNormal
   MsgBox "[SoloPaquete]" & vbCrLf & Err.Description & Erl
End Sub
