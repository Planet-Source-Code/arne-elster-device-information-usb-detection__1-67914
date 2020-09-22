VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Device Information"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdEject 
      Caption         =   "safe eject"
      Height          =   390
      Left            =   5100
      TabIndex        =   1
      Top             =   3150
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwDrives 
      Height          =   3015
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Drive"
         Object.Width           =   1041
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bus Type"
         Object.Width           =   1481
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Removable"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   5186
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iSubclass

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
Alias "GetLogicalDriveStringsA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String _
) As Long

Private m_clsSubcls As cSubclass

Private Sub cmdEject_Click()
    If EjectDevice(lvwDrives.SelectedItem.Tag) Then
        MsgBox "Successfully ejected the device from the system!", vbInformation
    Else
        MsgBox "Could not eject " & lvwDrives.SelectedItem.Tag & "!", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Set m_clsSubcls = New cSubclass
    
    m_clsSubcls.Subclass Me.hwnd, Me
    m_clsSubcls.AddMsg Me.hwnd, WM_DEVICECHANGE
    
    RefreshDriveList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_clsSubcls.Terminate
End Sub

Private Sub RefreshDriveList()
    Dim strDriveBuffer  As String
    Dim strDrives()     As String
    Dim i               As Long
    Dim udtInfo         As DEVICE_INFORMATION
    
    strDriveBuffer = Space(240)
    strDriveBuffer = Left$(strDriveBuffer, GetLogicalDriveStrings(Len(strDriveBuffer), strDriveBuffer))
    strDrives = Split(strDriveBuffer, Chr$(0))

    lvwDrives.ListItems.Clear

    For i = 0 To UBound(strDrives) - 1
        With lvwDrives.ListItems.Add(Text:=strDrives(i))
            udtInfo = GetDevInfo(strDrives(i))
            
            If udtInfo.Valid Then
                Select Case udtInfo.BusType
                    Case BusTypeUsb:        .SubItems(1) = "USB"
                    Case BusType1394:       .SubItems(1) = "1394"
                    Case BusTypeAta:        .SubItems(1) = "ATA"
                    Case BusTypeAtapi:      .SubItems(1) = "ATAPI"
                    Case BusTypeFibre:      .SubItems(1) = "Fibre"
                    Case BusTypeRAID:       .SubItems(1) = "RAID"
                    Case BusTypeScsi:       .SubItems(1) = "SCSI"
                    Case BusTypeSsa:        .SubItems(1) = "SSA"
                    Case BusTypeUnknown:    .SubItems(1) = "Unknown"
                End Select
                
                .SubItems(2) = IIf(udtInfo.Removable, "yes", "no")
                .SubItems(3) = Trim$(udtInfo.VendorID & " " & udtInfo.ProductID & " " & udtInfo.ProductRevision)
                
                .Tag = strDrives(i)
            End If
        End With
    Next
End Sub

Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    If uMsg = WM_DEVICECHANGE Then RefreshDriveList
End Sub
