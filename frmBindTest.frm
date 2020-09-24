VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBindTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Adapter Binding"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   315
      Left            =   4740
      TabIndex        =   3
      Top             =   2100
      Width           =   1455
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Top             =   2100
      Width           =   915
   End
   Begin MSComctlLib.ListView lvwNetAdapters 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2990
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Adapter"
         Object.Width           =   8062
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Left            =   3900
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPort 
      Caption         =   "Port:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblLvw 
      Caption         =   "Select Network Adapter:"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "frmBindTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdListen_Click()
    '-- If we're not listening then
    If cmdListen.Caption = "Listen" Then
        '--  Make sure the socket is closed
        sckTCP.Close
        '-- Listen
        Listen
        '-- Change the commandbutton caption
        cmdListen.Caption = "Stop Listening"
        '-- Set up the properties of the listview so the user can't change it and it looks nice
        lvwNetAdapters.Enabled = False
        lvwNetAdapters.BackColor = vbButtonFace
        txtPort.Enabled = False
        txtPort.BackColor = vbButtonFace
    Else
    '-- If we are already listening then
        '-- Close the socket
        sckTCP.Close
        '-- Change the caption so we can listen again
        cmdListen.Caption = "Listen"
        '-- Re-enable the listview
        lvwNetAdapters.Enabled = True
        lvwNetAdapters.BackColor = vbWhite
        txtPort.Enabled = True
        txtPort.BackColor = vbWhite
    End If
End Sub

Private Sub Form_Load()
    '-- Objects we need from WMI
    Dim objWMIService As SWbemServices
    Dim colNetAdapters As SWbemObjectSet
    Dim objNetAdapter As SWbemObject
    
    '-- Get the WMI Service object
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    '-- Fill the object set with the network adapter configs that have IPEnabled = True
    Set colNetAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    '-- Loop through the collection and add each adapter to the listview
    For Each objNetAdapter In colNetAdapters
        lvwNetAdapters.ListItems.Add , , objNetAdapter.Caption
        lvwNetAdapters.ListItems(lvwNetAdapters.ListItems.Count).SubItems(1) = objNetAdapter.IPAddress(0)
    Next
    
    '-- Clean up
    Set colNetAdapters = Nothing
    Set objWMIService = Nothing
    Set objNetAdapter = Nothing
End Sub

Private Sub sckTCP_Close()
    '-- Output in Immediate window that we've disconnected
    Debug.Print ":: Disconnected ::"
End Sub

Private Sub sckTCP_Connect()
    '-- Output in Immediate window that we've connected
    Debug.Print ":: Connected ::"
End Sub

Private Sub sckTCP_ConnectionRequest(ByVal requestID As Long)
    '-- Close the socket and accept the request
    sckTCP.Close
    sckTCP.Accept requestID
End Sub

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    
    '-- Get the data
    '-- NOTE: This data -SHOULD- be buffered, as it is TCP data. However, since this is only an
    '   example of the simplest type, I've not bothered
    sckTCP.GetData sData, vbString, bytesTotal
    
    '-- Show in Immediate window that data has arrived
    Debug.Print "<< " & StripCrLf(sData)
    
    '-- Send the same data back out again
    sckTCP.SendData StripCrLf(sData) & vbCrLf & vbCrLf
    '-- Show us sending out the data in the Immediate window
    Debug.Print ">> " & StripCrLf(sData)
End Sub

Private Function StripCrLf(ByVal s As String) As String
    '-- Remove CrLfs
    While Right$(s, 2) = vbCrLf: s = Left$(s, Len(s) - 2): Wend
    StripCrLf = s
End Function

Private Sub Listen()
    '-- Make sure we have a port to bind to
    If Trim$(txtPort.Text) = "" Then txtPort.Text = "128"
    '-- Bind to the selected port and network adapter
    sckTCP.Bind Val(Trim$(txtPort.Text)), lvwNetAdapters.SelectedItem.SubItems(1)
    '-- Listen!
    sckTCP.Listen
End Sub

Private Sub sckTCP_SendComplete()
    '-- Send is complete, close socket and listen again
    sckTCP.Close
    Listen
End Sub
