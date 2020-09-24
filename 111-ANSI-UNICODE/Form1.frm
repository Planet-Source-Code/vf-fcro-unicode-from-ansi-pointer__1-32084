VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "UNICODE FROM ANSI POINTER...AUTHOR:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create ANSI String in Memory"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get UNICODE String From Pointer Of ANSI String"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private MemH As Long
Private MemP As Long
Private Sub Command1_Click()
Dim Unicode As String
Dim lenstr As Long
lenstr = lstrlen(MemP)
Unicode = Space(lenstr)
lstrcpy Unicode, MemP
Label1 = Unicode
GlobalUnlock MemH
GlobalFree MemH
End Sub

Private Sub Command2_Click()
Dim StringX As String
StringX = "UNICODE From ANSI Pointer!" & Chr(CByte(0))
MemH = GlobalAlloc(&H2 Or &H40, Len(StringX))
MemP = GlobalLock(MemH)
CopyMemory ByVal MemP, ByVal StringX, Len(StringX)

End Sub

