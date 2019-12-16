VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton CmdConnenct 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub crcNum(cmdstring() As Byte, ByVal size As Integer, ByRef p_hiByte As Integer, ByRef p_loByte As Integer)
    Dim outData As Integer
    Dim ressreg_crc As Integer
    ressreg_crc = &HFFFF
    For i = LBound(cmdstring) To size
        ressreg_crc = ressreg_crc Xor (Int(cmdstring(i)) And &HFF)
        ressreg_crc = ressreg_crc And &HFFFF
        For j = 0 To 7
            outData = ressreg_crc And &H1
            ressreg_crc = Int(ressreg_crc / 2)
            ressreg_crc = ressreg_crc And &H7FFF
            If outData Then
                ressreg_crc = ressreg_crc Xor &HA001
            End If
        Next j
    Next i
    p_hiByte = ressreg_crc And &HFF
    p_loByte = (ressreg_crc And &HFF00) / 255
End Sub

Function crc(cmdstring() As Byte) As Boolean
    Dim HiByte As Integer
    Dim LoByte As Integer
    Dim outData As Integer
    Call crcNum(cmdstring, UBound(cmdstring) - 2, HiByte, LoByte)
    If HiByte = cmdstring(UBound(cmdstring)) And LoByte = cmdstring(UBound(cmdstring) - 1) Then
    crc = True
    Exit Function
    End If
    crc = 0
End Function
Private Sub CmdConnenct_Click()
    MSComm1.CommPort = 2
    MSComm1.Settings = "19200,e,8,1"
    MSComm1.InputLen = 0  ' 读入整个缓冲区。
    MSComm1.InBufferSize = 1024 '缓冲区大小
    MSComm1.OutBufferSize = 1024 ' 发送缓冲区大小
    MSComm1.InputMode = comInputModeBinary      '采用二进制传输
    MSComm1.InBufferCount = 0   '清空接受缓冲区
    MSComm1.OutBufferCount = 0  '清空传输缓冲区
    MSComm1.RThreshold = 1      '产生MSComm事件
    MSComm1.SThreshold = 0 '一次发送所有数据 ,发送数据时不产生OnComm 事件
    MSComm1.PortOpen = True
    
End Sub
Private Sub MSComm1_OnComm()
    Dim revalue() As Byte
    Dim StrData As String
    
    If (MSComm1.CommEvent = comEvReceive) Then
    revalue = MSComm1.Input
    
    If crc(revalue) Then
    List1.AddItem "crc 错误"
    Exit Sub
    End If
    For i = LBound(revalue) To UBound(revalue)
        If Len(Hex(revalue(i))) = 1 Then
            StrData = StrData & "0" & Hex(revalue(i))
        Else
            StrData = StrData & Hex(revalue(i))
        End If
        StrData = StrData + " "
    Next i
    End If
    List1.AddItem StrData
End Sub
