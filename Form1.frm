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

Sub printInformationFrame(cmdstring() As Byte)
    Dim StrData As String
    For i = LBound(cmdstring) To UBound(cmdstring)
        If Len(Hex(cmdstring(i))) = 1 Then
            StrData = StrData & "0" & Hex(cmdstring(i))
        Else
            StrData = StrData & Hex(cmdstring(i))
        End If
        StrData = StrData + " "
    Next i
    List1.AddItem StrData
End Sub
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
    p_loByte = Int((ressreg_crc And &HFF00) / 256) And &HFF
End Sub
Sub decode(cmdstring() As Byte)
    Dim i As Integer
    For i = 0 To 7
        If Int(cmdstring(1)) = Module1.commandCode(i) Then
        Exit For
        End If
    Next i
    Select Case i
          Case Is < 4:
              Call retrueRead(cmdstring, i) '读 返回
          Case 4:
              
          Case 5:
              
          Case 6:
             
          Case 7:
              
          Case Else
             
      End Select
End Sub


Sub retrueRead(cmdstring() As Byte, p_id As Integer)
Dim retrueStr() As Byte
Dim size As Integer
Dim start As Integer
Dim hiByte As Integer
Dim loByte As Integer
Dim data As Integer

start = cmdstring(3)
size = cmdstring(5) - 1
ReDim retrueStr(2)
retrueStr(0) = cmdstring(0)
retrueStr(1) = cmdstring(1)
    Select Case p_id
          Case 0:
              ReDim Preserve retrueStr(5)
              
              retrueStr(2) = 1
              retrueStr(3) = 0
              For j = size To 0 Step -1
              retrueStr(3) = retrueStr(3) * 2 + Module1.coils(start + j)
              Next j
              
              Call crcNum(retrueStr, 3, hiByte, loByte)
              retrueStr(4) = hiByte
              retrueStr(5) = loByte
             Call printInformationFrame(retrueStr)
             
             
          Case 1:
              
              ReDim Preserve retrueStr(5)
              retrueStr(2) = 1
              retrueStr(3) = 0
              
              For j = size To 0 Step -1
              retrueStr(3) = retrueStr(3) * 2 + Module1.discreteInputs(start + j)
              Next j
              Call crcNum(retrueStr, 3, hiByte, loByte)
              
              retrueStr(4) = hiByte
              retrueStr(5) = loByte
             Call printInformationFrame(retrueStr)
          Case 2:
                N = 2 * cmdstring(5) + 4
              ReDim Preserve retrueStr(N)
              retrueStr(2) = 2 * cmdstring(5)
              i = 0
              For j = 0 To size
              i = 2 * j
              retrueStr(3 + i) = Int(Module1.InputRegisters(start + j) / 256)
              retrueStr(3 + i + 1) = Int(Module1.InputRegisters(start + j) And &HFF)
              Next j
              Call crcNum(retrueStr, Int(retrueStr(2)) + 2, hiByte, loByte)
              
              retrueStr(N - 1) = hiByte
              retrueStr(N) = loByte
             Call printInformationFrame(retrueStr)
             
          Case 3:
                N = 2 * cmdstring(5) + 4
              ReDim Preserve retrueStr(N)
              retrueStr(2) = 2 * cmdstring(5)
              i = 0
              For j = 0 To size
              i = 2 * j
              retrueStr(3 + i) = Int(Module1.holdingRegisters(start + j) / 256)
              retrueStr(3 + i + 1) = Int(Module1.InputRegisters(start + j) And &HFF)
              Next j
              Call crcNum(retrueStr, Int(retrueStr(2)) + 2, hiByte, loByte)
              
              retrueStr(N - 1) = hiByte
              retrueStr(N) = loByte
             Call printInformationFrame(retrueStr)
          Case Else
          Exit Sub
      End Select
      MSComm1.Output = retrueStr
End Sub

Function crc(cmdstring() As Byte) As Boolean
    Dim hiByte As Integer
    Dim loByte As Integer
    Dim outData As Integer
    
    If UBound(cmdstring) < 6 Then
    crc = 0
    Exit Function
    End If
    
    Call crcNum(cmdstring, UBound(cmdstring) - 2, hiByte, loByte)
    If hiByte = cmdstring(UBound(cmdstring)) And loByte = cmdstring(UBound(cmdstring) - 1) Then
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
    
    Module1.commandCode(0) = &H1 '读线圈寄存器 colis
    Module1.commandCode(1) = &H2 '读离散寄存器 discreteInputs
    Module1.commandCode(2) = &H4 '读输入寄存器 InputRegisters
    Module1.commandCode(3) = &H3 '读保持寄存器 holdingRegisters
    
    Module1.commandCode(4) = &H5 '写单个线圈寄存器
    Module1.commandCode(5) = &HF '写多个线圈寄存器器
    Module1.commandCode(6) = &H6 '写单个保持寄存器
    Module1.commandCode(7) = &H10 '写多个保持寄存器
    
    
    
    MSComm1.PortOpen = True
End Sub
Private Sub MSComm1_OnComm()
    Dim revalue() As Byte
    If (MSComm1.CommEvent = comEvReceive) Then
    revalue = MSComm1.Input
    Call printInformationFrame(revalue)
    
    If crc(revalue) Then
    List1.AddItem "crc 错误"
    Exit Sub
    End If
    
    Call decode(revalue)
    
    End If

End Sub
