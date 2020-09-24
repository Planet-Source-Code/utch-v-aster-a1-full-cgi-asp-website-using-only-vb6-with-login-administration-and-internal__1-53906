VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wsHLData 
      Left            =   2520
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ControlChar = "ÿÿÿÿ"

Dim T As Long
Dim Pinging As Boolean
Dim Done As Boolean
Dim Players As Boolean

Public Sub GetServerStats(IP As String, Port As Integer)
    Players = False
    Done = False
    Pinging = False
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "infostring"
    Do Until Done
      DoEvents
    Loop
End Sub

Public Sub getPing(IP As String, Port As Integer)
    T = GetTickCount
    Pinging = True
    Players = False
    Done = False
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "ping"
    Do Until Done
      DoEvents
    Loop
End Sub

Public Sub GetPlayers(IP As String, Port As Integer)
    Players = True
    Done = False
    Pinging = False
    
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "players"
    Do Until Done
      DoEvents
    Loop
End Sub

Private Sub wsHLData_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo error
    
    Dim Data As String
    Dim R As Integer

    wsHLData.GetData Data, , bytesTotal

    Call ProcessData(Data)
    Exit Sub
error:
    MsgBox "An error occured while processing data. Make sure you correctly typed the host IP and port" & vbCrLf & "ERR: " & Err.Description & "NUM: " & Err.Number, vbCritical, "Error:"
End Sub

Public Function CharCount(Text As String, CText As String) As Integer
  Dim L As Integer
  Dim S As String
  L = Len(Text)
  S = Replace(Text, CText, "")
  CharCount = (L - Len(S)) / Len(CText)
End Function

Sub ProcessData(Data As String)
    Dim StartPoint As Integer
    Dim EndPoint As Integer
    Dim Info As String
    Static LastInfo As String
    Dim CurrentLine As Integer
    Dim X As Integer
    
    If Not Pinging Then
      If Not Players Then
        Dim R As Integer
        R = InStr(5, Data, Chr(0))
        If R Then
            Data = Mid(Data, R + 1)
        End If
        
        StartPoint = 1
        CurrentLine = 1
        
        Do
            StartPoint = InStr(StartPoint, Data, "\")
            EndPoint = InStr(StartPoint + 1, Data, "\")
    
            If EndPoint <> 0 Then
                Info = Mid(Data, StartPoint + 1, EndPoint - StartPoint - 1)
            Else
                Info = Mid(Data, StartPoint + 1, 1)
            End If
           
            CurrentLine = CurrentLine + 1
            If CurrentLine Mod 2 = 0 Then
              LastInfo = LCase$(Info)
            Else
                If LastInfo = "players" Then
                  ServerSettings.Players = CharCount(ServerSettings.PlayerList, Chr(9))
                ElseIf LastInfo = "max" Then
                  ServerSettings.MaxPlayers = Val(Info)
                ElseIf LastInfo = "hostname" Then
                  ServerSettings.Name = Info
                ElseIf LastInfo = "map" Then
                  ServerSettings.Map = Info
                ElseIf LastInfo = "password" Then
                  ServerSettings.Type = IIf(Val(Info) = 1, "Private", "Public")
                  Done = True
                End If
                
                If Info = "1" Or Info = "0" Then
                    Info = CBool(Info)
                End If
                
            End If
            
            If EndPoint <> 0 Then
                StartPoint = EndPoint
            Else
                Exit Do
            End If
        Loop
        
      Else
            
        Dim Spot As Integer
        Dim Curr As String
        Dim C As Integer
        
        Do Until Len(Data) = 0
          C = C + 1
          Data = Mid(Data, IIf(C = 1, 8, 10))
          If Data = "" Then Exit Do
          Spot = InStr(1, Data, Chr(0))
          Curr = Left(Data, Spot - 1)
          ServerSettings.PlayerList = ServerSettings.PlayerList & Curr & Chr(9)
          Data = Mid(Data, Spot + 1)
        Loop
      
        Done = True
      
      End If
    Else
     ServerSettings.Ping = (GetTickCount() - T)
     Done = True
    End If
End Sub
