Attribute VB_Name = "mTypes"
Option Explicit

Public Const APP_ACCEPTED As Integer = 1
Public Const APP_NEW As Integer = 0
Public Const APP_DECLINED As Integer = -1

Public uNews As String

Public Const uINVALID As Integer = 0
Public Const uUSER As Integer = 1
Public Const uMEMBER As Integer = 2
Public Const uADMINUSER As Integer = 3
Public Const uADMINMEMBER As Integer = 4
Public Const uADMINGOD As Integer = 10

Public Type tpMember
  UserName As String
  Password As String
  nName As String
  EMail As String
  URL As String
  Rank As Integer
  Admin As Boolean
  mMember As Boolean
  ID As Integer
  AIM As String
End Type
Public pMember As tpMember

Public Type tpServerSettings
  Name As String
  Players As Integer
  MaxPlayers As Integer
  Map As String
  Type As String
  PlayerList As String
  Ping As Long
End Type
Public ServerSettings As tpServerSettings

Public Type tpApplication
  Name As String
  UserName As String
  EMail As String
  PreviousClans As String
  Comments As String
  IPAddress As String
  SubmmittedTime As Date
  Status As Integer
End Type
Public Application As tpApplication

Public Type tpAbuse
  MemberName As String
  UserName As String
  EMail As String
  Comments As String
  vDate As Date
End Type
Public Abuse As tpAbuse
