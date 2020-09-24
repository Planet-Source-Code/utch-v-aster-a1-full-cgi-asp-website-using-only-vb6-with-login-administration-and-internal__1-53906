Attribute VB_Name = "CGI4VB"
Option Explicit

Public MYID As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const DBPath$ = "f:\vbscripts\jasongoldberg\swat\swat.mdb"
'Public Const DBPath$ = "c:\html\swat\scripts\swat.mdb"

Public DBDimmed As Boolean
Public IAmMember As Boolean
Public RequiredAccess As Integer
Public DB As Database
Public LoginRequired As Boolean
Public LoginStatus As Integer
Public mScreenName As String
Public mPassWord As String
Public Action As String
Public Section As String

'======================================
' CGI4VB.BAS
' CGI routines used with 32bit VB 4.0
'======================================

Declare Function GetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long) As Long

Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    lpOverlapped As Any) As Long

Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    lpOverlapped As Any) As Long

Declare Function SetFilePointer Lib "kernel32" _
   (ByVal hFile As Long, _
   ByVal lDistanceToMove As Long, _
   lpDistanceToMoveHigh As Long, _
   ByVal dwMoveMethod As Long) As Long

Declare Function SetEndOfFile Lib "kernel32" _
   (ByVal hFile As Long) As Long

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const FILE_BEGIN = 0&

' environment variables

Public EXEName As String
Public Const EXEPath As String = "http://www.jasongoldberg.com/cgi-bin/swat/"

Public CGI_Accept            As String
Public CGI_AuthType          As String
Public CGI_ContentLength     As String
Public CGI_ContentType       As String
Public CGI_GatewayInterface  As String
Public CGI_PathInfo          As String
Public CGI_PathTranslated    As String
Public CGI_QueryString       As String
Public CGI_Referer           As String
Public CGI_RemoteAddr        As String
Public CGI_RemoteHost        As String
Public CGI_RemoteIdent       As String
Public CGI_RemoteUser        As String
Public CGI_RequestMethod     As String
Public CGI_ScriptName        As String
Public CGI_ServerSoftware    As String
Public CGI_ServerName        As String
Public CGI_ServerPort        As String
Public CGI_ServerProtocol    As String
Public CGI_UserAgent         As String

Public lContentLength As Long   ' CGI_ContentLength converted to Long
Public hStdIn         As Long   ' handle of Standard Input
Public hStdOut        As Long   ' handle of Standard Output
Public sErrorDesc     As String ' constructed error message
Public sEmail         As String ' webmaster's/your email address
Public sFormData      As String ' url-encoded data sent by the server

Type pair
  Name As String
  Value As String
End Type

Public tPair() As pair           ' array of name=value pairs

Sub ErrorHandler()
Dim rc As Long

On Error Resume Next
        
' use SetFilePointer API to reset stdOut to BOF
' and SetEndOfFile to reset EOF
' Comment or remove these functions in order to debug a program

rc = SetFilePointer(hStdOut, 0&, 0&, FILE_BEGIN)


Send "<TABLE Width=550><TR><TD Class=ne><H1>Error in " & CGI_ScriptName & "</H1><H3>Last Called Procedure: " & CurrSub & "()</H3>"
    
Send "The following internal error has occurred:"
Send "<PRE>" & sErrorDesc & "</PRE>"
Send "<I>Please</I> note what you were doing when this problem occurred, "
Send "so we can identify and correct it. Write down the Web page you were "
Send "using, any data you may have entered into a form or search box, "
Send "and anything else that may help us duplicate the problem."
Send "Then contact the administrator of this service: "
Send "<A HREF=""mailto:" & sEmail & """>"
Send "<ADDRESS>&lt;" & sEmail & "&gt;</ADDRESS></A>"
Send "</TD></TR></TABLE>    "
rc = SetEndOfFile(hStdOut)

SendFooter

End Sub
Function GetCgiValue(cgiName As String) As String
'====================================================================
' Accept the name of a pair
' Return the value matching the name
'
' tPair(0) is always empty.
' An empty string will be returned
'    if cgiName is not defined in the form (programmer error)
'    or, a select type form item was used, but no item was selected.
'
' Multiple values, separated by a semi-colon, will be returned
'     if the form item uses the "multiple" option
'     and, more than one selection was chosen.
'     The calling procedure must parse this string as needed.
'====================================================================
Dim n As Integer
 
GetCgiValue = ""
For n = 1 To UBound(tPair)
    If UCase$(cgiName) = UCase$(tPair(n).Name) Then
       If GetCgiValue = "" Then
          GetCgiValue = tPair(n).Value
       Else             ' allow for multiple selections
          GetCgiValue = GetCgiValue & ";" & tPair(n).Value
       End If
    End If
Next n
GetCgiValue = Replace(GetCgiValue, "'", "")

End Function


Sub GetFormData()
'====================================================
' Get the CGI data from STDIN and/or from QueryString
' Store name/value pairs
'====================================================
Dim sBuff      As String    ' buffer to receive POST method data
Dim lBytesRead As Long      ' actual bytes read by ReadFile()
Dim rc         As Long      ' return code

' Method POST - get CGI data from STDIN
' Method GET  - get CGI data from QueryString environment variable
'
If CGI_RequestMethod = "POST" Then
   sBuff = String(lContentLength, Chr$(0))
   rc = ReadFile(hStdIn, ByVal sBuff, lContentLength, lBytesRead, ByVal 0&)
   sFormData = Left$(sBuff, lBytesRead)
   
   ' Make sure posted data is url-encoded
   ' Multipart content types, for example, are not necessarily encoded.
   '
   If InStr(1, CGI_ContentType, "www-form-urlencoded", 1) Then
      StorePairs sFormData
   End If
End If
StorePairs CGI_QueryString
End Sub
Sub StorePairs(sData As String)
'=====================================================================
' Parse and decode form data and/or query string
' Data is received from server as "name=value&name=value&...name=value"
' Names and values are URL-encoded
'
' Store name/value pairs in array tPair(), and decode them
'
' Note: if an element in the query string does not contain an "=",
'       then it will not be stored.
'
' /cgi-bin/pgm.exe?parm=1   "1" gets stored and can be
'                               retrieved with getCgiValue("parm")
' /cgi-bin/pgm.exe?1        "1" does not get stored, but can be
'                               retrieved with urlDecode(CGI_QueryString)
'
'======================================================================
Dim pointer    As Long      ' sData position pointer
Dim n          As Long      ' name/value pair counter
Dim delim1     As Long      ' position of "="
Dim delim2     As Long      ' position of "&"
Dim lastPair   As Long      ' size of tPair() array
Dim lPairs     As Long      ' number of name=value pairs in sData

lastPair = UBound(tPair)    ' current size of tPair()
pointer = 1
Do
  delim1 = InStr(pointer, sData, "=")
  If delim1 = 0 Then Exit Do
  pointer = delim1 + 1
  lPairs = lPairs + 1
Loop

If lPairs = 0 Then Exit Sub  'nothing to add

' redim tPair() based on the number of pairs found in sData
ReDim Preserve tPair(lastPair + lPairs) As pair

' assign values to tPair().name and tPair().value
pointer = 1
For n = (lastPair + 1) To UBound(tPair)
   delim1 = InStr(pointer, sData, "=") ' find next equal sign
   If delim1 = 0 Then Exit For         ' parse complete

   tPair(n).Name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))
   
   delim2 = InStr(delim1, sData, "&")

   ' if no trailing ampersand, we are at the end of data
   If delim2 = 0 Then delim2 = Len(sData) + 1
 
   ' value is between the "=" and the "&"
   tPair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
   pointer = delim2 + 1
Next n
End Sub


Public Function UrlDecode(ByVal sEncoded As String) As String
'========================================================
' Accept url-encoded string
' Return decoded string
'========================================================

Dim pointer    As Long      ' sEncoded position pointer
Dim pos        As Long      ' position of InStr target

If sEncoded = "" Then Exit Function

' convert "+" to space
pointer = 1
Do
   pos = InStr(pointer, sEncoded, "+")
   If pos = 0 Then Exit Do
   Mid$(sEncoded, pos, 1) = " "
   pointer = pos + 1
Loop
    
' convert "%xx" to character
pointer = 1

On Error GoTo errorUrlDecode
Do
   pos = InStr(pointer, sEncoded, "%")
   If pos = 0 Then Exit Do
   
   Mid$(sEncoded, pos, 1) = Chr$("&H" & (Mid$(sEncoded, pos + 1, 2)))
   sEncoded = Left$(sEncoded, pos) _
             & Mid$(sEncoded, pos + 3)
   pointer = pos + 1
Loop
On Error GoTo 0     'reset error handling
UrlDecode = sEncoded
Exit Function

errorUrlDecode:
'--------------------------------------------------------------------
' If this function was mistakenly called with the following:
'    UrlDecode("100% natural")
' a type-mismatch error would be raised when trying
' to convert " n" from hex to character.
' Instead, a more descriptive error message will be generated.
'--------------------------------------------------------------------
Err.Clear
Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
Resume Next
End Function

Sub InitCgi()

hStdIn = GetStdHandle(STD_INPUT_HANDLE)
hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

sEmail = "jag@alphamedia.net"

'==============================
' Get the environment variables
'==============================
'
' Environment variables will vary depending on the server.
' Replace any variables below with the ones used by your server.
'

EXEName = App.EXEName

CGI_Accept = Environ("HTTP_ACCEPT")
CGI_AuthType = Environ("AUTH_TYPE")
CGI_ContentLength = Environ("CONTENT_LENGTH")
CGI_ContentType = Environ("CONTENT_TYPE")
CGI_GatewayInterface = Environ("GATEWAY_INTERFACE")
CGI_PathInfo = Environ("PATH_INFO")
CGI_PathTranslated = Environ("PATH_TRANSLATED")
CGI_QueryString = Environ("QUERY_STRING")
CGI_Referer = Environ("HTTP_REFERER")
CGI_RemoteAddr = Environ("REMOTE_ADDR")
CGI_RemoteHost = Environ("REMOTE_HOST")
CGI_RemoteIdent = Environ("REMOTE_IDENT")
CGI_RemoteUser = Environ("REMOTE_USER")
CGI_RequestMethod = Environ("REQUEST_METHOD")
CGI_ScriptName = Environ("SCRIPT_NAME")
CGI_ServerSoftware = Environ("SERVER_SOFTWARE")
CGI_ServerName = Environ("SERVER_NAME")
CGI_ServerPort = Environ("SERVER_PORT")
CGI_ServerProtocol = Environ("SERVER_PROTOCOL")
CGI_UserAgent = Environ("HTTP_USER_AGENT")

lContentLength = Val(CGI_ContentLength)   'convert to long
ReDim tPair(0)                            'initialize name/value array

End Sub


Sub Main()

Call frmMain.GetPlayers("64.156.2.135", 27024)
Call frmMain.getPing("64.156.2.135", 27024)
Call frmMain.GetServerStats("64.156.2.135", 27024)


Unload frmMain

On Error GoTo ErrorRoutine
InitCgi          ' Load environment vars and perform other initialization
GetFormData      ' Read data sent by the server
CGI_Main         ' Process and return data to server

EndPgm:
   End           ' end program

ErrorRoutine:
   sErrorDesc = Err.Description & " Error Number = " & Str$(Err.Number)
   ErrorHandler
   Resume EndPgm
End Sub

Sub SendFooter()

Send "<!-Footer Start------------------------------------------------------------------------------------------------------------------------>"
    Send "<font class=ne><BR></font>"
    'Send "<BR><BR><BR><BR><BR><BR><BR>"
    Send "<TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><HR Color=666666><font color=666666>Website Designed & Maintained by</font> <A Style=""text-decoration:none"" Href=""http://www.jasongoldberg.com"" Target=""_new""><font color=666666>Jason ""DutchMaster"" Goldberg</TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><font color=666666>© Copyright " & GetCopyRight(2004) & ". All Rights Reserved.</TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><font color=666666>Website VBScript Code Version: <font color=666666>" & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00") & "</font><HR Color=336699></TD>"
      Send "</tr>"

    Send "</TABLE>"
  Send "</BODY>"

  Send "</CENTER>"
Send "</HTML>"
Send "<!-Footer end------------------------------------------------------------------------------------------------------------------------>"

End Sub

Public Sub SendWhosOnline(Optional ShowFullNames As Boolean)
CurrSub = "SendWhosOnline"
  Dim BG As String
  Dim FG As String
  
  Send "    <TABLE Border=0 CellPadding=1 CellSpacing=0 Width=100%>"
  Send "    <TR bgcolor=336699 align=center><TD ColSpan=2 Class=ne><font color=white><b>.: Who's Playing Now :.</B></Font></TD></tr>"
  
  If Replace(Trim(ServerSettings.PlayerList), "|", "") = "" Then
    Send "<TR><TD align=center Width=20 Class=ne><Img src=""http://www.jasongoldberg.com/swat/images/person.gif"" Height=15 Width=9></TD><TD Class=ne>No Players Online</TD></TR>"
  Else
    Dim Spot As Integer
    Dim Curr As String
    Dim P As String
    Dim X As Integer
    P = ServerSettings.PlayerList
    Spot = InStr(1, P, Chr(9))
    Do Until Spot = 0
      X = X + 1
      
      BG = IIf(BG = "000000", "333333", "000000")
      FG = IIf(BG = "000000", "FFFF99", "FFFFFF")
      
      
      
      Curr = Left(P, Spot - 1)
      
      If Left$(UCase$(Curr), 9) = "[S.W.A.T]" Then FG = "FF0000"
      
      If Not (ShowFullNames) And Len(Curr) > 16 Then Curr = Left(Curr, 16) & "..."
      
      P = Mid(P, Spot + 1)
      Send "<TR BGColor=" & BG & "><TD align=center Width=12 Class=ne><Img src=""http://www.jasongoldberg.com/swat/images/person.gif"" Height=15 Width=9></TD><TD Class=ne><font color=" & FG & ">" & Replace(Curr, " ", "&nbsp;") & "</TD></TR>"
      Spot = InStr(1, P, Chr(9))
      
    Loop
  End If
  Send "    <TR><TD ColSpan=2 Class=ne><HR Color=336699></TD></tr>"

  Send "    </TABLE>"
End Sub

Function HexConv(Text As String) As String
  Dim Curr As String
  Dim X As Integer
  For X = 1 To Len(Text)
    Curr = Mid(Text, X, 1)
    HexConv = HexConv & "&#" & Format(Asc(Curr), "0000") & ";"
  Next
End Function

Sub SendHeader(sTitle As String)

  Send "Status: 200 OK"
  Send "Content-type: text/html" & vbCrLf
  
  Send "<!-Disclaimer Start------------------------------------------------------------------------------------------------------------------------>"
  Send "<! NOTE: ALL HTML CODE ON THIS WEBPAGE IS 100% COPYRIGHTED BY JASON GOLDBERG - THE AUTHOR OF THIS WEBSITE.--------------------------------->"
  Send "<! COPYING ANY PORTION OF THIS CODE LEAVE YOU VOLNERABLE FOR LAW SUITS, OR MISC. LEGAL ACTION.--------------------------------------------->"
  Send "<!-Disclaimer End-------------------------------------------------------------------------------------------------------------------------->"
  
  Send "<!-Header Start---------------------------------------------------------------------------------------------------------------------------->"
  Send "<HTML>"
  Send "<HEAD>"
  Send "<TITLE>[S.W.A.T] Counter-Strike Clan</TITLE>"
  Send "  <STYLE type=""text/css"">"
  Send "<!--"
  Send ".exact {font-family: Times New Roman;font-size: 11px;}"
  Send ".ne {font-family: tahoma;font-size: 11px;color: #3366CC;}"
  Send ".nesm {font-family: tahoma;font-size: 10px;color: #3366CC;}"
  Send ".heading {font-family: tahoma;font-size: 13px;color: #3366CC;}"
  Send ".bigheading {font-family: tahoma;font-size: 23px;color: #3366CC;}"
  Send ".names {font-family: arial;font-size: 11px;color: #ffffff;}"
  Send ".yellownames {font-family: arial;font-size: 11px;color: #ffff00;}"
  Send ".nebold {font-family: tahoma;font-size: 12px;color: #3366CC;}"
  Send "-->"
  Send "</STYLE>"
  
  Send "</head>"
  Send "<BODY BGColor=000000 leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 text=336699 link=""#3366CC"" alink=""#ff0000"" vlink=ff0000>"
  
  Send "  <CENTER>"
  
  Send "  <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne Width=261 align=Center valign=Top><BR><BR>"
  
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=261>"
  Send "    <TR>"
  Send "    <TD Class=ne align=center valign=middle>"
  Send "      <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=261>"
  Send "      <TR><TD Class=nebold align=left Colspan=2><B><font color=white>[</font><font color=red>LIVE Server Statistics</font><font color=white>]</font></b><BR><BR></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Server:</B><BR></TD><TD Class=ne><B><font color=white>" & ServerSettings.Name & "</TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Server IP:</B><BR></TD><TD Class=ne><font color=white>64.156.2.135:27024</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Players:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Players & " of " & ServerSettings.MaxPlayers & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Map:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Map & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Game Type:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Type & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Ping:</B></TD><TD Class=ne><font color=white>~ " & ServerSettings.Ping & " ms</font></TD></TR>"
  Send "      </TABLE>"
  Send "    </TD>"
  Send "    </TR>"
  Send "    </TABLE>"
  
  Unload frmMain
  
  Send "  </TD>"
  Send "  <TD Class=ne Width=259> <IMG Src=""http://www.jasongoldberg.com/swat/images/logo.gif"" Width=259 Height=125><BR></TD>"
  Send "  <TD Class=ne Width=220 valign=top><BR><BR>"
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=220>"
  Send "    <TR>"
  Send "    <TD Class=ne align=right valign=middle>"
  Send "      <TABLE CellPadding=0 CellSpacing=0 Border=0>"
  Send "      <TR><TD Class=nebold align=left Colspan=2><B><font color=white>[</font><font color=red>Other Statistics</font><font color=white>]</font></b><BR><BR></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Your IP:</B></TD><TD Class=ne><font color=white>" & CGI_RemoteAddr & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Time:</B></TD><TD Class=ne><font color=white>" & Format(Time, "hh:mm:ss AMPM") & "&nbsp;EST</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Date:</B></TD><TD Class=ne><font color=white>" & Replace(Format(Date, "dddd mmmm dd, yyyy"), " ", "&nbsp;") & "</font></TD></TR>"
  
  Send "      </TABLE>"
  Send "    </TD>"
  Send "    </TR>"
  Send "    </TABLE>"
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
  
  Send "  <HR Color=336699 Width=750>"
  Call SendToolbar
  Send "  <HR Color=666666 Width=750>"
  
  Send "<!-Header End------------------------------------------------------------------------------------------------------------------------>"
End Sub

Private Sub SendToolbar()
  Send "  <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne align=center>"
  
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0>"
  
  Send "    <TR>"
  Send "    <TD Class=ne>"
  Send MeLink("Home", "Yellow", "action=", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Members", "Yellow", "action=MemberList", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Downloads", "Yellow", "action=Downloads", True, True)
  Send "    </TD>"

  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Contact", "Yellow", "action=Contact", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Clan Application", "Yellow", "action=apply", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Report Abuse", "Yellow", "action=reportabuse", True, True)
  Send "    </TD>"
  
  If LoginStatus > 0 Then
    Dim M As Integer
    M = -1
    Call GetMailCount(M)
    CurrSub = "SendToolBar"
    Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send MeLink("Mail (<font color=white>" & M & "</font>)", "Yellow", "action=MailIndex", True, True)
    'Send "<font color=white>Mail *</font>"
    Send "    </TD>"
  End If
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Admin. Console", "White", "action=admConsole", True, True)
  Send "    </TD>"
  
  If LoginStatus > 0 Then
    Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send MeLink("Sign Out", "White", "", True)
    Send "    </TD>"
  End If
  
  Send "    </TR>"
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Public Function GetName(ID As Integer) As String
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT * FROM USERS WHERE ID=" & ID)
  GetName = RS!UserName
End Function

Public Function GetMailCount(Optional ByRef M As Integer, Optional ByRef S As Integer, Optional ByRef T As Integer) As Integer
  
  Call InitDB
  
  CurrSub = "GetMailCount"
  
  If M <> 0 Then
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=FALSE and TO='" & Trim(MYID) & "'")
    If RS.RecordCount <> 0 Then
      RS.MoveLast
      M = RS.RecordCount
    End If
    RS.Close
  End If
  If M <= 0 Then M = 0
  
  
  If S <> 0 Then
    Dim RS2 As Recordset
    Set RS2 = DB.OpenRecordset("SELECT * FROM MAIL WHERE FROM='" & Trim(MYID) & "' Order by Sent Desc")
    If RS2.RecordCount <> 0 Then
      RS2.MoveLast
      S = RS2.RecordCount
    End If
    RS2.Close
  End If
  If S <= 0 Then S = 0
  
  If T <> 0 Then
    Dim RS3 As Recordset
    Set RS3 = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=TRUE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    If RS3.RecordCount <> 0 Then
      RS3.MoveLast
      T = RS3.RecordCount
    End If
    RS3.Close
  End If
  If T <= 0 Then T = 0
  
End Function

Sub Send(S As String)
'======================
' Send output to STDOUT
'======================
Dim rc            As Long
Dim lBytesWritten As Long

S = S & vbCrLf
rc = WriteFile(hStdOut, S, Len(S), lBytesWritten, ByVal 0&)
End Sub

Sub SendB(S As String)
'============================================
' Send output to STDOUT without vbCrLf.
' Use when sending binary data. For example,
' images sent with "Content-type image/jpeg".
'============================================
Dim rc            As Long
Dim lBytesWritten As Long

rc = WriteFile(hStdOut, S, Len(S), lBytesWritten, ByVal 0&)
End Sub

Public Function GetRank(Rank As Integer) As String
CurrSub = "GetRank"
  Select Case Rank
    Case 1
      GetRank = "Leader / Founder / Owner"
    Case 2
      GetRank = "Lieutenant / Owner"
    Case 3
      GetRank = "Lieutenant / Founder"
    Case 10
      GetRank = "Lieutenant"
    Case 5
      GetRank = "Lieutenant / Server-Op"
    Case 99
      GetRank = "Lieutenant / Web Design"
    Case 6
      GetRank = "Lieutenant / Web Design"
  End Select
End Function

Public Sub Show404()
CurrSub = "Show404"
  Send "<font class=ne><BR><BR><font color=white><B>Error: 404 (Page/Action Not Found)</B><BR><BR>"
  Send "The Page/Action you are trying to reach: '<font color=yellow>" & Action & "</font>' can not be found.<BR><BR>"
  Send "Please " & MeLink("Contact", "yellow", "action=contact", True, True) & " us if you feel this message was reached in error, or return " & MeLink("Home", "yellow", "", True, True) & "</font><BR><BR><BR></font>"
End Sub

Sub ShowLogin(Optional ErrMsg As String)
CurrSub = "ShowLogin"
  Send "<!--Sub: ShowLogin-->"
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""hidden"" Name=""action"" value=""" & Action & """>"
  Send "  <Input type=""hidden"" Name=""section"" value=""" & Section & """>"
  Send "  <Input type=""hidden"" Name=""SkipDecrypt"" value=""1"">"
  Send "  <Input type=""hidden"" Name=""Updatell"" value=""1"">"
  
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=2 Border=0 Width=400>"
  
  If Trim(ErrMsg) <> "" Then Send "  <TR><TD valign=top colspan=2 Class=ne><font color=red>" & ErrMsg & "</TD></TR>"

  Send "  </TR>"
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Screen Name<BR><Input Type=Text Name=""ScreenName"" Size=25 value=""" & mScreenName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Password<BR><Input Type=Password Name=""Password"" Size=25 value=""" & mPassWord & """></TD>"
  Send "  </TR>"
  Send "  <TR>"
  Send "  <TD align=center valign=top colspan=2 Class=ne><Input type=""Submit"" value=""Log into [S.W.A.T] Console""></TD>"

  Send "  </TR>"
  Send "  </TABLE>"
  Send "  </Form>"
  
End Sub

Function GetCopyRight(StartYear As String) As String
  If StartYear = Format(Date, "yyyy") Then
    GetCopyRight = StartYear
  Else
    GetCopyRight = StartYear & "-" & Format(Date, "yyyy")
  End If
End Function

Public Function GetLoginStatus(vUser As String, vPass As String) As Integer
  
  CurrSub = "GetLoginStatus"
  Send "<!--Sub: GetLoginStatus-->"

  Call InitDB

  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Users where username='" & vUser & "'")
  
  If RS.RecordCount = 0 Then
    
    GetLoginStatus = uINVALID
  
  Else
    
    If LCase$(RS!Password) <> LCase$(vPass) Then
      
      GetLoginStatus = uINVALID
    
    Else
    
      IAmMember = RS!member
      mScreenName = RS!UserName
      MYID = RS!ID
      
      If LCase(vUser) = "k!ller v!ru$" Or LCase(vUser) = "dutchmaster" Then
        GetLoginStatus = uADMINGOD
      ElseIf RS!Admin And RS!member Then
        GetLoginStatus = uADMINMEMBER
      ElseIf RS!Admin And Not (RS!member) Then
        GetLoginStatus = uADMINUSER
      ElseIf Not (RS!Admin) And RS!member Then
        GetLoginStatus = uMEMBER
      ElseIf Not (RS!Admin) And Not (RS!member) Then
        GetLoginStatus = uUSER
      End If
    
    End If
  
  End If
  
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    
    On Error GoTo ErrorHandler
    
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    Dim i As Integer
    
    If StringToDecrypt = "" Then Exit Function

    If AlphaDecoding Then
        
        Decrypt = StringToDecrypt
        StringToDecrypt = ""


        For i = 1 To Len(Decrypt)
            
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 147)
        
        Next i
    
    End If
    
    Decrypt = ""

    Do Until StringToDecrypt = ""
        
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
    
    Loop
    
    Exit Function
    
ErrorHandler:
    
    Decrypt = ""

End Function

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim Char As String
    Encrypt = ""
    
    If StringToEncrypt = "" Then Exit Function

    For i = 1 To Len(StringToEncrypt)
        
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    
    Next i
    


    If AlphaEncoding Then
        
        StringToEncrypt = Encrypt
        Encrypt = ""


        For i = 1 To Len(StringToEncrypt)
            
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        
        Next i
    
    End If
    
    Exit Function

ErrorHandler:
    
    Encrypt = ""

End Function

Public Function GetStatusName(AccessLevel As Integer) As String
CurrSub = "GetStatusName"
  Select Case AccessLevel
    Case uINVALID
      GetStatusName = "Invalid - No Access"
    Case uUSER
      GetStatusName = "User"
    Case uMEMBER
      GetStatusName = "Member"
    Case uADMINUSER
      GetStatusName = "Admin/User"
    Case uADMINMEMBER
      GetStatusName = "Admin/Member"
    Case uADMINGOD
      GetStatusName = "Master"
  End Select
End Function

Public Sub InitDB()
  
  CurrSub = "InitDB"
  
  If DBDimmed Then Exit Sub
  
  Set DB = OpenDatabase(DBPath$)
  DBDimmed = True
  
End Sub

Public Sub UpdateLastLogin()
CurrSub = "UpdateLastLogin"
On Error GoTo ErrPoint:
  
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * from users where username='" & mScreenName & "'")
  
  RS.Edit
  RS!LastLogin = Now
  RS.Update
  
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
  
End Sub

Public Function GetUserValueID(ID As Integer, Field As String) As Variant
CurrSub = "GetUserValue"
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select " & Field & " From Users Where ID=" & ID)
  GetUserValueID = RS.Fields(Field)
End Function

Public Function PullPlug() As Boolean
  CurrSub = "PullPlug"
  PullPlug = Not (Dir(App.Path & "\KillAll.txt") = "")
End Function

Public Sub PerformMailClean(Optional FlgMessage As Integer)
  
  Call InitDB
  Dim F As String
  CurrSub = "PerformMailClean"
  
  Dim S As String
  S = Format(Date - 60, "mm/dd/yyyy")

  Dim R As Recordset
  Set R = DB.OpenRecordset("SELECT * FROM LASTCLEAN")
  R.MoveFirst
  If Abs(DateDiff("D", R!lastclean, Date)) = 0 Then Exit Sub

  Dim RS As Recordset
  Set RS = DB.OpenRecordset("select Sent from mail order by sent asc")

  With RS
    If .RecordCount = 0 Then Exit Sub
    Send "<!--Got Here!!-->  "
    
    .MoveFirst
    Do While Not .EOF
      If Abs(DateDiff("D", Date, !sent)) > 60 Then DelMail (RS!ID)
      DoEvents
      .MoveNext
      Send "<!--Got Here!!-->"
    Loop
  End With

  R.Edit
  R!lastclean = Date
  R.Update
  R.Close
  RS.Close
  
End Sub

Sub DelMail(ID As String)
  DB.Execute ("DELETE * FROM MAIL WHERE ID='" & ID & "'")
End Sub
