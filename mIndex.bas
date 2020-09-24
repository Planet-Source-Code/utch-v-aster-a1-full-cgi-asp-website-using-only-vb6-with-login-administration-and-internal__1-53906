Attribute VB_Name = "mIndex"
Option Explicit

Public CurrSub As String
Public Const DutchID As Integer = 1
Public Const KillerID As Integer = 2
Public Const CatID As Integer = 3
Public Const MowadID As Integer = 9
Public Const FettID As Integer = 6

Public Const TRASHBOX As Integer = -1
Public Const INBOX As Integer = 0
Public Const SENTBOX As Integer = 1

Public LastMailClean As String

Public Sub CGI_Main()
  CurrSub = "CGI_Main"

  Action = LCase(GetCgiValue("action"))
  Section = LCase(GetCgiValue("section"))
  mScreenName = LCase(GetCgiValue("ScreenName"))
  
  If Val(GetCgiValue("skipdecrypt")) = 1 Then
    mPassWord = LCase(GetCgiValue("PassWord"))
  Else
    mPassWord = Decrypt(LCase(GetCgiValue("PassWord")))
  End If
  
  LoginStatus = GetLoginStatus(mScreenName, mPassWord)
  SendHeader ("[S.W.A.T] Counter-Strike Clan")
  
  If PullPlug() Then
    Send "Page temporarily out of service. Please try back in 10 minutes."
    SendFooter
    Exit Sub
  End If
  
  Call PerformMailClean
  
  CurrSub = "CGI_Main"
  
  If Action = "" Then
    Call SendIndex
  ElseIf Action = "contact" Then
    Call SendContactInfo
  ElseIf Action = "mailindex" Then
    Call ShowMailIndex
  ElseIf Action = "restoremail" Then
    Call RestoreMail
  ElseIf Action = "deletemail" Then
    Call DeleteMail
  ElseIf Action = "sendmail" Then
    Call sendMail
  ElseIf Action = "readmail" Then
    Call DisplayMail
  ElseIf Action = "replymail" Then
    Call ReplyMail
  ElseIf Action = "composemail" Then
    Call ComposeMail
  ElseIf Action = "serverstats" Then
    Call SendServerStats
  ElseIf Action = "memberlist" Then
    Call SendMemberList
  ElseIf Action = "apply" Then
    Call sendApplication
  ElseIf Action = "reportabuse" Then
    Call SendAbuseForm
  ElseIf Action = "submitapplication" Then
    Call ProcessApplication
  ElseIf Action = "downloads" Then
    Call SendDownloads
  ElseIf Action = "submitabuse" Then
    Call ProcessAbuse
  ElseIf Action = "admconsole" Then
    
    Send "<!--Action: admconsole-->"
    Send "<!--Section: " & Section & "-->"
    
    Call ProcessAdminClick
    
  Else
    Call Show404
  End If
  
  SendFooter
End Sub

Sub RestoreMail()
  Dim RS As Recordset
  Call InitDB
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  If RS.RecordCount > 0 Then
    RS.Edit
    RS!trash = False
    RS.Update
  End If
  ShowMailIndex (-1)
End Sub

Sub DeleteMail()
  Dim RS As Recordset
  Call InitDB
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  If RS.RecordCount > 0 Then
    If RS!trash Then
      RS.Delete
    Else
      RS.Edit
      RS!trash = True
      RS.Update
    End If
  End If
  ShowMailIndex (Val(GetCgiValue("mailbox")))
End Sub

Sub ReplyMail()
  Dim RS As Recordset
  Call InitDB
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  Call ComposeMail(RS!From, RS!Subject, RS!Message)
End Sub

Sub sendMail()
  Dim vTO As String
  Dim vSubject As String
  Dim vMessage As String
  
  vTO = GetCgiValue("MemberName")
  vSubject = GetCgiValue("Subject")
  vMessage = GetCgiValue("Message")

  Call InitDB
  
  If vSubject = "" Then vSubject = "[No Subject Provided]"
  If vMessage = "" Then vSubject = "[No Message Provided]"
    
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Mail")
    
  If Left(vTO, 4) = "[  A" Then
    
    Dim uRS As Recordset
    
    If InStr(1, LCase$(vTO), "upper") Then
      Set uRS = DB.OpenRecordset("SELECT * FROM USERS WHERE MEMBER and RANK<=6")
    Else
      Set uRS = DB.OpenRecordset("SELECT * FROM USERS WHERE MEMBER")
    End If
    
    uRS.MoveFirst
    Do While Not (uRS.EOF)
      RS.AddNew
      RS!To = uRS!ID
      RS!sent = Now
      RS!From = MYID
      RS!Message = vMessage
      RS!Subject = vSubject
      RS!read = False
      RS!ID = GetRandomMailID(5)
      RS.Update
      uRS.MoveNext
    Loop
    Send "<font class=ne><font color=red><B>Mail Sent To: ALL SWAT MEMBERS</B></FONT></FONT>"
  
  Else
    RS.AddNew
    RS!To = Trim$(vTO)
    RS!sent = Now
    RS!From = MYID
    RS!Message = vMessage
    RS!Subject = vSubject
    RS!read = False
    RS!ID = GetRandomMailID(5)
    RS.Update
    Send "<font class=ne><font color=red><B>Mail Sent" & "</B></FONT></FONT>"
    
  End If
  Call ShowMailIndex
  
End Sub

Function GetRandomMailID(Length As Integer) As String
  Dim RS As Recordset
  Dim Letts As String
  Letts = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
  
  Call InitDB
  
  Dim RandStr As String
  Dim X As Integer
  
TryAgain:
  
  RandStr = ""
  For X = 1 To Length
    RandStr = RandStr & Mid(Letts, Rand(1, 36), 1)
  Next
  
  
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & RandStr & "'")
  If RS.RecordCount <> 0 Then
    RS.Close
    GoTo TryAgain
  End If
  
  GetRandomMailID = RandStr

End Function

Public Function Rand(Min As Integer, Max As Integer) As Integer
10:
    Rand = Int((Rnd * Max) + Min)
    If Rand < Min Or Rand > Max Then GoTo 10
End Function

Sub ComposeMail(Optional ReplyName As String, Optional ReplySubject As String, Optional ReplyMessage As String)
  Dim vTO As String
  Dim vSubject As String
  Dim vMessage As String
  
  If ReplyName <> "" Then
    vTO = ReplyName
    vSubject = "Re: " & ReplySubject
    vMessage = ""
  Else
    vTO = GetCgiValue("MemberName")
    vSubject = GetCgiValue("Subject")
    vMessage = GetCgiValue("Message")
  End If

  Send "<font class=ne><BR><BR></font>"
  Send "<form action=""" & EXEPath & "index.exe"" Method=post>"
  Send ""
  Send "  <Input Type=""Hidden"" Name=""Action"" Value=""sendmail"">"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send ""
  Send "  <TABLE border=1 bordercolor=336699 BGColor=FFFFFF Width=500 CellPadding=0 CellSpacing=0><TR><TD Class=ne align=center>"
  Send "    <BR><TABLE BGColor=FFFFFF Width=90% CellPadding=0 CellSpacing=0>"
  Send "    <TR><TD Class=ne Colspan=3 Align=center>"
  Send "    <font color=9c1100><B><U>NOTE</U><BR><U>All</U> Mail gets deleted after 60 days, regardless<BR>of whether it is read, unread, or trashed.<hr width=95% size=2>"
  Send "    </TD></tr>"
  Send "    <TR>"
  Send "    <TD valign=middle Class=ne><font color=000000><B>To:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Call MemberCombo(vTO, True, "[  All Upper-Level Admins  ]", True, "[  All SWAT Members  ]")
  Send "    </TD>"
  Send "    </TR>"
  Send "    <TR>"
  Send "    <TD valign=middle Class=ne><font color=000000><B>Subject:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne><input Type=""text"" Name=""Subject"" Value=""" & vSubject & """ Size=50></TD>"
  Send "    </TR>"
  Send "    <TR>"
  Send "    <TD valign=top Class=ne><font color=000000><B>Message:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne><TEXTAREA Name=""Message"" Cols=50 Rows=6>" & vMessage & "</TEXTAREA></TD>"
  Send "    </TR>"
  Send "    <TR><TD Class=ne Colspan=3 align=center><BR><INPUT Type=""Submit"" Value=""Send Message""></TD></TR>"
  Send "    </TABLE><BR>"
  Send "  </TD></TR></TABLE>"
  Send "</form>"
  
End Sub

Sub DisplayMail()
  
  Dim ID As String
  ID = GetCgiValue("ID")
  
  Call InitDB
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Mail Where ID='" & ID & "'")
  
  If RS.RecordCount = 0 Then
    Send "<font Class=ne><font color=red><B>MESSAGE NOT FOUND</B></FONT></FONT><BR>"
    ShowMailIndex
  Else
    Send "<font class=ne><BR></font><TABLE cellspacing=0 Width=500 border=1 bordercolor=336699 BGColor=white>"
    Send "<TR>"
    Send "<TD Class=ne>"
    Send "  <TABLE Width=500 BGColor=white width=95%>"
    Send "  <TR>"
    Send "  <TD Class=ne><font color=black><B>To: " & GetName(Val(RS!To)) & "</TD>"
    Send "  <TD Class=ne align=right><font color=black><B>Sent: </B>" & Format(RS!sent, "mm/dd/yyyy hh:mm AMPM") & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD Class=ne colspan=2><font color=black><B>From: " & GetName(Val(RS!From)) & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><B>Subject: </B>" & RS!Subject & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><IMG src=""http://www.jasongoldberg.com/swat/images/line.gif"" Width=100% Height=2></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><TABLE Width=85%><TR><TD Class=ne><font color=336699><BR>" & RS!Message & "<BR><BR></TD></TR></TABLE></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><IMG src=""http://www.jasongoldberg.com/swat/images/line.gif"" Width=100% Height=2></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne>"
    Send "    <TABLE>"
    Send "    <TR>"
    Send "    <TD Class=ne>" & MeLink("Reply to " & GetName(Val(RS!From)), "9C1100", "Action=replymail&Id=" & ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
    Send "    <TD Class=ne>" & MeLink("Delete Message", "9C1100", "Action=deletemail&Id=" & ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
    Send "    <TD Class=ne>" & MeLink("Compose Message", "9C1100", "Action=composemail", True, True) & "</TD>"
    Send "    </TR>"
    Send "    </TABLE>"
    Send "  </TD>"
    Send "  </TR>"
    Send "  </TABLE>"
    Send "</TD>"
    Send "</TR>"
    Send "</TABLE>"
    
    If MYID = Val(RS!To) Then
      RS.Edit
      RS!read = True
      RS.Update
    End If
  End If
End Sub

Sub SendDownloads()

  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=600>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  Send "  <TR bgcolor=white><TD align=center Class=Heading colspan=2><B>Direct Downloads</TD><TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD Class=ne width=150 valign=top><A Href=""http://www.jasongoldberg.com/files/maps.zip""><font color=yellow><B>CS&nbsp;Map&nbsp;Pack</B></font></a></font></TD><TD Class=ne>"
  Send "<font color=red>75 Maps (~62.5 MB)<BR>"
  Send Replace("  <font color=white> aim_ak-colt.bsp, aimtrain.bsp, as_oilrig.bsp, as_tundra.bsp, awp_city.bsp, awp_map.bsp, awp_mapXL.bsp, awp_snowfun.bsp, cs_747.bsp, cs_assault.bsp, cs_assault_upc.bsp, cs_assault2k.bsp, cs_backalley.bsp, cs_beersel_f.bsp, cs_ciudad.bsp, cs_deagle5.bsp, cs_delta_assault.bsp, cs_estate.bsp, cs_grenadefrenzy.bsp, cs_havana.bsp, cs_italy.bsp, cs_mario_b2.bsp, cs_mice_final.bsp, cs_militia.bsp, cs_office.bsp, cs_office_old.bsp, cs_prison.bsp, cs_prospeedball.bsp, cs_rats2.bsp, cs_rats2_final.bsp, ", ",", "<font color=red>,</font>")
  Send Replace("  cs_reflex.bsp, cs_shogun.bsp, cs_siege.bsp, cs_tibet.bsp, cs_winternights.bsp, de_747.bsp, de_aztec.bsp, de_bridge.bsp, de_cbble.bsp, de_celtic.bsp, de_chateau.bsp, de_clan2_fire.bsp, de_dust.bsp, de_dust2.bsp, de_dust2002.bsp, de_flatout.bsp, de_iced2k.bsp, de_icestation.bsp, de_inferno.bsp, de_jeepathon6k.bsp, de_mog.bsp, de_nuke.bsp, de_pacman.bsp, de_piranesi.bsp, de_prodigy.bsp, de_rats.bsp, de_rats3.bsp, de_scud.bsp, de_storm.bsp, de_subway.bsp, de_survivor.bsp, de_torn.bsp, de_train.bsp, de_vegas.bsp, de_vertigo.bsp, de_village.bsp, de_volare.bsp, de_wastefacility.bsp, fy_iceworld.bsp, fy_iceworld_arena.bsp, fy_iceworld2k.bsp, ", ",", "<font color=red>,</font>")
  Send Replace("  fy_pool_day.bsp, he_tennis.bsp, Jay1.bsp, ka_legoland.bsp, motel.bsp, playground_x.bsp, playground3.bsp, rdw_hideout_b4.bsp, scout_map.bsp, starwars2A.bsp, the_hood.bsp, tr_1.bsp, tr_1a.bsp, tr_2.bsp, tr_3.bsp", ",", "<font color=red>,</font>")
  Send "  </TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  Send "  </TABLE>"
  
  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=600>"
  
  Send "  <TR bgcolor=white><TD align=center Class=Heading colspan=2><B>Download Links</TD><TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.counter-strike.net/mod_full.html""><font color=yellow><B>counter-strike.net</B></font></a></font></TD>"
  Send "<TD Class=ne><font color=red>The official Counter-Strike Website</font><BR><font color=white>Download the full CS Mod (for Half-Life), or upgrade your CS to the newest version. Anything you could possibly want for your CS Core can be found here.</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://spitfireservers.net/modules.php?name=Downloads&d_op=viewdownload&cid=1#cat""><font color=yellow><B>SpitFireServers.net</B></font></a></font></TD>"
  Send "<TD Class=ne><font color=red>A pretty nice collection of CS stuff.</font><BR><font color=white>Download sprays, console images, models, sounds, maps, crosshairs, and more...</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  </TABLE>"

End Sub



Sub SendContactInfo()

  CurrSub = "SendContactInfo"

  Send " <font class=ne><BR></font>"
  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=400>"
  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Clan Administration</B></Font></TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(KillerID, "username") & "</B></Font> - <Font Color=red>" & GetRank(Val(GetUserValueID(KillerID, "Rank"))) & "</Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(KillerID, "email") & """><Font Color=white>" & GetUserValueID(KillerID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(KillerID, "aim") & """><Font Color=white>" & GetUserValueID(KillerID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(CatID, "username") & "</B></Font> - <Font Color=red>" & GetRank(Val(GetUserValueID(CatID, "Rank"))) & "</Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(CatID, "email") & """><Font Color=white>" & GetUserValueID(CatID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(CatID, "aim") & """><Font Color=white>" & GetUserValueID(CatID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(MowadID, "username") & "</B></Font> - <Font Color=red>" & GetRank(Val(GetUserValueID(MowadID, "Rank"))) & "</Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(MowadID, "email") & """><Font Color=white>" & GetUserValueID(MowadID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(MowadID, "aim") & """><Font Color=white>" & GetUserValueID(MowadID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Game Server Administration</B></Font></TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(FettID, "username") & "</B></Font> - <Font Color=red>" & GetRank(Val(GetUserValueID(FettID, "Rank"))) & "</Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(FettID, "email") & """><Font Color=white>" & GetUserValueID(FettID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(FettID, "aim") & """><Font Color=white>" & GetUserValueID(FettID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Website Administration</B></Font></TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(DutchID, "username") & "</B></Font> - <Font Color=red>" & GetRank(Val(GetUserValueID(DutchID, "Rank"))) & "</Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(DutchID, "email") & """><Font Color=white>" & GetUserValueID(DutchID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(DutchID, "aim") & """><Font Color=white>" & GetUserValueID(DutchID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub SearchUsers(Term As String)
  
  CurrSub = "SearchUsers"
  
  Call InitDB
  Dim RS As Recordset
  
  Set RS = DB.OpenRecordset("Select * from users Where Name like '*" & Term & "*' OR " & _
                                                       "UserName like '*" & Term & "*' OR " & _
                                                       "EMail like '*" & Term & "*' OR " & _
                                                       "URL like '*" & Term & "*' OR " & _
                                                       "AIM like '*" & Term & "*' ORDER BY USerNAME")
             
  Send "<TABLE><TR><TD Class=ne Valign=Middle><B>Search Users</B></TD></TR><TR><TD Class=ne Valign=Middle>"
  Send "    <form action=""" & EXEPath & "index.exe"" Method=post>"
  Send "    <Input Type=""Hidden"" Name=""Action"" Value=""admConsole"">"
  Send "    <Input Type=""Hidden"" Name=""section"" Value=""searchusers"">"
  Send "    <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "    <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send "    <input Type=Text Size=25 name=""search"" Value=""" & GetCgiValue("search") & """>"

  Send "</TD></TR></TABLE>"

  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD colspan=6 class=ne valign=top align=center>"
  Send "  <B><font color=red><BR>Search Results For: '<font color=white>" & Term & "</font>'</b><BR><BR>"
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top><B><font color=black>Username</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Name</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail/URL</TD>"
  Send "  <TD class=ne valign=top align=center><B><font color=black>Member?/Admin?</TD>"
  Send "  <TD class=ne valign=top align=center><B><font color=black>AIM</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Last Login</TD>"
  Send "  </TR>"

  Dim BG As String

  If RS.RecordCount = 0 Then
  
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      
      BG = IIf(BG = "333333", "000000", "333333")
      
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne valign=top><B><font color=White><B>" & RS!UserName & "</TD>"
      Send "  <TD class=ne valign=top><B><font color=White><B>" & RS!Name & "</B></TD>"
      Send "  <TD class=ne valign=top><B><font color=White><A Href=""mailto:" & RS!EMail & """>" & RS!EMail & "</TD>"
      Send "  <TD class=ne valign=top align=center><B><font color=White>" & RS!member & " / " & RS!Admin & "<BR></TD>"
      
      If Len(RS!AIM) <> 0 Then
        Send "  <TD class=ne align=center valign=top><B><a href=""aim:goim?screenname=" & RS!AIM & """><font color=White>" & RS!AIM & "</TD>"
      Else
        Send "  <TD class=ne align=center valign=top><B><font color=White>?</TD>"
      End If
      
      If DateDiff("d", "01/01/1900", RS!LastLogin) = 0 Then
        Send "  <TD class=ne valign=top><B><font color=White>Never</TD>"
      Else
        Send "  <TD class=ne valign=top><B><font color=White>" & Format(RS!LastLogin, "mmm dd 'yy") & "</TD>"
      End If
      Send "  </TR>"
      
      RS.MoveNext
    Loop
  End If

  Send "</TABLE>"
  Send "    </form>"
End Sub

Public Function CheckForNulls(Text As Variant, Optional NullIsSpace As Boolean) As String
  If IsNull(Text) Then
    CheckForNulls = IIf(NullIsSpace, "&nbsp;", "")
  Else
    CheckForNulls = Text
  End If
End Function

Sub SendPostNews(Optional ID As Integer)

  CurrSub = "SendPostNews"

  If ID = 0 Then
    uNews = GetCgiValue("news")
    Send " <font class=ne><BR></font>"
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""submitnews"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=600>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><b>Post News</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=ne><font color=""RED""><br><br><b>This box suppports HTML ONLY. Use ""&lt;BR&gt;"" instead of carraige return. Keep HTML as simple as possible (Text, Links, iFrames, Images) and try to avoid tables, javascript, CSS, etc..<br><br><br></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>News Posted By: " & mScreenName & "<BR><TextArea Name=""news"" Cols=60 Rows=8>" & uNews & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit News""></TD></TR>"
    Send "  </TABLE>"
    Send "  </Form>"
  Else
    
    Call InitDB
    
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("Select * from news where id=" & ID)
    
    Send " <font class=ne><BR></font>"
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""processnews"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <Input type=""Hidden"" Name=""id"" value=""" & ID & """>"
    
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=600>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><b>Post News</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=ne><font color=""RED""><br><br><b>This box suppports HTML ONLY. Use ""&lt;BR&gt;"" instead of carraige return. Keep HTML as simple as possible (Text, Links, iFrames, Images) and try to avoid tables, javascript, CSS, etc..<br><br><br></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>News Posted By: " & RS!Postedby & "<BR><TextArea Name=""news"" Cols=60 Rows=8>" & RS!News & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit News""></TD></TR>"
    Send "  </TABLE>"
    Send "  </Form>"
  End If
End Sub

Sub ProcessAdminClick()
  
  CurrSub = "ProcessAdminClick"
  On Error GoTo ErrPoint:
    
    Dim vReqStatus As Integer
        
    If LoginStatus < 4 Then
      If LoginStatus = uINVALID Then
        Call ShowLogin
      Else
        Send "<TABLE Width=750><TR><TD Class=ne><B><font color=red>** ACESS DENIED!</font></b></TD></TR></TABLE>"
        Call SendIndex
      End If
      Exit Sub
    End If
    
    If LoginStatus > 0 Then Send "<font class=ne><font color=white><b>Welcome <font class=ne><font color=yellow>" & IIf(IAmMember, "[S.W.A.T] ", "") & mScreenName & "</font></font>. <font color=white>You are logged in with access level: <font color=yellow>" & LoginStatus & "</font> (<font color=yellow>" & GetStatusName(LoginStatus) & "</font>)</font></font></font></font></b><BR><HR Color=336699 Width=750>"
    
    If Val(GetCgiValue("updatell")) = 1 Then
      Send "<!--1-->"
      UpdateLastLogin
      Send "<!--2-->"
    End If
    
    If Section = "" Then
      Call ShowMainConsoleMenu
    ElseIf Section = "viewabuse" Then
      Call ListAbuse(50)
    ElseIf Section = "addprofile" Then
      Call AddNewMember
    ElseIf Section = "updatemember" Then
      Call UpdateMember
    ElseIf Section = "viewallabuse" Then
      Call ListAbuse(0)
    ElseIf Section = "viewapps" Then
      Call ListApplications(50)
    ElseIf Section = "cleanmail" Then
      Call PerformMailClean
      Send "<font class=ne><font color=red><B>Mail Cleaned</B></font></font>"
      Call ShowMainConsoleMenu
    ElseIf Section = "viewallapps" Then
      Call ListApplications(0)
    ElseIf Section = "searchusers" Then
      Call SearchUsers(GetCgiValue("search"))
    ElseIf Section = "deleteapp" Then
      Call DeleteApp(Val(GetCgiValue("appnumber")))
    ElseIf Section = "deleteabuse" Then
      Call DeleteAbuse(Val(GetCgiValue("appnumber")))
    ElseIf Section = "postnews" Then
      Call SendPostNews
    ElseIf Section = "editnews" Then
      Call SendPostNews(Val(GetCgiValue("id")))
    ElseIf Section = "submitnews" Then
      Call ProcessNewNews
    ElseIf Section = "processnews" Then
      Call ProcessNewNews(GetCgiValue("id"))
    ElseIf Section = "deletenews" Then
      Call DeleteNews(GetCgiValue("id"))
    ElseIf Section = "editprofile" Then
      Dim i As Integer
      i = Val(GetCgiValue("member"))
      If i = 0 Then
        Call SendMemberList(True)
      Else
          Call ShowMemberEdit(i)
      End If
    End If
    
    Send "<font class=ne><BR><BR></font>"
    Exit Sub

ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub
Sub AddNewMember()
  
  On Error Resume Next
  
  Dim T As Integer
  Call InitDB
  
  Dim RS As Recordset
  
  Set RS = DB.OpenRecordset("Select * From Users Where Username='NEW'")
  If RS.RecordCount = 0 Then
    RS.Close
    Set RS = DB.OpenRecordset("Users")
    
    RS.AddNew
    RS!UserName = "NEW"
    RS!member = True
    RS!Rank = 10
    RS!Password = "New"
    T = RS!ID
    
    RS.Update
  Else
    T = RS!ID
  End If
  
  
  ShowMemberEdit T
  
End Sub

Sub UpdateMember()

    pMember.ID = Val(GetCgiValue("member"))
    pMember.Admin = LCase$(GetCgiValue("admin")) = "on"
    pMember.AIM = GetCgiValue("aim")
    pMember.EMail = GetCgiValue("email")
    pMember.mMember = LCase$(GetCgiValue("ismember")) = "on"
    pMember.Password = GetCgiValue("pass")
    pMember.Rank = Val(GetCgiValue("rank"))
    pMember.UserName = GetCgiValue("username")
    pMember.nName = GetCgiValue("name")

    If pMember.UserName = mScreenName Then pMember.Admin = True
    

    If pMember.ID = CatID Or pMember.ID = MowadID Or pMember.ID = DutchID Or pMember.ID = KillerID Then _
      pMember.Admin = True

    If Trim$(pMember.UserName) = "" Then
      Call ShowMemberEdit(pMember.ID, "Please enter a valid UserName")
    ElseIf Trim$(pMember.Password) = "" Then
      Call ShowMemberEdit(pMember.ID, "Please enter a valid Password")
    Else
        
      Call InitDB
      Dim RS As Recordset
      
      Set RS = DB.OpenRecordset("Select * From Users Where ID=" & pMember.ID)
      
      With RS
        .Edit
        !Admin = pMember.Admin
        !AIM = pMember.AIM
        !EMail = pMember.EMail
        !member = pMember.mMember
        !Password = pMember.Password
        !Rank = pMember.Rank
        !UserName = pMember.UserName
        !Name = pMember.nName
        .Update
        
        Send "<BR><font class=ne><font color=red><B>Member Updated!</B><BR><BR>"
        Call SendMemberList(True)
        
      End With
        
    End If
End Sub

Sub ShowMemberEdit(vID As Integer, Optional formError As String)
  
  If formError = "" Then
  
    Call InitDB
    
    CurrSub = "ShowMemberEdit"

    Dim RS As Recordset
    Set RS = DB.OpenRecordset("SELECT * FROM Users WHERE ID=" & vID)
    
    Send "<!--" & RS!Admin & "-->"
    pMember.ID = vID
    pMember.Admin = (RS!Admin = "True")
    pMember.AIM = CheckForNulls(RS!AIM)
    pMember.EMail = CheckForNulls(RS!EMail)
    pMember.mMember = (RS!member = "True")
    pMember.Password = CheckForNulls(RS!Password)
    pMember.Rank = RS!Rank
    pMember.UserName = RS!UserName
    pMember.nName = CheckForNulls(RS!Name)
    
  End If
  
  Send " <font class=ne><BR></font>"

  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  
  Send "  <Input type=""hidden"" Name=""action"" value=""admconsole"">"
  Send "  <Input type=""hidden"" Name=""Member"" value=""" & pMember.ID & """>"
  Send "  <Input type=""hidden"" Name=""section"" value=""UpdateMember"">"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=500>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>Member Editor: " & pMember.UserName & "</TD></TR>"
  Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
  
  If formError <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & formError & "</TD></TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Screen Name: (<B>EXCLUDE</B> '[S.W.A.T]')<BR><Input Type=Text Name=""UserName"" Size=25 value=""" & pMember.UserName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Password:<BR><Input Type=password Name=""Pass"" Size=25 value=""" & pMember.Password & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>AIM Name:<BR><Input Type=Text Name=""aim"" Size=25 value=""" & pMember.AIM & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>E-Mail:<BR><Input Type=Text Name=""email"" Size=25 value=""" & pMember.EMail & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>Full Name:<BR><Input Type=Text Name=""name"" Size=25 value=""" & pMember.nName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  </TR>"
  
  Send "  <TR><TD valign=top align=center colspan=2 align=center Class=ne><HR></TD></TR>"
  
  Send "  <TR><TD valign=top align=center colspan=2 align=center Class=ne>"
  Send "    <TABLE><TR>"
  Send "    <TD Class=ne valign=top>"
  Send "    <input Type=Checkbox Name=""isMember""" & IIf(pMember.mMember, " CHECKED", "") & "> <font color=yellow>Active&nbsp;Member?"
  Send "    </TD>"
  Send "    <TD Class=ne valign=top>"
  Send "    &nbsp;&nbsp;&nbsp;<input Type=Checkbox Name=""Admin""" & IIf(pMember.Admin, " CHECKED", "") & "> <font color=yellow>Admin&nbsp;Access?"
  Send "    </TD></TR><TR>"
  Send "    <TD Class=ne Colspan=2 align=center valign=top>"
  Send "    <BR>&nbsp;&nbsp;&nbsp;<font color=yellow>Rank:</font>"
  Send "    <SELECT Name=""rank"">"
  
  Dim X As Integer
  Dim S As String
  For X = 0 To 10
    S = GetRank(X)
    If S <> "" Then
      If pMember.Rank = X Then
        Send "<option Value=""" & X & """ SELECTED>" & S
      Else
        Send "<option Value=""" & X & """>" & S
      End If
    End If
  Next
  
  Send "    </SELECT>"
  Send "    </TD>"
  
  
  Send "    </TR></TABLE><BR><BR>"
  Send "  </TD></TR>"
    
  Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Update""></TD></TR>"
  
  Send "  </TABLE>"
  Send "  </Form>"
  
End Sub

Sub DeleteNews(ID As Integer)
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT * FROM NEWS WHERE ID=" & ID)
  RS.Delete
  SendIndex
End Sub

Sub ProcessNewNews(Optional ID As Integer)
  
  uNews = GetCgiValue("news")
  
  If Trim$(uNews) = "" Then
    ShowMainConsoleMenu
    Exit Sub
  End If
  
  Call InitDB
  
  Dim RS As Recordset
  
  
  If ID > 0 Then
    Set RS = DB.OpenRecordset("SELECT * FROM News Where ID=" & ID)
    RS.Edit
    RS!News = GetCgiValue("news")
    RS.Update
    Send "<font class=ne><BR><B>News Edited Successfully</B><BR><BR></font>"
    SendIndex
  Else
    Set RS = DB.OpenRecordset("News")
    RS.AddNew
    RS!Postedby = mScreenName
    RS!PostedTime = Now
    RS!News = uNews
    RS.Update
    Send "<font class=ne><BR><B>News Posted Successfully</B><BR><BR></font>"
    ShowMainConsoleMenu
  End If
  
End Sub

Sub DeleteAbuse(AppNum As Integer)

  CurrSub = "DeleteAbuse"
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select ID from abuse Where ID=" & AppNum)
  
  RS.Delete
  Call ListAbuse(50)
    
End Sub

Sub DeleteApp(AppNum As Integer)
  
  Send "<!--DeleteApp-->"
  CurrSub = "DeleteApp"
  Call InitDB
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select ID from Applications Where ID=" & AppNum)
  
  RS.Delete
  Call ListApplications(50)
    
End Sub

Sub ListApplications(Optional Num As Integer)

  CurrSub = "ListApplications"
  Call InitDB
  Dim RS As Recordset
  
  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=600>"
  Send "  <TR>"
  Send "  <TD colspan=6 class=ne valign=top align=center>"
  
  If Num = 0 Then
    Send "  <B><font color=red><BR>ALL Membership Applications</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select * From Applications Order by Submitted desc")
  Else
    Send "  <B><font color=red><BR>" & Num & " Most Recent Membership Applications</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select Top " & Num & " * From Applications Order by Submitted desc")
  End If
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top Width=50>&nbsp;</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Name</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>CS Screenname</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Previous&nbsp;Clans</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Submitted:</TD>"
  Send "  </TR>"
  
  Dim X As Integer
  Dim BG As String
  
  If RS.RecordCount = 0 Then
    Send "<TR><TD Colspan=6 Align=center class=ne><BR><b><font color=white>No applications at this time.</TD></TR>"
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      BG = IIf(BG = "333333", "000000", "333333")
      X = X + 1
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne align=center rowspan=3 valign=top Width=50><b><font color=yellow>" & X & "</font><BR>"
      Send MeLink("Delete", , "appnumber=" & RS!ID & "&action=admconsole&section=deleteapp", True, True) & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Name, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!UserName, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><A Href=""mailto:" & RS!EMail & """><font color=white>" & RS!EMail & "</a></TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!PreviousClans, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Submitted, " ", "&nbsp;") & "</TD>"
      Send "  </TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne valign=top Colspan=5><font color=yellow>" & CheckForNulls(RS!Comments, True) & "</TD></TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne align=right valign=top Colspan=5>&nbsp;</TD></TR>"
      Send "<TR bgcolor=" & BG & "><TD class=ne align=right valign=top Colspan=6>"
      Send "    <hr color=336699>"
      Send "  </TD></TR>"
      RS.MoveNext
    Loop
  End If
  Send "  </TABLE>"
End Sub



Public Sub ShowMainConsoleMenu()

  Dim ARCount As Integer
  Dim APPCount As Integer
  Dim RS As Recordset
  
  Call InitDB

  Set RS = DB.OpenRecordset("SELECT * From Applications")
  If RS.RecordCount = 0 Then
    APPCount = 0
  Else
    RS.MoveLast
    APPCount = RS.RecordCount
  End If
  RS.Close

  Set RS = DB.OpenRecordset("SELECT * From Abuse")
  If RS.RecordCount = 0 Then
    ARCount = 0
  Else
    RS.MoveLast
    ARCount = RS.RecordCount
  End If
  RS.Close

  CurrSub = "ShowMainconsole"
  Send "  <font class=ne><BR><BR></font><TABLE CellPadding=2 CellSpacing=0 Border=0>"
  Send "  <TR>"
  Send "  <TD class=ne valign=top>"
  
  'Misc Table
  Send "    <TABLE CellPadding=2 CellSpacing=0 Border=0>"
  Send "    <TR><TD colspan=2 Class=ne><font color=white><B>Misc. Stuff</td></tr>"
  
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/glass.gif""></TD><TD Class=ne>" & MeLink("View Applications (<font color=white>" & APPCount & "</font>)", "red", "action=admConsole&Section=ViewApps", , True) & "</td></tr>"
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/glass.gif""></TD><TD Class=ne>" & MeLink("View Abuse Reports (<font color=white>" & ARCount & "</font>)", "red", "action=admConsole&Section=ViewAbuse", , True) & "</td></tr>"
  Send "    <TR><TD Class=ne>&nbsp;<TD Class=ne>&nbsp;</td></tr>"
  
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/news.gif""></TD><TD Class=ne>" & MeLink("Post News", "red", "action=admConsole&Section=PostNews", , True) & "</td></tr>"
  Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
  
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/myprofile.gif""></TD><TD Class=ne>" & MeLink("Modify My Profile", "red", "action=admconsole&section=editprofile&member=" & MYID, , True) & "</TD></TR>"
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/profile.gif""></TD><TD Class=ne>" & MeLink("Modify Member Profile", "red", "action=admconsole&section=editprofile", , True) & "</TD></TR>"
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/profile.gif""></TD><TD Class=ne>" & MeLink("Add a Member", "red", "action=admconsole&section=addprofile", , True) & "</TD></TR>"
  Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
  
  Send "    <TR><TD Class=ne><IMG src=""http://www.jasongoldberg.com/swat/images/fix.gif""></TD><TD Class=ne>" & MeLink("Force Mail Cleanup", "red", "action=admConsole&Section=cleanmail", , True) & "</td></tr>"
  Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
  
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  <TD class=ne valign=top Width=50>"
  Send "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  Send "  </TD>"

  
  Send "  <TD class=ne valign=top>"
  
    'Users Table
    Send "    <TABLE CellPadding=2 CellSpacing=0 Border=0>"
    Send "    <TR><TD colspan=2 Class=ne><font color=white><B>Users / Memebers</td></tr>"
    Send "    <TR><TD Class=ne>"
    Send "      <TABLE><TR><TD Class=ne Valign=Middle><font color=red>Search Users:</B></TD></TR><TR><TD Class=ne Valign=Middle>"
    Send "      <form action=""" & EXEPath & "index.exe"" Method=post>"
    Send "      <Input Type=""Hidden"" Name=""Action"" Value=""admConsole"">"
    Send "      <Input Type=""Hidden"" Name=""section"" Value=""searchusers"">"
    Send "      <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "      <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "      <input Type=Text Size=25 name=""search"" Value=""" & GetCgiValue("search") & """>"
    Send "      </form>"
    Send "      </TD></TR></TABLE>"
    Send "    </td>"
    Send "    </tr>"
    Send "    </TABLE>"
  

  
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub ProcessApplication()
CurrSub = "ProcessApplication"
  On Error GoTo ErrPoint
  Dim RS As Recordset
  
  Call InitDB
  
  With Application
  
    .Comments = Trim(GetCgiValue("comments"))
    .EMail = Trim(GetCgiValue("email"))
    .IPAddress = CGI_RemoteAddr
    .Name = Trim(GetCgiValue("name"))
    .UserName = Trim(GetCgiValue("UserName"))
    .PreviousClans = Trim(GetCgiValue("Previous"))
    .SubmmittedTime = Now
    
    
    Set RS = DB.OpenRecordset("Select * From Applications Where IP='" & .IPAddress & "'")
    
    If RS.RecordCount <> 0 Then
      
      Call sendApplication("An application from <font color=white>'" & .IPAddress & "</font>' has already been submitted.", True)
        
    Else
    
      If Len(.Name) < 2 Or Len(.Name) > 70 Then
        Call sendApplication("Please enter a valid name (2-70 chars).")
        Exit Sub
      End If
    
      If Len(.EMail) < 2 Or Len(.EMail) > 90 Or InStr(1, .EMail, "@") = 0 Or InStr(1, .EMail, ".") = 0 Then
        Call sendApplication("Please enter a valid e-mail address (2-90 chars).")
        Exit Sub
      End If
    
      If Len(.UserName) <= 2 Or Len(.UserName) > 50 Then
        Call sendApplication("Please enter a valid user name (2-50 chars).")
        Exit Sub
      End If
    
      Dim RS2 As Recordset
      Set RS2 = DB.OpenRecordset("Applications")
      RS2.AddNew
      RS2!Name = .Name
      RS2!EMail = .EMail
      RS2!PreviousClans = .PreviousClans
      RS2!UserName = .UserName
      RS2!Comments = .Comments
      RS2!IP = .IPAddress
      RS2!Submitted = .SubmmittedTime
      RS2.Update
    
      Send "<font class=ne><BR><BR><font color=white><B>Thank You for Applying to [S.W.A.T]</B><BR><BR>"
      Send "Your Application will be processed, and you will be contacted when a decision has been reached.<BR><BR>"
    
    End If
    
    
  End With
  
  
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub

Sub sendApplication(Optional errorMessage As String, Optional SkipBody As Boolean)
CurrSub = "SendApplication"
  Send " <font class=ne><BR></font>"
  If Not SkipBody Then
    
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""hidden"" Name=""action"" value=""submitapplication"">"
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=400>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>So You Wanna Be [S.W.A.T]?</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
    
    If errorMessage <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & errorMessage & "</TD></TR>"
    
    Send "  <TR>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Name<BR><Input Type=Text Name=""Name"" Size=25 value=""" & Application.Name & """></TD>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Your E-Mail Address<BR><Input Type=Text Name=""EMail"" Size=25 value=""" & Application.EMail & """></TD>"
    Send "  </TR>"
    
    Send "  <TR>"
    Send "  <TD valign=top Width=400 Class=ne><font color=white>Name used on [S.W.A.T] Server<BR><Input Type=Text Name=""Username"" Size=25 value=""" & Application.UserName & """></TD>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Previous Clan(s)<BR><Input Type=Text Name=""Previous"" Size=25 value=""" & Application.PreviousClans & """><BR><BR><BR></TD>"
    Send "  </TR>"
    
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>Comments:<BR><TextArea Name=""Comments"" Cols=40 Rows=8>" & Application.Comments & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Application""></TD></TR>"
  
  Else
    If errorMessage <> "" Then Send "  <font Class=ne><font color=yellow><B>" & errorMessage & "<BR></font>"
  End If
  
  Send "  </TABLE>"
  Send "  </Form>"

End Sub

Sub SendAbuseForm(Optional errorMessage As String)
CurrSub = "SendAbuseForm"
  Send " <font class=ne><BR></font>"
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""hidden"" Name=""action"" value=""submitabuse"">"
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=400>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>So You Wanna complain about [S.W.A.T]?</TD></TR>"
  Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
  
  If errorMessage <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & errorMessage & "</TD></TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your CS Name<BR><Input Type=Text Name=""UserName"" Size=25 value=""" & Abuse.UserName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your E-Mail Address<BR><Input Type=Text Name=""EMail"" Size=25 value=""" & Abuse.EMail & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD colspan=2 valign=top Width=400 align=center Class=ne><font color=white>[S.W.A.T] Member In Question:<BR>"
  Call MemberCombo(Abuse.MemberName, , "N/A")
  Send "  <BR><BR></TD>"
  Send "  </TR>"
  
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>Register Your Complaint:<BR><TextArea Name=""Comments"" Cols=40 Rows=8>" & Abuse.Comments & "</TEXTAREA></TD></TR>"
  Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Complaint ""></TD></TR>"

  Send "  </TABLE>"
  Send "  </Form>"

End Sub

Sub MemberCombo(Optional MemberToSelect As String, Optional OmitSWAT As Boolean, Optional FirstSelection As String, Optional UseIDs As Boolean, Optional SecondSelection As String)
  CurrSub = "MemberCombo"
  On Error GoTo Err
  
  Dim RS As Recordset
  Dim BG As String
  
  Call InitDB
  Set RS = DB.OpenRecordset("Select * From Users Where Member Order By Rank, username")
  
  MemberToSelect = Trim(MemberToSelect)
  
  Send "<Select Name=""Membername"">"
  
  If FirstSelection <> "" Then Send "<Option Value=""" & FirstSelection & """" & IIf(MemberToSelect = "", " SELECTED", "") & ">" & FirstSelection
  If SecondSelection <> "" Then Send "<Option Value=""" & SecondSelection & """>" & SecondSelection
  
  RS.MoveFirst
  Do While Not RS.EOF
    If LCase$(RS!UserName) <> "new" Then
      
      If UseIDs Then
        If RS!ID = Val(MemberToSelect) And Val(MemberToSelect) <> 0 Then
          Send "<Option Value=""" & RS!ID & """ SELECTED>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!UserName
        Else
          Send "<Option Value=""" & RS!ID & """>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!UserName
        End If
      Else
        If RS!UserName = MemberToSelect And MemberToSelect <> "" Then
          Send "<Option Value=""" & RS!UserName & """ SELECTED>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!UserName
        Else
          Send "<Option Value=""" & RS!UserName & """>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!UserName
        End If
      End If
    End If
    RS.MoveNext
  Loop
  
  Send "</Select>"
  
  Exit Sub
Err:
  Send "Error " & Err.Number & ": " & Err.Description
End Sub

Sub ProcessAbuse()
CurrSub = "ProcessAbuse"
  On Error GoTo ErrPoint
  Dim RS As Recordset
  
  Call InitDB
  
  With Abuse
  
    .Comments = Trim(GetCgiValue("comments"))
    .EMail = Trim(GetCgiValue("email"))
    .UserName = Trim(GetCgiValue("username"))
    .MemberName = Trim(GetCgiValue("memberName"))
    .vDate = Now
      
    If Len(.UserName) < 2 Or Len(.UserName) > 50 Then
      Call SendAbuseForm("Please enter a valid name (2-50 chars).")
      Exit Sub
    End If
  
    If Len(.EMail) < 2 Or Len(.EMail) > 90 Or InStr(1, .EMail, "@") = 0 Or InStr(1, .EMail, ".") = 0 Then
      Call sendApplication("Please enter a valid e-mail address (2-90 chars).")
      Exit Sub
    End If
  
    If Len(.MemberName) < 2 Or Len(.MemberName) > 70 Then
      Call SendAbuseForm("Please select a [S.W.A.T] member, or select N/A.")
      Exit Sub
    End If
  
    If Len(.Comments) < 10 Or Len(.MemberName) > 700 Then
      Call SendAbuseForm("Please enter a valid comaplaint (10-700 Chars).")
      Exit Sub
    End If
  
    Set RS = DB.OpenRecordset("Abuse")
    RS.AddNew
    RS!member = .MemberName
    RS!EMail = .EMail
    RS!Date = .vDate
    RS!User = .UserName
    RS!Complaint = .Comments
    RS.Update
  
    Send "<font class=ne><BR><BR><font color=white><B>Your complaint has been registered.</B><BR><BR>"
    Send "While you probably deserved whatever you got, we will review your complaint, check the server logs, and proceed accordingly.<BR><BR>"
    
  End With
  
  
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub


Sub SendMemberList(Optional AdminEdit As Boolean)
  
  CurrSub = "SendMemeberList"
  
  On Error GoTo Err
  
  Dim RS As Recordset
  Dim BG As String
  
  Call InitDB
  
  If AdminEdit Then
    Set RS = DB.OpenRecordset("Select * From Users Order By Rank, username")
  Else
    Set RS = DB.OpenRecordset("Select * From Users Where Member Order By Rank, username")
  End If
  
  Send " <font class=ne><BR></font>"
  If AdminEdit Then
    Send " <font class=ne><font color=yellow>Click the member you would like to edit (Red = Inactive Member)<BR><BR></font>"
  End If
  Send "  <TABLE CellPadding=3 bordercolor=336699 CellSpacing=0 Border=1>"
  Send "  <TR bgcolor=fffff>"
  Send "  <TD Class=ne><font color=000000><B>Member Name</TD>"
  
  If AdminEdit = False Then
    Send "  <TD Class=ne><font color=000000>&nbsp;&nbsp;<B>Rank</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Email</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>AIM</TD>"
  End If
  If RS.RecordCount = 0 Then
  
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      If LCase(RS!UserName) = "new" And AdminEdit = False Then GoTo 10
      BG = IIf(BG = "000000", "333333", "000000")
      
      If AdminEdit And Not (RS!member) Then BG = "660000"
      
      If AdminEdit Then
        Send "  <TR BGColor=" & BG & "><TD Class=yellownames>" & MeLink("[S.W.A.T] " & RS!UserName, "Yellow", "action=admconsole&section=editprofile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
      Else
        Send "  <TR BGColor=" & BG & "><TD Class=yellownames><font color=yellow>[S.W.A.T] " & RS!UserName & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
      End If
            
      If AdminEdit = False Then
        Send "<TD Class=ne>&nbsp;&nbsp;" & GetRank(RS!Rank) & "</TD>"
        If IsNull(RS!EMail) Then
          Send "    <TD Class=ne align=center><font color=white>?</TD>"
        ElseIf Len(RS!EMail) = 0 Then
          Send "    <TD Class=ne align=center><font color=white>?</TD>"
        Else
          Send "    <TD Class=ne align=center><A href=""mailto:" & RS!EMail & """><font color=white>" & CheckForNulls(RS!EMail, True) & "</a></TD>"
        End If
        
        If IsNull(RS!AIM) Then
          Send "    <TD Class=ne align=center><font color=white>?</TD>"
        ElseIf Len(RS!AIM) = 0 Then
          Send "    <TD Class=ne align=center><font color=white>?</TD>"
        Else
          Send "    <TD class=ne align=center><A href=""aim:goim?screenname=" & Replace(RS!AIM, " ", "") & """><font color=white>" & CheckForNulls(RS!AIM, True) & "</a></TD></TR>"
        End If
      End If
10:
      RS.MoveNext
            
    Loop
  End If
  
  Send "  </TABLE>"
  
  
  Exit Sub
Err:
  Send "Error " & Err.Number & ": " & Err.Description
End Sub

Sub SendServerStats()
  CurrSub = "SendServerStats"
  Send "  <TABLE Border=0 CellPadding=0 CellSpacing=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne valign=top>"
  
    Call SendWhosOnline(True)
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub SendIndex()
CurrSub = "SendIndex"
  Send " <font class=ne><BR></font>"
  Send "  <TABLE Border=0 CellPadding=0 CellSpacing=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne valign=top Width=130>"
  
    Call SendWhosOnline
  
  Send "  </TD>"
  Send "  <TD Class=ne align=center valign=top Width=620>"
  
    Call ShowNews
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Public Function MeLink(Text As String, Optional color As String, Optional EndofLink As String, Optional LinkUnderline As Boolean, Optional KeepLoginData As Boolean) As String
  Dim i As String
  CurrSub = "MeLink"
  Dim S As String
  Dim T As String
  
  i = Format(Abs(DateDiff("s", Now, "01/01/2004")), "00000000000000000000000000000")
  
  If LinkUnderline = False Then T = "Style=""text-decoration:none"""
  If KeepLoginData Then S = "&screenname=" & mScreenName & "&password=" & Encrypt(mPassWord)
  
  If color = "" Then color = "336699"
  
  MeLink = "<A " & T & " Href=""" & EXEPath & "index.exe?" & EndofLink & S & "&refresh=" & i & """><font color=" & color & ">" & Text & "</font></A>"
  
End Function

Sub ShowNews()
  CurrSub = "ShowNews"
  Call InitDB
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT TOP 3 * From News Order By Postedtime desc")
  
  RS.MoveFirst
  Do While Not RS.EOF
  
    Send "    <TABLE Border=0 CellPadding=1 CellSpacing=0 Width=550>"
    Send "    <TR bgcolor=336699 align=center><TD align=left width=475 Class=ne><font color=white>&nbsp;&nbsp;<b>.: [S.W.A.T] " & RS!Postedby & "</TD><TD Class=ne align=right><font color=white>" & Replace(Format(Trim(RS!PostedTime), "mm/dd hh:mm AMPM"), " ", "&nbsp;") & "</font></TD></tr>"
    Send "    <TR>"
    Send "    <TD colspan=2 bgcolor=333333 Class=ne valign=top align=center>"
    Send "    <TABLE Width=95%><TR><TD Class=ne><font color=white>"
    Send RS!News
    Send "</TD></TR></TABLE>"
    Send "    </font>"
    Send "    </TD></TR>"
    
    If LoginStatus >= 4 Then
      Send "    <TR><TD colspan=2 bgcolor=black align=right class=ne><B>" & MeLink("Edit", "yellow", "action=admconsole&section=editnews&ID=" & RS!ID, False, True) & "&nbsp;&nbsp;|&nbsp;&nbsp;" & MeLink("Delete", "yellow", "action=admconsole&section=deletenews&ID=" & RS!ID, False, True) & "</TD></TR>"
    End If
    
    Send "    <TR><TD colspan=2 bgcolor=black valign=top><HR Color=9C1100></TD></TR>"
    Send "    </TABLE><BR>"
    
    RS.MoveNext
  Loop
  
End Sub


Sub ListAbuse(Optional Num As Integer)
  CurrSub = "ListAbuse"
  Call InitDB
  Dim RS As Recordset
  
  
  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=600>"
  Send "  <TR>"
  Send "  <TD colspan=5 class=ne valign=top align=center>"
  
  If Num = 0 Then
    Send "  <B><font color=red><BR>ALL Abuse Reports</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select * From Abuse Order by date desc")
  Else
    Send "  <B><font color=red><BR>" & Num & " Most Recent Abuse Reports</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select Top " & Num & " * From Abuse Order by date desc")
  End If
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top Width=50>&nbsp;</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>CS Screenname</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Member</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Submitted</TD>"
  Send "  </TR>"
  
  Dim X As Integer
  Dim BG As String
  
  If RS.RecordCount = 0 Then
    Send "<TR><TD Colspan=5 Align=center class=ne><BR><b><font color=white>No abuse reports at this time.</TD></TR>"
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      BG = IIf(BG = "333333", "000000", "333333")
      X = X + 1
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne align=center rowspan=2 valign=top Width=50><b><font color=yellow>" & X & "</font><BR>"
      Send MeLink("Delete", , "appnumber=" & RS!ID & "&action=admconsole&section=deleteabuse", True, True) & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!User, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><A Href=""mailto:" & RS!EMail & """><font color=white>" & RS!EMail & "</a></TD>"
      Send "  <TD class=ne valign=top><font color=white>" & IIf(LCase$(RS!member) = "n/a", "", "[S.W.A.T] ") & Replace(RS!member, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Date, " ", "&nbsp;") & "</TD>"
      Send "  </TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne valign=top Colspan=4><font color=white>" & RS!Complaint & "</TD></TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne valign=top Colspan=5><hr color=336699>"
      Send "  </TD></TR>"
      RS.MoveNext
    Loop
  End If
  Send "  </TABLE>"
End Sub

Sub ShowMailIndex(Optional MB As Integer)
  Call InitDB
  CurrSub = "ShowMailIndex"
  Dim Mailbox As Integer
  Dim M As Integer
  Dim FolderTitle As String
  Dim BG As String
  Dim RS As Recordset
  
  Mailbox = Val(GetCgiValue("mailbox"))
  
  If MB <> 0 Then Mailbox = MB
  
  If Mailbox = INBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=FALSE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Inbox"
  ElseIf Mailbox = SENTBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE FROM='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Sent Messages"
  ElseIf Mailbox = TRASHBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=TRUE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Trashed Messages"
  End If
  
  Send "<!--m:" & M & "-->"
  M = RS.RecordCount
  Send "<!--m:" & M & "-->"
  Send "<TABLE CellSpacing=0 border=0 Width=650>"
  Send "<TR>"
  Send "<TD Class=ne><BR>"
  Send "  <TABLE CellSpacing=0 border=0 width=100%>"
  Send "  <TR><TD Colspan=2 class=bigheading align=center>.: My " & FolderTitle & " (" & M & "):.</TD></TR>"
  If Mailbox = SENTBOX Then Send "  <TR><TD Colspan=2 class=ne align=center><font color=red><B><BR>Note: These messages are deleted when the recipient deletes them.<BR><BR></TD></TR>"
  Send "  <TR>"
  Send "  <TD Class=ne>"
  Send MeLink("Compose Message", "Yellow", "Action=composemail", True, True)
  Send "  </TD>"
  Send "  <TD Class=ne align=right>"
  
  Send "    <TABLE CellSpacing=0 border=0>"
  Send "    <TR>"
  
  Dim A As Integer
  Dim S As Integer
  Dim T As Integer
  
  A = -1
  S = -1
  T = -1
  
  Call GetMailCount(A, S, T)
  
  CurrSub = "ShowMailIndex"
  
  If Mailbox <> INBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send "    <TD Class=ne>"
    Send MeLink("Inbox (" & A & ")", "Yellow", "Action=MailIndex", True, True)
    Send "    </TD>"
  End If
  
  If Mailbox <> SENTBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send MeLink("View Sent (" & S & ")", "Yellow", "action=MailIndex&mailbox=" & SENTBOX, True, True)
    Send "    </TD>"
  End If
    
  If Mailbox <> TRASHBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send "    <TD Class=ne>"
    Send MeLink("View Trash (" & T & ")", "Yellow", "Action=MailIndex&mailbox=" & TRASHBOX, True, True)
    Send "    </TD>"
  End If
  
  Send "    </TR>"
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE><BR>"
  
  Send "</TD>"
  Send "</TR>"
  Send "</TABLE>"
  
  Send "<TABLE bgcolor=""333333"" CellSpacing=0 border=0 Width=650>"
  Send "<TR BGCOLOR=FFFFFF><TD Class=ne>&nbsp</TD>"
  
  If Mailbox = SENTBOX Then
    Send "<TD Class=ne><B><font color=black>To</TD>"
  Else
    Send "<TD Class=ne><B><font color=black>From</TD>"
  End If
  
  Send "<TD class=ne><B><font color=black>Subject</TD>"
  Send "<TD class=ne align=right><B><font color=black>Date/Time</TD><TD Class=ne>&nbsp;</TD></TR>"
  
  
  
  If M > 0 Then
    On Error GoTo Err
    RS.MoveFirst
    Do While Not RS.EOF
      Send "<TR><TD Colspan=4>&nbsp;</TD></TR>"
      Send "<TR><TD width=25 Class=ne valign=middle>" & IIf(RS!read, "&nbsp", "<IMG Src=""http://www.jasongoldberg.com/swat/images/new.gif"">") & "</TD>"
      
      If Mailbox = SENTBOX Then
        Send "<TD Class=ne><font color=white><B>" & GetName(Val(RS!To)) & "</TD>"
      Else
        Send "<TD Class=ne><font color=white><B>" & GetName(Val(RS!From)) & "</TD>"
      End If
      
      Send "<TD class=ne><font color=Yellow>" & IIf(RS!read, "", "<B>") & MeLink(RS!Subject, "Yellow", "action=readmail&id=" & RS!ID, True, True) & "</TD>"
      Send "<TD class=ne align=right><font color=999999>" & Format(RS!sent, "mm.dd.yyyy - hh:mm AMPM") & "</TD>"
      Send "</TD>"
      Send "<TD class=ne align=right>"
      Send "    <TABLE cellpadding=0 cellspacing=0>"
      Send "    <TR>"
      If Mailbox <> SENTBOX Then
        If Mailbox = INBOX Then Send "    <TD Class=ne>" & MeLink("Reply", "white", "Action=replymail&Id=" & RS!ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
        If Mailbox = TRASHBOX Then Send "    <TD Class=ne>" & MeLink("Back To Inbox", "white", "Action=restoremail&Id=" & RS!ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
        Send "    <TD Class=ne>" & MeLink("Delete", "white", "Mailbox=" & Mailbox & "&Action=deletemail&Id=" & RS!ID, True, True) & "</TD>"
      Else
        Send "    <TD Class=ne>&nbsp;</TD>"
      End If
      Send "    </TR>"
      Send "    </TABLE>"
      Send "</TD>"
      Send "</TR>"
      RS.MoveNext
    Loop
    Send "<TR><TD Colspan=4>&nbsp;</TD></TR>"
    Send "</TABLE>"
  Else
Err:
    Send "<TR><TD Colspan=4 align=center>"
    Send "<font class=ne><font color=ffffff><B>Sorry, No New Messages</Font>"
    Send "&nbsp;</TD></TR>"
  End If
End Sub
