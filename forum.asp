<!--#include file="conn.asp"-->
<%
'dim in_count,str_rnd
randomize timer
str_rnd=""
  Rstring="abcdefghijklmnopqrstuvwxyz0123456789"
while in_count<>25 
 str_rnd=str_rnd&mid(Rstring,int(rnd(1)*36)+1,1)
in_count=in_count+1
wend
Dim Fo
Set Fo = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")
response.write "<HTML><HEAD><TITLE>"
sql="select * from config_setting"
set rstitle=conn.execute(sql)
if rstitle("en_reg")=true then%>
<!--#include file="pow.asp"-->
<%end if
page_num=rstitle("reply_counts")
response.write rstitle("bbsname")
response.write "</TITLE>"
response.write "</HEAD>"%>
<link rel="stylesheet" type="text/css" href=style1.css>
<%tid=request("tid")
UserID=request.cookies("UserID")
response.cookies("tid")=tid
Set rs=Server.CreateObject("ADODB.Recordset")

SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
set rs11=conn.execute(SQL)

'密碼確定
'*******************************************************
Set rs_forum=Server.CreateObject("ADODB.Recordset")
rs_forum.Open "Select * From tid Where tid="&request("tid"),conn,1,3

forum_cookies=bbsid+Cstr(tid)+request.cookies("UserID")
if request.cookies(forum_cookies)="yes" or  rs_forum("have_pass")=0 then

'*********************刪除文章程序****************
if rsidid("en_save")=true then

Set rs_delforum=Server.CreateObject("ADODB.Recordset")
SQL="Select * From Titles where tid="&tid&" and best<>1 and top_lock<>1"
rs_delforum.Open SQL,conn,1,3

While not rs_delforum.eof
if rs_delforum("save_date")<>empty then
if Datevalue(Dateadd("d",rs_delforum("save_date"),rs_delforum("CreateDate")))<=date then
   Set rs_del= Server.CreateObject("ADODB.Recordset")
   sql="select * from Details where TitleID="&rs_delforum("TitleID")
   rs_del.Open sql, conn,1,3
   while not rs_del.EOF
     if rs_del("file")<>empty then
		master123=split(rs_del("file"),",")
		for i = 0 to ubound(master123)
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("./upload/"&master123(i))) then
		FileName = "upload/" & Trim(master123(i))'刪除回覆文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
		next	
     end if
     rs_forum("replycounts")=rs_forum("replycounts")-1
     rs_forum.update
     rs_del.delete
     rs_del.movenext
   wend
     if rs_delforum("file")<>empty then
		master123=split(rs_delforum("file"),",")
		for i = 0 to ubound(master123)
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("./upload/"&master123(i))) then
		FileName = "upload/" & Trim(master123(i))'刪除文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
		next	
     end if
   rs_forum("counts")=rs_forum("counts")-1
   rs_forum.update
   rs_delforum.delete
end if
end if
	rs_delforum.MoveNext
Wend
end if
'*********************************************end**

if request("mod")="best" then
SQL="Select * From Titles where tid="&tid&" and best=1 and top_lock=0 order By LastNewsDate Desc" 
else
SQL="Select * From Titles where tid="&tid&" and top_lock=0 order By LastNewsDate Desc" 
end if
rs.Open  SQL,conn,1,1

'*******************************************************
response.write "<BODY leftmargin=0 topmargin=0>"%>
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"-->
<%response.write "<CENTER> <FONT COLOR=RED size=-1>"&Request("msg")&"</FONT></CENTER>"
response.write "<div align=left><div align=right><table border=0 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor=#000000 width=325 height=18 id=AutoNumber3><tr><td width=103 height=18></td><td width=1 height=18 bordercolor=#000000></td><td width=137 height=18></td><td width=30 height=18></td><td width=41 height=18></td></tr></table></div>"
Page = cint(Request("Page"))
myself=request.serverVariables("PATH_INFO")
rs.PageSize =15       ' 設定每頁顯示 15 筆
If Page<1 Then 
   Page=1
ELSEIF Page>rs.PageCount Then 
   Page=rs.PageCount
END IF 
response.write "<center>"
response.write "<table border=0 cellpadding=0 cellspacing=0 style=border-collapse: collapse bordercolor=#111111 width=96% height=28 id=AutoNumber8>"
response.cookies("tid")=tid%>
<!--#include file="choice.asp"-->
<%response.write "<tr><td width=363 height=28><p align=left><a href=newpost.asp?tid="&Request("tid")&"><img border=0 src=images/send.gif align=left></a><a href=newpoll.asp?tid="&Request("tid")&"><img border=0 src=images/newpoll.gif align=left></a></td><td width=24 height=28></td><td width=22 height=28><p></p><p style=margin-bottom: 0></td><td width=294 height=28>"%>
<p align="right"><img border="0" src="images/icon.gif"><a href=forum.asp?tid=<%=request("tid")%>&mod=best><font size="2">精華文章</font></a>&nbsp;&nbsp;&nbsp; <img src="images/manager.gif">
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option selected>本區管理版主...</option>
<%sql="select * from tid where tid="&tid
set rs_admin=conn.execute(sql)
tidadmin=split(rs_admin("tidadmin"),"|")
for i = 1 to ubound(tidadmin)-1%>
<option value=u2u.asp?touser=<%=tidadmin(i)%>><%=tidadmin(i)%></option>
<%next%>
</select></p>
<%response.write "</td></tr></table>"%><!--#include file="user_list_forum.asp"-->  
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse;WORD-BREAK: break-all" bordercolor="#000000" width="98%" height="46" id="AutoNumber1" bordercolorlight="#FFFFFF" bordercolordark="#E3E3E3">
<tr> 
<td bordercolor="#000000" align="center" background="images/Button.gif" height="27">
<font size="2">狀態</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" height="27" width="50%">
<font size="2">主題</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" height="27">
<font size="2">作者</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" height="27">
<font size="2">點閱數</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" height="27">
<font size="2">回覆數</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" bgcolor="#000000" height="27">
<font size="2">最後更新</font></td>
<td bordercolor="#000000" align="center" background="images/Button.gif" bgcolor="#EFF4FE" height="28">
<font size="2">關注</font></td>
</tr>
<%'置頂****************************************************************************
IF Page=1 then
sql="select * from Titles where tid=" & tid & " and top_lock=1 order by top_lock_order desc,LastNewsDate desc"
set rs_lock=conn.execute(sql)
if not rs_lock.EOf then
while not rs_lock.eof%>
<tr>
<td width="30" height="18" bgcolor="#E3E3E3" bordercolor="#000000" align="center">
<%if rs_lock("reply_lock")=1 then
    response.write "<img src=""images/lock.gif"">"
elseif rs_lock("count")>50 then
    response.write "<img src=""images/hot.gif"">"
elseIF Now - rs_lock("LastNewsDate") < 1 then 
 	response.write "<img src=""images/icon_folder_new_3.gif"">"
else
    response.write "<img src=""images/icon_folder_3.gif"">"
end if%>
</td>
<td width="50%" height="18" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif"><p align="left">
<%if rs_lock("top_lock_order")=3 then
response.write "<img src=""images/locktop.gif""><font size=-1 color=red>【置頂III】</font>"
elseif rs_lock("top_lock_order")=2 then
response.write "<img src=""images/locktop.gif""><font size=-1 color=red>【置頂II】</font>"
elseif rs_lock("top_lock_order")=1 then
response.write "<img src=""images/locktop.gif""><font size=-1 color=red>【置頂I】</font>"
end if%>
<img src="images/post/<%=rs_lock("R1")%>.gif" width=13 height=13 border=0>
<%if rs_lock("best")=1 then%>
<img src="images/best.gif">
<%end if%>
<a href=Detail.asp?TitleID=<%=rs_lock("TitleID")%>&tid=<%=Request("tid")%>&postname=<%=rs_lock("Name")%>&noc=<%=str_rnd%>><font color="black"size=-1><font color=black><%=rs_lock("subc")%></font><%=rs_lock("Subject")%></font></a><%if rs_lock("poll")=1 then%><img src="images/sel.jpg"><%end if%>
<%IF (Now - rs_lock("LastNewsDate"))<0.5 then
      response.write "<img src=""images/new.gif"">"		 
end if
mxPages = rs_lock("replycount")/page_num
if mxPages<>CInt(mxPages) then
mxpages=Int(mxPages)+1
end if

if (mxPages>1 and mxPages<=12) then 
Response.Write("<br>　　<font size=-1>第")
for counter = 1 to mxPages
ref = "&nbsp;<a href='Detail.asp?"
ref = ref & "tid=" & Request("tid")
ref = ref & "&TitleID=" & rs_lock("TitleID")
ref = ref & "&postname=" & rs_lock("Name")
ref = ref & "&Page=" & counter
ref = ref & "'><font color=ff0000><b>" & counter & "</font></b></a>"
response.write ref
next
Response.Write("&nbsp;頁。</font>")
elseif mxPages>12 then
Response.Write("<br>　　<font size=-1>第")
for counter = 1 to 12
ref = "&nbsp;<a href='Detail.asp?"
ref = ref & "tid=" & Request("tid")
ref = ref & "&TitleID=" & rs_lock("TitleID")
ref = ref & "&postname=" & rs_lock("Name")
ref = ref & "&Page=" & counter
ref = ref & "'><font color=ff0000><b>" & counter & "</font></b></a>"
response.write ref
next
response.write "......."
las = "&nbsp;<a href='Detail.asp?"
las = las & "tid=" & Request("tid")
las = las & "&TitleID=" & rs_lock("TitleID")
lasf = las & "&postname=" & rs_lock("Name")
las = las & "&Page=" & mxPages
las = las & "'><font color=ff0000><b>" & mxPages & "</font></b></a>"
response.write las
Response.Write("&nbsp;頁。</font>")
end if		   '特殊選項 
response.write "<font size=-1>"
Set passed3rs = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From tid where tid="&request("tid")
passed3rs.Open SQL, conn,1,1
if passed3rs("tidadmin")<>"" then
tidadmin=split(passed3rs("tidadmin"),"|")
 
for i = 1 to ubound(tidadmin)-1 '檢驗版主
 if tidadmin(i)=UserID and request.cookies("admin_pass")<>bbsid and rs_lock("Name")<>UserID then
  Response.Write " <br>  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a href=delete.asp?TitleID="& rs_lock("TitleID")&"&tid="&tid&">" & "刪除</a>"
  Response.Write "<a href=edit.asp?TitleID="& rs_lock("TitleID")&"&postname="&rs_lock("Name")&"&tid="&request("tid")&"&url=index.asp"&">" & "|&nbsp;編輯</a>"
 end if
next
end if

If request.cookies("admin_pass")=bbsid Or rs_lock("Name")= UserID Then
Response.Write " <br>  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a href=delete.asp?TitleID="& rs_lock("TitleID")&"&tid="&tid&">" & "刪除</a>"
Response.Write "<a href=edit.asp?TitleID="& rs_lock("TitleID")&"&postname="&rs_lock("Name")&"&tid="&request("tid")&"&url=index.asp"&">" & "|&nbsp;編輯</a>"
End If %></font></p></td>
<td width="130" height="18" bgcolor="#FFFFFF" align="center" bordercolor="#000000"> 
<font color="#000080" size="2"><a href="edituser.asp?UserID=<%=rs_lock("Name")%>"><%=rs_lock("Name")%></a></font></td>
<td width="50" height="18" bgcolor="#F7F5EE" align="center" bordercolor="#000000" background="images/postbit_lightbg.gif"> 
<font size="2">　<%=rs_lock("count")%></font></td>
<td width="50" height="18" align="center" bordercolor="#000000"> 
<font size="2">　<%=rs_lock("replycount")%></font></td>
<td width="160" height="18" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif"> <p align="right"><font color="black" size="2">
				<%if rs11 Is Nothing Or rs11.EOF Then%>  
				<%=DateAdd("h",rstitle("servertimezone"),rs_lock("LastNewsDate"))%>
				<%else%>
				<%=DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs_lock("LastNewsDate"))%>
				<%end if%>
<%sql="select * from Details where TitleID="&rs_lock("TitleID")&" order by DetailID Desc"
	set sortlast=conn.execute(sql)%>
				<br>by&nbsp;<a href="edituser.asp?UserID=<%=rs_lock("lastposter")%>"><%=rs_lock("lastposter")%></a></font>
				<%If sortlast Is Nothing Or sortlast.EOF Then%>
				<a href=detail.asp?TitleID=<%=rs_lock("TitleID")%>&tid=<%=tid%>&postname=<%=rs_lock("Name")%>>
<img src="images/lastpost2.gif" border="0"></a></td>
				<%Else%>
				<a href=detail.asp?Page=1000&TitleID=<%=rs_lock("TitleID")%>&tid=<%=tid%>&postname=<%=rs_lock("Name")%>#<%=sortlast("DetailID")%>><img src="images/lastpost2.gif"  border="0"></a></td>
				<%End IF%>
<td width="28" height="18" bgcolor="#E3E3E3" bordercolor="#000000" align="center"><a href="favorite.asp?id=<%=rs_lock("TitleID")%>"><img src="images/eye.gif"border=0></a></td>
</tr>
<%rs_lock.movenext
wend
end if
end if '假使第一頁的end if
'置頂結束****************************************************************************%>
<tr>
<%If rs.EOF Then
	response.write "<center><FONT size=-1>Sorry!!目前尚無留言...</FONT></center>"
Else
  rs.AbsolutePage = Page ' 將資料錄移至 PAGE 頁
FOR SH=1 to rs.PageSize      ' 顯示資料%> 
<td width="30" height="18" bgcolor="#E3E3E3" bordercolor="#000000" align="center">
<%if rs("reply_lock")=1 then
    response.write "<img src=""images/lock.gif"">"
elseif rs("poll")=1 then
    response.write "<img src=""images/f_poll.gif"">"
elseif rs("count")>50 then
    response.write "<img src=""images/hot.gif"">"
elseIF Now - rs("LastNewsDate") < 1 then 
 	response.write "<img src=""images/icon_folder_new_3.gif"">"
else
    response.write "<img src=""images/icon_folder_3.gif"">"
end if%>
</td>
<td width="50%" height="18" bgcolor="#000000" bordercolor="#000000" background="images/postbit_lightbg.gif"><p align="left">
<img src="images/post/<%=rs("R1")%>.gif" width=13 height=13 border=0>
<%if rs("best")=1 then%>
<img src="images/best.gif">
<%end if%>
<a href=Detail.asp?TitleID=<%=rs("TitleID")%>&tid=<%=Request("tid")%>&postname=<%=rs("Name")%>&noc=<%=str_rnd%>><font color="black"size=-1>
<font color=black><%=rs("subc")%></font><%=rs("Subject")%><font color=#CC3300>
				<%if rs11 Is Nothing Or rs11.EOF Then%>
				(<%=Datevalue(DateAdd("h",rstitle("servertimezone"),rs("CreateDate")))%>)
				<%else%>
				(<%=Datevalue(DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs("CreateDate")))%>)
				<%end if%></font></font></a><%if rs("poll")=1 then%><img src="images/sel.jpg"><%end if%>
<%IF (Now - rs("LastNewsDate"))<0.5 then
      response.write "<img src=""images/new.gif"">"		 
end if
mxPages = rs("replycount")/page_num
if mxPages<>CInt(mxPages) then
mxpages=Int(mxPages)+1
end if

if (mxPages>1 and mxPages<=12) then 

Response.Write("<br>　　<font size=-1>第")
for counter = 1 to mxPages
ref = "&nbsp;<a href='Detail.asp?"
ref = ref & "tid=" & Request("tid")
ref = ref & "&TitleID=" & rs("TitleID")
ref = ref & "&postname=" & rs("Name")
ref = ref & "&Page=" & counter
ref = ref & "'><font color=ff0000><b>" & counter & "</font></b></a>"
response.write ref
next
Response.Write("&nbsp;頁。</font>")
elseif mxPages>12 then
Response.Write("<br>　　<font size=-1>第")
for counter = 1 to 12
ref = "&nbsp;<a href='Detail.asp?"
ref = ref & "tid=" & Request("tid")
ref = ref & "&TitleID=" & rs("TitleID")
ref = ref & "&postname=" & rs("Name")
ref = ref & "&Page=" & counter
ref = ref & "'><font color=ff0000><b>" & counter & "</font></b></a>"
response.write ref
next
response.write "......."
las = "&nbsp;<a href='Detail.asp?"
las = las & "tid=" & Request("tid")
las = las & "&TitleID=" & rs("TitleID")
lasf = las & "&postname=" & rs("Name")
las = las & "&Page=" & mxPages
las = las & "'><font color=ff0000><b>" & mxPages & "</font></b></a>"
response.write las
Response.Write("&nbsp;頁。</font>")
end if '特殊選項 
response.write "<font size=-1>"
Set passed2rs = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From tid where tid="&request("tid")
passed2rs.Open SQL, conn,1,1
if passed2rs("tidadmin")<>"" then
tidadmin=split(passed2rs("tidadmin"),"|")
 
for i = 1 to ubound(tidadmin)-1 '檢驗版主
 if tidadmin(i)=UserID and request.cookies("admin_pass")<>bbsid and rs("Name")<>UserID then
  Response.Write " <br>  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a href=delete.asp?TitleID="& rs("TitleID")&"&tid="&tid&">" & "刪除</a>"
  Response.Write "<a href=edit.asp?TitleID="& rs("TitleID")&"&tid="&request("tid")&"&postname="&rs("Name")&"&url=index.asp"&">" & "|&nbsp;編輯</a>"
 end if
next
end if

If request.cookies("admin_pass")=bbsid Or rs("Name")= UserID Then
Response.Write " <br>  &nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp<a href=delete.asp?TitleID="& rs("TitleID")&"&tid="&tid&">" & "刪除</a>"
Response.Write "<a href=edit.asp?TitleID="& rs("TitleID")&"&tid="&request("tid")&"&postname="&rs("Name")&"&url=index.asp"&">" & "|&nbsp;編輯</a>"
End If%></font></p></td>
<td width="130" height="18" bgcolor="#FFFFFF" align="center" bordercolor="#000000"> 
<font color="#000080" size="2"><a href="edituser.asp?UserID=<%=rs("Name")%>"><%=rs("Name")%></a></font></td>
<td width="50" height="18" bgcolor="#F7F5EE" align="center" bordercolor="#000000" background="images/postbit_lightbg.gif"> 
<font size="2">　<%=rs("count")%></font></td>
<td width="50" height="18" bgcolor="#FFFFFF" align="center" bordercolor="#000000"> 
<font size="2">　<%=rs("replycount")%></font></td>
<td width="160" height="18" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif"> <p align="right"><font color="black" size="2">
				<%if rs11 Is Nothing Or rs11.EOF Then%>  
				<%=DateAdd("h",rstitle("servertimezone"),rs("LastNewsDate"))%>
				<%else%>
				<%=DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs("LastNewsDate"))%>
				<%end if%>
<%sql="select * from Details where TitleID="&rs("TitleID")&" order by DetailID Desc"
	set sortlast=conn.execute(sql)%>				
				<br>by&nbsp;<a href="edituser.asp?UserID=<%=rs("lastposter")%>"><%=rs("lastposter")%></a></font>
				<%If sortlast Is Nothing Or sortlast.EOF Then%>
				<a href=detail.asp?TitleID=<%=rs("TitleID")%>&tid=<%=tid%>&postname=<%=rs("Name")%>>
<img src="images/lastpost2.gif" border="0"></a></td>
				<%Else%>
				<a href=detail.asp?Page=1000&TitleID=<%=rs("TitleID")%>&tid=<%=tid%>&postname=<%=rs("Name")%>#<%=sortlast("DetailID")%>><img src="images/lastpost2.gif" border="0"></a></td>
				<%End IF%>
	<td width="28" height="18" bgcolor="#E3E3E3" bordercolor="#000000" align="center"><a href="favorite.asp?id=<%=rs("TitleID")%>"><img src="images/eye.gif"border=0></a></td>
</tr>
<%rs.MoveNext
IF rs.EOF THEN EXIT FOR
Next
End If%>
</table></div>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="98%" height="26">
<tr>
	<td width="100%" height="15" background="images/Button.gif" bordercolor="#000000">
		&nbsp;&nbsp; <img border="0" src="images/xp.gif">
	<span id="forum"><font color="#333333" size="2"><b> 
		<%SQL= "Select * From config_setting"
		Set bbsnamers=conn.Execute(SQL) %><%=bbsnamers("bbsname")%>論壇 論壇圖例</b></font>
		<font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border="0" src="images/icon_folder_new_3.gif"> 近期的文章&nbsp;&nbsp;&nbsp;
		<img border="0" src="images/icon_folder_3.gif"> 舊的文章&nbsp;&nbsp;&nbsp;
		<img border="0" src="images/lock.gif"> 已鎖定的文章&nbsp;&nbsp;&nbsp;
		<img border="0" src="images/hot.gif"> 熱門的文章</font>&nbsp;
		<img border="0" src="images/f_poll.gif" width="18" height="12"><font size="2">投票</font></span></td></tr>
</table>
  <!-- 快速跳頁功能 -->
<%If rs.PageCount=0 Then
response.write "<p align=center><font size=-1>頁次1/1</font></p>"
Else
response.write "<form action="&myself&" method=GET>"
response.write "<font size=-1><p align=center>共<font color=red><b>"&rs.PageCount&"</b></font>頁&nbsp;&nbsp;&nbsp;&nbsp;"
response.write "<font size=-1>目前頁次<font color=red><b>"&Page&"</b></font>&nbsp;&nbsp;&nbsp;&nbsp;"
If Page<>1 Then
response.write "<a href="&myself&"?Page=1&tid="&tid&">[1]</a>"
response.write "........."
response.write "<a href="&myself&"?Page="&(Page-1)&"&tid="&tid&">上一頁</a>"
End If
If Page<>rs.PageCount Then
response.write "<a href="&myself&"?Page="&(Page+1)&"&tid="&tid&">下一頁</a>"
response.write "........."
response.write "<a href="&myself&"?Page="&rs.PageCount-1&"&tid="&tid&">["&rs.PageCount-1&"]</a>"
response.write "<a href="&myself&"?Page="&rs.PageCount&"&tid="&tid&">["&rs.PageCount&"]</a>"
End If
response.write "&nbsp;&nbsp;&nbsp;&nbsp;輸入頁次:<INPUT TYPE=TEXT NAME=Page SIZE=2><INPUT TYPE=hidden NAME=tid VALUE="&tid&"></form></p>"
End If%><p align="right">
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option selected>跳轉論壇至...</option><%
Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
while not rsjump.eof
%><option value=forum.asp?tid=<%=rsjump("tid")%>&Page=1><%=rsjump("tidname")%></option><%
rsjump.movenext
wend
rsjump.close%></select><br></p>
<div align="CENTER">
<div align="center">
<center><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="22">
<tr>
<td width="524" height="22" bordercolor="#FFFFFF" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif"><font size=-1><b><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
%>論壇</b></font></td>
<td width="5%" height="22" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif">
<a href="#top"><img border="0" src="images/up.gif" width="20" height="20" alt="回頂端"></a></td>
</tr></table></center></div>
<%'*******************************************************
else
response.redirect "forum_check.asp?tid="&tid
end if
'********************************************************%>
<!--#include file="foot.asp"-->
</div></font>
</BODY></HTML></textarea>