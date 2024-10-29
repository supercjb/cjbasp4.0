<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"-->
<% response.cookies("tid")="搜尋結果"%><br><!--#include file="choice.asp"--><br>
<%SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
set rs11=conn.execute(SQL)
if request.cookies("UserPassed")<>"yes" then
  response.redirect "index.asp?msg=請先登錄論壇!"
end if
dim strSQL
dim strSelect, strWhere, strOrder

'假設 pkey 是 table 的 primary key
strSelect = "select * from Titles"
strSelect2 = "select * from Titles"
strSelect = strSelect & " where "
strOrder = " order by TitleID desc"
strOrder2 = " order by LastNewsDate desc"
strOrder3 = " order by count desc"
strOrder4 = " order by replycount desc"
strWhere = ""
if request("key")="ON" then'若關鍵字搜尋

if request("tit")<>"" then
if request("range")=0 then
 strWhere = strWhere & " Subject like '%" &request("tit") & "%'"
elseif request("range")=1 then
 strWhere = strWhere & " Name like '%" &request("tit") & "%'"
elseif request("range")=2 then
 strWhere = strWhere & " Words like '%" &request("tit") & "%'"
else
end if
end if

else'若不是關鍵字搜尋

'搜尋條件有四個，依序判斷是否有輸入
if request("tit")<>"" then
if request("range")=0 then
 strWhere = strWhere & " Subject = '" &request("tit") & "'"
elseif request("range")=1 then
 strWhere = strWhere & " Name = '" &request("tit") & "'"
elseif request("range")=2 then
 strWhere = strWhere & " Words like '%" &request("tit") & "%'"
else
end if
end if
end if'關鍵字endif

if request("order") = 0 then
strSQL = strSelect & strWhere & strOrder2
elseif request("order") = 1 then
strSQL = strSelect & strWhere & strOrder
elseif request("order") = 2 then
strSQL = strSelect & strWhere & strOrder3
elseif request("order") = 3 then
strSQL = strSelect & strWhere & strOrder4
else
end if

Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open strSQL,conn,3,3

     Page = cint(Request("Page"))
	 myself=request.serverVariables("PATH_INFO")
	
	  rs.PageSize =10      ' 設定每頁顯示 20 筆
     If Page<1 or Page=empty Then 
		Page=1
     ELSEIF Page>rs.PageCount Then 
		Page=rs.PageCount
     END IF

If rs.Eof Then 
    	RESPONSE.WRITE "<FONT COLOR=maroon size=-1><br><center>抱歉,找不到您要的資料</center></font>" 
Else
  rs.AbsolutePage = Page ' 將資料錄移至 PAGE 頁%><center><br>
<table border=1 height="17" bordercolor="#C0C0C0" style="border-collapse: collapse" cellpadding="0" cellspacing="0" width="96%">
<tr>
  <td height="14" background="images/Button.gif" width="227" colspan="3" align="center">
  <font size="2" color="#000080">主題</font></td>
  <td height="14" background="images/Button.gif" width="130" align="center">
  <font size="2" color="#000080">作者</font></td>
  <td height="14" background="images/Button.gif" width="101" align="center">
  <font size="2" color="#000080">論壇</font></td>
  <td height="14" background="images/Button.gif" width="33" align="center">
  <font size="2" color="#000080">回覆</font></td>
  <td height="14" background="images/Button.gif" width="31" align="center">
  <font size="2" color="#000080">瀏覽</font></td>
  <td height="14" background="images/Button.gif" width="130" align="center">
  <font size="2" color="#000080">最後更新</font></td>
  <td height="5" background="images/Button.gif" width="30" align="center">
  <font size="2" color="#000080">關注</font></td></tr>
<%FOR SH=1 to rs.PageSize      ' 顯示資料%>
 <tr>
  <td height="14" width="8" background="images/postbit_lightbg.gif"><%
if rs("reply_lock")=1 then
    response.write "<img src=""images/lock.gif"">"
elseif rs("count")>50 then
    response.write "<img src=""images/hot.gif"">"
elseIF Now - rs("LastNewsDate") < 1 then 
 	response.write "<img src=""images/icon_folder_new_3.gif"">"
else
    response.write "<img src=""images/icon_folder_3.gif"">"
end if%></td>
  <td height="14" width="15" align="center" background="images/postbit_lightbg.gif"><img src="images/post/<%=rs("R1")%>.gif" width=13 height=13 border=0></td>
  <td height="14" width="214" align="center" bgcolor="#EBEBEB" background="images/postbit_lightbg.gif"><font size=-1>
  <%if rs("replycount")=0 then%>
  <a href="Detail.asp?TitleID=<%=rs("TitleID")%>&tid=<%=rs("tid")%>&postname=<%=rs("Name")%>"><%=rs("Subject")%></a></td>
  <%else%>
<%sql="select * from Details where TitleID="&rs("TitleID")&" order by DetailID Desc"
	set sortlast=conn.execute(sql)%>				
  <a href="Detail.asp?page=100&TitleID=<%=rs("TitleID")%>&tid=<%=rs("tid")%>&postname=<%=rs("Name")%>#<%=sortlast("DetailID")%>"><%=rs("Subject")%></a></td>
  <%end if%>
  <td height="14" width="130" align="left" background="images/postbit_lightbg.gif"><font size=-1>
  			<%if rs11 Is Nothing Or rs11.EOF Then%>  
			<%=DateAdd("h",rstitle("servertimezone"),rs("createdate"))%>
			<%else%>
			<%=DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs("createdate"))%>
			<%end if%>
  <p align="right" style="margin-top: 0; margin-bottom: 0"><a href="edituser.asp?UserID=<%=rs("Name")%>"><%=rs("Name")%></a></p></font></td>
  <td height="14" width="101" align="center" bgcolor="#EBEBEB" background="images/postbit_lightbg.gif"><font size=-1>
  <%sql="select * from tid where tid="&rs("tid")
  set rs2=conn.execute(sql)%>
  <a href="forum.asp?tid=<%=rs("tid")%>"><%=rs2("tidname")%></a></font></td>
  <td height="14" width="33" align="center" background="images/postbit_lightbg.gif"><font size=-1><%=rs("replycount")%></font></td>
  <td height="14" width="31" align="center" bgcolor="#EBEBEB" background="images/postbit_lightbg.gif"><font size=-1><%=rs("count")%></font></td>
  <td height="14" width="130" align="left" background="images/postbit_lightbg.gif"><font size=-1>
  			<%if rs11 Is Nothing Or rs11.EOF Then%>  
			<%=DateAdd("h",rstitle("servertimezone"),rs("LastNewsDate"))%>
			<%else%>
			<%=DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs("LastNewsDate"))%>
			<%end if%>
  <p align="right" style="margin-top: 0; margin-bottom: 0"><a href="edituser.asp?UserID=<%=rs("lastposter")%>"><%=rs("lastposter")%></a></p></font></td>
  <td height="5" width="30" align="center" background="images/postbit_lightbg.gif"><font size=-1><a href="favorite.asp?id=<%=rs("TitleID")%>"><img src="images/eye.gif"border=0></a></font></td>
  </tr>
  <%rs.MoveNext
IF rs.EOF THEN EXIT FOR
Next%></table>
  <%END IF%>
  <%If rs.PageCount=0 Then%><p>
 <p align="center"><font size="2">頁次1/1</p>
   </font><%Else%> <font size="2">
   <form action=<%=myself%> method=GET>
   <p align="center">
共有<b><font color=red><%=rs.PageCount%></font></b>頁&nbsp;&nbsp;&nbsp;&nbsp;
目前頁次<b>
   <font color=red><%=Page%></font></b>
 &nbsp;&nbsp;<%If Page<>1 Then%>
<a href="<%=myself%>?Page=1&key=<%=request("key")%>&sort=<%=request("sort")%>&name=<%=request("name")%>&tit=<%=request("tit")%>&replycount=<%=request("replycount")%>&timediff=<%=request("timediff")%>&order=<%=request("order")%>">[1]</a>..........
<a href="<%=myself%>?Page=<%=(Page-1)%>&key=<%=request("key")%>&sort=<%=request("sort")%>&name=<%=request("name")%>&tit=<%=request("tit")%>&replycount=<%=request("replycount")%>&timediff=<%=request("timediff")%>&order=<%=request("order")%>">上一頁</a>
<%End If%>
<%If Page<>rs.PageCount Then%>
<a href="<%=myself%>?Page=<%=(Page+1)%>&key=<%=request("key")%>&sort=<%=request("sort")%>&name=<%=request("name")%>&tit=<%=request("tit")%>&replycount=<%=request("replycount")%>&timediff=<%=request("timediff")%>&order=<%=request("order")%>">下一頁</a>
..........
<a href="<%=myself%>?Page=<%=rs.PageCount-1%>&key=<%=request("key")%>&sort=<%=request("sort")%>&name=<%=request("name")%>&tit=<%=request("tit")%>&replycount=<%=request("replycount")%>&timediff=<%=request("timediff")%>&order=<%=request("order")%>">[<%=rs.PageCount-1%>]</a>
<a href="<%=myself%>?Page=<%=rs.PageCount%>&key=<%=request("key")%>&sort=<%=request("sort")%>&name=<%=request("name")%>&tit=<%=request("tit")%>&replycount=<%=request("replycount")%>&timediff=<%=request("timediff")%>&order=<%=request("order")%>">[<%=rs.PageCount%>]</a>
<%End If%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

輸入頁次：<INPUT TYPE=TEXT NAME=Page SIZE=2>  
<INPUT TYPE=hidden NAME=key VALUE="<%=request("key")%>">
<INPUT TYPE=hidden NAME=sort VALUE="<%=request("sort")%>">
<INPUT TYPE=hidden NAME=name VALUE="<%=request("name")%>">
<INPUT TYPE=hidden NAME=tit VALUE="<%=request("tit")%>">
<INPUT TYPE=hidden NAME=replycount VALUE="<%=request("replycount")%>">
<INPUT TYPE=hidden NAME=timediff VALUE="<%=request("timediff")%>">
<INPUT TYPE=hidden NAME=order VALUE="<%=request("order")%>">
</form></p>
<%End If%>
<p align="right">
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option selected>跳轉論壇至...</option><%
Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
while not rsjump.eof
%><option value=forum.asp?tid=<%=rsjump("tid")%>&Page=1><%=rsjump("tidname")%></option><%
rsjump.movenext
wend
rsjump.close%></select></p></font>
    <!--#include file="foot.asp"-->
</BODY></HTML>