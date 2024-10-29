<html><!--#include file="conn.asp"-->
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""><link rel="stylesheet" type="text/css" href=../style1.css>
</HEAD>
<BODY>
<%
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else%>
<FORM METHOD=GET>
<font size="2">帳號：</font><INPUT TYPE="text" NAME="user" size="20">
<INPUT TYPE="submit" value="查詢使用者">
</FORM>
<%
if request("act")="delete" then
  sql="select * from id where ID="&request("id")
 Set rs2= Server.CreateObject("ADODB.Recordset")
  rs2.Open sql, conn,1,3
  rs2.delete
  response.write "刪除成功!!<br>"
end if
a=0
if not request("user")=empty then
sql="Select * from id Where user like '%" & request("user") & "%'" 
set rs=conn.execute(sql)
 if not rs.eof then
  while not rs.eof
    response.write "<table border=0><tr><td width=250 height=20 background=images/title.gif>"
    response.write rs("user")
     response.write "<a href=admin_user_edit.asp?id="&rs("ID")&">&nbsp;修改</a>"
	  response.write "&nbsp;|&nbsp;"%>
	  <a href="admin_user.asp?act=delete&id=<%=rs("ID")%>&user=<%=request("user")%>">刪除</a>
	 <%response.write "</td></tr></table>"
    rs.movenext
	a=a+1
  wend
response.write "<font size=-1>※共找到<b>"&a&"</b>筆資料</font>"
 else
 response.write "抱歉,無符合結果"
 end if
end if
end if%>
</BODY>
</HTML>