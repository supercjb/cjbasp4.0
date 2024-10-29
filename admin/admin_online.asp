<!--#include file="conn.asp"--><HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<meta http-equiv="refresh" content="30; url=admin_online.asp">  
<link rel="stylesheet" type="text/css" href=../style1.css>
</HEAD>

<BODY>
<B><font color=red size=3>CJB-ASP論壇即時監控系統1.0</font></b>
<%sql="select * from online"
set rs=conn.execute(sql)
response.write "<hr><b>首頁</b><br>"
while not rs.eof
if rs("tid")="0" then
response.write "<font color=red size=-1>"&rs("name")&"</font>&nbsp;&nbsp;"
end if
rs.movenext
wend
sql="select * from tid"
set rs1=conn.execute(sql)
while not rs1.eof
response.write "<hr><b>"&rs1("tidname")&"</b><br>"
sql="select * from online where tid='"&rs1("tid")&"'"
set rs=conn.execute(sql)
while not rs.eof
if not rs.eof then
response.write "<font color=red size=-1>"&rs("name")&"</font>&nbsp;&nbsp;"
end if
rs.movenext
wend
rs1.movenext
wend
%>
</BODY>
</HTML>