<html><!--#include file="conn.asp"-->
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title><link rel="stylesheet" type="text/css" href=../style1.css>
</head>
<body>
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
   Set rs3= Server.CreateObject("ADODB.Recordset") 
   Set rs4= Server.CreateObject("ADODB.Recordset") 
   SQL="select * from sort where id="&request("id")
   rs3.Open SQL, conn,1,3
sql="select * from tid where sort='"&rs3("sort_name")&"'"
   rs4.Open sql, conn,1,3
if request("sure")="yes" then
   rs3("sort_name")=request("name")
   rs3("sort_order")=request("order")
   rs3.update
while not rs4.eof 
 rs4("sort")=request("name")
 rs4.update
 rs4.movenext
wend
Response.Redirect "admin_sort.asp"
end if%><font size=-1>
分類名稱：<%=rs3("sort_name")%><br>
分類順序：<%=rs3("sort_order")%>
<form method="GET">
<input type="hidden" name="sure" value="yes" size="20">
  更改名稱：<input type="text" name="name" size="20" value="<%=rs3("sort_name")%>"></p>
  <p>更改順序：<input type="text" name="order" size="5" value="<%=rs3("sort_order")%>"></p>
  <input type="hidden" name="id" value=<%=request("id")%> size="20">
  <p><input type="submit" value="確定" name="B1"><input type="reset" value="重新設定" name="B2"></p>
</form>
</font><!--#include file="../foot.asp"-->
<%end if%>
</body>  
</html>