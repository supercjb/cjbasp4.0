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
   SQL="select * from friend_web where id="&request("id")
   rs3.Open SQL, conn,1,3
if request("sure")="yes" then
   rs3("name")=request("name")
   rs3("url")=request("url")
   rs3("logo")=request("logo")
   rs3.update
Response.Redirect "admin_web.asp"
end if%><font size=-1>
聯盟名稱：<%=rs3("name")%><br>
聯盟網址：<%=rs3("url")%><br>
聯盟LOGO：<%=rs3("logo")%>
<form method="GET">
<input type="hidden" name="sure" value="yes" size="20">
更改名稱：<input type="text" name="name" size="20" value="<%=rs3("name")%>"><br>
更改網址：<input type="text" name="url" size="30" value="<%=rs3("url")%>"><br>
更改logo：<input type="text" name="logo" size="30" value="<%=rs3("logo")%>"><br>
<font color=red size=-1><b>ps.若沒有logo請在logo欄位填入0</b></font><br>
<input type="hidden" name="id" value=<%=request("id")%> size="20">
<input type="submit" value="確定" name="B1"><input type="reset" value="重新設定" name="B2"></p>
</form>
</font>
<!--#include file="../foot.vs"-->
<%end if %>
</body>  
</html>