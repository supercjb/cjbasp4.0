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
   SQL="select * from member_sort where id="&request("id")
   rs3.Open SQL, conn,1,3
if request("sure")="yes" then
   rs3("levelname")=request("levelname")
   rs3("level")=request("le")
   rs3("upload_size")=request("size")
   rs3("pm_size")=request("pm_size")
   rs3("fav_size")=request("fav_size")
    rs3("pic")=request("pic")
   rs3.update
   response.write "<font size=-1><b>儲存成功!</b></font><br>"
end if%><font size=-1>
頭銜：<%=rs3("levelname")%><br>
等級代號：<%=rs3("level")%><br>
上傳檔案大小：<%=rs3("upload_size")%>K<br>
簡訊pm大小：<%=rs3("pm_size")%>封<br>
關注話題大小：<%=rs3("fav_size")%>個<br>
代表圖案：<img src="<%=rs3("pic")%>"><br>
<form method="GET">
<input type="hidden" name="sure" value="yes" size="20">
更改頭銜：<input type="text" name="levelname" size="20" value="<%=rs3("levelname")%>"><br>
更改代號：<input type="text" name="le" size="5" value="<%=rs3("level")%>"><br>
更改上傳檔案大小：<input type="size" name="size" size="5" value="<%=rs3("upload_size")%>"><br>
更改簡訊pm大小：<input type="size" name="pm_size" size="5" value="<%=rs3("pm_size")%>"><br>
更改關注話題大小：<input type="size" name="fav_size" size="5" value="<%=rs3("fav_size")%>">
<input type="hidden" name="id" value=<%=request("id")%> size="20"><br>
代表圖案位置：<input type="size" name="pic" size="55" value="<%=rs3("pic")%>"><br>
<input type="submit" value="確定" name="B1"><input type="reset" value="重新設定" name="B2">
</form>
</font><!--#include file="../foot.asp"-->
<%end if%>
</body> 
</html>