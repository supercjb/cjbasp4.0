<%@ LANGUAGE="VBSCRIPT" CODEPAGE="950" %>
<!-- #include file="conn.asp" -->
<html>
<head><META HTTP-EQUIV="Content-Type" content="text/html; charset=big5"> </head>
<body bgcolor="#F5F7F4" background="images/line.gif"><font size=-1>
<%if request.cookies("UserPassed")="yes" then
if not request.cookies("upload")="1" then%>
<form method="POST" action="upp.asp" enctype="multipart/form-data" ><p>
附件檔案:<input type="file" name="fruit" size="20">
<!--多檔上傳複製：<input type="file" name="fruit" size="20">几次--> 
<input type="submit" value="上傳" name="subbutt"> 
<input type="reset" value="重設" name="rebutt"><br>
<%sql="select * from id where user='"&request.cookies("UserID")&"'"
set rs_level=conn.execute(sql)
sql2="select * from member_sort where level="&rs_level("level")
set rs_level2=conn.execute(sql2)%>
<img src="images/1.gif">您的等級是<font color=red><b><%=rs_level2("levelname")%></b></font>可以上傳<font color=red><b><%=rs_level2("upload_size")%></b></font>K
&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b><%=request("message")%></b></font>
</form> 
<%else%>
<font color=red><%=request("message")%></font>
<%end if
else
 response.write "你無權限!上傳"
end if %>
</body></html>