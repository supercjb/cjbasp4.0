<%@ LANGUAGE="VBSCRIPT" CODEPAGE="950" %>
<!-- #include file="conn.asp" -->
<html>
<head><META HTTP-EQUIV="Content-Type" content="text/html; charset=big5"> </head>
<body bgcolor="#F5F7F4" background="images/line.gif"><font size=-1>
<%if request.cookies("UserPassed")="yes" then
if not request.cookies("upload")="1" then%>
<form method="POST" action="upp.asp" enctype="multipart/form-data" ><p>
�����ɮ�:<input type="file" name="fruit" size="20">
<!--�h�ɤW�ǽƻs�G<input type="file" name="fruit" size="20">�L��--> 
<input type="submit" value="�W��" name="subbutt"> 
<input type="reset" value="���]" name="rebutt"><br>
<%sql="select * from id where user='"&request.cookies("UserID")&"'"
set rs_level=conn.execute(sql)
sql2="select * from member_sort where level="&rs_level("level")
set rs_level2=conn.execute(sql2)%>
<img src="images/1.gif">�z�����ŬO<font color=red><b><%=rs_level2("levelname")%></b></font>�i�H�W��<font color=red><b><%=rs_level2("upload_size")%></b></font>K
&nbsp;&nbsp;&nbsp;&nbsp;<font color=red><b><%=request("message")%></b></font>
</form> 
<%else%>
<font color=red><%=request("message")%></font>
<%end if
else
 response.write "�A�L�v��!�W��"
end if %>
</body></html>