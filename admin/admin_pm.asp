<HTML><!--#include file="conn.asp"-->
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>
<BODY>
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
sql="select * from member_sort"
set rs=conn.execute(sql)%>
<FORM METHOD=POST ACTION="admin_pm_action.asp">
傳送給<font style="font-size: 9pt">
<select name="to">
<option value=12>全部的帳號</option>
<%while not rs.eof%>
<option value=<%=rs("level")%>><%=rs("levelname")%></option>
<%rs.movenext
wend%>
</select>
</font><p>
<font style="font-size: 9pt">標題：<INPUT TYPE="text" NAME="tit" size="35"><br>
內容：<textarea rows="11" name="text" cols="50"></textarea></font></p>
<p><font style="font-size: 9pt" color="#FF0000">請確定後才傳出,以免造成系統負擔</font><br>
<INPUT TYPE="submit"></p>
</FORM>
<%end if %>
</BODY>
</HTML>