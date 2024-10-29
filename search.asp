<!--#include file="conn.asp"-->
<html><head>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<link rel="stylesheet" type="text/css" href=style1.css></head>
<body leftmargin="0" topmargin="0"><!--#include file="header.asp"--><!--#include file="admin/broad.asp"--><br>
<% response.cookies("tid")="搜尋文章"%>
<!--#include file="choice.asp"--><br><br><center>
<table width="50%" style="border-collapse: collapse" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" border="1">
<tr><td background="images/Button.gif"><font size="2">&nbsp;</font><font size="2" color="#000080">論壇搜尋</font></tr></td>
<FORM METHOD=POST ACTION="search_action.asp"><tr><td>
<font size="2">討論區&nbsp;&nbsp;&nbsp; ：</font><select name="sort">
<%Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
%><option Value="aa">所有的討論區</option><%
while not rsjump.eof%><option VALUE=<%=rsjump("tid")%>><%=rsjump("tidname")%></option>
<%rsjump.movenext
wend
rsjump.close%>
</select></tr></td>
	<tr>
		<td>
<font size="2">搜尋範圍：</font><select name="range" size="1">
<option selected value="0">標題</option>
<option value="1">作者</option>
<option value="2">內容</option>
</select></tr>
	</tr>
<tr><td><font size="2">搜尋內容：</font><INPUT TYPE="text" NAME="tit" size="35"></tr></td>
<tr><td>　</tr></td>
<tr>
	<td><font size="2">搜尋結果排序依</font><select size="1" name="order">
	<option value="0" selected>最後更新</option>
	<option value="1">發表日期</option>
		<option value="2">點閱數</option>
			<option value="3">回覆數</option>
	</select>
	
	</tr>
</tr>
<tr><td><font size="2">關鍵字搜尋：</font><INPUT TYPE="checkbox" NAME="key" value="ON"></tr></td>
<tr><td>
  <p align="center"><INPUT TYPE="submit" value="開始搜尋"></tr></td>
</FORM></table>
  <p align="right">
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option selected>跳轉論壇至...</option><%
Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
while not rsjump.eof
%><option value=forum.asp?tid=<%=rsjump("tid")%>&Page=1><%=rsjump("tidname")%></option><%
rsjump.movenext
wend
rsjump.close%></select></p>
<!--#include file="foot.asp"-->
</body></html>