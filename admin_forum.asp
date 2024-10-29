<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0"><!--#include file="header.asp"-->
<!--#include file="admin/broad.asp"--><center><font size=-1>
<%if request("mode")="top_lock" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
一級置頂<INPUT TYPE="radio" NAME="level" value=1 checked>
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
二級置頂<INPUT TYPE="radio" NAME="level" value=2>
三級置頂<INPUT TYPE="radio" NAME="level" value=3>
<INPUT TYPE="hidden" NAME="tid" value=<%=request("tid")%>>
<INPUT TYPE="submit" value="確定">
</FORM> 
<%elseif request("mode")="cancel_lock" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
取消置頂<INPUT TYPE="radio" NAME="level" value=0 checked>
一級置頂<INPUT TYPE="radio" NAME="level" value=1>
二級置頂<INPUT TYPE="radio" NAME="level" value=2>
三級置頂<INPUT TYPE="radio" NAME="level" value=3>
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="hidden" NAME="tid" value=<%=request("tid")%>>
<INPUT TYPE="submit" value="確定">
</form>
<%elseif request("mode")="reply_lock" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
取消鎖定<INPUT TYPE="radio" NAME="lock" value=0 checked>
鎖定<INPUT TYPE="radio" NAME="lock" value=1>
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="hidden" NAME="tid" value=<%=request("tid")%>>
<INPUT TYPE="submit" value="確定">
</form>
<%elseif request("mode")="tit_move" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
移轉至:<select name="tid">
<%Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
while not rsjump.eof%>
<option value=<%=rsjump("tid")%>><%=rsjump("tidname")%></option>
<%rsjump.movenext
wend%>
</select>
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="submit" value="確定">
</form>
<%elseif request("mode")="reply_count" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" name="tid" value=<%=request("tid")%>>
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="submit" value="確定重新統計該主題回覆數目">
</FORM>
<%elseif request("mode")="tid_count" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" name="tid" value=<%=request("tid")%>>
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="submit" value="確定重新統計此討論區的文章、回覆數目">
</FORM>
<%elseif request("mode")="best" then%>
<FORM METHOD=POST ACTION="admin_forum_action.asp">
取消精華文章<INPUT TYPE="radio" NAME="level" value=0 checked>
加入精華文章<INPUT TYPE="radio" NAME="level" value=1>

<INPUT TYPE="hidden" name="mode" value="<%=request("mode")%>">
<INPUT TYPE="hidden" NAME="TitleID" value="<%=request("TitleID")%>">
<INPUT TYPE="hidden" NAME="tid" value=<%=request("tid")%>>
<INPUT TYPE="submit" value="確定">
</form>
<%end if%> </font><!--#include file="foot.asp"-->
</BODY></HTML>