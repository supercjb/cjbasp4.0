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
<BODY leftmargin="0" topmargin="0">
<!-- 登錄 --><!--#include file="header.vs"--><!--#include file="admin/broad.vs"-->
<%tid=request("tid")
if not request("forum_pass")=empty then
sql="select * from tid where tid="&tid
set rs_forum=conn.execute(sql)
 if not request.cookies("UserPassed")="yes" then
  response.redirect "index.asp?msg=請先登入!"
 else
   if rs_forum("have_pass")=1 and rs_forum("password")=request("forum_pass") then
    forum_cookies=bbsid+Cstr(tid)+request.cookies("UserID")
    forum_pswds=bbsid+Cstr(tid)+request.cookies("UserID")+Cstr(tid)
    ExpireDate=DateAdd("d",30,Date)
    response.cookies(forum_cookies).Expires=FormatDateTime(ExpireDate)
    response.cookies(forum_cookies)="yes"
    response.cookies(forum_pswds).Expires=FormatDateTime(ExpireDate)   
    response.cookies(forum_pswds)=request("forum_pass")
    response.redirect "forum.asp?Page=1&tid="&tid
   else
    response.redirect "forum_check.asp?msg=密碼錯誤!&tid="&tid
   end if
  end if
end if%>
<center>
<FORM METHOD=GET>
<INPUT TYPE="hidden" name="tid" value="<%=tid%>">
　<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="322" height="42" id="AutoNumber1">
    <tr>
      <td width="322" height="8" background="images/Button.gif">
      <p align="center"><b><font color="#800000">論壇驗證</font></b></td>
    </tr>
    <tr>
      <td width="322" height="35">
<font size="2">請輸入安全密碼：</font><INPUT TYPE="password" NAME="forum_pass" size="15">
     <center> <p style="margin-top: 0; margin-bottom: 0"><font color=red><b><%=request("msg")%></b></font></p>
      <p><font size="2" color="#FF0000">說明：此討論類別需要密碼驗證方可進入</font></p>
      <p align="center">
<INPUT TYPE="submit" value="確定"></td>
    </tr>
  </table>
  </center>
</div>
<p>
　</p>
</FORM> <!--#include file="foot.vs"-->
</BODY></HTML>