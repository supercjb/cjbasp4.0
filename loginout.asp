<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<BODY>
<%UserID=request.cookies("UserID")
  SQL="SELECT * FROM online where name='"&UserID&"'"
  
  Set rs= Server.CreateObject("ADODB.Recordset") 
  rs.open SQL,conn,1,3
  if not rs.EOF then
  rs.delete
  rs.close
  end if
  set rs=nothing
	Session("Passed")= false
	session.abandon
	 ExpireDate=DateAdd("d",-100,Date)
'****************************®ø°£ªO¶ô±K½Xªºcookies**************************
sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
tid=request("tid")
sql="select * from tid "
set rs_forum=conn.execute(sql)
while not rs_forum.eof
 
    forum_cookies=bbsid+Cstr(tid)+request.cookies("UserID")
	if not request.cookies(forum_cookies)=empty then
    Response.Cookies(forum_cookies).Expires=FormatDateTime(ExpireDate)
	end if
	rs_forum.movenext 
wend
'*****************************************************************************

   Response.Cookies("UserID").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Password").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("UserPassed").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Passed2").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("admin_pass").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("level").Expires=FormatDateTime(ExpireDate)
    Response.Cookies("bbsid").Expires=FormatDateTime(ExpireDate)
	 'Response.Cookies("TEST_COOKIE").Expires=FormatDateTime(ExpireDate)

response.redirect "index.asp"%>
</BODY></HTML>