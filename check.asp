<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">

<META NAME="Description" CONTENT=""></HEAD>
<!--#include file="conn.asp"-->
<BODY>
<%
if Len(Request.Cookies("TEST_COOKIE")) = 0 then'檢查COOKIE
%><Script Language=Vbscript>
MsgBox "您的瀏覽器不支援cookie!請開啟後才能登入本站!!說明:IE>工具>選項>隱私權>調低",64,"錯誤！！"
location.href = "javascript:history.back()"
</Script><%
else'最外層的cookie檢查if%>
<%	
Response.Buffer =True
If Request.Form("UserID")<>Empty and Request.Form("Password")<>Empty Then
   UserID=replace(trim(Request.Form("UserID")),"'","‘")
   Password=replace(trim(Request.Form("Password")),"'","‘")
   Set Rs = Server.CreateObject("ADODB.Recordset")
   SQL = "Select * From ID Where user='"&UserID&"' and password='"&Password&"'"
   Rs.Open SQL, conn,1,3
If Not Rs.EOF Then
'cookie_save_day_sub
   Session.TimeOut=30
	If Rs("level")=1 or Rs("level")=10 Then '驗證是否為管理員
	   response.Cookies("admin_pass")=bbsid
		'Session("Passed")=True
    End If

    If Rs("level")=2 Then '驗證是否為版主
		response.Cookies("Passed2")="yes"
    End If

   Response.Cookies("UserID")=UserID
   Response.Cookies("Password")=Password
   Response.Cookies("UserPassed")="yes"'會員通過認證的cookies
   response.cookies("bbsid")=bbsid
   response.Cookies("level")=Rs("level")

if request("cookie_save_day")="365" then
   ExpireDate=DateAdd("d",365,Date)
   Response.Cookies("UserID").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Password").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("UserPassed").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Passed2").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("admin_pass").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("level").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("bbsid").Expires=FormatDateTime(ExpireDate)   
elseif request("cookie_save_day")="30" then
   ExpireDate=DateAdd("d",30,Date)
   Response.Cookies("UserID").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Password").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("UserPassed").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Passed2").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("admin_pass").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("level").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("bbsid").Expires=FormatDateTime(ExpireDate)
elseif request("cookie_save_day")="7" then
   ExpireDate=DateAdd("d",7,Date)
   Response.Cookies("UserID").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Password").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("UserPassed").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Passed2").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("admin_pass").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("level").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("bbsid").Expires=FormatDateTime(ExpireDate)
elseif request("cookie_save_day")="1" then
   ExpireDate=DateAdd("d",1,Date)
   Response.Cookies("UserID").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Password").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("UserPassed").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("Passed2").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("admin_pass").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("level").Expires=FormatDateTime(ExpireDate)
   Response.Cookies("bbsid").Expires=FormatDateTime(ExpireDate)
elseif request("cookie_save_day")="0" then
else
end if



   if request("act")="head_login" then
      Response.redirect request("url")
   else
      Response.Redirect "index.asp"
   end if%>
<%Else %>
<Script Language=Vbscript>
MsgBox "資料填寫錯誤!",64,"錯誤！！"
location.href = "javascript:history.back()"
</Script>
<% End If 
   Else   
   Response.Redirect "index.asp" 
   End If
   end if'最外層的cookieend if
   %>
</BODY></HTML>
