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
if Len(Request.Cookies("TEST_COOKIE")) = 0 then'�ˬdCOOKIE
%><Script Language=Vbscript>
MsgBox "�z���s�������䴩cookie!�ж}�ҫ�~��n�J����!!����:IE>�u��>�ﶵ>���p�v>�էC",64,"���~�I�I"
location.href = "javascript:history.back()"
</Script><%
else'�̥~�h��cookie�ˬdif%>
<%	
Response.Buffer =True
If Request.Form("UserID")<>Empty and Request.Form("Password")<>Empty Then
   UserID=replace(trim(Request.Form("UserID")),"'","��")
   Password=replace(trim(Request.Form("Password")),"'","��")
   Set Rs = Server.CreateObject("ADODB.Recordset")
   SQL = "Select * From ID Where user='"&UserID&"' and password='"&Password&"'"
   Rs.Open SQL, conn,1,3
If Not Rs.EOF Then
'cookie_save_day_sub
   Session.TimeOut=30
	If Rs("level")=1 or Rs("level")=10 Then '���ҬO�_���޲z��
	   response.Cookies("admin_pass")=bbsid
		'Session("Passed")=True
    End If

    If Rs("level")=2 Then '���ҬO�_�����D
		response.Cookies("Passed2")="yes"
    End If

   Response.Cookies("UserID")=UserID
   Response.Cookies("Password")=Password
   Response.Cookies("UserPassed")="yes"'�|���q�L�{�Ҫ�cookies
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
MsgBox "��ƶ�g���~!",64,"���~�I�I"
location.href = "javascript:history.back()"
</Script>
<% End If 
   Else   
   Response.Redirect "index.asp" 
   End If
   end if'�̥~�h��cookieend if
   %>
</BODY></HTML>
