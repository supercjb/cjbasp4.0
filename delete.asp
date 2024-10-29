<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<BODY><!--#include file="conn.asp"-->
<%TitleID=Request("TitleID")
UserID=request.cookies("UserID")
tid=Request("tid")
Dim Fo
Set Fo = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")

Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Titles Where TitleID="& TitleID,conn,1,3
Set tidadmin = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From tid where tid="&tid
tidadmin.Open SQL, conn,1,1

master=split(tidadmin("tidadmin"),"|")
for i = 1 to ubound(master)-1 '檢驗版主
if master(i)=UserID and rs("Name")<>UserID and request.cookies("admin_pass")<>bbsid then

sql="select * from Details where TitleID="&request("TitleID")&""
Set rs123=conn.execute(sql)

while not rs123.eof 
if rs123("file")<>empty then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&rs123("file"))) then
		FileName = "upload/" & Trim(rs123("file"))'刪除回覆文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
end if
rs123.movenext
wend

Set cmd=Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection=rs.ActiveConnection
SQL = "DELETE From Details "
SQL = SQL & " Where TitleID=" & TitleID
cmd.CommandText = SQL
cmd.Execute
Set rs3=Server.CreateObject("ADODB.Recordset")
rs3.Open "Select * From tid Where tid="&request("tid"),conn,1,3
'更新刪除後的主題回覆數
rs3("counts")=rs3("counts")-1'主題數減1
rs3.update
rs3("replycounts")=rs3("replycounts")-rs("replycount")'回覆數減掉該主題的原回覆數
rs3.update
sql="select * from tid where tid="&tid '刪除版的最後更新
Set rs_delas=Server.CreateObject("ADODB.Recordset")
rs_delas.Open sql, conn,3,3
rs_delas("lastposter")=""
rs_delas("lasttitle")=0
rs_delas("lasttitle2")=""
rs_delas.update
'判斷檔案是否半刪了
if rs("file")<>empty then	
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&rs("file"))) then
		FileName = "upload/" & Trim(rs("file"))'刪除文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
		
end if
rs.Delete'更新數目完成後 刪除該主題
Response.Redirect "forum.asp?msg=主題已刪除!&tid="&request("tid")
end if
next

If request.cookies("admin_pass")=bbsid Or rs("Name")= UserID Then'判斷是不是管理員或原作者

sql="select * from Details where TitleID="&request("TitleID")&""
Set rs123=conn.execute(sql)

while not rs123.eof 
if rs123("file")<>empty then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&rs123("file"))) then
		FileName = "upload/" & Trim(rs123("file"))'刪除回覆文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
end if
rs123.movenext
wend

Set cmd=Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection=rs.ActiveConnection
SQL = "DELETE From Details "
SQL = SQL & " Where TitleID=" & TitleID
cmd.CommandText = SQL
cmd.Execute
Set rs3=Server.CreateObject("ADODB.Recordset")
rs3.Open "Select * From tid Where tid="&request("tid"),conn,1,3
'更新刪除後的主題回覆數
rs3("counts")=rs3("counts")-1'主題數減1
rs3.update
rs3("replycounts")=rs3("replycounts")-rs("replycount")'回覆數減掉該主題的原回覆數
rs3.update

sql="select * from tid where tid="&tid'刪除版的最後更新
Set rs_delas=Server.CreateObject("ADODB.Recordset")
rs_delas.Open sql, conn,3,3
rs_delas("lastposter")=""
rs_delas("lasttitle")=0
rs_delas("lasttitle2")=""
rs_delas.update

if rs("file")<>empty then
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&rs("file"))) then
		FileName = "upload/" & Trim(rs("file"))'刪除文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
end if       

if rs("poll")=1 then'有投票則 刪除投票'1
	Set del_poll= Server.CreateObject("ADODB.Recordset")
   SQL="SELECT * FROM poll where titleid="&request("TitleID")
   del_poll.Open SQL, conn,1,3
   while not del_poll.EOF
   del_poll.delete
   del_poll.movenext
   wend
  	Set del_poller= Server.CreateObject("ADODB.Recordset")
   SQL="SELECT * FROM poller where titleid="&request("TitleID")
   del_poller.Open SQL, conn,1,3
   while not del_poller.EOF
   del_poller.delete
   del_poller.movenext
   wend
end if'1
rs.Delete'更新數目完成後 刪除該主題

Response.Redirect "forum.asp?msg=主題已刪除!&tid="&request("tid")
Else
Response.Redirect "forum.asp?msg=你無權限!&tid="&request("tid")
	'非管理者或作者,則無權限
End If
  tidadmin.close%>
</BODY></HTML>