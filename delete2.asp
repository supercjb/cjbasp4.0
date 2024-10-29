<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<!--#include file="conn.asp"-->
<BODY>
<%UserID=request.cookies("UserID")
DetailID=Request("DetailID")
TitleID=Request("TitleID")
Dim Fo
Set Fo = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")


Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Details Where DetailID="& DetailID,conn,1,3
Set tidadmin = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From tid where tid="&request("tid")
tidadmin.Open SQL, conn,1,1
master=split(tidadmin("tidadmin"),"|")
		for i = 1 to ubound(master)-1 '檢驗版主
         if master(i)=UserID and rs("Name")<>UserID and request.cookies("admin_pass")<>bbsid then
			
'判斷檔案是否半刪了
if rs("file")<>empty then
		master123=split(rs("file"),",")
		for j = 0 to ubound(master123)
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&master123(j))) then
		FileName = "upload/" & Trim(master123(j))'刪除文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
		next	
end if
            rs.Delete
			Set cmd=Server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection=rs.ActiveConnection
			
			SQL = "Update Titles Set [replycount]=[replycount]-1"
			SQL = SQL & " Where TitleID=" & TitleID
			cmd.CommandText = SQL
			cmd.Execute
			rs.close
			rs.Open "Select * From tid Where tid="&request("tid"),conn,1,3
		    rs("replycounts")=rs("replycounts")-1
            rs.update
        	Response.Redirect "Detail.asp?TitleID="&TitleID
         end if
		next

If request.cookies("admin_pass")=bbsid Or rs("Name")= UserID Then '判斷是否為管理員或原作者
'判斷檔案是否半刪了
if rs("file")<>empty then
		master123=split(rs("file"),",")
		for i = 0 to ubound(master123)
		Set FSO = CreateObject("Scripting.FileSystemObject")
		if Fo.FileExists(Server.MapPath("upload/"&master123(i))) then
		FileName = "upload/" & Trim(master123(i))'刪除回覆文章附件
		FSO.DeleteFile Server.MapPath(Trim(FileName)),True
		end if
		next	
end if
	rs.Delete
	Set cmd=Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection=rs.ActiveConnection
	SQL = "Update Titles Set [replycount]=[replycount]-1"
	SQL = SQL & " Where TitleID=" & TitleID
	cmd.CommandText = SQL
	cmd.Execute
rs.close
rs.Open "Select * From tid Where tid="&request("tid"),conn,1,3
  rs("replycounts")=rs("replycounts")-1
  rs.update

	Response.Redirect "Detail.asp?TitleID="&TitleID&"&tid="&request("tid")
Else
	Response.Redirect "index.asp"
End If%>
</BODY></HTML>