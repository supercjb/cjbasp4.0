<!--#include file="conn.asp"-->
<!--#include file="safe.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE> <%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<BODY>
<%UserID=Request.form("Name2")
Subject=Request.form("Subject")
Words=Request.form("Words")
Color1=Request.form("Color1")
tid=Request.form("tid")
subc=Request.form("subc")
R1=Request.form("R1")
file=request.cookies("filename")
if request("save_date")<>"不設定" then
save_date=request("save_date")
end if 
reply_pm=request("reply_pm")
sql="select * from config_setting"
set rstitle=conn.execute(sql)
	Set rs2= Server.CreateObject("ADODB.Recordset") '檢查帳號是否被用
   SQL="SELECT * FROM id WHERE user='" & UserID & "'"
   rs2.Open SQL, conn,1,1
if  request.cookies("UserPassed")<>"yes"  Then 'Or rs2.EOF
	Response.redirect "index.asp?msg=請登入!"%>
<%ElseIf Subject=Empty Or Words=Empty Or UserID =Empty Or Color1 =Empty Or tid =Empty Then
'Subject=Empty Or Words=Empty Or UserID =Empty Or Color1 =Empty Or tid =Empty Then
%>
        <Script Language=Vbscript>
		MsgBox "欄位留白!!!",64,"錯誤！！"
		location.href = "javascript:history.back()"
        </Script>
<%elseif len(Words)>rstitle("word_len") then%>
      <Script Language=Vbscript>
		MsgBox "字數超過!!!",64,"錯誤！！"
		location.href = "javascript:history.back()"
        </Script>
<%Else
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Titles",conn,1,3
	rs.AddNew
	rs("Name")=UserID
	rs("Subject")=Subject
	rs("Words")=Words
	rs("Color")=Color1
	rs("tid")=tid
	rs("lastposter")=UserID	
	rs("LastNewsDate")=now()
	rs("CreateDate")=now()
	rs("subc")=subc
	rs("R1")=R1
	rs("file")=file
	rs("save_date")=save_date
	rs("reply_pm")=reply_pm
	rs.Update
	Session("Subject")=""
	Session("Words")=""
	Set cmd=Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection=rs.ActiveConnection
		  
'		SQL = "Update tid Set lastnewtime=Now(),[counts]=[counts]+1,[lastposter]='"&UserID&"',[lasttitle2]='"&Subject&"'"&",[lasttitle]="&rs("TitleID")&""
		SQL = "Update tid Set lastnewtime=Now(),[counts]=[counts]+1,[lastposter]='"&UserID&"',[lasttitle]="&rs("TitleID")&""
		SQL = SQL & " Where tid=" & tid
		cmd.CommandText = SQL
		cmd.Execute
   
       rs.close'使用者文章數+1
	   SQL="SELECT * FROM id WHERE user='" & UserID & "'"
	   rs.Open SQL, conn,1,3
       rs("postsum")= rs("postsum")+1
       rs.Update

 sql="Select top 1 * From Titles Order By TitleID DESC"
		Set rs123=conn.Execute(sql) 
      
Response.Redirect "detail.asp?TitleID="&rs123("TitleID")&"&tid="&request("tid")&"&postname="&UserID
End If%>
</BODY></HTML>