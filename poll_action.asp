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
<%UserID=Request("Name2")
Subject=Request("Subject")
Words=Request("Words")
Color1=Request("Color1")
tid=Request("tid")
subc=Request("subc")
R1=Request("R1")
file=request.cookies("filename")
T1=request("T1")
T2=request("T2")
T3=request("T3")
T4=request("T4")
T5=request("T5")
T6=request("T6")
T7=request("T7")
T8=request("T8")
T9=request("T9")
T10=request("T10")
if request("save_date")<>"不設定" then
save_date=request("save_date")
end if 
reply_pm=request("reply_pm")
sql="select * from config_setting"
set rstitle=conn.execute(sql)
	Set rs2= Server.CreateObject("ADODB.Recordset") '檢查帳號是否被用
   SQL="SELECT * FROM id WHERE user='" & UserID & "'"
   rs2.Open SQL, conn,1,1
IF  request.cookies("UserPassed") <> "yes" Or rs2.EOF Then 
	Response.redirect "index.asp?msg=請登入!"%>
<%ElseIf Subject=Empty Or Words=Empty Or UserID =Empty Or Color1 =Empty Or tid =Empty or (T1=empty and T2=empty and T3=empty and T4=empty and T5=empty and T6=empty and T7=empty and T8=empty and T9=empty and T10=empty) Then %>
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
	rs("poll")=1

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
  'poll----------------------
  	set rspoll=Server.CreateObject("ADODB.Recordset")
	rspoll.open "poll",conn,1,3
	rspoll.addnew
	rspoll("titleid")=rs123("TitleID")
	rspoll("T1")=T1
	rspoll("T2")=T2
	rspoll("T3")=T3
	rspoll("T4")=T4
	rspoll("T5")=T5
	rspoll("T6")=T6
	rspoll("T7")=T7
	rspoll("T8")=T8
	rspoll("T9")=T9
	rspoll("T10")=T10
	rspoll.update
  'poll-end------------------
Response.Redirect "detail.asp?TitleID="&rs123("TitleID")&"&tid="&request("tid")&"&postname="&UserID
End If%>
</BODY></HTML>