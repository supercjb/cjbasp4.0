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
<!--#include file="admin/broad.asp"--><BR>
<%mode=request("mode")
if mode="top_lock" then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3
rs("top_lock")=1
rs("top_lock_order")=request("level")
rs.update
rs.close
set rs=nothing
response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
elseif mode="cancel_lock" then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3
	if request("level")=0 then
      rs("top_lock")=0
      rs.update
	  response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
	elseif request("level")=1 or request("level")=2 or request("level")=3 then
	  rs("top_lock")=1
      rs("top_lock_order")=request("level")
	  rs.update
      rs.close
      set rs=nothing
      response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
     end if
elseif mode="reply_lock" then
   	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3
	rs("reply_lock")=request("lock")
	rs.update
	response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
elseif mode="tit_move" then
   	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3
sql="select * from tid where tid="&rs("tid")'�ק��ª����̫��s
Set rs_delas0=Server.CreateObject("ADODB.Recordset")
rs_delas0.Open sql, conn,1,3
	rs_delas0("counts")=rs_delas0("counts")-1
	rs_delas0("replycounts")=rs_delas0("replycounts")-rs("replycount")
	rs_delas0("lasttitle")=0
	rs_delas0.update
	rs("tid")=request("tid")
	rs("LastNewsDate")=now()
	rs("lastposter")=UserID
	rs.update
	rs.close
	sql="select * from Details where TitleID="&request("TitleID")
    rs.Open sql,conn,1,3
	while not rs.eof
	rs("tid")=request("tid")
	rs.update
	rs.movenext
	wend
	rs.close
    sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3

	sql="select * from tid where tid="&request("tid")'�ק�s�����̫��s
	Set rs_delas=Server.CreateObject("ADODB.Recordset")
    rs_delas.Open sql, conn,1,3
	rs_delas("counts")=rs_delas("counts")+1
	rs_delas("replycounts")=rs_delas("replycounts")+rs("replycount")
	rs_delas("lasttitle")=rs("TitleID")
	Words=""&rs("Words")&chr(10)&chr(10)&chr(10)&"�i "&UserID&" �N���g�D�D�� "&rs_delas0("tidname")&" ���� "&rs_delas("tidname")&" ��x�W�ɶ� "&DateAdd("h",rstitle("servertimezone"),Now())&" �j"
	rs("Words")=Words	
	rs.update
	     sql="select * from u2u"'�t�γq���@���ಾ
		 Set pm=Server.CreateObject("ADODB.Recordset")
		 pm.Open sql,conn,1,3
		 pm.addnew
		 pm("fromuser")="�t��"
		 pm("totitle")="[�t��]�峹�ಾ�q���I�ФŦ^��"
		 pm("touser")=rs("Name")
         pm("totext")="�z���峹 <b><a href=Detail.asp?TitleID="&rs("TitleID")&"&tid="&rs("tid")&"&post_name="&rs("Name")&">"&rs("Subject")&"</b></a> �w�� "&rs_delas0("tidname")&" �Q�ಾ�� "&rs_delas("tidname")&" "
		 pm.update
		 
	rs_delas.update
	rs_delas.close
	rs.close
	pm.close
	response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
elseif mode="reply_count" then
    sql="select * from Details where TitleID="&request("TitleID")
	set rs=conn.execute(sql)
	sum=0
	while not rs.eof
	sum=sum+1
	rs.movenext
	wend
	Set rs1=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs1.Open sql,conn,1,3
	rs1("replycount")=sum
	rs1.update
	rs1.close
	set rs1=nothing
	wri=wri & "<center>"
	wri=wri & "��s����,�@���"
    wri=wri & sum
	wri=wri & "���^�Цb���D�D<br>"
	wri=wri & "<br><a href=forum.asp?tid=" & request("tid")
    wri=wri & ">�^�Q�װ�</a>"
    response.write wri
elseif mode="tid_count" then
    sql="select * from Titles where tid="&request("tid")
	set rs=conn.execute(sql)
    sql="select * from Details where tid="&request("tid")
	set rs1=conn.execute(sql)
	sum=0'�D�D
	while not rs.eof
	sum=sum+1
	rs.movenext
	wend
	sum1=0'�^��
	while not rs1.eof
	sum1=sum1+1
	rs1.movenext
	wend
	Set rs2=Server.CreateObject("ADODB.Recordset")
	sql="select * from tid where tid="&request("tid")
	rs2.Open sql,conn,1,3
	rs2("counts")=sum
	rs2("replycounts")=sum1
	rs2.update
	rs2.close
	wri=wri & "<center>"
	wri=wri & "��s����,�@���"
    wri=wri & sum
	wri=wri & "���D�D�b���D�D�A"
	wri=wri & sum1
	wri=wri & "���^�Цb���Q�װ�<br>"
	wri=wri & "<br><a href=forum.asp?tid=" & request("tid")
    wri=wri & ">�^�Q�װ�</a>"
    response.write wri
elseif mode="best" then
	Set rs=Server.CreateObject("ADODB.Recordset")
	sql="select * from Titles where TitleID="&request("TitleID")
	rs.Open sql,conn,1,3
    rs("best")=request("level")
	rs.update
	response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")
end if%><BR><BR><!--#include file="foot.asp"-->
</BODY></HTML>