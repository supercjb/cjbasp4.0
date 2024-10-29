<!--#include file="conn.asp"-->
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.Cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
tit=request("tit")
text=request("text")
if request("to")=12 then
sql="select * from id"
else
sql="select * from id where level="&request("to")
end if

Set rs1=Server.CreateObject("ADODB.Recordset")
rs1.Open sql,conn,1,3

sql="select * from u2u"
Set rs2=Server.CreateObject("ADODB.Recordset")
rs2.Open sql,conn,1,3

i=0

while not rs1.eof
 rs1("havepm")=1
 rs1.Update
	rs2.AddNew
	rs2("fromuser")="系統公告[請勿回覆]"
    rs2("touser")=rs1("user")
	rs2("totitle")=tit
	rs2("totext")=text
	rs2("time")=now
	rs2.update
     i=i+1
response.write "傳送給第&nbsp;"&i&"&nbsp;位~&nbsp;"&rs2("touser")&".......<br>"
rs1.movenext
wend
  response.write "<br>完成~共寄給&nbsp;"&i&"&nbsp;位會員!"
end if%>