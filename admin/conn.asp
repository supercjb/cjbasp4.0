<%'--------------------------------------------------------------有密碼的連接方式
Response.buffer=true 
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Path = "Data Source=" & Server.MapPath("../db/forum.mdb")'資料庫
Set conn= Server.CreateObject("ADODB.Connection")
dbpwd="jet oledb:database password=cjbasp4.0;"
p1=Provider& dbpwd &Path
conn.Open P1
'--------------------------------------------------------------無密碼的連結方式
'Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
'Path = "Data Source=" & Server.MapPath("/db/forum.mdb")
'Set conn= Server.CreateObject("ADODB.Connection")
'p1=Provider&Path
'conn.Open P1
sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
'以下是以UserID為狀態來分辨是否在線上
if not request.cookies("UserID")=empty then
UserID=request.cookies("UserID")
SQL1="SELECT * FROM online where name='"&UserID&"'"
Set sr1= Server.CreateObject("ADODB.Recordset")
sr1.Open SQL1, conn,1,3
if not sr1.eof then
sr1("time")=now'讀取到conn.vs檔時就更新最後活動時間
response.cookies("tid")=""
if request("tid")<>empty then
sr1("detail")=""
sr1("tid")=request("tid")
response.cookies("tid")=request("tid")
else
sr1("detail")=""
sr1("tid")="0"
end if
sr1.update
end if

else

SQL2="SELECT * FROM online where name='"&Request.ServerVariables("REMOTE_ADDR")&"'"
Set sr2= Server.CreateObject("ADODB.Recordset")
sr2.Open SQL2, conn,1,3
if not sr2.eof then
sr2("time")=now'讀取到conn.vs檔時就更新最後活動時間
response.cookies("tid")=""
if request("tid")<>empty then
   sr2("detail")=""
   sr2("tid")=request("tid")
   response.cookies("tid")=request("tid")
else
'sr2("detail")=""
sr2("tid")="0"
end if
sr2.update
end if

end if

'以下是刪除線上名單(刪除會員登入前的訪客資格)**************************************
  SQL="SELECT * FROM online where name='"&Request.ServerVariables("REMOTE_ADDR")&"'"
  Set rs12= Server.CreateObject("ADODB.Recordset") 
  rs12.open SQL,conn,1,3
  if not rs12.EOF and request.cookies("UserPassed")="yes" then
  rs12.delete
  rs12.close
  set rs12=nothing
  end if
'以上結束***************************************************************************

'判斷是否為該論壇的使用者

	
If not (request.cookies("UserPassed") = "yes" Or request.cookies("bbsid")=bbsid) Then 
ExpireDate=DateAdd("d",-100,Date)
response.cookies.Expires=FormatDateTime(ExpireDate)
end if

'假釋帳號不存在,刪除之前存在的cookies
SQL = "Select * From id Where user='"&request.cookies("UserID")&"' and password='"&request.cookies("Password")&"'"
set rs1=conn.execute(SQL)

IF request.cookies("UserID")=empty or rs1.eof THEN
UserID="訪客"
response.cookies("Passed2")="no"
response.cookies("admin_pass")=empty
 response.cookies("UserPassed")="no"
session.abandon
ELSEIF not rs1.EOF then
UserID=request.cookies("UserID")
end if
%>
