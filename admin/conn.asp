<%'--------------------------------------------------------------���K�X���s���覡
Response.buffer=true 
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Path = "Data Source=" & Server.MapPath("../db/forum.mdb")'��Ʈw
Set conn= Server.CreateObject("ADODB.Connection")
dbpwd="jet oledb:database password=cjbasp4.0;"
p1=Provider& dbpwd &Path
conn.Open P1
'--------------------------------------------------------------�L�K�X���s���覡
'Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
'Path = "Data Source=" & Server.MapPath("/db/forum.mdb")
'Set conn= Server.CreateObject("ADODB.Connection")
'p1=Provider&Path
'conn.Open P1
sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
'�H�U�O�HUserID�����A�Ӥ���O�_�b�u�W
if not request.cookies("UserID")=empty then
UserID=request.cookies("UserID")
SQL1="SELECT * FROM online where name='"&UserID&"'"
Set sr1= Server.CreateObject("ADODB.Recordset")
sr1.Open SQL1, conn,1,3
if not sr1.eof then
sr1("time")=now'Ū����conn.vs�ɮɴN��s�̫ᬡ�ʮɶ�
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
sr2("time")=now'Ū����conn.vs�ɮɴN��s�̫ᬡ�ʮɶ�
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

'�H�U�O�R���u�W�W��(�R���|���n�J�e���X�ȸ��)**************************************
  SQL="SELECT * FROM online where name='"&Request.ServerVariables("REMOTE_ADDR")&"'"
  Set rs12= Server.CreateObject("ADODB.Recordset") 
  rs12.open SQL,conn,1,3
  if not rs12.EOF and request.cookies("UserPassed")="yes" then
  rs12.delete
  rs12.close
  set rs12=nothing
  end if
'�H�W����***************************************************************************

'�P�_�O�_���ӽ׾ª��ϥΪ�

	
If not (request.cookies("UserPassed") = "yes" Or request.cookies("bbsid")=bbsid) Then 
ExpireDate=DateAdd("d",-100,Date)
response.cookies.Expires=FormatDateTime(ExpireDate)
end if

'�����b�����s�b,�R�����e�s�b��cookies
SQL = "Select * From id Where user='"&request.cookies("UserID")&"' and password='"&request.cookies("Password")&"'"
set rs1=conn.execute(SQL)

IF request.cookies("UserID")=empty or rs1.eof THEN
UserID="�X��"
response.cookies("Passed2")="no"
response.cookies("admin_pass")=empty
 response.cookies("UserPassed")="no"
session.abandon
ELSEIF not rs1.EOF then
UserID=request.cookies("UserID")
end if
%>
