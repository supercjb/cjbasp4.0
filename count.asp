<%
'��έp�ƾ����O �N�ѱ��ϥέp�ƾ��������ǭȦܦ�
'countobj=[guest/join] guest=�X�ȭp�ƾ� join=�s���p�ƾ�
'�Ϥ�s���p�ƾ��ɤ��B�~�ܼ�
'countobjjoin=[add/read] add=�s�W�p�� read=Ū���p��

'���o�s�u��IP���}
ipguest=request.servervariables("remote_addr")	

'�s����Ʈw count.mdb
'set concount = server.createobject("ADODB.Connection")
'mdbpath = server.mappath("/db/forum_counter.mdb")
'concount.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & mdbpath

Response.buffer=true 
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Path = "Data Source=" & Server.MapPath("db/forum_counter.asp")
Set concount= Server.CreateObject("ADODB.Connection")
dbpwd="jet oledb:database password=admin;"
p1=Provider& dbpwd &Path
concount.Open P1

'�s����ƪ� �X�ȭp��
set rsguest = Server.CreateObject("ADODB.Recordset")
sqlguest="Select * From Visitorcounts"
rsguest.Open sqlguest, concount,1, 3

'�s����ƪ� �X�ȰO��
set rsguests = Server.CreateObject("ADODB.Recordset")
sqlguests="Select * From Visitorrecords where visitor='" & ipguest & "'"
rsguests.Open sqlguests, concount,1, 3

'�s����ƪ� �����p��
set rsjoin = Server.CreateObject("ADODB.Recordset")
sqljoin="Select * From Webcounts"
rsjoin.Open sqljoin, concount,1, 3

'�ˬd�X�ȬO�_���
if session("nowip")<>ipguest then					'�X�Ȭ��D���
  
  session("nowip")=ipguest
  select case countobj
    case "guest"			 						'�X�ȭp�ƾ��{���ҩl%>
      <!--#include file="guestadd.asp" -->
<%  case "join"										'�s���p�ƾ��{���ҩl
  end select
else												'�X�����%>
  <!--#include file="guestread.asp" -->
<%end if
concount.close%>