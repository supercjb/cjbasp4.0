<%
'選用計數器類別 將由欲使用計數器之網頁傳值至此
'countobj=[guest/join] guest=訪客計數器 join=連結計數器
'使月連結計數器時之額外變數
'countobjjoin=[add/read] add=新增計數 read=讀取計數

'取得連線之IP為址
ipguest=request.servervariables("remote_addr")	

'連結資料庫 count.mdb
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

'連結資料表 訪客計數
set rsguest = Server.CreateObject("ADODB.Recordset")
sqlguest="Select * From Visitorcounts"
rsguest.Open sqlguest, concount,1, 3

'連結資料表 訪客記錄
set rsguests = Server.CreateObject("ADODB.Recordset")
sqlguests="Select * From Visitorrecords where visitor='" & ipguest & "'"
rsguests.Open sqlguests, concount,1, 3

'連結資料表 網頁計數
set rsjoin = Server.CreateObject("ADODB.Recordset")
sqljoin="Select * From Webcounts"
rsjoin.Open sqljoin, concount,1, 3

'檢查訪客是否灌水
if session("nowip")<>ipguest then					'訪客為非灌水
  
  session("nowip")=ipguest
  select case countobj
    case "guest"			 						'訪客計數器程式啟始%>
      <!--#include file="guestadd.asp" -->
<%  case "join"										'連結計數器程式啟始
  end select
else												'訪客灌水%>
  <!--#include file="guestread.asp" -->
<%end if
concount.close%>