<%
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
Path = "Data Source=" & Server.MapPath("../db/cjb-asp.mdb")
Set conn= Server.CreateObject("ADODB.Connection")
p1=Provider&Path
conn.Open P1
%>