    <div align="center">
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="16">
        <tr>
          <td width="28%" height="7" background="images/Button.gif">
          <font color="#333333" size="2">&nbsp;<img src="images/icon_online.gif"><b>&nbsp;此 分 類 論 壇 線 上 狀 態</b></font></td>
          <td width="72%" height="7" background="images/Button.gif" bgcolor="#EFF4FE">
          <img border="0" src="images/online.gif"><font size="2" color="#333333">線上名單圖例： <%sql="select * from member_sort order by fav_size desc"
		  set rsaa=conn.execute(sql)
		  while not rsaa.eof
		  response.write "&nbsp;&nbsp;<img src=images/"&rsaa("pic")&">"&rsaa("levelname")
		  rsaa.movenext
		  wend
		  %>&nbsp; </font><font size="2">&nbsp;<img border="0" src="images/messages5.gif">訪客</font></font></td>
        </tr>
        <tr>
          <td width="100%" height="19" bgcolor="#FFFFFF" colspan="2" background="images/postbit_lightbg.gif"><font size=-1>
<%sql="select * from config_setting"
set rstitle=conn.execute(sql)
 
    UserID=request.cookies("UserID")
    Set rson= Server.CreateObject("ADODB.Recordset") 
    Set rson1= Server.CreateObject("ADODB.Recordset") 
    Set rson2= Server.CreateObject("ADODB.Recordset")
    SQL="SELECT * FROM online where tid='"&request("tid")&"' order by id"
    SQL1="SELECT * FROM online where name='"&UserID&"'"
    SQL2="SELECT * FROM online where name='"&Cstr(Request.ServerVariables("REMOTE_ADDR"))&"'"
    rson1.Open SQL1, conn,1,3
    rson2.Open SQL2, conn,1,3
    rson.Open SQL, conn,1,3
    if rson1.EOF AND not UserID=empty then
      rson1.addnew
      rson1("name")=UserID
      rson1("ip")=Cstr(Request.ServerVariables("REMOTE_ADDR"))
      rson1("time")=now()
      rson1("tid")=request("tid")
      rson1.Update
    elseif UserID=empty and rson2.EOF then
      rson1.addnew
      rson1("name")=Cstr(Request.ServerVariables("REMOTE_ADDR"))
      rson1("ip")=Cstr(Request.ServerVariables("REMOTE_ADDR"))
      rson1("time")=now()
      rson1("tid")=request("tid")
      rson1.Update
    end if
	
dim online_sum,online_guest
  online_sum=0
  online_guest=0

   while not rson.EOF
       SQL2="SELECT * FROM id where user='"&rson("name")&"'"
       set rsname=conn.execute(SQL2)
       
	   newtime=datediff("n",rson("time"),now)

    if newtime>5 then
     rson.delete'刪除名單
	else
		if rson("tid")="0" or rson("tid")=empty then
			wh="首頁"
		else
        	sql="SELECT * FROM tid where [tid]="&Cint(rson("tid"))
        	set rswh=conn.execute(sql)
        	wh=rswh("tidname")
		end if
		if not rsname.EOF then
        	sql="select * from member_sort where level="&rsname("level")
			set rs_mem=conn.execute(sql)
			response.write "<img src=images/"&rs_mem("pic")&" alt='最後活動:"&DateAdd("h",rstitle("servertimezone")+rsname("timezone"),rson("time"))&"，正在"&wh&"，閒滯"&newtime&"分"&"'><a href=u2u.asp?touser="&rson("name")&">"&rson("name")&"</a>&nbsp;"
			online_sum=online_sum+1
		else
			response.write "<img src=images/messages5.gif alt='"&rson("ip")&"最後活動:"&rson("time")&"，正在"&wh&"，閒滯"&newtime&"分"&"'>訪客&nbsp;"
        	online_guest=online_guest+1
		end if		
	end if                  
         rson.movenext  	
	wend

   response.write "<br><font size=-1>共有<font color=red><b>"&online_sum&"</font></b>位會員在此，"
   response.write "共有<font color=red><b>"&online_guest&"</font></b>位訪客在此。</font>"%>
	          </td>
        </tr>
      </table>
</div>