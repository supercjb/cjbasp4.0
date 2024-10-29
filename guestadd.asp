<%sql="select * from config_setting"
set rstitle=conn.execute(sql)

totalguest=rsguest("totalvisitor")+1					'取得目前資料庫中資料 並加入此次訪客之計數值
dayguestarea=rsguest("visitdate")
dayguest=rsguest("dayvisitor")
dayguestu=rsguest("lastdayvisitor")
monthguestarea=rsguest("visitmonth")
monthguest=rsguest("monthvisitor")
monthguestu=rsguest("lastmonthvisitor")
seasonguestarea=rsguest("visitseason")
seasonguest=rsguest("seasonvisitor")
seasonguestu=rsguest("lastseasonvisitor")
yearguestarea=rsguest("visityear")
yearguest=rsguest("yearvisitor")
yearguestu=rsguest("lastyearvisitor")
nowyear=cstr(year(DateAdd("h",rstitle("servertimezone"),now))-1911)					'取得今日之年字串
nowmonth=cstr(month(DateAdd("h",rstitle("servertimezone"),now)))						'取得今日之月字串
nowday=cstr(day(DateAdd("h",rstitle("servertimezone"),now)))							'取得今日之日字串
if len(nowyear)<2 then							'若取得之今日之年字串未滿兩未數則在前方補0
  nowyear="0"+nowyear
end if
if len(nowmonth)<2 then							'若取得之今日之月字串未滿兩未數則在前方補0
  nowmonth="0"+nowmonth
end if  
if len(nowday)<2 then							'若取得之今日之日字串未滿兩未數則在前方補0
  nowday="0"+nowday
end if
nowdatestring=nowyear+nowmonth+nowday			'今日之日期字串
nowmonthstring=nowyear+nowmonth					'今日之月期字串
select case	month(DateAdd("h",rstitle("servertimezone"),now))							'今日之季期字串
  case 1,2,3
    nowseasonstring=nowyear+"1"
  case 4,5,6
    nowseasonstring=nowyear+"2"
  case 7,8,9
    nowseasonstring=nowyear+"3"
  case 10,11,12
    nowseasonstring=nowyear+"4"
end select
if nowdatestring<>dayguestarea then				'若取得之今日日期字串與資料庫中記載之日期字串不同
  dayguestarea=nowdatestring      				'  則將今日之日期字串視為新的日期字串
  dayguestu=dayguest							'  並將原今日訪客移至昨日訪客
  dayguest=1									'  並將計數歸1
else											'若相同
  dayguest=dayguest+1							'  則將今日之訪客計數+1
end if
if nowmonthstring<>monthguestarea then			'若取得之今日月期字串與資料庫中記載之月期字串不同
  monthguestarea=nowmonthstring					'  則將今日之月期字串視為新的月期字串
  monthguestu=monthguest						'  並將原今月訪客移至上月訪客
  monthguest=1									'  並將計數歸1
else											'若相同
  monthguest=monthguest+1						'  則將今日之月訪客計數+1
end if
if nowseasonstring<>seasonguestarea then		'若取得之今日季期字串與資料庫中記載之季期字串不同
  seasonguestarea=nowseasonstring				'  則將今日之季期字串視為新的季期字串
  seasonguestu=seasonguest						'  並將原今季訪客移至上季訪客
  seasonguest=1									'  並將計數歸1
else											'若相同
  seasonguest=seasonguest+1						'  則將今日之季訪客計數+1
end if
if nowyear<>yearguestarea then					'若取得之今日年期字串與資料庫中記載之年期字串不同
  yearguestarea=nowyear							'  則將今日之年期字串視為新的年期字串
  yearguestu=yearguest							'  並將原今年訪客移至去年訪客
  yearguest=1									'  並將計數歸1
else											'若相同
  yearguest=yearguest+1							'  則將今日之年訪客計數+1
end if
												'將今日取得之計數值存回資料庫
sql="update Visitorcounts set totalvisitor=" & totalguest & ",visitdate='" & dayguestarea & "',dayvisitor=" & dayguest & ",lastdayvisitor=" & dayguestu & ",visitmonth='" & monthguestarea & "',monthvisitor=" & monthguest & ",lastmonthvisitor=" & monthguestu & ",visitseason='" & seasonguestarea & "',seasonvisitor=" & seasonguest & ",lastseasonvisitor=" & seasonguestu & ",visityear='" & yearguestarea & "',yearvisitor=" & yearguest & ",lastyearvisitor=" & yearguestu
set rs=concount.execute(sql)

if rsguests.eof then							'如果 訪客記錄 中未發現本次連線之訪客記錄 則新增該訪客計數
  youguest=1
  sql="insert into Visitorrecords(visitor,counts) values('" & ipguest & "',1)"
else											'若 訪客記錄中 發現本次連線之訪客記錄 則更新計數
  sql="update Visitorrecords set counts=" & rsguests("counts")+1 & " where visitor='" & ipguest & "'"
  youguest=rsguests("counts")+1
end if
set rs=concount.execute(sql)%>