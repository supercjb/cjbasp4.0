<%sql="select * from config_setting"
set rstitle=conn.execute(sql)

totalguest=rsguest("totalvisitor")+1					'���o�ثe��Ʈw����� �å[�J�����X�Ȥ��p�ƭ�
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
nowyear=cstr(year(DateAdd("h",rstitle("servertimezone"),now))-1911)					'���o���餧�~�r��
nowmonth=cstr(month(DateAdd("h",rstitle("servertimezone"),now)))						'���o���餧��r��
nowday=cstr(day(DateAdd("h",rstitle("servertimezone"),now)))							'���o���餧��r��
if len(nowyear)<2 then							'�Y���o�����餧�~�r�ꥼ���⥼�ƫh�b�e���0
  nowyear="0"+nowyear
end if
if len(nowmonth)<2 then							'�Y���o�����餧��r�ꥼ���⥼�ƫh�b�e���0
  nowmonth="0"+nowmonth
end if  
if len(nowday)<2 then							'�Y���o�����餧��r�ꥼ���⥼�ƫh�b�e���0
  nowday="0"+nowday
end if
nowdatestring=nowyear+nowmonth+nowday			'���餧����r��
nowmonthstring=nowyear+nowmonth					'���餧����r��
select case	month(DateAdd("h",rstitle("servertimezone"),now))							'���餧�u���r��
  case 1,2,3
    nowseasonstring=nowyear+"1"
  case 4,5,6
    nowseasonstring=nowyear+"2"
  case 7,8,9
    nowseasonstring=nowyear+"3"
  case 10,11,12
    nowseasonstring=nowyear+"4"
end select
if nowdatestring<>dayguestarea then				'�Y���o���������r��P��Ʈw���O��������r�ꤣ�P
  dayguestarea=nowdatestring      				'  �h�N���餧����r������s������r��
  dayguestu=dayguest							'  �ñN�줵��X�Ȳ��ܬQ��X��
  dayguest=1									'  �ñN�p���k1
else											'�Y�ۦP
  dayguest=dayguest+1							'  �h�N���餧�X�ȭp��+1
end if
if nowmonthstring<>monthguestarea then			'�Y���o���������r��P��Ʈw���O��������r�ꤣ�P
  monthguestarea=nowmonthstring					'  �h�N���餧����r������s������r��
  monthguestu=monthguest						'  �ñN�줵��X�Ȳ��ܤW��X��
  monthguest=1									'  �ñN�p���k1
else											'�Y�ۦP
  monthguest=monthguest+1						'  �h�N���餧��X�ȭp��+1
end if
if nowseasonstring<>seasonguestarea then		'�Y���o������u���r��P��Ʈw���O�����u���r�ꤣ�P
  seasonguestarea=nowseasonstring				'  �h�N���餧�u���r������s���u���r��
  seasonguestu=seasonguest						'  �ñN�줵�u�X�Ȳ��ܤW�u�X��
  seasonguest=1									'  �ñN�p���k1
else											'�Y�ۦP
  seasonguest=seasonguest+1						'  �h�N���餧�u�X�ȭp��+1
end if
if nowyear<>yearguestarea then					'�Y���o������~���r��P��Ʈw���O�����~���r�ꤣ�P
  yearguestarea=nowyear							'  �h�N���餧�~���r������s���~���r��
  yearguestu=yearguest							'  �ñN�줵�~�X�Ȳ��ܥh�~�X��
  yearguest=1									'  �ñN�p���k1
else											'�Y�ۦP
  yearguest=yearguest+1							'  �h�N���餧�~�X�ȭp��+1
end if
												'�N������o���p�ƭȦs�^��Ʈw
sql="update Visitorcounts set totalvisitor=" & totalguest & ",visitdate='" & dayguestarea & "',dayvisitor=" & dayguest & ",lastdayvisitor=" & dayguestu & ",visitmonth='" & monthguestarea & "',monthvisitor=" & monthguest & ",lastmonthvisitor=" & monthguestu & ",visitseason='" & seasonguestarea & "',seasonvisitor=" & seasonguest & ",lastseasonvisitor=" & seasonguestu & ",visityear='" & yearguestarea & "',yearvisitor=" & yearguest & ",lastyearvisitor=" & yearguestu
set rs=concount.execute(sql)

if rsguests.eof then							'�p�G �X�ȰO�� �����o�{�����s�u���X�ȰO�� �h�s�W�ӳX�ȭp��
  youguest=1
  sql="insert into Visitorrecords(visitor,counts) values('" & ipguest & "',1)"
else											'�Y �X�ȰO���� �o�{�����s�u���X�ȰO�� �h��s�p��
  sql="update Visitorrecords set counts=" & rsguests("counts")+1 & " where visitor='" & ipguest & "'"
  youguest=rsguests("counts")+1
end if
set rs=concount.execute(sql)%>