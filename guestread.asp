<%totalguest=rsguest("totalvisitor")
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
if not rsguests.eof then
  youguest=rsguests("counts")
end if%>