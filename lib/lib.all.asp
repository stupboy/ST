<%
'########## Stupboy 个人自定义函数库       #########
'########## UPDATE 2015.08.18              #########
'--函数汇总 及功能说明--
'-1. SC              输出函数-
'-2. caidan(a,b)     菜单输出函数,a为菜单名,b为"子菜单名$链接|子菜单$链接"的格式-
'-3. LimitCheck(a)   权限检测函数，若无权限则终端输出-
'-4. qx(a)           判断是否有权限，返回boolen值 TRUE OR FALSE-
'-5. str_x(x,y)      字符补位函数x为原字符,y为位数不足用0补齐-
'-6. date2str(x,y)   日期转字符函数，x为日期，y为类型，y为1则到日150801，y为2则到秒150801120025-
'-7. DanHao(x)       单号生成函数，x为单号前缀，后连接当期日期【类型2】-
'-8. getip()         获取IP函数-
'-
'-
'-函数明细列表-
'-输出函数SC -
Sub sc(str)
Response.write str
End Sub
'-菜单下拉显示函数,a为菜单名称,b为菜“单名$链接”的格式-
function caidan(a,b)
 mx=split(b,"|")     '-b为菜单名称及链接，多个菜单用“|”区分开，用SPLIT函数拆为数组-
 ms=ubound(mx,1)     
 caidan1="<ul class='nav navbar-nav'>"&_  
        "<!--<li class='active'><a href='#'>Link <span class='sr-only'>(current)</span></a></li>-->"&_
        "<!--<li><a href='#'>刷新</a></li>-->"&_
        "<li class='dropdown'>"&_
        "<a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>"&a&"<span class='caret'></span></a>"&_
        "<ul class='dropdown-menu'>"
 for i = 0 to ms            '-循环输出数组中的菜单 For循环-
    mt=split(mx(i),"$")     '-用$区分菜单名和链接-
    caidan3=caidan3&"<li><a href='"&mt(1)&"' target='MainF'>"&mt(0)&"</a></li>"    '-菜单字符串的拼接-
 next                       '-循环输出结束-
 caidan2="<!--<li role='separator' class='divider'></li>-->"&_                     
        "<!--<li><a href='#'>One more separated link</a></li>-->"&_
        "</ul>"&_
        "</li>"&_
        "</ul>"
 caidan=caidan1&caidan3&caidan2                                                     '-字符串的拼接输出-
end function 
'-权限检测函数[中断输出]-
sub LimitCheck(a)                                         
 if instr(session("session(UserLimit)"),a)=0 and len(a&"0")>1 then         
  sc "没有权限，权限代码：" & a
  response.end                                            
 end if                                                   
end sub  
'-权限检测函数[输出返回值1为是0为否]                                                '-函数结束-
function qx(a)
 if instr(session("session(UserLimit)"),a)=0 then 
   qx=false
 else
   qx=true
 end if   
end function 
'--字符转多位数--
function str_x(x,y)
 if len(trim(x))<y then
  dim a,b
  a=y-len(trim(x))
  for i = 1 to a
  b=b&"0"
  next 
  str_x=b&x
 else 
  str_x=x
 end if 
end function
'-日期转字符函数 1为到日 2为到秒-
function date2str(x,y) 
 a=right(year(x),2)
 if y=1 then 
 date2str=a&str_x(month(x),2)&str_x(day(x),2)
 elseif y=2 then 
 date2str=a&str_x(month(x),2)&str_x(day(x),2)&str_x(hour(x),2)&str_x(minute(x),2)&str_x(second(x),2)
 end if 
end function
'-单号生成函数-
function DanHao(x)
 DanHao=x&date2str(now(),2)
end function 
'-IP获取函数-
Private Function getIP()   
Dim strIPAddr   
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then   
strIPAddr = Request.ServerVariables("REMOTE_ADDR")   
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then   
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)   
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then   
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)   
Else   
strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")   
End If   
getIP = Trim(Mid(strIPAddr, 1, 30))   
End Function

%>