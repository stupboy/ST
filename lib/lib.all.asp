<%
'##########Stupboy 个人自定义函数库#########
'##########2015.08.18              #########
'输出函数SC 
Sub sc(str)
Response.write str
End Sub

'菜单下拉显示函数,a为菜单名称,b为菜“单名$链接”的格式
function caidan(a,b)
 mx=split(b,"|")     'b为菜单名称及链接，多个菜单用“|”区分开，用SPLIT函数拆为数组
 ms=ubound(mx,1)     '求数组个数
 caidan1="<ul class='nav navbar-nav'>"&_  
        "<!--<li class='active'><a href='#'>Link <span class='sr-only'>(current)</span></a></li>-->"&_
        "<!--<li><a href='#'>刷新</a></li>-->"&_
        "<li class='dropdown'>"&_
        "<a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>"&a&"<span class='caret'></span></a>"&_
        "<ul class='dropdown-menu'>"
 for i = 0 to ms            '循环输出数组中的菜单 For循环
    mt=split(mx(i),"$")     '用$区分菜单名和链接
    caidan3=caidan3&"<li><a href='"&mt(1)&"' target='MainF'>"&mt(0)&"</a></li>"    '菜单字符串的拼接
 next                       '循环输出结束
 caidan2="<!--<li role='separator' class='divider'></li>-->"&_                     
        "<!--<li><a href='#'>One more separated link</a></li>-->"&_
        "</ul>"&_
        "</li>"&_
        "</ul>"
 caidan=caidan1&caidan3&caidan2                                                     '字符串的拼接输出
end function 

'权限检测函数[中断输出]
sub LimitCheck(a)                                         'a为为需检测值，c为条件为真时输出值
 if instr(session("session(UserLimit)"),a)=0 then         '如果没有权限
  sc "没有权限，权限代号"&a                               '权限提示
  response.end                                            '输出提示
 end if                                                   '条件判断结束
end sub  

'权限检测函数[输出返回值1为是0为否]                                                 '函数结束
function qx(a)
 if instr(session("session(UserLimit)"),a)=0 then 
   qx=false
 else
   qx=true
 end if   
end function 

'IP获取函数
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