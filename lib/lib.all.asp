<%
'##########Stupboy 个人自定义函数库#########
'##########2015.08.18              #########
'输出函数SC 
Sub sc(str)
Response.write str
End Sub

'菜单下拉显示函数,a为菜单名称,b为菜“单名$链接”的格式
function caidan(a,b)
mx=split(b,"$")
ms=ubound(mx,1)
caidan="<ul class='nav navbar-nav'>"&_
        "<!--<li class='active'><a href='#'>Link <span class='sr-only'>(current)</span></a></li>-->"&_
        "<!--<li><a href='#'>刷新</a></li>-->"&_
        "<li class='dropdown'>"&_
          "<a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>"&a&"<span class='caret'></span></a>"&_
          "<ul class='dropdown-menu'>"&_
            "<li><a href='"&mx(1)&"' target='MainF'>"&mx(0)&"</a></li>"&_
            "<!--<li role='separator' class='divider'></li>-->"&_
            "<!--<li><a href='#'>One more separated link</a></li>-->"&_
          "</ul>"&_
        "</li>"&_
      "</ul>"
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
function rs()
set rs=Server.CreateObject("ADODB.RecordSet")
end function 
%>