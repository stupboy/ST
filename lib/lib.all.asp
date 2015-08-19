<%
'##########Stupboy 个人自定义函数库#########
'##########2015.08.18              #########
'输出函数SC 
Sub sc(str)
Response.write str
End Sub
'菜单下来函数
function caidan(a)
caidan="<ul class='nav navbar-nav'>"&_
        "<!--<li class='active'><a href='#'>Link <span class='sr-only'>(current)</span></a></li>-->"&_
        "<!--<li><a href='#'>刷新</a></li>-->"&_
        "<li class='dropdown'>"&_
          "<a href='#' class='dropdown-toggle' data-toggle='dropdown' role='button' aria-haspopup='true' aria-expanded='false'>"&a&"<span class='caret'></span></a>"&_
          "<ul class='dropdown-menu'>"&_
            "<li><a href='http://www.baidu.com' target='MainF'>百度</a></li>"&_
            "<li><a href='#'>Another action</a></li>"&_
            "<li><a href='#'>Something else here</a></li>"&_
            "<li role='separator' class='divider'></li>"&_
            "<li><a href='#'>Separated link</a></li>"&_
            "<li role='separator' class='divider'></li>"&_
            "<li><a href='#'>One more separated link</a></li>"&_
          "</ul>"&_
        "</li>"&_
      "</ul>"
end function 


%>