<%
'1.如果SESSION USERNAME值为空则跳转登入页面
'如果SESSION为空则跳转主页面
If session("UserName")="" then                            
response.write "<script>parent.location.href='pro/action.login.asp';</script>"    
response.End()                                                                   
end If                                                                             
'2.设置SESSION过期时间
Session.Timeout=15                                                                 'SEESION有效时间为15分钟 
'3.网站配置文件
'网站标题
WebTitle="网站主页"
'公司名称
ComName="希尼亚D"
%>