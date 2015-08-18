<%
'1.如果SESSION USERNAME值为空则跳转登入页面
if session("UserName")=""  then                                                    '如果用户名SESSION值为空
response.write "<script>parent.location.href='pro/action.login.asp';</script>"     '跳转登入页面
response.End()                                                                     '输出终止
end if                                                                             '条件结束
'2.设置SESSION过期时间
Session.Timeout=15                                                                 'SEESION有效时间为15分钟 
'3.网站配置文件
'网站标题
WebTitle="销售库存管理"


%>