<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- 引用自定义【函数库】及【数据库连接】 -->
<!--#include file="../lib/lib.all.asp" -->
<!--#include file="../inc/sql.conn.asp" -->
<%
    Response.Charset = "utf-8"                                   '输出为UTF8格式
	ip=getIP()                                                   '获取IP值
    username=request.Form("UserName")                            '获取传递用户名和密码
	pwd=request.Form("PassWord")                                 '查询表User_Info
	sql="select * from [User_Info] where UserName = '"&username&"' and Password = '"&pwd&"'"
	set rs=server.CreateObject("adodb.recordset")                '创建记录集对象                
	rs.open sql,conn,1,3                                         '打开连接字符串
	if not rs.eof then                                           '如果记录不为空则
	  session("UserName")=trim(rs("UserName"))                   '设置SESSION值USERNAME
	  session("UserLimit")=trim(rs("UserLimit"))                 '设置SESSION值USERLimit权限字符串
	  session("RealName")=trim(rs("RealName"))                   '设置SESSION值RealName
	  rs.close                                                   '关闭记录集
	  set rs = nothing                                           '关闭记录集
	  '增加登入日志
      sql="select * from [Sys_log]"                              '登入记录表SYS_LOG
	  set rs=server.CreateObject("adodb.recordset")              '创建记录集对象
	  rs.open sql,conn,1,3                                       '打开连接
	  rs.addnew()                                                '增加登入记录
	  rs("UserName")=session("UserName")                         '用户名                      
	  rs("loginip")=ip                                           '登入IP
	  rs.update                                                  '更新字段
	  rs.close                                                   '关闭记录集
	  conn.close                                                 '关闭连接
	  set conn = Nothing                                         '关闭连接
	  response.Redirect("../index.asp")                          '登入成功跳转主页面
	else                                                         '如果查询无用户记录
     response.write"<SCRIPT language=JavaScript>alert('密码输入错误');" '弹窗提示密码错误
     response.write"location.href='../pro/action.login.asp'</SCRIPT>"   '验证码错误返回登入页面
	end if                                                       '结束条件判断
	rs.close                                                     '关闭记录集
    set rs=nothing                                               '关闭记录集
    conn.close                                                   '关闭连接
    set conn=nothing                                             '关闭连接
%>