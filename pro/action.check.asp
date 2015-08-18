<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- 引用自定义函数库及数据库连接 -->
<!--#include file="../lib/lib.all.asp" -->
<!--#include file="../inc/sql.conn.asp" -->
<%
    Response.Charset = "utf-8"                                   '输出为UTF8格式
    username=request.Form("UserName")                            '获取传递用户名和密码
	pwd=request.Form("PassWord")                                 '查询表User_Info
	sql="select * from [User_Info] where UserName = '"&username&"' and Password = '"&pwd&"'"
	set rs=server.CreateObject("adodb.recordset")                
	rs.open sql,conn,1,3
	if not rs.eof then
	  session("UserName")=rs("UserName")                         '设置SESSION值USERNAME
	  session("RealName")=rs("RealName")                         '设置SESSION值RealName
	  rs.close                                                   '关闭记录集
	  set rs = nothing                                           '关闭记录集
	  '增加登入记录
      sql="select * from [Sys_log]"                              '登入记录表SYS_LOG
	  set rs=server.CreateObject("adodb.recordset")
	  rs.open sql,conn,1,3
	  rs.addnew()                                                '增加登入记录
	  rs("UserName")=1'session("UserName")                       '用户名                      
	  'rs("logip")=ip                                            '登入IP
	  rs.update
	  rs.close
	  conn.close                                                 '关闭连接
	  set conn = Nothing                                         '关闭连接
	  response.Redirect("../index.asp")                          '登入成功跳转主页面
	else
     response.write"<SCRIPT language=JavaScript>alert('密码输入错误');" '弹窗提示密码错误
     response.write"location.href='../pro/action.login.asp'</SCRIPT>"   '验证码错误返回登入页面
	end if  
	rs.close
    set rs=nothing
    conn.close
    set conn=nothing
%>