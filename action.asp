<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../lib/lib.all.asp" -->
<!--#include file="../inc/sql.conn.asp" -->
<%
    Response.Charset = "utf-8"
    username=request.Form("UserName")
	pwd=request.Form("PassWord")
	sql="select * from [User_Info] where UserName = '"&username&"' and Password = '"&pwd&"'"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
	  session("UserName")=rs("UserName")
	  session("RealName")=rs("RealName")
	  rs.close
	  set rs = nothing
	  '增加登入记录
      sql="select * from [sys_log]"
	  set rs=server.CreateObject("adodb.recordset")
	  rs.open sql,conn,1,3
	  rs.addnew()
	  rs("username")=session("UserName")
	  rs("logip")=ip
	  rs.update
	  rs.close

	  conn.close
	  set conn = Nothing
	  response.Redirect("../index.asp")
	else
     response.write"<SCRIPT language=JavaScript>alert('密码输入错误');"
     response.write"location.href='../pro/action.login.asp'</SCRIPT>"  '验证码错误返回登入页面
	end if  
	rs.close
    set rs=nothing
    conn.close
    set conn=nothing
%>