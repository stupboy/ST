<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>登入页面</title>
<!-- 新 Bootstrap 核心 CSS 文件 -->
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap-theme.min.css">
<script src="inc/jquery-1.11.3.min.js"></script>
<script src="plugin/bootstrap/js/bootstrap.min.js"></script>
<!-- 全局CSS引用 -->
<link rel="stylesheet" href="inc/sys.main.css" />
<!--#include file="lib/lib.all.asp" -->
<!--#include file="inc/sql.conn.asp" -->
</head>
<body>
<div class="panel panel-default">
  <!-- Default panel contents -->
  <div class="panel-heading">商品基础资料</div>
  <div class="panel-body">
<%
	sql="select * from sys_goods"                    '查询商品基础资料库【范例】
	set rs=Server.CreateObject("ADODB.RecordSet")
	rs.open sql,conn,1,1
	if rs.EOF then 
	sc "NO answer!"
	else 
	 sc "<table class='table'>"
	 sc "<tr>"
	 for i=0 to rs.Fields.Count-1
	 sc "<td>"&rs.Fields(i).Name&"</td>"
	 next
	 sc "</tr>"
	 while not rs.EOF
	  sc "<tr>"
	  for each x in rs.Fields
	  sc "<td>"&x.Value&"</td>"
	  next
	  sc "</tr>"
	  rs.MoveNext
	 wend 
	 rs.Close
	 set rs=nothing 
	end if 
%>
  </div>
  <!-- List group -->
  <ul class="list-group">
	
</div>
</body>
</html>