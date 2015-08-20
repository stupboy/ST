<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="inc/sys.right.asp" -->
<title>登入页面</title>
<!-- 新 Bootstrap 核心 CSS 文件 -->
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap-theme.min.css">
<script src="inc/jquery-1.11.3.min.js"></script>
<script src="plugin/bootstrap/js/bootstrap.min.js"></script>
<!-- 全局CSS引用 -->
<link rel="stylesheet" href="inc/sys.main.css" />
<!-- 引用自定义【函数库】及【数据库连接】 -->
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
	set rs=Server.CreateObject("ADODB.RecordSet")    '创建连接对象
	rs.open sql,conn,1,1                             '打开连接字符
	if rs.EOF then                                   '如果记录集为空
	sc "NO answer!"                                  '则输出NO answer
	else                                             '如果不为空则输出记录
	 sc "<table class='table'>"                      '输出表格头利用bootstrap类
	 sc "<tr>"                                       '输出换行标题
	 for i=0 to rs.Fields.Count-1                    '表标题循环输出  rs.Fields.count 为标题个数 0为开始数量 FOR循环
	 sc "<td>"&rs.Fields(i).Name&"</td>"             '循环输出表的字段名 .Name为字段名 .value为值
	 next                                            '循环结束
	 sc "</tr>"                                      '输出表格换行</tr>
	 while not rs.EOF                                '循环输出值 while循环
	  sc "<tr>"                                      '表格换行开始
	  for each x in rs.Fields                        '循环开始 对应记录集rs.Fields中的每个对象x
	  sc "<td>"&x.Value&"</td>"                      '输出值
	  next                                           '循环结束
	  sc "</tr>"                                     '表格换行结束
	  rs.MoveNext                                    '记录下移一条
	 wend                                            '循环输出结束 while循环
	 rs.Close                                        '关闭记录集
	 set rs=nothing                                  '关闭记录集
	end if                                           '判断记录集if结束
%>
  </div>
</div>
</body>
</html>