<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="inc/sys.right.asp" -->
<title><%=webtitle%></title>
<!-- 新 Bootstrap 核心 CSS 文件 -->
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="plugin/bootstrap/css/bootstrap-theme.min.css">
<script src="inc/jquery-1.11.3.min.js"></script>
<script src="plugin/bootstrap/js/bootstrap.min.js"></script>
</head>
<frameset rows="110,*,60" cols="*" frameborder="no" border="0" framespacing="0">
  <frame src="top.asp" name="topFrame" scrolling="No" noresize="noresize" id="topFrame" title="topFrame" />
  <frameset cols="180,*" frameborder="no" border="1" framespacing="0">
    <frame src="left.asp" name="leftFrame" scrolling="No" noresize="noresize" id="leftFrame" title="leftFrame" />
    <frame src="main.asp" name="mainFrame" id="mainFrame" title="mainFrame" />
  </frameset>
  <frame src="bottom.asp" name="botFrame" scrolling="No" noresize="noresize" id="botFrame" title="botFrame" />
</frameset>
<noframes><body>
</body></noframes>
</html>
