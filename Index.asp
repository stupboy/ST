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
<script>
var llq=document.documentElement.clientHeight;
var MF=document.getElementByID("MainF");
MF.height=llq;
</script>
<!-- 全局CSS引用 -->
<link rel="stylesheet" href="inc/sys.main.css" />
<!-- 权限文件引用 -->
<!--#include file="inc/sys.right.asp" -->
<!--#include file="lib/lib.all.asp" -->

</head>
<body>
<div class="panel panel-default">
<nav class="navbar navbar-default">

  <div class="container-fluid">
    <!-- Brand and toggle get grouped for better mobile display -->
    <div class="navbar-header">
      <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1" aria-expanded="false">
        <span class="sr-only">Toggle navigation</span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
        <span class="icon-bar"></span>
      </button>
      <a class="navbar-brand" href="#"><%=ComName%></a>
    </div>
    <!-- Collect the nav links, forms, and other content for toggling -->
    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
      <ul class="nav navbar-nav">
        <!--<li class="active"><a href="#">Link <span class="sr-only">(current)</span></a></li>-->
        <li><a href="index.asp">刷新</a></li>
        <li class="dropdown">
          <a href="#" class="dropdown-toggle" data-toggle="dropdown" role="button" aria-haspopup="true" aria-expanded="false">销售管理<span class="caret"></span></a>
          <ul class="dropdown-menu">
            <li><a href="http://www.baidu.com" target="MainF">百度</a></li>
            <li><a href="#">Another action</a></li>
            <li><a href="#">Something else here</a></li>
            <li role="separator" class="divider"></li>
            <li><a href="#">Separated link</a></li>
            <li role="separator" class="divider"></li>
            <li><a href="#">One more separated link</a></li>
          </ul>
        </li>
      </ul>
	  <%
	  '权限管理 S1 为生产管理
	  if qx("s1") then 
	  sc caidan("生产管理","新浪$http://www.sina.com.cn")
	  end if 
	  sc caidan("仓库管理","生产$main.asp|百度$http://www.baidu.com")
	  %>
      <form class="navbar-form navbar-left" role="search">
        <div class="form-group">
          <input type="text" class="form-control" placeholder="Search">
        </div>
        <button type="submit" class="btn btn-default">搜索</button>
      </form>
      <ul class="nav navbar-nav navbar-right">
        <li><a href="#">管理</a></li>
        <li class="dropdown">
          <a href="#" class="dropdown-toggle" data-toggle="dropdown" role="button" aria-haspopup="true" aria-expanded="false">登入管理 <span class="caret"></span></a>
          <ul class="dropdown-menu">
            <li><a href="pro/action.logout.asp" target="_parent">注销</a></li>
			<li role="separator" class="divider"></li>
            <li><a href="#">修改密码</a></li>
          </ul>
        </li>
      </ul>
    </div><!-- /.navbar-collapse -->
  </div><!-- /.container-fluid -->
</nav>
<div class="panel-body"><iframe id="MainF" name="MainF" src="Main.asp" width="100%" frameborder="0" height="620px"  scrolling="auto">xxx</iframe></div>
<div class="panel-footer panel-right"><p class="text-center">Center aligned text.</p></div>
</div>
</body>
</html>