<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>登入页面</title>
<!-- 新 Bootstrap 核心 CSS 文件 -->
<link rel="stylesheet" href="../plugin/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="../plugin/bootstrap/css/bootstrap-theme.min.css">
<script src="../inc/jquery-1.11.3.min.js"></script>
<script src="../plugin/bootstrap/js/bootstrap.min.js"></script>
<!-- 全局CSS引用 -->
<link rel="stylesheet" href="../inc/sys.main.css" />
</head>
<body>
<div id="LoginPage" class="col-lg-6" >
<form id="form1" method="post" action="action.check.asp">
<div class="form-group">
<label for="UserName">输入用户名</label>
<input type="text" class="form-control" id="UserName" placeholder="UserName" name="UserName">
</div>
<div class="form-group">
<label for="exampleInputPassword1">Password</label>
<input name="password" type="password" class="form-control" id="exampleInputPassword1" placeholder="Password">
</div>
<button type="submit" class="btn btn-default">登入</button>
</form>
</div>
</body>
</html>