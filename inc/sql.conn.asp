
<%
'数据库连接文件
set conn=server.CreateObject("adodb.connection")
'“.”为服务器地址、ST为连接数据库名称、sa为数据库用户名、PWD为数据库密码
ConnStr="server=.;driver={sql server};database=ST;uid=sa;pwd=!@#$%asdfg"
conn.Open connstr
'如果连接出错则报错
If Err Then
  err.Clear
  Set Conn = Nothing
  Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
  response.Write connstr
  Response.End
End If
'自定义函数关闭数据库连接
sub CloseConn()
 On Error Resume Next
 If IsObject(Conn) Then
  conn.close
  set conn=nothing
 end if
end Sub
%>