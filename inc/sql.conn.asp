
<%
'on error resume next
set conn=server.CreateObject("adodb.connection")
ConnStr="server=.;driver={sql server};database=ST;uid=sa;pwd=!@#$%asdfg"
conn.Open connstr
 
If Err Then
  err.Clear
  Set Conn = Nothing
  Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
  response.Write connstr
  Response.End
 End If
sub CloseConn()
 On Error Resume Next
 If IsObject(Conn) Then
  conn.close
  set conn=nothing
 end if
end Sub
 
%>