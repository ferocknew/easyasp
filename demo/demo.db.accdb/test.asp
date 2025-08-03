<!--#include file="../../easyasp/easp.asp" -->
<%
' 测试数据库连接
Response.Write "<h2>EasyASP 数据库连接测试</h2>"

' 配置数据库连接
Easp.Db.SetConn "ACCESS", "../db/demo.accdb", ""

' 测试连接
On Error Resume Next
Dim conn
Set conn = Easp.Db.GetConn()
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>❌ 数据库连接失败: " & Err.Description & "</p>"
    Response.Write "<p>错误代码: " & Err.Number & "</p>"
Else
    Response.Write "<p style='color:green'>✅ 数据库连接成功</p>"

    ' 测试获取数据库类型
    Dim dbType
    dbType = Easp.Db.Type()
    Response.Write "<p>数据库类型: " & dbType & "</p>"

    ' 测试获取数据库版本
    Dim dbVersion
    dbVersion = Easp.Db.Version()
    Response.Write "<p>数据库版本: " & dbVersion & "</p>"

    ' 测试获取表列表
    Response.Write "<h3>数据库表列表:</h3>"
    Dim rs
    Set rs = Easp.Db.GetRS("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
    If Easp.Has(rs) Then
        Response.Write "<ul>"
        While Not rs.Eof
            Response.Write "<li>" & rs("Name") & "</li>"
            rs.MoveNext
        Wend
        Response.Write "</ul>"
    Else
        Response.Write "<p>没有找到用户表</p>"
    End If
    Easp.Db.Close(rs)
End If
On Error GoTo 0
%>