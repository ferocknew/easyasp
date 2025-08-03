<%
' ========================================
' ACCDB 连接修复测试页面
' 文件名: test_fix.asp
' 用途: 测试 EasyASP 对 .accdb 文件的支持
' ========================================

Response.CharSet = "utf-8"
Response.ContentType = "text/html"

' 引入 EasyASP
%><!--#include file="../../easyasp/easp.asp" --><%

' 测试连接
On Error Resume Next

' 设置数据库连接
Easp.Db.SetConn "ACCESS", "../db/demo.accdb", ""

' 尝试获取连接
Dim conn
Set conn = Easp.Db.GetConn()

If Err.Number <> 0 Then
    Response.Write "<h2>连接失败</h2>"
    Response.Write "<p><strong>错误代码:</strong> " & Err.Number & "</p>"
    Response.Write "<p><strong>错误描述:</strong> " & Err.Description & "</p>"
    Response.Write "<p><strong>错误来源:</strong> " & Err.Source & "</p>"
Else
    Response.Write "<h2>连接成功!</h2>"
    Response.Write "<p><strong>数据库类型:</strong> " & Easp.Db.GetType() & "</p>"
    Response.Write "<p><strong>数据库版本:</strong> " & Easp.Db.GetVersion() & "</p>"

    ' 尝试获取表列表
    Dim rs
    Set rs = Easp.Db.GetRS("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")

    If Easp.Has(rs) Then
        Response.Write "<h3>数据库表列表:</h3>"
        Response.Write "<ul>"
        While Not rs.Eof
            Response.Write "<li>" & rs("Name") & "</li>"
            rs.MoveNext
        Wend
        Response.Write "</ul>"
        Easp.Db.Close(rs)
    Else
        Response.Write "<p>无法获取表列表</p>"
    End If

    ' 关闭连接
    Easp.Db.CloseConn()
End If

On Error Goto 0
%>

<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8">
    <title>ACCDB 连接修复测试</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .success { color: green; }
        .error { color: red; }
        .info { color: blue; }
    </style>
</head>
<body>
    <h1>ACCDB 连接修复测试</h1>
    <p class="info">此页面用于测试 EasyASP 对 .accdb 文件的支持</p>

    <h3>测试结果:</h3>
    <div id="result">
        <!-- 结果将在这里显示 -->
    </div>

    <h3>修复说明:</h3>
    <ul>
        <li>修改了 <code>easyasp/core/easp.db.asp</code> 文件</li>
        <li>添加了文件扩展名检测逻辑</li>
        <li>对于 .accdb 文件使用 Microsoft.ACE.OLEDB.12.0 提供程序</li>
        <li>对于 .mdb 文件继续使用 Microsoft.Jet.OLEDB.4.0 提供程序</li>
    </ul>

    <p><a href="index.asp">返回主页面</a></p>
</body>
</html>