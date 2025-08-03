<%
' ========================================
' 基本数据库连接代码示例
' 文件名: base.asp
' 用途: 提供基本的 ADODB 连接示例
' ========================================

' 设置响应头
Response.CharSet = "utf-8"
Response.ContentType = "text/html"

' 定义数据库路径
Dim dbPath
dbPath = Server.MapPath("../db/demo.accdb")

' 方法1: 使用 Microsoft.ACE.OLEDB.12.0 (推荐用于 .accdb 文件)
Function ConnectWithACE()
    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    Dim connString
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    conn.Open connString

    If Err.Number <> 0 Then
        Response.Write "<p><strong>ACE 连接失败:</strong> " & Err.Description & "</p>"
        Response.Write "<p>错误代码: " & Err.Number & "</p>"
        Set conn = Nothing
        ConnectWithACE = False
    Else
        Response.Write "<p><strong>ACE 连接成功!</strong></p>"
        conn.Close
        Set conn = Nothing
        ConnectWithACE = True
    End If

    On Error Goto 0
End Function

' 方法2: 使用 Microsoft.ACE.OLEDB.14.0 (备用选项)
Function ConnectWithACE14()
    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    Dim connString
    connString = "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & dbPath & ";"

    conn.Open connString

    If Err.Number <> 0 Then
        Response.Write "<p><strong>ACE 14.0 连接失败:</strong> " & Err.Description & "</p>"
        Response.Write "<p>错误代码: " & Err.Number & "</p>"
        Set conn = Nothing
        ConnectWithACE14 = False
    Else
        Response.Write "<p><strong>ACE 14.0 连接成功!</strong></p>"
        conn.Close
        Set conn = Nothing
        ConnectWithACE14 = True
    End If

    On Error Goto 0
End Function

' 方法3: 使用 Microsoft.Jet.OLEDB.4.0 (仅适用于 .mdb 文件)
Function ConnectWithJet()
    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    Dim connString
    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";"

    conn.Open connString

    If Err.Number <> 0 Then
        Response.Write "<p><strong>Jet 连接失败:</strong> " & Err.Description & "</p>"
        Response.Write "<p>错误代码: " & Err.Number & "</p>"
        Set conn = Nothing
        ConnectWithJet = False
    Else
        Response.Write "<p><strong>Jet 连接成功!</strong></p>"
        conn.Close
        Set conn = Nothing
        ConnectWithJet = True
    End If

    On Error Goto 0
End Function

' 方法4: 使用 ODBC 连接
Function ConnectWithODBC()
    On Error Resume Next

    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    Dim connString
    connString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & dbPath & ";"

    conn.Open connString

    If Err.Number <> 0 Then
        Response.Write "<p><strong>ODBC 连接失败:</strong> " & Err.Description & "</p>"
        Response.Write "<p>错误代码: " & Err.Number & "</p>"
        Set conn = Nothing
        ConnectWithODBC = False
    Else
        Response.Write "<p><strong>ODBC 连接成功!</strong></p>"
        conn.Close
        Set conn = Nothing
        ConnectWithODBC = True
    End If

    On Error Goto 0
End Function

' 检查数据库文件是否存在
Function CheckDatabaseFile()
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(dbPath) Then
        Response.Write "<p><strong>数据库文件存在:</strong> " & dbPath & "</p>"
        CheckDatabaseFile = True
    Else
        Response.Write "<p><strong>数据库文件不存在:</strong> " & dbPath & "</p>"
        CheckDatabaseFile = False
    End If

    Set fso = Nothing
End Function

' 显示服务器信息
Function ShowServerInfo()
    Response.Write "<h3>服务器信息</h3>"
    Response.Write "<p><strong>服务器软件:</strong> " & Request.ServerVariables("SERVER_SOFTWARE") & "</p>"
    Response.Write "<p><strong>操作系统:</strong> " & Request.ServerVariables("OS") & "</p>"
    Response.Write "<p><strong>当前时间:</strong> " & Now() & "</p>"
End Function

' 主程序
Response.Write "<html><head><title>数据库连接测试</title>"
Response.Write "<meta charset='utf-8'>"
Response.Write "<style>"
Response.Write "body { font-family: Arial, sans-serif; margin: 20px; }"
Response.Write ".success { color: green; }"
Response.Write ".error { color: red; }"
Response.Write ".info { color: blue; }"
Response.Write "</style>"
Response.Write "</head><body>"

Response.Write "<h1>数据库连接测试</h1>"
Response.Write "<p class='info'>此页面用于测试不同的数据库连接方法</p>"

' 显示服务器信息
ShowServerInfo()

' 检查数据库文件
Response.Write "<h3>数据库文件检查</h3>"
CheckDatabaseFile()

' 测试不同的连接方法
Response.Write "<h3>连接测试</h3>"

Response.Write "<h4>1. 测试 Microsoft.ACE.OLEDB.12.0</h4>"
ConnectWithACE()

Response.Write "<h4>2. 测试 Microsoft.ACE.OLEDB.14.0</h4>"
ConnectWithACE14()

Response.Write "<h4>3. 测试 Microsoft.Jet.OLEDB.4.0</h4>"
ConnectWithJet()

Response.Write "<h4>4. 测试 ODBC 连接</h4>"
ConnectWithODBC()

Response.Write "<h3>故障排除建议</h3>"
Response.Write "<ul>"
Response.Write "<li>如果所有连接都失败，请确保服务器已安装 Microsoft Access Database Engine</li>"
Response.Write "<li>对于 .accdb 文件，推荐使用 Microsoft.ACE.OLEDB.12.0 或 14.0</li>"
Response.Write "<li>对于 .mdb 文件，可以使用 Microsoft.Jet.OLEDB.4.0</li>"
Response.Write "<li>确保数据库文件路径正确且有读取权限</li>"
Response.Write "<li>在 64 位系统上，可能需要安装 32 位版本的 Access Database Engine</li>"
Response.Write "</ul>"

Response.Write "</body></html>"
%>
