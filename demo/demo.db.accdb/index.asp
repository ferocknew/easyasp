<!--#include file="../../easyasp/easp.asp" -->
<%
' 配置数据库连接
Easp.Db.SetConn "ACCESS", "../db/demo.accdb", ""

' 获取所有表名
Function GetTableList()
    Dim rs, tables
    Set tables = CreateObject("Scripting.Dictionary")

    ' 获取所有表名
    Set rs = Easp.Db.GetRS("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
    If Easp.Has(rs) Then
        While Not rs.Eof
            tables.Add rs("Name"), rs("Name")
            rs.MoveNext
        Wend
    End If
    Easp.Db.Close(rs)

    Set GetTableList = tables
End Function

' 获取表结构
Function GetTableStructure(tableName)
    Dim rs, fields
    Set fields = CreateObject("Scripting.Dictionary")

    ' 获取表结构
    Set rs = Easp.Db.GetRS("SELECT * FROM [" & tableName & "] WHERE 1=0")
    If Easp.Has(rs) Then
        For Each field In rs.Fields
            fields.Add field.Name, Array(field.Type, field.DefinedSize, field.Attributes)
        Next
    End If
    Easp.Db.Close(rs)

    Set GetTableStructure = fields
End Function

' 获取表数据
Function GetTableData(tableName, limit)
    Dim rs, data
    Set data = CreateObject("Scripting.Dictionary")

    ' 获取表数据
    Set rs = Easp.Db.GetRS("SELECT TOP " & limit & " * FROM [" & tableName & "]")
    If Easp.Has(rs) Then
        Set data("rs") = rs
        data("hasData") = True
    Else
        data("hasData") = False
    End If

    Set GetTableData = data
End Function

' 处理页面逻辑
Dim selectedTable, tableList, tableStructure, tableData
selectedTable = Easp.Get("table")

If selectedTable <> "" Then
    Set tableStructure = GetTableStructure(selectedTable)
    Set tableData = GetTableData(selectedTable, 10)
End If

Set tableList = GetTableList()
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>数据库表结构查看器 - EasyASP Demo</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <link href="style.css" rel="stylesheet">
</head>
<body>
    <!--#include file="_nav.asp" -->
    <div class="container-fluid">
        <div class="row">
            <!-- 侧边栏 -->
            <div class="col-md-3 col-lg-2 d-md-block bg-light sidebar">
                <div class="position-sticky pt-3">
                    <h6 class="sidebar-heading d-flex justify-content-between align-items-center px-3 mb-3 text-muted">
                        <span>数据库表列表</span>
                        <span class="badge bg-primary rounded-pill"><%=tableList.Count%></span>
                    </h6>
                    <ul class="nav flex-column">
                        <%
                        For Each tableName In tableList.Keys
                            Dim isSelected
                            isSelected = (LCase(tableName) = LCase(selectedTable))
                        %>
                        <li class="nav-item">
                            <a class="nav-link table-name <%=IIf(isSelected, "selected-table", "")%>"
                               href="?table=<%=Server.URLEncode(tableName)%>">
                                <i class="bi bi-table"></i>
                                <%=tableName%>
                            </a>
                        </li>
                        <% Next %>
                    </ul>
                </div>
            </div>

            <!-- 主内容区 -->
            <div class="col-md-9 col-lg-10 ms-sm-auto px-md-4">
                <div class="d-flex justify-content-between flex-wrap flex-md-nowrap align-items-center pt-3 pb-2 mb-3 border-bottom">
                    <h1 class="h2">数据库表结构查看器</h1>
                    <div class="btn-toolbar mb-2 mb-md-0">
                        <div class="btn-group me-2">
                            <button type="button" class="btn btn-sm btn-outline-secondary" onclick="location.reload()">
                                <i class="bi bi-arrow-clockwise"></i> 刷新
                            </button>
                        </div>
                    </div>
                </div>

                <% If selectedTable <> "" Then %>
                <!-- 表结构信息 -->
                <div class="row">
                    <div class="col-12">
                        <div class="card mb-4">
                            <div class="card-header">
                                <h5 class="card-title mb-0">
                                    <i class="bi bi-info-circle"></i>
                                    表结构: <%=selectedTable%>
                                </h5>
                            </div>
                            <div class="card-body">
                                <div class="table-responsive">
                                    <table class="table table-striped table-sm">
                                        <thead>
                                            <tr>
                                                <th>字段名</th>
                                                <th>数据类型</th>
                                                <th>长度</th>
                                                <th>属性</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <%
                                            If Not tableStructure Is Nothing Then
                                                For Each fieldName In tableStructure.Keys
                                                    Dim fieldInfo
                                                    fieldInfo = tableStructure(fieldName)
                                            %>
                                            <tr>
                                                <td><strong><%=fieldName%></strong></td>
                                                <td><span class="field-type"><%=GetDataTypeName(fieldInfo(0))%></span></td>
                                                <td><%=fieldInfo(1)%></td>
                                                <td><%=GetFieldAttributes(fieldInfo(2))%></td>
                                            </tr>
                                            <%
                                                Next
                                            End If
                                            %>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 表数据 -->
                <div class="row">
                    <div class="col-12">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="card-title mb-0">
                                    <i class="bi bi-list-ul"></i>
                                    表数据 (前10条)
                                </h5>
                            </div>
                            <div class="card-body">
                                <% If tableData("hasData") Then %>
                                <div class="table-responsive table-container">
                                    <table class="table table-striped table-hover">
                                        <thead class="table-dark sticky-top">
                                            <tr>
                                                <%
                                                Dim rs
                                                Set rs = tableData("rs")
                                                For Each field In rs.Fields
                                                %>
                                                <th><%=field.Name%></th>
                                                <% Next %>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <%
                                            While Not rs.Eof
                                            %>
                                            <tr>
                                                <%
                                                For Each field In rs.Fields
                                                    Dim fieldValue
                                                    fieldValue = rs(field.Name)
                                                    If IsNull(fieldValue) Then
                                                        fieldValue = "<span class='text-muted'>NULL</span>"
                                                    ElseIf IsDate(fieldValue) Then
                                                        fieldValue = FormatDateTime(fieldValue, 2)
                                                    ElseIf IsNumeric(fieldValue) Then
                                                        fieldValue = fieldValue
                                                    Else
                                                        fieldValue = Server.HTMLEncode(fieldValue)
                                                    End If
                                                %>
                                                <td><%=fieldValue%></td>
                                                <% Next %>
                                            </tr>
                                            <%
                                                rs.MoveNext
                                            Wend
                                            %>
                                        </tbody>
                                    </table>
                                </div>
                                <% Else %>
                                <div class="alert alert-info">
                                    <i class="bi bi-info-circle"></i>
                                    该表没有数据或无法访问
                                </div>
                                <% End If %>
                            </div>
                        </div>
                    </div>
                </div>
                <% Else %>
                <!-- 欢迎页面 -->
                <div class="row">
                    <div class="col-12">
                        <div class="card">
                            <div class="card-body text-center">
                                <i class="bi bi-database" style="font-size: 4rem; color: #6c757d;"></i>
                                <h3 class="mt-3">欢迎使用数据库表结构查看器</h3>
                                <p class="text-muted">请从左侧选择一个表来查看其结构和数据</p>
                                <div class="mt-4">
                                    <div class="row">
                                        <div class="col-md-4">
                                            <div class="card bg-primary text-white">
                                                <div class="card-body">
                                                    <h5><i class="bi bi-table"></i> 表数量</h5>
                                                    <h2><%=tableList.Count%></h2>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-md-4">
                                            <div class="card bg-success text-white">
                                                <div class="card-body">
                                                    <h5><i class="bi bi-check-circle"></i> 连接状态</h5>
                                                    <h2>正常</h2>
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-md-4">
                                            <div class="card bg-info text-white">
                                                <div class="card-body">
                                                    <h5><i class="bi bi-gear"></i> 数据库</h5>
                                                    <h2>demo.accdb</h2>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <% End If %>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // 添加一些交互效果
        document.addEventListener('DOMContentLoaded', function() {
            // 表格行悬停效果
            const tableRows = document.querySelectorAll('tbody tr');
            tableRows.forEach(row => {
                row.addEventListener('mouseenter', function() {
                    this.style.backgroundColor = '#f8f9fa';
                });
                row.addEventListener('mouseleave', function() {
                    this.style.backgroundColor = '';
                });
            });
        });
    </script>
</body>
</html>

<%
' 辅助函数
Function GetDataTypeName(dataType)
    Select Case dataType
        Case 2: GetDataTypeName = "SmallInt"
        Case 3: GetDataTypeName = "Integer"
        Case 4: GetDataTypeName = "Single"
        Case 5: GetDataTypeName = "Double"
        Case 6: GetDataTypeName = "Currency"
        Case 7: GetDataTypeName = "Date"
        Case 8: GetDataTypeName = "BSTR"
        Case 11: GetDataTypeName = "Boolean"
        Case 12: GetDataTypeName = "Variant"
        Case 14: GetDataTypeName = "Decimal"
        Case 16: GetDataTypeName = "TinyInt"
        Case 17: GetDataTypeName = "UnsignedTinyInt"
        Case 18: GetDataTypeName = "UnsignedSmallInt"
        Case 19: GetDataTypeName = "UnsignedInt"
        Case 20: GetDataTypeName = "BigInt"
        Case 21: GetDataTypeName = "UnsignedBigInt"
        Case 72: GetDataTypeName = "GUID"
        Case 128: GetDataTypeName = "Binary"
        Case 129: GetDataTypeName = "Char"
        Case 130: GetDataTypeName = "WChar"
        Case 131: GetDataTypeName = "Numeric"
        Case 132: GetDataTypeName = "UserDefined"
        Case 133: GetDataTypeName = "DBDate"
        Case 134: GetDataTypeName = "DBTime"
        Case 135: GetDataTypeName = "DBTimeStamp"
        Case 200: GetDataTypeName = "VarChar"
        Case 201: GetDataTypeName = "LongVarChar"
        Case 202: GetDataTypeName = "VarWChar"
        Case 203: GetDataTypeName = "LongVarWChar"
        Case 204: GetDataTypeName = "VarBinary"
        Case 205: GetDataTypeName = "LongVarBinary"
        Case Else: GetDataTypeName = "Unknown(" & dataType & ")"
    End Select
End Function

Function GetFieldAttributes(attributes)
    Dim attrList
    attrList = ""

    If (attributes And 1) <> 0 Then attrList = attrList & "Fixed, "
    If (attributes And 2) <> 0 Then attrList = attrList & "Nullable, "
    If (attributes And 4) <> 0 Then attrList = attrList & "Long, "
    If (attributes And 8) <> 0 Then attrList = attrList & "Signed, "
    If (attributes And 16) <> 0 Then attrList = attrList & "AutoIncrement, "
    If (attributes And 32) <> 0 Then attrList = attrList & "RowVersion, "
    If (attributes And 64) <> 0 Then attrList = attrList & "RowID, "
    If (attributes And 128) <> 0 Then attrList = attrList & "Chapter, "
    If (attributes And 256) <> 0 Then attrList = attrList & "IsRowID, "
    If (attributes And 512) <> 0 Then attrList = attrList & "Deferred, "
    If (attributes And 1024) <> 0 Then attrList = attrList & "CacheDeferred, "

    If attrList <> "" Then
        GetFieldAttributes = Left(attrList, Len(attrList) - 2)
    Else
        GetFieldAttributes = "Normal"
    End If
End Function
%>
