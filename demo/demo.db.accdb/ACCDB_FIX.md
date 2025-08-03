# ACCDB 连接修复说明

## 问题描述

在使用 EasyASP 连接 `.accdb` 文件时出现错误：
```
错误代码 : -2147467259
错误描述 : 不可识别的数据库格式 'C:\Users\Administrator\Documents\www\easyasp\demo\db\demo.accdb'。
错误来源 : Microsoft JET Database Engine
```

## 问题原因

EasyASP 的数据库连接逻辑中，ACCESS 类型统一使用 `Microsoft.Jet.OLEDB.4.0` 提供程序，但这个提供程序只支持 `.mdb` 文件格式，不支持 `.accdb` 文件格式。

## 解决方案

修改了 `easyasp/core/easp.db.asp` 文件中的 ACCESS 连接逻辑，添加了文件扩展名检测：

### 修改前
```vbscript
Case "ACCESS"
  Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
  ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
```

### 修改后
```vbscript
Case "ACCESS"
  Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
  ' 根据文件扩展名选择提供程序
  Dim fileExt : fileExt = LCase(Right(tDb, 4))
  If fileExt = ".accdb" Then
    ' 对于 .accdb 文件使用 ACE 提供程序
    ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
  Else
    ' 对于 .mdb 文件使用 JET 提供程序
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
  End If
```

## 提供程序说明

- **Microsoft.ACE.OLEDB.12.0**: 支持 `.accdb` 和 `.mdb` 文件
- **Microsoft.Jet.OLEDB.4.0**: 仅支持 `.mdb` 文件

## 系统要求

要使用 `.accdb` 文件，服务器需要安装 Microsoft Access Database Engine：

1. **下载地址**: Microsoft 官方网站
2. **版本要求**: 2010 或更新版本
3. **架构匹配**: 确保安装的版本与 IIS 应用程序池的位数匹配

## 测试页面

- `test_fix.asp`: 测试 ACCDB 连接修复是否有效
- `base.asp`: 提供多种连接方法的测试页面

## 兼容性

此修改保持了向后兼容性：
- `.mdb` 文件继续使用 JET 提供程序
- `.accdb` 文件使用 ACE 提供程序
- 现有代码无需修改

## 注意事项

1. 确保服务器已安装 Microsoft Access Database Engine
2. 在 64 位系统上，可能需要使用 32 位应用程序池
3. 如果仍然出现提供程序错误，请检查服务器配置