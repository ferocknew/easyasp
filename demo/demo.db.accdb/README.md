# 数据库表结构查看器

这是一个基于 EasyASP 框架开发的数据库表结构查看器，用于连接和展示 Access 数据库中的表结构和数据。

## 功能特性

### 🎯 核心功能
- **数据库连接**: 自动连接到 `../db/demo.accdb` 数据库
- **表列表展示**: 显示数据库中所有表的列表
- **表结构查看**: 详细展示选中表的字段信息，包括：
  - 字段名称
  - 数据类型
  - 字段长度
  - 字段属性
- **数据预览**: 显示表中前 10 条数据记录
- **响应式设计**: 支持桌面和移动设备

### 🎨 界面特性
- **现代化 UI**: 使用 Bootstrap 5 和 Bootstrap Icons
- **侧边栏导航**: 左侧显示表列表，右侧显示详细信息
- **数据表格**: 支持滚动查看大量数据
- **状态指示**: 显示数据库连接状态
- **交互效果**: 悬停高亮、选中状态等

## 文件结构

```
demo/demo.db.accdb/
├── index.asp          # 主页面
├── _nav.asp           # 导航栏
├── style.css          # 样式文件
└── README.md          # 说明文档
```

## 技术实现

### 数据库连接
```vbscript
Easp.Db.SetConn "ACCESS", "../db/demo.accdb", ""
```

### 获取表列表
```vbscript
Set rs = Easp.Db.GetRS("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
```

### 获取表结构
```vbscript
Set rs = Easp.Db.GetRS("SELECT * FROM [" & tableName & "] WHERE 1=0")
```

### 获取表数据
```vbscript
Set rs = Easp.Db.GetRS("SELECT TOP " & limit & " * FROM [" & tableName & "]")
```

## 数据类型映射

系统支持显示以下 Access 数据类型：

| 类型代码 | 类型名称 | 说明 |
|---------|---------|------|
| 2 | SmallInt | 小整数 |
| 3 | Integer | 整数 |
| 4 | Single | 单精度浮点数 |
| 5 | Double | 双精度浮点数 |
| 6 | Currency | 货币类型 |
| 7 | Date | 日期时间 |
| 11 | Boolean | 布尔值 |
| 12 | Variant | 变体类型 |
| 129 | Char | 字符 |
| 200 | VarChar | 可变字符 |
| 201 | LongVarChar | 长可变字符 |
| 204 | VarBinary | 可变二进制 |

## 字段属性说明

- **Fixed**: 固定长度字段
- **Nullable**: 可为空字段
- **Long**: 长字段
- **Signed**: 有符号数值
- **AutoIncrement**: 自动递增
- **RowVersion**: 行版本
- **RowID**: 行标识符

## 使用方法

1. **访问页面**: 打开 `index.asp` 文件
2. **选择表**: 从左侧表列表中选择要查看的表
3. **查看结构**: 在右侧查看表的字段结构信息
4. **预览数据**: 查看表中前 10 条数据记录
5. **刷新页面**: 点击刷新按钮重新加载数据

## 依赖项

- **EasyASP 框架**: 提供数据库操作功能
- **Bootstrap 5**: UI 框架
- **Bootstrap Icons**: 图标库
- **Access 数据库**: 目标数据库文件

## 浏览器兼容性

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## 注意事项

1. 确保 `../db/demo.accdb` 数据库文件存在且可访问
2. 服务器需要支持 ASP 和 Access 数据库驱动
3. 某些系统表可能无法访问，这是正常现象
4. 大数据表的数据预览可能较慢，请耐心等待

## 扩展功能

可以考虑添加以下功能：

- [ ] 数据导出功能
- [ ] 表结构对比
- [ ] SQL 查询执行器
- [ ] 数据统计图表
- [ ] 表关系图
- [ ] 数据编辑功能

## 故障排除

### 常见问题

1. **数据库连接失败**
   - 检查数据库文件路径是否正确
   - 确认服务器支持 Access 数据库

2. **表列表为空**
   - 检查数据库是否包含用户表
   - 确认数据库权限设置

3. **页面显示异常**
   - 检查 Bootstrap 和图标库是否正确加载
   - 确认浏览器兼容性

### 调试信息

页面包含详细的错误处理和日志记录，可以通过浏览器开发者工具查看控制台信息。