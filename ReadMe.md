# VBA JsonParser

一个强大的VBA类模块，用于在Microsoft Office应用程序中解析和处理JSON数据。

## 特性

- **高性能**: 使用HTML Document Object Model (DOM) 和JavaScript引擎进行JSON解析
- **功能完整**: 支持JSON的解析、获取、设置、删除等操作
- **嵌套支持**: 支持深度嵌套的JSON结构，使用点号路径访问
- **类型检测**: 自动识别JSON数据类型（对象、数组、基本类型）
- **易用性**: 简洁的API设计，易于集成到现有VBA项目中

## 系统要求

- Microsoft Office 2007或更高版本
- Windows操作系统
- Internet Explorer组件（用于HTML文档对象）

## 安装

1. 在VBA编辑器中，右键点击项目
2. 选择"插入" → "类模块"
3. 将`JsonParser.cls`文件的内容复制到新建的类模块中
4. 将类模块命名为`JsonParser`

## 快速开始

```vba
Sub Example()
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    ' 设置JSON字符串
    jp.JsonString = "{""name"":""张三"",""age"":30,""city"":""北京""}"
    
    ' 获取值
    Debug.Print jp.Item("name")  ' 输出: 张三
    Debug.Print jp.Item("age")   ' 输出: 30
    
    ' 添加新键值对
    jp.Add "email", "zhangsan@example.com"
    
    ' 检查键是否存在
    Debug.Print jp.ContainsKey("email")  ' 输出: True
    
    ' 删除键
    jp.Delete "city"
    
    ' 格式化输出JSON
    Debug.Print jp.Stringify(2)
End Sub
```

## API参考

### 属性

#### JsonString
设置或获取JSON字符串。

```vba
' 设置
jp.JsonString = "{""key"":""value""}"

' 获取
Dim json As String
json = jp.JsonString
```

#### JsonType
获取JSON对象的类型。

```vba
' 返回数字格式 (0=数组, 1=对象, 2=其他)
Dim typeNum As Integer
typeNum = jp.JsonType(0)

' 返回符号格式 ([]/{}/Other)
Dim typeSym As String
typeSym = jp.JsonType(1)

' 返回文字格式 (Array/Object/Other)
Dim typeText As String
typeText = jp.JsonType(2)
```

#### Count
获取JSON对象或数组的元素数量。

```vba
Dim count As Integer
count = jp.Count
```

#### Keys
获取顶层键名（逗号分隔）。

```vba
Dim keys As String
keys = jp.Keys  ' 例如: "name,age,city"
```

#### ContainsKey
检查JSON中是否包含指定键。

```vba
Dim exists As Boolean
exists = jp.ContainsKey("name")
```

#### FindKeyPaths
查找指定键的所有路径。

```vba
Dim paths As Variant
paths = jp.FindKeyPaths("name")  ' 返回所有包含"name"键的路径
```

#### KeyValuePairs
获取所有键值对（逗号分隔）。

```vba
Dim pairs As Variant
pairs = jp.KeyValuePairs
```

### 方法

#### Item
根据键名或路径获取值。

```vba
' 简单键
Dim name As String
name = jp.Item("name")

' 嵌套路径
Dim street As String
street = jp.Item("address.street")

' 数组元素
Dim hobby As String
hobby = jp.Item("hobbies.0")
```

#### Add
添加键值对（支持嵌套路径）。

```vba
' 添加简单键值对
jp.Add "email", "user@example.com"

' 添加嵌套键值对
jp.Add "address.zipcode", "100000"

' 添加数组元素
jp.Add "hobbies.3", "旅游"
```

#### Delete
删除指定键。

```vba
' 删除简单键
jp.Delete "email"

' 删除嵌套键
jp.Delete "address.zipcode"
```

#### GetAllKeys
获取所有唯一的键名（包括嵌套）。

```vba
Dim allKeys As Variant
allKeys = jp.GetAllKeys
```

#### Stringify
将JSON对象格式化为字符串。

```vba
' 带缩进的格式化
Dim formattedJson As String
formattedJson = jp.Stringify(2)  ' 2个空格缩进

' 无缩进
formattedJson = jp.Stringify(0)
```

## 使用示例

### 处理复杂嵌套JSON

```vba
Sub ComplexJsonExample()
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    Dim complexJson As String
    complexJson = "{""users"":[{""id"":1,""profile"":{""name"":""用户1"",""settings"":{""theme"":""dark""}}}],""config"":{""version"":""1.0""}}"
    
    jp.JsonString = complexJson
    
    ' 访问嵌套数据
    Debug.Print jp.Item("users.0.profile.name")      ' "用户1"
    Debug.Print jp.Item("users.0.profile.settings.theme")  ' "dark"
    Debug.Print jp.Item("config.version")            ' "1.0"
    
    ' 添加新用户
    jp.Add "users.1", "{""id"":2,""profile"":{""name"":""用户2"",""settings"":{""theme"":""light""}}}"
    
    ' 输出格式化JSON
    Debug.Print jp.Stringify(2)
End Sub
```

### 处理数组JSON

```vba
Sub ArrayJsonExample()
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    Dim arrayJson As String
    arrayJson = "[{""name"":""商品A"",""price"":100},{""name"":""商品B"",""price"":200}]"
    
    jp.JsonString = arrayJson
    
    Debug.Print "数组类型: " & jp.JsonType(2)        ' "Array"
    Debug.Print "数组长度: " & jp.Count              ' 2
    Debug.Print "第一个商品: " & jp.Item("0.name")   ' "商品A"
    Debug.Print "第二个价格: " & jp.Item("1.price")  ' 200
End Sub
```

## 错误处理

当设置无效的JSON字符串时，类会抛出错误：

```vba
On Error GoTo ErrorHandler
jp.JsonString = "{invalid json string"
Exit Sub

ErrorHandler:
Debug.Print "错误: " & Err.Description
```

## 限制

- 仅支持Windows平台
- 依赖Internet Explorer组件，可能在某些系统配置下受限
- JSON字符串长度受VBA字符串限制

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。

## 许可证

[MIT License](LICENSE)

## 作者

longyongchong

创建于: 2025年10月30日