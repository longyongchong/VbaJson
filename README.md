# VbaJson - JSON解析器 for VBA

VbaJson 是一个用于VBA的轻量级JSON解析器类模块，提供简单易用的接口来解析和处理JSON数据。

## 功能特性

- ✅ 解析标准JSON字符串
- ✅ 支持嵌套对象和数组访问
- ✅ 提供多种查询和遍历方法
- ✅ 支持JSON格式化输出
- ✅ 轻量级实现，不依赖外部库

## 安装方法

1. 将 `VbaJson.cls` 类模块复制到您的VBA项目中
2. 在代码中创建 `VbaJson` 类的实例即可使用

## 快速开始

```vba
Sub Demo()
    Dim jsonParser As New VbaJson
    Dim jsonString As String
    
    ' 示例JSON数据
    jsonString = "{""name"":""John"",""age"":30,""address"":{""city"":""New York""}}"
    
    ' 设置JSON数据
    jsonParser.Json = jsonString
    
    ' 获取值
    Debug.Print "Name: " & jsonParser.Item("name")  ' 输出: John
    Debug.Print "City: " & jsonParser.Item("address.city")  ' 输出: New York
    
    ' 格式化输出
    Debug.Print jsonParser.Stringify(2)
End Sub
```

## 完整API文档

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `Json` | String | 设置要解析的JSON字符串 |
| `KeyValuePair` | String | 获取所有键值对（逗号分隔） |
| `Data` | String | 获取所有键和值的扁平化数组（逗号分隔） |
| `Count` | Variant | 获取顶层元素数量 |
| `Keys` | String | 获取所有顶层键名（逗号分隔） |

### 方法

| 方法 | 返回值 | 说明 |
|------|--------|------|
| `HasKey(key As String, [bln As Boolean = False])` | Variant | 检查键是否存在，bln为True时返回所有匹配路径 |
| `Item(key As Variant)` | Variant | 根据键名或路径获取值 |
| `getAllKeys()` | Variant | 获取JSON中所有唯一的键名（包括嵌套的） |
| `Stringify([iNumber As Integer = 4])` | String | 将JSON对象格式化为字符串，iNumber为缩进空格数 |

## 使用示例

```vba
Sub ExampleUsage()
    Dim json As New VbaJson
    json.Json = "{""user"":{""name"":""Alice"",""age"":25,""hobbies"":[""reading"",""coding""]}}"
    
    ' 检查键是否存在
    Debug.Print json.HasKey("user.name")  ' 输出: True
    Debug.Print json.HasKey("user.email")  ' 输出: False
    
    ' 获取嵌套值
    Debug.Print json.Item("user.name")  ' 输出: Alice
    Debug.Print json.Item("user.hobbies.1")  ' 输出: coding
    
    ' 获取所有键
    Debug.Print json.Keys  ' 输出: user
    Debug.Print Join(json.getAllKeys(), ", ")  ' 输出: user, name, age, hobbies, 0, 1
    
    ' 格式化输出
    Debug.Print json.Stringify(2)
End Sub
```

## 贡献指南

欢迎提交问题和拉取请求！如果您有任何改进建议或发现了bug，请通过GitHub Issues提交。

## 许可证

本项目采用 LICENSE。
