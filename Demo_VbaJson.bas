Attribute VB_Name = "Demo_VbaJson"
Sub Demo_VbaJson()
    ' 创建 VbaJson 实例
    Dim jsonParser As New VbaJson
    
    ' 模拟 JSON 数据
    Dim jsonString As String
    jsonString = "{""name"":""John"",""age"":30,""isActive"":true,""address"":{""street"":""123 Main St"",""city"":""New York""},""hobbies"":[""reading"",""swimming"",""coding""],""projects"":[{""name"":""Project A"",""status"":""completed""},{""name"":""Project B"",""status"":""in progress""}]}"
    
    ' 1. 设置 JSON 属性
    jsonParser.Json = jsonString
    
    ' 2. 使用 Stringify 方法格式化输出 JSON
    Debug.Print "格式化后的 JSON:"
    Debug.Print jsonParser.Stringify(2)
    Debug.Print "----------------------"
    
    ' 3. 使用 Count 属性获取顶层元素数量
    Debug.Print "顶层元素数量: " & jsonParser.Count
    Debug.Print "----------------------"
    
    ' 4. 使用 Keys 属性获取所有顶层键名
    Debug.Print "所有顶层键名: " & jsonParser.Keys
    Debug.Print "----------------------"
    
    ' 5. 使用 Item 方法获取特定值
    Debug.Print "获取 name 的值: " & jsonParser.Item("name")
    Debug.Print "获取 address.city 的值: " & jsonParser.Item("address.city")
    Debug.Print "获取 hobbies.1 的值: " & jsonParser.Item("hobbies.1")
    Debug.Print "获取 projects.0.name 的值: " & jsonParser.Item("projects.0.name")
    Debug.Print "----------------------"
    
    ' 6. 使用 HasKey 方法检查键是否存在
    Debug.Print "检查是否存在 'name' 键: " & jsonParser.HasKey("name")
    Debug.Print "检查是否存在 'address.zip' 键: " & jsonParser.HasKey("address.zip")
    Debug.Print "检查所有 'name' 键的路径: " & jsonParser.HasKey("name", True)
    Debug.Print "----------------------"
    
    ' 7. 使用 KeyValuePair 属性获取所有键值对
    Debug.Print "所有键值对:"
    Dim pairs As Variant
    pairs = Split(jsonParser.KeyValuePair, ",")
    Dim i As Integer
    For i = LBound(pairs) To UBound(pairs)
        Debug.Print pairs(i)
    Next i
    Debug.Print "----------------------"
    
    ' 8. 使用 Data 属性获取所有键和值的扁平化数组
    Debug.Print "所有键和值的扁平化数组:"
    Dim dataItems As Variant
    dataItems = Split(jsonParser.Data, ",")
    For i = LBound(dataItems) To UBound(dataItems)
        Debug.Print dataItems(i)
    Next i
    Debug.Print "----------------------"
    
    ' 9. 使用 getAllKeys 方法获取所有唯一键名
    Debug.Print "所有唯一键名:"
    Dim allKeys As Variant
    allKeys = Split(jsonParser.getAllKeys, ",")
    For i = LBound(allKeys) To UBound(allKeys)
        Debug.Print allKeys(i)
    Next i
    
    ' 清理对象
    Set jsonParser = Nothing
End Sub
