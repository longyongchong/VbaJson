Attribute VB_Name = "Demo"
Sub JsonParserDemo()
    ' 创建JsonParser实例
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    ' 演示用的JSON数据
    Dim JsonString As String
    JsonString = "{""name"":""张三"",""age"":30,""city"":""北京"",""hobbies"":[""读书"",""游泳"",""编程""],""address"":{""street"":""长安街"",""number"":123,""details"":{""room"":""A101"",""floor"":5}},""friends"":[{""name"":""李四"",""age"":28},{""name"":""王五"",""age"":32}]}"
    
    Debug.Print "=== JsonParser类模块演示 ==="
    Debug.Print ""
    
    ' 1. 演示JsonString属性（设置）
    Debug.Print "1. 设置JSON字符串："
    jp.JsonString = JsonString
    Debug.Print "JSON字符串设置成功"
    Debug.Print ""
    
    ' 2. 演示JsonString属性（获取）
    Debug.Print "2. 获取JSON字符串（格式化）："
    Debug.Print jp.JsonString
    Debug.Print ""
    
    ' 3. 演示JsonType属性
    Debug.Print "3. JSON对象类型："
    Debug.Print "数字格式: " & jp.JsonType(0)  ' 0=数组,1=对象,2=其他
    Debug.Print "符号格式: " & jp.JsonType(1)  ' []/{}/Other
    Debug.Print "文字格式: " & jp.JsonType(2)  ' Array/Object/Other
    Debug.Print ""
    
    ' 4. 演示Count属性
    Debug.Print "4. JSON对象元素数量："
    Debug.Print "顶层元素数量: " & jp.Count
    Debug.Print ""
    
    ' 5. 演示Item方法 - 获取值
    Debug.Print "5. 获取值演示："
    Debug.Print "姓名: " & jp.Item("name")
    Debug.Print "年龄: " & jp.Item("age")
    Debug.Print "城市: " & jp.Item("city")
    Debug.Print "第一个爱好: " & jp.Item("hobbies.0")
    Debug.Print "第二个爱好: " & jp.Item("hobbies.1")
    Debug.Print "街道: " & jp.Item("address.street")
    Debug.Print "门牌号: " & jp.Item("address.number")
    Debug.Print "房间号: " & jp.Item("address.details.room")
    Debug.Print "第一个朋友姓名: " & jp.Item("friends.0.name")
    Debug.Print "第二个朋友年龄: " & jp.Item("friends.1.age")
    Debug.Print ""
    
    ' 6. 演示Keys属性
    Debug.Print "6. 获取顶层键名："
    Debug.Print "顶层键名: " & jp.Keys
    Debug.Print ""
    
    ' 7. 演示GetAllKeys方法
    Debug.Print "7. 获取所有键名（包括嵌套）："
    Debug.Print "所有键名: " & jp.GetAllKeys
    Debug.Print ""
    
    ' 8. 演示ContainsKey属性
    Debug.Print "8. 检查键是否存在："
    Debug.Print "包含name键: " & jp.ContainsKey("name")
    Debug.Print "包含email键: " & jp.ContainsKey("email")
    Debug.Print "包含address.street键: " & jp.ContainsKey("address.street")
    Debug.Print "包含address.postalCode键: " & jp.ContainsKey("address.postalCode")
    Debug.Print ""
    
    ' 9. 演示FindKeyPaths属性
    Debug.Print "9. 查找键的所有路径："
    Debug.Print "查找name键的路径: " & jp.FindKeyPaths("name")
    Debug.Print "查找age键的路径: " & jp.FindKeyPaths("age")
    Debug.Print "查找city键的路径: " & jp.FindKeyPaths("city")
    Debug.Print ""
    
    ' 10. 演示Add方法 - 添加键值对
    Debug.Print "10. 添加键值对："
    jp.Add "email", "zhangsan@example.com"
    jp.Add "phone", "13800138000"
    jp.Add "address.postalCode", "100000"
    jp.Add "hobbies.3", "旅游"
    jp.Add "work", "{""company"":""ABC公司"",""position"":""工程师""}"
    Debug.Print "添加email, phone, postalCode, 新爱好, 工作信息成功"
    Debug.Print "新邮箱: " & jp.Item("email")
    Debug.Print "新电话: " & jp.Item("phone")
    Debug.Print "新邮编: " & jp.Item("address.postalCode")
    Debug.Print "新爱好: " & jp.Item("hobbies.3")
    Debug.Print "工作公司: " & jp.Item("work.company")
    Debug.Print ""
    
    ' 11. 再次查看完整的JSON
    Debug.Print "11. 更新后的JSON（格式化）："
    Debug.Print jp.Stringify(2)  ' 2个空格缩进
    Debug.Print ""
    
    ' 12. 演示Delete方法 - 删除键
    Debug.Print "12. 删除键值对："
    Dim deleteResult As Boolean
    deleteResult = jp.Delete("phone")
    Debug.Print "删除phone键: " & deleteResult
    deleteResult = jp.Delete("address.details.floor")
    Debug.Print "删除address.details.floor键: " & deleteResult
    Debug.Print "删除后尝试获取phone: " & jp.Item("phone")  ' 应该为空或默认值
    Debug.Print "删除后尝试获取floor: " & jp.Item("address.details.floor")  ' 应该为空或默认值
    Debug.Print ""
    
    ' 13. 演示KeyValuePairs属性
    Debug.Print "13. 获取所有键值对："
    Debug.Print "所有键值对: " & jp.KeyValuePairs
    Debug.Print ""
    
    ' 14. 演示Stringify方法 - 不同缩进格式
    Debug.Print "14. 不同缩进格式的JSON字符串："
    Debug.Print "无缩进："
    Debug.Print jp.Stringify(0)
    Debug.Print ""
    Debug.Print "4个空格缩进："
    Debug.Print jp.Stringify(4)
    Debug.Print ""
    
    ' 15. 演示数组类型的JSON
    Debug.Print "15. 数组类型JSON演示："
    Dim arrayJson As String
    arrayJson = "[{""name"":""商品A"",""price"":100},{""name"":""商品B"",""price"":200},{""name"":""商品C"",""price"":150}]"
    jp.JsonString = arrayJson
    Debug.Print "数组JSON类型: " & jp.JsonType(2)
    Debug.Print "数组长度: " & jp.Count
    Debug.Print "第一个商品名: " & jp.Item("0.name")
    Debug.Print "第二个商品价格: " & jp.Item("1.price")
    Debug.Print "所有键名: " & jp.GetAllKeys
    Debug.Print ""
    
    ' 16. 演示错误处理 - 无效JSON
    Debug.Print "16. 错误处理演示（无效JSON）："
    On Error Resume Next
    jp.JsonString = "{invalid json"
    If Err.Number <> 0 Then
        Debug.Print "捕获到错误: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' 17. 最终状态展示
    Debug.Print "17. 最终JSON状态："
    jp.JsonString = JsonString  ' 重置为原始JSON
    jp.Add "demo", "completed"  ' 添加演示完成标记
    Debug.Print jp.Stringify(2)
    Debug.Print ""
    
    Debug.Print "=== JsonParser演示完成 ==="
    
    ' 清理对象
    Set jp = Nothing
End Sub

' 额外的测试函数 - 测试特殊场景
Sub JsonParserAdvancedDemo()
    Debug.Print "=== 高级功能演示 ==="
    
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    ' 测试嵌套数组和对象的复杂结构
    Dim complexJson As String
    complexJson = "{""users"":[{""id"":1,""profile"":{""name"":""用户1"",""settings"":{""theme"":""dark"",""notifications"":true}}},{""id"":2,""profile"":{""name"":""用户2"",""settings"":{""theme"":""light"",""notifications"":false}}}],""config"":{""version"":""1.0"",""features"":[""feature1"",""feature2"",""feature3""]}}"
    
    jp.JsonString = complexJson
    
    Debug.Print "复杂JSON结构演示："
    Debug.Print "用户1主题: " & jp.Item("users.0.profile.settings.theme")
    Debug.Print "用户2通知: " & jp.Item("users.1.profile.settings.notifications")
    Debug.Print "配置版本: " & jp.Item(jp.FindKeyPaths("version"))
    Debug.Print "第一个功能: " & jp.Item("config.features.0")
    
    ' 查找特定键的所有路径
    Debug.Print "所有theme键的路径: " & jp.FindKeyPaths("theme")
    Debug.Print "所有name键的路径: " & jp.FindKeyPaths("name")
    
    ' 添加新的用户
    jp.Add "users.2", "{""id"":3,""profile"":{""name"":""用户3"",""settings"":{""theme"":""auto"",""notifications"":true}}}"
    Debug.Print "添加新用户后用户数量: " & jp.Item("users.2.profile.name")
    
    ' 演示删除嵌套对象
    jp.Delete "users.1.profile.settings.notifications"
    Debug.Print "删除用户2通知设置后，值为: " & jp.Item("users.1.profile.settings.notifications")
    
    ' 完整输出
    Debug.Print "最终复杂JSON结构："
    Debug.Print jp.Stringify(2)
    
    Set jp = Nothing
    Debug.Print "=== 高级功能演示完成 ==="
End Sub
