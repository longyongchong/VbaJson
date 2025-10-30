VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'类模块 JsonParser
'作者：longyongchong
'时间：2025年10月30日
'功能：提供JSON解析和处理功能的VBA类模块

Option Explicit

Private htmlDoc As Object      ' HTML文档对象，用于执行JavaScript
Private jsWindow As Object     ' HTML窗口对象，用于访问JavaScript环境
Private m_jsonString As String ' 存储JSON字符串

'--- JavaScript辅助函数库 (合并为单次注入，提升性能) ---
Private Const JAVASCRIPT_LIBRARY As String = _
    "function isJSON(str) { try { JSON.parse(str); return true; } catch(e) { return false; } };/*检查字符串是否为有效JSON*/" & _
    "function parseJSON(str) { try {return JSON.parse(str);} catch(e) { return str;} };/*解析JSON字符串*/" & _
    "function getValue(obj, path) { for (var keys = path.split('.'), i = 0, len = keys.length; i < len; i++) { var k = keys[i]; if (obj == null) return void 0; obj = obj[isNaN(k) || k === '' ? k : parseInt(k, 10)]; } return obj; };/*根据路径获取JSON中的值*/" & _
    "function getItem(obj,path){var t=getValue(obj,path);var item=getType(t)<2 ? JSON.stringify(t,null,0) : t;return item;};/*在getValue基础上将对象转为字符返回*/" & _
    "function getKeys(obj) { var keys = []; for(var k in obj) keys.push(k); return keys.join(','); };/*获取对象的所有键名*/" & _
    "function getType(obj){var type=Object.prototype.toString.call(obj)==='[object Array]' ? 0 : Object.prototype.toString.call(obj)==='[object Object]' ? 1 : 2 ; return type;};/*获取值的类型：0(数组),1(对象),2(其他基本类型)*/" & _
    "function processArray(arr,currentKey){var result=[];for(var i=0;i<arr.length;i++){var item=arr[i],arrayKey=currentKey+'.'+i;if(getType(item)===2){result.push(arrayKey+':'+item);}else{var subResult=getData(item,arrayKey);for(var j=0;j<subResult.length;j++){result.push(subResult[j]);}}}return result;};/*处理数组类型数据*/" & _
    "function processObject(obj,currentKey){return getData(obj,currentKey);};/*处理对象类型数据*/" & _
    "function getData(obj,prefix){prefix=prefix||'';var d=[];for(var key in obj)if(obj.hasOwnProperty(key)){var value=obj[key],currentKey=prefix?prefix+'.'+key:key,type=getType(value);if(type===2){d.push(currentKey+':'+value);}else if(type===0){var subArray=processArray(value,currentKey);for(var i=0;i<subArray.length;i++){d.push(subArray[i]);}}else{var subObject=processObject(value,currentKey);for(var j=0;j<subObject.length;j++){d.push(subObject[j]);}}}return d;};/*获取结构化数据*/" & _
    "function findKey(obj, key) { if (obj === null || typeof obj !== 'object') return false; if (Object.prototype.hasOwnProperty.call(obj, key)) return true; for (var k in obj) { if (Object.prototype.hasOwnProperty.call(obj, k) && findKey(obj[k], key)) return true; } return false;};/*查找是否存在key值*/" & _
    "function findKeyPaths(obj,targetKey){var paths=[];(function traverse(currentObj,currentPath){if(currentObj===null||currentObj===undefined)return;if(Object.prototype.toString.call(currentObj)==='[object Array]'){for(var i=0;i<currentObj.length;i++){var newPath=currentPath?currentPath+'.'+i:''+i;if(String(i)===String(targetKey))paths.push(newPath);traverse(currentObj[i],newPath);}}else if(typeof currentObj==='object'){for(var key in currentObj){if(currentObj.hasOwnProperty(key)){var newPath=currentPath?currentPath+'.'+key:key;if(key===targetKey)paths.push(newPath);traverse(currentObj[key],newPath);}}}})(obj,'');return paths;};/*查找指定键的所有路径*/" & _
    "function setDeepValue(o,p,v){var k=p.replace(/\[(\d+)\]/g,'.$1').split('.'),c=o,i,l=k.length-1;for(i=0;i<l;i++){var t=k[i];if(t===''||t==null)continue;if(!(t in c)||c[t]==null||typeof c[t]!=='object')c[t]=/^\d+$/.test(k[i+1])?[]:{};c=c[t];}c[k[l]]=v;};/*增加key、value*/" & _
    "function delDeepKey(o,p){var k=p.replace(/\[(\d+)\]/g,'.$1').split('.'),c=o,i,l=k.length-1;for(i=0;i<l;i++){var t=k[i];if(!c||typeof c!=='object'||!(t in c))return;c=c[t];}if(c&&typeof c==='object')delete c[k[l]];};/*删除key*/" & _
    "function extractKeys(arr){var r=[],i=0,l=arr.length;for(;i<l;i++){var k=arr[i].split(':');r.push(k[0]);}return r;};/*获取JSON中所有唯一的键名（包括点号.嵌套的）*/"
    

' 类初始化事件
Private Sub Class_Initialize()
    Set htmlDoc = CreateObject("htmlfile") ' 创建HTML文档对象并设置兼容模式
    htmlDoc.Open
    htmlDoc.Write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"
    htmlDoc.Close
    
    Set jsWindow = htmlDoc.parentWindow
    jsWindow.execScript JAVASCRIPT_LIBRARY
End Sub

' 安全地在JS环境中设置变量
Private Sub SetJsVariable(ByVal Name As String, ByVal Value As Variant)
    jsWindow.execScript "var " & Name & " = " & Value & ";", "JavaScript"
End Sub

' 转义字符串，用于注入JS字符串字面量
Private Function EscapeForJS(str As String) As String
    EscapeForJS = Replace(str, "\", "\\")
    EscapeForJS = Replace(EscapeForJS, "'", "\'")
    EscapeForJS = Replace(EscapeForJS, """", "\""")
    EscapeForJS = Replace(EscapeForJS, vbCrLf, "\n")
End Function

'【属性】设置JSON字符串   ' 设置JSON字符串并验证格式有效性   ' 如果格式无效会抛出错误
Public Property Let JsonString(ByVal json As String)
    m_jsonString = json
    If Not jsWindow.isJSON(m_jsonString) Then
        Err.Raise vbObjectError + 1000, "VbaJson.JsonString", "提供的字符串不是有效的JSON格式。"
        Exit Property
    End If
    SetJsVariable "json", "parseJSON('" & EscapeForJS(m_jsonString) & "');"
End Property

'【属性】获取JSON字符串   ' 返回格式化后的JSON字符串
Public Property Get JsonString() As String
    JsonString = Me.Stringify(0)
End Property

'【属性】获取所有键值对   ' 返回所有键值对的数组，格式为"key:value"
Public Property Get KeyValuePairs() As Variant
    KeyValuePairs = jsWindow.GetData(jsWindow.json)
End Property

'【属性】获取JSON对象的类型
' @param iNumber - 返回格式：0=数字(0=数组,1=对象,2=其他), 1=符号([]/{}/Other), 2=文字(Array/Object/Other)
Public Property Get JsonType(Optional ByVal inumber As Integer = 0) As Variant
    On Error GoTo ErrorHandler
    
    Dim typeValue As Integer
    typeValue = jsWindow.eval("getType(json)")
    Select Case inumber
    Case 0: JsonType = typeValue
    Case 1: JsonType = Array("[]", "{}", "Other")(typeValue)
    Case Else: JsonType = Array("Array", "Object", "Other")(typeValue)
    End Select
    Exit Property
    
ErrorHandler:
    JsonType = "Invalid"
End Property

'【属性】获取JSON对象或数组的元素数量   ' 如果是数组返回长度，如果是对象返回键的数量，其他情况返回0
Public Property Get Count() As Integer
    On Error GoTo ErrorHandler
    
    Select Case jsWindow.getType(jsWindow.json)
    Case 0: Count = jsWindow.eval("json.length")  'Array
    Case 1: Count = jsWindow.eval("Object.keys(json).length")   ' Object
    Case Else: Count = 0
    End Select
    Exit Property
    
ErrorHandler:
    Count = 0
End Property

'【方法】在JSON中增加键值对   ' 支持嵌套路径，如"user.name"或"user[0].name"
' @param Key - 键名或路径     ' @param Value - 对应的值
Public Function Add(ByVal Key As String, ByVal Value As Variant) As Variant
    jsWindow.setDeepvalue jsWindow.json, Key, jsWindow.ParseJson(Value)
End Function

'【方法】删除JSON中的指定键   ' @param Key - 要删除的键名或路径
Public Function Delete(ByVal Key As String) As Boolean
    Dim blnQian As Boolean, blnHou As Boolean
    blnQian = Not (IsEmpty(Me.Item(Key)))
    jsWindow.DelDeepKey jsWindow.json, Key
    blnHou = Not (jsWindow.FindKey(jsWindow.json, Key))
    Delete = blnQian And blnHou
End Function

'【属性】检查JSON中是否包含指定键   ' @param Key - 要查找的键名
Public Property Get ContainsKey(ByVal Key As String) As Boolean
    ContainsKey = jsWindow.FindKey(jsWindow.json, Key)
End Property

'【属性】查找指定键的所有路径
' @param Key - 要查找的键名   ' @return 包含所有匹配路径的数组
Public Property Get FindKeyPaths(ByVal Key As String) As Variant
    FindKeyPaths = jsWindow.FindKeyPaths(jsWindow.json, Key)
End Property

'【方法】根据键名或路径获取值   ' 支持嵌套路径，如"user.name"或"user[0].name"
' @param Key - 键名或路径       ' @return 对应的值，如果是对象则返回JSON字符串
Public Function Item(ByVal KeyPath As Variant) As Variant
    Item = jsWindow.GetItem(jsWindow.json, KeyPath)
End Function

'【属性】获取所有顶层键名   ' 返回顶层键名的字符串，以逗号分隔
Public Property Get Keys() As String
    Keys = jsWindow.getKeys(jsWindow.json)
End Property

'【方法】获取JSON中所有唯一的键名   ' 包括嵌套的键名，返回以逗号分隔的字符串
Public Function GetAllKeys() As Variant
    GetAllKeys = jsWindow.eval("extractKeys(getData(json)).join(',')")
End Function

'【方法】将JSON对象格式化为字符串
' @param iNumber - 缩进空格数，默认为4，设为0则不格式化   ' @return 格式化后的JSON字符串
Public Function Stringify(Optional ByVal inumber As Integer = 4) As String
    Stringify = jsWindow.eval("JSON.stringify(json,null," & inumber & ")")
End Function

' 类终止事件，清理对象
Private Sub Class_Terminate()
    Set htmlDoc = Nothing
    Set jsWindow = Nothing
End Sub

