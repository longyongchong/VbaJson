VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'类模块 VbaJson
'作者：longyongchong
'时间：2025年10月7日
'功能：提供JSON解析和处理功能的VBA类模块

Option Explicit

' 私有变量声明
Dim Html As Object        ' HTML文档对象，用于执行JavaScript
Dim oW As Object         ' HTML窗口对象，用于访问JavaScript环境
Dim sJS As String        ' 存储JSON字符串

' 类初始化事件
Private Sub Class_Initialize()
    Set Html = CreateObject("htmlfile") ' 创建HTML文档对象并设置兼容模式
    Html.write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"
    Set oW = Html.parentWindow
    
    With oW ' 在JavaScript环境中添加JSON处理函数
        .execScript "function getLength(arr) { return arr.length; };" ' 获取数组长度
        .execScript "function isJSON(str) { try { JSON.parse(str); return true; } catch(e) { return false; } };" ' 检查字符串是否为有效JSON
        .execScript "function parseJSON(str) { return JSON.parse(str); };" ' 解析JSON字符串
        .execScript "function getValue(obj, path) { for (var keys = path.split('.'), i = 0, len = keys.length; i < len; i++) { var k = keys[i]; if (obj == null) return void 0; obj = obj[isNaN(k) || k === '' ? k : parseInt(k, 10)]; } return obj; }" ' 根据路径获取JSON中的值
        .execScript "function getKeys(obj) { var keys = []; for(var k in obj) keys.push(k); return keys.join(','); };" ' 获取对象的所有键名
        .execScript "function f(o,r){r=r||[];if(o!==null&&typeof o==='object'){if(Object.prototype.toString.call(o)==='[object Array]'){for(var i=0;i<o.length;i++)f(o[i],r);}else{for(var k in o)if(o.hasOwnProperty(k)){r.push(k);f(o[k],r);}}}else r.push(o);return r;};" ' 递归获取所有键和值
        .execScript "function isObject(obj) { return Object.prototype.toString.call(obj) === '[object Object]';};" ' 检查是否为对象
        .execScript "function getType(obj){var type=Object.prototype.toString.call(obj)==='[object Array]'?0:Object.prototype.toString.call(obj)==='[object Object]'?1:2;return type;};" ' 获取值的类型：0(数组),1(对象),2(其他基本类型)
        .execScript "function processArray(arr,currentKey){var result=[];for(var i=0;i<arr.length;i++){var item=arr[i],arrayKey=currentKey+'.'+i;if(getType(item)===2){result.push(arrayKey+':'+item);}else{var subResult=getData(item,arrayKey);for(var j=0;j<subResult.length;j++){result.push(subResult[j]);}}}return result;};" ' 处理数组类型数据
        .execScript "function processObject(obj,currentKey){return getData(obj,currentKey);};" ' 处理对象类型数据
        .execScript "function getData(obj,prefix){prefix=prefix||'';var d=[];for(var key in obj)if(obj.hasOwnProperty(key)){var value=obj[key],currentKey=prefix?prefix+'.'+key:key,type=getType(value);if(type===2){d.push(currentKey+':'+value);}else if(type===0){var subArray=processArray(value,currentKey);for(var i=0;i<subArray.length;i++){d.push(subArray[i]);}}else{var subObject=processObject(value,currentKey);for(var j=0;j<subObject.length;j++){d.push(subObject[j]);}}}return d;};" ' 获取结构化数据
        .execScript "function findKey(obj, key) { if (obj === null || typeof obj !== 'object') return false; if (Object.prototype.hasOwnProperty.call(obj, key)) return true; for (var k in obj) { if (Object.prototype.hasOwnProperty.call(obj, k) && findKey(obj[k], key)) return true; } return false;};"
        .execScript "function findKeyPaths(obj,targetKey){var paths=[];(function traverse(currentObj,currentPath){if(currentObj===null||currentObj===undefined)return;if(Object.prototype.toString.call(currentObj)==='[object Array]'){for(var i=0;i<currentObj.length;i++){var newPath=currentPath?currentPath+'.'+i:''+i;if(String(i)===String(targetKey))paths.push(newPath);traverse(currentObj[i],newPath);}}else if(typeof currentObj==='object'){for(var key in currentObj){if(currentObj.hasOwnProperty(key)){var newPath=currentPath?currentPath+'.'+key:key;if(key===targetKey)paths.push(newPath);traverse(currentObj[key],newPath);}}}})(obj,'');return paths;};" ' 查找指定键的所有路径
        .execScript "function getAllKeys(obj){var keys=[],seen={};function extractKeys(data){if(data!==null&&typeof data==='object'){for(var k in data){if(data.hasOwnProperty(k)){var keyStr=String(k);if(!seen[keyStr]){seen[keyStr]=true;keys.push(keyStr);}extractKeys(data[k]);}}}};extractKeys(obj);return keys;};" ' 获取所有唯一键名
    End With
End Sub

' 属性：设置JSON字符串
Public Property Let Json(ByVal Json As String)
    sJS = Json
    With oW
        .execScript "var str='" & Replace(sJS, "'", "\'") & "';" ' 将JSON字符串传递给JavaScript环境

        If Not .eval("isJSON(str)") Then        ' 检查是否为有效JSON
            MsgBox prompt:="非标准 json 数据字符串，请检查！", Buttons:=vbOKOnly + vbCritical, Title:="警告"
            Exit Property
        End If
        .execScript "var json=parseJSON(str);" ' 解析JSON字符串
    End With
End Property

' 方法：检查JSON中是否包含指定键
' 参数：key - 要查找的键名
'       bln - 可选参数，为True时返回所有匹配路径，为False时返回布尔值
Public Function HasKey(ByVal key As String, Optional ByVal bln As Boolean = False) As Variant
    oW.execScript "var key='" & key & "';"
    HasKey = IIf(bln, oW.eval("findKeyPaths(json,key)"), oW.eval("findKey(json, key)"))
End Function

' 属性：获取所有键值对（以逗号分隔的字符串）
Public Property Get KeyValuePair() As String
    KeyValuePair = oW.eval("getData(json).join(',')")
End Property

' 属性：获取所有键和值的扁平化数组（以逗号分隔的字符串）
Public Property Get Data() As String
    Data = oW.eval("f(json)")
End Property

' 属性：获取JSON对象或数组的元素数量
Public Property Get Count() As Variant
    Count = oW.eval("getLength(json)")
End Property

' 方法：根据键名或路径获取值
' 参数：key - 键名或路径（如"user.name"）
Public Function Item(ByVal key As Variant) As Variant
    With oW
        .execScript "var key='" & Replace(key, "'", "\'") & "';"
        .execScript "var temp=getValue(json,key);"
        .execScript "var item=getType(temp)<2?JSON.stringify(temp,null,0):temp;" ' 如果是对象则转换为JSON字符串，否则直接返回值
        Item = .eval("item")
    End With
End Function

' 属性：获取所有顶层键名（以逗号分隔的字符串）
Public Property Get Keys() As String
    Keys = oW.eval("getKeys(json)")
End Property

' 方法：获取JSON中所有唯一的键名（包括点号.嵌套的）
Public Function getAllKeys() As Variant
    getAllKeys = oW.eval("getAllKeys(json)")
End Function

' 方法：将JSON对象格式化为字符串
' 参数：iNumber - 可选参数，缩进空格数（默认4）
Public Function Stringify(Optional ByVal iNumber As Integer = 4) As String
    With oW
        .execScript "var num=" & iNumber & ";"
        Stringify = .eval("JSON.stringify(json,null,num)")
    End With
End Function

' 类终止事件，清理对象
Private Sub Class_Terminate()
    Set Html = Nothing
    Set oW = Nothing
End Sub
