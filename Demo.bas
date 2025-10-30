Attribute VB_Name = "Demo"
Sub JsonParserDemo()
    ' ����JsonParserʵ��
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    ' ��ʾ�õ�JSON����
    Dim JsonString As String
    JsonString = "{""name"":""����"",""age"":30,""city"":""����"",""hobbies"":[""����"",""��Ӿ"",""���""],""address"":{""street"":""������"",""number"":123,""details"":{""room"":""A101"",""floor"":5}},""friends"":[{""name"":""����"",""age"":28},{""name"":""����"",""age"":32}]}"
    
    Debug.Print "=== JsonParser��ģ����ʾ ==="
    Debug.Print ""
    
    ' 1. ��ʾJsonString���ԣ����ã�
    Debug.Print "1. ����JSON�ַ�����"
    jp.JsonString = JsonString
    Debug.Print "JSON�ַ������óɹ�"
    Debug.Print ""
    
    ' 2. ��ʾJsonString���ԣ���ȡ��
    Debug.Print "2. ��ȡJSON�ַ�������ʽ������"
    Debug.Print jp.JsonString
    Debug.Print ""
    
    ' 3. ��ʾJsonType����
    Debug.Print "3. JSON�������ͣ�"
    Debug.Print "���ָ�ʽ: " & jp.JsonType(0)  ' 0=����,1=����,2=����
    Debug.Print "���Ÿ�ʽ: " & jp.JsonType(1)  ' []/{}/Other
    Debug.Print "���ָ�ʽ: " & jp.JsonType(2)  ' Array/Object/Other
    Debug.Print ""
    
    ' 4. ��ʾCount����
    Debug.Print "4. JSON����Ԫ��������"
    Debug.Print "����Ԫ������: " & jp.Count
    Debug.Print ""
    
    ' 5. ��ʾItem���� - ��ȡֵ
    Debug.Print "5. ��ȡֵ��ʾ��"
    Debug.Print "����: " & jp.Item("name")
    Debug.Print "����: " & jp.Item("age")
    Debug.Print "����: " & jp.Item("city")
    Debug.Print "��һ������: " & jp.Item("hobbies.0")
    Debug.Print "�ڶ�������: " & jp.Item("hobbies.1")
    Debug.Print "�ֵ�: " & jp.Item("address.street")
    Debug.Print "���ƺ�: " & jp.Item("address.number")
    Debug.Print "�����: " & jp.Item("address.details.room")
    Debug.Print "��һ����������: " & jp.Item("friends.0.name")
    Debug.Print "�ڶ�����������: " & jp.Item("friends.1.age")
    Debug.Print ""
    
    ' 6. ��ʾKeys����
    Debug.Print "6. ��ȡ���������"
    Debug.Print "�������: " & jp.Keys
    Debug.Print ""
    
    ' 7. ��ʾGetAllKeys����
    Debug.Print "7. ��ȡ���м���������Ƕ�ף���"
    Debug.Print "���м���: " & jp.GetAllKeys
    Debug.Print ""
    
    ' 8. ��ʾContainsKey����
    Debug.Print "8. �����Ƿ���ڣ�"
    Debug.Print "����name��: " & jp.ContainsKey("name")
    Debug.Print "����email��: " & jp.ContainsKey("email")
    Debug.Print "����address.street��: " & jp.ContainsKey("address.street")
    Debug.Print "����address.postalCode��: " & jp.ContainsKey("address.postalCode")
    Debug.Print ""
    
    ' 9. ��ʾFindKeyPaths����
    Debug.Print "9. ���Ҽ�������·����"
    Debug.Print "����name����·��: " & jp.FindKeyPaths("name")
    Debug.Print "����age����·��: " & jp.FindKeyPaths("age")
    Debug.Print "����city����·��: " & jp.FindKeyPaths("city")
    Debug.Print ""
    
    ' 10. ��ʾAdd���� - ��Ӽ�ֵ��
    Debug.Print "10. ��Ӽ�ֵ�ԣ�"
    jp.Add "email", "zhangsan@example.com"
    jp.Add "phone", "13800138000"
    jp.Add "address.postalCode", "100000"
    jp.Add "hobbies.3", "����"
    jp.Add "work", "{""company"":""ABC��˾"",""position"":""����ʦ""}"
    Debug.Print "���email, phone, postalCode, �°���, ������Ϣ�ɹ�"
    Debug.Print "������: " & jp.Item("email")
    Debug.Print "�µ绰: " & jp.Item("phone")
    Debug.Print "���ʱ�: " & jp.Item("address.postalCode")
    Debug.Print "�°���: " & jp.Item("hobbies.3")
    Debug.Print "������˾: " & jp.Item("work.company")
    Debug.Print ""
    
    ' 11. �ٴβ鿴������JSON
    Debug.Print "11. ���º��JSON����ʽ������"
    Debug.Print jp.Stringify(2)  ' 2���ո�����
    Debug.Print ""
    
    ' 12. ��ʾDelete���� - ɾ����
    Debug.Print "12. ɾ����ֵ�ԣ�"
    Dim deleteResult As Boolean
    deleteResult = jp.Delete("phone")
    Debug.Print "ɾ��phone��: " & deleteResult
    deleteResult = jp.Delete("address.details.floor")
    Debug.Print "ɾ��address.details.floor��: " & deleteResult
    Debug.Print "ɾ�����Ի�ȡphone: " & jp.Item("phone")  ' Ӧ��Ϊ�ջ�Ĭ��ֵ
    Debug.Print "ɾ�����Ի�ȡfloor: " & jp.Item("address.details.floor")  ' Ӧ��Ϊ�ջ�Ĭ��ֵ
    Debug.Print ""
    
    ' 13. ��ʾKeyValuePairs����
    Debug.Print "13. ��ȡ���м�ֵ�ԣ�"
    Debug.Print "���м�ֵ��: " & jp.KeyValuePairs
    Debug.Print ""
    
    ' 14. ��ʾStringify���� - ��ͬ������ʽ
    Debug.Print "14. ��ͬ������ʽ��JSON�ַ�����"
    Debug.Print "��������"
    Debug.Print jp.Stringify(0)
    Debug.Print ""
    Debug.Print "4���ո�������"
    Debug.Print jp.Stringify(4)
    Debug.Print ""
    
    ' 15. ��ʾ�������͵�JSON
    Debug.Print "15. ��������JSON��ʾ��"
    Dim arrayJson As String
    arrayJson = "[{""name"":""��ƷA"",""price"":100},{""name"":""��ƷB"",""price"":200},{""name"":""��ƷC"",""price"":150}]"
    jp.JsonString = arrayJson
    Debug.Print "����JSON����: " & jp.JsonType(2)
    Debug.Print "���鳤��: " & jp.Count
    Debug.Print "��һ����Ʒ��: " & jp.Item("0.name")
    Debug.Print "�ڶ�����Ʒ�۸�: " & jp.Item("1.price")
    Debug.Print "���м���: " & jp.GetAllKeys
    Debug.Print ""
    
    ' 16. ��ʾ������ - ��ЧJSON
    Debug.Print "16. ��������ʾ����ЧJSON����"
    On Error Resume Next
    jp.JsonString = "{invalid json"
    If Err.Number <> 0 Then
        Debug.Print "���񵽴���: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Debug.Print ""
    
    ' 17. ����״̬չʾ
    Debug.Print "17. ����JSON״̬��"
    jp.JsonString = JsonString  ' ����ΪԭʼJSON
    jp.Add "demo", "completed"  ' �����ʾ��ɱ��
    Debug.Print jp.Stringify(2)
    Debug.Print ""
    
    Debug.Print "=== JsonParser��ʾ��� ==="
    
    ' �������
    Set jp = Nothing
End Sub

' ����Ĳ��Ժ��� - �������ⳡ��
Sub JsonParserAdvancedDemo()
    Debug.Print "=== �߼�������ʾ ==="
    
    Dim jp As JsonParser
    Set jp = New JsonParser
    
    ' ����Ƕ������Ͷ���ĸ��ӽṹ
    Dim complexJson As String
    complexJson = "{""users"":[{""id"":1,""profile"":{""name"":""�û�1"",""settings"":{""theme"":""dark"",""notifications"":true}}},{""id"":2,""profile"":{""name"":""�û�2"",""settings"":{""theme"":""light"",""notifications"":false}}}],""config"":{""version"":""1.0"",""features"":[""feature1"",""feature2"",""feature3""]}}"
    
    jp.JsonString = complexJson
    
    Debug.Print "����JSON�ṹ��ʾ��"
    Debug.Print "�û�1����: " & jp.Item("users.0.profile.settings.theme")
    Debug.Print "�û�2֪ͨ: " & jp.Item("users.1.profile.settings.notifications")
    Debug.Print "���ð汾: " & jp.Item(jp.FindKeyPaths("version"))
    Debug.Print "��һ������: " & jp.Item("config.features.0")
    
    ' �����ض���������·��
    Debug.Print "����theme����·��: " & jp.FindKeyPaths("theme")
    Debug.Print "����name����·��: " & jp.FindKeyPaths("name")
    
    ' ����µ��û�
    jp.Add "users.2", "{""id"":3,""profile"":{""name"":""�û�3"",""settings"":{""theme"":""auto"",""notifications"":true}}}"
    Debug.Print "������û����û�����: " & jp.Item("users.2.profile.name")
    
    ' ��ʾɾ��Ƕ�׶���
    jp.Delete "users.1.profile.settings.notifications"
    Debug.Print "ɾ���û�2֪ͨ���ú�ֵΪ: " & jp.Item("users.1.profile.settings.notifications")
    
    ' �������
    Debug.Print "���ո���JSON�ṹ��"
    Debug.Print jp.Stringify(2)
    
    Set jp = Nothing
    Debug.Print "=== �߼�������ʾ��� ==="
End Sub
