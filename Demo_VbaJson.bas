Attribute VB_Name = "Demo_VbaJson"
Sub Demo_VbaJson()
    ' ���� VbaJson ʵ��
    Dim jsonParser As New VbaJson
    
    ' ģ�� JSON ����
    Dim jsonString As String
    jsonString = "{""name"":""John"",""age"":30,""isActive"":true,""address"":{""street"":""123 Main St"",""city"":""New York""},""hobbies"":[""reading"",""swimming"",""coding""],""projects"":[{""name"":""Project A"",""status"":""completed""},{""name"":""Project B"",""status"":""in progress""}]}"
    
    ' 1. ���� JSON ����
    jsonParser.Json = jsonString
    
    ' 2. ʹ�� Stringify ������ʽ����� JSON
    Debug.Print "��ʽ����� JSON:"
    Debug.Print jsonParser.Stringify(2)
    Debug.Print "----------------------"
    
    ' 3. ʹ�� Count ���Ի�ȡ����Ԫ������
    Debug.Print "����Ԫ������: " & jsonParser.Count
    Debug.Print "----------------------"
    
    ' 4. ʹ�� Keys ���Ի�ȡ���ж������
    Debug.Print "���ж������: " & jsonParser.Keys
    Debug.Print "----------------------"
    
    ' 5. ʹ�� Item ������ȡ�ض�ֵ
    Debug.Print "��ȡ name ��ֵ: " & jsonParser.Item("name")
    Debug.Print "��ȡ address.city ��ֵ: " & jsonParser.Item("address.city")
    Debug.Print "��ȡ hobbies.1 ��ֵ: " & jsonParser.Item("hobbies.1")
    Debug.Print "��ȡ projects.0.name ��ֵ: " & jsonParser.Item("projects.0.name")
    Debug.Print "----------------------"
    
    ' 6. ʹ�� HasKey ���������Ƿ����
    Debug.Print "����Ƿ���� 'name' ��: " & jsonParser.HasKey("name")
    Debug.Print "����Ƿ���� 'address.zip' ��: " & jsonParser.HasKey("address.zip")
    Debug.Print "������� 'name' ����·��: " & jsonParser.HasKey("name", True)
    Debug.Print "----------------------"
    
    ' 7. ʹ�� KeyValuePair ���Ի�ȡ���м�ֵ��
    Debug.Print "���м�ֵ��:"
    Dim pairs As Variant
    pairs = Split(jsonParser.KeyValuePair, ",")
    Dim i As Integer
    For i = LBound(pairs) To UBound(pairs)
        Debug.Print pairs(i)
    Next i
    Debug.Print "----------------------"
    
    ' 8. ʹ�� Data ���Ի�ȡ���м���ֵ�ı�ƽ������
    Debug.Print "���м���ֵ�ı�ƽ������:"
    Dim dataItems As Variant
    dataItems = Split(jsonParser.Data, ",")
    For i = LBound(dataItems) To UBound(dataItems)
        Debug.Print dataItems(i)
    Next i
    Debug.Print "----------------------"
    
    ' 9. ʹ�� getAllKeys ������ȡ����Ψһ����
    Debug.Print "����Ψһ����:"
    Dim allKeys As Variant
    allKeys = Split(jsonParser.getAllKeys, ",")
    For i = LBound(allKeys) To UBound(allKeys)
        Debug.Print allKeys(i)
    Next i
    
    ' �������
    Set jsonParser = Nothing
End Sub
