VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'http://www.excelworld.ru/forum/3-4357-4

'������ ��� ��� ��������� �� ������, ����� ������ �� ADO

'���� ��������� [test1$A:C] - ������ �� 65000 ����� !
'� ���� ������, ��� ������� (��� ����) ������, �������� ������ "����������� �������� ��� ������ ���� ���������� ��������� ����������".
'���� ��������� �� ���������� [test1$A:C]
'FROM [A1:D1048570] �������� ������ ��� ������ �����, ������������ ���� ������
'������ �������� � �������� ���������� ��������� � ������� 3,16*1
'�ells.NumberFormat = "@" ��������� ��� ������ ''
'����� �������� ������ ������ ������ �������
' ��������� - FROM [A1:D1048570] /// FROM [" & ActiveSheet.Name & "$]
' ��������� ��� ��� - WHERE F1 BETWEEN #5/1/2018# AND #5/25/2018#
'--------------------------------
' @author nerv
' @version 07/05/2013, 0.7
'--------------------------------
' ����� ��� ������ � ��������� Connection � Recordset; ����������� SQL �������� � ������ ������, ��������� ������, ����� ������.
'
' ADO (�� ����. ActiveX Data Objects � �������� ������ ActiveX�) � ��������� ���������������� ���������� ��� ������� � ������,
' ������������� ��������� Microsoft (MS Access, MS SQL Server) � ���������� �� ���������� ����������� ActiveX.
' ADO ��������� ������������ ������ �� ������������� ���������� (����������� ��� ������, ��������� ������ � �. �.)
' � ��������-��������������� ����.
'
' reference - http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx
' wiki - http://ru.wikipedia.org/wiki/ADO
'
' http://support.microsoft.com/kb/257819/ru
' http://support.microsoft.com/kb/316934/ru
' http://support.microsoft.com/kb/246335/ru
'
' http://www.script-coding.com/ADO.html
' http://www.w3schools.com/ado/default.asp

' ������:
'   Create     - ������� ����� �����������. ������������� ���������� ��� �������������.
'   Connect    - ��������� ����������. ������������� ���������� ��� �������.
'   Destroy    - ���������� ������ ����������� � ������ �������. ���������� ������������� ��� ������ �� ���������.
'   Disconnect - ��������� �������� ������ � �����������. ���������� ������������� ��� ������ �� ���������.
'   Query      - ��������� SQL ������. ��������� ������� �������� � ������ Recordset. ���������� �����, � ������� ��� �������� ������.
'   ToArray    - ���������� ��������� ������� � ���� �������

' ��������:
'   Connection - ������ ����������.
'   Recordset  - ��������� ��������� �������.
'   DataSoure  - �������� ������. ������ ��� ����� ������. �� ��������� ������� �����.
'   Header     - ��������� ��������� (��/���). �� ��������� ���. � ���� ������ ����� ����� ����������� ������������� F1 ... Fn.
'                ���� ��, ������ ������ ��������� ��������� ���������� ����.

' � ������ �������� ��������� (ConnectionString) � ����� Connect, �������� ������� DataSoure � Header �� �����������,
' � ������������ ������ ���������� �������� ��������� �� ������������.


Option Explicit


Public Connection As Object
Public Recordset As Object
Public DataSource As String
Public Header As Boolean


Public Sub Create()
    Set Me.Connection = CreateObject("ADODB.Connection")
    Set Me.Recordset = CreateObject("ADODB.Recordset")
    
    Me.Recordset.CursorLocation = 3 ' adUseClient
End Sub


Public Sub Destroy()
    Call Me.Disconnect
    
    Set Me.Recordset = Nothing
    Set Me.Connection = Nothing
End Sub


Public Sub Connect(Optional ByVal ConnectionString As String)
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If

    If ConnectionString = "" Then
        Me.Connection.Open GetExcelConnectionString()
    Else
        Me.Connection.Open ConnectionString
    End If
End Sub


Public Sub Disconnect()
    If Not Me.Recordset Is Nothing Then
        If Me.Recordset.State = 1 Then
            Me.Recordset.Close
        End If
    End If
    
    If Not Me.Connection Is Nothing Then
        If Me.Connection.State = 1 Then
            Me.Connection.Close
        End If
    End If
End Sub


Public Function Query(ParamArray QueryString() As Variant) As Date
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If
    
    If Me.Connection.State = 0 Then
        Call Me.Connect
    End If
    
    If Me.Recordset.State = 1 Then
        Me.Recordset.Close
    End If
    
    Me.Recordset.Open Join(QueryString, " "), Me.Connection
    
    Query = Now
End Function


Public Function ToArray() As Variant
    Dim Field As Object
    Dim arr() As Variant
    Dim Column As Integer
    Dim row As Long
    
    If Me.Recordset Is Nothing Then
        Exit Function
    End If
    
    If Me.Recordset.BOF And Me.Recordset.EOF Then
        Exit Function
    End If
    
    ReDim arr(1 To Me.Recordset.RecordCount, 1 To Me.Recordset.Fields.count)
    
    Me.Recordset.MoveFirst
    
    Do Until Me.Recordset.EOF
        row = row + 1
        
        For Each Field In Me.Recordset.Fields
            Column = Column + 1
            arr(row, Column) = Field.Value
        Next
        
        Me.Recordset.MoveNext
        Column = 0
    Loop
    ToArray = arr
End Function


Private Sub Class_Initialize()
    Call Me.Create
End Sub


Private Sub Class_Terminate()
    Call Me.Destroy
End Sub


Private Function GetExcelConnectionString() As String
    Dim Params As String
    Params = IIf(val(Application.Version) < 12, _
        "Provider='Microsoft.Jet.OLEDB.4.0';Data Source=':1';Extended Properties='Excel 4.0;HDR=:2;IMEX=1';", _
        "Provider='Microsoft.ACE.OLEDB.12.0';Data Source=':1';Extended Properties='Excel 12.0;HDR=:2;IMEX=1';")
    Params = Replace(Params, ":1", IIf(Me.DataSource = "", ThisWorkbook.FullName, Me.DataSource), , 1)
    Params = Replace(Params, ":2", IIf(Me.Header, "Yes", "No"), , 1)
    GetExcelConnectionString = Params
End Function





