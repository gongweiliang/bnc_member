
Sub Main()


GetDataFromDB


End Sub




'********************************
'�����ݿ��������õ�����
'********************************
Sub GetDataFromDB()
   dim ret
   Dim DataBase_DAO
   Dim rs_DAO 
   dim ora_connection,oraRet,dsql
'���ӷ�����oracle,��recordset
'****************************************
set DataBase_DAO = CreateObject("ADODB.connection")
ora_connection="Provider=OraOLEDB.Oracle.1;Password=buynowcan;Persist Security Info=True;User ID=asset_can_sales;Data Source=hjnew"
DataBase_DAO.open ora_connection
'����Ѿ������ѵ������ݣ�����ѵ������ݣ������ظ�

'*********************************************************************************
dsql="delete from gl.gl_interface where accounting_date=to_date('" & parameters("��ʼ����").value & "','yyyy-mm-dd')"

dsql=dsql & vblf & "and created_by=123"

 DataBase_DAO.Execute(DSQL)

dsql="commit"
 
DataBase_DAO.Execute(dsql) 
set oraRet=CreateObject("ADODB.recordset")
oraRet.locktype=2
oraRet.Open "select * from GL.GL_INTERFACE where status is null",DataBase_DAO
'*******************************************************************************************************************


dim DateFilter
  '�ж���������
  if Parameters("��ʼ����").Value = "" and Parameters("��������").Value = "" then
        DateFilter = " "
  elseif Parameters("��ʼ����").Value = "" and Parameters("��������").Value <> "" then
        DateFilter = " AND A.�տ����� <= TO_DATE('" & Parameters("��������").Value & "','yyyy-mm-dd') "  
  elseif Parameters("��ʼ����").Value <> "" and Parameters("��������").Value = "" then
        DateFilter = " AND A.�տ����� >= TO_DATE('" & Parameters("��ʼ����").Value & "','yyyy-mm-dd') "
  elseif Parameters("��ʼ����").Value <> "" and Parameters("��������").Value <> "" then
        DateFilter = " AND A.�տ����� BETWEEN TO_DATE('" & Parameters("��ʼ����").Value & "','yyyy-mm-dd')"
        DateFilter = DateFilter & " AND TO_DATE('" & Parameters("��������").Value & "','yyyy-mm-dd') "
  end if

  dim DumpFilter
  dim sSQL,sSQL1,sSQL2,SQL
  dim rs
  dim RcptID
  dim RowTxt
  dim UnitNo
  dim RcptYM
  dim RcptDate
  dim DueYear
  dim DueMonth
  dim date
  dim DueDateStr
  dim RType
  dim Dc
  dim Dec
  dim Dc1
  dim Dec1
  dim AcctID,AcctID1 
  dim JourType
  dim i,LiuNum,a,a1
  dim ItemName
  dim SourceName
'�����ƿ�Ŀ��
dim num1,num2,num3,num4,num5,num6,num7
'******************************************************
'�������ģ�飬����ϸ�跽�����ܴ�����һ��ƾ֤
'*******************************************************
'���Ƶ�������
'********************************
date=year(now)
if len(month(now))=1 then
   date=date & "0" & month(now)
else 
   date=date & month(now)
end if

if len(day(now))=1 then
    date=date & "0" & day(now)
else
    date=date & day(now)
end if



  'ȡ���������������������վݼ�¥�̵��������,�������(�跽��ϸ)
'***********************************************************************
  sSQL = "SELECT A.�վ�ID,A.���ݱ��,A.�տ�����,A.���,A.����������,A.����,"
  sSQL = sSQL & vblf & "A.����,A.���ҽ��,A.��������,A.�������,A.��ע,A.ժҪ,A.������,"
  sSQL = sSQL & vblf & "A.����,A.���,A.֧Ʊ����,A.�����տ�,A.���,A.�շ���Ŀ,A.����,A.����״̬,"
  sSQL = sSQL & vblf & "B.�Ϲ����,B.ҵ������,C.¥������,C.¥������,C.¥��,C.����,C.��Ԫ���,"
  sSQL = sSQL & vblf & "DECODE(C.¥������,NULL,' ',C.¥������) ¥������,"
  sSQL = sSQL & vblf & "DECODE(D.�Ƿ�����ת��,NULL,'��',D.�Ƿ�����ת��) �Ƿ�����ת��,E.��Ԫ"
  sSQL = sSQL & vblf & " FROM climb.V_�������_�վ���ϸ A,climb.���۹���_�Ϲ��� B,climb.V_��Ԫ C,climb.�Զ�������_�վ� D,climb.�Զ�������_��Ԫ E "
  sSQL = sSQL & vblf & "WHERE A.���� = 1 " & DateFilter
  sSQL = sSQL & vblf & "AND A.�Ϲ���ID = B.�Ϲ���ID "
  sSQL = sSQL & vblf & "AND B.��ԪID = C.��ԪID "
 sSQL = sSQL & vblf & "AND B.��ԪID = E.��ԪID "
  sSQL = sSQL & vblf & "AND A.�վ�ID = D.�վ�ID "
  sSQL = sSQL & vblf & "AND D.�Ƿ�����ת��='��' "
  sSQL = sSQL & vblf & "ORDER BY A.�տ�����,A.���ݱ��"
  Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
  'ȡ���������������������վݼ�¥�̵��������,�������(��������)
'***********************************************************************
  SQL = "SELECT nvl(sum(A.�����տ�),0) as �����տ�"
  SQL = SQL & vblf & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D"
 SQL = SQL & vblf & " WHERE A.���� = 1"  & DateFilter 
  SQL = SQL & vblf & " AND A.�վ�ID = D.�վ�ID "
  SQL = SQL & vblf & " AND D.�Ƿ�����ת��='��'" 
  Set rs1=Database.CreateDynaset(SQL,ORADYN_READONLY)
  '�����վ�ID
'***********************************************************************************************
  RcptID = ""     '�վ�ID����
  a=0             ' ƾ֤�ֱ���
if rs.eof=false then
   a=a+1      '����ƾ֤��
end if
           i=0
  While not rs.Eof
'	if RcptID <> CStr(rs("�վ�ID")) Then
           '��ʼ�����µ�һ���վ�
SourceName="�յ�" & rs("¥������")+rs("¥������")+rs("¥��")+rs("����")+rs("����������") & "�ĳ����"
'     *****************************

'�����ת�루�����ת��Ϊ������Ӧ�տ��Ŀ)���ʻ����������Ŀ����
   '****************************************************************************************************************************
          '�����վݵ��շ���Ŀ����ƾ֤����(���ô���)
num1="61"
num2="01"
num3="0000"
num4="2181"
num5="299"
num6="0000"
num7="0000"

'�����������ID


' ����ע

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ(�跽)
'***********************************************
oraRet.addnew

       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value        'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_dr")=CDbl(rs("���"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("���"))       '���ҽ�
       oraRet.fields("reference4")="�����ת����"
       oraRet.fields("reference10")=SourceName
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
       oraRet.update
      rs.movenext
wend
'End If
          '��ʼ������ϸ��¼
       '�����շ���Ŀ���ÿ�Ŀ����
'��Ŀ
num1="61"
num2="01"
num3="0000"
num4="2131"
num5="121"
num6="0000"
num7="0000"
'�����¼��

   '��ʼд����ϸ��¼(����)
   '*****************************
if rs1.eof=false and rs1("�����տ�")<>"0" then
oraRet.AddNew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       'oraRet.fields("entered_dr")=CDbl(rs("���"))          'ԭ�ҽ�
      ' oraRet.fields("accounted_dr")=CDBL(rs("���"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�����ת����"
       oraRet.fields("reference10")="�����ת����"
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)

      oraRet.update
end if
        '������һ����¼       


  'ȡ���������������������վݼ�¥�̵��������,���ֽ�ǳ������(�跽��������ϸ,��������Ŀ����)
'***********************************************************************
'�跽��ѯ����
'*****************************************************************************************
  sSQL = "SELECT A.�վ�ID,A.���ݱ��,A.�տ�����,A.���,A.����������,A.����,A.��������,"
  sSQL = sSQL & vblf & "A.����,A.���ҽ��,A.��������,A.�������,A.��ע,A.ժҪ,A.������,"
  sSQL = sSQL & vblf & "A.����,A.���,A.֧Ʊ����,A.�����տ�,A.���,A.�շ���Ŀ,A.����,A.����״̬,"
  sSQL = sSQL & vblf & "decode(��������,'�㷢����',1,'�㷢����(Y)',2,'��������·֧��',3,4) as kk,"
  sSQL = sSQL & vblf & "B.�Ϲ����,B.ҵ������,C.¥������,C.¥������,C.¥��,C.����,C.��Ԫ���,"
  sSQL = sSQL & vblf & "DECODE(C.¥������,NULL,' ',C.¥������) ¥������,"
  sSQL = sSQL & vblf & "DECODE(D.�Ƿ�����ת��,NULL,'��',D.�Ƿ�����ת��) �Ƿ�����ת��"
  sSQL = sSQL & vblf & " FROM climb.V_�������_�վ���ϸ A,climb.���۹���_�Ϲ��� B,climb.V_��Ԫ C,�Զ�������_�վ� D "
  sSQL = sSQL & vblf & "WHERE A.���� = 1 " & DateFilter & DumpFilter
  sSQL = sSQL & vblf & "AND A.�Ϲ���ID = B.�Ϲ���ID "
  sSQL = sSQL & vblf & "AND B.��ԪID = C.��ԪID "
  sSQL = sSQL & vblf & "AND A.�վ�ID = D.�վ�ID "
  sSQL = sSQL & vblf & "AND A.��������<>'�ֽ�' and A.��������<>'POS��' and (D.�Ƿ�����ת�� is null or D.�Ƿ�����ת��<>'��')"
  sSQL = sSQL & vblf & "ORDER BY kk,A.��������,A.���ݱ��,A.�շ���Ŀ,A.�տ�����"
  Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
'������ѯ����
'******************************************************************
SQL="SELECT sum(A.�����տ�) as �����տ�,A.�շ���Ŀ,A.��������,"
SQL = SQL & vblf & "decode(��������,'�㷢����',1,'�㷢����(Y)',2,'��������·֧��',3,4) as kk"
SQL=SQL & vblf & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D" 
SQL=SQL & vblf & " WHERE A.���� = 1" & DateFilter & DumpFilter   
SQL=SQL & vblf & " AND A.�վ�ID = D.�վ�ID AND A.��������<>'�ֽ�' and A.��������<>'POS��'" 
SQL=SQL & vblf & " and (D.�Ƿ�����ת�� is null or D.�Ƿ�����ת��<>'��')"
SQL=SQL & vblf & " group by A.��������,A.�շ���Ŀ"
SQL=SQL & vblf & " ORDER BY kk,A.��������,A.�շ���Ŀ"
set rs1=Database.CreateDynaset(SQL,ORADYN_READONLY)
'�跽����,�����з�ƾ֤:�㷢����,����ϸ��Ŀ
  'ÿ���վݿ����ж�ʷ�¼����RcptID��¼��ǰ���ڴ�����վ�ID�����վ�ID�Ǳ�ʶÿ���վݵ�Ψһ��
'***********************************************************************************************
dim rowID1,rowID2,rowID3
rowID1=1
rowID2=1
'�㷢����
          ' ƾ֤�ֱ���
  a1=2          '����Ӧ�տ�����
if rs.eof=false  and  rs("��������")="�㷢����" then
  a=a+1  ' ƾ֤�ֱ���
end if
   i=0

  While not rs.Eof and rs("��������")="�㷢����"
           '��ʼ�����µ�һ���վ�
'     *****************************
           a1=a1+1
           RcptID = CStr(rs("�վ�ID"))
           SourceName="�յ����ݱ��" & rs("���ݱ��") & "," & rs("¥������")+rs("¥������")+rs("¥��")+rs("����")+rs("����������")




           '���ݸ�������(�ֽ�/����)�Լ����ֻ�����,���Ƿ�����ת�루�����ת��Ϊ������Ӧ�տ��Ŀ)���ʻ����������Ŀ����
   '****************************************************************************************************************************
           if CStr(rs("��������")) = "�㷢����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"
           elseif CStr(rs("��������")) = "��������·֧��" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '�����վݵ��շ���Ŀ����ƾ֤����(���ô���)

       ItemName="����"

 
' ����ע
    ItemName=SourceName+rs("��������")+rs("�շ���Ŀ")+ItemName

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
'***************************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
     '  oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
      ' oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�㷢�����տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
         oraRet.update
i=i+1
if rs.eof=false then
rowID1=rs.rowposition
end if
       rs.movenext

Wend	

       
        '��ʼ������ϸ��¼


'�����¼��

   '��ʼд����ϸ��¼
   '*****************************
while not rs1.eof and rs1("��������")="�㷢����"
       '�����շ���Ŀ���ÿ�Ŀ����
           if CStr(rs1("�շ���Ŀ")) = "����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ҷ���" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "����" or CStr(rs1("�շ���Ŀ"))="���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "ӡ��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "������" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num1="61"
              num2="01"
              num3="0000"
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if
    ItemName=rs1("�շ���Ŀ")
    ItemName=ItemName & "����"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
     '  oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
      ' oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�㷢�����տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)

     oraRet.update
        '������һ����¼    
i=i+1    
if rs1.eof=false then
rowID2=rs1.rowposition
end if
         rs1.MoveNext
 Wend	

'�㷢����(Y)
  RcptID = ""     '�վ�ID����          
if rs.eof =false and CDBL(rowID1)<>1 then
rs.moveto CDBL(rowID1)+1
end if

if rs.eof=false  and  rs("��������")="�㷢����(Y)" then
  a=a+1  ' ƾ֤�ֱ���
end if
   i=0
  While not rs.Eof and rs("��������")="�㷢����(Y)"
           '��ʼ�����µ�һ���վ�
'     *****************************
           SourceName="�յ����ݱ��" & rs("���ݱ��") & "," & rs("¥������")+rs("¥������")+rs("¥��")+rs("����")+rs("����������")


           '���ݸ�������(�ֽ�/����)�Լ����ֻ�����,���Ƿ�����ת�루�����ת��Ϊ������Ӧ�տ��Ŀ)���ʻ����������Ŀ����
   '****************************************************************************************************************************
           if CStr(rs("��������")) = "�㷢����(Y)" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="215"
              num6="0000"
              num7="0000"
          elseif CStr(rs("��������")) = "��������·֧��" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '�����վݵ��շ���Ŀ����ƾ֤����(���ô���)

       ItemName="����"

 
' ����ע
    ItemName=SourceName+rs("��������")+rs("�շ���Ŀ")+ItemName

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
'************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
     '  oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
      ' oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�㷢����(Y)�տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
oraRet.update
i=i+1
if rs.eof=false then
rowID1=rs.rowposition
end if
       rs.movenext

Wend	


         
 '��ʼ������ϸ��¼




'�����¼��

   '��ʼд����ϸ��¼(����)
   '*****************************
if rs1.eof =false and CDBL(rowID2)<>1 then
rs1.moveto CDBL(rowID2)+1
end if
while not rs1.eof and rs1("��������")="�㷢����(Y)"
       '�����շ���Ŀ���ÿ�Ŀ����
           if CStr(rs1("�շ���Ŀ")) = "����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ҷ���" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "����" or CStr(rs1("�շ���Ŀ"))="���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "ӡ��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "������" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num1="61"
              num2="01"
              num3="0000"
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if

    ItemName=rs1("�շ���Ŀ")
    ItemName=ItemName & "����"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
     '  oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
      ' oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�㷢����(Y)�տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
      oraRet.update
        '������һ����¼       
i=i+1 
if rs1.eof=false then
rowID2=rs1.rowposition
end if
         rs1.MoveNext
 Wend	



'����
          ' ƾ֤�ֱ���
  a1=2          '����Ӧ�տ�����
if rs.eof=false  and  rs("��������")="��������·֧��" then
  a=a+1  ' ƾ֤�ֱ���
end if
   i=0
if rs.eof =false and CDBL(rowID1)<>1 then
rs.moveto rowID1+1
end if
  While not rs.Eof and rs("��������")="��������·֧��"
           '��ʼ�����µ�һ���վ�
'     *****************************
           SourceName="�յ����ݱ��" & rs("���ݱ��") & "," & rs("¥������")+rs("¥������")+rs("¥��")+rs("����")+rs("����������")
           '��������
    '***************************************

           '���ݸ�������(�ֽ�/����)�Լ����ֻ�����,���Ƿ�����ת�루�����ת��Ϊ������Ӧ�տ��Ŀ)���ʻ����������Ŀ����
   '****************************************************************************************************************************
           if CStr(rs("��������")) = "�㷢����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"
           elseif CStr(rs("��������")) = "��������·֧��" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '�����վݵ��շ���Ŀ����ƾ֤����(���ô���)

       ItemName="����"
 
' ����ע
    ItemName=SourceName+rs("��������")+rs("�շ���Ŀ")+ItemName

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
'******************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
      ' oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�����տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
       oraRet.update
i=i+1
       rs.movenext

Wend	

          '��ʼ������ϸ��¼




'�����¼��

   '��ʼд����ϸ��¼
   '*****************************
if rs1.eof =false and CDBL(rowID2)<>1 then
rs1.moveto rowID2+1
end if
while not rs1.eof and rs1("��������")="��������·֧��"
       '�����շ���Ŀ���ÿ�Ŀ����
            if CStr(rs1("�շ���Ŀ")) = "����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "���ҷ���" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "����" or CStr(rs1("�շ���Ŀ"))="���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "ӡ��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "������" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num1="61"
              num2="01"
              num3="0000"
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if
    ItemName=rs1("�շ���Ŀ")
    ItemName=ItemName & "����"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       'oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
       'oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�����տ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=ItemName                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '������һ����¼    
i=i+1    
         rs1.MoveNext
 Wend	



  'ȡ���������������������վݼ�¥�̵��������,�ֽ���,(Ĭ���ڹ㷢����)
'*******************************************************
sSQL="SELECT sum(A.�����տ�) as ���,A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D,"
sSQL=sSQL & " (SELECT sum(A.�����տ�) as �ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D" 
sSQL=sSQL & " WHERE A.���� = 1" & DateFilter & DumpFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='�ֽ�') B"
sSQL=sSQL & " WHERE A.���� = 1"  & DateFilter & DumpFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='�ֽ�'"
sSQL=sSQL & " group by A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " order by �շ���Ŀ"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 a=a              'ƾ֤��

'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("�ܽ��")) Then
 '          '��ʼ�����µ�һ���վ�
           a=a+1
          RcptID = CStr(rs("�ܽ��"))
           i=0
  '���ݱ��δ�����վ����տ���˿�������վݽ�����跽���Ǵ���
'**********************************************************************

        '���ݣ��ֽ�������У����������Ŀ����(����)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"

'�����ݱ�ţ���ע��Ҫ��ϸ���ݱ��
'************************************************************
sSQL2="select distinct ���ݱ��"
sSQL2=sSQL2 & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D"
sSQL2=sSQL2 & " where A.���� = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.�վ�ID = D.�վ�ID"
sSQL2=sSQL2 & " AND A.��������='�ֽ�'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="�ֽ𵥾ݱ��Ϊ:"
while not rs1.eof
    ItemName=ItemName+rs1("���ݱ��") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "ת�뽨������֧��"

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
'************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
      ' oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�ܽ��"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�ܽ��"))       '���ҽ�
       'oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="�ֽ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")="�ֽ�ת�뽨������֧��"              '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
  '******************************************************************
   End If
          '��ʼ������ϸ��¼
       '�����շ���Ŀ���ÿ�Ŀ����
            if CStr(rs("�շ���Ŀ")) = "����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ҷ���" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "����" or CStr(rs("�շ���Ŀ"))="���ڿ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "ӡ��˰" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "������" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num1="61"
              num2="01"
              num3="0000"
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if
    i=i+1
Name=rs("�շ���Ŀ") & "�ֽ�ת��"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("���"))          'ԭ�Ҵ�
       'oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
       'oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs("���"))       '���Ҵ�
       oraRet.fields("reference4")="�ֽ�ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=Name                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '������һ����¼        
         rs.MoveNext
wend
'posY������(Ĭ�Ϲ㷢����)
  'ȡ���������������������վݼ�¥�̵��������,�ֽ���,(Ĭ���ڹ㷢����)
'*******************************************************
sSQL="SELECT sum(A.�����տ�) as ���,A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D,"
sSQL=sSQL & " (SELECT sum(A.�����տ�) as �ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D" 
sSQL=sSQL & " WHERE A.���� = 1" & DateFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='POS��(Y)') B"
sSQL=sSQL & " WHERE A.���� = 1"  & DateFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='POS��(Y)'"
sSQL=sSQL & " group by A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " order by �շ���Ŀ"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 RcptID = " "     '�ֽ�������б���
 a=a              'ƾ֤��
'����ƾ֤����,��Ϊ�ǻ���,û����ȷ���տ�����,�������������ʼ����Ϊ׼

'��ʼ�����¼
'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("�ܽ��")) Then
 '          '��ʼ�����µ�һ���վ�
           a=a+1
          RcptID = CStr(rs("�ܽ��"))
           i=0
  '���ݱ��δ�����վ����տ���˿�������վݽ�����跽���Ǵ���
'**********************************************************************

        '���ݣ��ֽ�������У����������Ŀ����(�㷢����)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="215"
              num6="0000"
              num7="0000"


'�����ݱ�ţ���ע��Ҫ��ϸ���ݱ��
'************************************************************
sSQL2="select distinct ���ݱ��"
sSQL2=sSQL2 & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D"
sSQL2=sSQL2 & " where A.���� = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.�վ�ID = D.�վ�ID"
sSQL2=sSQL2 & " AND A.��������='POS��(Y)'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="pos���տ�ݱ��Ϊ:"
while not rs1.eof
    ItemName=ItemName+rs1("���ݱ��") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "ת��㷢����(Y)"
'�����վݵ��շ���Ŀ�Լ�¥�����ƾ���ƾ֤����(���ô���)
'*********************************************************************

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
  '****************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�ܽ��"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�ܽ��"))       '���ҽ�
       'oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="pos��ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")="POS��ת��㷢����(Y)"             '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
   End If
          '��ʼ������ϸ��¼
       '�����շ���Ŀ���ÿ�Ŀ����
              num1="61"
              num2="01"
              num3="0000"
           if CStr(rs("�շ���Ŀ")) = "����" then
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ڿ�" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ҷ���" then
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "����" or CStr(rs("�շ���Ŀ"))="���ڿ�" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��˰" then
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "ӡ��˰" then
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "������" then
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if
    i=i+1
Name=rs("�շ���Ŀ") & "pos��ת��"
'**********************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("���"))          'ԭ�Ҵ�
       'oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
      ' oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs("���"))       '���Ҵ�
       oraRet.fields("reference4")="pos��ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=Name                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '������һ����¼        
         rs.MoveNext
wend
'Pos��
'pos������(Ĭ�Ϲ㷢����)
  'ȡ���������������������վݼ�¥�̵��������,�ֽ���,(Ĭ���ڹ㷢����)
'*******************************************************
sSQL="SELECT sum(A.�����տ�) as ���,A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D,"
sSQL=sSQL & " (SELECT sum(A.�����տ�) as �ܽ��"
sSQL=sSQL & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D" 
sSQL=sSQL & " WHERE A.���� = 1" & DateFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='POS��') B"
sSQL=sSQL & " WHERE A.���� = 1"  & DateFilter
sSQL=sSQL & " AND A.�վ�ID = D.�վ�ID" 
sSQL=sSQL & " AND A.��������='POS��'"
sSQL=sSQL & " group by A.�շ���Ŀ,B.�ܽ��"
sSQL=sSQL & " order by �շ���Ŀ"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 RcptID = " "     '�ֽ�������б���
 a=a              'ƾ֤��
'����ƾ֤����,��Ϊ�ǻ���,û����ȷ���տ�����,�������������ʼ����Ϊ׼

'��ʼ�����¼
'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("�ܽ��")) Then
 '          '��ʼ�����µ�һ���վ�
           a=a+1
          RcptID = CStr(rs("�ܽ��"))
           i=0
  '���ݱ��δ�����վ����տ���˿�������վݽ�����跽���Ǵ���
'**********************************************************************

        '���ݣ��ֽ�������У����������Ŀ����(�㷢����)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"


'�����ݱ�ţ���ע��Ҫ��ϸ���ݱ��
'************************************************************
sSQL2="select distinct ���ݱ��"
sSQL2=sSQL2 & " FROM climb.V_�������_�վ���ϸ A,climb.�Զ�������_�վ� D"
sSQL2=sSQL2 & " where A.���� = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.�վ�ID = D.�վ�ID"
sSQL2=sSQL2 & " AND A.��������='POS��'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="pos���տ�ݱ��Ϊ:"
while not rs1.eof
    ItemName=ItemName+rs1("���ݱ��") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "ת��㷢����(Y)"
'�����վݵ��շ���Ŀ�Լ�¥�����ƾ���ƾ֤����(���ô���)
'*********************************************************************

           'RowTxt ��¼��д���м��ļ���ÿ����Ϣ
  '****************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("�����տ�"))          'ԭ�Ҵ�
       oraRet.fields("entered_dr")=CDbl(rs("�ܽ��"))          'ԭ�ҽ�
       oraRet.fields("accounted_dr")=CDBL(rs("�ܽ��"))       '���ҽ�
       'oraRet.fields("accounted_cr")=CDBL(rs1("�����տ�"))       '���Ҵ�
       oraRet.fields("reference4")="pos��ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")="POS��ת��㷢����(Y)"             '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
   End If
          '��ʼ������ϸ��¼
       '�����շ���Ŀ���ÿ�Ŀ����
              num1="61"
              num2="01"
              num3="0000"
           if CStr(rs("�շ���Ŀ")) = "����" then
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ڿ�" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "���ҷ���" then
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "����" or CStr(rs("�շ���Ŀ"))="���ڿ�" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��˰" then
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "ӡ��˰" then
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "������" then
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ�ǼǷ�" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("�շ���Ŀ")) = "��Ȩ֤ӡ����" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           else
              num4="1133"
              num5="299"
              num6="6101"
              num7="0001"
           end if
    i=i+1
Name=rs("�շ���Ŀ") & "pos��ת��"
'**********************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("��ʼ����").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("��ʼ����").value
       oraRet.fields("created_by")=parameters("�û�ID").value         'From Oracle User_id��ORACLE�û�ID��
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="��������"
       oraRet.fields("user_je_source_name")="����ϵͳ_����"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("���"))          'ԭ�Ҵ�
       'oraRet.fields("entered_dr")=CDbl(rs("�����տ�"))          'ԭ�ҽ�
      ' oraRet.fields("accounted_dr")=CDBL(rs("�����տ�"))       '���ҽ�
       oraRet.fields("accounted_cr")=CDBL(rs("���"))       '���Ҵ�
       oraRet.fields("reference4")="pos��ƾ֤"                'ƾ֤����
       oraRet.fields("reference10")=Name                '��ϸժҪ
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '������һ����¼        
         rs.MoveNext
wend
oraRet.close	
 msgbox("�����ɹ�����")	
end sub