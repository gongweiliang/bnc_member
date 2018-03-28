
Sub Main()


GetDataFromDB


End Sub




'********************************
'从数据库生成适用的数据
'********************************
Sub GetDataFromDB()
   dim ret
   Dim DataBase_DAO
   Dim rs_DAO 
   dim ora_connection,oraRet,dsql
'连接服务器oracle,打开recordset
'****************************************
set DataBase_DAO = CreateObject("ADODB.connection")
ora_connection="Provider=OraOLEDB.Oracle.1;Password=buynowcan;Persist Security Info=True;User ID=asset_can_sales;Data Source=hjnew"
DataBase_DAO.open ora_connection
'如果已经当期已导入数据，清楚已导入数据，避免重复

'*********************************************************************************
dsql="delete from gl.gl_interface where accounting_date=to_date('" & parameters("起始日期").value & "','yyyy-mm-dd')"

dsql=dsql & vblf & "and created_by=123"

 DataBase_DAO.Execute(DSQL)

dsql="commit"
 
DataBase_DAO.Execute(dsql) 
set oraRet=CreateObject("ADODB.recordset")
oraRet.locktype=2
oraRet.Open "select * from GL.GL_INTERFACE where status is null",DataBase_DAO
'*******************************************************************************************************************


dim DateFilter
  '判断日期条件
  if Parameters("起始日期").Value = "" and Parameters("截至日期").Value = "" then
        DateFilter = " "
  elseif Parameters("起始日期").Value = "" and Parameters("截至日期").Value <> "" then
        DateFilter = " AND A.收款日期 <= TO_DATE('" & Parameters("截至日期").Value & "','yyyy-mm-dd') "  
  elseif Parameters("起始日期").Value <> "" and Parameters("截至日期").Value = "" then
        DateFilter = " AND A.收款日期 >= TO_DATE('" & Parameters("起始日期").Value & "','yyyy-mm-dd') "
  elseif Parameters("起始日期").Value <> "" and Parameters("截至日期").Value <> "" then
        DateFilter = " AND A.收款日期 BETWEEN TO_DATE('" & Parameters("起始日期").Value & "','yyyy-mm-dd')"
        DateFilter = DateFilter & " AND TO_DATE('" & Parameters("截至日期").Value & "','yyyy-mm-dd') "
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
'定义会计科目段
dim num1,num2,num3,num4,num5,num6,num7
'******************************************************
'诚意金处理模块，按明细借方，汇总贷方成一张凭证
'*******************************************************
'定制当天日期
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



  '取出符合日期条件的所有收据及楼盘的相关资料,诚意金处理(借方明细)
'***********************************************************************
  sSQL = "SELECT A.收据ID,A.单据编号,A.收款日期,A.金额,A.交款人姓名,A.币种,"
  sSQL = sSQL & vblf & "A.汇率,A.本币金额,A.付款类型,A.单据类别,A.备注,A.摘要,A.经办人,"
  sSQL = sSQL & vblf & "A.出纳,A.会计,A.支票号码,A.本次收款,A.类别,A.收费项目,A.期数,A.导出状态,"
  sSQL = sSQL & vblf & "B.认购书号,B.业主姓名,C.楼盘名称,C.楼阁名称,C.楼层,C.房号,C.单元编号,"
  sSQL = sSQL & vblf & "DECODE(C.楼梯名称,NULL,' ',C.楼梯名称) 楼梯名称,"
  sSQL = sSQL & vblf & "DECODE(D.是否诚意金转入,NULL,'否',D.是否诚意金转入) 是否诚意金转入,E.单元"
  sSQL = sSQL & vblf & " FROM climb.V_财务管理_收据明细 A,climb.销售管理_认购书 B,climb.V_单元 C,climb.自定义内容_收据 D,climb.自定义内容_单元 E "
  sSQL = sSQL & vblf & "WHERE A.记帐 = 1 " & DateFilter
  sSQL = sSQL & vblf & "AND A.认购书ID = B.认购书ID "
  sSQL = sSQL & vblf & "AND B.单元ID = C.单元ID "
 sSQL = sSQL & vblf & "AND B.单元ID = E.单元ID "
  sSQL = sSQL & vblf & "AND A.收据ID = D.收据ID "
  sSQL = sSQL & vblf & "AND D.是否诚意金转入='是' "
  sSQL = sSQL & vblf & "ORDER BY A.收款日期,A.单据编号"
  Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
  '取出符合日期条件的所有收据及楼盘的相关资料,诚意金处理(贷方汇总)
'***********************************************************************
  SQL = "SELECT nvl(sum(A.本次收款),0) as 本次收款"
  SQL = SQL & vblf & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D"
 SQL = SQL & vblf & " WHERE A.记帐 = 1"  & DateFilter 
  SQL = SQL & vblf & " AND A.收据ID = D.收据ID "
  SQL = SQL & vblf & " AND D.是否诚意金转入='是'" 
  Set rs1=Database.CreateDynaset(SQL,ORADYN_READONLY)
  '处理收据ID
'***********************************************************************************************
  RcptID = ""     '收据ID变量
  a=0             ' 凭证字变量
if rs.eof=false then
   a=a+1      '处理凭证号
end if
           i=0
  While not rs.Eof
'	if RcptID <> CStr(rs("收据ID")) Then
           '开始处理新的一张收据
SourceName="收到" & rs("楼阁名称")+rs("楼梯名称")+rs("楼层")+rs("房号")+rs("交款人姓名") & "的诚意金"
'     *****************************

'诚意金转入（诚意金转入为“其他应收款”科目)的帐户决定财务科目编码
   '****************************************************************************************************************************
          '根据收据的收费项目决定凭证类型(费用代码)
num1="61"
num2="01"
num3="0000"
num4="2181"
num5="299"
num6="0000"
num7="0000"

'处理诚意金核算ID


' 处理备注

           'RowTxt 记录将写入中间文件的每条信息(借方)
'***********************************************
oraRet.addnew

       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value        'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_dr")=CDbl(rs("金额"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("金额"))       '本币借
       oraRet.fields("reference4")="诚意金转定金"
       oraRet.fields("reference10")=SourceName
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
       oraRet.update
      rs.movenext
wend
'End If
          '开始处理明细分录
       '根据收费项目设置科目编码
'科目
num1="61"
num2="01"
num3="0000"
num4="2131"
num5="121"
num6="0000"
num7="0000"
'处理记录号

   '开始写入明细分录(贷方)
   '*****************************
if rs1.eof=false and rs1("本次收款")<>"0" then
oraRet.AddNew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       'oraRet.fields("entered_dr")=CDbl(rs("金额"))          '原币借
      ' oraRet.fields("accounted_dr")=CDBL(rs("金额"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="诚意金转定金"
       oraRet.fields("reference10")="诚意金转定金"
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)

      oraRet.update
end if
        '处理下一条记录       


  '取出符合日期条件的所有收据及楼盘的相关资料,非现金非诚意金处理(借方按银行明细,贷方按项目汇总)
'***********************************************************************
'借方查询代码
'*****************************************************************************************
  sSQL = "SELECT A.收据ID,A.单据编号,A.收款日期,A.金额,A.交款人姓名,A.币种,A.入帐银行,"
  sSQL = sSQL & vblf & "A.汇率,A.本币金额,A.付款类型,A.单据类别,A.备注,A.摘要,A.经办人,"
  sSQL = sSQL & vblf & "A.出纳,A.会计,A.支票号码,A.本次收款,A.类别,A.收费项目,A.期数,A.导出状态,"
  sSQL = sSQL & vblf & "decode(付款类型,'广发总行',1,'广发总行(Y)',2,'建行龙口路支行',3,4) as kk,"
  sSQL = sSQL & vblf & "B.认购书号,B.业主姓名,C.楼盘名称,C.楼阁名称,C.楼层,C.房号,C.单元编号,"
  sSQL = sSQL & vblf & "DECODE(C.楼梯名称,NULL,' ',C.楼梯名称) 楼梯名称,"
  sSQL = sSQL & vblf & "DECODE(D.是否诚意金转入,NULL,'否',D.是否诚意金转入) 是否诚意金转入"
  sSQL = sSQL & vblf & " FROM climb.V_财务管理_收据明细 A,climb.销售管理_认购书 B,climb.V_单元 C,自定义内容_收据 D "
  sSQL = sSQL & vblf & "WHERE A.记帐 = 1 " & DateFilter & DumpFilter
  sSQL = sSQL & vblf & "AND A.认购书ID = B.认购书ID "
  sSQL = sSQL & vblf & "AND B.单元ID = C.单元ID "
  sSQL = sSQL & vblf & "AND A.收据ID = D.收据ID "
  sSQL = sSQL & vblf & "AND A.付款类型<>'现金' and A.付款类型<>'POS机' and (D.是否诚意金转入 is null or D.是否诚意金转入<>'是')"
  sSQL = sSQL & vblf & "ORDER BY kk,A.付款类型,A.单据编号,A.收费项目,A.收款日期"
  Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
'贷方查询代码
'******************************************************************
SQL="SELECT sum(A.本次收款) as 本次收款,A.收费项目,A.付款类型,"
SQL = SQL & vblf & "decode(付款类型,'广发总行',1,'广发总行(Y)',2,'建行龙口路支行',3,4) as kk"
SQL=SQL & vblf & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D" 
SQL=SQL & vblf & " WHERE A.记帐 = 1" & DateFilter & DumpFilter   
SQL=SQL & vblf & " AND A.收据ID = D.收据ID AND A.付款类型<>'现金' and A.付款类型<>'POS机'" 
SQL=SQL & vblf & " and (D.是否诚意金转入 is null or D.是否诚意金转入<>'是')"
SQL=SQL & vblf & " group by A.付款类型,A.收费项目"
SQL=SQL & vblf & " ORDER BY kk,A.付款类型,A.收费项目"
set rs1=Database.CreateDynaset(SQL,ORADYN_READONLY)
'借方处理,按银行分凭证:广发银行,并明细项目
  '每张收据可能有多笔分录，用RcptID记录当前正在处理的收据ID——收据ID是标识每张收据的唯一号
'***********************************************************************************************
dim rowID1,rowID2,rowID3
rowID1=1
rowID2=1
'广发银行
          ' 凭证字变量
  a1=2          '其他应收款核算号
if rs.eof=false  and  rs("付款类型")="广发总行" then
  a=a+1  ' 凭证字变量
end if
   i=0

  While not rs.Eof and rs("付款类型")="广发总行"
           '开始处理新的一张收据
'     *****************************
           a1=a1+1
           RcptID = CStr(rs("收据ID"))
           SourceName="收到单据编号" & rs("单据编号") & "," & rs("楼阁名称")+rs("楼梯名称")+rs("楼层")+rs("房号")+rs("交款人姓名")




           '根据付款类型(现金/银行)以及币种或银行,或是否诚意金转入（诚意金转入为“其他应收款”科目)的帐户决定财务科目编码
   '****************************************************************************************************************************
           if CStr(rs("付款类型")) = "广发总行" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"
           elseif CStr(rs("付款类型")) = "建行龙口路支行" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '根据收据的收费项目决定凭证类型(费用代码)

       ItemName="交费"

 
' 处理备注
    ItemName=SourceName+rs("付款类型")+rs("收费项目")+ItemName

           'RowTxt 记录将写入中间文件的每条信息
'***************************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
     '  oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
      ' oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="广发银行收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
         oraRet.update
i=i+1
if rs.eof=false then
rowID1=rs.rowposition
end if
       rs.movenext

Wend	

       
        '开始处理明细分录


'处理记录号

   '开始写入明细分录
   '*****************************
while not rs1.eof and rs1("付款类型")="广发总行"
       '根据收费项目设置科目编码
           if CStr(rs1("收费项目")) = "定金" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "首期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "按揭房款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "房款" or CStr(rs1("收费项目"))="分期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "契税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "印花税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "工本费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权登记费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权证印花费" then
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
    ItemName=rs1("收费项目")
    ItemName=ItemName & "交款"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
     '  oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
      ' oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="广发银行收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)

     oraRet.update
        '处理下一条记录    
i=i+1    
if rs1.eof=false then
rowID2=rs1.rowposition
end if
         rs1.MoveNext
 Wend	

'广发总行(Y)
  RcptID = ""     '收据ID变量          
if rs.eof =false and CDBL(rowID1)<>1 then
rs.moveto CDBL(rowID1)+1
end if

if rs.eof=false  and  rs("付款类型")="广发总行(Y)" then
  a=a+1  ' 凭证字变量
end if
   i=0
  While not rs.Eof and rs("付款类型")="广发总行(Y)"
           '开始处理新的一张收据
'     *****************************
           SourceName="收到单据编号" & rs("单据编号") & "," & rs("楼阁名称")+rs("楼梯名称")+rs("楼层")+rs("房号")+rs("交款人姓名")


           '根据付款类型(现金/银行)以及币种或银行,或是否诚意金转入（诚意金转入为“其他应收款”科目)的帐户决定财务科目编码
   '****************************************************************************************************************************
           if CStr(rs("付款类型")) = "广发总行(Y)" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="215"
              num6="0000"
              num7="0000"
          elseif CStr(rs("付款类型")) = "建行龙口路支行" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '根据收据的收费项目决定凭证类型(费用代码)

       ItemName="交费"

 
' 处理备注
    ItemName=SourceName+rs("付款类型")+rs("收费项目")+ItemName

           'RowTxt 记录将写入中间文件的每条信息
'************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
     '  oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
      ' oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="广发总行(Y)收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
oraRet.update
i=i+1
if rs.eof=false then
rowID1=rs.rowposition
end if
       rs.movenext

Wend	


         
 '开始处理明细分录




'处理记录号

   '开始写入明细分录(贷方)
   '*****************************
if rs1.eof =false and CDBL(rowID2)<>1 then
rs1.moveto CDBL(rowID2)+1
end if
while not rs1.eof and rs1("付款类型")="广发总行(Y)"
       '根据收费项目设置科目编码
           if CStr(rs1("收费项目")) = "定金" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "首期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "按揭房款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "房款" or CStr(rs1("收费项目"))="分期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "契税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "印花税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "工本费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权登记费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权证印花费" then
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

    ItemName=rs1("收费项目")
    ItemName=ItemName & "交款"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
     '  oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
      ' oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="广发总行(Y)收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
      oraRet.update
        '处理下一条记录       
i=i+1 
if rs1.eof=false then
rowID2=rs1.rowposition
end if
         rs1.MoveNext
 Wend	



'建行
          ' 凭证字变量
  a1=2          '其他应收款核算号
if rs.eof=false  and  rs("付款类型")="建行龙口路支行" then
  a=a+1  ' 凭证字变量
end if
   i=0
if rs.eof =false and CDBL(rowID1)<>1 then
rs.moveto rowID1+1
end if
  While not rs.Eof and rs("付款类型")="建行龙口路支行"
           '开始处理新的一张收据
'     *****************************
           SourceName="收到单据编号" & rs("单据编号") & "," & rs("楼阁名称")+rs("楼梯名称")+rs("楼层")+rs("房号")+rs("交款人姓名")
           '处理日期
    '***************************************

           '根据付款类型(现金/银行)以及币种或银行,或是否诚意金转入（诚意金转入为“其他应收款”科目)的帐户决定财务科目编码
   '****************************************************************************************************************************
           if CStr(rs("付款类型")) = "广发银行" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"
           elseif CStr(rs("付款类型")) = "建行龙口路支行" then
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"
           end if

           '根据收据的收费项目决定凭证类型(费用代码)

       ItemName="交费"
 
' 处理备注
    ItemName=SourceName+rs("付款类型")+rs("收费项目")+ItemName

           'RowTxt 记录将写入中间文件的每条信息
'******************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
      ' oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="建行收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
       oraRet.update
i=i+1
       rs.movenext

Wend	

          '开始处理明细分录




'处理记录号

   '开始写入明细分录
   '*****************************
if rs1.eof =false and CDBL(rowID2)<>1 then
rs1.moveto rowID2+1
end if
while not rs1.eof and rs1("付款类型")="建行龙口路支行"
       '根据收费项目设置科目编码
            if CStr(rs1("收费项目")) = "定金" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "首期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "按揭房款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "房款" or CStr(rs1("收费项目"))="分期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "契税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "印花税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "工本费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权登记费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs1("收费项目")) = "产权证印花费" then
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
    ItemName=rs1("收费项目")
    ItemName=ItemName & "交款"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       'oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
       'oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="建行收款凭证"                '凭证名称
       oraRet.fields("reference10")=ItemName                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '处理下一条记录    
i=i+1    
         rs1.MoveNext
 Wend	



  '取出符合日期条件的所有收据及楼盘的相关资料,现金处理,(默认在广发银行)
'*******************************************************
sSQL="SELECT sum(A.本次收款) as 金额,A.收费项目,B.总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D,"
sSQL=sSQL & " (SELECT sum(A.本次收款) as 总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D" 
sSQL=sSQL & " WHERE A.记帐 = 1" & DateFilter & DumpFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='现金') B"
sSQL=sSQL & " WHERE A.记帐 = 1"  & DateFilter & DumpFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='现金'"
sSQL=sSQL & " group by A.收费项目,B.总金额"
sSQL=sSQL & " order by 收费项目"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 a=a              '凭证号

'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("总金额")) Then
 '          '开始处理新的一张收据
           a=a+1
          RcptID = CStr(rs("总金额"))
           i=0
  '根据本次处理的收据是收款或退款，决定将收据金额记入借方还是贷方
'**********************************************************************

        '根据（现金存入银行）决定财务科目编码(建行)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="211"
              num6="0000"
              num7="0000"

'处理单据编号，备注需要明细单据编号
'************************************************************
sSQL2="select distinct 单据编号"
sSQL2=sSQL2 & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D"
sSQL2=sSQL2 & " where A.记帐 = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.收据ID = D.收据ID"
sSQL2=sSQL2 & " AND A.付款类型='现金'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="现金单据编号为:"
while not rs1.eof
    ItemName=ItemName+rs1("单据编号") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "转入建行龙口支行"

           'RowTxt 记录将写入中间文件的每条信息
'************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
      ' oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("总金额"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("总金额"))       '本币借
       'oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="现金凭证"                '凭证名称
       oraRet.fields("reference10")="现金转入建行龙口支行"              '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
  '******************************************************************
   End If
          '开始处理明细分录
       '根据收费项目设置科目编码
            if CStr(rs("收费项目")) = "定金" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "首期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "按揭房款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "房款" or CStr(rs("收费项目"))="分期款" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "契税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "印花税" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "工本费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权登记费" then
              num1="61"
              num2="01"
              num3="0000"
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权证印花费" then
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
Name=rs("收费项目") & "现金转入"
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("金额"))          '原币贷
       'oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
       'oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs("金额"))       '本币贷
       oraRet.fields("reference4")="现金凭证"                '凭证名称
       oraRet.fields("reference10")=Name                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '处理下一条记录        
         rs.MoveNext
wend
'posY机处理(默认广发银行)
  '取出符合日期条件的所有收据及楼盘的相关资料,现金处理,(默认在广发银行)
'*******************************************************
sSQL="SELECT sum(A.本次收款) as 金额,A.收费项目,B.总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D,"
sSQL=sSQL & " (SELECT sum(A.本次收款) as 总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D" 
sSQL=sSQL & " WHERE A.记帐 = 1" & DateFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='POS机(Y)') B"
sSQL=sSQL & " WHERE A.记帐 = 1"  & DateFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='POS机(Y)'"
sSQL=sSQL & " group by A.收费项目,B.总金额"
sSQL=sSQL & " order by 收费项目"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 RcptID = " "     '现金存入银行变量
 a=a              '凭证号
'处理凭证日期,因为是汇总,没有明确的收款日期,所以已运算的起始日期为准

'开始处理记录
'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("总金额")) Then
 '          '开始处理新的一张收据
           a=a+1
          RcptID = CStr(rs("总金额"))
           i=0
  '根据本次处理的收据是收款或退款，决定将收据金额记入借方还是贷方
'**********************************************************************

        '根据（现金存入银行）决定财务科目编码(广发银行)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="215"
              num6="0000"
              num7="0000"


'处理单据编号，备注需要明细单据编号
'************************************************************
sSQL2="select distinct 单据编号"
sSQL2=sSQL2 & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D"
sSQL2=sSQL2 & " where A.记帐 = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.收据ID = D.收据ID"
sSQL2=sSQL2 & " AND A.付款类型='POS机(Y)'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="pos机收款单据编号为:"
while not rs1.eof
    ItemName=ItemName+rs1("单据编号") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "转入广发总行(Y)"
'根据收据的收费项目以及楼阁名称决定凭证类型(费用代码)
'*********************************************************************

           'RowTxt 记录将写入中间文件的每条信息
  '****************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("总金额"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("总金额"))       '本币借
       'oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="pos机凭证"                '凭证名称
       oraRet.fields("reference10")="POS机转入广发总行(Y)"             '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
   End If
          '开始处理明细分录
       '根据收费项目设置科目编码
              num1="61"
              num2="01"
              num3="0000"
           if CStr(rs("收费项目")) = "定金" then
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "首期款" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "按揭房款" then
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "房款" or CStr(rs("收费项目"))="分期款" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "契税" then
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "印花税" then
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "工本费" then
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权登记费" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权证印花费" then
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
Name=rs("收费项目") & "pos机转入"
'**********************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("金额"))          '原币贷
       'oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
      ' oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs("金额"))       '本币贷
       oraRet.fields("reference4")="pos机凭证"                '凭证名称
       oraRet.fields("reference10")=Name                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '处理下一条记录        
         rs.MoveNext
wend
'Pos机
'pos机处理(默认广发银行)
  '取出符合日期条件的所有收据及楼盘的相关资料,现金处理,(默认在广发银行)
'*******************************************************
sSQL="SELECT sum(A.本次收款) as 金额,A.收费项目,B.总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D,"
sSQL=sSQL & " (SELECT sum(A.本次收款) as 总金额"
sSQL=sSQL & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D" 
sSQL=sSQL & " WHERE A.记帐 = 1" & DateFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='POS机') B"
sSQL=sSQL & " WHERE A.记帐 = 1"  & DateFilter
sSQL=sSQL & " AND A.收据ID = D.收据ID" 
sSQL=sSQL & " AND A.付款类型='POS机'"
sSQL=sSQL & " group by A.收费项目,B.总金额"
sSQL=sSQL & " order by 收费项目"
Set rs=Database.CreateDynaset(sSQL,ORADYN_READONLY)
 RcptID = " "     '现金存入银行变量
 a=a              '凭证号
'处理凭证日期,因为是汇总,没有明确的收款日期,所以已运算的起始日期为准

'开始处理记录
'**********************************************************************
while not rs.eof
         if RcptID <> CStr(rs("总金额")) Then
 '          '开始处理新的一张收据
           a=a+1
          RcptID = CStr(rs("总金额"))
           i=0
  '根据本次处理的收据是收款或退款，决定将收据金额记入借方还是贷方
'**********************************************************************

        '根据（现金存入银行）决定财务科目编码(广发银行)
'***********************************************************************       
              num1="61"
              num2="01"
              num3="0000"
              num4="1002"
              num5="953"
              num6="0000"
              num7="0000"


'处理单据编号，备注需要明细单据编号
'************************************************************
sSQL2="select distinct 单据编号"
sSQL2=sSQL2 & " FROM climb.V_财务管理_收据明细 A,climb.自定义内容_收据 D"
sSQL2=sSQL2 & " where A.记帐 = 1" & DateFilter & DumpFilter
sSQL2=sSQL2 & " and A.收据ID = D.收据ID"
sSQL2=sSQL2 & " AND A.付款类型='POS机'" 
Set rs1=Database.CreateDynaset(sSQL2,ORADYN_READONLY)
       ItemName="pos机收款单据编号为:"
while not rs1.eof
    ItemName=ItemName+rs1("单据编号") & ","
    rs1.movenext
wend
ItemName=left(ItemName,len(ItemName)-1) 
ItemName=ItemName & "转入广发总行(Y)"
'根据收据的收费项目以及楼阁名称决定凭证类型(费用代码)
'*********************************************************************

           'RowTxt 记录将写入中间文件的每条信息
  '****************************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       'oraRet.fields("entered_cr")=CDbl(rs1("本次收款"))          '原币贷
       oraRet.fields("entered_dr")=CDbl(rs("总金额"))          '原币借
       oraRet.fields("accounted_dr")=CDBL(rs("总金额"))       '本币借
       'oraRet.fields("accounted_cr")=CDBL(rs1("本次收款"))       '本币贷
       oraRet.fields("reference4")="pos机凭证"                '凭证名称
       oraRet.fields("reference10")="POS机转入广发总行(Y)"             '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
   End If
          '开始处理明细分录
       '根据收费项目设置科目编码
              num1="61"
              num2="01"
              num3="0000"
           if CStr(rs("收费项目")) = "定金" then
              num4="2131"
              num5="121"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "首期款" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "按揭房款" then
              num4="2131"
              num5="123"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "房款" or CStr(rs("收费项目"))="分期款" then
              num4="2131"
              num5="122"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "契税" then
              num4="2181"
              num5="501"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "印花税" then
              num4="2181"
              num5="510"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "工本费" then
              num4="2181"
              num5="506"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权登记费" then
              num4="2181"
              num5="509"
              num6="0000"
              num7="0000"
           elseif CStr(rs("收费项目")) = "产权证印花费" then
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
Name=rs("收费项目") & "pos机转入"
'**********************************************
oraRet.addnew
       oraRet.fields("status")="NEW"
       oraRet.fields("set_of_books_id")=5019
       oraRet.fields("accounting_date")=parameters("起始日期").value
       oraRet.fields("currency_code")="CNY"
       oraRet.fields("date_created")=parameters("起始日期").value
       oraRet.fields("created_by")=parameters("用户ID").value         'From Oracle User_id（ORACLE用户ID）
       oraRet.fields("actual_flag")="A"
       oraRet.fields("user_je_category_name")="房屋销售"
       oraRet.fields("user_je_source_name")="销售系统_广州"
       oraRet.fields("segment1")=num1
       oraRet.fields("segment2")=num2
       oraRet.fields("segment3")=num3
       oraRet.fields("segment4")=num4
       oraRet.fields("segment5")=num5
       oraRet.fields("segment6")=num6
       oraRet.fields("segment7")=num7
       oraRet.fields("entered_cr")=CDbl(rs("金额"))          '原币贷
       'oraRet.fields("entered_dr")=CDbl(rs("本次收款"))          '原币借
      ' oraRet.fields("accounted_dr")=CDBL(rs("本次收款"))       '本币借
       oraRet.fields("accounted_cr")=CDBL(rs("金额"))       '本币贷
       oraRet.fields("reference4")="pos机凭证"                '凭证名称
       oraRet.fields("reference10")=Name                '明细摘要
       oraRet.fields("reference21")=date & "0" & a
       oraRet.fields("group_id")=CDBL(date & "0" & a)
  oraRet.update
        '处理下一条记录        
         rs.MoveNext
wend
oraRet.close	
 msgbox("导出成功！！")	
end sub