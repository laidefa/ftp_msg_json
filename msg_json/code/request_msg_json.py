# encoding: utf-8
import time
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import re
time1 = time.time()
import json
import os
import datetime
import pymysql
today=datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
print yesterday

#####################创建昨天的文件夹########################
isExists=os.path.exists("/home/laidefa/msg_json/data/%s" %yesterday)
if not isExists:
    os.mkdir("/home/laidefa/msg_json/data/%s" %yesterday)
else:
    print "目录已存在"



######################从mysql数据库读数据###########################################

## 加上字符集参数，防止中文乱码
dbconn=pymysql.connect(
  host="XXXXX",
  database="cgjr",
  user="XXXXX",
  password="XXXXXX",
  port=XXXXXX,
  charset='utf8'
 )

#sql语句
sqlcmd="SELECT order_no,debt_no,request_msg,create_time,return_msg from t_order_debt_log  WHERE return_msg like '%success%' and request_msg like '%微农贷%' and debt_no like '%WND%' and `status` in(0,1) and substr(create_time,1,10)=date_sub(curdate(),interval 1 day)"

#利用pandas 模块导入mysql数据
df=pd.read_sql(sqlcmd,dbconn)
# print df


######################################################写入excel设置问题#########################################
import xlsxwriter
workbook = xlsxwriter.Workbook("/home/laidefa/msg_json/data/%s/debt_wnd_%s.xlsx" %(yesterday,yesterday), options={'strings_to_urls': False})

format=workbook.add_format()
format.set_border(1)
format_title=workbook.add_format()
format_title.set_border(1)
format_title.set_bg_color('#cccccc')
format_title.set_align('center')
format_title.set_bold()
format_ave=workbook.add_format()
format_ave.set_border(1)
format_ave.set_num_format('0.00')

data_format=workbook.add_format()
data_format.set_num_format('yyyy-mm-dd HH:MM:SS')
data_format.set_border(1)




########################################5、产品基本信息######################################################
order_no=[]

debtType=[]
productName=[]
serialNumber=[]
amount=[]
balanceAmount=[]
contractRate=[]
category1=[]
category2=[]
category3=[]
feeRate=[]
repayment=[]
startTime=[]
endTime=[]
channel=[]
borrowingDays=[]
productLimit=[]
creditFeeMoney=[]
userInterestFrom=[]
interestFrom=[]
creditRepayment=[]
creditDeposit=[]
borrowerUserId=[]
assureUserId=[]
scaleFlag=[]
publishCompany=[]


for i in range(0,len(df)):
    order_no1 = df.iloc[i,0]
    m=re.findall('data=(.*?), appId=',df.iloc[i,2],re.S)
    data_json1=json.loads(m[0])
    product1=data_json1["product"]
    debtType1=product1['debtType']
    productName1=product1['productName']
    serialNumber1=product1['serialNumber']
    amount1=product1['amount']
    balanceAmount1=product1['balanceAmount']
    contractRate1 =product1['contractRate']
    category11 = product1['category1']
    category21 =product1['category2']
    category31 = product1['category3']
    feeRate1 =product1['feeRate']
    repayment1 =product1['repayment']
    startTime1 = product1['startTime']
    endTime1 =product1['endTime']
    channel1 = product1['channel']
    borrowingDays1 =product1['borrowingDays']
    productLimit1 =product1['productLimit']
    creditFeeMoney1 =product1['creditFeeMoney']
    userInterestFrom1 =product1['userInterestFrom']
    interestFrom1 = product1['interestFrom']
    creditRepayment1 =product1['creditRepayment']
    creditDeposit1 = product1['creditDeposit']
    borrowerUserId1 =data_json1['borrowerUserId']
    assureUserId1 = data_json1['assureUserId']
    scaleFlag1 =product1['scaleFlag']
    publishCompany1=product1['publishCompany']

    order_no.append(order_no1)

    debtType.append(debtType1)
    productName.append(productName1)
    serialNumber.append(serialNumber1)
    amount.append(amount1)
    balanceAmount.append(balanceAmount1)
    contractRate.append(contractRate1)
    category1.append(category11)
    category2.append(category21)
    category3.append(category31)
    feeRate.append(feeRate1)
    repayment.append(repayment1)
    startTime.append(startTime1)
    endTime.append(endTime1)
    channel.append(channel1)
    borrowingDays.append(borrowingDays1)
    productLimit.append(productLimit1)
    creditFeeMoney.append(creditFeeMoney1)
    userInterestFrom.append(userInterestFrom1)
    interestFrom.append(interestFrom1)
    creditRepayment.append(creditRepayment1)
    creditDeposit.append(creditDeposit1)
    borrowerUserId.append(borrowerUserId1)
    assureUserId.append(assureUserId1)
    scaleFlag.append(scaleFlag1)
    publishCompany.append(publishCompany1)


    # print   debtType1,productName1,serialNumber1,amount1,balanceAmount1,contractRate1,category11,category21,category31,feeRate1,repayment1,startTime1,endTime1,channel1,borrowingDays1, \
    #         productLimit1,creditFeeMoney1,userInterestFrom1,interestFrom1,creditRepayment1,creditDeposit1,borrowerUserId1,assureUserId1,scaleFlag1,publishCompany1


product=pd.DataFrame({"订单号":order_no,"债权类型":debtType,"产品名称":productName,"债权编号":serialNumber,"产品总额":amount,"产品剩余总额":balanceAmount,"合同利率":contractRate,\
                      "一级分类":category1,"二级分类":category2,"三级分类":category3,"手续费":feeRate,"付息方式":repayment,"借款开始时间":startTime,"借款结束时间":endTime,\
                      "渠道":channel,"借款周期":borrowingDays,"产品期限":productLimit,"信贷服务费":creditFeeMoney,"用户起息方式":userInterestFrom,"起息日":interestFrom,\
                      "信贷付息类型":creditRepayment,"信贷保证金":creditDeposit,"借款人id":borrowerUserId,"担保人id":assureUserId,"募集期标识":scaleFlag,"分公司":publishCompany})



# print product.columns



##################################################写入excel数据##############################################
worksheet5 = workbook.add_worksheet('产品基本信息')
title5 = [u'订单号',u'债权类型',u'产品名称',u'债权编号',u'产品总额',u'产品剩余总额',u'合同利率',\
          u'一级分类',u'二级分类',u'三级分类',u'手续费',u'付息方式',u'借款开始时间',u'借款结束时间',\
          u'渠道',u'借款周期',u'产品期限',u'信贷服务费',u'用户起息方式',u'起息日',u'信贷付息类型',\
          u'信贷保证金',u'借款人id',u'担保人id',u'募集期标识',u'分公司'

          ]

worksheet5.write_row('A1',title5,format_title)
worksheet5.write_column('A2:', product.iloc[:,24],format)
worksheet5.write_column('B2:', product.iloc[:,15],format)
worksheet5.write_column('C2', product.iloc[:,4],format)
worksheet5.write_column('D2', product.iloc[:,16],format)
worksheet5.write_column('E2', product.iloc[:,5],format)
worksheet5.write_column('F2', product.iloc[:,3],format)
worksheet5.write_column('G2', product.iloc[:,19],format)
worksheet5.write_column('H2', product.iloc[:,0],format)
worksheet5.write_column('I2', product.iloc[:,2],format)
worksheet5.write_column('J2', product.iloc[:,1],format)
worksheet5.write_column('K2', product.iloc[:,20],format)
worksheet5.write_column('L2', product.iloc[:,7],format)
worksheet5.write_column('M2', product.iloc[:,13],data_format)
worksheet5.write_column('N2', product.iloc[:,14],data_format)
worksheet5.write_column('O2', product.iloc[:,22],format)
worksheet5.write_column('P2', product.iloc[:,12],format)
worksheet5.write_column('Q2', product.iloc[:,6],format)
worksheet5.write_column('R2', product.iloc[:,10],format)
worksheet5.write_column('S2', product.iloc[:,23],format)
worksheet5.write_column('T2', product.iloc[:,25],format)
worksheet5.write_column('U2', product.iloc[:,8],format)
worksheet5.write_column('V2', product.iloc[:,9],format)
worksheet5.write_column('W2', product.iloc[:,11],format)
worksheet5.write_column('X2', product.iloc[:,21],format)
worksheet5.write_column('Y2', product.iloc[:,18],format)
worksheet5.write_column('Z2', product.iloc[:,17],format)



########################################################4、信贷借款人历史信息########################################
order_no=[]
applyNum=[]
applyAmount=[]
unpayMoney=[]
unpayNum=[]
normalAmount=[]
overdueNum=[]
monthIncome=[]
companyType=[]
workTime=[]
workPlace=[]
personId=[]


for i in range(0,len(df)):
    order_no1 = df.iloc[i,0]
    m=re.findall('data=(.*?), appId=',df.iloc[i,2],re.S)
    data_json1=json.loads(m[0])
    debtPersonCredit1=data_json1["debtPersonCredit"]
    applyNum1=debtPersonCredit1['applyNum']
    applyAmount1=debtPersonCredit1['applyAmount']
    unpayMoney1=debtPersonCredit1['unpayMoney']
    unpayNum1=debtPersonCredit1['unpayNum']
    normalAmount1=debtPersonCredit1['normalAmount']
    overdueNum1=debtPersonCredit1['overdueNum']
    monthIncome1=debtPersonCredit1['monthIncome']
    companyType1=debtPersonCredit1['companyType']
    workTime1=debtPersonCredit1['workTime']
    workPlace1=debtPersonCredit1['workPlace']
    personId1=debtPersonCredit1['personId']
    order_no.append(order_no1)
    applyNum.append(applyNum1)
    applyAmount.append(applyAmount1)
    unpayMoney.append(unpayMoney1)
    unpayNum.append(unpayNum1)
    normalAmount.append(normalAmount1)
    overdueNum.append(overdueNum1)
    monthIncome.append(monthIncome1)
    companyType.append(companyType1)
    workTime.append(workTime1)
    workPlace.append(workPlace1)
    personId.append(personId1)
    # print order_no1,applyNum1,applyAmount1,unpayMoney1,unpayNum1,normalAmount1,overdueNum1,monthIncome1,companyType1,workTime1,workPlace1,personId1

debtPersonCredit=pd.DataFrame({"订单号":order_no,"申请借款笔数":applyNum,"累计借款":applyAmount,"待还本息":	unpayMoney,"未还借款笔数":unpayNum, \
                               "正常还清笔数":normalAmount,"逾期还款笔数":overdueNum,"月收入":monthIncome,"公司性质":companyType,"工作时间":workTime, \
                               "工作地点":workPlace,"借款人ID":personId
                               })
# print debtPersonCredit



##################################################写入excel数据##############################################
worksheet4 = workbook.add_worksheet('信贷借款人历史信息')
title4 = [u'订单号',u'申请借款笔数',u'累计借款',u'待还本息',u'未还借款笔数',u'正常还清笔数',u'逾期还款笔数',u'月收入',u'公司性质',u'工作时间',u'工作地点',u'借款人ID']

worksheet4.write_row('A1',title4,format_title)
worksheet4.write_column('A2:', debtPersonCredit.iloc[:,10],format)
worksheet4.write_column('B2:', debtPersonCredit.iloc[:,8],format)
worksheet4.write_column('C2', debtPersonCredit.iloc[:,9],format)
worksheet4.write_column('D2', debtPersonCredit.iloc[:,4],format)
worksheet4.write_column('E2', debtPersonCredit.iloc[:,6],format)
worksheet4.write_column('F2', debtPersonCredit.iloc[:,7],format)
worksheet4.write_column('G2', debtPersonCredit.iloc[:,11],format)
worksheet4.write_column('H2', debtPersonCredit.iloc[:,5],format)
worksheet4.write_column('I2', debtPersonCredit.iloc[:,1],format)
worksheet4.write_column('J2', debtPersonCredit.iloc[:,3],data_format)
worksheet4.write_column('K2', debtPersonCredit.iloc[:,2],format)
worksheet4.write_column('L2', debtPersonCredit.iloc[:,0],format)




##########################################3、借款人审核信息##########################################
order_no=[]
auditType=[]
auditResult=[]
auditDate=[]

for i in range(0,len(df)):
    order_no1 = df.iloc[i,0]
    m=re.findall('data=(.*?), appId=',df.iloc[i,2],re.S)
    data_json1=json.loads(m[0])
    debtPersonAudit1=data_json1["debtPersonAudit"]
    for each in debtPersonAudit1:
        order_no.append(order_no1)
        auditType1=each["auditType"]
        auditResult1=each["auditResult"]
        auditDate1=each['auditDate']

        auditType.append(auditType1)
        auditResult.append(auditResult1)
        auditDate.append(auditDate1)

        # print order_no1,auditType1,auditType1,auditResult1,auditDate1

debtPersonAudit=pd.DataFrame({"订单号":order_no,"审核项目":auditType,"审核结果":auditResult,"审核日期":auditDate})
# print debtPersonAudit


##################################################写入excel数据##############################################
worksheet3 = workbook.add_worksheet('借款人审核信息')
title3 = [u'订单号',u'审核项目',u'审核结果',u'审核日期']

worksheet3.write_row('A1',title3,format_title)
worksheet3.write_column('A2:', debtPersonAudit.iloc[:,3],format)
worksheet3.write_column('B2:', debtPersonAudit.iloc[:,2],format)
worksheet3.write_column('C2', debtPersonAudit.iloc[:,1],format)
worksheet3.write_column('D2', debtPersonAudit.iloc[:,0],data_format)








###########################################1、债权描述######################################
order_no=[]
desc=[]
use=[]
pledge=[]
source=[]
risk=[]
advice=[]

for i in range(0,len(df)):
    order_no1 = df.iloc[i, 0]
    m=re.findall('data=(.*?), appId=',df.iloc[i,2],re.S)
    data_json1=json.loads(m[0])
    debtDesc1=data_json1["debtDesc"]
    desc1=debtDesc1["desc"]
    use1=debtDesc1["use"]
    pledge1=debtDesc1["pledge"]
    source1=debtDesc1["source"]
    risk1=debtDesc1["risk"]
    advice1 = debtDesc1["advice"]

    order_no.append(order_no1)
    desc.append(desc1)
    use.append(use1)
    pledge.append(pledge1)
    source.append(source1)
    risk.append(risk1)
    advice.append(advice1)

    # print order_no1,desc1,use1,pledge1,source1,risk1,advice1



debtDesc=pd.DataFrame({"订单号":order_no,"债权描述":desc,"资金用途":use,"抵押物描述":pledge,"还款来源":source,"风控措施":risk,"意见":advice})
# print debtDesc


# ##################################写入excel################################################################


worksheet1 = workbook.add_worksheet('债权描述')
title1 = [u'订单号',u'债权编号',u'资金用途',u'抵押物描述',u'还款来源',u'风控措施',u'意见']



worksheet1.write_row('A1',title1,format_title)
worksheet1.write_column('A2:', debtDesc.iloc[:,3],format)
worksheet1.write_column('B2:', debtDesc.iloc[:,0],format)
worksheet1.write_column('C2', debtDesc.iloc[:,4],format)
worksheet1.write_column('D2', debtDesc.iloc[:,2],format)
worksheet1.write_column('E2', debtDesc.iloc[:,5],format)
worksheet1.write_column('F2', debtDesc.iloc[:,6],format)
worksheet1.write_column('G2', debtDesc.iloc[:,1],format)



#######################################2、付息数据#######################################################
order_no=[]
startTime=[]
endTime=[]
days=[]
unitInterest=[]
payableInterest=[]
payablePrincipal=[]
number=[]
period=[]


for i in range(0,len(df)):
    order_no1 = df.iloc[i,0]
    m=re.findall('data=(.*?), appId=',df.iloc[i,2],re.S)
    data_json1=json.loads(m[0])
    debtInterest1=data_json1["debtInterest"]
    for each in debtInterest1:
        startTime1=each["startTime"]
        endTime1=each["endTime"]
        days1=each["days"]
        unitInterest1=each['unitInterest']
        payableInterest1=each['payableInterest']
        payablePrincipal1=each['payablePrincipal']
        number1=each['number']
        period1=each['period']

        order_no.append(order_no1)
        startTime.append(startTime1)
        endTime.append(endTime1)
        days.append(days1)
        unitInterest.append(unitInterest1)
        payableInterest.append(payableInterest1)
        payablePrincipal.append(payablePrincipal1)
        number.append(number1)
        period.append(period1)

        # print order_no1,startTime1,endTime1,days1,unitInterest1,payableInterest1,payablePrincipal1,number1,period1

debtInterest=pd.DataFrame({"订单号":order_no,"本期开始时间":startTime,"本期结束时间":endTime,"本期天数":days,"单位天息":unitInterest,"待支付利息":payableInterest,"待支付本金":payablePrincipal,\

                       "当前期数":number,"总期数":period})
# print debtInterest




##################################################写入excel数据##############################################
worksheet2 = workbook.add_worksheet('息付息数据')
title2 = [u'订单号',u'本期开始时间',u'本期结束时间',u'本期天数',u'单位天息',u'待支付利息',u'待支付本金',u'当前期数',u'总期数']


worksheet2.write_row('A1',title2,format_title)
worksheet2.write_column('A2:', debtInterest.iloc[:,8],format)
worksheet2.write_column('B2:', debtInterest.iloc[:,6],data_format)
worksheet2.write_column('C2', debtInterest.iloc[:,7],data_format)
worksheet2.write_column('D2', debtInterest.iloc[:,5],format)
worksheet2.write_column('E2', debtInterest.iloc[:,0],format)
worksheet2.write_column('F2', debtInterest.iloc[:,2],format)
worksheet2.write_column('G2', debtInterest.iloc[:,3],format)
worksheet2.write_column('H2', debtInterest.iloc[:,1],format)
worksheet2.write_column('I2', debtInterest.iloc[:,4],format)



workbook.close()
time2 = time.time()
print u'ok,解析json结束!'
print u'总共耗时：' + str(time2 - time1) + 's'





