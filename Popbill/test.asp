<!--#include file="Popbill.asp"--> 
<!--#include file="TaxinvoiceService.asp"--> 
<html>
<head>
	<title>ASP �� ��������.</title>
	<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
</head>
<body>
<div>
<%

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize "TESTER", "t4B19Ph5K2aIh9oNd91Q99Vwe9jST2/2IJbWjxhCgsA="
	m_TaxinvoiceService.IsTest = True
	
	On Error Resume Next

	remainPoint = m_TaxinvoiceService.getBalance("1231212312")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "getBalance : " + CStr(remainpoint)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	remainPoint = m_TaxinvoiceService.getPartnerBalance("1231212312")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "getPartnerBalance : " + CStr(remainpoint)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	url = m_TaxinvoiceService.GetPopbillURL("1231212312","userid","CHRG")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "GetPopbillURL : " + url
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.CheckIsMember("1231212312","TESTER")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "CheckIsMember : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0


	Response.write "<br/>"
	
	Set newTaxinvoice = New Taxinvoice
	newTaxinvoice.writeDate = "20150105"             '�ʼ�, ����� �ۼ�����
    newTaxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    newTaxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    newTaxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    newTaxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    newTaxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    
    
    newTaxinvoice.invoicerCorpNum = "1231212312"
    newTaxinvoice.invoicerTaxRegID = "" '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    newTaxinvoice.invoicerCorpName = "������ ��ȣ&%$@<>^^"
    newTaxinvoice.invoicerMgtKey = "1234567890"    '������ ��Ʈ�� ������ȣ
    newTaxinvoice.invoicerCEOName = "������"" ��ǥ�� ����"
    newTaxinvoice.invoicerAddr = "������ �ּ�"
    newTaxinvoice.invoicerBizClass = "������ ����"
    newTaxinvoice.invoicerBizType = "������ ����,����2"
    newTaxinvoice.invoicerContactName = "������ ����ڸ�"
    newTaxinvoice.invoicerEmail = "test@test.com"
    newTaxinvoice.invoicerTEL = "070-7070-0707"
    newTaxinvoice.invoicerHP = "010-000-2222"
    newTaxinvoice.invoicerSMSSendYN = True '����� ���ڹ߼۱�� ���� Ȱ��
    
    newTaxinvoice.invoiceeType = "�����"
    newTaxinvoice.invoiceeCorpNum = "8888888888"
    newTaxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    newTaxinvoice.invoiceeMgtKey = ""
    newTaxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    newTaxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    newTaxinvoice.invoiceeBizClass = "���޹޴��� ����"
    newTaxinvoice.invoiceeBizType = "���޹޴��� ����"
    newTaxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    newTaxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    newTaxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    newTaxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    newTaxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    newTaxinvoice.modifyCode = "" '�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    newTaxinvoice.originalTaxinvoiceKey = "" '�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��.
    newTaxinvoice.serialNum = "123"
    newTaxinvoice.cash = ""          '����
    newTaxinvoice.chkBill = ""       '��ǥ
    newTaxinvoice.note = ""          '����
    newTaxinvoice.credit = ""        '�ܻ�̼���
    newTaxinvoice.remark1 = "���1"
    newTaxinvoice.remark2 = "���2"
    newTaxinvoice.remark3 = "���3"
    newTaxinvoice.kwon = "1"
    newTaxinvoice.ho = "1"
    
    newTaxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    newTaxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
  

	'���׸� �߰�.
    
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"

    newTaxinvoice.AddDetail newDetail

    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2"
    
    newTaxinvoice.AddDetail newDetail
 

	'�߰������ �߰�. �ɼ�.
    set newContact = New Contact
    newContact.contactName = "����� ����"
    newContact.email = "test2@test.com"
    
    newTaxinvoice.AddContact newContact
    
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.Register("1231212312",newTaxinvoice,false,"")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Register : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0


	Response.write "<br/>"
	On Error Resume Next

	Set taxinvoiceInfo = m_TaxinvoiceService.GetDetailInfo("1231212312",SELL,"1234567890")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "GetDetailInfo : " & taxinvoiceInfo.InvoicerCorpName & "|" &  (taxinvoiceInfo.detailList.Get(0).itemName)
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presonse = m_TaxinvoiceService.AttachFile("1231212312",SELL,"1234567890","C:\Inetpub\wwwroot\Popbill\�ΰ�.gif","userid")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "AttachFile : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0

	Response.write "<br/>"
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.Delete("1231212312",SELL,"1234567890","")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Delete : [" + CStr(Presponse.code) & " ] "  & Presponse.message
	End If

	On Error GoTo 0

%>
</div>
</body>
</html>