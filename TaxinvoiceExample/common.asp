<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	'�������� �߱޹��� �������̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey =  "yj1oEHdWJI0fMMsagD0JxBnYnbNRQuTD3MwxNwty2Tg="

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_TaxinvoiceService.IsTest = True
%>