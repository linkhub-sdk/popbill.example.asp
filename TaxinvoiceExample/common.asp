<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	'�������� �߱޹��� �������̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey ="ut8QMlWBgUYLCgvHqit0rmPdyBPNeWUziQLT0osDvXQ="

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_TaxinvoiceService.IsTest = True
%>