<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	'�������� �߱޹��� ��ũ���̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_TaxinvoiceService = new TaxinvoiceService
	m_TaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_TaxinvoiceService.IsTest = True
%>