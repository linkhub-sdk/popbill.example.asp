<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/CashbillService.asp"-->
<%
	'�������� �߱޹��� �������̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey = "lJRNJTfXcNpLTrQTTRpR3vTxTwSTA7Mg0IomDbNy6RA="

	set m_CashbillService = new CashbillService
	m_CashbillService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_CashbillService.IsTest = True
%>