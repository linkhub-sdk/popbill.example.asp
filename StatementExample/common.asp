<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/StatementService.asp"-->
<%
	'�������� �߱޹��� �������̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey = "lJRNJTfXcNpLTrQTTRpR3vTxTwSTA7Mg0IomDbNy6RA="

	set m_StatementService = new StatementService
	m_StatementService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_StatementService.IsTest = True
%>