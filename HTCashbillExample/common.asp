<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTCashbillService.asp"-->
<%
	'�������� �߱޹��� ��ũ���̵� 
	LinkID = "TESTER"

	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTCashbillService = new HTCashbillService
	m_HTCashbillService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_HTCashbillService.IsTest = True
%>