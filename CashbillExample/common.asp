<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/CashbillService.asp"-->
<%
	'**************************************************************
	' �˺� ���ݿ����� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577
	' - ������Ʈ ���� : 2017-08-17
	' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
	' - ���� ������� �̸��� : code@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 19, 22�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
	'**************************************************************
	
	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_CashbillService = new CashbillService

	m_CashbillService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_CashbillService.IsTest = True
%>