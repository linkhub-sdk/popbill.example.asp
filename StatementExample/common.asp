<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/StatementService.asp"-->
<%
	'**************************************************************
	' �˺� ���ڸ��� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577
	' - ������Ʈ ���� : 2016-11-14
	' - ���� ������� ����ó : 1600-8536 / 070-4304-2991~2
	' - ���� ������� �̸��� : dev@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 19, 22�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
	'**************************************************************

	' ��ũ���̵� 
	LinkID = "TESTER"

	' ���Ű
	SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_StatementService = new StatementService

	m_StatementService.Initialize LinkID, SecretKey

	' ����ȯ�� ������, ���߿�(True), �����(False)
	m_StatementService.IsTest = True
%>