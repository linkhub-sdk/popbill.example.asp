<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
	'**************************************************************
	' �˺� �ѽ� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577
	' - ������Ʈ ���� : 2017-05-24
	' - ���� ������� ����ó : 1600-8536 / 070-4304-2991
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

	set m_FaxService = new FaxService

	m_FaxService.Initialize LinkID, SecretKey

	'����ȯ�漳����, ���߿�(True), �����(False)
	m_FaxService.IsTest = True
%>