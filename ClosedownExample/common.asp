<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/ClosedownService.asp"-->
<%
	'**************************************************************
	' �˺� �������ȸ API ASP SDK Example
	'
	' - ������Ʈ ���� : 2020-07-15
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
	
	set m_ClosedownService = new ClosedownService

	m_ClosedownService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_ClosedownService.IsTest = True

	' ������ū IP���ѱ�� ��뿩��, ����(True)
	m_ClosedownService.IPRestrictOnOff = True

	' �˺� API ���� ���� IP ��뿩��(GA), Ture-���, False-�̻��, �⺻��(False)
	m_ClosedownService.UseStaticIP = False
%>