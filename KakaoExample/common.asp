<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/KakaoService.asp"-->
<%
	'**************************************************************
	' �˺� īī���� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577
	' - ������Ʈ ���� : 2018-03-15
	' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
	' - ���� ������� �̸��� : code@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 19, 22�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	'
	'**************************************************************

	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_KakaoService = new KakaoService

	m_KakaoService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_KakaoService.IsTest = True
%>