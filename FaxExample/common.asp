<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
	'�������� �߱޹��� �������̵� 
	LinkID = "TESTER"
	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey = "EGh1WjSul2JcPazL6AtQy7VTGamL5i14SK4/qGZvz6E="

	set m_FaxService = new FaxService
	m_FaxService.Initialize LinkID, SecretKey

	'����ȯ�漳����, �׽�Ʈ�Ϸ��� ����� ��ȯ�� False�� ���� �����ϰų� �ּ�ó��.
	m_FaxService.IsTest = True
%>