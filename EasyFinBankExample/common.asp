<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/EasyFinBankService.asp"-->
<%

	'**************************************************************'
	' �˺� ������ȸ API ASP SDK Example
	'
	' - ������Ʈ ���� : 2020-07-24
	' - ������� ����ó : 1600-9854 / 070-4304-2991
	' - ������� �̸��� : code@linkhub.co.kr
	'
	'**************************************************************
	
	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_EasyFinBankService = new EasyFinBankService
	m_EasyFinBankService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_EasyFinBankService.IsTest = True

	' ������ū IP���ѱ�� ��뿩��, ����(True)
	m_EasyFinBankService.IPRestrictOnOff = True

	' �˺� API ���� ���� IP ��뿩��(GA), Ture-���, False-�̻��, �⺻��(False)
	m_EasyFinBankService.UseStaticIP = False
%>