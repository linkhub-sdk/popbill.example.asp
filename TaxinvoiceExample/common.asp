<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
	'**************************************************************
	' �˺� ���ڼ��ݰ�꼭 API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577
	' - ������Ʈ ���� : 2016-11-09
	' - ���� ������� ����ó : 1600-8536 / 070-4304-2991~2
	' - ���� ������� �̸��� : dev@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 24, 27�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
	' 3) ���ڼ��ݰ�꼭 ������ ���� ������������ ����մϴ�.
	'    - �˺�����Ʈ �α��� > [���ڼ��ݰ�꼭] > [ȯ�漳��]
	'      > [���������� ����]
	'    - ���������� ��� �˾� URL (GetPopbillURL API)�� �̿��Ͽ� ���
	'**************************************************************

	' ��ũ���̵� 
	LinkID = "TESTER"

	' ���Ű
	SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_TaxinvoiceService = new TaxinvoiceService
	
	' ���ݰ�꼭 API ���� ��� �ʱ�ȭ
	m_TaxinvoiceService.Initialize LinkID, SecretKey

	' ����ȯ�� ������, ���߿�(True),  �����(False)
	m_TaxinvoiceService.IsTest = True
%>