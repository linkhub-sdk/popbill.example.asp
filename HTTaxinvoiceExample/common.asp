<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%

	'**************************************************************'
	' �˺� Ȩ�ý� ���ڼ��ݰ�꼭 ���� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577/
	' - ������Ʈ ���� : 2016-11-11
	' - ���� ������� ����ó : 1600-8536 / 070-4304-2991~2
	' - ���� ������� �̸��� : dev@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 24, 27�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
	' 3) Ȩ�ý����� �̿밡���� ������������ ����մϴ�.
	'    - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���������� ����] �޴�
	'    - ���������� ���(GetCertificatePopUpURL API) ��ȯ�� URL�� �̿��Ͽ�
	'      �˾� ���������� ���������� ���
	'**************************************************************
	
	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTTaxinvoiceService = new HTTaxinvoiceService
	m_HTTaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_HTTaxinvoiceService.IsTest = True
%>