<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%

	'**************************************************************'
	' �˺� Ȩ�ý� ���ڼ��ݰ�꼭 ���� API ASP SDK Example
	'
	' - ������Ʈ ���� : 2019-10-24
	' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
	' - ���� ������� �̸��� : code@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 24, 27�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
	'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
	' 3) Ȩ�ý� ���� ó���� �մϴ�. (�μ�����ڵ�� / ���������� ���)
	'     - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���� ����] �޴�
	'     - Ȩ�ý����� ���� ���� �˾� URL(GetCertificatePopUpURL API) ��ȯ�� URL�� �̿��Ͽ�
	'       Ȩ�ý� ���� ó���� �մϴ�.
	'**************************************************************
	
	'��ũ���̵� 
	LinkID = "TESTER"

	'���Ű
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTTaxinvoiceService = new HTTaxinvoiceService
	m_HTTaxinvoiceService.Initialize LinkID, SecretKey

	'����ȯ�� ������, ���߿�(True), �����(False)
	m_HTTaxinvoiceService.IsTest = True

	' ������ū IP���ѱ�� ��뿩��, ����(True)
	m_HTTaxinvoiceService.IPRestrictOnOff = True
%>