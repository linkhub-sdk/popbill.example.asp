<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTCashbillService.asp"-->
<%

	'**************************************************************'
	' �˺� Ȩ�ý� ���ݿ����� ���� API ASP SDK Example
	'
	' - ������Ʈ ���� : 2020-07-15
	' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
	' - ���� ������� �̸��� : code@linkhub.co.kr
	'
	' <�׽�Ʈ �������� �غ����>
	' 1) 23, 26�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
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

	set m_HTCashbillService = new HTCashbillService

	m_HTCashbillService.Initialize LinkID, SecretKey

	'����ȯ�� ������ ���߿�(True), �����(False)
	m_HTCashbillService.IsTest = True

	' ������ū IP���ѱ�� ��뿩��, ����(True)
	m_HTCashbillService.IPRestrictOnOff = True

	' �˺� API ���� ���� IP ��뿩��(GA), Ture-���, False-�̻��, �⺻��(False)
	m_HTCashbillService.UseStaticIP = False
%>