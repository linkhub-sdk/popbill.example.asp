<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTCashbillService.asp"-->
<%

	'**************************************************************'
	' �˺� Ȩ�ý� ���ݿ����� ���� API ASP SDK Example
	'
	' - ASP SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/577/
	' - ������Ʈ ���� : 2018-11-22
	' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
	' - ���� ������� �̸��� : code@linkhub.co.kr
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

	'�������� �߱޹��� ���Ű, ���⿡ ����
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTCashbillService = new HTCashbillService

	m_HTCashbillService.Initialize LinkID, SecretKey

	'����ȯ�� ������ ���߿�(True), �����(False)
	m_HTCashbillService.IsTest = True
%>