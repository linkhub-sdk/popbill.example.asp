<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%

    '**************************************************************'
    ' �˺� Ȩ�ý� ���ڼ��ݰ�꼭 ���� API ASP SDK Example
    '
    ' - ������Ʈ ���� : 2021-06-01
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 22, 25�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
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

    ' �˺� API ���� ���� IP ��뿩��, Ture-���, False-�̻��, �⺻��(False)
    m_HTTaxinvoiceService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩�� True-���(�⺻��-����), false-�̻��
    m_HTTaxinvoiceService.UseLocalTimeYN = True
%>