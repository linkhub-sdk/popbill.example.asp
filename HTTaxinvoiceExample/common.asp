<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTTaxinvoiceService.asp"-->
<%

    '**************************************************************'
    ' �˺� Ȩ�ý� ���ڼ��ݰ�꼭 ���� API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/htcashbill/tutorial/asp
    ' - ������Ʈ ���� : 2022-07-20
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 23, 26�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
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

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_HTTaxinvoiceService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_HTTaxinvoiceService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_HTTaxinvoiceService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_HTTaxinvoiceService.UseLocalTimeYN = True
%>