<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
    '**************************************************************
    ' �˺� ���ڼ��ݰ�꼭 API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/taxinvoice/tutorial/asp
    ' - ������Ʈ ���� : 2022-07-20
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 22, 25�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    ' 2) ���ڼ��ݰ�꼭 ������ ���� ������������ ����մϴ�.
    '    - �˺�����Ʈ �α��� > [���ڼ��ݰ�꼭] > [ȯ�漳��]
    '      > [���������� ����]
    '    - ���������� ��� �˾� URL (GetTaxCertURL API)�� �̿��Ͽ� ���
    '**************************************************************

    ' ��ũ���̵� 
    LinkID = "TESTER"

    ' ���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_TaxinvoiceService = new TaxinvoiceService
    
    ' ���ݰ�꼭 API ���� ��� �ʱ�ȭ
    m_TaxinvoiceService.Initialize LinkID, SecretKey

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_TaxinvoiceService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_TaxinvoiceService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_TaxinvoiceService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_TaxinvoiceService.UseLocalTimeYN = True
%>