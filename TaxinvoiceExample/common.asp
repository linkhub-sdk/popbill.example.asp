<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
    '**************************************************************
    ' �˺� ���ڼ��ݰ�꼭 API ASP SDK Example
    '
    ' - ������Ʈ ���� : 2021-06-01
    ' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
    ' - ���� ������� �̸��� : code@linkhub.co.kr
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 23, 26�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    ' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
    ' 3) ���ڼ��ݰ�꼭 ������ ���� ������������ ����մϴ�.
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

    ' ����ȯ�� ������, ���߿�(True),  �����(False)
    m_TaxinvoiceService.IsTest = True

    ' ������ū IP���ѱ�� ��뿩��, ����(True)
    m_TaxinvoiceService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��(GA), Ture-���, False-�̻��, �⺻��(False)
    m_TaxinvoiceService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� True-���(�⺻��-����), false-�̻��
    m_TaxinvoiceService.UseLocalTimeYN = True
%>