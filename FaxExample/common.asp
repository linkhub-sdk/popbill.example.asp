<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
    '**************************************************************
    ' �˺� �ѽ� API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/fax/tutorial/asp
    ' - ������Ʈ ���� : 2022-05-09
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 18, 21�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    '**************************************************************

    '��ũ���̵� 
    LinkID = "TESTER"

    '���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_FaxService = new FaxService

    m_FaxService.Initialize LinkID, SecretKey

    '����ȯ�� ������, Ture-���, False-�̻��
    m_FaxService.IsTest = True

    '������ū IP���ѱ�� ��뿩��, Ture-���, False-�̻��, �⺻��(True)
    m_FaxService.IPRestrictOnOff = True
    
    '�˺� API ���� ���� IP ��뿩��, Ture-���, False-�̻��, �⺻��(False)
    m_FaxService.UseStaticIP = False
    
    '���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_FaxService.UseLocalTimeYN = True
%>