<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/MessageService.asp"-->
<%
    '**************************************************************
    ' �˺� ���� API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/message/tutorial/asp
    ' - ������Ʈ ���� : 2021-06-01
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

    set m_MessageService = new MessageService

    m_MessageService.Initialize LinkID, SecretKey

    '����ȯ�� ������, Ture-���, False-�̻��
    m_MessageService.IsTest = True

    '������ū IP���ѱ�� ��뿩��, Ture-���, False-�̻��, �⺻��(True)
    m_MessageService.IPRestrictOnOff = True
    
    '�˺� API ���� ���� IP ��뿩��, Ture-���, False-�̻��, �⺻��(False)
    m_MessageService.UseStaticIP = False
    
    '���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_MessageService.UseLocalTimeYN = True
%>