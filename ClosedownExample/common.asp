<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/ClosedownService.asp"-->
<%
    '**************************************************************
    ' �˺� �������ȸ API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/closedown/tutorial/asp
    ' - ������Ʈ ���� : 2022-07-20
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
    
    set m_ClosedownService = new ClosedownService

    m_ClosedownService.Initialize LinkID, SecretKey

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_ClosedownService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_ClosedownService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_ClosedownService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_ClosedownService.UseLocalTimeYN = True
%>