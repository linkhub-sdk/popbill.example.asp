<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/AccountCheckService.asp"-->
<%
    '**************************************************************
    ' �˺� ��������ȸ API ASP SDK Example
    '
    ' - ������Ʈ ���� : 2021-06-01
    ' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
    ' - ���� ������� �̸��� : code@linkhub.co.kr
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 17, 20�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    '**************************************************************
    
    '��ũ���̵� 
    LinkID = "TESTER"
    
    '���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
    
    set m_AccountCheckService = new AccountCheckService

    m_AccountCheckService.Initialize LinkID, SecretKey

    '����ȯ�� ������, ���߿�(True), �����(False)
    m_AccountCheckService.IsTest = True

    ' ������ū IP���ѱ�� ��뿩��, ����(True)
    m_AccountCheckService.IPRestrictOnOff = True

    ' �˺� API ���� ���� IP ��뿩��(GA), Ture-���, False-�̻��, �⺻��(False)
    m_AccountCheckService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩�� True-���(�⺻��-����), false-�̻��
    m_AccountCheckService.UseLocalTimeYN = True
%>