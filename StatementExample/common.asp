<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/StatementService.asp"-->
<%
    '**************************************************************
    ' �˺� ���ڸ��� API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://developers.popbill.com/guide/statement/asp/getting-started/environment-set-up
    ' - ������Ʈ ���� : 2022-07-20
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 18, 21�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    '**************************************************************

    ' ��ũ���̵� 
    LinkID = "TESTER"

    ' ���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_StatementService = new StatementService

    m_StatementService.Initialize LinkID, SecretKey

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_StatementService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_StatementService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_StatementService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_StatementService.UseLocalTimeYN = True
%>