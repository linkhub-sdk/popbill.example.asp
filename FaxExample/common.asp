<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
    '**************************************************************
    ' �˺� �ѽ� API ASP SDK Example
    '
    ' - ������Ʈ ���� : 2021-12-29
    ' - ���� ������� ����ó : 1600-9854
    ' - ���� ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 17, 20�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    '**************************************************************

    '��ũ���̵� 
    LinkID = "TESTER"

    '���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_FaxService = new FaxService

    m_FaxService.Initialize LinkID, SecretKey

    '����ȯ�漳����, ���߿�(True), �����(False)
    m_FaxService.IsTest = True

    ' ������ū IP���ѱ�� ��뿩��, ����(True)
    m_FaxService.IPRestrictOnOff = True

    ' �˺� API ���� ���� IP ��뿩��, Ture-���, False-�̻��, �⺻��(False)
    m_FaxService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩�� True-���(�⺻��-����), false-�̻��
    m_FaxService.UseLocalTimeYN = True
%>