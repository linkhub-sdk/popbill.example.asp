<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/BizInfoCheckService.asp"-->
<%
    '**************************************************************
    ' �˺� ���������ȸ API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://developers.popbill.com/guide/bizinfocheck/asp/getting-started/environment-set-up
    ' - ������Ʈ ���� : 2022-09-26
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
    
    set m_BizInfoCheckService = new BizInfoCheckService

    m_BizInfoCheckService.Initialize LinkID, SecretKey

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_BizInfoCheckService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_BizInfoCheckService.IPRestrictOnOff = True

    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_BizInfoCheckService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_BizInfoCheckService.UseLocalTimeYN = True
%>