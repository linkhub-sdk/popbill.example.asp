<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/EasyFinBankService.asp"-->
<%

    '**************************************************************'
    ' �˺� ������ȸ API ASP SDK Example
    '
    ' ASP SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/easyfinbank/tutorial/asp
    ' - ������Ʈ ���� : 2022-07-20
    ' - ������� ����ó : 1600-9854
    ' - ������� �̸��� : code@linkhubcorp.com
    '
    ' <�׽�Ʈ �������� �غ����>
    ' 1) 19, 22�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
    '    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
    '**************************************************************
    
    ' ��ũ���̵� 
    LinkID = "TESTER"

    ' ���Ű
    SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_EasyFinBankService = new EasyFinBankService
    m_EasyFinBankService.Initialize LinkID, SecretKey

    ' ����ȯ�� ������, True-���߿�, false-�����
    m_EasyFinBankService.IsTest = True

    ' ������ū �߱� IP ���� On/Off, True-���, false-�̻��, �⺻��(True)
    m_EasyFinBankService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, false-�̻��, �⺻��(false)
    m_EasyFinBankService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� Ture-���, False-�̻��, �⺻��(True)
    m_EasyFinBankService.UseLocalTimeYN = True
%>