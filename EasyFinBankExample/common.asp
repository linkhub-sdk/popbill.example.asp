<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/EasyFinBankService.asp"-->
<%

    '**************************************************************'
    ' �˺� ������ȸ API ASP SDK Example
    '
    ' - ������Ʈ ���� : 2021-12-29
    ' - ������� ����ó : 1600-9854
    ' - ������� �̸��� : code@linkhubcorp.com
    '
    '**************************************************************
    
    '��ũ���̵� 
    LinkID = "TESTER"

    '���Ű
    SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_EasyFinBankService = new EasyFinBankService
    m_EasyFinBankService.Initialize LinkID, SecretKey

    '����ȯ�� ������, ���߿�(True), �����(False)
    m_EasyFinBankService.IsTest = True

    ' ������ū IP���ѱ�� ��뿩��, ����(True)
    m_EasyFinBankService.IPRestrictOnOff = True

    ' �˺� API ���� ���� IP ��뿩��, Ture-���, False-�̻��, �⺻��(False)
    m_EasyFinBankService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩�� True-���(�⺻��-����), false-�̻��
    m_EasyFinBankService.UseLocalTimeYN = True
%>