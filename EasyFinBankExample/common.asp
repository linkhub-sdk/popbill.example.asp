<!--#include virtual="/Popbill/Popbill.asp"-->
<!--#include virtual="/Popbill/EasyFinBankService.asp"-->
<%

    '**************************************************************'
    ' �˺� ������ȸ API ASP SDK Example
    ' ASP ���� Ʃ�丮�� �ȳ� : https://developers.popbill.com/guide/fax/asp/getting-started/tutorial
    '
    ' ������Ʈ ���� : 2024-02-27
    ' ����������� ����ó : 1600-9854
    ' ����������� �̸��� : code@linkhubcorp.com
    '         
    ' <�׽�Ʈ �������� �غ����>
    ' 1) API Key ���� (������û �� ���Ϸ� ���޵� ����)
    '     - LinkID : ��ũ��꿡�� �߱��� ��ũ���̵�
    '     - SecretKey : ��ũ��꿡�� �߱��� ���Ű
    ' 2) SDK ȯ�漳�� �ɼ� ����
    '     - IsTest : ����ȯ�� ����, True-�׽�Ʈ, False-�(Production), (�⺻��:True)
    '     - IPRestrictOnOff : ������ū IP ���� ����, True-���, False-�̻��, (�⺻��:True)
    '     - UseStaticIP : ��� IP ����, True-���, False-�̻��, (�⺻��:False)
    '     - UseLocalTimeYN : ���ýý��� �ð� ��뿩��, True-���, False-�̻��, (�⺻��:True)
    '**************************************************************

    ' ��ũ���̵�
    LinkID = "TESTER"

    ' ���Ű
    SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_EasyFinBankService = new EasyFinBankService

    m_EasyFinBankService.Initialize LinkID, SecretKey

    ' ����ȯ�� ����, True-�׽�Ʈ, False-�(Production), (�⺻��:True)
    m_EasyFinBankService.IsTest = True

    ' ������ū IP ���� ����, True-���, False-�̻��, (�⺻��:True)
    m_EasyFinBankService.IPRestrictOnOff = True

    ' ��� IP ����, True-���, False-�̻��, (�⺻��:False)
    m_EasyFinBankService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩��, True-���, False-�̻��, (�⺻��:True)
    m_EasyFinBankService.UseLocalTimeYN = True
%>
