<!--#include virtual="/Popbill/Popbill.asp"-->
<!--#include virtual="/Popbill/HTCashbillService.asp"-->
<%

    '**************************************************************'
    ' �˺� Ȩ�ý� ���ݿ����� API ASP SDK Example
    ' ASP ���� Ʃ�丮�� �ȳ� : https://developers.popbill.com/guide/htcashbill/asp/getting-started/tutorial
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
    ' 3) Ȩ�ý� �α��� ���������� ����մϴ�. (�μ�����ڵ�� / ���������� ���)
    '    - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���� ����] �޴�
    '    - Ȩ�ý����� ���� ���� �˾� URL(GetCertificatePopUpURL API) ��ȯ�� URL�� �̿��Ͽ�
    '      Ȩ�ý� ���� ó���� �մϴ�.
    '**************************************************************

    ' ��ũ���̵�
    LinkID = "TESTER"

    ' ���Ű
    SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_HTCashbillService = new HTCashbillService

    m_HTCashbillService.Initialize LinkID, SecretKey

    ' ����ȯ�� ����, True-�׽�Ʈ, False-�(Production), (�⺻��:True)
    m_HTCashbillService.IsTest = True

    ' ������ū IP ���� ����, True-���, False-�̻��, (�⺻��:True)
    m_HTCashbillService.IPRestrictOnOff = True

    ' ��� IP ����, True-���, False-�̻��, (�⺻��:False)
    m_HTCashbillService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩��, True-���, False-�̻��, (�⺻��:True)
    m_HTCashbillService.UseLocalTimeYN = True
%>
