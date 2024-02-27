<!--#include virtual="/Popbill/Popbill.asp"-->
<!--#include virtual="/Popbill/KakaoService.asp"-->
<%
    '**************************************************************
    ' �˺� īī���� API ASP SDK Example
    ' ASP ���� Ʃ�丮�� �ȳ� : https://developers.popbill.com/guide/kakaotalk/asp/getting-started/tutorial
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
    ' 3) �߽Ź�ȣ ��������� �մϴ�. (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
    '    - 1. �˺� ����Ʈ �α��� > [����/�ѽ�] > [īī����] > [�߽Ź�ȣ �������] �޴����� ���
    '    - 2. getSenderNumberMgtURL API�� ���� ��ȯ�� URL�� �̿��Ͽ� �߽Ź�ȣ ���
    ' 4) ����Ͻ� ä�� ��� �� �˸��� ���ø��� ��û�մϴ�.
    '    - 1. ����Ͻ� ä�� ��� (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
    '        �� �˺� ����Ʈ �α��� [����/�ѽ�] > [īī����] > [īī���� ����] > 'īī���� ä�� ����' �޴����� ���
    '        �� GetPlusFriendMgtURL API �� ���� ��ȯ�� URL�� �̿��Ͽ� ���
    '    - 2. �˸��� ���ø� ��û (��Ϲ���� ����Ʈ/API �ΰ��� ����� �ֽ��ϴ�.)
    '        �� �˺� ����Ʈ �α��� [����/�ѽ�] > [īī����] > [īī���� ����] > '�˸��� ���ø� ����' �޴����� ���
    '        �� GetATSTemplateMgtURL API �� ���� URL�� �̿��Ͽ� ���.
    '**************************************************************

    ' ��ũ���̵�
    LinkID = "TESTER"

    ' ���Ű
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_KakaoService = new KakaoService

    m_KakaoService.Initialize LinkID, SecretKey

    ' ����ȯ�� ����, True-�׽�Ʈ, False-�(Production), (�⺻��:True)
    m_KakaoService.IsTest = True

    ' ������ū IP ���� ����, True-���, False-�̻��, (�⺻��:True)
    m_KakaoService.IPRestrictOnOff = True

    ' ��� IP ����, True-���, False-�̻��, (�⺻��:False)
    m_KakaoService.UseStaticIP = False

    ' ���ýý��� �ð� ��뿩��, True-���, False-�̻��, (�⺻��:True)
    m_KakaoService.UseLocalTimeYN = True
%>
