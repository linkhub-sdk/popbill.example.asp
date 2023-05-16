<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ������ȸ ���񽺸� �̿��� ���¸� �˺��� ����մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#GetBankAccountInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' �������� ��ü ����
    Set infoObj = New EasyFinBankAccountForm

    ' ����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    infoObj.BankCode = ""

    ' ���¹�ȣ ������('-') ����
    infoObj.AccountNumber = ""

    ' ���º�й�ȣ
    infoObj.AccountPWD = ""

    ' ��������, "����" �Ǵ� "����" �Է�
    infoObj.AccountType = ""

    ' ������ �ĺ����� (��-�� ����)
    ' ���������� "����"�� ��� : ����ڹ�ȣ(10�ڸ�)
    ' ���������� "����"�� ��� : ������ ������� (6�ڸ�-YYMMDD)
    infoObj.IdentityNumber = ""

    ' ���� ��Ī
    infoObj.AccountName = ""

    ' ���ͳݹ�ŷ ���̵� (�������� �ʼ�)
    infoObj.BankID = ""

    ' ��ȸ���� ���� ���̵� (�뱸����, ����, �������� �ʼ�)
    infoObj.FastID = ""

    ' ��ȸ���� ���� ��й�ȣ (�뱸����, ����, �������� �ʼ�
    infoObj.FastPWD = ""

    ' �����Ⱓ(����), 1~12 �Է°���, �̱���� �⺻��(1) ó��
    ' - ��Ʈ�� ���ݹ���� ��� �Է°��� ������� 1���� ó��
    infoObj.UsePeriod = ""

    ' �޸�
    infoObj.Memo = ""

    On Error Resume Next
        Set Presponse = m_EasyFinBankService.RegistBankAccount(CorpNum, infoObj, UserID)

        If Err.Number <> 0 Then
            code = Err.Number
            message = Err.Description
            Err.Clears
        Else
            code = Presponse.code
            message =Presponse.message
        End If
    On Error GoTo 0
%>

    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���� ���</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>