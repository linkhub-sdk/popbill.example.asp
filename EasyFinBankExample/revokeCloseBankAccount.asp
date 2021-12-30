<html xmlns="http:'www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ��������� ������ ������û�� ����մϴ�.
    ' - https://docs.popbill.com/easyfinbank/asp/api#RevokeCloseBankAccount
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    CorpNum = "1234567890"	

    ' �˺�ȸ�� ���̵� 
    UserID = "testkorea"			
    

    ' [�ʼ�] ����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    BankCode = ""

    ' [�ʼ�] ���¹�ȣ ������('-') ����
    AccountNumber = ""

    On Error Resume Next
        Set Presponse = m_EasyFinBankService.RevokeCloseBankAccount(CorpNum, BankCode, AccountNumber, UserID)
        
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
                <legend>���� ������ ������û ���</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>