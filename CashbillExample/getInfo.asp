<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 1���� ���ݿ����� ����/��� ������ Ȯ���մϴ�.
    ' - https://docs.popbill.com/cashbill/asp/api#GetInfo
    '**************************************************************

    '�˺� ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	 

    '�˺� ȸ�� ���̵�
    userID = "testkorea"		 
    
    '������ȣ
    mgtKey = "20211201-001"

    On Error Resume Next

    Set Presponse = m_CashbillService.GetInfo(testCorpNum, mgtKey, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�˺� ���ݿ����� ����/��� ����Ȯ�� </legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>itemKey (���ݿ����� ������Ű) : <%=Presponse.itemKey%></li>
                        <li>confirmNum (����û ���ι�ȣ) : <%=Presponse.confirmNum%></li>
                        <li>mgtKey (������ȣ) : <%=Presponse.mgtKey%></li>
                        <li>tradeDate (�ŷ�����) : <%=Presponse.tradeDate%></li>
                        <li>issueDT (�����Ͻ�) : <%=Presponse.issueDT%></li>
                        <li>regDT (����Ͻ�) : <%=Presponse.regDT%></li>
                        <li>taxationType (��������) : <%=Presponse.taxationType%></li>
                        <li>totalAmount (�ŷ��ݾ�) : <%=Presponse.totalAmount%></li>
                        <li>tradeUsage (�ŷ�����) : <%=Presponse.tradeUsage%></li>
                        <li>tradeOpt (�ŷ�����) : <%=Presponse.tradeOpt%></li>
                        <li>tradeType (��������) : <%=Presponse.tradeType%></li>
                        <li>stateCode (�����ڵ�) : <%=Presponse.stateCode%></li>
                        <li>stateDT (���º����Ͻ�) : <%=Presponse.stateDT%></li>
                        <li>stateMemo (���¸޸�) : <%=Presponse.stateMemo%></li>
                        <li>identityNum (�ŷ�ó �ĺ���ȣ) : <%=Presponse.identityNum%></li>
                        <li>itemName (��ǰ��) : <%=Presponse.itemName%></li>
                        <li>customerName (����) : <%=Presponse.customerName%></li>
                        <li>ntssendDT (����û �����Ͻ�) : <%=Presponse.ntssendDT%></li>
                        <li>ntsresultDT (����û ó����� �����Ͻ�) : <%=Presponse.ntsResultDT%></li>
                        <li>ntsresultCode (����û ó����� �����ڵ�) : <%=Presponse.ntsResultCode%></li>
                        <li>ntsresultMessage (����û ó����� �޽���) : <%=Presponse.ntsResultMessage%></li>
                        <li>orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : <%=Presponse.orgConfirmNum%></li>
                        <li>orgTradeDate (���� ���ݿ����� �ŷ�����) : <%=Presponse.orgTradeDate%></li>
                        <li>printYN (�μ⿩��) : <%=Presponse.printYN%></li>
                    <% Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If%> 
                </ul>
            </fieldset>
         </div>
    </body>
</html>