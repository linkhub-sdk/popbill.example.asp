<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ���� ��û�� �۾� ����� Ȯ���մϴ�.
    ' - ���� ��û �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
    ' - https://docs.popbill.com/easyfinbank/asp/api#ListActiveJob
    '**************************************************************

    ''�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"
    
    On Error Resume Next

    Set result = m_EasyFinBankService.ListActiveJob(testCorpNum, UserID)
    
    If Err.Number <> 0 Then
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
                <legend>���� ��� ��ȸ</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">					
                                <legend>ListActiveJob [ <%=i+1%> / <%=result.Count%> ] </legend>
                                    <ul>
                                        <li> jobID (�۾����̵�) : <%=result.Item(i).jobID%></li>
                                        <li> startDate (��������) : <%=result.Item(i).startDate%></li>
                                        <li> endDate (��������) : <%=result.Item(i).endDate%></li>
                                        <li> errorCode (�����ڵ�) : <%=result.Item(i).errorCode%></li>
                                        <li> errorReason (�����޽���) : <%=result.Item(i).errorReason%></li>
                                        <li> jobStartDT (�۾� �����Ͻ�) : <%=result.Item(i).jobStartDT%></li>
                                        <li> jobEndDT (�۾� �����Ͻ�) : <%=result.Item(i).jobEndDT%></li>
                                        <li> regDT (���� ��û�Ͻ�) : <%=result.Item(i).regDT%></li>
                                    </ul>
                                </fieldset>
                <%
                        Next
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>	
                <%	
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>
