<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ���� ����Ʈ ��볻���� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/point#GetUseHistory
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    '��ȸ �Ⱓ�� ��������
    SDate = "20230501"

    '��ȸ �Ⱓ�� ��������
    EDate = "20230530"

    '��� ��������ȣ
    Page = 1

    '�������� ǥ���� ��� ����
    PerPage = 500

    '�ŷ����ڸ� �������� �ϴ� ��� ���� ���� : "D" / "A" �� �� 1
    Order = "D"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_CashbillService.GetUseHistory(CorpNum, SDate,EDate,Page,PerPage,Order, UserID)

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
                <legend>����ȸ�� ����Ʈ ��볻�� Ȯ��</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> code (�����ڵ�) : <%=result.code%></li>
                        <li> total (�� �˻���� �Ǽ�) : <%=result.total%></li>
                        <li> perPage (�������� �˻�����) : <%=result.perPage%></li>
                        <li> pageNum (������ ��ȣ) : <%=result.pageNum%></li>
                        <li> pageCount (������ ����) : <%=result.pageCount%></li>
                    </ul>
                <%
                    Dim i
                    For i = 0 to UBound(result.list)-1
                %>
                    <fieldset class="fieldset2">
                        <legend> UseHistory [ <%= i+1%> / <%=UBound(result.list)%>]</legend>
                        <ul>
                        <li> itemCode (�����ڵ�) : <%=result.list(i).itemCode%></li>
                        <li> txType (����Ʈ ���� ����) : <%=result.list(i).txType%></li>
                        <li> txPoint (���� ����Ʈ) : <%=result.list(i).txPoint%></li>
                        <li> balance (�ܿ� ����Ʈ) : <%=result.list(i).balance%></li>
                        <li> txDT (����Ʈ ���� �Ͻ�) : <%=result.list(i).txDT%></li>
                        <li> UserID (����� ���̵�) : <%=result.list(i).UserID%></li>
                        <li> userName (����ڸ�) : <%=result.list(i).userName%></li>
                        </ul>
                    </fieldset>
                <%
                Next
                %>
                </fieldset>
                <%

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
