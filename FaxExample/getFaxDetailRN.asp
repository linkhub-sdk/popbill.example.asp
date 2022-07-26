<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �ѽ� ���ۻ��� �� ����� Ȯ���մϴ�.
    ' - https://docs.popbill.com/fax/asp/api#GetFaxDetailRN
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"		

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"					

    ' �ѽ����� ��û�� �Ҵ��� ���ۿ�û��ȣ
    requestNum = "" 
 
    On Error Resume Next

    Set result = m_FaxService.GetFaxDetailRN(testCorpNum, requestNum, userID)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�ѽ����� ���۰�� Ȯ�� </legend>
                <% If code = 0 Then 
                        For i=0 To result.Count-1
                %>
                    <fieldset class="fieldset2">
                            <legend> �ѽ� ���۰�� [<%=i+1%>] </legend>
                            <ul>
                                <li>state (���ۻ��� �ڵ�) : <%=result.Item(i).state%> </li>
                                <li>result (���۰�� �ڵ�) : <%=result.Item(i).result%> </li>
                                <li>sendNum (�߽Ź�ȣ) : <%=result.Item(i).sendNum%> </li>
                                <li>senderName (�߽��ڸ�) : <%=result.Item(i).senderName%> </li>
                                <li>receiveNum (���Ź�ȣ) : <%=result.Item(i).receiveNum%> </li>
                                <li>receiveNumType (���Ź�ȣ ����) : <%=result.Item(i).receiveNumType%> </li>
                                <li>receiveName (�����ڸ�) : <%=result.Item(i).receiveName%> </li>
                                <li>title (�ѽ� ����) : <%=result.Item(i).title %> </li>
                                <li>sendPageCnt (��������) : <%=result.Item(i).sendPageCnt%></li>
                                <li>successPageCnt (���� ��������) : <%=result.Item(i).successPageCnt%></li>
                                <li>failPageCnt (���� ��������) : <%=result.Item(i).failPageCnt%></li>
                                <li>cancelPageCnt (��� ��������) : <%=result.Item(i).cancelPageCnt%></li>
                                <li>reserveDT (����ð�) : <%=result.Item(i).reserveDT%></li>
                                <li>sendDT (�߼۽ð�) : <%=result.Item(i).sendDT%></li>
                                <li>receiptDT (���� �����ð�) : <%=result.Item(i).receiptDT%></li>
                                <li>fileNames (�������ϸ� �迭) : <%=result.Item(i).fileNames%></li>
                                <li>receiptNum (������ȣ) : <%=result.Item(i).receiptNum%> </li>
                                <li>requestNum (��û��ȣ) : <%=result.Item(i).requestNum%> </li>
                                <li>interOPRefKey (��Ʈ�� ����Ű) : <%=result.Item(i).interOPRefKey%> </li>
                                <li>chargePageCnt (���� ��������) : <%=result.Item(i).chargePageCnt%> </li>
                                <li>refundPageCnt (ȯ�� ��������) : <%=result.Item(i).refundPageCnt%></li>
                                <li>tiffFileSize (��ȯ���Ͽ뷮 (���� : byte)) : <%=result.Item(i).tiffFileSize%> </li>							
                            </ul>
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
                <%	End If	%>

            </fieldset>
         </div>
    </body>
</html>