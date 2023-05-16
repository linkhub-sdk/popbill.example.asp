
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ε� �˸��� ���ø� ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/template#ListATSTemplate
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    Set resultObj = m_KakaoService.ListATSTemplate(testCorpNum)

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
                <legend>�˸��� ���ø� ��� ��ȸ </legend>
                    <%
                        If code = 0 Then
                            For i=0 To resultObj.Count-1
                    %>
                        <fieldset class="fieldset2">
                            <legend>  ���ø� ���� [ <%=i+1%> / <%= resultObj.Count %> ] </legend>
                            <ul>
                                <li> templateCode (���ø� �ڵ�) : <%=resultObj(i).templateCode%></li>
                                <li> templateName (���ø� ����) : <%=resultObj(i).templateName%></li>
                                <li> template (���ø� ����) : <%=resultObj(i).template%></li>
                                <li> plusFriendID (�˻��� ���̵�) : <%=resultObj(i).plusFriendID%></li>
                                <li> ads (�����޽��� ����) : <%=resultObj(i).ads%></li>
                                <li> appendix (�ΰ��޽��� ����) : <%=resultObj(i).appendix%></li>
                                <li> secureYN (�������ø� ����) : <%=resultObj(i).secureYN%></li>
                                <li> state (���ø� ����) : <%=resultObj(i).state%></li>
                                <li> stateDT (���ø� ���� �Ͻ�) : <%=resultObj(i).stateDT%></li>
                            </ul>
                        <%
                            For j=0 To UBound(resultObj(i).btns) -1
                        %>
                                <fieldset class="fieldset3">
                                    <legend> ��ư���� [ <%=j+1%> / <%= UBound(resultObj(i).btns)%> ] </legend>
                                    <ul>
                                        <li>n : <%=resultObj(i).btns(j).n%> </li>
                                        <li>t : <%=resultObj(i).btns(j).t%> </li>
                                        <li>u1 : <%=resultObj(i).btns(j).u1%> </li>
                                        <li>u2 : <%=resultObj(i).btns(j).u2%> </li>
                                    </ul>
                            </fieldset>
                        <%
                                Next
                        %>
                        </fieldset>
                        <%
                            Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>

            </fieldset>
         </div>
    </body>
</html>