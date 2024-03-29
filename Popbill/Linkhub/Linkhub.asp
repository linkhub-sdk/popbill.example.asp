<!--#include file="json2.asp"-->
<%
Const linkhub_ServiceURL = "https://auth.linkhub.co.kr"
Const linkhub_ServiceURL_Static = "https://static-auth.linkhub.co.kr"
Const linkhub_ServiceURL_GA = "https://ga-auth.linkhub.co.kr"

class Linkhub

'http://support.microsoft.com/kb/299692
'UniStrToUTF8 - CopyMemory

Private m_linkID
Private m_secretKey
Private m_sha1

Public Property Let LinkID(ByVal value)
    m_linkID = value
End Property

Public Property Get LinkID()
    LinkID = m_linkID
End Property

Public Property Let SecretKey(ByVal value)
    m_secretKey = value
End Property

Public Property Get SecretKey()
    SecretKey = m_secretKey
End Property

Public Function b64md5(postData)
    b64md5 = m_sha1.b64_md5(postData)
End Function

Function b64sha1(d)
    b64sha1 = m_sha1.b64_sha1(d)
End Function

Public Function b64_sha256(postData)
    b64_sha256 = m_sha1.b64_sha256(postData)
End Function

Public Function b64hmacsha1(secretkey, target)
    b64hmacsha1 = m_sha1.b64_sha256(secretkey, target)
End Function

Public Function b64_hmac_sha256(secretkey, target)
    b64_hmac_sha256 = m_sha1.b64_hmac_sha256(secretkey, target)
End Function

Public Sub Class_Initialize
    Set m_sha1 = GetObject( "script:" & Request.ServerVariables("APPL_PHYSICAL_PATH") + "Popbill\Linkhub" & "\sha1.wsc" )
End Sub

Public Sub Class_Terminate
    Set m_sha1 = Nothing
End Sub

Public function getTime(useStaticIP, useLocalTimeYN, useGAIP)
    Dim result

    If useLocalTimeYN Then
        result = m_sha1.getLocalTime()
    Else
        Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
        Call winhttp1.setRequestHeader("User-Agent", "Classic ASP LINKHUB SDK")
        Call winhttp1.Open("GET", getTargetURL(useStaticIP, useGAIP) + "/Time")

        winhttp1.send
        winhttp1.WaitForResponse
        result = winhttp1.responseText

        If winhttp1.Status <> 200 Then
            Dim er : Set er = parse(result)
            Err.raise er.code , "LINKHUB", er.message
        End If

         Set winhttp1 = Nothing
    End If

    getTime = result

End Function

public function getToken(serviceID , access_id, Scope, forwardIP, useStaticIP, useLocalTimeYN, useGAIP)

    Dim postObject : Set postObject = JSON.parse("{}")
    postObject.set "access_id", access_id
    postObject.Set "scope",Scope

    Dim postData : postData = toString(postObject)

    Dim xDate : xDate = getTime(useStaticIP, useLocalTimeYN, useGAIP)
    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("POST", getTargetURL(useStaticIP, useGAIP) + "/" + serviceID + "/Token")
    Call winhttp1.setRequestHeader("x-lh-date", xdate)
    Call winhttp1.setRequestHeader("x-lh-version", "2.0")
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP LINKHUB SDK")
    If forwardIP <> "" Then
            Call winhttp1.setRequestHeader("x-lh-forwarded", forwardIP)
    End If

    Dim target
    target = "POST" + Chr(10)
    target = target + m_sha1.b64_sha256(postData) + Chr(10)
    target = target + xDate + Chr(10)
    If forwardIP <> "" Then
        target = target + forwardIP + Chr(10)
    End If
    target = target + "2.0" + Chr(10)
    target = target + "/" + serviceID + "/Token"

    Dim Bearer : Bearer =  m_sha1.b64_hmac_sha256(m_secretKey,target)

    Call winhttp1.setRequestHeader("Authorization", "LINKHUB " + m_linkID + " " + Bearer)

    winhttp1.send (postData)
    winhttp1.WaitForResponse
    Dim result : result =  winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Dim er : Set er = parse(result)
        Err.raise er.code ,"LINKHUB", er.message
    End if
    Set getToken = parse(result)

    Set winhttp1 = nothing

end function

Public Function GetBalance(BearerToken, serviceID, useStaticIP, useGAIP)

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("GET", getTargetURL(useStaticIP, useGAIP) + "/" + serviceID + "/Point")
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP LINKHUB SDK")

    winhttp1.send
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Dim er : Set er = parse(result)
        Err.raise er.code , "LINKHUB", er.message
    End If

    Set winhttp1 = Nothing
    Dim parsedDic : Set parsedDic = parse(result)

    GetBalance = CDbl(parsedDic.remainPoint)

End Function

Public Function GetPartnerBalance(BearerToken, serviceID, useStaticIP, useGAIP)

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("GET", getTargetURL(useStaticIP, useGAIP) + "/" + serviceID + "/PartnerPoint")
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP LINKHUB SDK")

    winhttp1.send
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Dim er : Set er = parse(result)
        Err.raise er.code , "LINKHUB", er.message
    End If

    Set winhttp1 = Nothing
    Dim parsedDic : Set parsedDic = parse(result)
    GetPartnerBalance = CDbl(parsedDic.remainPoint)

End Function

' ????? ????? ???? ??? URL - 2017/08/29 ???
Public Function GetPartnerURL(BearerToken, serviceID, TOGO, useStaticIP, useGAIP)

    Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

    Call winhttp1.Open("GET", getTargetURL(useStaticIP, useGAIP) + "/" + serviceID + "/URL?TG=" + TOGO)
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("User-Agent", "Classic ASP LINKHUB SDK")

    winhttp1.send
    winhttp1.WaitForResponse
    Dim result : result = winhttp1.responseText

    If winhttp1.Status <> 200 Then
        Dim er : Set er = parse(result)
        Err.raise er.code , "LINKHUB", er.message
    End If

    Set winhttp1 = Nothing
    Dim parsedDic : Set parsedDic = parse(result)
    GetPartnerURL = parsedDic.url

End Function

Private Function getTargetURL(useStaticIP, useGAIP)
    If useGAIP Then
        getTargetURL = linkhub_ServiceURL_GA
    ElseIf useStaticIP Then
        getTargetURL = linkhub_ServiceURL_Static
    Else
        getTargetURL = linkhub_ServiceURL
    End If
End Function

Private Function IIf(condition , trueState,falseState)
    If condition Then
        IIf = trueState
    Else
        IIf = falseState
    End if
End Function

public Function toString(object)
    toString = JSON.Stringify(object)
End Function

Public Function parse(jsonString)
    Set parse = JSON.parse(jsonString)
End Function

'end of class
end Class
%>
