<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
    '**************************************************************
    ' 팝빌 팩스 API ASP SDK Example
    '
    ' ASP SDK 연동환경 설정방법 안내 : https://docs.popbill.com/fax/tutorial/asp
    ' - 업데이트 일자 : 2022-05-09
    ' - 연동 기술지원 연락처 : 1600-9854
    ' - 연동 기술지원 이메일 : code@linkhubcorp.com
    '
    ' <테스트 연동개발 준비사항>
    ' 1) 18, 21번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
    '    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
    '**************************************************************

    '링크아이디 
    LinkID = "TESTER"

    '비밀키
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_FaxService = new FaxService

    m_FaxService.Initialize LinkID, SecretKey

    '연동환경 설정값, Ture-사용, False-미사용
    m_FaxService.IsTest = True

    '인증토큰 IP제한기능 사용여부, Ture-사용, False-미사용, 기본값(True)
    m_FaxService.IPRestrictOnOff = True
    
    '팝빌 API 서비스 고정 IP 사용여부, Ture-사용, False-미사용, 기본값(False)
    m_FaxService.UseStaticIP = False
    
    '로컬시스템 시간 사용여부 Ture-사용, False-미사용, 기본값(True)
    m_FaxService.UseLocalTimeYN = True
%>