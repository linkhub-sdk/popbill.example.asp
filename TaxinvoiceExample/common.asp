<!--#include virtual="/Popbill/Popbill.asp"-->
<!--#include virtual="/Popbill/TaxinvoiceService.asp"-->
<%
    '**************************************************************
    ' 팝빌 전자세금계산서 API ASP SDK Example
    ' ASP 연동 튜토리얼 안내 : https://developers.popbill.com/guide/taxinvoice/asp/getting-started/tutorial
    '
    ' 업데이트 일자 : 2024-02-27
    ' 연동기술지원 연락처 : 1600-9854
    ' 연동기술지원 이메일 : code@linkhubcorp.com
    '         
    ' <테스트 연동개발 준비사항>
    ' 1) API Key 변경 (연동신청 시 메일로 전달된 정보)
    '     - LinkID : 링크허브에서 발급한 링크아이디
    '     - SecretKey : 링크허브에서 발급한 비밀키
    ' 2) SDK 환경설정 옵션 설정
    '     - IsTest : 연동환경 설정, True-테스트, False-운영(Production), (기본값:True)
    '     - IPRestrictOnOff : 인증토큰 IP 검증 설정, True-사용, False-미사용, (기본값:True)
    '     - UseStaticIP : 통신 IP 고정, True-사용, False-미사용, (기본값:False)
    '     - UseLocalTimeYN : 로컬시스템 시간 사용여부, True-사용, False-미사용, (기본값:True)
    ' 3) 전자세금계산서 발행을 위해 공동인증서를 등록합니다.
    '    - 팝빌사이트 로그인 > [전자세금계산서] > [환경설정] > [공동인증서 관리]
    '    - 공동인증서 등록 팝업 URL (GetTaxCertURL API)을 이용하여 등록
    '**************************************************************

    ' 링크아이디
    LinkID = "TESTER"

    ' 비밀키
    SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

    set m_TaxinvoiceService = new TaxinvoiceService

    m_TaxinvoiceService.Initialize LinkID, SecretKey

    ' 연동환경 설정, True-테스트, False-운영(Production), (기본값:True)
    m_TaxinvoiceService.IsTest = True

    ' 인증토큰 IP 검증 설정, True-사용, False-미사용, (기본값:True)
    m_TaxinvoiceService.IPRestrictOnOff = True

    ' 통신 IP 고정, True-사용, False-미사용, (기본값:False)
    m_TaxinvoiceService.UseStaticIP = False

    ' 로컬시스템 시간 사용여부, True-사용, False-미사용, (기본값:True)
    m_TaxinvoiceService.UseLocalTimeYN = True
%>
