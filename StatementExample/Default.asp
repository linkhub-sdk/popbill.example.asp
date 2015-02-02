<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>
	<body>
		<div id="content">
			
			<p class="heading1">팝빌 전자명세서 SDK ASP Example.</p>

			<br/>

			<fieldset class="fieldset1">
				<legend>팝빌 기본 API</legend>

				<fieldset class="fieldset2">
					<legend>회원정보</legend>
					<ul>
						<li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원사 가입 여부 확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원사 가입 요청</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원사 잔여포인트 확인</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>전자명세서 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>등록/수정/발행/삭제</legend>
					<ul>
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - 연동관리번호 사용여부 확인</li>
						<li><a href="register.asp">register</a> - 전자명세서 임시저장</li>
						<li><a href="update.asp">update</a> - 전자명세서 수정</li>
						<li><a href="issue.asp">issue</a> - 전자명세서 발행</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - 전자명세서발행취소</li>
						<li><a href="delete.asp">delete</a> - 전자명세서 삭제</li>
						<li><a href="attachFile.asp">attachFile</a> - 전자명세서 첨부파일 추가</li>
						<li><a href="getFiles.asp">getFiles</a> - 전자명세서 첨부파일 목록확인</li>
						<li><a href="deleteFile.asp">deleteFile</a> - 전자명세서 첨부파일 1개 삭제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>정보 확인</legend>
					<ul>
						<li><a href="getInfo.asp">getInfo</a> - 전자명세서 상태확인</li>
						<li><a href="getInfos.asp">getInfos</a> - 전자명세서 상태 대량 확인</li>
						<li><a href="getLogs.asp">getLogs</a> - 전자명세서 이력 확인</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - 전자명세서 상세정보 확인</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가기능</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - 알림메일 재전송</li>
						<li><a href="sendSMS.asp">sendSMS</a> - 알림문자 재전송</li>
						<li><a href="sendFAX.asp">sendFAX</a> - 팩스 전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>팝빌 전자명세서 SSO URL 기능</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 전자명세서 관련 SSO URL 확인</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - 해당 전자명세서의 팝빌 화면을 표시하는 URL 확인</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - 해당 전자명세서의 팝빌 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - 해당 전자명세서의 팝빌 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - 다량(최대100건)의 전자명세서 인쇄 화면을 표시하는 URL 확인 </li>
						<li><a href="getMailURL.asp">getMailURL</a> - 해당 전자명세서의 전송메일상의 "보기" 버튼에 해당하는 URL 확인 </li>
					</ul>
				</fieldset>
				<fieldset class="fieldset2">
					<legend>기타</legend>
					<ul>
						<li><a href="getUnitCost.asp">getUnitCost</a> - 전자명세서 발행단가 확인</li>
					</ul>
				</fieldset>

			</fieldset>
		 </div>
	</body>
</html>