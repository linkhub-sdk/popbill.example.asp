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
						<li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원 가입 여부 확인</li>
						<li><a href="checkID.asp">checkID</a> - 아이디 중복확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원 가입 요청</li>
						<li><a href="getChargeInfo.asp">getChargeInfo</a> - 과금정보 확인</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원 잔여포인트 확인</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청</li>
						<li><a href="listContact.asp">listContact</a> - 담당자 목록 조회</li>
						<li><a href="updateContact.asp">updateContact</a> - 담당자 정보 수정</li>
						<li><a href="registContact.asp">registContact</a> - 담당자 추가</li>
						<li><a href="updateCorpInfo.asp">updateCorpInfo</a> - 회사정보 수정</li>
						<li><a href="getCorpInfo.asp">getCorpInfo</a> - 회사정보 확인</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>전자명세서 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>등록/수정/발행/삭제</legend>
					<ul>
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - 문서관리번호 사용여부 확인</li>
						<li><a href="registIssue.asp">registIssue</a> - 전자명세서 즉시발행</li>
						<li><a href="register.asp">register</a> - 전자명세서 임시저장</li>
						<li><a href="update.asp">update</a> - 전자명세서 수정</li>
						<li><a href="issue.asp">issue</a> - 전자명세서 발행</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - 전자명세서 발행취소</li>
						<li><a href="delete.asp">delete</a> - 전자명세서 삭제</li>
						<li><a href="attachFile.asp">attachFile</a> - 전자명세서 첨부파일 추가</li>
						<li><a href="getFiles.asp">getFiles</a> - 전자명세서 첨부파일 목록확인</li>
						<li><a href="deleteFile.asp">deleteFile</a> - 전자명세서 첨부파일 삭제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>정보 확인</legend>
					<ul>
						<li><a href="getInfo.asp">getInfo</a> - 전자명세서 상태/요약정보 확인</li>
						<li><a href="getInfos.asp">getInfos</a> - 전자명세서 상태/요약정보 확인 - 대량</li>
						<li><a href="getLogs.asp">getLogs</a> - 전자명세서 상태변경 이력 확인</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - 전자명세서 상세정보 확인</li>
						<li><a href="search.asp">search</a> - 전자명세서 목록 조회</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가기능</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - 알림메일 재전송</li>
						<li><a href="sendSMS.asp">sendSMS</a> - 알림문자 재전송</li>
						<li><a href="sendFAX.asp">sendFAX</a> - 전자명세서 팩스 전송</li>
						<li><a href="FAXSend.asp">FAXSend</a> - 선팩스 전송</li>
						<li><a href="attachStatement.asp">attachStatement</a> - 다른 전자명세서 첨부</li>
						<li><a href="detachStatement.asp">detachStatement</a> - 다른 전자명세서 첨부해제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>팝빌 전자명세서 SSO URL 기능</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 전자명세서 관련 SSO URL</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - 전자명세서 보기 팝업 URL</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - 전자명세서 인쇄 팝업 URL</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - 전자명세서 인쇄 팝업 URL - 공급받는자용 </li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - 전자명세서 인쇄 팝업 URL - 대량</li>
						<li><a href="getMailURL.asp">getMailURL</a> - 전자명세서 메일링크 URL</li>
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