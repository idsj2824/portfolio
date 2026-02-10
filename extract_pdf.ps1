$word = New-Object -ComObject Word.Application
$word.Visible = $false
$pdfPath = "c:\Users\KYJ\Desktop\김연주 개인\포트폴리오 자료\2025 하반기 재원생 만족도 조사 리포트_최종 - 복사본.pdf"
$txtPath = "c:\Users\KYJ\Desktop\김연주 개인\포트폴리오 자료\pdf_text.txt"

try {
    $doc = $word.Documents.Open($pdfPath, $false, $true)
    $text = $doc.Content.Text
    $text | Out-File -FilePath $txtPath -Encoding utf8
    $doc.Close()
} catch {
    Write-Error $_
} finally {
    $word.Quit()
}
