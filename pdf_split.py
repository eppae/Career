from PyPDF4 import PdfFileReader, PdfFileWriter

# PDF 읽기
input_pdf = PdfFileReader("활동지원사 양성 교육교재.pdf")
output_pdf = PdfFileWriter()

# 특정 페이지 추출 (예: 1~3 페이지)
for page_num in range(126, 137):
    output_pdf.addPage(input_pdf.getPage(page_num - 1))

# 새 파일로 저장
with open("활동지원사 양성 교육교재(126~137).pdf", "wb") as output_file:
    output_pdf.write(output_file)
