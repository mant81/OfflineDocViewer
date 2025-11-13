package com.sschoi.offlinedocviewer;

import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractMethod;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.convert.out.pdf.viaXSLFO.Conversion;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.*;
import java.util.List;

public class DocumentConverter {

    public void convertToPdf(String inputPath, String outputPath) {
        try {
            if (inputPath.endsWith(".docx")) {
                convertDocxToPdf(inputPath, outputPath);
            } else if (inputPath.endsWith(".xlsx")) {
                convertXlsxToPdf(inputPath, outputPath);
            } else if (inputPath.endsWith(".pptx")) {
                convertPptxToPdf(inputPath, outputPath);
            } else if (inputPath.endsWith(".hwp")) {
                convertHwpToPdf(inputPath, outputPath);
            } else if (inputPath.endsWith(".hwpx")) {
                convertHwpxToPdf(inputPath, outputPath);
            } else {
                System.out.println("지원하지 않는 형식: " + inputPath);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ------------------ DOCX ------------------
    private void convertDocxToPdf(String inputPath, String outputPath) throws Exception {
        try {
            FileInputStream fis = new FileInputStream(inputPath);
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(fis);
            OutputStream os = new FileOutputStream(outputPath);
            Conversion c = new Conversion(wordMLPackage);
            // PdfSettings 대신 null 사용
            c.output(os, null);
            os.close();
            fis.close();
            System.out.println("DOCX → PDF 변환 완료: " + outputPath);
        } catch (Exception e) {
            System.err.println("DOCX 변환 실패: " + e.getMessage());
            e.printStackTrace();
            // 실패 시 텍스트 문서 생성
            createTextPdf(outputPath, "DOCX 변환 실패: " + e.getMessage());
        }
    }

    // XLSX 파일을 PDF로 변환 (간단한 테이블 형태)
    private void convertXlsxToPdf(String inputPath, String outputPath) throws Exception {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(inputPath));
            PDDocument document = new PDDocument();
            
            for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                StringBuilder content = new StringBuilder();
                content.append("Sheet: ").append(workbook.getSheetName(sheetIdx)).append("\n\n");
                
                workbook.getSheetAt(sheetIdx).forEach(row -> {
                    row.forEach(cell -> {
                        content.append(cell.toString()).append("\t");
                    });
                    content.append("\n");
                });
                
                // PDF 페이지 추가
                PDPage page = new PDPage();
                document.addPage(page);
                PDPageContentStream contentStream = new PDPageContentStream(document, page);
                contentStream.beginText();
                contentStream.setFont(PDType1Font.HELVETICA, 10);
                contentStream.newLineAtOffset(50, 750);
                
                String[] lines = content.toString().split("\n");
                for (String line : lines) {
                    contentStream.showText(line.length() > 100 ? line.substring(0, 100) : line);
                    contentStream.newLine();
                }
                contentStream.endText();
                contentStream.close();
            }
            
            document.save(new FileOutputStream(outputPath));
            document.close();
            workbook.close();
            System.out.println("XLSX → PDF 변환 완료: " + outputPath);
        } catch (Exception e) {
            System.err.println("XLSX 변환 실패: " + e.getMessage());
            e.printStackTrace();
            createTextPdf(outputPath, "XLSX 변환 실패: " + e.getMessage());
        }
    }

    // ------------------ PPTX ------------------
    private void convertPptxToPdf(String inputPath, String outputPath) throws Exception {
        try {
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputPath));
            PDDocument document = new PDDocument();
            
            for (XSLFSlide slide : ppt.getSlides()) {
                PDPage page = new PDPage();
                document.addPage(page);
                PDPageContentStream contentStream = new PDPageContentStream(document, page);
                contentStream.beginText();
                contentStream.setFont(PDType1Font.HELVETICA, 14);
                contentStream.newLineAtOffset(50, 750);
                String title = slide.getTitle() != null ? slide.getTitle() : "Slide";
                contentStream.showText(title);
                contentStream.endText();
                contentStream.close();
            }
            
            document.save(new FileOutputStream(outputPath));
            document.close();
            ppt.close();
            System.out.println("PPTX → PDF 변환 완료: " + outputPath);
        } catch (Exception e) {
            System.err.println("PPTX 변환 실패: " + e.getMessage());
            e.printStackTrace();
            createTextPdf(outputPath, "PPTX 변환 실패: " + e.getMessage());
        }
    }

    // ------------------ HWP ------------------
    private void convertHwpToPdf(String inputPath, String outputPath) throws Exception {
        try {
            // HWP 파일에서 텍스트 추출
            Object bodyTextObj = HWPReader.fromFile(inputPath).getBodyText();
            String text = bodyTextObj != null ? bodyTextObj.toString() : "HWP 파일에서 텍스트를 추출할 수 없습니다.";
            
            if (text == null || text.isEmpty()) {
                text = "HWP 파일에서 텍스트를 추출할 수 없습니다.";
            }
            
            // 텍스트를 PDF로 저장
            createTextPdf(outputPath, text);
            System.out.println("HWP → PDF 변환 완료: " + outputPath);
        } catch (Exception e) {
            System.err.println("HWP 변환 실패: " + e.getMessage());
            e.printStackTrace();
            createTextPdf(outputPath, "HWP 변환 실패: " + e.getMessage());
        }
    }

    // ------------------ HWPX ------------------
    private void convertHwpxToPdf(String inputPath, String outputPath) throws Exception {
        // HWPX는 완전한 지원이 어려움. 텍스트 추출만 시도
        try {
            // hwpxlib이 없으면 간단한 오류 메시지 생성
            String text = "HWPX 형식은 완전히 지원되지 않습니다.\n" +
                    "문서 변환을 위해 HWP 또는 DOCX 형식을 사용해주세요.";
            createTextPdf(outputPath, text);
            System.out.println("HWPX 변환 미지원 - 안내문 PDF 생성: " + outputPath);
        } catch (Exception e) {
            System.err.println("HWPX 처리 실패: " + e.getMessage());
            e.printStackTrace();
            createTextPdf(outputPath, "HWPX 처리 실패: " + e.getMessage());
        }
    }

    // ------------------ HTML → PDF ------------------
    private void htmlToPdf(String htmlContent, String pdfPath) throws Exception {
        // 간단한 텍스트 PDF 생성 (HTML 파싱 불가)
        try {
            // HTML 태그 제거
            String plainText = htmlContent.replaceAll("<[^>]*>", "");
            createTextPdf(pdfPath, plainText);
        } catch (Exception e) {
            System.err.println("HTML → PDF 변환 실패: " + e.getMessage());
            e.printStackTrace();
            createTextPdf(pdfPath, "HTML 변환 실패: " + e.getMessage());
        }
    }

    // 간단한 텍스트 PDF 생성
    private void createTextPdf(String pdfPath, String text) throws Exception {
        try {
            PDDocument document = new PDDocument();
            PDPage page = new PDPage();
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA, 12);
            contentStream.newLineAtOffset(50, 750);
            contentStream.showText(text != null ? text : "No content");
            contentStream.endText();
            contentStream.close();

            document.save(new FileOutputStream(pdfPath));
            document.close();
            System.out.println("텍스트 PDF 생성: " + pdfPath);
        } catch (Exception e) {
            System.err.println("PDF 생성 실패: " + e.getMessage());
            throw e;
        }
    }
}
