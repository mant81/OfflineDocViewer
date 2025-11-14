package com.sschoi.offlinedocviewer;

import kr.dogfoot.hwplib.reader.HWPReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.*;

public class DocumentConverter {

    public void convertToPdf(String inputPath, String outputPath) throws Exception {
        if (inputPath.endsWith(".docx")) convertDocxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".xlsx")) convertXlsxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".pptx")) convertPptxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".hwp")) convertHwpToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".hwpx")) convertHwpxToPdf(inputPath, outputPath);
        else createTextPdf(outputPath, "지원하지 않는 파일 형식: " + inputPath);
    }

    private void convertDocxToPdf(String inputPath, String outputPath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(inputPath))) {
            StringBuilder text = new StringBuilder();
            doc.getParagraphs().forEach(p -> text.append(p.getText()).append("\n"));
            createTextPdf(outputPath, text.toString());
        }
    }

    private void convertXlsxToPdf(String inputPath, String outputPath) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(inputPath))) {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                text.append("Sheet: ").append(workbook.getSheetName(i)).append("\n");
                workbook.getSheetAt(i).forEach(row -> {
                    row.forEach(cell -> text.append(cell.toString()).append("\t"));
                    text.append("\n");
                });
                text.append("\n");
            }
            createTextPdf(outputPath, text.toString());
        }
    }

    private void convertPptxToPdf(String inputPath, String outputPath) throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputPath));
             PDDocument document = new PDDocument()) {

            for (XSLFSlide slide : ppt.getSlides()) {
                PDPage page = new PDPage();
                document.addPage(page);

                try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
                    cs.beginText();
                    cs.setFont(PDType1Font.HELVETICA_BOLD, 14);
                    cs.newLineAtOffset(50, 750);
                    cs.showText(slide.getTitle() != null ? slide.getTitle() : "Slide");
                    cs.endText();
                }
            }

            document.save(outputPath);
        }
    }

    private void convertHwpToPdf(String inputPath, String outputPath) throws Exception {
        Object bodyTextObj = HWPReader.fromFile(inputPath).getBodyText();
        String text = bodyTextObj != null ? bodyTextObj.toString() : "HWP 텍스트 추출 실패";
        createTextPdf(outputPath, text);
    }

    private void convertHwpxToPdf(String inputPath, String outputPath) throws Exception {
        createTextPdf(outputPath, "HWPX 형식은 지원되지 않습니다.\nHWP 또는 DOCX를 사용하세요.");
    }

    private void createTextPdf(String pdfPath, String text) throws Exception {
        try (PDDocument document = new PDDocument()) {
            PDPage page = new PDPage();
            document.addPage(page);

            try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
                cs.beginText();
                cs.setFont(PDType1Font.HELVETICA, 12);
                cs.newLineAtOffset(50, 750);

                String[] lines = text.split("\n");
                for (String line : lines) cs.showText(line != null ? line : "");

                cs.endText();
            }

            document.save(pdfPath);
        }
    }
}
