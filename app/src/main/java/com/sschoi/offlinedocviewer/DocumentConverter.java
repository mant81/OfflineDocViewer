package com.sschoi.offlinedocviewer;

import android.content.Context;
import android.net.Uri;

import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import kr.dogfoot.hwplib.reader.HWPReader;

import java.io.*;

public class DocumentConverter {

    private final Context context;

    public DocumentConverter(Context context) {
        this.context = context;
    }

    // 기존 String 기반 변환 유지
    public void convertToPdf(String inputPath, String outputPath) throws Exception {
        if (inputPath.endsWith(".docx")) convertDocxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".xlsx")) convertXlsxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".pptx")) convertPptxToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".hwp")) convertHwpToPdf(inputPath, outputPath);
        else if (inputPath.endsWith(".hwpx")) convertHwpxToPdf(outputPath);
        else createTextPdf(outputPath, "지원하지 않는 파일 형식입니다: " + inputPath);
    }

    // 스트림 버전
    public void convertToPdf(InputStream input, OutputStream output) throws Exception {
        // 임시 파일 생성
        File tempInput = File.createTempFile("input", null, context.getCacheDir());
        try (FileOutputStream fos = new FileOutputStream(tempInput)) {
            byte[] buf = new byte[8192];
            int len;
            while ((len = input.read(buf)) > 0) {
                fos.write(buf, 0, len);
            }
        }

        // 기존 String 버전 재사용
        File tempOutput = File.createTempFile("output", ".pdf", context.getCacheDir());
        convertToPdf(tempInput.getAbsolutePath(), tempOutput.getAbsolutePath());

        // 결과 복사
        try (FileInputStream fis = new FileInputStream(tempOutput)) {
            byte[] buf = new byte[8192];
            int len;
            while ((len = fis.read(buf)) > 0) {
                output.write(buf, 0, len);
            }
        }

        tempInput.deleteOnExit();
        tempOutput.deleteOnExit();
    }

    // DOCX
    private void convertDocxToPdf(String inputPath, String outputPath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(inputPath))) {
            StringBuilder text = new StringBuilder();
            doc.getParagraphs().forEach(p -> text.append(p.getText()).append("\n"));
            createTextPdf(outputPath, text.toString());
        }
    }

    // XLSX
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

    // PPTX
    private void convertPptxToPdf(String inputPath, String outputPath) throws Exception {
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputPath))) {
            StringBuilder text = new StringBuilder();
            int idx = 1;
            for (XSLFSlide slide : ppt.getSlides()) {
                text.append("Slide ").append(idx++).append("\n");
                if (slide.getTitle() != null)
                    text.append("Title: ").append(slide.getTitle()).append("\n");
                text.append("\n");
            }
            createTextPdf(outputPath, text.toString());
        }
    }

    // HWP
    private void convertHwpToPdf(String inputPath, String outputPath) throws Exception {
        Object bodyTextObj = HWPReader.fromFile(inputPath).getBodyText();
        String text = bodyTextObj != null ? bodyTextObj.toString() : "HWP 텍스트 추출 실패";
        createTextPdf(outputPath, text);
    }

    // HWPX
    private void convertHwpxToPdf(String outputPath) throws Exception {
        createTextPdf(outputPath, "HWPX 형식은 지원되지 않습니다.\nHWP 또는 DOCX를 사용하세요.");
    }

    // PDF 생성
    private void createTextPdf(String pdfPath, String text) throws Exception {
        try (OutputStream os = new FileOutputStream(pdfPath)) {
            Document pdf = new Document();
            PdfWriter.getInstance(pdf, os);
            pdf.open();
            for (String line : text.split("\n")) {
                pdf.add(new Paragraph(line));
            }
            pdf.close();
        }
    }
}
