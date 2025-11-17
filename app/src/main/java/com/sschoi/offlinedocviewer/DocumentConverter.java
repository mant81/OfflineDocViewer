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

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class DocumentConverter {

    private final Context context;

    public DocumentConverter(Context context) {
        this.context = context;
    }

    // SAF content:// 또는 file:// 모두 처리
    private InputStream openInput(String pathOrUri) throws Exception {
        if (pathOrUri.startsWith("content://")) {
            return context.getContentResolver().openInputStream(Uri.parse(pathOrUri));
        } else {
            return new FileInputStream(pathOrUri);
        }
    }

    // 메인 변환 함수 (OutputStream으로 PDF 생성)
    public void convertToPdf(Uri inputUri, OutputStream outStream) throws Exception {
        String path = inputUri.toString();

        if (path.endsWith(".docx")) convertDocxToPdf(inputUri, outStream);
        else if (path.endsWith(".xlsx")) convertXlsxToPdf(inputUri, outStream);
        else if (path.endsWith(".pptx")) convertPptxToPdf(inputUri, outStream);
        else if (path.endsWith(".hwp")) convertHwpToPdf(inputUri, outStream);
        else createTextPdf(outStream, "지원하지 않는 파일 형식입니다.");
    }

    // DOCX
    private void convertDocxToPdf(Uri uri, OutputStream out) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri);
             XWPFDocument doc = new XWPFDocument(is)) {

            StringBuilder text = new StringBuilder();
            doc.getParagraphs().forEach(p -> text.append(p.getText()).append("\n"));

            createTextPdf(out, text.toString());
        }
    }

    // XLSX
    private void convertXlsxToPdf(Uri uri, OutputStream out) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri);
             XSSFWorkbook workbook = new XSSFWorkbook(is)) {

            StringBuilder text = new StringBuilder();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                text.append("Sheet: ").append(workbook.getSheetName(i)).append("\n");
                workbook.getSheetAt(i).forEach(row -> {
                    row.forEach(cell -> text.append(cell.toString()).append("\t"));
                    text.append("\n");
                });
                text.append("\n");
            }
            createTextPdf(out, text.toString());
        }
    }

    // PPTX
    private void convertPptxToPdf(Uri uri, OutputStream out) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri);
             XMLSlideShow ppt = new XMLSlideShow(is)) {

            StringBuilder text = new StringBuilder();
            int idx = 1;
            for (XSLFSlide slide : ppt.getSlides()) {
                text.append("Slide ").append(idx++).append("\n");
                try {
                    if (slide.getTitle() != null)
                        text.append("Title: ").append(slide.getTitle()).append("\n");
                } catch (Exception ignored) {}
                text.append("\n");
            }

            createTextPdf(out, text.toString());
        }
    }

    // HWP
    private void convertHwpToPdf(Uri uri, OutputStream out) throws Exception {
        try (InputStream is = context.getContentResolver().openInputStream(uri)) {
            String text = "HWP 텍스트 추출 실패";
            try {
                text = HWPReader.fromInputStream(new BufferedInputStream(is))
                        .getBodyText()
                        .toString();
            } catch (Exception e) {
                text = "HWP 파싱 오류: " + e.getMessage();
            }
            createTextPdf(out, text);
        }
    }

    // 공통 PDF 생성
    private void createTextPdf(OutputStream out, String text) throws Exception {
        Document pdf = new Document();
        PdfWriter.getInstance(pdf, out);
        pdf.open();
        for (String line : text.split("\n")) {
            pdf.add(new Paragraph(line));
        }
        pdf.close();
    }
}
