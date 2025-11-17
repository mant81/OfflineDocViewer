package com.sschoi.offlinedocviewer;

import android.Manifest;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import com.github.barteksc.pdfviewer.PDFView;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {

    private DocumentConverter converter;
    private ActivityResultLauncher<Intent> filePickerLauncher;
    private PDFView pdfView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        converter = new DocumentConverter(this);
        pdfView = findViewById(R.id.pdfView);

        // 권한 요청
        ActivityCompat.requestPermissions(this,
                new String[]{Manifest.permission.READ_EXTERNAL_STORAGE,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);

        filePickerLauncher = registerForActivityResult(
                new ActivityResultContracts.StartActivityForResult(),
                result -> {
                    if (result.getResultCode() == RESULT_OK && result.getData() != null) {
                        Uri uri = result.getData().getData();
                        if (uri != null) handleFile(uri);
                    }
                });

        Button btnPick = findViewById(R.id.btnPickFile);
        btnPick.setOnClickListener(v -> pickFile());
    }

    private void pickFile() {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        intent.setType("*/*");
        filePickerLauncher.launch(intent);
    }

    private void handleFile(Uri uri) {
        try {
            // 임시 PDF 파일
            String tempName = "converted_" + System.currentTimeMillis() + ".pdf";
            File tempFile = new File(getCacheDir(), tempName);

            // PDF 변환
            try (OutputStream out = new FileOutputStream(tempFile)) {
                converter.convertToPdf(uri, out);
            }

            if (!tempFile.exists() || tempFile.length() == 0) {
                Toast.makeText(this, "PDF 변환 실패: 파일 생성 안됨", Toast.LENGTH_LONG).show();
                return;
            }

            // PDFView 소프트웨어 렌더링 적용
            pdfView.setLayerType(View.LAYER_TYPE_SOFTWARE, null);

            // PDFView 로드
            pdfView.fromFile(tempFile)
                    .enableSwipe(true)
                    .swipeHorizontal(false)
                    .enableDoubletap(true)
                    .enableAntialiasing(true)
                    .load();

            Toast.makeText(this, "PDF 변환 완료!", Toast.LENGTH_SHORT).show();

        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "PDF 변환 실패: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }
}
