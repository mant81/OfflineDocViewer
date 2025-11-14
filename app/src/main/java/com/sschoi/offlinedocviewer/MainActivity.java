package com.sschoi.offlinedocviewer;

import android.Manifest;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.DocumentsContract;
import android.widget.Button;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import java.io.File;

public class MainActivity extends AppCompatActivity {

    private final DocumentConverter converter = new DocumentConverter();

    private final ActivityResultLauncher<Intent> filePickerLauncher =
            registerForActivityResult(
                    new ActivityResultContracts.StartActivityForResult(),
                    result -> {
                        if (result.getResultCode() == RESULT_OK && result.getData() != null) {
                            Uri uri = result.getData().getData();
                            if (uri != null) handleFile(uri);
                        }
                    });

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // 권한 요청
        ActivityCompat.requestPermissions(this,
                new String[]{Manifest.permission.READ_EXTERNAL_STORAGE,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);

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
            String fileName = DocumentsContract.getDocumentId(uri).replace(":", "_");
            File outputDir = new File(Environment.getExternalStorageDirectory(), "OfflineDocViewer");
            if (!outputDir.exists()) outputDir.mkdirs();

            File outputFile = new File(outputDir, fileName + ".pdf");
            converter.convertToPdf(uri.getPath(), outputFile.getAbsolutePath());

            Toast.makeText(this, "PDF 변환 완료: " + outputFile.getAbsolutePath(), Toast.LENGTH_LONG).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, "PDF 변환 실패: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }
}
