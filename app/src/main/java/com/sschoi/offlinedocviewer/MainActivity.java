package com.sschoi.offlinedocviewer;

import android.Manifest;
import android.content.pm.PackageManager;
import android.widget.Toast;
import android.os.Bundle;
import android.os.Environment;
import android.widget.Button;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import com.github.barteksc.pdfviewer.PDFView;
import java.io.File;

public class MainActivity extends AppCompatActivity {

    PDFView pdfView;
    Button btnConvert;
    String inputPath, outputPath;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        pdfView = findViewById(R.id.pdfView);
        btnConvert = findViewById(R.id.btnConvert);

        // 권한 체크: 읽기/쓰기 둘다 요청
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED
                || ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE}, 1);
        }

        inputPath = Environment.getExternalStorageDirectory() + "/Documents/sample.docx"; // 변경 가능
        outputPath = Environment.getExternalStorageDirectory() + "/Documents/output.pdf";

        // 입력 파일 존재 여부 확인 — 파일이 없으면 변환 버튼 비활성화 및 알림
        File inFile = new File(inputPath);
        if (!inFile.exists()) {
            btnConvert.setEnabled(false);
            Toast.makeText(this, "입력 파일이 없습니다:\n" + inputPath + "\n파일을 기기/에뮬레이터에 넣어주세요.", Toast.LENGTH_LONG).show();
            return;
        }

        btnConvert.setOnClickListener(v -> {
            // 권한 재확인
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED
                    || ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                Toast.makeText(this, "저장소 권한이 필요합니다.", Toast.LENGTH_SHORT).show();
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE}, 1);
                return;
            }

            try {
                DocumentConverter converter = new DocumentConverter();
                converter.convertToPdf(inputPath, outputPath);

                File pdfFile = new File(outputPath);
                if (pdfFile.exists()) {
                    pdfView.fromFile(pdfFile)
                            .enableSwipe(true)
                            .enableDoubletap(true)
                            .load();
                } else {
                    Toast.makeText(this, "PDF 변환 실패: 출력 파일이 생성되지 않았습니다.", Toast.LENGTH_LONG).show();
                }
            } catch (Exception e) {
                e.printStackTrace();
                Toast.makeText(this, "변환 중 예외 발생: " + e.getClass().getSimpleName() + " - " + e.getMessage(), Toast.LENGTH_LONG).show();
            }
        });
    }
}
