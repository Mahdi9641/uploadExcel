package com.example.upload_excel;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.security.GeneralSecurityException;


@RestController
public class ExcelController {

    private final ExcelFileRepository excelFileRepository;

    public ExcelController(ExcelFileRepository excelFileRepository) {
        this.excelFileRepository = excelFileRepository;
    }

    @PostMapping(consumes = {MediaType.MULTIPART_FORM_DATA_VALUE}, produces = {MediaType.APPLICATION_JSON_VALUE})
    public String uploadExcelBase64(@Validated @RequestParam("file") MultipartFile file, @RequestParam("name") String name) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheet("test");

            if (sheet != null) {
                // رمز عبور برای محافظت از فایل
                String password = "pdr@102030";

                // ایجاد POIFSFileSystem برای ذخیره‌سازی فایل رمزگذاری شده
                POIFSFileSystem fs = new POIFSFileSystem();

                // ایجاد اطلاعات رمزگذاری با حالت Agile
                EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
                Encryptor encryptor = info.getEncryptor();
                encryptor.confirmPassword(password);

                // ذخیره فایل اکسل با محافظت از رمز عبور
                try (OutputStream os = encryptor.getDataStream(fs)) {
                    workbook.write(os);
                } catch (GeneralSecurityException e) {
                    throw new RuntimeException(e.getMessage());
                }

                byte[] fileData;
                // نوشتن POIFSFileSystem به فایل
                try (FileOutputStream fos = new FileOutputStream("C:/Users/m_khodam/Desktop/excel/" + name + ".xlsx")) {
                    fs.writeFilesystem(fos);
                }

                try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                    fs.writeFilesystem(bos);
                    fileData = bos.toByteArray();
                }
                ExcelFile excelFile = new ExcelFile();
                excelFile.setName(name);
                excelFile.setData(fileData);
                excelFileRepository.save(excelFile);

                return "File uploaded, saved with password protection, and stored in the database successfully!";
            } else {
                return "Sheet with name 'test' not found in the Excel file.";
            }
        } catch (IOException e) {
            return "Error processing the file: " + e.getMessage();
        }
    }

    @GetMapping("/download/{id}")
    public ResponseEntity<Resource> downloadFile(@PathVariable Long id) {
        ExcelFile excelFile = excelFileRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("File not found with id " + id));

        ByteArrayResource resource = new ByteArrayResource(excelFile.getData());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + excelFile.getName() + ".xlsx\"")
                .contentLength(resource.contentLength())
                .body(new InputStreamResource(resource));
    }

}


