package com.example.certgenerator.demo.controller;

import com.example.certgenerator.demo.model.CertificateData;
import com.example.certgenerator.demo.service.CertificateService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@CrossOrigin(origins = "*") // Allow React frontend
@RestController
@RequestMapping("/cert")
public class CertificateController {

    @Autowired
    private CertificateService certificateService;

    @PostMapping("/generate")
    public ResponseEntity<ByteArrayResource> generateCertificates(
            @RequestParam("excel") MultipartFile excel,
            @RequestParam("bg") MultipartFile bgImage,
            @RequestParam("signs") List<MultipartFile> signImages,
            @RequestParam("signNames") List<String> signNames,
            @RequestParam("mainText") String mainText,
            @RequestParam("headingText") String headingText,
            @RequestParam("headingFont") String headingFont,
            @RequestParam("bodyFont") String bodyFont,
            @RequestParam("nameFontSize") int nameFontSize,
            @RequestParam("nameFontFamily") String nameFontFamily
    ) {
        try {
            // Extract student data from Excel
            List<CertificateData> recipients = certificateService.extractDataFromExcel(excel.getInputStream());

            // Prepare ZIP output stream
            ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
            ZipOutputStream zip = new ZipOutputStream(zipOutputStream);

            // Generate a certificate for each recipient
            for (CertificateData data : recipients) {
                ByteArrayOutputStream pdf = certificateService.generateCertificate(
                        data,
                        bgImage.getInputStream(),
                        signImages,
                        signNames,
                        mainText,
                        headingText,
                        headingFont,
                        bodyFont,
                        nameFontSize,
                        nameFontFamily
                );

                // Add PDF to zip
                ZipEntry entry = new ZipEntry(data.getName().replaceAll("\\s+", "_") + ".pdf");
                zip.putNextEntry(entry);
                zip.write(pdf.toByteArray());
                zip.closeEntry();
            }

            zip.close();

            // Prepare zipped certificates as downloadable
            ByteArrayResource resource = new ByteArrayResource(zipOutputStream.toByteArray());

            HttpHeaders headers = new HttpHeaders();
            headers.setContentDisposition(ContentDisposition.attachment().filename("certificates.zip").build());
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

            return new ResponseEntity<>(resource, headers, HttpStatus.OK);

        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).build();
        }
    }
}

