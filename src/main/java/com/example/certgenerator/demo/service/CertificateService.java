
package com.example.certgenerator.demo.service;

import com.example.certgenerator.demo.model.CertificateData;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.awt.Color;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class CertificateService {

    public List<CertificateData> extractDataFromExcel(InputStream inputStream) throws IOException {
        List<CertificateData> dataList = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        Row headerRow = rowIterator.next();
        int nameCol = -1, activityCol = -1, dateCol = -1;

        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue().trim().toLowerCase();
            if (header.contains("name")) nameCol = cell.getColumnIndex();
            else if (header.contains("activity") || header.contains("event")) activityCol = cell.getColumnIndex();
            else if (header.contains("date")) dateCol = cell.getColumnIndex();
        }

        if (nameCol == -1 || activityCol == -1 || dateCol == -1) {
            workbook.close();
            throw new IOException("Required columns (Name, Activity, Date) not found in header row.");
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            CertificateData data = new CertificateData();

            data.setName(getCellString(row.getCell(nameCol)));
            data.setActivity(getCellString(row.getCell(activityCol)));

            Cell dateCell = row.getCell(dateCol);
            if (dateCell != null) {
                if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                    Date date = dateCell.getDateCellValue();
                    SimpleDateFormat formatter = new SimpleDateFormat("dd MMMM yyyy");
                    data.setDate(formatter.format(date));
                } else {
                    data.setDate(getCellString(dateCell));
                }
            } else {
                data.setDate("Missing");
            }

            dataList.add(data);
        }

        workbook.close();
        return dataList;
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }

    public ByteArrayOutputStream generateCertificate(
            CertificateData data,
            InputStream bgImageStream,
            List<MultipartFile> signImages,
            List<String> signNames,
            String mainTextTemplate,
            String headingText,
            String headingFont, String bodyFont, int nameFontSize, String nameFontFamily) throws IOException {

        PDDocument document = new PDDocument();
        PDRectangle landscape = new PDRectangle(PDRectangle.LETTER.getHeight(), PDRectangle.LETTER.getWidth());
        PDPage page = new PDPage(landscape);
        document.addPage(page);

        PDPageContentStream contentStream = new PDPageContentStream(document, page);

        PDImageXObject bgImage = PDImageXObject.createFromByteArray(document, bgImageStream.readAllBytes(), "bg");
        contentStream.drawImage(bgImage, 0, 0, page.getMediaBox().getWidth(), page.getMediaBox().getHeight());

        contentStream.beginText();
        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 30);
        contentStream.setNonStrokingColor(Color.DARK_GRAY);
        float headingWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(headingText) / 1000 * 30;
        contentStream.newLineAtOffset((landscape.getWidth() - headingWidth) / 2, landscape.getHeight() - 230);
        contentStream.showText(headingText);
        contentStream.endText();

        String rawText = mainTextTemplate
                .replace("{activity}", data.getActivity())
                .replace("{date}", data.getDate());

        String[] parts = rawText.split("\\{name}");
        String beforeName = parts[0];
        String afterName = parts.length > 1 ? parts[1] : "";

        float currentY = 320;
        float maxWidth = landscape.getWidth() - 200;

        java.util.List<String> beforeLines = wrapText(beforeName, PDType1Font.HELVETICA_BOLD, 20, maxWidth);
        java.util.List<String> afterLines = wrapText(afterName, PDType1Font.HELVETICA_BOLD, 20, maxWidth);

        for (String line : beforeLines) {
            float lineWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(line) / 1000 * 20;
            float startX = (landscape.getWidth() - lineWidth) / 2;

            contentStream.beginText();
            contentStream.setNonStrokingColor(Color.BLACK);
            contentStream.setFont(PDType1Font.HELVETICA_BOLD, 20);
            contentStream.newLineAtOffset(startX, currentY);
            contentStream.showText(line);
            contentStream.endText();

            currentY -= 24;
        }

        float nameWidth = PDType1Font.TIMES_ITALIC.getStringWidth(data.getName()) / 1000 * 28;
        float nameX = (landscape.getWidth() - nameWidth) / 2;

        contentStream.beginText();
        contentStream.setFont(PDType1Font.TIMES_ITALIC, 28);
        contentStream.setNonStrokingColor(Color.BLUE);
        contentStream.newLineAtOffset(nameX, currentY);
        contentStream.showText(data.getName());
        contentStream.endText();

        currentY -= 30;

        for (String line : afterLines) {
            float lineWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(line) / 1000 * 20;
            float startX = (landscape.getWidth() - lineWidth) / 2;

            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA_BOLD, 20);
            contentStream.setNonStrokingColor(Color.BLACK);
            contentStream.newLineAtOffset(startX, currentY);
            contentStream.showText(line);
            contentStream.endText();

            currentY -= 24;
        }

        float startX = 100;
        float yPosition = 150;
        for (int i = 0; i < signImages.size(); i++) {
            MultipartFile signFile = signImages.get(i);
            String signName = signNames.get(i);
            PDImageXObject signImg = PDImageXObject.createFromByteArray(document, signFile.getBytes(), "sign" + i);
            contentStream.drawImage(signImg, startX, yPosition, 100, 50);

            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA_OBLIQUE, 14);
            contentStream.setNonStrokingColor(Color.DARK_GRAY);
            contentStream.newLineAtOffset(startX, yPosition - 20);
            contentStream.showText(signName);
            contentStream.endText();

            startX += 150;
        }

        contentStream.close();

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.save(outputStream);
        document.close();

        return outputStream;
    }

    private java.util.List<String> wrapText(String text, PDType1Font font, int fontSize, float maxWidth) throws IOException {
        java.util.List<String> lines = new ArrayList<>();
        String[] words = text.split(" ");
        StringBuilder line = new StringBuilder();

        for (String word : words) {
            String testLine = line.length() == 0 ? word : line + " " + word;
            float size = font.getStringWidth(testLine) / 1000 * fontSize;
            if (size > maxWidth) {
                lines.add(line.toString());
                line = new StringBuilder(word);
            } else {
                line = new StringBuilder(testLine);
            }
        }
        if (!line.toString().isEmpty()) {
            lines.add(line.toString());
        }

        return lines;
    }
}


