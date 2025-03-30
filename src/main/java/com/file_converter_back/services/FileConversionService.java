package com.file_converter_back.services;

import com.file_converter_back.DTO.ConversionOption;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.text.PDFTextStripper;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.pdfbox.pdmodel.font.PDType1Font;

@Service
public class FileConversionService {

    public List<ConversionOption> getAvailableConversions(String sourceFormat) {
        List<ConversionOption> options = new ArrayList<>();



        switch (sourceFormat.toLowerCase()) {
            case "docx":
                options.add(new ConversionOption("Documento PowerPoint", "pptx"));
                options.add(new ConversionOption("Documento PDF", "pdf"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                options.add(new ConversionOption("Documento RTF", "rtf"));
                options.add(new ConversionOption("Documento ODT", "odt"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("XML", "xml"));
                break;

            case "xlsx":
                options.add(new ConversionOption("CSV", "csv"));
                options.add(new ConversionOption("PDF", "pdf"));
                options.add(new ConversionOption("Texto Simples (TSV)", "tsv"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("XML", "xml"));
                break;

            case "pptx":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("PDF", "pdf"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("JSON", "json"));
                break;

            case "csv":
                options.add(new ConversionOption("Excel", "xlsx"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("XML", "xml"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("PDF", "pdf"));
                break;

            case "pdf":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("PowerPoint", "pptx"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("JSON", "json"));
                break;

            case "txt":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("PDF", "pdf"));
                options.add(new ConversionOption("HTML", "html"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("XML", "xml"));

                break;

            case "html":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("PDF", "pdf"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("XML", "xml"));
                break;

            case "json":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("Excel", "xlsx"));
                options.add(new ConversionOption("CSV", "csv"));
                options.add(new ConversionOption("XML", "xml"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                break;

            case "xml":
                options.add(new ConversionOption("Documento Word", "docx"));
                options.add(new ConversionOption("Excel", "xlsx"));
                options.add(new ConversionOption("JSON", "json"));
                options.add(new ConversionOption("Texto Simples", "txt"));
                break;
        }

        return options;
    }

    public byte[] convert(byte[] fileBytes, String sourceFormat, String targetFormat) throws IOException {
        switch (sourceFormat + "-" + targetFormat) {
            // Conversões de Excel
            case "xlsx-csv": return convertXlsxToCsv(fileBytes);
            case "xlsx-json": return convertXlsxToJson(fileBytes);
            case "xlsx-pdf": return convertXlsxToPdf(fileBytes);
            case "xlsx-xml": return convertXlsxToXml(fileBytes);
            case "xlsx-html": return convertXlsxToHtml(fileBytes);
            case "xlsx-tsv": return convertXlsxToTsv(fileBytes);


            // Conversões de CSV
            case "csv-xlsx": return convertCsvToXlsx(fileBytes);
            case "csv-json": return convertCsvToJson(fileBytes);
            case "csv-xml": return convertCsvToXml(fileBytes);
            case "csv-html": return convertCsvToHtml(fileBytes);
            case "csv-pdf": return convertCsvToPdf(fileBytes);

            // Conversões de Word
            case "docx-txt": return convertDocxToTxt(fileBytes);
            case "docx-pdf": return convertDocxToPdf(fileBytes);
            case "docx-pptx": return convertDocxToPptx(fileBytes);
            case "docx-html": return convertDocxToHtml(fileBytes);
            case "docx-json": return convertDocxToJson(fileBytes);
            case "docx-xml": return convertDocxToXml(fileBytes);
            case "docx-rtf": return convertDocxToRtf(fileBytes);

            // Conversões de PowerPoint
            case "pptx-docx": return convertPptxToDocx(fileBytes);
            case "pptx-pdf": return convertPptxToPdf(fileBytes);
            case "pptx-txt": return convertPptxToTxt(fileBytes);
            case "pptx-html": return convertPptxToHtml(fileBytes);
            case "pptx-json": return convertPptxToJson(fileBytes);


            // Conversões de PDF
            case "pdf-txt": return convertPdfToTxt(fileBytes);
            case "pdf-docx": return convertPdfToDocx(fileBytes);
            case "pdf-pptx": return convertPdfToPptx(fileBytes);
            case "pdf-html": return convertPdfToHtml(fileBytes);
            case "pdf-json": return convertPdfToJson(fileBytes);


            // Conversões de texto simples
            case "txt-pdf": return convertTxtToPdf(fileBytes);
            case "txt-docx": return convertTxtToDocx(fileBytes);
            case "txt-html": return convertTxtToHtml(fileBytes);
            case "txt-json": return convertTxtToJson(fileBytes);
            case "txt-xml": return convertTxtToXml(fileBytes);

            // Conversões de HTML
            case "html-txt": return convertHtmlToTxt(fileBytes);
            case "html-pdf": return convertHtmlToPdf(fileBytes);
            case "html-docx": return convertHtmlToDocx(fileBytes);
            case "html-json": return convertHtmlToJson(fileBytes);
            case "html-xml": return convertHtmlToXml(fileBytes);

            // Conversões de JSON
            case "json-xlsx": return convertJsonToXlsx(fileBytes);
            case "json-csv": return convertJsonToCsv(fileBytes);
            case "json-xml": return convertJsonToXml(fileBytes);
            case "json-txt": return convertJsonToTxt(fileBytes);

            // Conversões de XML

            case "xml-txt": return convertXmlToTxt(fileBytes);

            default:
                throw new UnsupportedOperationException(
                        "Conversão de " + sourceFormat + " para " + targetFormat + " não implementada");
        }
    }

    private byte[] convertXlsxToJson(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes))) {
            Sheet sheet = workbook.getSheetAt(0);
            List<Map<String, String>> data = new ArrayList<>();
            Row headerRow = sheet.getRow(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Map<String, String> rowData = new HashMap<>();

                for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                    String header = headerRow.getCell(j).getStringCellValue();
                    String value = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).toString();
                    rowData.put(header, value);
                }
                data.add(rowData);
            }

            return new ObjectMapper().writeValueAsBytes(data);
        }
    }

    private byte[] convertTxtToPdf(byte[] txtBytes) throws IOException {
        try (PDDocument document = new PDDocument()) {
            PDPage page = new PDPage();
            document.addPage(page);

            try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                // Usar uma fonte mais compatível (Helvetica)
                contentStream.setFont(PDType1Font.HELVETICA, 12);

                // Iniciar o bloco de texto
                contentStream.beginText();

                // Posicionar o texto
                contentStream.newLineAtOffset(25, 700);

                String text = new String(txtBytes, StandardCharsets.UTF_8);
                String[] lines = text.split("\n");

                for (String line : lines) {
                    // Substituir caracteres não suportados
                    line = line.replaceAll("[^\\x00-\\x7F]", "?");

                    // Escrever a linha
                    contentStream.showText(line);

                    // Mover para a próxima linha
                    contentStream.newLineAtOffset(0, -15);
                }

                // Finalizar o bloco de texto
                contentStream.endText();
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            document.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertCsvToXlsx(byte[] csvBytes) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(csvBytes)))) {

            Sheet sheet = workbook.createSheet("Dados");
            String line;
            int rowNum = 0;

            while ((line = reader.readLine()) != null) {
                Row row = sheet.createRow(rowNum++);
                String[] values = line.split(",");

                for (int i = 0; i < values.length; i++) {
                    row.createCell(i).setCellValue(values[i]);
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertDocxToTxt(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes))) {
            StringBuilder text = new StringBuilder();

            for (XWPFParagraph para : doc.getParagraphs()) {
                text.append(para.getText()).append("\n");
            }

            return text.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertDocxToPdf(byte[] docxBytes) throws IOException {
        try (PDDocument pdf = new PDDocument()) {
            String text = new String(convertDocxToTxt(docxBytes), StandardCharsets.UTF_8);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pdf.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertPptxToDocx(byte[] pptxBytes) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(pptxBytes));
             XWPFDocument doc = new XWPFDocument()) {

            for (XSLFSlide slide : ppt.getSlides()) {
                XWPFParagraph para = doc.createParagraph();
                XWPFRun run = para.createRun();
                run.setText("Slide " + (slide.getSlideNumber() + 1));
                run.addBreak();

                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        run.setText(((XSLFTextShape) shape).getText());
                        run.addBreak();
                    }
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertPdfToTxt(byte[] pdfBytes) throws IOException {
        try (PDDocument doc =  PDDocument.load(pdfBytes);) {
            PDFTextStripper stripper = new PDFTextStripper();
            return stripper.getText(doc).getBytes(StandardCharsets.UTF_8);
        }
    }
    private byte[] convertTxtToDocx(byte[] txtBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument();
             BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(txtBytes)))) {

            String line;
            while ((line = reader.readLine()) != null) {
                XWPFParagraph para = doc.createParagraph();
                XWPFRun run = para.createRun();
                run.setText(line);
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertXlsxToCsv(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes));
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            StringBuilder csvData = new StringBuilder();

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (Row row : sheet) {
                boolean firstCell = true;

                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (!firstCell) {
                        csvData.append(",");
                    }
                    firstCell = false;

                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = getProperCellValue(cell, formatter, evaluator);
                    csvData.append(escapeCsvValue(cellValue));
                }
                csvData.append("\n");
            }

            out.write(csvData.toString().getBytes(StandardCharsets.UTF_8));
            return out.toByteArray();
        }
    }

    private String getProperCellValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        switch (cell.getCellType()) {
            case FORMULA:
                return formatter.formatCellValue(cell, evaluator);
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return formatter.formatCellValue(cell);
            default:
                return formatter.formatCellValue(cell);
        }
    }

    private String escapeCsvValue(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }

        boolean needsQuotes = value.contains(",") || value.contains("\n") || value.contains("\"");
        String escapedValue = value.replace("\"", "\"\"");

        return needsQuotes ? "\"" + escapedValue + "\"" : escapedValue;
    }

    private byte[] convertDocxToPptx(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes));
             XMLSlideShow ppt = new XMLSlideShow()) {

            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
            XSLFSlide titleSlide = ppt.createSlide(titleLayout);
            titleSlide.getPlaceholder(0).setText("Documento Convertido");

            for (XWPFParagraph para : doc.getParagraphs()) {
                if (!para.getText().isEmpty()) {
                    XSLFSlideLayout contentLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
                    XSLFSlide slide = ppt.createSlide(contentLayout);
                    slide.getPlaceholder(0).setText("Seção");
                    slide.getPlaceholder(1).setText(para.getText());
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertXlsxToPdf(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes));
             PDDocument pdfDoc = new PDDocument()) {

            PDPage page = new PDPage();
            pdfDoc.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(pdfDoc, page);
            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA, 12);
            contentStream.newLineAtOffset(50, 750);

            Sheet sheet = workbook.getSheetAt(0);
            int rowNum = 0;

            String tabReplacement = "    ";

            for (Row row : sheet) {
                StringBuilder rowData = new StringBuilder();
                for (Cell cell : row) {
                    rowData.append(cell.toString()).append(tabReplacement);
                }

                contentStream.showText(rowData.toString().trim());
                contentStream.newLineAtOffset(0, -15);

                rowNum++;
                if (rowNum > 30) {
                    contentStream.endText();
                    contentStream.close();

                    page = new PDPage();
                    pdfDoc.addPage(page);
                    contentStream = new PDPageContentStream(pdfDoc, page);
                    contentStream.beginText();
                    contentStream.setFont(PDType1Font.HELVETICA, 12);
                    contentStream.newLineAtOffset(50, 750);
                    rowNum = 0;
                }
            }

            contentStream.endText();
            contentStream.close();

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pdfDoc.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertTxtToHtml(byte[] txtBytes) throws IOException {
        String text = new String(txtBytes, StandardCharsets.UTF_8);

        text = text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#39;");

        text = text.replace("\n", "<br>\n");

        String html = "<!DOCTYPE html>\n" +
                "<html>\n" +
                "<head>\n" +
                "    <meta charset=\"UTF-8\">\n" +
                "    <title>Converted Text</title>\n" +
                "    <style>\n" +
                "        body {\n" +
                "            font-family: Arial, sans-serif;\n" +
                "            white-space: pre-wrap;\n" +
                "            margin: 20px;\n" +
                "            line-height: 1.5;\n" +
                "        }\n" +
                "    </style>\n" +
                "</head>\n" +
                "<body>\n" +
                text + "\n" +
                "</body>\n" +
                "</html>";

        return html.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertXlsxToXml(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes))) {
            Sheet sheet = workbook.getSheetAt(0);
            StringBuilder xml = new StringBuilder();
            xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.append("<workbook>\n");
            xml.append("  <sheet name=\"").append(sheet.getSheetName()).append("\">\n");

            for (Row row : sheet) {
                xml.append("    <row>\n");
                for (Cell cell : row) {
                    xml.append("      <cell>")
                            .append(cell.toString())
                            .append("</cell>\n");
                }
                xml.append("    </row>\n");
            }

            xml.append("  </sheet>\n");
            xml.append("</workbook>");

            return xml.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertXlsxToHtml(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes))) {
            Sheet sheet = workbook.getSheetAt(0);
            StringBuilder html = new StringBuilder();
            html.append("<!DOCTYPE html>\n");
            html.append("<html>\n");
            html.append("<head>\n");
            html.append("  <meta charset=\"UTF-8\">\n");
            html.append("  <title>").append(sheet.getSheetName()).append("</title>\n");
            html.append("  <style>table { border-collapse: collapse; } td, th { border: 1px solid #ddd; padding: 8px; }</style>\n");
            html.append("</head>\n");
            html.append("<body>\n");
            html.append("<table>\n");

            for (Row row : sheet) {
                html.append("  <tr>\n");
                for (Cell cell : row) {
                    html.append("    <td>").append(cell.toString()).append("</td>\n");
                }
                html.append("  </tr>\n");
            }

            html.append("</table>\n");
            html.append("</body>\n");
            html.append("</html>");

            return html.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertCsvToJson(byte[] csvBytes) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(csvBytes)));
        String line;
        List<Map<String, String>> data = new ArrayList<>();
        String[] headers = reader.readLine().split(",");

        while ((line = reader.readLine()) != null) {
            String[] values = line.split(",");
            Map<String, String> row = new HashMap<>();
            for (int i = 0; i < headers.length && i < values.length; i++) {
                row.put(headers[i], values[i]);
            }
            data.add(row);
        }

        return new ObjectMapper().writeValueAsBytes(data);
    }

    private byte[] convertCsvToXml(byte[] csvBytes) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(csvBytes)));
        String line;
        StringBuilder xml = new StringBuilder();
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        xml.append("<data>\n");

        String[] headers = reader.readLine().split(",");

        while ((line = reader.readLine()) != null) {
            xml.append("  <row>\n");
            String[] values = line.split(",");
            for (int i = 0; i < headers.length && i < values.length; i++) {
                xml.append("    <").append(headers[i]).append(">")
                        .append(values[i])
                        .append("</").append(headers[i]).append(">\n");
            }
            xml.append("  </row>\n");
        }

        xml.append("</data>");
        return xml.toString().getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertCsvToHtml(byte[] csvBytes) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new ByteArrayInputStream(csvBytes)));
        String line;
        StringBuilder html = new StringBuilder();
        html.append("<!DOCTYPE html>\n");
        html.append("<html>\n");
        html.append("<head>\n");
        html.append("  <meta charset=\"UTF-8\">\n");
        html.append("  <title>CSV Converted</title>\n");
        html.append("  <style>table { border-collapse: collapse; } td, th { border: 1px solid #ddd; padding: 8px; }</style>\n");
        html.append("</head>\n");
        html.append("<body>\n");
        html.append("<table>\n");

        // Headers
        String[] headers = reader.readLine().split(",");
        html.append("  <tr>\n");
        for (String header : headers) {
            html.append("    <th>").append(header).append("</th>\n");
        }
        html.append("  </tr>\n");

        // Data rows
        while ((line = reader.readLine()) != null) {
            html.append("  <tr>\n");
            String[] values = line.split(",");
            for (String value : values) {
                html.append("    <td>").append(value).append("</td>\n");
            }
            html.append("  </tr>\n");
        }

        html.append("</table>\n");
        html.append("</body>\n");
        html.append("</html>");

        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertCsvToPdf(byte[] csvBytes) throws IOException {
        try (PDDocument doc = new PDDocument()) {
            PDPage page = new PDPage();
            doc.addPage(page);

            String text = new String(csvBytes, StandardCharsets.UTF_8);

            try (PDPageContentStream contentStream = new PDPageContentStream(doc, page)) {
                contentStream.beginText();
                contentStream.setFont(PDType1Font.COURIER, 12); // ✅ Usar PDType1Font
                contentStream.newLineAtOffset(50, 700);

                String[] lines = text.split("\n");
                for (String line : lines) {
                    contentStream.showText(line);
                    contentStream.newLineAtOffset(0, -15);
                }

                contentStream.endText();
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertDocxToHtml(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes))) {
            StringBuilder html = new StringBuilder();
            html.append("<!DOCTYPE html>\n");
            html.append("<html>\n");
            html.append("<head>\n");
            html.append("  <meta charset=\"UTF-8\">\n");
            html.append("  <title>Converted Document</title>\n");
            html.append("</head>\n");
            html.append("<body>\n");

            for (XWPFParagraph para : doc.getParagraphs()) {
                html.append("<p>").append(para.getText()).append("</p>\n");
            }

            html.append("</body>\n");
            html.append("</html>");

            return html.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertDocxToJson(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes))) {
            List<Map<String, String>> paragraphs = new ArrayList<>();

            for (XWPFParagraph para : doc.getParagraphs()) {
                Map<String, String> paragraph = new HashMap<>();
                paragraph.put("text", para.getText());
                paragraph.put("style", para.getStyle());
                paragraphs.add(paragraph);
            }

            return new ObjectMapper().writeValueAsBytes(paragraphs);
        }
    }

    private byte[] convertDocxToXml(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes))) {
            StringBuilder xml = new StringBuilder();
            xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.append("<document>\n");

            for (XWPFParagraph para : doc.getParagraphs()) {
                xml.append("  <paragraph>\n");
                xml.append("    <text>").append(para.getText()).append("</text>\n");
                xml.append("  </paragraph>\n");
            }

            xml.append("</document>");
            return xml.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertPptxToPdf(byte[] pptxBytes) throws IOException {
        try (PDDocument pdf = new PDDocument()) {
            String text = new String(convertPptxToTxt(pptxBytes), StandardCharsets.UTF_8);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pdf.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertPptxToTxt(byte[] pptxBytes) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(pptxBytes))) {
            StringBuilder text = new StringBuilder();

            for (XSLFSlide slide : ppt.getSlides()) {
                text.append("Slide ").append(slide.getSlideNumber() + 1).append("\n");

                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        text.append(((XSLFTextShape) shape).getText()).append("\n");
                    }
                }
                text.append("\n");
            }

            return text.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertPptxToHtml(byte[] pptxBytes) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(pptxBytes))) {
            StringBuilder html = new StringBuilder();
            html.append("<!DOCTYPE html>\n");
            html.append("<html>\n");
            html.append("<head>\n");
            html.append("  <meta charset=\"UTF-8\">\n");
            html.append("  <title>Presentation</title>\n");
            html.append("</head>\n");
            html.append("<body>\n");

            for (XSLFSlide slide : ppt.getSlides()) {
                html.append("<div class=\"slide\">\n");
                html.append("  <h2>Slide ").append(slide.getSlideNumber() + 1).append("</h2>\n");

                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        html.append("  <p>").append(((XSLFTextShape) shape).getText()).append("</p>\n");
                    }
                }

                html.append("</div>\n");
            }

            html.append("</body>\n");
            html.append("</html>");

            return html.toString().getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertPptxToJson(byte[] pptxBytes) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow(new ByteArrayInputStream(pptxBytes))) {
            List<Map<String, Object>> slides = new ArrayList<>();

            for (XSLFSlide slide : ppt.getSlides()) {
                Map<String, Object> slideData = new HashMap<>();
                slideData.put("slideNumber", slide.getSlideNumber() + 1);

                List<String> contents = new ArrayList<>();
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        contents.add(((XSLFTextShape) shape).getText());
                    }
                }

                slideData.put("contents", contents);
                slides.add(slideData);
            }

            return new ObjectMapper().writeValueAsBytes(slides);
        }
    }

    private byte[] convertPdfToDocx(byte[] pdfBytes) throws IOException {
        try (PDDocument pdf = PDDocument.load(pdfBytes);  // ✅ Usar PDDocument.load() diretamente
             XWPFDocument doc = new XWPFDocument()) {

            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(pdf);

            for (String line : text.split("\n")) {
                XWPFParagraph para = doc.createParagraph();
                XWPFRun run = para.createRun();
                run.setText(line);
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertPdfToPptx(byte[] pdfBytes) throws IOException {
        try (PDDocument pdf =PDDocument.load(pdfBytes);
             XMLSlideShow ppt = new XMLSlideShow()) {


            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(pdf);

            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
            XSLFSlide titleSlide = ppt.createSlide(titleLayout);
            titleSlide.getPlaceholder(0).setText("Documento Convertido");

            for (String paragraph : text.split("\n\n")) {
                if (!paragraph.trim().isEmpty()) {
                    XSLFSlideLayout contentLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
                    XSLFSlide slide = ppt.createSlide(contentLayout);
                    slide.getPlaceholder(1).setText(paragraph);
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertPdfToHtml(byte[] pdfBytes) throws IOException {
        try (PDDocument pdf = PDDocument.load(pdfBytes);) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(pdf);

            String html = "<!DOCTYPE html>\n" +
                    "<html>\n" +
                    "<head>\n" +
                    "  <meta charset=\"UTF-8\">\n" +
                    "  <title>PDF Converted</title>\n" +
                    "</head>\n" +
                    "<body>\n" +
                    "<pre>" + text + "</pre>\n" +
                    "</body>\n" +
                    "</html>";

            return html.getBytes(StandardCharsets.UTF_8);
        }
    }

    private byte[] convertPdfToJson(byte[] pdfBytes) throws IOException {
        try (PDDocument pdf = PDDocument.load(pdfBytes);) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(pdf);

            Map<String, String> pdfData = new HashMap<>();
            pdfData.put("content", text);
            pdfData.put("pageCount", String.valueOf(pdf.getNumberOfPages()));

            return new ObjectMapper().writeValueAsBytes(pdfData);
        }
    }

    private byte[] convertTxtToJson(byte[] txtBytes) throws IOException {
        String text = new String(txtBytes, StandardCharsets.UTF_8);
        Map<String, String> data = new HashMap<>();
        data.put("content", text);
        return new ObjectMapper().writeValueAsBytes(data);
    }

    private byte[] convertTxtToXml(byte[] txtBytes) throws IOException {
        String text = new String(txtBytes, StandardCharsets.UTF_8);
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<text>\n" +
                "  <content>" + text + "</content>\n" +
                "</text>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertHtmlToTxt(byte[] htmlBytes) throws IOException {
        String html = new String(htmlBytes, StandardCharsets.UTF_8);
        String text = html.replaceAll("<[^>]*>", "");
        return text.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertHtmlToPdf(byte[] htmlBytes) throws IOException {
        try (PDDocument pdf = new PDDocument()) {
            String text = new String(convertHtmlToTxt(htmlBytes), StandardCharsets.UTF_8);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            pdf.save(out);
            return out.toByteArray();
        }
    }

    private byte[] convertHtmlToDocx(byte[] htmlBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            String text = new String(convertHtmlToTxt(htmlBytes), StandardCharsets.UTF_8);

            for (String line : text.split("\n")) {
                XWPFParagraph para = doc.createParagraph();
                XWPFRun run = para.createRun();
                run.setText(line);
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            doc.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertHtmlToJson(byte[] htmlBytes) throws IOException {
        String html = new String(htmlBytes, StandardCharsets.UTF_8);
        Map<String, String> data = new HashMap<>();
        data.put("html", html);
        data.put("text", new String(convertHtmlToTxt(htmlBytes), StandardCharsets.UTF_8));
        return new ObjectMapper().writeValueAsBytes(data);
    }

    private byte[] convertHtmlToXml(byte[] htmlBytes) throws IOException {
        String html = new String(htmlBytes, StandardCharsets.UTF_8);
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<html>\n" +
                "  <content><![CDATA[" + html + "]]></content>\n" +
                "</html>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertJsonToXlsx(byte[] jsonBytes) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        List<Map<String, String>> data = mapper.readValue(jsonBytes,
                mapper.getTypeFactory().constructCollectionType(List.class, Map.class));

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            // Headers
            if (!data.isEmpty()) {
                Row headerRow = sheet.createRow(0);
                int colNum = 0;
                for (String key : data.get(0).keySet()) {
                    headerRow.createCell(colNum++).setCellValue(key);
                }

                // Data
                int rowNum = 1;
                for (Map<String, String> row : data) {
                    Row dataRow = sheet.createRow(rowNum++);
                    colNum = 0;
                    for (String value : row.values()) {
                        dataRow.createCell(colNum++).setCellValue(value);
                    }
                }
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertJsonToCsv(byte[] jsonBytes) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        List<Map<String, String>> data = mapper.readValue(
                jsonBytes,
                mapper.getTypeFactory().constructCollectionType(List.class, Map.class)
        );

        if (data.isEmpty()) {
            return new byte[0];
        }

        StringBuilder csv = new StringBuilder();

        // Headers
        csv.append(escapeCsvRow(data.get(0).keySet())).append("\n");

        // Data
        for (Map<String, String> row : data) {
            csv.append(escapeCsvRow(row.values())).append("\n");
        }

        return csv.toString().getBytes(StandardCharsets.UTF_8);
    }

    private String escapeCsvRow(Collection<String> values) {
        return values.stream()
                .map(value -> {
                    if (value == null) {
                        return "";
                    }
                    String escaped = value.replace("\"", "\"\"");
                    if (escaped.contains(",") || escaped.contains("\n") || escaped.contains("\"")) {
                        return "\"" + escaped + "\"";
                    }
                    return escaped;
                })
                .collect(Collectors.joining(","));
    }

    private byte[] convertJsonToXml(byte[] jsonBytes) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> data = mapper.readValue(jsonBytes, Map.class);

        StringBuilder xml = new StringBuilder();
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        xml.append("<root>\n");

        for (Map.Entry<String, Object> entry : data.entrySet()) {
            xml.append("  <").append(entry.getKey()).append(">")
                    .append(entry.getValue().toString())
                    .append("</").append(entry.getKey()).append(">\n");
        }

        xml.append("</root>");
        return xml.toString().getBytes(StandardCharsets.UTF_8);
    }

    private byte[] convertJsonToTxt(byte[] jsonBytes) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        Object json = mapper.readValue(jsonBytes, Object.class);
        return mapper.writerWithDefaultPrettyPrinter()
                .writeValueAsString(json)
                .getBytes(StandardCharsets.UTF_8);
    }




    private byte[] convertXmlToTxt(byte[] xmlBytes) throws IOException {
        String json = new String(convertXmlToTxt(xmlBytes), StandardCharsets.UTF_8);
        return convertJsonToTxt(json.getBytes(StandardCharsets.UTF_8));
    }

    private byte[] convertDocxToRtf(byte[] docxBytes) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(docxBytes));
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            String rtfHeader = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}\n";
            StringBuilder rtfContent = new StringBuilder(rtfHeader);

            for (XWPFParagraph para : doc.getParagraphs()) {
                String text = para.getText();
                text = text.replace("\\", "\\\\")
                        .replace("{", "\\{")
                        .replace("}", "\\}")
                        .replace("\n", "\\par\n");
                rtfContent.append(text).append("\\par\n");
            }

            rtfContent.append("}");
            out.write(rtfContent.toString().getBytes(StandardCharsets.UTF_8));
            return out.toByteArray();
        }
    }



    private byte[] convertXlsxToTsv(byte[] xlsxBytes) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(xlsxBytes));
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            StringBuilder tsvData = new StringBuilder();

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (Row row : sheet) {
                boolean firstCell = true;

                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (!firstCell) {
                        tsvData.append("\t");
                    }
                    firstCell = false;

                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = getProperCellValue(cell, formatter, evaluator);
                    tsvData.append(escapeTsvValue(cellValue));
                }
                tsvData.append("\n");
            }

            out.write(tsvData.toString().getBytes(StandardCharsets.UTF_8));
            return out.toByteArray();
        }
    }

    private String escapeTsvValue(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }

        String escaped = value.replace("\n", " ").replace("\r", " ").replace("\t", " ");

        if (escaped.contains("\"")) {
            escaped = escaped.replace("\"", "\"\"");
        }

        if (escaped.contains("\"") || escaped.contains("\t") || escaped.contains("\n") || escaped.contains(",")) {
            escaped = "\"" + escaped + "\"";
        }

        return escaped;
    }

}