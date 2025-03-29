package com.file_converter_back.controller;

import com.file_converter_back.DTO.ConversionOption;
import com.file_converter_back.services.FileConversionService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api")
@CrossOrigin(origins = "http://localhost:5173")
public class FileController {

    private final FileConversionService conversionService;

    public FileController(FileConversionService conversionService) {
        this.conversionService = conversionService;
    }

    @PostMapping(value = "/get-conversion-options", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> getConversionOptions(@RequestParam("file") MultipartFile file) {
        try {
            String fileExtension = getFileExtension(file.getOriginalFilename());
            List<ConversionOption> options = conversionService.getAvailableConversions(fileExtension);
            return ResponseEntity.ok(options);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(createErrorResponse("Erro ao obter opções", e));
        }
    }

    @PostMapping(value = "/convert", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> convertFile(
            @RequestParam("file") MultipartFile file,
            @RequestParam("targetFormat") String targetFormat) {

        try {
            byte[] convertedBytes = conversionService.convert(
                    file.getBytes(),
                    getFileExtension(file.getOriginalFilename()),
                    targetFormat
            );

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted." + targetFormat)
                    .contentType(getMediaType(targetFormat))
                    .body(convertedBytes);
        } catch (UnsupportedOperationException e) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(createErrorResponse("Conversão não suportada", e));
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .contentType(MediaType.APPLICATION_JSON)
                    .body(createErrorResponse("Erro na conversão", e));
        }
    }

    private String getFileExtension(String filename) {
        if (filename == null || !filename.contains(".")) {
            throw new IllegalArgumentException("Nome de arquivo inválido");
        }
        return filename.substring(filename.lastIndexOf(".") + 1).toLowerCase();
    }

    private MediaType getMediaType(String format) {
        return switch (format.toLowerCase()) {
            case "docx" -> MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            case "xlsx" -> MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            case "pptx" -> MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            case "csv" -> MediaType.TEXT_PLAIN;
            case "pdf" -> MediaType.APPLICATION_PDF;
            case "txt" -> MediaType.TEXT_PLAIN;
            case "json" -> MediaType.APPLICATION_JSON;
            case "xml" -> MediaType.APPLICATION_XML;
            default -> MediaType.APPLICATION_OCTET_STREAM;
        };
    }

    private Map<String, String> createErrorResponse(String error, Exception e) {
        Map<String, String> response = new HashMap<>();
        response.put("error", error);
        response.put("message", e.getMessage());
        return response;
    }


}