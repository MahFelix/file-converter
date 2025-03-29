package com.file_converter_back.DTO;

public class ConversionOption {
    private String label;
    private String targetFormat;

    // Construtor, getters e setters
    public ConversionOption(String label, String targetFormat) {
        this.label = label;
        this.targetFormat = targetFormat;
    }

    // Getters
    public String getLabel() { return label; }
    public String getTargetFormat() { return targetFormat; }
}