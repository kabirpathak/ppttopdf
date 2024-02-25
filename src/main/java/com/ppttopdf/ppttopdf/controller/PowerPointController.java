package com.ppttopdf.ppttopdf.controller;

import com.ppttopdf.ppttopdf.service.PowerPointService;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

@Controller
public class PowerPointController {

    private final PowerPointService powerPointService;

    public PowerPointController(PowerPointService powerPointService) {
        this.powerPointService = powerPointService;
    }

    @GetMapping("/upload")
    public String showUploadForm() {
        return "upload";
    }

    @PostMapping("/api/powerpoint/read")
    public ResponseEntity<List<String>> readPowerPointFiles(@RequestParam("files") List<MultipartFile> powerpointFiles) {
        // Convert MultipartFiles to regular Files
        List<File> files = convertMultipartFiles(powerpointFiles);

        // Use PowerPointService to read content from PowerPoint files
        return ResponseEntity.ok(powerPointService.readPowerPointFiles(files));
    }

    @PostMapping("/api/powerpoint/mergeAndConvert")
    public ResponseEntity<File> mergeAndConvertPowerPointFiles(@RequestParam("files") List<MultipartFile> powerpointFiles) {
        // Convert MultipartFiles to regular Files
        List<File> files = convertMultipartFiles(powerpointFiles);

        // Use PowerPointService to merge PowerPoint files
        File mergedPowerPointFile = powerPointService.mergePowerPointFiles(files);

        // Use PowerPointService to convert the merged PowerPoint file to PDF
        File pdfFile = powerPointService.convertToPdf(mergedPowerPointFile);

        // Return the resulting PDF file
        return ResponseEntity.ok(pdfFile);
    }

    private List<File> convertMultipartFiles(List<MultipartFile> multipartFiles) {
        List<File> convertedFiles = new ArrayList<>();

        for (MultipartFile multipartFile : multipartFiles) {
            File file = convertToFile(multipartFile);
            convertedFiles.add(file);
        }

        return convertedFiles;
    }

    private File convertToFile(MultipartFile multipartFile) {
        File file = new File(multipartFile.getOriginalFilename());

        try {
            multipartFile.transferTo(file);
        } catch (Exception e) {
            e.printStackTrace(); // Handle the exception appropriately
        }

        return file;
    }
}
