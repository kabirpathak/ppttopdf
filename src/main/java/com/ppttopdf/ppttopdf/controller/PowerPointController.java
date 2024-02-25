package com.ppttopdf.ppttopdf.controller;

import com.ppttopdf.ppttopdf.service.PowerPointService;
import java.util.ArrayList;
import java.util.List;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/powerpoint")
public class PowerPointController {

    private final PowerPointService powerPointService;

    public PowerPointController(PowerPointService powerPointService) {
        this.powerPointService = powerPointService;
    }

    @PostMapping("/read")
    public ResponseEntity<List<String>> readPowerPointFiles(@RequestParam("files") List<MultipartFile> powerpointFiles) {
        // Convert MultipartFiles to regular Files
        List<File> files = convertMultipartFiles(powerpointFiles);

        // Use PowerPointService to read content from PowerPoint files
        return ResponseEntity.ok(powerPointService.readPowerPointFiles(files));
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

        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(multipartFile.getBytes());
        } catch (IOException e) {
            e.printStackTrace(); // Handle the exception appropriately
        }

        return file;
    }

    @PostMapping("/mergeAndConvert")
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

}
