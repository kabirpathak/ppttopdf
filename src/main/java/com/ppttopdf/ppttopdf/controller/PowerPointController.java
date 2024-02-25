package com.ppttopdf.ppttopdf.controller;

import com.ppttopdf.ppttopdf.service.PowerPointService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
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
    public String mergeAndConvertPowerPointFiles(@RequestParam("files") List<MultipartFile> powerpointFiles, Model model) {
        // Convert MultipartFiles to regular Files
        List<File> files = convertMultipartFiles(powerpointFiles);

        // Use PowerPointService to merge PowerPoint files
        File mergedPowerPointFile = powerPointService.mergePowerPointFiles(files);

        // Use PowerPointService to convert the merged PowerPoint file to PDF
        File pdfFile = powerPointService.convertToPdf(mergedPowerPointFile);

        // Add the PDF file path to the model
        model.addAttribute("pdfFilePath", pdfFile.getAbsolutePath());

        // Redirect to the result page
        return "redirect:/result";
    }

    @GetMapping("/api/powerpoint/downloadPdf")
    public ResponseEntity<Resource> downloadPdf(@RequestParam("pdfFilePath") String pdfFilePath) {
        try {
            Path path = Paths.get(pdfFilePath);
            ByteArrayResource resource = new ByteArrayResource(Files.readAllBytes(path));

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + path.getFileName().toString())
                    .contentType(MediaType.APPLICATION_PDF)
                    .contentLength(resource.contentLength())
                    .body(resource);
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
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
