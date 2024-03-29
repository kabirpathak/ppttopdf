package com.ppttopdf.ppttopdf.service;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import java.time.LocalDateTime;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;

@Service
public class PowerPointService {

    public List<String> readPowerPointFiles(List<File> powerpointFiles) {
        List<String> contentList = new ArrayList<>();

        for (File file : powerpointFiles) {
            try (FileInputStream fis = new FileInputStream(file)) {
                XMLSlideShow ppt = new XMLSlideShow(fis);

                // Iterate through slides and extract content
                for (XSLFSlide slide : ppt.getSlides()) {
                    // Extract content from the slide
                    String content = extractContentFromSlide(slide);
                    contentList.add(content);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return contentList;
    }

    private String extractContentFromSlide(XSLFSlide slide) {
        StringBuilder content = new StringBuilder();

        // Iterate through the shapes on the slide
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                // Check if the shape is a text-containing shape
                XSLFTextShape textShape = (XSLFTextShape) shape;

                // Retrieve and append the text from the shape
                String textContent = textShape.getText();
                content.append(textContent).append("\n");
            }
        }

        return content.toString();
    }

    public File mergePowerPointFiles(List<File> powerpointFiles) {
        try {
            XMLSlideShow mergedPpt = new XMLSlideShow();

            for (File file : powerpointFiles) {
                try (FileInputStream fis = new FileInputStream(file)) {
                    XMLSlideShow ppt = new XMLSlideShow(fis);

                    // Iterate through slides and add them to the merged presentation
                    for (XSLFSlide slide : ppt.getSlides()) {
                        mergedPpt.createSlide().importContent(slide);
                    }
                }
            }

            // Save the merged presentation to a file
            File mergedFile = new File("mergedPresentation.pptx");
            try (FileOutputStream fos = new FileOutputStream(mergedFile)) {
                mergedPpt.write(fos);
            }

            return mergedFile;
        } catch (IOException e) {
            e.printStackTrace();
            return null; // Handle the exception appropriately
        }
    }

    public File convertToPdf(File mergedPowerPointFile) {
        try {
            // Load the merged PowerPoint file
            XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(mergedPowerPointFile));

            // Create a PDF document
            PDDocument pdfDocument = new PDDocument();

            // Iterate through slides and add them to the PDF document
            for (XSLFSlide slide : ppt.getSlides()) {
                // Create a PDF page for each slide
                PDPage page = new PDPage();
                pdfDocument.addPage(page);

                // Create content stream to write on the PDF page
                PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);

                // Extract text content from the slide and write it to the PDF page
                String content = extractContentFromSlide(slide);

                // Split content into lines
                String[] lines = content.split("\n");

                // Set font and initial position
                contentStream.beginText();
                contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD), 12);
                contentStream.newLineAtOffset(10, 700); // Adjust the position as needed

                // Write each line with adjusted position
                for (String line : lines) {
                    contentStream.showText(line);
                    contentStream.newLineAtOffset(0, -12); // Adjust the vertical offset as needed
                }

                contentStream.endText();
                contentStream.close();
            }

            // Save the PDF document to a file
            File pdfFile = new File("output" + LocalDateTime.now() + ".pdf");
            pdfDocument.save(pdfFile);
            pdfDocument.close();

            return pdfFile;
        } catch (IOException e) {
            e.printStackTrace();
            return null; // Handle the exception appropriately
        }
    }
}
