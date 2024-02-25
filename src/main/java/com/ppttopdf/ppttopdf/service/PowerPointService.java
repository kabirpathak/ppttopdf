package com.ppttopdf.ppttopdf.service;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
}
