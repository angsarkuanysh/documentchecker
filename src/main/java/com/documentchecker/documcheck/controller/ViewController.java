package com.documentchecker.documcheck.controller;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.documentchecker.documcheck.service.DocxToHtmlConverter;
import com.documentchecker.documcheck.service.StyleApplier;

import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

@Controller
public class ViewController {
    @GetMapping("/") 
    public String index() {
        return "index";
    }

    @GetMapping("/upload")
    public String getUpload(Model model) {
        model.addAttribute("fontSizeValue", 14);
        model.addAttribute("indentValue", 1.25);
        model.addAttribute("lineSpacingValue", 1.5);
        return "upload";
    }
    
    @GetMapping("/test")
    public String test() {
        return "test";
    }
    
    @GetMapping("/login")
    public String login() {
        return "login"; 
    }

    @GetMapping("/download-styled")
    public void downloadStyledFile(@RequestParam("fontSize") int fontSize,HttpSession session, HttpServletResponse response) {
        try {
            byte[] fileBytes = (byte[]) session.getAttribute("lastUploadedFileBytes");
            String filename = (String) session.getAttribute("lastUploadedFileName");

            if (fileBytes == null) {
                response.sendError(HttpServletResponse.SC_NOT_FOUND, "Файл для скачивания не найден. Пожалуйста, сначала проверьте документ.");
                return;
            }

            XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(fileBytes));

            StyleApplier.applyGostStyles(document, fontSize);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            document.write(baos);
            document.close();

            String styledFilename = "Styled_" + (filename != null ? filename.replace(" ", "_") : "document.docx");

            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + styledFilename + "\"");
            response.setContentLength(baos.size());

            FileCopyUtils.copy(new ByteArrayInputStream(baos.toByteArray()), response.getOutputStream());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @PostMapping("/upload")
    public String handleUpload(@RequestParam("fontSize") int fontSize,
    @RequestParam("indent") double indent, @RequestParam("lineSpacing") 
    double lineSpacing,
    @RequestParam("file") MultipartFile file, 
    Model model,
    HttpSession session) {
        try {
            DocxToHtmlConverter htmlConverter = new DocxToHtmlConverter();
            String htmlContent = htmlConverter.convertDocxToHtmlWithErrors(file, indent, lineSpacing, fontSize);
            // String htmlContent = htmlConverter.convertDocxToHtml(file);
            model.addAttribute("fontSizeValue", fontSize);
            model.addAttribute("file", file);
            model.addAttribute("indentValue", indent);
            model.addAttribute("lineSpacingValue", lineSpacing);
            model.addAttribute("html", htmlContent);
            session.setAttribute("lastUploadedFileBytes", file.getBytes());
            session.setAttribute("lastUploadedFileName", file.getOriginalFilename());
            // model.addAttribute("history", htmlContent);

            

        } catch (Exception e) {
            System.out.println("ОШИБКА : "+ e.getMessage());
            model.addAttribute("errors", List.of("Ошибка при проверке: " + e.getMessage()));
        }
        return "upload";
    }

}
