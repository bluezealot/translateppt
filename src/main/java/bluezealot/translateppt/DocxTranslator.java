package bluezealot.translateppt;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class DocxTranslator {
    private final Translator translator;

    public DocxTranslator(Translator translator) {
        this.translator = translator;
    }

    public void translateDocx(String sourceFilePath, String targetFilePath, String sourceLang, String targetLang) throws IOException {
        log.info("Starting DOCX translation process");
        
        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             FileOutputStream fos = new FileOutputStream(targetFilePath)) {
            
            XWPFDocument doc = new XWPFDocument(fis);
            translateDocument(doc, sourceLang, targetLang);
            doc.write(fos);
            log.info("Translated DOCX file saved successfully to: {}", targetFilePath);
        }
    }

    private void translateDocument(XWPFDocument doc, String sourceLang, String targetLang) {
        // Process paragraphs
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            String originalText = paragraph.getText();
            if (originalText != null && !originalText.trim().isEmpty()) {
                try {
                    log.info("Translating paragraph: {}", originalText);
                    String translatedText = translator.translate(originalText, sourceLang, targetLang);
                    log.info("Translated to: {}", translatedText);
                    
                    // Get all runs in the paragraph
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs != null && !runs.isEmpty()) {
                        // Clear existing runs and create a single new run with translated text
                        for (XWPFRun run : runs) {
                            clearText(run);
                        }
                        XWPFRun newRun = runs.get(0);
                        newRun.setText(translatedText, 0);
                    }
                } catch (Exception e) {
                    log.error("Error translating paragraph: " + originalText, e);
                    // Keep original text if translation fails
                    continue;
                }
            }
        }

        // Process tables
        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        String originalText = paragraph.getText();
                        if (originalText != null && !originalText.trim().isEmpty()) {
                            try {
                                log.info("Translating table cell: {}", originalText);
                                String translatedText = translator.translate(originalText, sourceLang, targetLang);
                                log.info("Translated to: {}", translatedText);
                                
                                // Get all runs in the paragraph
                                List<XWPFRun> runs = paragraph.getRuns();
                                if (runs != null && !runs.isEmpty()) {
                                    // Clear existing runs and create a single new run with translated text
                                    for (XWPFRun run : runs) {
                                        clearText(run);
                                    }
                                    XWPFRun newRun = runs.get(0);
                                    newRun.setText(translatedText, 0);
                                }
                            } catch (Exception e) {
                                log.error("Error translating table cell: " + originalText, e);
                                continue;
                            }
                        }
                    }
                }
            }
        }
    }

    public void clearAndSetText(XWPFRun run, String newText) {
        CTR ctr = run.getCTR();
    
        // Remove all <w:t> elements
        while (ctr.sizeOfTArray() > 0) {
            ctr.removeT(0);
        }
    
        // Add a new one
        run.setText(newText, 0);
    }

    public void clearText(XWPFRun run) {
        CTR ctr = run.getCTR();
    
        // Remove all <w:t> elements
        while (ctr.sizeOfTArray() > 0) {
            ctr.removeT(0);
        }
    }
}
