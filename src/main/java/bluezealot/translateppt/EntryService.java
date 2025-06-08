package bluezealot.translateppt;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.stereotype.Service;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class EntryService implements CommandLineRunner, ExitCodeGenerator{
    @Override
    public int getExitCode() {
        // TODO Auto-generated method stub
        return 0;
    }

    @Override
    public void run(String... args) throws Exception {
        log.info("Starting PPT translation process");
        
        // Initialize translator
        Translator translator = new Translator();
        String sourceLang = "ja-jp"; // Source language
        String targetLang = "zh-cn"; // Target language
        
        try(var fis = new FileInputStream("/Volumes/Seagate/ts_test/販売店おもてなしロボットAI搭載型の機能 20240119.pptx")){
            XMLSlideShow ppt = new XMLSlideShow(fis);
            
            // Collect all text shapes first, including those in group shapes
            List<XSLFTextShape> textShapes = new ArrayList<>();
            
            // Process shapes recursively
            for (XSLFSlide slide : ppt.getSlides()) {
                List<XSLFShape> shapesToProcess = new ArrayList<>();
                shapesToProcess.addAll(slide.getShapes());
                
                while (!shapesToProcess.isEmpty()) {
                    XSLFShape shape = shapesToProcess.remove(0);
                    if (shape instanceof XSLFTextShape) {
                        textShapes.add((XSLFTextShape) shape);
                    } else if (shape instanceof XSLFGroupShape) {
                        shapesToProcess.addAll(0, ((XSLFGroupShape) shape).getShapes());
                    } else if (shape instanceof XSLFTable) {
                        // Process table cells
                        XSLFTable table = (XSLFTable) shape;
                        for (int row = 0; row < table.getNumberOfRows(); row++) {
                            for (int col = 0; col < table.getNumberOfColumns(); col++) {
                                XSLFTextShape cellShape = table.getCell(row, col);
                                if (cellShape != null) {
                                    textShapes.add(cellShape);
                                }
                            }
                        }
                    }
                }
            }
            
            // Process collected text shapes
            for (XSLFTextShape shape : textShapes) {
                // Process each paragraph separately
                for (XSLFTextParagraph paragraph : shape.getTextParagraphs()) {
                    StringBuilder paragraphText = new StringBuilder();
                    List<XSLFTextRun> runsInParagraph = new ArrayList<>();
                    
                    // Collect text and runs for this paragraph
                    for (XSLFTextRun run : paragraph.getTextRuns()) {
                        String runText = run.getRawText();
                        if (runText != null && !runText.trim().isEmpty()) {
                            paragraphText.append(runText);
                            runsInParagraph.add(run);
                        }
                    }
                    
                    String originalParagraphText = paragraphText.toString();
                    if (!originalParagraphText.trim().isEmpty()) {
                        log.info("Translating paragraph: {}", originalParagraphText);
                        try {
                            // Translate the paragraph
                            String translatedParagraph = translator.translate(originalParagraphText, sourceLang, targetLang);
                            
                            // Split translated text by spaces to match runs
                            String[] translatedParts = translatedParagraph.split("\\s+");
                            int partIndex = 0;
                            
                            // Distribute translated text back to runs in this paragraph
                            for (XSLFTextRun run : runsInParagraph) {
                                if (partIndex < translatedParts.length) {
                                    String translatedPart = translatedParts[partIndex++];
                                    run.setText(translatedPart);
                                } else {
                                    // If we run out of translated parts, just clear the remaining runs
                                    run.setText("");
                                }
                            }
                            
                            log.info("Translated to: {}", translatedParagraph);
                        } catch (Exception e) {
                            log.error("Error translating paragraph:", e);
                            // Restore original text for runs in this paragraph if translation fails
                            for (XSLFTextRun run : runsInParagraph) {
                                run.setText(run.getRawText());
                            }
                        }
                    }
                }
            }
            
            // Save the translated PPT
            try(FileOutputStream fos = new FileOutputStream("/Volumes/Seagate/ts_test/販売店おもてなしロボットAI搭載型の機能 20240119_translated.pptx", false)) {
                ppt.write(fos);
                log.info("Successfully saved translated PPT");
            }
        } catch (Exception e) {
            log.error("Error during PPT translation:", e);
            throw e;
        }
    }
}
