package bluezealot.translateppt;

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
        
        PptTranslator pptTranslator = new PptTranslator(translator);
        pptTranslator.translatePpt("/Volumes/Seagate/work/robot/ts_test/001268136.pptx", "/Volumes/Seagate/work/robot/ts_test/001268136_translated.pptx", sourceLang, targetLang);
    }
}
