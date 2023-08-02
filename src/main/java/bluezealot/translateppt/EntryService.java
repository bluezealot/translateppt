package bluezealot.translateppt;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
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
        // TODO Auto-generated method stub
        log.info("run");
        try(var fis = new FileInputStream("/Users/qinxizhou/work/iot/RmindTarget.pptx")){
            XMLSlideShow ppt = new XMLSlideShow(fis);
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape sh : slide.getShapes()) {
                    if (sh instanceof XSLFTextShape) {
                        XSLFTextShape shape = (XSLFTextShape) sh;
                        log.info(shape.getText());
                        shape.setText("test");
                    }
                }
                break;
            }
            try(FileOutputStream fos = new FileOutputStream("/Users/qinxizhou/work/iot/edited_RmindTarget.pptx", false))
            {
                ProcessBuilder ps=new ProcessBuilder("node","/Users/qinxizhou/work/git/github/translator/translatorcmd.js","123");
                Process pr = ps.start();
                String line;
                BufferedReader in = new BufferedReader(new InputStreamReader(pr.getInputStream()));
                while ((line = in.readLine()) != null) {
                    System.out.println(line);
                }
                pr.waitFor();
                in.close();
                System.out.println("ok!");
                ppt.write(fos);
            }
        }
        
    }
    
}
