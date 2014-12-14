package source;

import java.io.File;
import javax.swing.filechooser.FileFilter;

public class VCFFileFilter extends FileFilter {
    
    @Override
    public boolean accept(File f) {
        return f.isDirectory() || f.getName().toLowerCase().endsWith(".vcf");
    }

    @Override
    public String getDescription() {
        return "*.vcf";
    }
    
}
