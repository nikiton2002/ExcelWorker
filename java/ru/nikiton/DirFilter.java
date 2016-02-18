package ru.nikiton;

import java.io.File;
import java.io.FilenameFilter;
import java.util.regex.Pattern;

public class DirFilter implements FilenameFilter{
    private Pattern pattern;

    public DirFilter (String regex){
        pattern = Pattern.compile(regex);
    }
    public boolean accept(File dir, String fileName){
        return pattern.matcher(fileName).matches();
    }
}
