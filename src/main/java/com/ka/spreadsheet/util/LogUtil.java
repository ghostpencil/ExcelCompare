package com.ka.spreadsheet.util;

import org.apache.log4j.FileAppender;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;

public class LogUtil {
    public static final String ROOT_DIRECTORY = "root_directory";

    public static void initLogging(String logFileName){
        FileAppender fa = new FileAppender();
        fa.setName("FileLogger");
        String rd = System.getProperty(ROOT_DIRECTORY);
        rd = rd == null ? "." : rd;
        fa.setFile(rd + "/logs/" +logFileName);
        fa.setLayout(new PatternLayout("%d %-5p [%c{1}] %m%n"));
        fa.setThreshold(Level.DEBUG);
        fa.setAppend(false);
        fa.activateOptions();
        Logger.getRootLogger().addAppender(fa);
    }
}