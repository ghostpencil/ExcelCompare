package com.ka.spreadsheet.util;

import org.apache.log4j.*;

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

    public static void initCleanLogging(String logFileName){
        FileAppender fa = new FileAppender();
        fa.setName("FileDisplayLogger");
        String rd = System.getProperty(ROOT_DIRECTORY);
        rd = rd == null ? "." : rd;
        fa.setFile(rd + "/logs/" +logFileName);
        fa.setLayout(new PatternLayout("%m%n"));
        fa.setThreshold(Level.DEBUG);
        fa.setAppend(false);
        fa.activateOptions();
        Logger.getLogger("DisplayLogger").addAppender(fa);
        ConsoleAppender ca = new ConsoleAppender();
        ca.setLayout(new PatternLayout("%m%n"));
        ca.setThreshold(Level.DEBUG);
        ca.activateOptions();
        Logger.getLogger("DisplayLogger").addAppender(ca);


    }
}