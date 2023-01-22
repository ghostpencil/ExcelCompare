package com.ka.spreadsheet.diff;


import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import java.util.LinkedHashSet;
import java.util.Set;

public class LoggerSpreadSheetDiffCallback extends SpreadSheetDiffCallbackBase {
  private static final Logger cleanLogger = LogManager.getLogger("DisplayLogger");
  private final Set<Object> sheets = new LinkedHashSet<Object>();
  private final Set<Object> rows = new LinkedHashSet<Object>();
  private final Set<Object> cols = new LinkedHashSet<Object>();
  private final Set<Object> macros = new LinkedHashSet<Object>();

  private final Set<Object> sheets1 = new LinkedHashSet<Object>();
  private final Set<Object> rows1 = new LinkedHashSet<Object>();
  private final Set<Object> cols1 = new LinkedHashSet<Object>();

  private final Set<Object> sheets2 = new LinkedHashSet<Object>();
  private final Set<Object> rows2 = new LinkedHashSet<Object>();
  private final Set<Object> cols2 = new LinkedHashSet<Object>();

  private final Set<Object> macros1 = new LinkedHashSet<Object>();
  private final Set<Object> macros2 = new LinkedHashSet<Object>();

  private String file1;
  private String file2;

  @Override
  public void init(String file1, String file2) {
    super.init(file1, file2);
    this.file1 = file1;
    this.file2 = file2;
  }

  @Override
  public void reportWorkbooksDiffer(boolean differ) {
    super.reportWorkbooksDiffer(differ);
    reportSummary("DIFF", sheets, rows, cols, macros);
    reportSummary("EXTRA WB1", sheets1, rows1, cols1, macros1);
    reportSummary("EXTRA WB2", sheets2, rows2, cols2, macros2);
    cleanLogger.info("-----------------------------------------");
    cleanLogger.info("Excel files " + file1 + " and " + file2 + " "
            + (differ ? "differ" : "match"));
  }

  @Override
  public void reportMacroOnlyIn(boolean inFirstSpreadSheet) {
    super.reportMacroOnlyIn(inFirstSpreadSheet);
    String name = "unknown";
    (inFirstSpreadSheet ? macros1 : macros2).add(name);
    cleanLogger.info("EXTRA macro name: " + name + " found only in " + wb(inFirstSpreadSheet));
  }

  @Override
  public void reportExtraCell(boolean inFirstSpreadSheet, CellPos c) {
    super.reportExtraCell(inFirstSpreadSheet, c);
    if (inFirstSpreadSheet) {
      sheets1.add(c.getSheetName());
      rows1.add(c.getRow());
      cols1.add(c.getColumn());
    } else {
      sheets2.add(c.getSheetName());
      rows2.add(c.getRow());
      cols2.add(c.getColumn());
    }
    cleanLogger.info("EXTRA Cell in " + wb(inFirstSpreadSheet) + " " + c.getCellPosition()
        + " => '" + c.getCellValue() + "'");
  }

  @Override
  public void reportDiffCell(CellPos c1, CellPos c2) {
    super.reportDiffCell(c1, c2);
    sheets.add(c1.getSheetName());
    rows.add(c1.getRow());
    cols.add(c1.getColumn());
    cleanLogger.info("DIFF  Cell at     " + c1.getCellPosition() + " => '" + c1.getCellValue()
        + "' v/s '" + c2.getCellValue() + "'");
  }

  private void reportSummary(String what, Set<Object> sheets, Set<Object> rows, Set<Object> cols,
      Set<Object> macros) {
    cleanLogger.info("----------------- " + what + " -------------------");
    cleanLogger.info("Sheets: " + sheets);
    cleanLogger.info("Rows: " + rows);
    cleanLogger.info("Cols: " + cols);
    if (!macros.isEmpty()) {
      cleanLogger.info("Macros: " + macros);
    }
  }

  private String wb(boolean inFirstSpreadSheet) {
    return inFirstSpreadSheet ? "WB1" : "WB2";
  }
}
