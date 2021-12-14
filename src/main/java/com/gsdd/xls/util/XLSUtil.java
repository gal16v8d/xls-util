package com.gsdd.xls.util;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.gsdd.constants.XLSConstants;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class XLSUtil {

  /**
   * Get a workbook by extension.
   * 
   * @param is
   * @param excelFilePath output path
   * @param errorMsg message to show if is not a valid extension.
   * @return
   */
  public static Workbook getWorkbook(FileInputStream is, String excelFilePath, String errorMsg) {
    Workbook workbook = null;
    try {
      if (excelFilePath.endsWith(XLSConstants.EXT_XLSX)) {
        workbook = (is != null ? new XSSFWorkbook(is) : new XSSFWorkbook());
      } else if (excelFilePath.endsWith(XLSConstants.EXT_XLS)) {
        workbook = (is != null ? new HSSFWorkbook(is) : new HSSFWorkbook());
      } else {
        throw new IllegalArgumentException(errorMsg);
      }
    } catch (IOException ioe) {
      log.error(ioe.toString(), ioe);
    }
    return workbook;
  }
}
