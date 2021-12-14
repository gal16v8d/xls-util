package com.gsdd.xls.util;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class XLSUtilTest {

  private static final String TEST_XLS = "test.xls";
  private static final String TEST_XLSX = "test.xlsx";
  private static final String ERROR = "error!";

  @Test
  void testGetWorkbookWrongExt(@Mock FileInputStream stream) {
    Assertions.assertThrows(IllegalArgumentException.class,
        () -> XLSUtil.getWorkbook(stream, "xls.xlst", ERROR));
  }

  @Test
  void testGetWorkbookIsXLSX() {
    try (FileInputStream stream =
        (FileInputStream) XLSUtilTest.class.getResourceAsStream(TEST_XLSX)) {
      Assertions.assertTrue(XLSUtil.getWorkbook(stream, TEST_XLSX, ERROR) instanceof XSSFWorkbook);
    } catch (IOException e) {
      Assertions.fail(ERROR);
    }
  }

  @Test
  void testGetWorkbookIsXLS() {
    try (FileInputStream stream =
        (FileInputStream) XLSUtilTest.class.getResourceAsStream(TEST_XLS)) {
      Assertions.assertTrue(XLSUtil.getWorkbook(stream, TEST_XLS, ERROR) instanceof HSSFWorkbook);
    } catch (IOException e) {
      Assertions.fail(ERROR);
    }

  }
}
