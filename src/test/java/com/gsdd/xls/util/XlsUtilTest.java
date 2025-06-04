package com.gsdd.xls.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.stream.Stream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.NullSource;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

@ExtendWith(MockitoExtension.class)
class XlsUtilTest {

  private static final String TEST_XLS = "test.xls";
  private static final String TEST_XLSX = "test.xlsx";
  private static final String ERROR = "error!";

  @Test
  void testGetWorkbookWrongExt(@Mock FileInputStream stream) {
    Assertions.assertThrows(
        IllegalArgumentException.class, () -> XlsUtil.getWorkbook(stream, "xls.xlst", ERROR));
  }

  @ParameterizedTest
  @NullSource
  @MethodSource({"getResourceArg"})
  void testGetWorkbookIsXlsx(FileInputStream streamFile) {
    try (FileInputStream stream = streamFile) {
      Assertions.assertInstanceOf(
          XSSFWorkbook.class, XlsUtil.getWorkbook(stream, TEST_XLSX, ERROR));
    } catch (IOException e) {
      Assertions.fail(ERROR);
    }
  }

  @ParameterizedTest
  @NullSource
  @MethodSource({"getResourceArg"})
  void testGetWorkbookIsXls(FileInputStream streamFile) {
    try (FileInputStream stream = streamFile) {
      Assertions.assertInstanceOf(HSSFWorkbook.class, XlsUtil.getWorkbook(stream, TEST_XLS, ERROR));
    } catch (IOException e) {
      Assertions.fail(ERROR);
    }
  }

  static Stream<Arguments> getResourceArg() {
    return Stream.of(Arguments.of(XlsUtilTest.class.getResourceAsStream(TEST_XLS)));
  }
}
