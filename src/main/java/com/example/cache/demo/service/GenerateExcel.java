package com.example.cache.demo.service;

import com.example.cache.demo.dto.UserDto;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static com.example.cache.demo.utils.ExcelUtils.*;
import static java.math.BigInteger.ZERO;
import static java.time.format.DateTimeFormatter.ofPattern;

@Component
public class GenerateExcel {

    private static final String REPORT_NAME = "qrcode recebidos";
    public static final int FIRST_ROW_WITH_CONTENT = 1;

    public List<UserDto> mock() {
        UserDto userDto = new UserDto();
        userDto.setId(1);
        userDto.setNome(null);
        userDto.setEmail("tese@gmail");
        userDto.setSexo("m");


        return Arrays.asList(userDto);
    }

    public ByteArrayOutputStream process(List<UserDto> data) throws Exception {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        List<String> headerReport = getReaderReport();

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            XSSFSheet firstSheet = workbook.createSheet(REPORT_NAME);
            XSSFCellStyle styleHeader = defineStyleHeader(workbook);
            XSSFCellStyle style = workbook.createCellStyle();
            style.setFont(defineFontDefault(workbook.createFont()));

            createHeaderExcel(firstSheet, REPORT_NAME.toUpperCase(), workbook);

            int rowNum = FIRST_ROW_WITH_CONTENT;
            Row row = firstSheet.createRow(rowNum++);
            int colNum = ZERO.intValue();
            for (Object field : headerReport) {
                Cell cell = row.createCell(colNum++);
                cell.setCellStyle(styleHeader);
                cell.setCellValue((String) field);
            }

            createRows(data, firstSheet, style, rowNum, headerReport);
            defineAutoSizeColumn(firstSheet, FIRST_ROW_WITH_CONTENT);

            workbook.write(byteArrayOutputStream);

        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }

        return byteArrayOutputStream;
    }

    private void createRows(List<UserDto> userData,
                            XSSFSheet firstSheet, XSSFCellStyle style, int rowNum, List<String> headerReport) {
        Row row;
        for (UserDto controlledReleaseRequest : userData) {
            List<Object> headerReportValues = headerReport.stream().map(f -> {
                try {
                    Method method = controlledReleaseRequest.getClass().getMethod(String.format("%s%s", "get", f));
                    Object value = method.invoke(controlledReleaseRequest);
                    DateTimeFormatter formatter = ofPattern("yyyy-MM-dd");
                    return treatValueFromMethod(value, formatter);
                } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                    return new Object();
                }
            }).collect(Collectors.toList());

            row = firstSheet.createRow(rowNum++);
            defineCellInRow(row, headerReportValues, style);
        }
    }

    private Object treatValueFromMethod(Object value, DateTimeFormatter formatter) {
        if (Objects.isNull(value)) {
            return "-";
        }
        if (value instanceof LocalDate) {
            value = ((LocalDate) value).format(formatter);
        } else if (value instanceof LocalDateTime) {
            value = ((LocalDateTime) value).format(formatter);
        }
        return value;
    }

    private List<String> getReaderReport() {
        return Arrays.stream(UserDto.class.getDeclaredMethods())
                .filter(f -> f.getName().startsWith("get"))
                .map(f -> f.getName().replace("get", ""))
                .collect(Collectors.toList());
    }

    public ByteArrayOutputStream process1() throws Exception {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        InputStream is = GenerateExcel.class.getClassLoader().getResourceAsStream("qrcode.xlsx");

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(is);
            XSSFSheet firstSheet = workbook.getSheet(REPORT_NAME);

            XSSFCellStyle style = workbook.createCellStyle();
            style.setFont(defineFontDefault(workbook.createFont()));

            List<String> headerReport = getReaderReport();

            int rowNum = 2;

            createRows(mock(), firstSheet, style, rowNum, headerReport);
            defineAutoSizeColumn(firstSheet, FIRST_ROW_WITH_CONTENT);

            workbook.write(byteArrayOutputStream);
        } catch (Exception e) {
            throw new Exception(e.getMessage());
        }

        return byteArrayOutputStream;
    }
}
