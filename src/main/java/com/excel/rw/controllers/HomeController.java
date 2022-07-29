package com.excel.rw.controllers;

import com.excel.rw.domains.User;
import com.excel.rw.repo.UserRepo;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@RestController
public class HomeController {


    @Autowired
    private UserRepo userRepo;

    @GetMapping
    public String home()
    {
        return "home";
    }

    @PostMapping("/readExcel")
    public ResponseEntity<?> readExcelFile(@RequestParam("file")MultipartFile file)
    {
        try {
                if(!file.isEmpty() )
                {
                    //First Way
                   //List<User> userList = this.getExcelData(file.getInputStream());

                    //Second Way
                    List<User> userList = this.readExcelDataSecondWay(file.getInputStream());

                   System.out.println(userList.toString());

                   return ResponseEntity.status(HttpStatus.OK).body("File Upload Successfully");
                }
                else {
                    throw  new FileNotFoundException("File Not Found Here..");
                }
        }
        catch (Exception e)
        {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }


    public  List<User>  readExcelDataFirstWay(InputStream inputStream)
    {
        List<User> userList = new ArrayList<>();

        DataFormatter dataFormatter = new DataFormatter();

        try {
            Workbook workbook = new HSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet("sheet1");

           Iterator<Row> rows = sheet.iterator();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();
                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }
                Iterator<Cell> cellsInRow = currentRow.iterator();
                int cellIdx = 0;
                while (cellsInRow.hasNext()) {

                    //Creating Object
                    User user = new User();

                    Cell currentCell = cellsInRow.next();
                    switch (cellIdx) {
                        case 0:
                            user.setUserId((long) currentCell.getNumericCellValue());
                            break;
                        case 1:
                            user.setFirstName(String.valueOf(currentCell.getStringCellValue()));
                            break;
                        case 2:
                            user.setLastName(String.valueOf(currentCell.getStringCellValue()));
                            break;
                        case 3:
                            user.setGender(String.valueOf(currentCell.getStringCellValue()));
                            break;
                        case 4:
                            user.setCountry(String.valueOf(currentCell.getStringCellValue()));
                            break;
                        case 5:
                            user.setAge(String.valueOf(currentCell.getNumericCellValue()));
                            break;
                        case 6:
                            user.setDate(String.valueOf(currentCell.getStringCellValue()));
                            break;
                        case 7:
                            user.setId((long)currentCell.getNumericCellValue());
                            break;
                    }
                    userList.add(user);

                    cellIdx++;
                }
            }
            workbook.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return userList;
    }


    //Second way
    public  List<User>  readExcelDataSecondWay(InputStream inputStream)
    {
        List<User> userList = new ArrayList<>();

        DataFormatter dataFormatter = new DataFormatter();

        try {
            Workbook workbook = new HSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet("sheet1");

            int index = 0;

            for(Row row : sheet) {

                User user  = new User();

                if(index++ == 0) continue;
                user.setId(Long.parseLong(dataFormatter.formatCellValue(row.getCell(0))));
                user.setFirstName(String.valueOf(dataFormatter.formatCellValue(row.getCell(1))));
                user.setLastName(String.valueOf(dataFormatter.formatCellValue(row.getCell(2))));
                user.setGender(String.valueOf(dataFormatter.formatCellValue(row.getCell(3))));
                user.setCountry(String.valueOf(dataFormatter.formatCellValue(row.getCell(4))));
                user.setAge(String.valueOf(dataFormatter.formatCellValue(row.getCell(5))));
                user.setDate(String.valueOf(dataFormatter.formatCellValue(row.getCell(6))));
                user.setUserId(Long.parseLong(dataFormatter.formatCellValue(row.getCell(7))));

                userList.add(user);
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return userList;
    }





    //write And Download Excel Report

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<User> listUsers;

    @GetMapping("/writeAndDownloadExcelReport")
    public void writeAndDownloadExcelReport(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=users_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

         listUsers = userRepo.findAll();

         workbook = new XSSFWorkbook();

        this.export(response);
    }

    private void writeHeaderLine() {
        sheet = workbook.createSheet("Users");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(12);
        style.setFont(font);

        createCell(row, 0, "User ID", style);

        createCell(row, 1, "Title", style);

        createCell(row, 2, "body", style);

    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else if (value instanceof Long){
            cell.setCellValue((Long) value);
        }
        else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void writeDataLines() {
        int rowCount = 1;

        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(10);
        style.setFont(font);

        for (User user : listUsers) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;

            createCell(row, columnCount++, user.getId(), style);

            createCell(row, columnCount++, user.getTitle(), style);

            createCell(row, columnCount++, user.getBody(), style);

        }
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeDataLines();

        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();

        outputStream.close();
    }

    //    public  void UserExcelExporter(List<User> listUsers) {
//        this.listUsers = listUsers;
//
//    }

}
