package controller;

import model.Pair;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;

public class Calendar {

    Workbook wb;
    HashMap<String, Pair> mapHour;
    Integer indexRows = 2;

    public Calendar() {
        wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("Sheet1");
        Row row = sheet1.createRow(0);
        Cell cell;

       String []init = {"Subject", "Start Date", "Start Time", "End Date",
               "End Time", "All Day Event", "Description", "Location", "Private"};
       for (int i = 0; i < init.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(init[i]);
       }
       mapHour = new HashMap<>();
       mapHour.put("1,2,3", new Pair("7:30", "9:25"));
       mapHour.put("4,5,6", new Pair("9:35", "12:00"));
       mapHour.put("7,8,9", new Pair("13:00", "15:25"));
       mapHour.put("10,11,12", new Pair("15:35", "18:00"));
    }



    private String[] processDuration(String example) {
        String []fromTo = example.trim().split("\\s+");
        return new String[]{fromTo[1], fromTo[3].substring(0, fromTo[3].length() - 1)};
    }

    private String[] processDay(String example) {
        String []fromTo = example.trim().split("\\s+");
        String []result = new String[3];
        result[2] = "";
        result[0] = fromTo[1];
        result[1] = fromTo[3];
        for (int i = 5; i < fromTo.length; i++) {
            if (fromTo[i] == null) continue;
            result[2] += fromTo[i] + " ";
        }
        System.out.println(result[2]);
        return result;
    }

    private void processParams(String s, String subject) throws IOException {
        try (BufferedReader br = new BufferedReader(new StringReader(s))) {
            String line;
            String []process = new String[2];
            Sheet outputSheet = wb.getSheetAt(0);
            String []process2;
            while ((line = br.readLine()) != null) {
                String tmp = line.trim().split("\\s+")[0];

                if (tmp.equals("Từ")) {
                    // process duration
                    process = processDuration(line);

                } else {
                    // process time and location
                    process2 = processDay(line);

                    // setup everyday
                    LocalDate dateStart = LocalDate.parse(process[0], DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                    LocalDate dateEnd = LocalDate.parse(process[1], DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                    dateStart = dateStart.plusDays(Integer.parseInt(process2[0]) - 2);

                    // timeline in day
                    Pair time = mapHour.get(process2[1]);
                    while (dateStart.isBefore(dateEnd)) {
                        Row outputRows = outputSheet.createRow(indexRows);

                        // set subject
                        Cell outputCell = outputRows.createCell(0);
                        outputCell.setCellValue(subject);

                        // set time start and time end
                        outputCell = outputRows.createCell(2);
                        outputCell.setCellValue(time.getFrom());
                        outputCell = outputRows.createCell(4);
                        outputCell.setCellValue(time.getTo());

                        // set day
                        outputCell = outputRows.createCell(1);
                        outputCell.setCellValue(dateStart.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
                        outputCell = outputRows.createCell(3);
                        outputCell.setCellValue(dateStart.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

                        // set location
                        outputCell = outputRows.createCell(7);
                        outputCell.setCellValue("3 Đ. Cầu Giấy, Ngọc Khánh, Đống Đa, Hà Nội, Việt Nam");
                        // set all-day event
                        outputCell = outputRows.createCell(5);
                        outputCell.setCellValue("False");

                        // set private
                        outputCell = outputRows.createCell(8);
                        outputCell.setCellValue("True");

                        // set description
                        outputCell = outputRows.createCell(6);
                        outputCell.setCellValue(process2[2]);

                        dateStart = dateStart.plusDays(7);
                        indexRows++;
                    }
                }

            }
        }
    }

    public void processExcel(String path, String start, String end) throws IOException {
        InputStream inputStream = new FileInputStream(path);
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int[] startCell = convert(start);
        int[] endCell = convert(end);

        // create output


        for (int i = startCell[0] - 1; i < endCell[0]; i++) {
            // name of subject
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(startCell[1]);
            String c1 = cell.getStringCellValue();
            if (c1 == null) break;
            // calendar
            cell = row.getCell( startCell[1] + 2);
            String c2 = cell.getStringCellValue();
            processParams(c2, c1);
            //System.out.println(c1 + "\n" + c2);
        }
        exportToExcel();
        exportToCSV();
    }

    private int[] convert(String s) {
        int row = Integer.parseInt(s.substring(1));
        int col = s.charAt(0) - 'A';
        return new int[]{row, col};
    }

    private void exportToExcel() {
        try (FileOutputStream fileOut = new FileOutputStream("export_calendar.xls")) {
            wb.write(fileOut);
            System.out.println("Excel file has been exported successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void exportToCSV() {
        Sheet sheet = wb.getSheetAt(0);

        try (BufferedWriter bw = new BufferedWriter(new FileWriter("export_calendar.csv"))) {
            for (Row row : sheet) {
                StringBuilder sb = new StringBuilder();

                for (Cell cell : row) {
                    String cellValue = cell.getStringCellValue();
                    sb.append(cellValue).append(",");
                }

                if (!sb.isEmpty()) {
                    sb.setLength(sb.length() - 1);
                }

                bw.write(sb.toString());
                bw.newLine();
            }

            System.out.println("Excel file has been converted to CSV successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
