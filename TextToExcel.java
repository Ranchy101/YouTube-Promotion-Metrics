package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class TextToExcel {

    public static void main(String[] args) {
        String path = "promotions.txt";
        List<VideoData> videos = readTextFile(path);

        List<VideoData> activePromotions = new ArrayList<>();
        List<VideoData> endedPromotions = new ArrayList<>();

        for (VideoData video : videos) {
            if ("Active".equalsIgnoreCase(video.status)) {
                activePromotions.add(video);
            } else {
                endedPromotions.add(video);
            }
        }

        activePromotions.sort(Comparator.comparingDouble(VideoData::getImpressionToSubRatio).reversed());
        endedPromotions.sort(Comparator.comparingDouble(VideoData::getImpressionToSubRatio).reversed());

        List<VideoData> sortedVideos = new ArrayList<>(activePromotions);
        sortedVideos.addAll(endedPromotions);

        writeToExcel(sortedVideos, "output.xlsx");
    }

    public static List<VideoData> readTextFile(String path) {
        List<VideoData> videos = new ArrayList<>();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(path)))) {
            // Read all lines into a list
            List<String> lines = reader.lines().collect(Collectors.toList());

            // Skip the first 6 lines and exclude the last 3 lines
            for (int i = 6; i < lines.size() - 3; i++) {
                String line = lines.get(i);
                if (line.startsWith("Video thumbnail:")) {
                    VideoData video = new VideoData();
                    i++;  // Skip next line
                    video.title = lines.get(++i);
                    video.status = lines.get(++i);
                    video.cost = Double.parseDouble(lines.get(++i).substring(1));
                    video.impressions = Integer.parseInt(lines.get(++i).replaceAll(",", ""));
                    video.views = Integer.parseInt(lines.get(++i).replaceAll(",", ""));
                    video.subscribers = Integer.parseInt(lines.get(++i).replaceAll(",", ""));
                    video.costPerSub = (video.subscribers == 0) ? 0 : video.cost / video.subscribers;
                    video.impressionToSubRatio = (video.subscribers == 0) ? 0 : ((double) video.subscribers / video.impressions) * 100; // as a percentage
                    videos.add(video);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return videos;
    }

    public static void writeToExcel(List<VideoData> videos, String outputPath) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Videos");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Title");
        header.createCell(1).setCellValue("Status");
        header.createCell(2).setCellValue("Cost");
        header.createCell(3).setCellValue("Impressions");
        header.createCell(4).setCellValue("Views");
        header.createCell(5).setCellValue("Subscribers");
        header.createCell(6).setCellValue("Cost per Sub");
        header.createCell(7).setCellValue("Impression to Sub Ratio");
        header.createCell(9).setCellValue("Total Active:");
        header.createCell(10).setCellValue("Total:");

        // Styling for the header row
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        for (Cell cell : header) {
            cell.setCellStyle(headerStyle);
        }

        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setAlignment(HorizontalAlignment.RIGHT);
        DataFormat format = workbook.createDataFormat();
        currencyStyle.setDataFormat(format.getFormat("$0.00"));

        DecimalFormat df = new DecimalFormat("0.00");

        // Define pink alternating color
        XSSFCellStyle customPinkStyle = (XSSFCellStyle) workbook.createCellStyle();
        customPinkStyle.setAlignment(HorizontalAlignment.RIGHT);
        XSSFColor customPink = new XSSFColor(new java.awt.Color(253, 225, 232));
        customPinkStyle.setFillForegroundColor(customPink);
        customPinkStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle titleStatusPinkStyle = (XSSFCellStyle) workbook.createCellStyle();
        titleStatusPinkStyle.setFillForegroundColor(customPink);
        titleStatusPinkStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle currencyPinkStyle = (XSSFCellStyle) workbook.createCellStyle();
        currencyPinkStyle.setDataFormat(format.getFormat("$0.00"));
        currencyPinkStyle.setFillForegroundColor(customPink);
        currencyPinkStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        currencyPinkStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle whiteStyle = workbook.createCellStyle();
        whiteStyle.setAlignment(HorizontalAlignment.RIGHT);
        whiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle titleStatusWhiteStyle = (XSSFCellStyle) workbook.createCellStyle();
        titleStatusWhiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        titleStatusWhiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle currencyWhiteStyle = (XSSFCellStyle) workbook.createCellStyle();
        currencyWhiteStyle.setDataFormat(format.getFormat("$0.00"));
        currencyWhiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        currencyWhiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        currencyWhiteStyle.setAlignment(HorizontalAlignment.RIGHT);

        int currentRowNum = 1;
        boolean endedLabelAdded = false;
        int colorCounter = 2;

        for (VideoData video : videos) {
            if (!endedLabelAdded && !"Active".equalsIgnoreCase(video.status)) {
                Row labelRow = sheet.createRow(currentRowNum++);
                CellStyle endedStyle = workbook.createCellStyle();
                Font endedFont = workbook.createFont();
                endedFont.setBold(true);
                endedStyle.setFont(endedFont);
                endedStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
                endedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                for (int i = 0; i <= 7; i++) {  // Stops at "Impression to Sub Ratio" column
                    Cell labelCell = labelRow.createCell(i);
                    labelCell.setCellStyle(endedStyle);
                }

                labelRow.getCell(0).setCellValue("Ended Promotions");
                endedLabelAdded = true;
                colorCounter++;
            }

            Row row = sheet.createRow(currentRowNum++);

            // Set alternate colors for "Title" and "Status" columns
            CellStyle titleStatusStyle = (colorCounter % 2 == 0) ? titleStatusWhiteStyle : titleStatusPinkStyle;
            Cell titleCell = row.createCell(0);
            titleCell.setCellValue(video.title);
            titleCell.setCellStyle(titleStatusStyle);
            Cell statusCell = row.createCell(1);
            statusCell.setCellValue(video.status);
            statusCell.setCellStyle(titleStatusStyle);

            CellStyle currentStyle = (colorCounter % 2 == 0) ? whiteStyle : customPinkStyle;

            for (int i = 2; i < 8; i++) {
                row.createCell(i).setCellStyle(currentStyle);
            }

            row.getCell(2).setCellValue("$" + df.format(video.cost));
            row.getCell(3).setCellValue(video.impressions);
            row.getCell(4).setCellValue(video.views);
            row.getCell(5).setCellValue(video.subscribers);

            Cell costPerSubCell = row.getCell(6);
            costPerSubCell.setCellValue(video.costPerSub);
            costPerSubCell.setCellStyle((colorCounter % 2 == 0) ? currencyWhiteStyle : currencyPinkStyle);

            row.getCell(7).setCellValue(df.format(video.impressionToSubRatio) + "%");

            colorCounter++;
        }

        Row totalValueRow = sheet.createRow(1);

        CellStyle totalActiveStyle = workbook.createCellStyle();
        totalActiveStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        totalActiveStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle totalStyle = workbook.createCellStyle();
        totalStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Cell totalActiveHeader = header.getCell(9);
        totalActiveHeader.setCellStyle(totalActiveStyle);
        Cell totalActiveCell = totalValueRow.createCell(9);
        totalActiveCell.setCellValue(videos.stream().filter(v -> "Active".equalsIgnoreCase(v.status)).count());
        totalActiveCell.setCellStyle(totalActiveStyle);

        Cell totalHeader = header.getCell(10);
        totalHeader.setCellStyle(totalStyle);
        Cell totalCell = totalValueRow.createCell(10);
        totalCell.setCellValue(videos.size());
        totalCell.setCellStyle(totalStyle);

        sheet.setColumnWidth(0, 88 * 256);
        sheet.setColumnWidth(1, 15 * 256);
        sheet.setColumnWidth(2, 12 * 256);
        sheet.setColumnWidth(3, 12 * 256);
        sheet.setColumnWidth(4, 9 * 256);
        sheet.setColumnWidth(5, 11 * 256);
        sheet.setColumnWidth(6, 12 * 256);
        sheet.setColumnWidth(7, 24 * 256);
        sheet.setColumnWidth(9, (int)(11.5 * 256));
        sheet.setColumnWidth(10, (int)(5.4 * 256));

        try (FileOutputStream fileOut = new FileOutputStream(outputPath)) {
            workbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static class VideoData {
        String title;
        String status;
        double cost;
        int impressions;
        int views;
        int subscribers;
        double costPerSub;
        double impressionToSubRatio;

        public double getImpressionToSubRatio() {
            return impressionToSubRatio;
        }
    }
}