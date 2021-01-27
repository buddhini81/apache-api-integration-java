package com.ourcompany.app;

import com.ourcompany.app.data.QuoteItem;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HeaderFooter;

public class Quotation {
    public void generateQuotation(List<QuoteItem> quoteData, boolean isDomestic) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        try(OutputStream fileOut = new FileOutputStream("src/main/resources/quotation.xlsx")) {
            XSSFSheet sheet = workbook.createSheet("Sheet");

            sheet.setColumnWidth(2, 30 * 256);
            sheet.setColumnWidth(4, 30 * 256);
            sheet.setColumnWidth(10, 20 * 256);

            addCompanyDetails(workbook, sheet);
            addQuoteDetails(workbook, sheet, quoteData);

            PropertyTemplate pt = new PropertyTemplate();
            pt.drawBorders(new CellRangeAddress(12,12,7,10),
                    BorderStyle.MEDIUM,
                    IndexedColors.GREEN.getIndex(),
                    BorderExtent.OUTSIDE);

            pt.drawBorders(new CellRangeAddress(14, 14, 7, 10),
                    BorderStyle.MEDIUM, IndexedColors.BROWN.getIndex(), BorderExtent.OUTSIDE);

            pt.applyBorders(sheet);

            Header header = sheet.getHeader();
            header.setLeft("Furniture Store");
            header.setRight("Quotation");

            Footer footer = sheet.getFooter();
            footer.setCenter("THANK YOU FOR YOUR BUSINESS!");
            footer.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

            sheet.groupRow(0, 4);
            sheet.groupRow(9, 14);

            addInfo(sheet, isDomestic);

            workbook.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addCompanyDetails(XSSFWorkbook workbook, XSSFSheet sheet) {
        XSSFRow titleRow = sheet.createRow(0);
        XSSFCell titleCell = titleRow.createCell(4);

        XSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short)24);

        XSSFCellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setFont(titleFont);
        titleCell.setCellStyle(titleCellStyle);
        titleCell.setCellValue("Product Quote");

        XSSFRow detailsRow1 =  sheet.createRow(1);
        XSSFCell contactNo = detailsRow1.createCell(2);
        contactNo.setCellValue("Contact No: xxxxxxxxxx");
        XSSFCell companyName = detailsRow1.createCell(10);
        companyName.setCellValue("The Furniture Store");

        XSSFRow detailsRow2 =  sheet.createRow(2);
        XSSFCell email = detailsRow2.createCell(2);
        email.setCellValue("Email: contact@furniture.com");

        XSSFRow detailsRow3 =  sheet.createRow(3);
        XSSFCell web = detailsRow3.createCell(2);
        web.setCellValue("Web: www.furniture.com");

        sheet.addMergedRegion(new CellRangeAddress(2, 4, 10, 10));

        CreationHelper helper = workbook.getCreationHelper();

        byte[] bytes;
        try (InputStream is = new FileInputStream("src/main/resources/images/logo.jpg")) {
            bytes = IOUtils.toByteArray(is);
            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            Drawing drawing = sheet.createDrawingPatriarch();

            ClientAnchor anchor = helper.createClientAnchor();

            anchor.setCol1(10);
            anchor.setRow1(2);
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        CellStyle hlink_style = workbook.createCellStyle();
        Font hlinkFont = workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(IndexedColors.BLUE.getIndex());
        hlink_style.setFont(hlinkFont);

        XSSFCell cellWeb = sheet.getRow(3).getCell(2);
        Hyperlink linkUrl = helper.createHyperlink(HyperlinkType.URL);
        linkUrl.setAddress("https://poi.apache.org/");
        cellWeb.setHyperlink(linkUrl);
        cellWeb.setCellStyle(hlink_style);

        XSSFCell cellEmail = sheet.getRow(2).getCell(2);
        Hyperlink linkEmail = helper.createHyperlink(HyperlinkType.EMAIL);
        linkEmail.setAddress("mailto:contact@furniture.com?subject=Inquiry");
        cellEmail.setHyperlink(linkEmail);
        cellEmail.setCellStyle(hlink_style);
    }

    private void addQuoteDetails(XSSFWorkbook workbook, XSSFSheet sheet, List<QuoteItem> quoteData) {
        XSSFRow row1 = sheet.createRow(5);
        XSSFCell cellTo = row1.createCell(2);
        cellTo.setCellValue("To:");

        XSSFCell cellDate = row1.createCell(10);
        cellDate.setCellValue("Date:");

        XSSFRow row2 = sheet.createRow(6);
        XSSFCell cellToVal1 = row2.createCell(2);
        cellToVal1.setCellValue("Mr.John Doe");

        XSSFRow row3 = sheet.createRow(7);
        XSSFCell cellToVal2 = row3.createCell(2);
        cellToVal2.setCellValue("London");

        XSSFRow row4 = sheet.createRow(8);
        XSSFCell cellToVal3 = row4.createCell(2);
        cellToVal3.setCellValue("UK");

        XSSFCellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setAlignment(HorizontalAlignment.LEFT);
        XSSFCreationHelper helper = workbook.getCreationHelper();
        dateCellStyle.setDataFormat(helper.createDataFormat().getFormat("dd/mm/yyyy"));
        XSSFCell cellDateVal = row2.createCell(10);
        cellDateVal.setCellStyle(dateCellStyle);
        cellDateVal.setCellValue(new Date());

        XSSFCellStyle priceCellStyle = workbook.createCellStyle();
        priceCellStyle.setDataFormat(helper.createDataFormat().getFormat("$#,##0.00"));

        XSSFFont font = workbook.createFont();
        font.setBold(true);
        IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
        XSSFColor col = new XSSFColor(IndexedColors.BLUE, colorMap);
        font.setColor(col);

        XSSFCellStyle headerRowCellStyle = workbook.createCellStyle();
        headerRowCellStyle.setFont(font);
        headerRowCellStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFCellStyle boldItalicCellStyle = workbook.createCellStyle();
        XSSFFont boldItalicFont = workbook.createFont();
        boldItalicFont.setBold(true);
        boldItalicFont.setItalic(true);
        boldItalicCellStyle.setFont(boldItalicFont);

        XSSFRow row5 = sheet.createRow(9);
        XSSFCell cell1 = row5.createCell(2);
        cell1.setCellStyle(headerRowCellStyle);
        cell1.setCellValue("Qty");

        XSSFCell cell2 = row5.createCell(4);
        cell2.setCellStyle(headerRowCellStyle);
        cell2.setCellValue("Description");

        XSSFCell cell3 = row5.createCell(7);
        cell3.setCellStyle(headerRowCellStyle);
        cell3.setCellValue("Unit Price");

        XSSFCell cell4 = row5.createCell(10);
        cell4.setCellStyle(headerRowCellStyle);
        cell4.setCellValue("Line Total");

        XSSFCellStyle lineRowCellStyle = workbook.createCellStyle();
        lineRowCellStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFCellStyle taxCellStyle = workbook.createCellStyle();
        taxCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        taxCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        taxCellStyle.setDataFormat(helper.createDataFormat().getFormat("$#,##0.00"));


        int rowIndex = 10;
        double subTotal = 0;
        double tax = 25;

        for(QuoteItem qi : quoteData) {
            XSSFRow linerow = sheet.createRow(rowIndex);
            XSSFCell lineCell1 = linerow.createCell(2);
            lineCell1.setCellStyle(lineRowCellStyle);
            lineCell1.setCellValue(qi.getQty());

            XSSFCell lineCell2 = linerow.createCell(4);
            lineCell2.setCellStyle(lineRowCellStyle);
            lineCell2.setCellValue(qi.getProduct().getName());

            XSSFCell lineCell3 = linerow.createCell(7);
            lineCell3.setCellStyle(priceCellStyle);
            lineCell3.setCellValue(qi.getProduct().getPrice());
            double lineTotal = qi.getQty() * qi.getProduct().getPrice();

            XSSFCell lineCell4 = linerow.createCell(10);
            lineCell4.setCellStyle(priceCellStyle);
            lineCell4.setCellValue(lineTotal);

            subTotal = subTotal + lineTotal;
            rowIndex++;
        }

        XSSFRow subTotalrow = sheet.createRow(rowIndex++);
        XSSFCell subTotalCell1 = subTotalrow.createCell(7);
        subTotalCell1.setCellStyle(boldItalicCellStyle);
        subTotalCell1.setCellValue("Sub Total");

        XSSFCell subTotalCell2 = subTotalrow.createCell(10);
        subTotalCell2.setCellStyle(priceCellStyle);
        subTotalCell2.setCellValue(subTotal);

        XSSFRow taxrow = sheet.createRow(rowIndex++);
        XSSFCell taxCell1 = taxrow.createCell(7);
        taxCell1.setCellStyle(boldItalicCellStyle);
        taxCell1.setCellValue("Tax");

        XSSFCell taxCell2 = taxrow.createCell(10);
        taxCell2.setCellStyle(taxCellStyle);
        taxCell2.setCellValue(tax);

        XSSFRow totalrow = sheet.createRow(rowIndex++);
        XSSFCell totalCell1 = totalrow.createCell(7);
        totalCell1.setCellStyle(boldItalicCellStyle);
        totalCell1.setCellValue("Total");

        XSSFCell totalCell2 = totalrow.createCell(10);
        totalCell2.setCellValue(subTotal + tax);

    }

    private void addInfo(XSSFSheet sheet, boolean isDomestic) {
        XSSFRow infoRo1 = sheet.createRow(16);
        XSSFCell infoCell1 = infoRo1.createCell(2);
        infoCell1.setCellValue("Delivery can be arranged within city limits at a cost");

        if(isDomestic) {
            infoRo1.setZeroHeight(false);
        } else {
            infoRo1.setZeroHeight(true);
        }

        XSSFRow infoRo2 = sheet.createRow(18);
        XSSFCell infoCell2 = infoRo2.createCell(4);
        infoCell2.setCellValue("Should you have any inquiries please contact us");
    }

    public void readQuoteDate() {
        try(InputStream is = new FileInputStream("src/main/resources/quotation.xlsx")) {
            Workbook workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(6);
            Cell cell = row.getCell(10);
            System.out.println("Quote Date: " + cell.getDateCellValue());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void readLogoImage() {
        try (InputStream is = new FileInputStream("src/main/resources/quotation.xlsx")) {
            Workbook workbook = WorkbookFactory.create(is);
            List pictures = workbook.getAllPictures();
            if(pictures != null && pictures.size() == 1) {
                PictureData pict = (PictureData)pictures.get(0);
                byte[] data = pict.getData();

                try(OutputStream out = new FileOutputStream("src/main/resources/images/logo_ext.jpg")) {
                    out.write(data);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void readHyperLink() {
        try {
            try (InputStream is = new FileInputStream("src/main/resources/quotation.xlsx")) {
                Workbook workbook = WorkbookFactory.create(is);
                Sheet sheet = workbook.getSheetAt(0);
                Cell cell = sheet.getRow(3).getCell(2);
                Hyperlink link = cell.getHyperlink();
                if(link != null){
                    System.out.println(link.getAddress());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
