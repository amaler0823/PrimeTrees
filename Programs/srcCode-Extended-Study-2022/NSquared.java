package com;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class NSquared {
//    int row = 900;
    int row = 100;
    int tmprelt = 0;

    public NSquared() {
    }

    public static void main(String[] args) {
        NSquared nSquared = new NSquared();
        nSquared.writeExcel();
    }

    void writeExcel() {
        String fileName = "Y:\\2021job\\@tree\\测试NSquared.xlsx";
        ((ExcelWriterSheetBuilder)EasyExcel.write(fileName).sheet("NSquared模板").registerWriteHandler(new NSquaredRecordBorrowExcelHandler())).doWrite(this.dataList());
    }

    //    https://segmentfault.com/a/1190000038566393
    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList();
//        int ma = 1;
        int ma = 0;
        int a = 0;

        for(int i = 1; i < this.row; ++i) {
            StringBuilder stringBuilder = new StringBuilder();

            for(int j = ma+1; j < i*i+1; ++j) {
//                a = Integer.valueOf(j + ma)+1;
                a = Integer.valueOf(j);
                String stmp=a+"";
                if (this.numIsP(a)) {
                    stmp = "P" + stmp;
                }
                if(a>(i*i-i) && a<(i*i-i+2)){
                    stmp = "O" + stmp;
                }


//                if (this.numIsInZ(a)) {
//                    stmp = "Z" + stmp;
//                }

                if (this.numIsEven(a)) {
                    stmp = "E" + stmp;
                }

                stringBuilder.append(stmp + ",");
            }

            ma = a;//            https://blog.csdn.net/qq_38974638/article/details/117395831
            list.add(new ArrayList(Arrays.asList(stringBuilder.toString().split(","))));
            stringBuilder.delete(0, stringBuilder.toString().length());
        }

        return list;
    }

    Boolean numIsInZ(int num) {//6- 22,30,39,49,60,72

        //6,10,15,21,28,36,45,55,66,78,91
        int z = 3;
        int a = 3;
        for (int i = 0; i < this.row; i++) {
            z = z + a;
            if (reltIsInt(z, num) && 0!=tmprelt) {
                if((z+1)<=tmprelt){
                    return true;
                }
            }
            tmprelt=0;
            a++;
        }
        return false;
    }

    Boolean reltIsInt(int z, int num) {
        //nn+n-6*2-num*2=0
        //-c
        BigInteger c = BigInteger.valueOf(z).multiply(BigInteger.valueOf(2)).add((BigInteger.valueOf(num).multiply(BigInteger.valueOf(2))));
        //-4ac
        BigInteger fac = BigInteger.valueOf(4).multiply(c);
        //bb-4ac
        BigInteger bbfac = BigInteger.ONE.add(fac);

        if (BigInteger.ZERO.compareTo(bbfac) <= 0) {
            //have solution is
            //sqrt bb-4ac
            BigDecimal dividendONE = BigDecimal.valueOf(Math.sqrt(Double.valueOf(bbfac.toString())));
            //-b +  sqrt bb-4ac
            BigDecimal dividend = dividendONE.subtract(BigDecimal.ONE);
            //2a
            BigDecimal divisor = BigDecimal.valueOf(2);
            /// (-b +  sqrt bb-4ac)/2a
            BigDecimal relt = dividend.divide(divisor);
//                    String ssolu = solu.toString();170.0

            try {
//                System.out.println(relt.intValueExact());
                relt.intValueExact();
            } catch (Exception e) {
                return false;
            }

            tmprelt = relt.intValueExact();
        }

        return true;
    }

    Boolean numIsP(int num) {
        Boolean b = true;
        for (int i = 2; i < num; i++) {
            if (BigInteger.ZERO.equals(BigInteger.valueOf(num).mod(BigInteger.valueOf(i)))) {
                return false;
            }
        }
        return b;
    }

    Boolean numIsEven(int num) {
        Boolean b = true;
//        for (int i = 2; i < num; i++) {
            if (!BigInteger.ZERO.equals(BigInteger.valueOf(num).mod(BigInteger.valueOf(2)))) {
                return false;
            }
//        }
        return b;
    }

}


//https://blog.csdn.net/qq_38637558/article/details/116571437
class NSquaredRecordBorrowExcelHandler implements CellWriteHandler {

    NSquaredRecordBorrowExcelHandler() {
    }

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {
    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean isHead) {
//        if (cell.getColumnIndex() == 2) {
//            if ("20".equals(cell.getNumericCellValue())) {
        if (cell.getStringCellValue().contains("P")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(3);
            sheet.setDefaultRowHeight((short) 170);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }

        if (cell.getStringCellValue().contains("Z")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.RED.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }


        if (cell.getStringCellValue().contains("E")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }



        if (cell.getStringCellValue().contains("OP")) {
            Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
            //size
            Sheet sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short) 180);

            CellStyle cellStyle = workbook.createCellStyle();
            // 填充色, 是设置前景色不是背景色
            cellStyle.setFillForegroundColor(IndexedColors.GOLD.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 边框
            cellStyle.setBorderLeft(BorderStyle.THICK);
            cellStyle.setBorderRight(BorderStyle.THICK);
            cellStyle.setBorderTop(BorderStyle.THICK);
            cellStyle.setBorderBottom(BorderStyle.THICK);
            // 居中
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            //font

            // 将自定义风格写入
            cell.setCellStyle(cellStyle);
        }



    }
//    }
}
