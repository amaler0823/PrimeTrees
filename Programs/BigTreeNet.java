package com;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;


public class BigTreeNet {
//    int row = 900;
    int row = 100;
    int tmprelt = 0;
    StringBuffer jcoll = new StringBuffer();

    public BigTreeNet() {
    }

    public static void main(String[] args) {
        BigTreeNet bigTreeNet = new BigTreeNet();
        bigTreeNet.writeExcel();
    }

    void writeExcel() {
        this.getJCollection();
        String fileName = "Y:\\2021job\\@tree\\测试big.xlsx";
        ((ExcelWriterSheetBuilder)EasyExcel.write(fileName).sheet("big模板").registerWriteHandler(new RzhongRecordBorrowExcelHandler())).doWrite(this.dataList());
    }

    //    https://segmentfault.com/a/1190000038566393
    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList();
        int ma = 1;
        int a = 0;

        for(int i = 0; i < this.row; ++i) {
            StringBuilder stringBuilder = new StringBuilder();

            for(int j = 0; j < i; ++j) {
                a = Integer.valueOf(j + ma)+1;String stmp=a+"";
                if (this.numIsP(a)) {
                    stmp = "P" + stmp;
                }

                if (this.numIsInZ(a)) {
                    stmp = "Z" + stmp;
                }

                String[] sj = this.jcoll.toString().split(",");
                String[] var9 = sj;
                int var10 = sj.length;

                for(int var11 = 0; var11 < var10; ++var11) {
                    String tmp = var9[var11];
                    if ((a+"").equals(tmp)) {
                        stmp = "J" + stmp;
                    }
                }

                stringBuilder.append(stmp + ",");
            }

            ma = a;//            https://blog.csdn.net/qq_38974638/article/details/117395831
            list.add(new ArrayList(Arrays.asList(stringBuilder.toString().split(","))));
            stringBuilder.delete(0, stringBuilder.toString().length());
        }

        return list;
    }

    void getJCollection() {
        StringBuffer tmpColl = new StringBuffer();
        int c = 6;
        int top = 3;
        int bot = 5;
        Boolean hasP = false;
        do {
            for(int i = 0; i < 2; ++i) {

//                System.out.println(top);//3,,3,,4
//                System.out.println(bot);//5,,6,,7
//                System.out.println(c);//6,,7,,8

                for(int j = top; j <= bot; ++j) {
                    int tmpNum = j * (j - 1) / 2 + c;

                    if (this.numIsP(tmpNum)) {
                        hasP = true;
                    }
                    tmpColl.append(tmpNum + ",");
                }

                if (hasP) {
                    tmpColl.delete(0, tmpColl.toString().length());
                    hasP=false;
                } else {
                    this.jcoll.append(tmpColl);
                    tmpColl.delete(0, tmpColl.toString().length());
                }
                ++c;++bot;
            }
            ++top;
        } while(c <= this.row + 6);

    }

    Boolean numIsInZ(int num) {
        int z = 3;
        int a = 3;
        for(int i = 0; i < row; ++i) {
            z += a;
            if (this.rzReltIsInt(z, num) && 0 != this.tmprelt && z + 1 <= this.tmprelt) {
                return true;
            }
            this.tmprelt = 0;++a;
        }
        return false;
    }

    Boolean rzReltIsInt(int z, int num) {
        BigInteger c = BigInteger.valueOf((long)z).multiply(BigInteger.valueOf(2)).add(BigInteger.valueOf((long)num).multiply(BigInteger.valueOf(2)));
        BigInteger fac = BigInteger.valueOf(4L).multiply(c);
        BigInteger bbfac = BigInteger.ONE.add(fac);
        if (BigInteger.ZERO.compareTo(bbfac) <= 0) {
            BigDecimal dividendONE = BigDecimal.valueOf(Math.sqrt(Double.valueOf(bbfac.toString())));
            BigDecimal dividend = dividendONE.subtract(BigDecimal.ONE);
            BigDecimal divisor = BigDecimal.valueOf(2L);
            BigDecimal relt = dividend.divide(divisor);
            try {
                relt.intValueExact();
            } catch (Exception var11) {
                return false;
            }
            this.tmprelt = relt.intValueExact();
        }
        return true;
    }

    Boolean numIsP(int num) {
        Boolean b = true;
        for(int i = 2; i < num; ++i) {
            if (BigInteger.ZERO.equals(BigInteger.valueOf((long)num).mod(BigInteger.valueOf((long)i)))) {
                return false;
            }
        }
        return b;
    }
}


class RzhongRecordBorrowExcelHandler implements CellWriteHandler {
    RzhongRecordBorrowExcelHandler() {
    }

    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer integer, Integer integer1, Boolean aBoolean) {
    }

    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer integer, Boolean aBoolean) {
    }

    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> list, Cell cell, Head head, Integer integer, Boolean isHead) {
        Workbook workbook;
        Sheet sheet;
        CellStyle cellStyle;
        if (cell.getStringCellValue().contains("P")) {
            workbook = writeSheetHolder.getSheet().getWorkbook();
            sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short)180);
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(cellStyle);
        }

        if (cell.getStringCellValue().contains("Z")) {
            workbook = writeSheetHolder.getSheet().getWorkbook();
            sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short)180);
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(cellStyle);
        }

        if (cell.getStringCellValue().contains("J")) {
            workbook = writeSheetHolder.getSheet().getWorkbook();
            sheet = writeSheetHolder.getSheet();
            sheet.setDefaultColumnWidth(5);
            sheet.setDefaultRowHeight((short)180);
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.RED.index);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cell.setCellStyle(cellStyle);
        }

    }
}
