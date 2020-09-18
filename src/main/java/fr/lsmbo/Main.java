package fr.lsmbo;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.javatuples.Triplet;
import org.apache.poi.xssf.usermodel.*;
import org.apache.log4j.Logger;

public class Main {

    private static final Logger logger = Logger.getLogger(Main.class);
    private static final CellCopyPolicy policy = new CellCopyPolicy();

    public static void main(String[] args) {
        File input = null, output = null;
        int sheetNumber = 0;
        char c1 = 'A', c2 = 'B';
        Character[] conds = {};
        try {
            // get input arguments
            logger.info("Get input arguments");
            input = new File(args[0]);
            sheetNumber = Integer.parseInt(args[1]);
            c1 = args[2].charAt(0);
            c2 = args[3].charAt(0);
            conds = Arrays.stream(args[4].split(",")).map(c -> c.charAt(0)).distinct().sorted().toArray(Character[]::new);
            output = new File(args[5]);

            // TODO add some verifications

        } catch(Throwable t) {
            logger.error("Error in the list of arguments"+t.getMessage(), t);
            System.err.println("\nUsage: "+Main.class.getPackage().getName()+" <input file path> <sheet number> <main column letter> <secondary column letter> <condition columns letters> <output file path>");
            System.err.println("\t<input file path> : The path of the XLSX file to read");
            System.err.println("\t<sheet number> : The number of the sheet to process (first sheet is 0)");
            System.err.println("\t<main column letter> : The letter of the main column to consider, such as Sample Name (ie. A)");
            System.err.println("\t<secondary column letter> : The letter of the secondary column, such as Protein Accession Number (ie. B)");
            System.err.println("\t<condition columns letters> : The letters of the columns used as conditions, separated by a coma (ie. C,D,E)");
            System.err.println("\t<output file path> : The path of the XLSX file to generate\n");
            System.exit(1);
        }

        try {
            // prepare variables
            policy.setCopyCellStyle(false);
            logger.info("Prepare variables");
            HashMap<String, Boolean> c1Values = new HashMap<>();
            HashMap<String, Boolean> c2Values = new HashMap<>();
            HashMap<Triplet<String, String, String>, ArrayList<Object>> values = new HashMap<>();
            HashMap<Triplet<String, String, String>, ArrayList<XSSFCell>> cells = new HashMap<>();

            int column1 = getNumber(c1);
            int column2 = getNumber(c2);
            Integer[] conditions = Arrays.stream(conds).map(Main::getNumber).toArray(Integer[]::new);
            String c1Name = "";
            String[] conditionNames = new String[conditions.length];

            // read input file and get unique values for main and second columns
            logger.info("Read input file and get unique values for main and second columns");
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(input);
            XSSFSheet sheet = inputWorkbook.getSheetAt(sheetNumber);
            for(int ln = sheet.getFirstRowNum(); ln <= sheet.getLastRowNum(); ln++) {
                XSSFRow row = inputWorkbook.getSheetAt(sheetNumber).getRow(ln);
                if(row.getRowNum() == sheet.getFirstRowNum()) {
                    c1Name = getString(row, column1);
                    for(int i = 0; i < conditions.length; i++) {
                        conditionNames[i] = getString(row, conditions[i]);
                    }
                } else {
                    String cell1 = getString(row, column1);
                    String cell2 = getString(row, column2);
                    c1Values.put(cell1, true);
                    c2Values.put(cell2, true);
                    for (int i = 0; i < conditions.length; i++) {
                        Triplet<String, String, String> key = new Triplet<>(cell1, cell2, conditionNames[i]);
                        if (!values.containsKey(key)) values.put(key, new ArrayList<>());
                        if (!cells.containsKey(key)) cells.put(key, new ArrayList<>());
                        cells.get(key).add(row.getCell(conditions[i]));
                    }
                }
            }
            inputWorkbook.close();

            List<String> c1Names = c1Values.keySet().stream().sorted().collect(Collectors.toList());
            List<String> c2Names = c2Values.keySet().stream().sorted().collect(Collectors.toList());

            // create the output file
            logger.info("Create the output file");
            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            sheet = outputWorkbook.createSheet();
            XSSFRow row1 = sheet.createRow(0);
            XSSFRow row2 = sheet.createRow(1);
            // writing header lines
            write(row2, 0, c1Name);
            for(int i = 0; i < conditionNames.length; i++) {
                write(row1, c2Values.size()*(i+1)-1, conditionNames[i]);
                // merge the header cells
                sheet.addMergedRegion(new CellRangeAddress(0, 0, c2Values.size()*(i+1)-1, c2Values.size()*(i+1)));
                for(int j = 0; j < c2Names.size(); j++) {
                    write(row2, i*c2Names.size()+j+1, c2Names.get(j));
                }
            }
            // add autofilter on second line
            sheet.setAutoFilter(new CellRangeAddress(1, 1, 0, conditionNames.length*c2Values.size()));
            // writing the rest of the file
            logger.info("Writing the rest of the file");
            int i = 2;
            for(String first : c1Names) {
                XSSFRow row = sheet.createRow(i++);
                write(row, 0, first);
                for(int j = 0; j < conditionNames.length; j++) {
                    for(int k = 0; k < c2Names.size(); k++) {
                        Triplet<String, String, String> key = new Triplet<>(first, c2Names.get(k), conditionNames[j]);
                        if(cells.containsKey(key)) {
                            writeArray(row, j * c2Names.size() + k + 1, cells.get(key));
                        }
                    }
                }
            }

            // add borders
            PropertyTemplate pt = new PropertyTemplate();
            pt.drawBorders(new CellRangeAddress(1, 1, 0, conditionNames.length*c2Values.size()), BorderStyle.MEDIUM, BorderExtent.BOTTOM);
            pt.drawBorders(new CellRangeAddress(0, c1Names.size() + 1, 0, 0), BorderStyle.MEDIUM, BorderExtent.RIGHT);
            for(i = 0; i < conditionNames.length; i++) {
                pt.drawBorders(new CellRangeAddress(0, c1Names.size() + 1, (i+1) * c2Values.size(), (i+1) * c2Values.size()), BorderStyle.MEDIUM, BorderExtent.RIGHT);
            }
            pt.applyBorders(sheet);

            // write the file
            logger.info("Write the file");
            FileOutputStream outputStream = new FileOutputStream(output);
            outputWorkbook.write(outputStream);
            outputWorkbook.close();

        } catch (Throwable t) {
            logger.error(t.getMessage(), t);
//            t.printStackTrace();
        }
    }

    private static int getNumber(Character column) {
        return (int)column - 65;
    }

    private static String getString(XSSFRow row, Integer index) {
        try {
            if(row.getCell(index).getCellTypeEnum().equals(CellType.STRING)) {
                return row.getCell(index).getRichStringCellValue().getString();
            } else if(row.getCell(index).getCellTypeEnum().equals(CellType.NUMERIC)) {
                return ""+row.getCell(index).getNumericCellValue();
            } else if(row.getCell(index).getCellTypeEnum().equals(CellType.BOOLEAN)) {
                return ""+row.getCell(index).getBooleanCellValue();
            } else return null;
        } catch(Throwable t) {
            return null;
        }
    }
    private static Object get(XSSFRow row, Integer index) {
        try {
            if(row.getCell(index).getCellTypeEnum().equals(CellType.STRING)) {
                return row.getCell(index).getRichStringCellValue().getString();
            } else if(row.getCell(index).getCellTypeEnum().equals(CellType.NUMERIC)) {
                return row.getCell(index).getNumericCellValue();
            } else if(row.getCell(index).getCellTypeEnum().equals(CellType.BOOLEAN)) {
                return row.getCell(index).getBooleanCellValue();
            } else return null;
        } catch(Throwable t) {
            return null;
        }
    }

//    private static String getString(XSSFRow row, Integer index) {
//        try {
//            return row.getCell(index).getStringCellValue();
//        } catch(Throwable t) {
//            return null;
//        }
//    }

//    private static void write(XSSFRow row, Integer index, String value) {
//        XSSFCell cell = row.createCell(index);
//        cell.setCellValue(value);
//    }

    private static void write(XSSFRow row, Integer index, Object value) {
        XSSFCell cell = row.createCell(index);
        if(value.getClass() == Integer.class) cell.setCellValue((Integer)value);
        else if(value.getClass() == Boolean.class) cell.setCellValue((Boolean)value);
        else cell.setCellValue(value.toString());
    }

//    private static void writeArray(XSSFRow row, Integer index, ArrayList<Object> values) {
//        if(values.size() == 1) write(row, index, values.get(0));
//        else if(values.size() > 1) {
//            write(row, index, values.stream().map(v -> ""+v).collect(Collectors.joining(", ")));
//        }
//    }

    private static void writeArray(XSSFRow row, Integer index, ArrayList<XSSFCell> cells) {
        XSSFCell cell = row.createCell(index);
        if(cells.size() == 1) {
            cell.copyCellFrom(cells.get(0), policy);
        } else if(cells.size() > 1) {
            StringBuilder value = new StringBuilder();
            for(int i = 0; i < cells.size(); i++) {
                XSSFCell c = cells.get(i);
                logger.debug("ABU "+c.getCellTypeEnum());
                if(c.getCellTypeEnum().equals(CellType.STRING)) {
                    value.append(row.getCell(index).getRichStringCellValue().getString());
                } else if(c.getCellTypeEnum().equals(CellType.NUMERIC)) {
                    value.append(row.getCell(index).getNumericCellValue());
                } else if(c.getCellTypeEnum().equals(CellType.BOOLEAN)) {
                    value.append(row.getCell(index).getBooleanCellValue());
                }
                if(i != cells.size() - 1) value.append(", ");
            }
            cell.setCellValue(value.toString());
        }
    }

}
