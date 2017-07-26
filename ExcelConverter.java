package excelconverter;  
 import java.io.FileOutputStream;  
 import java.io.IOException;  
 import java.io.FileInputStream;  
 import java.io.InputStream;  
 import java.util.Iterator;  
 import java.util.HashMap;  
 import org.apache.poi.poifs.filesystem.POIFSFileSystem;  
 import org.apache.poi.ss.usermodel.Sheet;  
 import org.apache.poi.ss.usermodel.Workbook;  
 import org.apache.poi.hssf.usermodel.HSSFCell;  
 import org.apache.poi.hssf.usermodel.HSSFSheet;  
 import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
 import org.apache.poi.hssf.usermodel.HSSFRow;  
 import org.apache.poi.ss.usermodel.*;  
 /**  
 * A simple POI example of opening an Excel spreadsheet  
 * and writing its contents to the command line.  
 */  
 public class ExcelConverter {  
 private HashMap<Integer, String> columnToColumnNameMap;  
 public static void main( String [] args ) {  
 ExcelConverter converter = new ExcelConverter();  
 converter.convert("SourceFile - Copy");  
 }  
 public void convert(String inputFileName){  
 try{  
 InputStream inp = new FileInputStream(inputFileName + ".xls");  
 HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));  
 HSSFSheet sheet = wb.getSheetAt(0);  
 // out put file  
 Workbook workbook = new HSSFWorkbook();  
 Sheet outputSheet = workbook.createSheet("sheet1.xls");  
 HSSFRow firstRow = (HSSFRow)outputSheet.createRow(0);  
 Cell outputCell = firstRow.createCell(0);  
 outputCell.setCellValue("Student ID");  
 outputCell = firstRow.createCell(1);  
 outputCell.setCellValue("Term Start Date");  
 outputCell = firstRow.createCell(2);  
 outputCell.setCellValue("Entry Curr");  
 outputCell = firstRow.createCell(3);  
 outputCell.setCellValue("1st Major Chg Date");  
 outputCell = firstRow.createCell(4);  
 outputCell.setCellValue("Major Chg CURR");  
 /**  
 outputCell = firstRow.createCell(5);  
 outputCell.setCellValue("Expected Grad Date");  
 outputCell = firstRow.createCell(6);  
 outputCell.setCellValue("School Transfered To");**/  
 outputCell = firstRow.createCell(5);  
 outputCell.setCellValue("2nd Major Chg Date");  
 outputCell = firstRow.createCell(6);  
 outputCell.setCellValue("Major Chg CURR");  
 outputCell = firstRow.createCell(7);  
 outputCell.setCellValue("3rd Major Chg Date");  
 outputCell = firstRow.createCell(8);  
 outputCell.setCellValue("Major Chg CURR");  
 outputCell = firstRow.createCell(9);  
 outputCell.setCellValue("4th Major Chg Date");  
 outputCell = firstRow.createCell(10);  
 outputCell.setCellValue("Major Chg CURR");  
 outputCell = firstRow.createCell(11);  
 outputCell.setCellValue("5th Major Chg Date");  
 outputCell = firstRow.createCell(12);  
 outputCell.setCellValue("Major Chg CURR");  
 outputCell = firstRow.createCell(13);  
 outputCell.setCellValue("6th Major Chg Date");  
 outputCell = firstRow.createCell(14);  
 outputCell.setCellValue("Major Chg CURR");  
 outputCell = firstRow.createCell(15);  
 outputCell.setCellValue("Last Term");  
 outputCell = firstRow.createCell(16);  
 outputCell.setCellValue("Last Major CURR");  
 outputCell = firstRow.createCell(17);  
 outputCell.setCellValue("Last semester b4 1st major change");  
 outputCell = firstRow.createCell(18);  
 outputCell.setCellValue("Last Major b4 1st major change");  
 outputCell = firstRow.createCell(19);  
 outputCell.setCellValue("Last semester b4 2nd major change");  
 outputCell = firstRow.createCell(20);  
 outputCell.setCellValue("Last Major b4 2nd major change");  
 outputCell = firstRow.createCell(21);  
 outputCell.setCellValue("Last semester b4 3rd major change");  
 outputCell = firstRow.createCell(22);  
 outputCell.setCellValue("Last Major b4 3rd major change");  
 outputCell = firstRow.createCell(23);  
 outputCell.setCellValue("Last semester b4 4th major change");  
 outputCell = firstRow.createCell(24);  
 outputCell.setCellValue("Last Major b4 4th major change");  
 outputCell = firstRow.createCell(25);  
 outputCell.setCellValue("Last semester b4 5th major change");  
 outputCell = firstRow.createCell(26);  
 outputCell.setCellValue("Last Major b4 5th major change");  
 outputCell = firstRow.createCell(27);  
 outputCell.setCellValue("Last semester b4 6th major change");  
 outputCell = firstRow.createCell(28);  
 outputCell.setCellValue("Last Major b4 6th major change");  
 // Iterate over each row in the sheet  
 Iterator rows = sheet.rowIterator();  
 HSSFRow firstInputRow = sheet.getRow(0);  
 this.columnToColumnNameMap = this.createColumnIndexToColumnNameMap(firstInputRow);  
 while( rows.hasNext() ) {  
 HSSFRow row = (HSSFRow) rows.next();  
 if(row.getRowNum() == 0) {  
 firstInputRow = row;  
 continue;  
 }  
 HSSFRow outputRow = (HSSFRow)outputSheet.createRow(row.getRowNum());  
 this.produceOutputRow(row, outputRow);  
 }  
 String outputFileName = inputFileName + "-out.xls";  
 FileOutputStream output = new FileOutputStream(outputFileName);  
 workbook.write(output);  
 output.close();  
 System.out.println("finished");  
 }  
 catch ( IOException ex ) {  
 ex.printStackTrace();  
 }  
 catch (Exception e){ System.out.println(e);  
 }  
 }  
 private HashMap<Integer, String> createColumnIndexToColumnNameMap(HSSFRow firstRow){  
 HashMap<Integer, String> columnIndexToColumnNameMap = new HashMap<Integer, String> ();  
 Iterator cells = firstRow.iterator();  
 while (cells.hasNext()){  
 Cell cell = (Cell)cells.next();  
 columnIndexToColumnNameMap.put(cell.getColumnIndex(), cell.getStringCellValue());  
 }  
 return columnIndexToColumnNameMap;  
 }  
 private void produceOutputRow(HSSFRow row, HSSFRow outputRow){  
 // Iterate over each cell in the row and print out the cell's content  
 Iterator cells = row.cellIterator();  
 String firstMajor = null;  
 String secondMajor = null;  
 String thirdMajor = null;  
 String fourthMajor = null;  
 String fifthMajor = null;  
 String sixthMajor = null;  
 String LastTerm = null;  
 String LastTermString = null;  
 String LastTermBefore1stChange = null;  
 String LastMajorBefore1stChange = null;  
 String LastTermBefore2ndChange = null;  
 String LastMajorBefore2ndChange = null;  
 String LastTermBefore3rdChange = null;  
 String LastMajorBefore3rdChange = null;  
 String LastTermBefore4thChange = null;  
 String LastMajorBefore4thChange = null;  
 String LastTermBefore5thChange = null;  
 String LastMajorBefore5thChange = null;  
 String LastTermBefore6thChange = null;  
 String LastMajorBefore6thChange = null;  
 boolean LastTermHit = false;  
 int numOfMajorChange = 0;  
 boolean started = false;  
 Cell outputCell;  
 while(cells.hasNext()) {  
 HSSFCell cell = (HSSFCell) cells.next();  
 String check = columnToColumnNameMap.get(cell.getColumnIndex());  
 if(cell.getColumnIndex() >= 10 && !check.equals("TR_COL") && !check.contains("CURR2"))  
 {  
 String curricum = null;  
 double curricum2 = 0;  
 if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC ) {curricum2 = cell.getNumericCellValue(); curricum = Double.toString(curricum2);}  
 else  
 {curricum = cell.getStringCellValue(); }  
 if(curricum != null && !curricum.trim().equals("") && !check.equals("GRAD_CURR") && !curricum.equals("GRADUATE") && !curricum.equals("TRANSFER"))  
 {  
 LastTerm = curricum;  
 String Result = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermString = Result.replaceAll("_CURR1","");  
 }  
 if (check.equals("GRAD_CURR"))  
 {  
 LastTermHit = true;  
 }  
 if (LastTermHit == true)  
 {  
 outputCell = outputRow.createCell(15);  
 outputCell.setCellValue(LastTermString);  
 outputCell = outputRow.createCell(16);  
 outputCell.setCellValue(LastTerm);  
 }  
 }  
 if(cell.getColumnIndex() == 0 ){  
 // ID column.  
 outputCell = outputRow.createCell(0);  
 if(cell.getCellType() == Cell.CELL_TYPE_STRING)  
 outputCell.setCellValue(cell.getStringCellValue());  
 else {  
 Integer cellValue = new Integer((int)cell.getNumericCellValue());  
 outputCell.setCellValue(cellValue.toString());  
 }  
 }  
 if(cell.getColumnIndex() == 3 ){  
 // Entry Curriculum column.  
 outputCell = outputRow.createCell(2);  
 if(cell.getCellType() == Cell.CELL_TYPE_STRING)  
 outputCell.setCellValue(cell.getStringCellValue());  
 else {  
 Integer cellValue = (int)cell.getNumericCellValue();  
 outputCell.setCellValue(cellValue.toString());  
 }  
 firstMajor = outputCell.getStringCellValue();  
 }  
 //Start semester.  
 if(cell.getColumnIndex() >= 10 && !check.equals("GRAD_CURR") && !check.equals("TR_COL") && !check.contains("CURR2")){  
 String curricum = cell.getStringCellValue();  
 if(curricum != null && !curricum.trim().equals("") && !started){  
 outputCell = outputRow.createCell(1);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 started = true;  
 }  
 //major change date and curriculum for 1st major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 0){  
 //captures last major and last term before 1st major change  
 if(!curricum.trim().equals("") && curricum.equals(firstMajor)){  
 LastTermBefore1stChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore1stChange = LastTermBefore1stChange.replaceAll("_CURR1","");  
 LastMajorBefore1stChange = curricum;  
 }  
 //captures major and term for 1st major change  
 if(!curricum.equals(firstMajor)){  
 //outputs last major and last term before 1st major change  
 outputCell = outputRow.createCell(17);  
 outputCell.setCellValue(LastTermBefore1stChange);  
 outputCell = outputRow.createCell(18);  
 outputCell.setCellValue(LastMajorBefore1stChange);  
 //outputs major and term for 1st major change  
 outputCell = outputRow.createCell(3);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(4);  
 outputCell.setCellValue(curricum);  
 secondMajor = outputCell.getStringCellValue();  
 }  
 }  
 //major change date and curriculum for 2nd major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 1){  
 //captures last major and last term before 2nd major change  
 if(!curricum.trim().equals("") && curricum.equals(secondMajor)){  
 LastTermBefore2ndChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore2ndChange = LastTermBefore2ndChange.replaceAll("_CURR1","");  
 LastMajorBefore2ndChange = curricum;  
 }  
 //captures major and term for 2nd major change  
 if(!curricum.equals(secondMajor)&&!curricum.equals(secondMajor)){  
 //outputs last major and last term before 2nd major change  
 outputCell = outputRow.createCell(19);  
 outputCell.setCellValue(LastTermBefore2ndChange);  
 outputCell = outputRow.createCell(20);  
 outputCell.setCellValue(LastMajorBefore2ndChange);  
 //outputs major and term for 2nd major change  
 outputCell = outputRow.createCell(5);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(6);  
 outputCell.setCellValue(curricum);  
 thirdMajor = outputCell.getStringCellValue();  
 }  
 }  
 //major change date and curriculum for 3rd major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 2){  
 //captures last major and last term before 2nd major change  
 if(!curricum.trim().equals("") && curricum.equals(thirdMajor)){  
 LastTermBefore3rdChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore3rdChange = LastTermBefore3rdChange.replaceAll("_CURR1","");  
 LastMajorBefore3rdChange = curricum;  
 }  
 if(!curricum.equals(secondMajor)&&!curricum.equals(firstMajor)&&!curricum.equals(thirdMajor)){  
 //outputs last major and last term before 3rd major change  
 outputCell = outputRow.createCell(21);  
 outputCell.setCellValue(LastTermBefore3rdChange);  
 outputCell = outputRow.createCell(22);  
 outputCell.setCellValue(LastMajorBefore3rdChange);  
 outputCell = outputRow.createCell(7);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(8);  
 outputCell.setCellValue(curricum);  
 fourthMajor = outputCell.getStringCellValue();  
 }  
 }  
 //major change date and curriculum for 4th major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 3){  
 //captures last major and last term before 4th major change  
 if(!curricum.trim().equals("") && curricum.equals(fourthMajor)){  
 LastTermBefore4thChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore4thChange = LastTermBefore4thChange.replaceAll("_CURR1","");  
 LastMajorBefore4thChange = curricum;  
 }  
 if(!curricum.equals(secondMajor)&&!curricum.equals(fourthMajor)&&!curricum.equals(firstMajor)&&!curricum.equals(thirdMajor)){  
 //outputs last major and last term before 4th major change  
 outputCell = outputRow.createCell(23);  
 outputCell.setCellValue(LastTermBefore4thChange);  
 outputCell = outputRow.createCell(24);  
 outputCell.setCellValue(LastMajorBefore4thChange);  
 outputCell = outputRow.createCell(9);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(10);  
 outputCell.setCellValue(curricum);  
 fifthMajor = outputCell.getStringCellValue();  
 }  
 }  
 //major change date and curriculum for 5th major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 4){  
 //captures last major and last term before 5th major change  
 if(!curricum.trim().equals("") && curricum.equals(fifthMajor)){  
 LastTermBefore5thChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore5thChange = LastTermBefore5thChange.replaceAll("_CURR1","");  
 LastMajorBefore5thChange = curricum;  
 }  
 if(!curricum.equals(secondMajor)&&!curricum.equals(fifthMajor)&&!curricum.equals(fourthMajor)&&!curricum.equals(firstMajor)&&!curricum.equals(thirdMajor)){  
 //outputs last major and last term before 5th major change  
 outputCell = outputRow.createCell(25);  
 outputCell.setCellValue(LastTermBefore5thChange);  
 outputCell = outputRow.createCell(26);  
 outputCell.setCellValue(LastMajorBefore5thChange);  
 outputCell = outputRow.createCell(11);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(12);  
 outputCell.setCellValue(curricum);  
 sixthMajor = outputCell.getStringCellValue();  
 }  
 }  
 //major change date and curriculum for 6th major change  
 if(curricum != null && !curricum.trim().equals("") && numOfMajorChange == 5){  
 //captures last major and last term before 6th major change  
 if(!curricum.trim().equals("") && curricum.equals(sixthMajor)){  
 LastTermBefore6thChange = columnToColumnNameMap.get(cell.getColumnIndex());  
 LastTermBefore6thChange = LastTermBefore6thChange.replaceAll("_CURR1","");  
 LastMajorBefore6thChange = curricum;  
 }  
 if(!curricum.equals(secondMajor)&&!curricum.equals(sixthMajor)&&!curricum.equals(fifthMajor)&&!curricum.equals(fourthMajor)&&!curricum.equals(firstMajor)&&!curricum.equals(thirdMajor)){  
 //outputs last major and last term before 6th major change  
 outputCell = outputRow.createCell(27);  
 outputCell.setCellValue(LastTermBefore6thChange);  
 outputCell = outputRow.createCell(28);  
 outputCell.setCellValue(LastMajorBefore6thChange);  
 outputCell = outputRow.createCell(13);  
 String term = columnToColumnNameMap.get(cell.getColumnIndex());  
 String Result = term.replaceAll("_CURR1","");  
 outputCell.setCellValue(Result);  
 numOfMajorChange++;  
 outputCell = outputRow.createCell(14);  
 outputCell.setCellValue(curricum);  
 }  
 }  
 }  
 }  
 }}