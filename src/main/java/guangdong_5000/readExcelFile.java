package guangdong_5000;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class readExcelFile {
    public static void main(String[] args){
        readExcelFile obj = new readExcelFile();
        File file = new File("D:/5000.xls");
//        List excelList = obj.ReadExcel(file);
//        for (int i = 0; i < excelList.size(); i++) {
//            List list = (List) excelList.get(i);
//            for (int j = 0; j < list.size(); j++) {
//                System.out.print(list.get(j));
//            }
//            System.out.println();
//        }
        List<String> list = obj.ReadExcelByColumn(file,1,4);
        for(String s:list){
            System.out.println(s);
        }
    }

   public List ReadExcel(File file) {
       InputStream is = null;
       try {
           is = new FileInputStream(file.getAbsolutePath());
           // jxl提供的Workbook类
           Workbook wb = Workbook.getWorkbook(is);
           int sheet_size = wb.getNumberOfSheets();
           for (int index = 0; index < sheet_size; index++) {
               List<List> outerList=new ArrayList<List>();
               Sheet sheet = wb.getSheet(index);
               for (int i = 0; i < sheet.getRows(); i++) {
                   List innerList=new ArrayList();
                   for (int j = 0; j < sheet.getColumns(); j++) {
                       String cellinfo = sheet.getCell(j, i).getContents();
                       if(cellinfo.isEmpty()){
                           continue;
                       }
                       innerList.add(cellinfo);
                       System.out.print(cellinfo);
                   }
                   outerList.add(i,innerList);
                   System.out.println();
               }
                   return outerList;

               }

           } catch (FileNotFoundException e) {
           e.printStackTrace();
       } catch (IOException e) {
           e.printStackTrace();
       } catch (BiffException e) {
           e.printStackTrace();
       }


       return null;
    }
    public List ReadExcelByColumn(File file,int sheetNum,int columnNum){
        InputStream is = null;
        try {
            is = new FileInputStream(file.getAbsolutePath());
            Workbook wb = Workbook.getWorkbook(is);
            int sheet_size = wb.getNumberOfSheets();
            Sheet sheet = wb.getSheet(sheetNum-1);
            List list = new ArrayList();
            Cell[] cells = sheet.getColumn(columnNum-1);
            for(int i=1;i<cells.length;i++){
                list.add(cells[i].getContents());
            }
            return list;

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

        return null;
    }
}
