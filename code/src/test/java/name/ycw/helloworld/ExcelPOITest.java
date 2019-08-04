package name.ycw.helloworld;

import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelPOITest {

    @Test
    public void excel(){

        try (OutputStream fileOut = new FileOutputStream("D:\\work\\workbook1.xls")) {

            Workbook wb = WorkbookFactory.create(new File("D:\\work\\abc.xls"));

            Map<String, List<Map<String, Object>>> data = new HashMap<>();
            Map<String, Object> map = new HashMap<>();
//            map.put("SBYXXMZ","是不允许心目中");
            List<Map<String, Object>> list = new ArrayList<>();
            list.add(map);
            data.put("LIST1",list);

            Map<String, Object> map2 = new HashMap<>();
            map2.put("ABC","ABC");
            List<Map<String, Object>> list2 = new ArrayList<>();
//            list2.add(map2);
            data.put("LIST2",list2);

            Map<String, Object> map3 = new HashMap<>();
            map3.put("ZKSBBH","ZKSBBH");
            map3.put("SJ","SJ");
            map3.put("SBYXXMDM","SBYXXMDM");
            map3.put("BB",999L);
            map3.put("CC",110);
            map3.put("DD",new Date());
//            map3.put("FF","FF");
            map3.put("AA","AA");
            Map<String, Object> map4= new HashMap<>();
            map4.put("ZKSBBH","goodgood");
            map4.put("SJ","SJ");
//            map4.put("SBYXXMDM","SBYXXMDM4");
            map4.put("BB","BB4");
            map4.put("CC","CC4");
            map4.put("DD","DD4");
            map4.put("FF","FF4");
            map4.put("AA","AA4");
            List<Map<String, Object>> list3 = new ArrayList<>();
            list3.add(map3);
            list3.add(map4);
            data.put("LIST3",list3);

            fillTemplate(wb , data);


            wb.write(fileOut);
            
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 替换模板数据
     *
     * 单行：
     * ${LIST[0].FIELD} 替换数据为data.get("LIST").get(0).get("FIELD")
     * 多行：
     * #{LIST, colCount, rowCount, TRUE} 和%{}组合使用。替换此cell往后的colCount个cell，并向下插入行，总行数为rowCount
     * rowCount为0时，表示插入全部数据
     * %{FIELD} 和#{}组合使用。
     * @param wb Workbook对象
     * @param data 数据
     */
    private void fillTemplate(Workbook wb, Map<String, List<Map<String, Object>>> data) {

        DataFormatter formatter = new DataFormatter();
        CreationHelper creationHelper = wb.getCreationHelper();
        //日期单元格样式
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));

        List<Cell> multiRow = new ArrayList<>();

        //第一次遍历，替换单行数据，多行#{}所在cell放入multiRow
        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = formatter.formatCellValue(cell);
                    // ${LIST[0].FIELD}替换为data中的数据
                    if (text.matches("^\\$\\{.*}$")){
                        String[] split = text.substring(2, text.length() - 1).split("\\.");
                        String listName = split[0].replaceAll("\\[\\d*]$","").toUpperCase();
                        String fieldName = split[1].toUpperCase();

                        try {
                            Object d = data.get(listName).get(0).get(fieldName);
                            if(d instanceof Date){
                                cell.setCellStyle(dateStyle);
                                cell.setCellValue((Date) d);
                            }else {
                                cell.setCellValue(d.toString());
                            }
                        } catch (NullPointerException e) {
                            System.err.println("匹配 " + text + " 对应数据失败");
                            cell.setCellValue("");
                        } catch (IndexOutOfBoundsException e){
                            System.err.println("匹配 list:" + listName + " 对应数据失败");
                            cell.setCellValue("");
                        }
                        continue;
                    }
                    if (text.matches("^#\\{.*}$")){
                        multiRow.add(cell);
                    }
                }
            }
        }

        // 多行 #{LIST, colCount, rowCount, TRUE}%{}替换
        Matcher matcher = Pattern.compile("#\\{.*}").matcher("");
        Matcher matcher1 = Pattern.compile("%\\{\\w*}").matcher("");
        for (Cell cell : multiRow) {
            String stringCellValue = formatter.formatCellValue(cell);
            matcher.reset(stringCellValue);
            if(matcher.find()){
                // 以逗号分割#{}中的参数
                String[] split = matcher.group().substring(2, matcher.group().length() - 1).split(",");
                String listName = split[0].toUpperCase();
                int colCount = Integer.parseInt(split[1]);
//                int rowCount = Integer.parseInt(split[2]);//TODO: 行数参数未实现

                List<Map<String, Object>> dataList = data.get(listName);
                if(dataList == null){
                    System.err.println("匹配 list:" + listName + " 对应数据失败");
                    continue;
                }

                Sheet sheet = cell.getSheet();
//                下移
                if (cell.getRowIndex() != sheet.getLastRowNum())
                    sheet.shiftRows(cell.getRowIndex() + 1, sheet.getLastRowNum(),dataList.size() - 1);
//                替换数据
                int rowIndex = cell.getRowIndex();
                int colIndex = cell.getColumnIndex();
                for (int i = 0; i < colCount; i++) {
                    matcher1.reset(formatter.formatCellValue(sheet.getRow(rowIndex).getCell(colIndex + i)));
                    if (matcher1.find()){
                        String key = matcher1.group().substring(2, matcher1.group().length() - 1).toUpperCase();
                        for (int j = 0; j < dataList.size(); j++) {
                            Row row = sheet.getRow(rowIndex + j);
                            if (row == null)
                                row = sheet.createRow(rowIndex + j);
                            Cell cell1 = row.getCell(colIndex + i);
                            if(cell1 == null)
                                cell1 = row.createCell(colIndex + i);

                            try {
                                Object d = dataList.get(j).get(key);
                                if(d instanceof Date){
                                    cell1.setCellStyle(dateStyle);
                                    cell1.setCellValue((Date) d);
                                }else {
                                    cell1.setCellValue(d.toString());
                                }
                            } catch (NullPointerException e) {
                                System.err.println("匹配 " + listName + "." + key +
                                        "(row:" + j + ") 对应数据失败");
                                cell1.setCellValue("");
                            }
                        }
                    }
                }

            }
        }
    }
}
