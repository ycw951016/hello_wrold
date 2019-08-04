package name.ycw.helloworld;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.*;

public class JavaIOTest {
    @Test
    public void test() throws Exception {
        StringReader in = new StringReader("hahahahah");
        int c;
        while ((c=in.read()) != -1){
            System.out.println((char)c);
        }
    }
    @Test
    public void readExcel ()throws Exception{
        String s = "hello world!!!!!";
        InputStream in = new BufferedInputStream(new FileInputStream("D:\\work\\abc.xls"));
        HSSFWorkbook workbook = new HSSFWorkbook(in);
        workbook.write(new File("D:\\work\\abc1.xls"));
    }
}
