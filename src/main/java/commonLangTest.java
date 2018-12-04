import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.util.List;

public class commonLangTest {
    public static void main(String[] args){
        readExcelFile obj = new readExcelFile();
        File file = new File("D:/5000.xls");
        List<String> contentList = obj.ReadExcelByColumn(file,1,4);
        int dis = 0;
        String s1 = contentList.get(110);
        for(String s2:contentList){
            dis = StringUtils.getLevenshteinDistance(s1, s2);
            if(dis<3){
                System.out.println(s1+":"+s2+" ("+dis+")");
            }
        }
    }


}
