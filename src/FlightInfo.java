import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.poi.xssf.usermodel.XSSFFont.DEFAULT_FONT_SIZE;

/**
 * Created by Mac on 2017/1/25.
 */
public class FlightInfo {



    /**
     * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
     * @throws Exception
     */

    public void printInfo() throws Exception {

        InputStream xis = new FileInputStream("./data/data.xls");
        HSSFWorkbook wbs = new HSSFWorkbook(xis);

        HSSFSheet sheet = wbs.getSheetAt(0);

        int infoRow = 4;//关键信息栏行数 - 1

        short totalCell = 7;//机组人数在第几格 - 1

        short airplaneCell = 5;//飞机编号在第几格 - 1

        short dateCell = 8;//日期在第几格 - 1

        short crewCell = 2;//全体成员在第几格 - 1

        short titleCell = 7;//全体成员职称8在第几格 - 1

        int crewstart = 6;//机组信息7栏行数 - 1

        int crewend = 1;//机组信息栏2后面行数 - 1

        int pilotNum = 0;

        int totalNum = Integer.parseInt(sheet.getRow(infoRow).getCell(totalCell).toString());

        String airplane =  sheet.getRow(infoRow).getCell(airplaneCell).toString().replace("\n","/");

        String date =  sheet.getRow(infoRow).getCell(dateCell).toString();

        StringBuilder crewBuilder = new StringBuilder();

        String titles = "";

        for (int i = crewstart; i < sheet.getLastRowNum() - crewend; i++ ) {
            HSSFRow row = sheet.getRow(i);

            HSSFCell theCrewCell = row.getCell(crewCell);
            String crew = theCrewCell.toString().toUpperCase().replace("/"," ");
            crewBuilder.append(crew.concat("\n"));

            HSSFCell theTitleCell = row.getCell(titleCell);
            String title = theTitleCell.toString().toUpperCase();

            if ( title.contains("CAPT") ){

                pilotNum ++;

                if( !titles.contains("CAP")){
                    titles += "CAP\n";
                }else{
                    titles += "\n" ;
                }

            }else if( title.contains("F/O") ){

                pilotNum ++;

                if( !titles.contains("F/O")){
                    titles += "F/O\n";
                }else{
                    titles += "\n";
                }

            }else if(title.contains("F/P") ){

                if( !titles.contains("F/P")){
                    titles += "F/P\n";
                }else{
                    titles += "\n";
                }

            }else if(title.contains("STW") ){

                if( !titles.contains("STW")){
                    titles += "STW\n";
                }else{
                    titles += "\n";
                }

            }else {
                if( !titles.contains("SEC")){
                    titles += "SEC\n";
                }else{
                    titles += "\n";
                }
            }

        }

        String count = String.valueOf(pilotNum).concat("/").concat(String.valueOf(totalNum - pilotNum));

        String crews =  crewBuilder.toString();

        //System.out.println(titles);

        try {
            xis.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

        this.writeToDoc(airplane,date,count,titles,crews);

    }

    /**
     * 输出到docx文档
     * @param airplane,date,count,titles,crews
     */
    private void writeToDoc(String airplane, String date, String count, String titles,String crews) {
        //模板变量初始化
        Map<String, Object> params = new HashMap<String, Object>();
        params.put("airplane", airplane);
        params.put("date", date);
        params.put("count", count);
        params.put("titles", titles);
        params.put("crews", crews);

        InputStream is = null;
        try {
            is = new FileInputStream("./template/template.docx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //替换段落里面的变量
        this.replaceInDoc(doc, params);
        //替换表格里面的变量
        this.replaceInTable(doc, params);
        OutputStream os = null;
        try {
            os = new FileOutputStream("./result/output.docx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            doc.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            os.close();
            is.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.print("succeed!!");
    }

    /**
     * 正则匹配字符串
     * @param str
     * @return
     */
    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 替换段落里面的变量
     * @param doc 要替换的文档
     * @param params 参数
     */
    private void replaceInDoc(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    /**
     * 替换段落里面的变量
     * @param para 要替换的段落
     * @param params 参数
     */
    private void replaceInPara(XWPFParagraph para, Map<String, Object> params) {
        List<XWPFRun> runs;
        Matcher matcher;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            for (int i=0; i<runs.size(); i++) {
                XWPFRun run = runs.get(i);
                int runFontSize = run.getFontSize();
                String runText = run.toString();
                String runColor = run.getColor();
                String runFontFamily = run.getFontFamily();
                Boolean runItalic = run.isItalic();
                Boolean runBold = run.isBold();
                UnderlinePatterns runUnderline = run.getUnderline();
                matcher = this.matcher(runText);
                if (matcher.find()) {
                    while ((matcher = this.matcher(runText)).find()) {
                        runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
                    }
                    //直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                    //所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                    para.removeRun(i);

                    String[] content = runText.split("\n");

                    for (int j = content.length - 1; j >= 0; j--){

                        XWPFRun newRun = para.insertNewRun(i);

                        // Copy text
                        newRun.setText(content[j]);

                        // Apply the same style
                        newRun.setFontSize( ( runFontSize == -1) ? DEFAULT_FONT_SIZE : runFontSize  );
                        newRun.setFontFamily( runFontFamily );
                        newRun.setBold( runBold );
                        newRun.setUnderline( runUnderline );
                        newRun.setItalic( runItalic );
                        newRun.setColor( runColor );

                        if (content.length > 1){
                            newRun.addBreak();
                        }

                    }

                }
            }
        }
    }

    /**
     * 替换表格里面的变量
     * @param doc 要替换的文档
     * @param params 参数
     */
    private void replaceInTable(XWPFDocument doc, Map<String, Object> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        this.replaceInPara(para, params);
                    }
                }
            }
        }
    }

    public static void main(String[] args){
        FlightInfo info = new FlightInfo();
        try {
            info.printInfo();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
