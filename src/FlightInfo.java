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

    public static void main(String[] args){
        FlightInfo info = new FlightInfo();
        try {
            info.printInfo();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 用一个docx文档作为模板，读取data下的.xls文件，然后替换其中的内容，再写入目标文档中。
     * @throws Exception
     */

    public void printInfo() throws Exception {

        short totalCell = 7;//机组人数在第8格 - 1

        short airplaneCell = 5;//飞机编号在第6格 - 1

        short dateCell = 8;//日期在第9格 - 1

        short crewCell = 2;//全体成员名字在第3格 - 1

        short titleCell = 7;//全体成员职称在第8格 - 1

        short flightCell = 1;//航班号在第2格 - 1

        short arrivalCell = 4;//航班号在第5格 - 1

        int infoRow = 4;//关键信息栏行数5 - 1

        int crewstart = 6;//机组信息栏行数7 - 1

        int pilotNum = 0;

        String path = "./data/";

        File file = new File(path);

        File[] files = file.listFiles();

        for(int i=0; i<files.length; i++) {

            if (files[i].isFile()) {

                String filename = files[i].getName();

                //System.out.println(filename);

                //String filename = name.substring(name.indexOf(" ") + 1, name.indexOf(".XLS"));

                InputStream xis = new FileInputStream(path + filename);
                HSSFWorkbook wbs = new HSSFWorkbook(xis);

                HSSFSheet sheet = wbs.getSheetAt(0);

                int totalNum = Integer.parseInt(sheet.getRow(infoRow).getCell(totalCell).toString());

                String[] aimArray = sheet.getRow(infoRow).getCell(arrivalCell).toString().split("\n");

                String aim = aimArray[1].substring(aimArray[1].lastIndexOf("(") + 1, aimArray[1].lastIndexOf(")"));

                String arrival = "";

                switch (aim) {
                    case "BKK":
                        arrival += "BANGKOK,THAILAND";
                        break;
                    case "CJU":
                        arrival += "CHEJU,KOREA";
                        break;
                    case "CXR":
                        arrival += "NHA TRANG,VIETNAM";
                        break;
                    case "HKG":
                        arrival += "HONGKONG";
                        break;
                    case "HKT":
                        arrival += "PHUKET,THAILAND";
                        break;
                    case "ICN":
                        arrival += "SEOUL,KOREA";
                        break;
                    case "KBV":
                        arrival += "KRABI,THAILAND";
                        break;
                    case "KIX":
                        arrival += "OSAKA,JAPAN";
                        break;
                    case "NRT":
                        arrival += "NARITA,JAPAN";
                        break;
                    case "PRG":
                        arrival += "PRAGUE,CZECH";
                        break;
                    case "SIN":
                        arrival += "SINGAPORE";
                        break;
                    case "SVO":
                        arrival += "MOSCOW,RUSSIA";
                        break;
                    case "TSA":
                        arrival += "TAIPEISONGSHAN";
                        break;
                    case "DXB":
                        arrival += "DUBAI,UNITEDARABEMIRATES";
                        break;
                    case "MEL":
                        arrival += "MELBOURNE,AUSTRIALIA";
                        break;
                    case "CNX":
                        arrival += "CHIANGMAI,THAILAND";
                        break;
                    case "SPN":
                        arrival += "SAIPAN,AMERICA";
                        break;
                    case "LAX":
                        arrival += "LOSANGELES,AMERICA";
                        break;
                    case "YVR":
                        arrival += "VANCOUVER,CANADA";
                        break;
                    case "KTM":
                        arrival += "KATHMANDU,NEPAL";
                        break;
                }

                String flight = sheet.getRow(infoRow).getCell(flightCell).toString().toUpperCase();

                if (flight.contains("/")){
                    flight = flight.substring(0,flight.indexOf("/"));
                }

                String airplane = sheet.getRow(infoRow).getCell(airplaneCell).toString().replace("\n", "/");

                String date = sheet.getRow(infoRow).getCell(dateCell).toString();

                StringBuilder crewBuilder = new StringBuilder();

                String titles = "";

                int crewend = crewstart + totalNum;

                for (int j = crewstart; j <  crewend; j++) {

                    HSSFRow row = sheet.getRow(j);
                    HSSFCell theCrewCell = row.getCell(crewCell);
                    String crew = theCrewCell.toString().toUpperCase().replace("/", " ");
                    crewBuilder.append(crew.concat("\n"));

                    HSSFCell theTitleCell = row.getCell(titleCell);
                    String title = theTitleCell.toString().toUpperCase();

                    //System.out.println(title);

                    if (title.contains("CAPT")) {

                        pilotNum++;

                        if (!titles.contains("CAP")) {
                            titles += "CAP\n";
                        } else {
                            titles += "\n";
                        }

                    } else if (title.contains("F/O")) {

                        pilotNum++;

                        if (!titles.contains("F/O")) {
                            titles += "F/O\n";
                        } else {
                            titles += "\n";
                        }

                    }  else if (title.contains("F/C")) {

                        pilotNum++;

                        if (!titles.contains("F/C")) {
                            titles += "F/C\n";
                        } else {
                            titles += "\n";
                        }

                    } else if (title.contains("PS")) {

                        if (!titles.contains("F/P")) {
                            titles += "F/P\n";
                        } else {
                            titles += "\n";
                        }

                    } else if (title.contains("D/H")) {

                        if (!titles.contains("D/H")) {
                            titles += "D/H\n";
                        } else {
                            titles += "\n";
                        }

                    } else if (title.contains("SECU")) {

                        if (!titles.contains("SEC")) {
                            titles += "SEC\n";
                        } else {
                            titles += "\n";
                        }

                    }else {

                        if (!titles.contains("STW")) {
                            titles += "STW\n";
                        } else {
                            titles += "\n";
                        }

                    }

                }

                String count = String.valueOf(pilotNum).concat("/").concat(String.valueOf(totalNum - pilotNum));

                pilotNum = 0;

                String crews = crewBuilder.toString();

                try {
                    xis.close();

                } catch (IOException e) {
                    e.printStackTrace();
                }

                this.writeToDoc(airplane, flight, date, aim, arrival, count, titles, crews);

            }else {
                System.out.println("请将.xls拷贝到data目录");
            }
        }

    }

    /**
     * 输出到docx文档
     * @param airplane,date,ap,count,titles,crews
     */
    private void writeToDoc(String airplane, String flight, String date,  String aim, String arrival, String count, String titles, String crews) {

        String filename = aim + date.substring(date.indexOf(".")).replace(".","");

        String stop = "";

        String stopover = "";

        switch (flight) {
            case "3U8645":
                stop += "CAN";
                stopover += "GUANGZHOU";
                break;
            case "3U8647":
                stop += "PVG";
                stopover += "SHANGHAI";
                break;
            case "3U8631":
                stop += "HGH";
                stopover += "HANGZHOU";
                break;
            case "3U8699":
                stop += "TNA";
                stopover += "JINAN";
                break;
            case "3U8579":
                stop += "SHE";
                stopover += "SHENYANG";
                break;
            case "3U8501":
                stop += "CGO";
                stopover += "ZHENGZHOU";
                break;
            case "3U8719":
                stop += "LXA";
                stopover += "LHASHA";
                break;
            case "3U8945":
                stop += "INC";
                stopover += "YINCHUAN";
                break;
        }

        //模板变量初始化
        InputStream is = null;
        Map<String, Object> params = new HashMap<String, Object>();

        if (stop == ""){
            //普通模板
            try {
                is = new FileInputStream("./template/template.docx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }else {
            //经停模板
            try {
                is = new FileInputStream("./template/templateStop.docx");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            params.put("stop", stop);
            params.put("stopover", stopover);
        }

        //模板变量
        params.put("airplane", airplane);
        params.put("flight", flight);
        params.put("date", date);
        params.put("aim", aim);
        params.put("arrival", arrival);
        params.put("count", count);
        params.put("titles", titles);
        params.put("crews", crews);

        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //替换段落里面的变量
        //this.replaceInDoc(doc, params);
        //替换表格里面的变量
        this.replaceInTable(doc, params);
        OutputStream os = null;
        try {
            os = new FileOutputStream("./result/" + filename + ".docx");
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
        System.out.println(flight + " finished~");
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
                //System.out.println(runText+" |");
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

}
