/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package intelligentanalytics;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author karthi
 */
public class IntelligentAnalytics {

    IntelligentAnalytics() {

    }

    String[] email_arry = {"m.karthikeyan@ncl.res.in", "karthincl@gmail.com", "..karthi@gmail...com"};
    String[] cc_arry = {"<MasterCard>2222 4053 4324 8877",
        "<MasterCard>2222 9909 0525 7051",
        "<MasterCard>2223 0076 4872 6984",
        "<MasterCard>2223 5771 2001 7656",
        "<MasterCard>5105 1051 0510 5100",
        "<MasterCard>5111 0100 3017 5156",
        "<MasterCard>5185 5408 1000 0019",
        "<MasterCard>5200 8282 8282 8210",
        "<MasterCard>5204 2300 8000 0017",
        "<MasterCard>5204 7400 0990 0014",
        "<MasterCard>5420 9238 7872 4339",
        "<MasterCard>5455 3307 6000 0018",
        "<MasterCard>5506 9004 9000 0436",
        "<MasterCard>5506 9004 9000 0444",
        "<MasterCard>5506 9005 1000 0234",
        "<MasterCard>5506 9208 0924 3667",
        "<MasterCard>5506 9224 0063 4930",
        "<MasterCard>5506 9274 2731 7625",
        "<MasterCard>5553 0422 4198 4105",
        "<MasterCard>5555 5537 5304 8194",
        "<MasterCard>5555 5555 5555 4444",
        "<Visa>4012 8888 8888 1881",
        "<Visa>4111 1111 1111 1111",
        "<Discover>6011 0009 9013 9424",
        "<Discover>6011 1111 1111 1117",
        "<American Express>3714 496353 98431",
        "<American Express>3782 822463 10005",
        "<Diners>3056 9309 0259 04",
        "<Diners>3852 0000 0232 37",
        "<JCB>3530 1113 3330 0000",
        "<JCB>3566 0020 2036 0505"};
    String[] phone_numbers = {
        /* Following are valid phone number examples */
        "(123)4567890", "1234567890", "+1-123-456-7890", "(123)456-7890",
        /* Following are invalid phone numbers */
        "(1234567890)", "123)4567890", "12345678901", "(1)234567890",
        "(123)-4567890", "1", "12-3456-7890", "Hello world",
        "(91)9767427981",
        "(+91)976 742 7981",
        "+91 9767427981",
        "+91-020-2590 2483",
        "+91-020-2590-2483",
        "+91-202-590-2483",
        "+91 202 590 2483"};

    String[] regex_phone = {
        //    "\\d{10}",
        //     "(?:\\d{3}-){2}\\d{4}",
        //     "\\(\\d{3}\\)\\d{3}-?\\d{4}",
        "\\s?((\\+[1-9]{1,4}[ \\-]*)|(\\([0-9]{2,3}\\)[ \\-]*)|([0-9]{2,4})[ \\-]*)*?[0-9]{3,4}?[ \\-]*[0-9]{3,4}?\\s"
    };
    String[] regex_email = {
        //      "^(.+)@(.+)$", //simple email pattern-0
        //     "^[A-Za-z0-9+_.-]+@(.+)$", //email pattern-1   //Adding Restrictions on User Name
        //    "^[a-zA-Z0-9_!#$%&amp;'*+/=?`{|}~^.-]+@[a-zA-Z0-9.-]+$", //email pattern-2 Regex for RFC-5322 Validation
        //    "^[a-zA-Z0-9_!#$&'*+/=?`{|}~^-]+(?:\\.[a-zA-Z0-9_!#$&'*+/=?`{|}~^-]+)*@[a-zA-Z0-9-]+(?:\\.[a-zA-Z0-9-]+)*$",
        "[\\w!#$%&;'*+/=?`{|}~^-]+(?:\\.[\\w!#$%&'*+/=?`{|}~^-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,6}"
    };

    String[] regex_cc = {
        "<Visa>[0-9]{4}.[0-9]{4}.[0-9]{4}.[0-9]{4}",
        "<MasterCard>[0-9]{4}.[0-9]{4}.[0-9]{4}.[0-9]{4}",
        "<Discover>[0-9]{4}.[0-9]{4}.[0-9]{4}.[0-9]{4}",
        "<American Express>[0-9]{4}.[0-9]{6}.[0-9]{5}",
        "<Diners>[0-9]{4}.[0-9]{4}.[0-9]{4}.[0-9]{2}",
        "<JCB>[0-9]{4}.[0-9]{4}.[0-9]{4}.[0-9]{4}"
    };

    String fileName = "intelligent.sqdb";
    String[] pattern = {
        regex_cc[0], regex_cc[1], regex_cc[2], regex_cc[3], regex_cc[4],
        regex_email[0],
        regex_phone[0]
    };

    public static void main(String[] args) {

        IntelligentAnalytics ia = new IntelligentAnalytics();

        String fname = "d:\\mytext.txt"; // input (with duplicates)
        String csv_fname = "d:\\mytext.csv"; //unique list (with frequency count)

        /*
        String emails = ia.EmailMatchTest();
        String phone_nos = ia.PhoneNumberMatchTest();
        String cc_nos = ia.CreditCardMatchTest();

        //   System.out.println(emails.replaceAll(";", "\n") + "\n" + phone_nos.replaceAll(";", "\n") + "\n" + cc_nos.replaceAll(";", "\n"));
        Map<String, Integer> term_freq = ia.getWordFreq(ia.getText(fname).split("\n"));
        term_freq = ia.sortByComparator(term_freq);
        WriteWordFreqCSV(term_freq, csv_fname);
         */
 /*
        String valid_ph = "";
        String invalid_ph = "";
        String valid_cc = "";
        String valid_emails = "";
        String[] csv_terms = new String[3];

        for (int i = 0; i < ia.pattern.length; i++) {
            String term_list = ia.GetTextEmails(ia.getText(fname), ia.pattern[i]);
            if (i == ia.pattern.length - 1) {
                String[] term_list_array = term_list.split("\n");
                for (int j = 0; j < term_list_array.length; j++) {
                    if (term_list_array[j].trim().length() > 12) {
                        valid_ph += (term_list_array[j].trim()) + "\n";
                    } else {
                        if (term_list_array[j].length() > 0) {
                            invalid_ph += (term_list_array[j].trim()) + "\n";
                        }
                    }
                }//for
                //   System.out.print(valid_ph);
            }//if
            else {
                if (i < 5) {
                    valid_cc += (term_list);
                } else if (i == 5) {
                    valid_emails += (term_list);
                }
            }

        }
        
        System.out.println("Credit cards:\n" + valid_cc + "\nEmails:\n" + valid_emails + "\nPhones:\n" + valid_ph);
         */
        File f = new File("case1000.xlsx");
        String out = ia.ReadExcel(f);
        System.out.println(out);

        //ArrayList<String> data=ia.extractExcelContentByColumnIndex(ci,f); 
    }

    public String ReadExcel(File file) {
        String out = "";

        try {

            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            int sheetcount = wb.getNumberOfSheets();

//            System.out.println("sheetcount: " + sheetcount);
//            for (int i = 0; i < sheetcount; i++) {
//                System.out.print(wb.getSheetName(i) + " ");
//            }
//
//            List<String> sheetNames = new ArrayList<String>();
//            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
//                sheetNames.add(wb.getSheetName(i));
//            }
            int sid = 0;
            for (int i = 0; i < sheetcount; i++) {
                System.out.println("\n****Data from : " + wb.getSheetName(i));
                sid++;
                // out += sid + "\t";
                XSSFSheet sheet = wb.getSheetAt(i);     //creating a Sheet object to retrieve object  
                Iterator<Row> itr = sheet.iterator();    //iterating over excel file  

                int rid = 0;
                while (itr.hasNext()) {
                    Row row = itr.next();
                    rid++;
                    //    out += rid + "\t";
                    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  

                    int cid = 0;
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        cid++;
                        //     out += cid + "\t";
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                                //   System.out.print(cell.getStringCellValue() + ";");
                                out += sid + "\t" + wb.getSheetName(i) + "\t" + rid + "\t" + cid + "\t" + cell.getStringCellValue() + "\n";
                                break;
                            case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                                //    System.out.print(cell.getNumericCellValue() + ";");
                                out += sid + "\t" + wb.getSheetName(i) + "\t" + rid + "\t" + cid + "\t" + cell.getNumericCellValue() + "\n";
                                break;
                            default:
                        }
                    }
                    //  System.out.println("");
                }
            }//all tbls
            System.out.println(out);

            
            
            
        } catch (Exception e) {
            e.printStackTrace();
        }

        return out;
    }

    public ArrayList<String> extractExcelContentByColumnIndex(int columnIndex, File f) {
        ArrayList<String> columndata = null;
        try {
            // File f = new File("sample.xlsx")

            /*
          
          //https://jar-download.com/artifact-search/poi-ooxml-schemas
          
          //https://poi.apache.org/
          
          //https://archive.apache.org/dist/poi/release/bin/poi-bin-5.2.3-20220909.zip
          
          
          poi-3.9.jar,
          poi-ooxml-3.9.jar,
          poi-ooxml-schemas-3.9.jar,
          xbea‌​n-2.3.0.jar,
          xmlbeans‌​-xmlpublic-2.4.0.jar‌​,
          dom4j-1.5.jar
             */
            FileInputStream ios = new FileInputStream(f);
            XSSFWorkbook workbook = new XSSFWorkbook(ios);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            columndata = new ArrayList<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (row.getRowNum() > 0) { //To filter column headings
                        if (cell.getColumnIndex() == columnIndex) {// To match column index
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    columndata.add(cell.getNumericCellValue() + "");
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    columndata.add(cell.getStringCellValue());
                                    break;
                            }
                        }
                    }
                }
            }
            ios.close();
            System.out.println(columndata);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return columndata;
    }

    private static Map<String, Integer> sortByComparator(Map<String, Integer> unsortMap) {
        // Convert Map to List
        List<Map.Entry<String, Integer>> list = new LinkedList<Map.Entry<String, Integer>>(unsortMap.entrySet());
        // Sort list with comparator, to compare the Map values
        Collections.sort(list, new Comparator<Map.Entry<String, Integer>>() {
            public int compare(Map.Entry<String, Integer> o1,
                    Map.Entry<String, Integer> o2) {
                return (o1.getValue()).compareTo(o2.getValue());
            }
        });
        // Convert sorted map back to a Map
        Map<String, Integer> sortedMap = new LinkedHashMap<String, Integer>();
        for (Iterator<Map.Entry<String, Integer>> it = list.iterator(); it.hasNext();) {
            Map.Entry<String, Integer> entry = it.next();
            sortedMap.put(entry.getKey(), entry.getValue());
        }
        return sortedMap;
    }

    public static Map<String, Integer> getWordFreq(String[] words) {
        Map<String, Integer> map = new HashMap<>();
        for (String w : words) {
            Integer n = map.get(w);
            n = (n == null) ? 1 : ++n;
            map.put(w, n);
        }
        return map;
    }

    public static void WriteWordFreqCSV(Map<String, Integer> final_map, String fname) {
        try {
            BufferedWriter bw = new BufferedWriter(new FileWriter(new File(fname)));
            for (Map.Entry<String, Integer> entry : final_map.entrySet()) {
                bw.write(entry.getKey().trim() + "\t" + entry.getValue() + "\n");
            }
            bw.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public String getText(String fname) {
        String out = "";
        try {
            BufferedReader br = new BufferedReader(new FileReader(new File(fname)));
            String s = "";
            while ((s = br.readLine()) != null) {
                String[] t = s.trim().split("\t");
                for (int i = 0; i < t.length; i++) {
                    out += t[i].trim() + "\n";
                }
            }
            br.close();
        } catch (Exception e) {
            System.out.println(e);
        }
        return out;
    }

    public String PhoneNumberMatchTest() {
        String out = "";
        List<String> ph_lst = Arrays.asList(phone_numbers);
        // System.out.println("Input phone numbers for matching:");
//        for (String e : ph_lst) {
//            System.out.println(e);
//        }
        for (int i = 0; i < regex_phone.length; i++) {
            //   System.out.println("Results for partern: " + regex_phone[i]);
            String[] cc = getMatchingIDs(ph_lst, regex_phone[i]);
            for (int j = 0; j < cc.length; j++) {
                out += (cc[j]) + ";";
            }
        }
        return out;
    }

    public String EmailMatchTest() {
        String out = "";
        List<String> email_lst = Arrays.asList(email_arry);
        // System.out.println("Input emails for matching:");
//        for (String e : email_lst) {
//            System.out.println(e);
//        }

        for (int i = 0; i < regex_email.length; i++) {
            //  System.out.println("Results for partern: " + regex_email[i]);
            String[] emails = getMatchingIDs(email_lst, regex_email[i]);
            for (int j = 0; j < emails.length; j++) {
                out += (emails[j]) + ";";;
            }
        }
        return out;

    }

    public String CreditCardMatchTest() {
        String out = "";
        List<String> cc_lst = Arrays.asList(cc_arry);
        //System.out.println("Input credit cards for matching:");
//        for (String e : cc_lst) {
//            System.out.println(e);
//        }

        for (int i = 0; i < regex_cc.length; i++) {
            //  System.out.println("Results for partern: " + regex_cc[i]);
            String[] cc = getMatchingIDs(cc_lst, regex_cc[i]);
            for (int j = 0; j < cc.length; j++) {
                out += (cc[j]) + ";";;
            }
        }
        return out;
    }

    public String[] getMatchingIDs(List<String> query, String regex) {
        String[] output = new String[query.size()];
        int cnt = 0;
        Pattern pattern = Pattern.compile(regex);
        for (String term : query) {
            try {
                Matcher m = pattern.matcher(term);
                if (m.matches()) {
                    output[cnt] = m.group();
                    //System.out.println("Found! " + output[cnt]);
                    cnt++;
                } else {
                    output[cnt] = "Not Found! or Error Pattern" + regex + "  -> " + term;
                }
            } catch (Exception e) {
                System.out.println(e + "  " + regex + "  -> " + term);
            }
        }
        //  System.out.println(cnt);
        String[] c_terms = new String[cnt];
        for (int j = 0; j < cnt; j++) {
            c_terms[j] = output[j];
        }
        return c_terms;
    }

    public void WriteCSV(String fname, String data) {
        try {
            PrintWriter pw = new PrintWriter(new FileWriter(new File(fname)));
            pw.append(data);
            pw.close();
        } catch (Exception e) {
            System.out.println(e);
        }

    }

    public void createDB(String fileName) {
        String url = "jdbc:sqlite:" + fileName;
        try {
            Connection conn = DriverManager.getConnection(url);
            if (conn != null) {
                DatabaseMetaData meta = conn.getMetaData();
            }
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
    }

    public static void createNewTable(String fileName, String sql) {
        // SQLite connection string  
        String url = "jdbc:sqlite:" + fileName;;
        try {
            Connection conn = DriverManager.getConnection(url);
            Statement stmt = conn.createStatement();
            stmt.execute(sql);
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
    }

    public String getData(String field, String table, String from, String max) {
        String out = "";
        Connection c = null;
        Statement stmt = null;
        try {
            Class.forName("org.sqlite.JDBC");
            c = DriverManager.getConnection("jdbc:sqlite:" + fileName);
            //  c.setAutoCommit(false);
            stmt = c.createStatement();
            ResultSet rs = stmt.executeQuery("SELECT " + field + " FROM " + table + " limit " + from + "," + max);
            while (rs.next()) {
                out += rs.getString(1) + "\n";
            }
            rs.close();
            stmt.close();
            c.close();
        } catch (Exception e) {
            System.out.println(e);
        }
        return out;
    }

    public String GetTextEmails(String input, String regex) {
        String out = "";
//        Scanner scanner = new Scanner(System.in);
//        String input = scanner.nextLine();
//String regex="([a-z0-9_.-]+)@([a-z0-9_.-]+[a-z])";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(input);

        while (matcher.find()) {
            out += (matcher.group() + "\n");
        }

        return out;

    }

}
