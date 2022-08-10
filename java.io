1. Write a program to read text from .txt file using InputStream.

import java.io.FileInputStream;
import java.io.InputStream;

public class ReadUsingInputStream 
{
    public static void main(String[] args) 
    {
        try 
        {
            InputStream fis = new FileInputStream("os.txt");
            int i;
           
            while ((i = fis.read()) != -1) 
            {
                System.out.print((char) i);
            }
           
            fis.close();
        } 
        catch (Exception e) 
        {
            e.getStackTrace();
        }
    }
}

2. Write a program to write text to .txt file using OutputStream.                                                         

import java.io.FileOutputStream;
import java.io.OutputStream;

public class WriteUsingOutputStream 
{
    public static void main(String args[]) 
    {
        String data = "This text is written using OutputStream.";

        try 
        {
           OutputStream fos = new FileOutputStream("os.txt");
           byte[] db = data.getBytes();
           fos.write(db);
           System.out.println("Data is written to the file.");
           fos.close();
        } 
        catch(Exception e) 
        {
           e.getStackTrace();
        }
    }
}

3. Read text from a .txt file using BufferedInputStream.

import java.io.*;  
public class BufferedInputStreamExample
{    
    public static void main(String args[])
    {    
        try
        {    
            FileInputStream fin=new FileInputStream("D:\\testout.txt");    
            BufferedInputStream bin=new BufferedInputStream(fin);    
            int i;    
            while((i=bin.read())!=-1)
            {    
                System.out.print((char)i);    
            }    
            bin.close();    
            fin.close();    
        }
        catch(Exception e)
        {System.out.println(e);}    
    }    
}  

4.Write text to a .txt file using BufferedOutputStream.

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
public class OutputStream 
{
    public static void main(String[] args) throws IOException 
    {
        BufferedOutputStream bufferedOutput=new BufferedOutputStream(new FileOutputStream("yourFile.txt"));
        bufferedOutput.write("Line one".getBytes());
        bufferedOutput.write("\n".getBytes());
        bufferedOutput.write(65);
        bufferedOutput.close();
    }
}

5. Write a program to read text from .txt file using FileReader.

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
 
class ReaderFile
{
    public static void main(String[] args)
    {
        File file = new File("doc.txt");
 
        try (FileReader fr = new FileReader(file))
        {
            int content;
            while ((content = fr.read()) != -1) {
                System.out.print((char) content);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

6. Write a program to write text to .txt file using FileWriter.

import java.io.FileWriter;  
public class FileWriter1
{  
    public static void main(String args[]) 
    {    
         try
         {    
           FileWriter fw=new FileWriter("D:\\testout.txt");    
           fw.write("Welcome to javaTpoint.");    
           fw.close();    
         }
         catch(Exception e)
         {System.out.println(e);}    
         System.out.println("Success");    
     }    
}  

7. Read text from a .txt file using BufferedReader.

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
 
class reader
{
    public static void main(String[] args)
    {
        File file = new File("demo.txt");
 
        try (BufferedReader br = new BufferedReader(new FileReader(file)))
        {
            String line;
            while ((line = br.readLine()) != null)
            {
                System.out.println(line);
            }
        } 
        catch (IOException e) 
        {
            e.printStackTrace();
        }
    }
}

8. Write text to a .txt file using BufferedWriter.

import java.io.*;  
public class BufferedWriterExample 
{  
    public static void main(String[] args) throws IOException
    {     
        FileWriter writer = new FileWriter("D:\\testout.txt");  
        BufferedWriter buffer = new BufferedWriter(writer);  
        buffer.write("Welcome to javaTpoint.");  
        buffer.close();  
        System.out.println("Success");  
    }  
}  

9. Write a program to read data from properties file.

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import java.util.Set;
 
public class PropertiesCache
{
   private final Properties configProp = new Properties();
    
   private PropertiesCache()
   {
      InputStream in =this.getClass().getClassLoader().getResourceAsStream("application.properties");
      System.out.println("Reading all properties from the file");
      try 
      {
          configProp.load(in);
      } 
      catch (IOException e) 
      {
          e.printStackTrace();
      }
   }
 
   private static class LazyHolder
   {
      private static final PropertiesCache INSTANCE = new PropertiesCache();
   }
 
   public static PropertiesCache getInstance()
   {
      return LazyHolder.INSTANCE;
   }
    
   public String getProperty(String key)
   {
      return configProp.getProperty(key);
   }
    
   public Set<String> getAllPropertyNames()
   {
      return configProp.stringPropertyNames();
   }
    
   public boolean containsKey(String key)
   {
      return configProp.containsKey(key);
   }
}

10. Write a program to read data from excel.

import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  
public class ReadExcelFileDemo  
{  
    public static void main(String args[]) throws IOException  
    {  
        FileInputStream fis=new FileInputStream(new File("C:\\demo\\student.xls"));  
        HSSFWorkbook wb=new HSSFWorkbook(fis);   
        HSSFSheet sheet=wb.getSheetAt(0);  
        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
        for(Row row: sheet)     
            for(Cell cell: row)      
            {  
                switch(formulaEvaluator.evaluateInCell(cell).getCellType())  
                {  
                    case Cell.CELL_TYPE_NUMERIC:  
                    System.out.print(cell.getNumericCellValue()+ "\t\t");   
                    break;  
                    case Cell.CELL_TYPE_STRING:   
                    System.out.print(cell.getStringCellValue()+ "\t\t");  
                    break;  
                }  
            }  
        System.out.println();  
    }  
}  

11. Write a program to write data to excel.

import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
  
public class WriteDataToExcel 
{
    public static void main(String[] args) throws Exception
    {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet spreadsheet=workbook.createSheet(" Student Data ");
        XSSFRow row;
        Map<String, Object[]> studentData=new TreeMap<String, Object[]>();
        studentData.put(
            "1",
            new Object[] { "Roll No", "NAME", "Year" });
  
        studentData.put("2", new Object[] { "128", "Aditya",
                                            "2nd year" });
  
        studentData.put(
            "3",
            new Object[] { "129", "Narayana", "2nd year" });
  
        studentData.put("4", new Object[] { "130", "Mohan",
                                            "2nd year" });
  
        studentData.put("5", new Object[] { "131", "Radha",
                                            "2nd year" });
  
        studentData.put("6", new Object[] { "132", "Gopal",
                                            "2nd year" });
  
        Set<String> keyid = studentData.keySet();
        int rowid = 0;
        for (String key : keyid) 
        {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = studentData.get(key);
            int cellid = 0;
  
            for (Object obj : objectArr) 
            {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
  
        FileOutputStream out = new FileOutputStream(new File("C:/savedexcel/GFGsheet.xlsx"));
        workbook.write(out);
        out.close();
    }
}

