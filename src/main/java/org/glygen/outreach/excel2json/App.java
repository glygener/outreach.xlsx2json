package org.glygen.outreach.excel2json;

import java.io.IOException;

import org.glygen.outreach.excel2json.io.ExcelReader;
import org.glygen.outreach.excel2json.io.JSONWriter;
import org.glygen.outreach.excel2json.om.FileReadingResult;
import org.glygen.outreach.excel2json.om.ImportError;

public class App
{

    public static void main(String[] a_args)
    {
        String t_dataFolderPath = "./data/";
        ExcelReader t_readerExcel = new ExcelReader();
        JSONWriter t_writerJSON = new JSONWriter();
        try
        {
            FileReadingResult t_results = t_readerExcel
                    .readFile(t_dataFolderPath + "/Outreach-overview.xlsx");
            if (t_results.getErrors().size() != 0)
            {
                System.out.println("There have been errors while reading the Excel file:");
                for (ImportError t_error : t_results.getErrors())
                {
                    StringBuffer t_errorMessage = new StringBuffer();
                    t_errorMessage.append("\tSheet " + t_error.getSheet());
                    if (t_error.getRow() != null)
                    {
                        t_errorMessage.append(" - Row " + t_error.getRow().toString());
                    }
                    t_errorMessage.append(": " + t_error.getMessage());
                    System.out.println(t_errorMessage.toString());
                }
            }
            t_writerJSON.writeFile(t_dataFolderPath + "/outreach.json", t_results.getRecords());
        }
        catch (IOException e)
        {
            System.err.println("Unable to process file: " + e.getMessage());
        }
        System.out.println("Finished!");
    }
}
