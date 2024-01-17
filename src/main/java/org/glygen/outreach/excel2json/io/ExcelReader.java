package org.glygen.outreach.excel2json.io;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.glygen.outreach.excel2json.om.FileMetadata;
import org.glygen.outreach.excel2json.om.FileReadingResult;
import org.glygen.outreach.excel2json.om.ImportError;
import org.glygen.outreach.excel2json.om.OutreachRecord;

public class ExcelReader
{
    private static String SHEET_PAPER = "Paper";
    private static String SHEET_MEETING_TALK = "Meeting Talk";
    private static String SHEET_POSTER = "Poster";
    private static String SHEET_DEMO = "Demo";
    private static String SHEET_ONLINE_TALK = "Online Talk";
    private static String SHEET_WORKSHOP = "Workshop";
    private static String SHEET_TRAINING = "Training";

    private HashMap<Integer, String> m_headersPaper = new HashMap<Integer, String>();
    private HashMap<Integer, String> m_headersMeetingTalkPosterDemo = new HashMap<Integer, String>();
    private HashMap<Integer, String> m_headersOnlineTalk = new HashMap<Integer, String>();
    private HashMap<Integer, String> m_headersTraining = new HashMap<Integer, String>();
    private HashMap<Integer, String> m_headersWorkshop = new HashMap<Integer, String>();

    private FileReadingResult m_result = null;
    private FileMetadata m_file1 = null;
    private FileMetadata m_file2 = null;
    private FileMetadata m_file3 = null;

    public ExcelReader()
    {
        super();
        // paper headers
        Integer t_counter = 0;
        this.m_headersPaper.put(t_counter++, "Title");
        this.m_headersPaper.put(t_counter++, "Autors");
        this.m_headersPaper.put(t_counter++, "Journal");
        this.m_headersPaper.put(t_counter++, "Date");
        this.m_headersPaper.put(t_counter++, "Issue");
        this.m_headersPaper.put(t_counter++, "Pages");
        this.m_headersPaper.put(t_counter++, "PMID");
        this.m_headersPaper.put(t_counter++, "DOID");
        this.m_headersPaper.put(t_counter++, "Cite");
        this.m_headersPaper.put(t_counter++, "Funding 1");
        this.m_headersPaper.put(t_counter++, "Funding 2");
        this.m_headersPaper.put(t_counter++, "Funding 3");
        this.m_headersPaper.put(t_counter++, "File 1 - Type");
        this.m_headersPaper.put(t_counter++, "File 1 - Format");
        this.m_headersPaper.put(t_counter++, "File 1 - URL");
        this.m_headersPaper.put(t_counter++, "File 2 - Type");
        this.m_headersPaper.put(t_counter++, "File 2 - Format");
        this.m_headersPaper.put(t_counter++, "File 2 - URL");
        this.m_headersPaper.put(t_counter++, "File 3 - Type");
        this.m_headersPaper.put(t_counter++, "File 3 - Format");
        this.m_headersPaper.put(t_counter++, "File 3 - URL");

        t_counter = 0;
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Title");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Presenter");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Date");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Meeting");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Location");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Meeting URL");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Funding 1");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Funding 2");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "Funding 3");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 1 - Type");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 1 - Format");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 1 - URL");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 2 - Type");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 2 - Format");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 2 - URL");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 3 - Type");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 3 - Format");
        this.m_headersMeetingTalkPosterDemo.put(t_counter++, "File 3 - URL");

        t_counter = 0;
        this.m_headersOnlineTalk.put(t_counter++, "Title");
        this.m_headersOnlineTalk.put(t_counter++, "Presenter");
        this.m_headersOnlineTalk.put(t_counter++, "Date");
        this.m_headersOnlineTalk.put(t_counter++, "Meeting");
        this.m_headersOnlineTalk.put(t_counter++, "Meeting URL");
        this.m_headersOnlineTalk.put(t_counter++, "Funding 1");
        this.m_headersOnlineTalk.put(t_counter++, "Funding 2");
        this.m_headersOnlineTalk.put(t_counter++, "Funding 3");
        this.m_headersOnlineTalk.put(t_counter++, "File 1 - Type");
        this.m_headersOnlineTalk.put(t_counter++, "File 1 - Format");
        this.m_headersOnlineTalk.put(t_counter++, "File 1 - URL");
        this.m_headersOnlineTalk.put(t_counter++, "File 2 - Type");
        this.m_headersOnlineTalk.put(t_counter++, "File 2 - Format");
        this.m_headersOnlineTalk.put(t_counter++, "File 2 - URL");
        this.m_headersOnlineTalk.put(t_counter++, "File 3 - Type");
        this.m_headersOnlineTalk.put(t_counter++, "File 3 - Format");
        this.m_headersOnlineTalk.put(t_counter++, "File 3 - URL");

        t_counter = 0;
        this.m_headersTraining.put(t_counter++, "Title");
        this.m_headersTraining.put(t_counter++, "Presenter");
        this.m_headersTraining.put(t_counter++, "Date");
        this.m_headersTraining.put(t_counter++, "Meeting");
        this.m_headersTraining.put(t_counter++, "Location");
        this.m_headersTraining.put(t_counter++, "Meeting URL");
        this.m_headersTraining.put(t_counter++, "Participants");
        this.m_headersTraining.put(t_counter++, "Funding 1");
        this.m_headersTraining.put(t_counter++, "Funding 2");
        this.m_headersTraining.put(t_counter++, "Funding 3");
        this.m_headersTraining.put(t_counter++, "File 1 - Type");
        this.m_headersTraining.put(t_counter++, "File 1 - Format");
        this.m_headersTraining.put(t_counter++, "File 1 - URL");
        this.m_headersTraining.put(t_counter++, "File 2 - Type");
        this.m_headersTraining.put(t_counter++, "File 2 - Format");
        this.m_headersTraining.put(t_counter++, "File 2 - URL");
        this.m_headersTraining.put(t_counter++, "File 3 - Type");
        this.m_headersTraining.put(t_counter++, "File 3 - Format");
        this.m_headersTraining.put(t_counter++, "File 3 - URL");

        t_counter = 0;
        this.m_headersWorkshop.put(t_counter++, "Workshop title");
        this.m_headersWorkshop.put(t_counter++, "Workshop URL");
        this.m_headersWorkshop.put(t_counter++, "Organizer");
        this.m_headersWorkshop.put(t_counter++, "Date");
        this.m_headersWorkshop.put(t_counter++, "Meeting");
        this.m_headersWorkshop.put(t_counter++, "Location");
        this.m_headersWorkshop.put(t_counter++, "Meeting URL");
        this.m_headersWorkshop.put(t_counter++, "Participants");
        this.m_headersWorkshop.put(t_counter++, "Funding 1");
        this.m_headersWorkshop.put(t_counter++, "Funding 2");
        this.m_headersWorkshop.put(t_counter++, "Funding 3");
        this.m_headersWorkshop.put(t_counter++, "File 1 - Type");
        this.m_headersWorkshop.put(t_counter++, "File 1 - Format");
        this.m_headersWorkshop.put(t_counter++, "File 1 - URL");
        this.m_headersWorkshop.put(t_counter++, "File 2 - Type");
        this.m_headersWorkshop.put(t_counter++, "File 2 - Format");
        this.m_headersWorkshop.put(t_counter++, "File 2 - URL");
        this.m_headersWorkshop.put(t_counter++, "File 3 - Type");
        this.m_headersWorkshop.put(t_counter++, "File 3 - Format");
        this.m_headersWorkshop.put(t_counter++, "File 3 - URL");

    }

    public FileReadingResult readFile(String t_fileNamePath) throws IOException
    {
        FileInputStream t_file = new FileInputStream(new File(t_fileNamePath));
        this.m_result = new FileReadingResult();

        HashMap<String, XSSFSheet> t_sheetsMap = new HashMap<String, XSSFSheet>();
        // Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook t_workbook = new XSSFWorkbook(t_file);
        t_file.close();
        // Get first/desired sheet from the workbook
        for (int i = 0; i < t_workbook.getNumberOfSheets(); i++)
        {
            XSSFSheet t_sheet = t_workbook.getSheetAt(i);
            // Process your sheet here.
            if (t_sheet.getSheetName().equals(SHEET_PAPER))
            {
                t_sheetsMap.put(SHEET_PAPER, t_sheet);
                this.readSheet(t_sheet, this.m_headersPaper, SHEET_PAPER);
            }
            else if (t_sheet.getSheetName().equals(SHEET_MEETING_TALK))
            {
                t_sheetsMap.put(SHEET_MEETING_TALK, t_sheet);
                this.readSheet(t_sheet, this.m_headersMeetingTalkPosterDemo, SHEET_MEETING_TALK);
            }
            else if (t_sheet.getSheetName().equals(SHEET_POSTER))
            {
                t_sheetsMap.put(SHEET_POSTER, t_sheet);
                this.readSheet(t_sheet, this.m_headersMeetingTalkPosterDemo, SHEET_POSTER);
            }
            else if (t_sheet.getSheetName().equals(SHEET_DEMO))
            {
                t_sheetsMap.put(SHEET_DEMO, t_sheet);
                this.readSheet(t_sheet, this.m_headersMeetingTalkPosterDemo, SHEET_DEMO);
            }
            else if (t_sheet.getSheetName().equals(SHEET_ONLINE_TALK))
            {
                t_sheetsMap.put(SHEET_ONLINE_TALK, t_sheet);
                this.readSheet(t_sheet, this.m_headersOnlineTalk, SHEET_ONLINE_TALK);
            }
            else if (t_sheet.getSheetName().equals(SHEET_WORKSHOP))
            {
                t_sheetsMap.put(SHEET_WORKSHOP, t_sheet);
                this.readSheet(t_sheet, this.m_headersWorkshop, SHEET_WORKSHOP);
            }
            else if (t_sheet.getSheetName().equals(SHEET_TRAINING))
            {
                t_sheetsMap.put(SHEET_TRAINING, t_sheet);
                this.readSheet(t_sheet, this.m_headersTraining, SHEET_TRAINING);
            }
            else
            {
                this.m_result.getErrors()
                        .add(new ImportError(t_sheet.getSheetName(), "Found unknown sheet name"));
            }
        }
        this.checkCompleteness(t_sheetsMap);
        t_workbook.close();
        return this.m_result;
    }

    private void checkCompleteness(HashMap<String, XSSFSheet> a_sheetsMap)
    {
        if (a_sheetsMap.get(SHEET_PAPER) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_MEETING_TALK) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_MEETING_TALK, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_POSTER) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_POSTER, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_DEMO) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_DEMO, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_ONLINE_TALK) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_WORKSHOP) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_WORKSHOP, "Missing sheet"));
        }
        if (a_sheetsMap.get(SHEET_TRAINING) == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING, "Missing sheet"));
        }
    }

    private void readSheet(XSSFSheet a_sheet, HashMap<Integer, String> a_header, String a_sheetName)
    {
        Integer t_counterRow = 0;
        // Iterate through each rows one by one
        Iterator<Row> t_rowIterator = a_sheet.iterator();
        while (t_rowIterator.hasNext())
        {
            t_counterRow++;
            Row t_row = t_rowIterator.next();
            if (t_counterRow > 1)
            {
                this.m_file1 = null;
                this.m_file2 = null;
                this.m_file3 = null;
                if (t_row.getFirstCellNum() != -1)
                {
                    if (a_sheetName.equals(SHEET_PAPER))
                    {
                        this.processRowPaper(t_row, t_counterRow);
                    }
                    else if (a_sheetName.equals(SHEET_MEETING_TALK))
                    {
                        this.processRowMeetingTalkPosterDemo(t_row, t_counterRow,
                                SHEET_MEETING_TALK);
                    }
                    else if (a_sheetName.equals(SHEET_POSTER))
                    {
                        this.processRowMeetingTalkPosterDemo(t_row, t_counterRow, SHEET_POSTER);
                    }
                    else if (a_sheetName.equals(SHEET_DEMO))
                    {
                        this.processRowMeetingTalkPosterDemo(t_row, t_counterRow, SHEET_DEMO);
                    }
                    else if (a_sheetName.equals(SHEET_ONLINE_TALK))
                    {
                        this.processRowOnlineTalk(t_row, t_counterRow);
                    }
                    else if (a_sheetName.equals(SHEET_WORKSHOP))
                    {
                        // this.processRowWorkshop(t_row, t_counterRow);
                    }
                    else if (a_sheetName.equals(SHEET_TRAINING))
                    {
                        this.processRowTraining(t_row, t_counterRow);
                    }
                }
            }
            else if (t_counterRow == 1)
            {
                if (!this.validHeader(t_row, a_header, a_sheetName))
                {
                    // errors are already stored
                    return;
                }
            }
        }
        if (t_counterRow < 1)
        {
            this.m_result.getErrors()
                    .add(new ImportError(a_sheetName, "Sheet is empty - no header found."));
        }
    }

    private boolean validHeader(Row a_row, HashMap<Integer, String> a_header, String a_sheetName)
    {
        boolean t_foundError = false;
        for (Integer t_columnNumber : a_header.keySet())
        {
            Cell t_cell = a_row.getCell(t_columnNumber);
            String t_valueString = this.sanitizeString(t_cell);
            if (t_valueString != null)
            {
                if (!t_valueString.equalsIgnoreCase(a_header.get(t_columnNumber)))
                {
                    this.m_result.getErrors()
                            .add(new ImportError(a_sheetName, "Wrong column header (column "
                                    + Integer.toString(t_columnNumber + 1) + "): expected "
                                    + a_header.get(t_columnNumber) + " got " + t_valueString));
                    t_foundError = true;
                }
            }
            else
            {
                this.m_result.getErrors()
                        .add(new ImportError(a_sheetName,
                                "Missing column (" + Integer.toString(t_columnNumber + 1) + "): "
                                        + a_header.get(t_columnNumber)));
                t_foundError = true;
            }
        }
        return !t_foundError;
    }

    private void processRowPaper(Row a_row, Integer a_counterRow)
    {
        OutreachRecord t_record = new OutreachRecord();
        t_record.setCategory("paper");
        // For each row, iterate through all the columns
        for (int t_counterColumn = 0; t_counterColumn < 21; t_counterColumn++)
        {
            Cell t_cell = a_row.getCell(t_counterColumn);
            if (t_counterColumn == 0)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setTitle(t_valueString);
            }
            else if (t_counterColumn == 1)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setAuthorList(t_valueString);
            }
            else if (t_counterColumn == 2)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setJournal(t_valueString);
            }
            else if (t_counterColumn == 3)
            {
                Date t_valueDate = this.sanitizeDate(t_cell);
                if (t_valueDate != null)
                {
                    t_record.setDate(this.formatDate(t_valueDate));
                }
            }
            else if (t_counterColumn == 4)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setIssue(t_valueString);
            }
            else if (t_counterColumn == 5)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setPages(t_valueString);
            }
            else if (t_counterColumn == 6)
            {
                Integer t_valueInt = this.sanitizeInteger(t_cell);
                if (t_valueInt != null)
                {
                    t_record.setPmid(t_valueInt.toString());
                }
            }
            else if (t_counterColumn == 7)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setDoid(t_valueString);
            }
            else if (t_counterColumn == 8)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    if (t_valueString.equalsIgnoreCase("yes"))
                    {
                        t_record.setCite(true);
                    }
                    else if (t_valueString.equalsIgnoreCase("no"))
                    {
                        t_record.setCite(false);
                    }
                }
            }
            else if (t_counterColumn == 9)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 10)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 11)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 12)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 13)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 14)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 15)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 16)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 17)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 18)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 19)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 20)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setUrl(t_valueString);
                }
            }
        }
        if (this.m_file1 != null)
        {
            t_record.getFiles().add(this.m_file1);
        }
        if (this.m_file2 != null)
        {
            t_record.getFiles().add(this.m_file2);
        }
        if (this.m_file3 != null)
        {
            t_record.getFiles().add(this.m_file3);
        }
        this.checkPaper(t_record, a_counterRow);
    }

    private void checkPaper(OutreachRecord a_record, Integer a_counterRow)
    {
        boolean t_foundError = false;
        if (a_record.getTitle() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing title in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getAuthorList() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing author list in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getJournal() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing journal in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getDate() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing date or wrong format in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getPages() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing pages in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getPmid() == null)
        {
            if (a_record.getDoid() == null)
            {
                this.m_result.getErrors()
                        .add(new ImportError(SHEET_PAPER,
                                "Missing PMID and DOID (one must be given) in row: "
                                        + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
        }
        if (a_record.isCite() == null)
        {
            this.m_result.getErrors().add(
                    new ImportError(SHEET_PAPER, "Missing cite value or invalid format in row: "
                            + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getFunding().size() == 0)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                    "Missing funding information in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        for (FileMetadata t_file : a_record.getFiles())
        {
            if (t_file.getFormat() == null)
            {
                this.m_result.getErrors()
                        .add(new ImportError(SHEET_PAPER, "File with missing format found in row: "
                                + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getType() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                        "File with missing type found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getUrl() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_PAPER,
                        "File with missing URL found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
        }
        if (!t_foundError)
        {
            this.m_result.getRecords().add(a_record);
        }
    }

    private FileMetadata accessFile(int a_fileNumber)
    {
        if (a_fileNumber == 1)
        {
            if (this.m_file1 == null)
            {
                this.m_file1 = new FileMetadata();
            }
            return this.m_file1;
        }
        if (a_fileNumber == 2)
        {
            if (this.m_file2 == null)
            {
                this.m_file2 = new FileMetadata();
            }
            return this.m_file2;
        }
        if (this.m_file3 == null)
        {
            this.m_file3 = new FileMetadata();
        }
        return this.m_file3;
    }

    private Integer sanitizeInteger(Cell a_cell)
    {
        if (a_cell == null)
        {
            return null;
        }
        Integer t_value = (int) a_cell.getNumericCellValue();
        return t_value;
    }

    private String sanitizeString(Cell a_cell)
    {
        if (a_cell == null)
        {
            return null;
        }
        String t_value = a_cell.getStringCellValue();
        if (t_value == null)
        {
            return null;
        }
        String t_string = t_value.trim();
        if (t_string.length() == 0)
        {
            return null;
        }
        return t_string;
    }

    private Date sanitizeDate(Cell a_cell)
    {
        if (a_cell == null)
        {
            return null;
        }
        Date t_value = a_cell.getDateCellValue();
        if (t_value == null)
        {
            return null;
        }
        return t_value;
    }

    private String formatDate(Date a_valueDate)
    {
        SimpleDateFormat t_formater = new SimpleDateFormat("yyyy MMM dd");
        // Format the date to Strings
        String t_string = t_formater.format(a_valueDate);
        return t_string;
    }

    private void processRowTraining(Row a_row, Integer a_counterRow)
    {
        OutreachRecord t_record = new OutreachRecord();
        t_record.setCategory("training");
        // For each row, iterate through all the columns
        for (int t_counterColumn = 0; t_counterColumn < 19; t_counterColumn++)
        {
            Cell t_cell = a_row.getCell(t_counterColumn);
            if (t_counterColumn == 0)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setTitle(t_valueString);
            }
            else if (t_counterColumn == 1)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setPresenter(t_valueString);
            }
            else if (t_counterColumn == 2)
            {
                Date t_valueDate = this.sanitizeDate(t_cell);
                if (t_valueDate != null)
                {
                    t_record.setDate(this.formatDate(t_valueDate));
                }
            }
            else if (t_counterColumn == 3)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeeting(t_valueString);
            }
            else if (t_counterColumn == 4)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setLocation(t_valueString);
            }
            else if (t_counterColumn == 5)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeetingUrl(t_valueString);
            }
            else if (t_counterColumn == 6)
            {
                Integer t_valueInt = this.sanitizeInteger(t_cell);
                t_record.setParticipants(t_valueInt);
            }
            else if (t_counterColumn == 7)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 8)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 9)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 10)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 11)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 12)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 13)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 14)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 15)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 16)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 17)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 18)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setUrl(t_valueString);
                }
            }
        }
        if (this.m_file1 != null)
        {
            t_record.getFiles().add(this.m_file1);
        }
        if (this.m_file2 != null)
        {
            t_record.getFiles().add(this.m_file2);
        }
        if (this.m_file3 != null)
        {
            t_record.getFiles().add(this.m_file3);
        }
        this.checkTraining(t_record, a_counterRow);
    }

    private void checkTraining(OutreachRecord a_record, Integer a_counterRow)
    {
        boolean t_foundError = false;
        if (a_record.getTitle() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing title in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getPresenter() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing presenter list in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getDate() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing date or wrong format in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getMeeting() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing meeting name in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getParticipants() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing number of participants in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getFunding().size() == 0)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing funding information in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        for (FileMetadata t_file : a_record.getFiles())
        {
            if (t_file.getFormat() == null)
            {
                this.m_result.getErrors().add(
                        new ImportError(SHEET_TRAINING, "File with missing format found in row: "
                                + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getType() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                        "File with missing type found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getUrl() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                        "File with missing URL found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
        }
        if (!t_foundError)
        {
            this.m_result.getRecords().add(a_record);
        }
    }

    private void processRowOnlineTalk(Row a_row, Integer a_counterRow)
    {
        OutreachRecord t_record = new OutreachRecord();
        t_record.setCategory("online_talk");
        // For each row, iterate through all the columns
        for (int t_counterColumn = 0; t_counterColumn < 17; t_counterColumn++)
        {
            Cell t_cell = a_row.getCell(t_counterColumn);
            if (t_counterColumn == 0)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setTitle(t_valueString);
            }
            else if (t_counterColumn == 1)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setPresenter(t_valueString);
            }
            else if (t_counterColumn == 2)
            {
                Date t_valueDate = this.sanitizeDate(t_cell);
                if (t_valueDate != null)
                {
                    t_record.setDate(this.formatDate(t_valueDate));
                }
            }
            else if (t_counterColumn == 3)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeeting(t_valueString);
            }
            else if (t_counterColumn == 4)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeetingUrl(t_valueString);
            }
            else if (t_counterColumn == 5)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 6)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 7)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 8)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 9)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 10)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 11)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 12)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 13)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 14)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 15)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 16)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setUrl(t_valueString);
                }
            }
        }
        if (this.m_file1 != null)
        {
            t_record.getFiles().add(this.m_file1);
        }
        if (this.m_file2 != null)
        {
            t_record.getFiles().add(this.m_file2);
        }
        if (this.m_file3 != null)
        {
            t_record.getFiles().add(this.m_file3);
        }
        this.checkOnlineTalk(t_record, a_counterRow);
    }

    private void checkOnlineTalk(OutreachRecord a_record, Integer a_counterRow)
    {
        boolean t_foundError = false;
        if (a_record.getTitle() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                    "Missing title in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getPresenter() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                    "Missing presenter list in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getDate() == null)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                    "Missing date or wrong format in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getFunding().size() == 0)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                    "Missing funding information in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        for (FileMetadata t_file : a_record.getFiles())
        {
            if (t_file.getFormat() == null)
            {
                this.m_result.getErrors().add(
                        new ImportError(SHEET_ONLINE_TALK, "File with missing format found in row: "
                                + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getType() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                        "File with missing type found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getUrl() == null)
            {
                this.m_result.getErrors().add(new ImportError(SHEET_ONLINE_TALK,
                        "File with missing URL found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
        }
        if (!t_foundError)
        {
            this.m_result.getRecords().add(a_record);
        }
    }

    private void processRowMeetingTalkPosterDemo(Row a_row, Integer a_counterRow,
            String a_sheetName)
    {
        OutreachRecord t_record = new OutreachRecord();
        if (a_sheetName.equals(SHEET_MEETING_TALK))
        {
            t_record.setCategory("meeting_talk");
        }
        else if (a_sheetName.equals(SHEET_DEMO))
        {
            t_record.setCategory("demo");
        }
        else if (a_sheetName.equals(SHEET_POSTER))
        {
            t_record.setCategory("poster");
        }
        // For each row, iterate through all the columns
        for (int t_counterColumn = 0; t_counterColumn < 18; t_counterColumn++)
        {
            Cell t_cell = a_row.getCell(t_counterColumn);
            if (t_counterColumn == 0)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setTitle(t_valueString);
            }
            else if (t_counterColumn == 1)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setPresenter(t_valueString);
            }
            else if (t_counterColumn == 2)
            {
                Date t_valueDate = this.sanitizeDate(t_cell);
                if (t_valueDate != null)
                {
                    t_record.setDate(this.formatDate(t_valueDate));
                }
            }
            else if (t_counterColumn == 3)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeeting(t_valueString);
            }
            else if (t_counterColumn == 4)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setLocation(t_valueString);
            }
            else if (t_counterColumn == 5)
            {
                String t_valueString = this.sanitizeString(t_cell);
                t_record.setMeetingUrl(t_valueString);
            }
            else if (t_counterColumn == 6)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 7)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 8)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    t_record.getFunding().add(t_valueString);
                }
            }
            else if (t_counterColumn == 9)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 10)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 11)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(1);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 12)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 13)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 14)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(2);
                    t_file.setUrl(t_valueString);
                }
            }
            else if (t_counterColumn == 15)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setType(t_valueString);
                }
            }
            else if (t_counterColumn == 16)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setFormat(t_valueString);
                }
            }
            else if (t_counterColumn == 17)
            {
                String t_valueString = this.sanitizeString(t_cell);
                if (t_valueString != null)
                {
                    FileMetadata t_file = this.accessFile(3);
                    t_file.setUrl(t_valueString);
                }
            }
        }
        if (this.m_file1 != null)
        {
            t_record.getFiles().add(this.m_file1);
        }
        if (this.m_file2 != null)
        {
            t_record.getFiles().add(this.m_file2);
        }
        if (this.m_file3 != null)
        {
            t_record.getFiles().add(this.m_file3);
        }
        this.checkTraining(t_record, a_counterRow, a_sheetName);
    }

    private void checkTraining(OutreachRecord a_record, Integer a_counterRow, String a_sheetName)
    {
        boolean t_foundError = false;
        if (a_record.getTitle() == null)
        {
            this.m_result.getErrors().add(new ImportError(a_sheetName,
                    "Missing title in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getPresenter() == null)
        {
            this.m_result.getErrors().add(new ImportError(a_sheetName,
                    "Missing presenter list in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getDate() == null)
        {
            this.m_result.getErrors().add(new ImportError(a_sheetName,
                    "Missing date or wrong format in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getMeeting() == null)
        {
            this.m_result.getErrors().add(new ImportError(a_sheetName,
                    "Missing meeting name in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        if (a_record.getFunding().size() == 0)
        {
            this.m_result.getErrors().add(new ImportError(SHEET_TRAINING,
                    "Missing funding information in row: " + Integer.toString(a_counterRow)));
            t_foundError = true;
        }
        for (FileMetadata t_file : a_record.getFiles())
        {
            if (t_file.getFormat() == null)
            {
                this.m_result.getErrors()
                        .add(new ImportError(a_sheetName, "File with missing format found in row: "
                                + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getType() == null)
            {
                this.m_result.getErrors().add(new ImportError(a_sheetName,
                        "File with missing type found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
            if (t_file.getUrl() == null)
            {
                this.m_result.getErrors().add(new ImportError(a_sheetName,
                        "File with missing URL found in row: " + Integer.toString(a_counterRow)));
                t_foundError = true;
            }
        }
        if (!t_foundError)
        {
            this.m_result.getRecords().add(a_record);
        }
    }
}
