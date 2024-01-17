package org.glygen.outreach.excel2json.om;

public class ImportError
{
    private Integer m_row = null;
    private String m_message = null;
    private String m_sheet = null;

    public ImportError()
    {
        super();
    }

    public ImportError(String a_sheet, String a_message)
    {
        super();
        this.m_sheet = a_sheet;
        this.m_message = a_message;
    }

    public ImportError(String a_sheet, Integer a_row, String a_message)
    {
        super();
        this.m_sheet = a_sheet;
        this.m_row = a_row;
        this.m_message = a_message;
    }

    public Integer getRow()
    {
        return m_row;
    }

    public void setRow(Integer a_row)
    {
        m_row = a_row;
    }

    public String getMessage()
    {
        return m_message;
    }

    public void setMessage(String a_message)
    {
        m_message = a_message;
    }

    public String getSheet()
    {
        return this.m_sheet;
    }

    public void setSheet(String a_sheet)
    {
        this.m_sheet = a_sheet;
    }
}
