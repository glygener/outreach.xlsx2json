package org.glygen.outreach.excel2json.om;

import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

public class OutreachRecord
{
    private String m_title = null;
    private String m_authorList = null;
    private String m_journal = null;
    private String m_date = null;
    private String m_issue = null;
    private String m_pages = null;
    private String m_pmid = null;
    private String m_doid = null;
    private String m_category = null;
    private Boolean m_cite = null;
    private List<String> m_funding = new ArrayList<>();
    private String m_presenter = null;
    private String m_meeting = null;
    private String m_location = null;
    private String m_meetingUrl = null;
    private List<FileMetadata> m_files = new ArrayList<>();
    private Integer m_participants = null;

    @JsonProperty("title")
    public String getTitle()
    {
        return m_title;
    }

    public void setTitle(String a_title)
    {
        m_title = a_title;
    }

    @JsonProperty("author_list")
    public String getAuthorList()
    {
        return m_authorList;
    }

    public void setAuthorList(String a_authorList)
    {
        m_authorList = a_authorList;
    }

    @JsonProperty("journal")
    public String getJournal()
    {
        return m_journal;
    }

    public void setJournal(String a_journal)
    {
        m_journal = a_journal;
    }

    @JsonProperty("date")
    public String getDate()
    {
        return m_date;
    }

    public void setDate(String a_date)
    {
        m_date = a_date;
    }

    @JsonProperty("issue")
    public String getIssue()
    {
        return m_issue;
    }

    public void setIssue(String a_issue)
    {
        m_issue = a_issue;
    }

    @JsonProperty("pages")
    public String getPages()
    {
        return m_pages;
    }

    public void setPages(String a_pages)
    {
        m_pages = a_pages;
    }

    @JsonProperty("pmid")
    public String getPmid()
    {
        return m_pmid;
    }

    public void setPmid(String a_pmid)
    {
        m_pmid = a_pmid;
    }

    @JsonProperty("doid")
    public String getDoid()
    {
        return m_doid;
    }

    public void setDoid(String a_doid)
    {
        m_doid = a_doid;
    }

    @JsonProperty("category")
    public String getCategory()
    {
        return m_category;
    }

    public void setCategory(String a_category)
    {
        m_category = a_category;
    }

    @JsonProperty("cite")
    public Boolean isCite()
    {
        return m_cite;
    }

    public void setCite(Boolean a_cite)
    {
        m_cite = a_cite;
    }

    @JsonProperty("funding")
    public List<String> getFunding()
    {
        return m_funding;
    }

    public void setFunding(List<String> a_funding)
    {
        m_funding = a_funding;
    }

    @JsonProperty("presenter")
    public String getPresenter()
    {
        return m_presenter;
    }

    public void setPresenter(String a_presenter)
    {
        m_presenter = a_presenter;
    }

    @JsonProperty("meeting")
    public String getMeeting()
    {
        return m_meeting;
    }

    public void setMeeting(String a_meeting)
    {
        m_meeting = a_meeting;
    }

    @JsonProperty("location")
    public String getLocation()
    {
        return m_location;
    }

    public void setLocation(String a_location)
    {
        m_location = a_location;
    }

    @JsonProperty("meeting_url")
    public String getMeetingUrl()
    {
        return m_meetingUrl;
    }

    public void setMeetingUrl(String a_meetingUrl)
    {
        m_meetingUrl = a_meetingUrl;
    }

    @JsonProperty("files")
    public List<FileMetadata> getFiles()
    {
        return m_files;
    }

    public void setFiles(List<FileMetadata> a_files)
    {
        m_files = a_files;
    }

    @JsonProperty("participants")
    public Integer getParticipants()
    {
        return m_participants;
    }

    public void setParticipants(Integer a_participants)
    {
        m_participants = a_participants;
    }
}
