package org.glygen.outreach.excel2json.om;

import java.util.ArrayList;
import java.util.List;

public class FileReadingResult 
{
	private List<OutreachRecord> m_records = new ArrayList<>();
	private List<ImportError> m_errors = new ArrayList<>();
	public List<OutreachRecord> getRecords() {
		return m_records;
	}
	public void setRecords(List<OutreachRecord> a_records) {
		m_records = a_records;
	}
	public List<ImportError> getErrors() {
		return m_errors;
	}
	public void setErrors(List<ImportError> a_errors) {
		m_errors = a_errors;
	}
}
