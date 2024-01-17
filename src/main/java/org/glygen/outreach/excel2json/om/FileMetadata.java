package org.glygen.outreach.excel2json.om;

import com.fasterxml.jackson.annotation.JsonProperty;

public class FileMetadata {
	private String m_type = null;
	private String m_format = null;
	private String m_url = null;
	
	@JsonProperty("type")
	public String getType() {
		return m_type;
	}
	public void setType(String a_type) {
		m_type = a_type;
	}
	
	@JsonProperty("format")
	public String getFormat() {
		return m_format;
	}
	public void setFormat(String a_format) {
		m_format = a_format;
	}
	
	@JsonProperty("url")
	public String getUrl() {
		return m_url;
	}
	public void setUrl(String a_url) {
		m_url = a_url;
	}
}
