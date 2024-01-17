package org.glygen.outreach.excel2json.io;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import org.glygen.outreach.excel2json.om.OutreachRecord;

import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.core.exc.StreamWriteException;
import com.fasterxml.jackson.databind.DatabindException;
import com.fasterxml.jackson.databind.ObjectMapper;

public class JSONWriter
{

    public void writeFile(String a_fileNamePath, List<OutreachRecord> a_list)
            throws StreamWriteException, DatabindException, IOException
    {
        ObjectMapper t_mapper = new ObjectMapper();
        t_mapper.setSerializationInclusion(Include.NON_NULL);
        t_mapper.writerWithDefaultPrettyPrinter()
                .writeValue(new FileWriter(new File(a_fileNamePath)), a_list);
    }

}
