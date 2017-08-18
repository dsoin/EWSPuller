package com.dsoin.ewspuller;

import com.pff.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.elasticsearch.action.admin.indices.delete.DeleteIndexRequest;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.action.search.SearchResponse;
import org.elasticsearch.action.search.SearchType;
import org.elasticsearch.client.Client;
import org.elasticsearch.client.Requests;
import org.elasticsearch.index.query.QueryBuilders;

import java.io.IOException;
import java.io.InputStream;
import java.net.UnknownHostException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;
import java.util.Vector;

/**
 * Created by soind on 11/21/2014.
 */
public class PSTHelper {
    final private static Logger log = LogManager.getLogger(PSTHelper.class);
    final private ESHelper esHelper;


    public PSTHelper(String esHost, int esPort, String indexName, String dataType ) throws UnknownHostException {
        esHelper = new ESHelper(esHost,esPort,indexName,dataType);
    }

    public void indexPST(String filename, String dataType) throws Exception {
        PSTFile pstFile = new PSTFile(filename);
        indexPSTFolder(pstFile.getRootFolder(), dataType);

    }


    private void indexPSTFolder(PSTFolder folder, String dataType)
            throws Exception {


        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
                indexPSTFolder(childFolder, dataType);
            }
        }

        // and now the emails for this folder
        int pushed = 0;
        if (folder.getContentCount() > 0) {
            log.info(folder.getContentCount());
            PSTMessage email = (PSTMessage) folder.getNextChild();
            while (email != null) {

                if (!esHelper.alreadyIndexed(email.getBody(),email.getBodyHTML())) {
                    String id = esHelper.pushEmail(email.getSubject(),email.getBodyHTML(),
                            email.getClientSubmitTime(),email.getSenderName(),email.hasAttachments());
                    if (email.hasAttachments()) {

                        pushAttachments(email, id);

                    }
                    pushed++;
                }
                try {
                    email = (PSTMessage) folder.getNextChild();
                    //break;
                } catch (PSTException ex) {
                    log.error("Error getting next child", ex);
                    email = null;
                }
            }
        }
        log.info("Pushed " + pushed + " emails");
    }

    private void pushAttachments(PSTMessage email, String id) throws Exception {
        for (int i = 0; i < email.getNumberOfAttachments(); i++) {
            PSTAttachment attach = email.getAttachment(i);
            String filename = attach.getLongFilename().isEmpty() ? attach.getFilename() : attach.getLongFilename();
            byte[] encodedAttach = getAttachment(attach);
            if (encodedAttach != null) {
                esHelper.pushAttachment(filename,encodedAttach,attach.getAttachSize(),attach.getMimeTag(),id);
            }
        }
    }

    private byte[] getAttachment(PSTAttachment attach) throws IOException, PSTException {
        InputStream attachmentStream = attach.getFileInputStream();
        byte[] buffer = new byte[attach.getAttachSize()];
        int count = attachmentStream.read(buffer);
        if (count <= 0)
            return null;
        byte[] endBuffer = new byte[count];
        System.arraycopy(buffer, 0, endBuffer, 0, count);
        attachmentStream.close();
        return Base64.getEncoder().encode(endBuffer);

    }


}
