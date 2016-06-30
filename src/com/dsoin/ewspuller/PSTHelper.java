package com.dsoin.ewspuller;

import com.pff.*;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.logging.impl.SimpleLog;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.action.search.SearchResponse;
import org.elasticsearch.action.search.SearchType;
import org.elasticsearch.client.Client;
import org.elasticsearch.client.Requests;
import org.elasticsearch.client.transport.TransportClient;
import org.elasticsearch.common.transport.InetSocketTransportAddress;
import org.elasticsearch.index.query.QueryBuilders;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.Vector;

/**
 * Created by soind on 11/21/2014.
 */
public class PSTHelper {
    final static SimpleLog log = new SimpleLog(PSTHelper.class.getName());
    private int esPort=9300;

    private  Client client = new TransportClient().
            addTransportAddress(new InetSocketTransportAddress("localhost", esPort));

    public PSTHelper(int port) {
        esPort=port;
    }

    public void indexPST(String filename) throws IOException, PSTException {
        PSTFile pstFile = new PSTFile(filename);
        indexPSTFolder(pstFile.getRootFolder());

    }


    private void indexPSTFolder(PSTFolder folder)
            throws PSTException, IOException {


        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
                indexPSTFolder(childFolder);
            }
        }

        // and now the emails for this folder
        int pushed = 0;
        if (folder.getContentCount() > 0) {
            log.info(folder.getContentCount());
            PSTMessage email = (PSTMessage) folder.getNextChild();
            while (email != null) {

                if (!alreadyIndexed(email)) {
                    String id = pushEmail(email);
                    if (email.hasAttachments()) {
                        pushAttachments(email, id);

                    }
                    pushed++;
                }
                try {
                    email = (PSTMessage) folder.getNextChild();
                    //break;
                } catch (PSTException ex) {
                    log.error("Error getting next child",ex);
                    email=null;
                }
            }
        }
        log.info("Pushed " + pushed +" emails");
    }

    private  boolean alreadyIndexed(PSTMessage email) {
        if ("".equals(email.getBody()) && "".equals(email.getBodyHTML()))
            return true;
        SearchResponse response = client.prepareSearch("emails").setSearchType(SearchType.QUERY_THEN_FETCH).
                setQuery(QueryBuilders.matchPhrasePrefixQuery("body",email.getBody())).execute().actionGet();
        if (response.getHits().getHits().length>0)
            return true;
        response = client.prepareSearch("emails").setSearchType(SearchType.QUERY_THEN_FETCH).
                setQuery(QueryBuilders.matchPhrasePrefixQuery("body_html",email.getBodyHTML())).execute().actionGet();
        if (response.getHits().getHits().length>0)
            return true;
        return false;

    }
    private  void pushAttachments(PSTMessage email, String id) throws PSTException, IOException {
        for (int i=0;i<email.getNumberOfAttachments();i++) {
            PSTAttachment attach = email.getAttachment(i);
            Map<String, Object> emailJson = new HashMap<>();
            String filename = attach.getLongFilename().isEmpty()?attach.getFilename():attach.getLongFilename();
            byte[] encodedAttach = getAttachment(attach);
            if (encodedAttach!=null) {
                emailJson.put("filename", filename);
                emailJson.put("email_id", id);
                emailJson.put("mime", attach.getMimeTag());
                emailJson.put("size", attach.getAttachSize());
                emailJson.put("attachment", new String(encodedAttach));

                IndexResponse response = client.prepareIndex("attachments", "ssc")
                        .setSource(emailJson
                        )
                        .execute()
                        .actionGet();

            }
        }
    }

    private  byte[] getAttachment(PSTAttachment attach) throws IOException, PSTException {
        InputStream attachmentStream = attach.getFileInputStream();
        byte[] buffer = new byte[attach.getAttachSize()];
        int count = attachmentStream.read(buffer);
        if (count<=0)
            return null;
        byte[] endBuffer = new byte[count];
        System.arraycopy(buffer, 0, endBuffer, 0, count);
        attachmentStream.close();
        return Base64.encodeBase64(endBuffer);

    }

    private  String  pushEmail(PSTMessage email) {
        Map<String, Object> emailJson = new HashMap<>();
        emailJson.put("topic", email.getConversationTopic());
        emailJson.put("body", email.getBody());
        emailJson.put("body_html", email.getBodyHTML());
        emailJson.put("submit_time", email.getClientSubmitTime());
        emailJson.put("sender", email.getSenderName());
        emailJson.put("sender_email", email.getSenderEmailAddress());
        emailJson.put("sent_to", email.getDisplayTo());
        if (email.hasAttachments())
            emailJson.put("has_attachment",true);



        IndexResponse response = client.prepareIndex("emails", "ssc")
                .setSource(emailJson
                )
                .execute()
                .actionGet();
        return response.getId();

    }

    private  void prepareIndexesAndMappings() throws IOException {
/*
        try {

            client.admin().indices().delete(new DeleteIndexRequest("emails"));
            client.admin().indices().delete(new DeleteIndexRequest("attachments"));
            client.admin().indices().deleteMapping(new DeleteMappingRequest("emails").types("*")).actionGet();
            client.admin().indices().deleteMapping(new DeleteMappingRequest("attachments").types("*")).actionGet();
        } catch (IndexMissingException ex) {
        }
        */
        client.admin().indices().preparePutTemplate("emails").
                setSource(new String(Files.readAllBytes(Paths.get("emails-template.json")))).

                execute().actionGet();
        client.admin().indices().preparePutTemplate("attachments").
                setSource(new String(Files.readAllBytes(Paths.get("attachments-template.json")))).

                execute().actionGet();

        client.admin().indices().create(Requests.createIndexRequest("emails")).actionGet();
    }


}
