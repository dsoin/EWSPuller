package com.dsoin.ewspuller;

import microsoft.exchange.webservices.data.Attachment;
import microsoft.exchange.webservices.data.FileAttachment;
import org.apache.commons.codec.binary.Base64;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.elasticsearch.action.admin.indices.delete.DeleteIndexRequest;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.action.search.SearchResponse;
import org.elasticsearch.action.search.SearchType;
import org.elasticsearch.client.Requests;
import org.elasticsearch.client.transport.TransportClient;
import org.elasticsearch.common.settings.Settings;
import org.elasticsearch.common.transport.InetSocketTransportAddress;
import org.elasticsearch.index.query.QueryBuilders;
import org.elasticsearch.transport.client.PreBuiltTransportClient;
import org.jsoup.Jsoup;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by freki on 17/08/17.
 */
public class ESHelper {
    private static Logger log = LogManager.getLogger(ESHelper.class);
    private  TransportClient client;
    private  String indexName;
    private  String dataType;

    public ESHelper(String esHost, int esPort, String indexName, String dataType) throws UnknownHostException {
        this.indexName = indexName;
        this.dataType = dataType;
        client = new PreBuiltTransportClient(Settings.EMPTY)
                .addTransportAddress(
                        new InetSocketTransportAddress(InetAddress.getByName(esHost), esPort));

    }


    public  void pushAttachment(String fileName,byte[] content, int size, String contentType, String id) throws Exception {

        Map<String, Object> emailJson = new HashMap<>();
        ByteArrayOutputStream bs = new ByteArrayOutputStream();
        emailJson.put("filename", fileName);
        emailJson.put("email_id", id);
        emailJson.put("size", size);
        emailJson.put("attachment", content);
        emailJson.put("mime", contentType);

        IndexResponse response = client.prepareIndex("attachments",dataType)
                .setSource(emailJson
                )
                .execute()
                .actionGet();

    }

    public  String pushEmail(String subject, String bodyHtml, Date timeSent,
                                    String from, boolean hasAttachments ) throws Exception {
        Map<String, Object> emailJson = new HashMap<>();
        String topic = subject.replaceAll("^(FW|Fwd|fwd|FWD|RE|fw|re|Re|Fw|RFP|Rfp|rfp):\\s", "");
        emailJson.put("topic", topic);
        emailJson.put("body", Jsoup.parse(bodyHtml).body().text());
        emailJson.put("body_html", bodyHtml);
        emailJson.put("submit_time", timeSent);
        emailJson.put("sender", from);
        if (hasAttachments)
            emailJson.put("has_attachment", true);

        IndexResponse response = client.prepareIndex(indexName, dataType)
                .setSource(emailJson
                )
                .execute()
                .actionGet();
        return response.getId();

    }

    public void prepareIndexesAndMappings() throws IOException {


        client.admin().indices().delete(new DeleteIndexRequest(indexName));
        client.admin().indices().delete(new DeleteIndexRequest("attachments"));

        client.admin().indices().preparePutTemplate(indexName).
                setSource(new String(Files.readAllBytes(Paths.get("data-template.json")))).
                execute().actionGet();
        client.admin().indices().preparePutTemplate("attachments").
                setSource(new String(Files.readAllBytes(Paths.get("attachments-template.json")))).
                execute().actionGet();

        client.admin().indices().create(Requests.createIndexRequest(indexName)).actionGet();
        client.admin().indices().create(Requests.createIndexRequest("attachments")).actionGet();
    }

    public boolean alreadyIndexed(String body, String bodyHtml) {
        if ("".equals(body) && "".equals(bodyHtml))
            return true;
        SearchResponse response = client.prepareSearch("data").
                setTypes(dataType).
                setSearchType(SearchType.QUERY_THEN_FETCH).
                setQuery(QueryBuilders.matchPhrasePrefixQuery("body", body)).
                execute().actionGet();
        if (response.getHits().getHits().length > 0)
            return true;
        response = client.prepareSearch("data").
                setTypes(dataType).
                setSearchType(SearchType.QUERY_THEN_FETCH).
                setQuery(QueryBuilders.matchPhrasePrefixQuery("body_html", bodyHtml)).
                execute().actionGet();
        if (response.getHits().getHits().length > 0)
            return true;
        return false;

    }
}
