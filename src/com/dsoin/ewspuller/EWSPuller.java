package com.dsoin.ewspuller;

import microsoft.exchange.webservices.data.*;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.logging.impl.SimpleLog;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.client.Client;
import org.elasticsearch.client.transport.TransportClient;
import org.elasticsearch.common.transport.InetSocketTransportAddress;
import org.jsoup.Jsoup;
import org.kohsuke.args4j.Argument;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;

import java.io.ByteArrayOutputStream;
import java.io.Console;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;


/**
 * Created by Dmitrii Soin on 28/11/14.
 */
public class EWSPuller {

    @Option(name = "-ews_url", required = false, usage = "Exchange/365 EWS URL, defaults to public 365 URL")
    private static String ews_url = "https://outlook.office365.com/EWS/Exchange.asmx";

    @Option(name = "-user", required = false, usage = "Exchange username/email")
    private static String username = "";

    @Option(name = "-pass", required = false, usage = "Exchange password. Will be read from console if omitted.")
    private static String password = "";

    @Option(name = "-folder", required = false, usage = "Exchange Folder to pull emails from. *Note emails will be deleted after syncing")
    private static String folderName = "";

    @Option(name = "-pst", required = false, usage = "PST file to index")
    private static String pstFile = "";

    @Option(name = "-esport", required = false, usage = "ES port")
    private static int esPort= 9300;

    @Option(name = "-interval", required = false, usage = "Pull interval in minutes")
    private static int pullInterval= 60;

    @Option(name = "-initES", required = false, usage = "Initialize ES index")
    private static boolean initES= false;

    private static SimpleLog log = new SimpleLog(EWSPuller.class.getName());

    private static Client client = new TransportClient().
            addTransportAddress(new InetSocketTransportAddress("localhost", 9300));

    @Argument
    private List<String> arguments = new ArrayList<String>();

    public static void main(final String[] args) throws Exception {
        new EWSPuller().doMain(args);
    }

    public void doMain(String[] args) throws Exception {

        CmdLineParser parser = new CmdLineParser(this);
        PSTHelper pstHelper = new PSTHelper(esPort);

        parser.parseArgument(args);
        if (args.length == 0) {
            parser.printUsage(System.err);
            System.exit(0);
        }

        if (!"".equals(pstFile)) {

            pstHelper.indexPST(pstFile);
            System.exit(0);
        }

        if (initES) {
            pstHelper.prepareIndexesAndMappings();
            System.exit(0);
        }

        Console console = System.console();
        if (console == null) {

            System.out.println("Cannot obtain system console");
            System.exit(0);
        }

        if ("".equals(pstFile) && "".equals(password)) {
            password = new String(console.readPassword("Exchange password:"));
        }

        ExchangeService service = getExchangeService(username, password);
        final FolderId folderID = getFolderId(folderName, service);

        if (folderID == null)
            throw new IllegalStateException("Cannot find folder " + folderName);

        String confirm = console.readLine("Folder " + folderName+" will be synced to ES and all synced " +
                "emails will be deleted. OK to continue? [y/n]:");
        if (!"y".equalsIgnoreCase(confirm))
            throw new IllegalStateException("Exit");

        Runnable puller = () -> {
            try {
                EWSConnection(username, password, folderName);
            } catch (Exception e) {
                log.error(e);
            }
        };
        ScheduledExecutorService executorService = Executors.newSingleThreadScheduledExecutor();
        executorService.scheduleAtFixedRate(puller, 0, pullInterval, TimeUnit.MINUTES);
    }


    private static void EWSConnection(String user, String password, String folderName) throws Exception {
        ExchangeService service = getExchangeService(user, password);


        FolderId indexingId = getFolderId(folderName, service);
        FindItemsResults<Item> emails = service.findItems(indexingId, new ItemView(Integer.MAX_VALUE));
        PropertySet itemPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
        itemPropertySet.setRequestedBodyType(BodyType.HTML);
        for (Item item : emails.getItems()) {
            log.info(item.getSubject() + ":" + ((EmailMessage) item).getFrom().getName() + ":" + item.getHasAttachments());
            String id = pushEmail(item);
            if (item.getHasAttachments()) {
                item.load();
                log.info(item.getAttachments().getCount());
                for (Attachment attach : item.getAttachments()) {
                    pushAttachment(attach, id);

                }
            }
            item.delete(DeleteMode.HardDelete);
        }
    }

    private static ExchangeService getExchangeService(String user, String password) throws URISyntaxException {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials(user, password);
        service.setCredentials(credentials);
        service.setUrl(new URI(ews_url));
        return service;
    }

    private static FolderId getFolderId(String folderName, ExchangeService service) throws Exception {
        Folder rootFolder = Folder.bind(service, WellKnownFolderName.Inbox);
        for (Folder folder:rootFolder.findFolders(new FolderView(100))) {
            log.info(folder.getDisplayName());
            if (folderName.equalsIgnoreCase(folder.getDisplayName())) {
                log.info("Found folder " + folder.getDisplayName() + " with " + folder.getTotalCount() + " emails");
                return folder.getId();
            }
        }
        return null;
    }


    private static void pushAttachment(Attachment attach, String id) throws Exception {

        if (!(attach instanceof FileAttachment))
            return;

        Map<String, Object> emailJson = new HashMap<>();
        String filename = attach.getName();
        ByteArrayOutputStream bs = new ByteArrayOutputStream();
        ((FileAttachment) attach).load(bs);
        byte data[] = Base64.encodeBase64(bs.toByteArray());
        bs.close();
        emailJson.put("filename", filename);
        emailJson.put("email_id", id);
        emailJson.put("size", attach.getSize());
        emailJson.put("attachment", new String(data));
        emailJson.put("mime", attach.getContentType());

        IndexResponse response = client.prepareIndex("attachments", "ssc")
                .setSource(emailJson
                )
                .execute()
                .actionGet();

    }

    private static String pushEmail(Item email) throws Exception {
        Map<String, Object> emailJson = new HashMap<>();

        email.load();
        String topic = email.getSubject().replaceAll("^(FW|Fwd|fwd|FWD|RE|fw|re|Re|Fw):\\s", "");

        emailJson.put("topic", topic);
        emailJson.put("body", Jsoup.parse(email.getBody().toString()).body().text());
        emailJson.put("body_html", email.getBody().toString());
        emailJson.put("submit_time", email.getDateTimeSent());
        emailJson.put("sender", ((EmailMessage) email).getFrom().getName());
        if (email.getHasAttachments())
            emailJson.put("has_attachment", true);


        IndexResponse response = client.prepareIndex("emails", "ssc")
                .setSource(emailJson
                )
                .execute()
                .actionGet();
        return response.getId();

    }

}
