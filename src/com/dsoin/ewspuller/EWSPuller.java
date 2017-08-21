package com.dsoin.ewspuller;

import microsoft.exchange.webservices.data.*;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.logging.impl.SimpleLog;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.sax.BodyContentHandler;
import org.apache.tika.sax.ToXMLContentHandler;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.client.transport.TransportClient;
import org.elasticsearch.common.settings.Settings;
import org.elasticsearch.common.transport.InetSocketTransportAddress;
import org.elasticsearch.transport.client.PreBuiltTransportClient;
import org.jsoup.Jsoup;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

import java.io.*;
import java.net.InetAddress;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Date;
import java.util.HashMap;
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

    @Option(name = "-dataType", required = false, usage = "Data type to use")
    private static String dataType = "emails";

    @Option(name = "-indexName", required = false, usage = "ES index name")
    private static String indexName = "data";

    @Option(name = "-esport", required = false, usage = "ES port")
    private static int esPort = 9300;

    @Option(name = "-eshost", required = false, usage = "ES host")
    private static String esHost = "127.0.0.1";

    @Option(name = "-interval", required = false, usage = "Pull interval in minutes")
    private static int pullInterval = 60;

    @Option(name = "-initES", required = false, usage = "Initialize ES index")
    private static boolean initES = false;

    @Option(name = "-batch_mode", required = false, usage = "Run without user interaction")
    private static boolean batch_mode = false;

    @Option(name = "-ingest_attachments", required = false, usage = "Ingest and index attachments. Email body will be replaced with parsed attachment content.")
    private static boolean ingest_attachments = false;

    private static Logger log = LogManager.getLogger(EWSPuller.class);
    private static ESHelper esHelper;


    public static void main(final String[] args) throws Exception {
        new EWSPuller().doMain(args);
    }

    public void doMain(String[] args) throws Exception {

        CmdLineParser parser = new CmdLineParser(this);
        parser.parseArgument(args);
        if (args.length == 0) {
            parser.printUsage(System.err);
            System.exit(0);
        }

        PSTHelper pstHelper = new PSTHelper(esHost, esPort, indexName, dataType);
        esHelper = new ESHelper(esHost,esPort,indexName, dataType);

        if (!"".equals(pstFile)) {

            pstHelper.indexPST(pstFile, dataType);
            System.exit(0);
        }

        if (initES) {
            esHelper.prepareIndexesAndMappings();
            System.exit(0);
        }

        Console console = null;
        if (!batch_mode) {
            console = System.console();
            if (console == null) {

                System.out.println("Cannot obtain system console");
                System.exit(0);
            }

            if ("".equals(pstFile) && "".equals(password)) {
                password = new String(console.readPassword("Exchange password:"));
            }
        }

        ExchangeService service = getExchangeService(username, password);
        final FolderId folderID = getFolderId(folderName, service);

        if (folderID == null)
            throw new IllegalStateException("Cannot find folder " + folderName);

        if (!batch_mode) {
            String confirm = console.readLine("Folder " + folderName + " will be synced to ES and all synced " +
                    "emails will be deleted. OK to continue? [y/n]:");
            if (!"y".equalsIgnoreCase(confirm))
                throw new IllegalStateException("Exit");
        }

        Runnable puller = () -> {
            try {
                EWSConnection(username, password, folderName);
            } catch (Exception e) {
                log.error("Exiting current pull cycle: ",e);
            }
        };
        ScheduledExecutorService executorService = Executors.newSingleThreadScheduledExecutor();
        executorService.scheduleAtFixedRate(puller, 0, pullInterval, TimeUnit.MINUTES);
    }


    private static void EWSConnection(String user, String password, String folderName) throws Exception {
        log.info("Logging to " +ews_url);
        ExchangeService service = getExchangeService(user, password);
        FolderId indexingId = getFolderId(folderName, service);
        FindItemsResults<Item> emails = service.findItems(indexingId, new ItemView(Integer.MAX_VALUE));
        PropertySet itemPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
        itemPropertySet.setRequestedBodyType(BodyType.HTML);
        for (Item item : emails.getItems()) {
            try {
                log.info(item.getSubject() + ":" + ((EmailMessage) item).getFrom().getName() + ":" + item.getHasAttachments());
                item.load();
                if (ingest_attachments && item.getHasAttachments())
                    ingestAttachmentMode(item);
                else
                    pushEmailMode(item);
                item.delete(DeleteMode.HardDelete);
            } catch (Exception ex) {
                log.error("Skipping email: "+item.getSubject(),ex);
            }
        }
    }

    private static void pushEmailMode(Item item) throws Exception {
        String id = esHelper.pushEmail(item.getSubject(),item.getBody().toString(),
                item.getDateTimeSent(),((EmailMessage) item).getFrom().getName(),item.getHasAttachments());
        if (item.getHasAttachments()) {
            item.load();
            log.info(item.getAttachments().getCount());
            for (Attachment attach : item.getAttachments()) {
                pushAttachment(attach, id);
            }
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
        if (folderName.equals(rootFolder.getDisplayName()))
            return rootFolder.getId();
        for (Folder folder : rootFolder.findFolders(new FolderView(100))) {
            log.info(folder.getDisplayName());
            if (folderName.equalsIgnoreCase(folder.getDisplayName())) {
                log.info("Found folder " + folder.getDisplayName() + " with " + folder.getTotalCount() + " emails");
                return folder.getId();
            }
        }
        return null;
    }


    private static void ingestAttachmentMode(Item email) throws ServiceLocalException {
        for (Attachment attach : email.getAttachments()) {
            ContentHandler handler = new ToXMLContentHandler();
            AutoDetectParser parser = new AutoDetectParser();
            Metadata metadata = new Metadata();

            log.info("Parsing " + attach.getName() + " " + attach.getSize() + " bytes");
            try {
                attach.load();
                try (InputStream stream = new ByteArrayInputStream(((FileAttachment) attach).getContent())) {
                    parser.parse(stream, handler, metadata);
                    log.info(Jsoup.parse(handler.toString()).body().text());
                    if (!"".equals(Jsoup.parse(handler.toString()).body().text())) {
                        String id = esHelper.pushEmail(email.getSubject(),handler.toString(),email.getDateTimeSent(),
                                ((EmailMessage) email).getFrom().getName(),true);
                        pushAttachment(attach,id);
                    }
                }
            } catch (Exception e) {
                log.error("Skipping attachment: " + ((FileAttachment) attach).getName(),e);
            }
        }
    }

    private static void pushAttachment(Attachment attach, String id) throws Exception {
        if (!(attach instanceof FileAttachment))
            return;
        String filename = attach.getName();
        ByteArrayOutputStream bs = new ByteArrayOutputStream();
        ((FileAttachment) attach).load(bs);
        byte data[] = Base64.encodeBase64(bs.toByteArray());
        bs.close();
        esHelper.pushAttachment(filename,data,attach.getSize(),attach.getContentType(),id);
    }

}
