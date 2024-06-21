package btobs;

import com.pff.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Vector;
import java.util.function.BiFunction;
import java.util.function.BiPredicate;
import java.util.stream.IntStream;

/**
 * This simple class allows extracting the attachments of a single message or all messages of an PST-file.
 * All code written here is only a thin facade to the work of Richard Johnson (https://github.com/rjohnsondev) https://github.com/rjohnsondev/java-libpst
 */
public class Extraktor {

        static PSTFile pstFile;
        static Path outputDir;
        static String folderSearchExpression;
        static String messageSearchExpression;

    /**
     * This predicates let pass all objects...
     */
    static public BiPredicate<PSTFolder, String> folderVisitor = (folder, searchText) -> true;
    static public BiPredicate<PSTMessage, String> messageVisitor = (message, searchText) -> true;


    /**
     * Folder filter predicate text search.
     */
    static BiPredicate<PSTFolder, String> folderDisplayNameFilter = (folder, searchText) -> {
        try {
            if (searchText==null || searchText.isBlank()) return false;
            return folder.getDisplayName().toUpperCase().contains(searchText.toUpperCase());
        } catch (Exception e){
            log("error accessing folders displayname", e);
        }
        return false;
    };

    /**
     * Message filter predicate text search.
     */
    static BiPredicate<PSTMessage, String> messageContentFilter = (message, searchText) -> {
        try {
            if (searchText==null || searchText.isBlank()) return false;
            String messageText = MessagteToString(message);
            if (messageText==null || messageText.isBlank()) return false;
            return messageText.toUpperCase().contains(searchText.toUpperCase());
        } catch (Exception e){
            log("error accessing message text", e);
        }
        return false;
    };

    /**
     * Default file naming convention: prefix every filename with the message id.
     */
    BiFunction<Long, String, String> defaultFileNamingRule = (messageId, filename)-> {
            return messageId + "_" + filename;
    };

    /**
     * Main with simple commandline args interpreter.
     * @param args
     */
    public static void main(final String[] args) {

        long messageId = -1;

        // help text
        if (args.length < 3) {
            log("there must be a least two arguments:");
            log("1: a valid path to a pst file");
            log("2:  - a message id ");
            log("    - OR 'ALL' if all attachments of all emails should be exported");
            log("3: a valid path to a writeable directory");
            log("+:    - -f for a string which has to be part of folders displayname (if ALL is given)");
            log("+:    - -m for a string which has to be part of message text (if ALL is given)");
            System.exit(1);
        }

        // pst must be readable
        if(!Files.isReadable(Path.of(args[0]))){
            log("given path <" + args[0] + "> seems not to be a readable pst file");
            System.exit(1);
        }

        // scnd arg must be all or a messageId
        String command = args[1].trim().toUpperCase();
        if (command.equals("ALL")){
            log("export all attachments....");
        } else {
            try {
                messageId = Long.parseLong(args[1]);
            } catch (NumberFormatException e){
                log("given messageId <" + args[1] + "> is not a number");
                System.exit(1);
            }
        }

        // outputdir must be writable
        if(!Files.isWritable(Path.of(args[2]))|| !Files.isDirectory(Path.of(args[2]))){
            log("given path <" + args[2] + "> is not writeable or no directory");
            System.exit(1);
        }
        outputDir = Path.of(args[2]);

        // if searchexpression is given, build expression predicates
        for (int i = 3; i < args.length; i++) {
            if (args[i].toLowerCase().contains("-f")){
                folderSearchExpression = args[i].trim().substring(2);
                folderVisitor = folderDisplayNameFilter;
            }

            if (args[i].toLowerCase().contains("-m")){
                messageSearchExpression= args[i].trim().substring(2);
                messageVisitor = messageContentFilter;
            }
        }

        // try open pst
        try {
            pstFile = new PSTFile(args[0]);
        } catch (PSTException | IOException e) {
            log("can't open pst file <" + args[0] + ">, reason see stackTrace", e);
            System.exit(1);
        }

        // start extraction
        try {
            Extraktor extraktor = new Extraktor();
            if (command.equals("ALL")){
                extraktor.processPstFile(args[0]);
            } else {
                extraktor.processMessageExtraction(messageId);
            }
        } catch (PSTException | IOException e) {
            log("error on extraction pst file <" + args[0] + ">, reason see stackTrace", e);
            System.exit(1);
        }

    }

    /**
     * Default entrypoint with default config.
     * @param filename
     */
    public void processPstFile(String filename) {
        processPstFile(filename, folderVisitor, messageVisitor);
    }

    /**
     * Public entrypoint for executing extraction with own visitor/filter expressions or as simple callback events...
     * @param filename
     * @param pstFolderFilter
     * @param pstMessageFilter
     */
    public void processPstFile(String filename, BiPredicate<PSTFolder, String> pstFolderFilter, BiPredicate<PSTMessage, String> pstMessageFilter) {
        try {
            PSTFile pstFile = new PSTFile(filename);
            log("process PST file <" + pstFile.getMessageStore().getDisplayName() + ">");
            processFolder(pstFile.getRootFolder(), pstFolderFilter, pstMessageFilter);
        } catch (Exception err) {
            err.printStackTrace();
        }
    }

    public void processFolder(PSTFolder folder, BiPredicate<PSTFolder, String> pstFolderFilter, BiPredicate<PSTMessage, String> pstMessageFilter) throws PSTException, IOException {

            // go through the folders...
            if (folder.hasSubfolders()) {
                Vector<PSTFolder> childFolders = folder.getSubFolders();
                for (PSTFolder childFolder : childFolders) {
                    if (pstFolderFilter.test(folder, folderSearchExpression))
                        processFolder(childFolder, pstFolderFilter, pstMessageFilter);
                }
            }
            // and now the emails for this folder
            if (folder.getContentCount() > 0) {
                PSTMessage email = (PSTMessage)folder.getNextChild();
                while (email != null) {
                    if (pstMessageFilter.test(email, messageSearchExpression))
                        saveAttachments(outputDir,defaultFileNamingRule, email);
                    email = (PSTMessage)folder.getNextChild();
                }
            }
    }

    private void processMessageExtraction(long messageId) throws PSTException, IOException {
        PSTMessage message = (PSTMessage) PSTObject.detectAndLoadPSTObject(pstFile, messageId);
        saveAttachments(outputDir, defaultFileNamingRule, message);
    }

    private void saveAttachments(Path outputDir, BiFunction<Long, String, String> fileNameingRule, PSTMessage message)  {

        if (message != null) {
            final int numAttach = message.getNumberOfAttachments();
            if (numAttach == 0) {
                return;
            }

            try {

                // check if the filename is unique, otherwise give him an ordinal suffix
                HashMap<Integer, String> attachmentNames = new HashMap<>();
                for (int x = 0; x < numAttach; x++) {
                    String fileName = fileNameingRule.apply(message.getDescriptorNodeId(), message.getAttachment(x).getLongFilename());
                    int i = 1;
                    while (attachmentNames.containsValue(fileName.toUpperCase()))
                        fileName = fileNameingRule.apply(message.getDescriptorNodeId(), message.getAttachment(x).getLongFilename()+"_"+ ++i);
                    attachmentNames.put(x, fileName.toUpperCase());
                }

                for (int x = 0; x < numAttach; x++) {
                    final PSTAttachment attach = message.getAttachment(x);
                    final InputStream attachmentStream = attach.getFileInputStream();
//                    String filename = attach.getLongFilename();
                    String filename = attachmentNames.get(x);
                    if (filename.isEmpty()) {
                        filename = attach.getFilename();
                    }
                    final FileOutputStream out = new FileOutputStream(outputDir + File.separator + filename);
                    // 8176 is the block size used internally and should
                    // give the best performance
                    final int bufferSize = 8176;
                    final byte[] buffer = new byte[bufferSize];
                    int count;
                    do {
                        count = attachmentStream.read(buffer);
                        out.write(buffer, 0, count);
                    } while (count == bufferSize);
                    out.close();
                    attachmentStream.close();
                }
            } catch (final IOException ioe) {
                log( "Failed writing to file", ioe);
            } catch (final PSTException pste) {
                log( "Error in PST file", pste);
            }
        }
    }


    /**
     * 'PST-object to text' method which enables simple text searches....
     * @implNote not alle possible relevant items are currently extracted, e.g. PSTMessage has a lot of attributes which could be interesting.
     * @param message
     * @return
     */
    public static String MessagteToString(PSTMessage message){
        if (message instanceof PSTContact) {
            final PSTContact contact = (PSTContact) message;
            return contact.toString();
        } else if (message instanceof PSTAppointment) {
            final PSTAppointment task = (PSTAppointment) message;
            return task.toString();
        } else if (message instanceof PSTTask) {
            final PSTTask task = (PSTTask) message;
            return task.toString();
        } else if (message instanceof PSTActivity) {
            final PSTActivity journalEntry = (PSTActivity) message;
            return journalEntry.toString();
        } else if (message instanceof PSTRss) {
            final PSTRss rss = (PSTRss) message;
            return rss.toString();
        } else if (message != null) {
            return message.getSubject() + System.lineSeparator()
                    + message.getRecipientsString() + System.lineSeparator()
                    + message.getSentRepresentingEmailAddress() + System.lineSeparator()
                    + message.getDisplayCC() + System.lineSeparator()
                    + message.getDisplayBCC() + System.lineSeparator()
                    + message.getSenderEmailAddress() + System.lineSeparator()
                    + message.getSenderName() + System.lineSeparator()
                    + message.getBody() + System.lineSeparator()
                    + IntStream.range(1,message.getNumberOfAttachments()).mapToObj(i -> {
                try {
                    return message.getAttachment(i).getLongFilename();
                } catch (PSTException | IOException e) {
                    log("error accessing attachments filename", e);
                }
                return "";
            });
        }
        return null;
    }

    private static void log(String msg){
        System.out.println(msg);
    }

    private static void log(String msg, Exception e){
        log(msg);
        if (e!=null){
            e.printStackTrace();
        }
    }
}
