package psttools;

import com.pff.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Scan PST files.
 * Modes:
 *   1 = Find PST containing an email whose subject matches any in a list (from file).
 *   2 = Find first email that has sender name but no sender email address; print its details and PST name.
 *   3 = For a given PST and subject, print all details we can read for every matching email.
 *
 * Java 8 compatible.
 *
 * Mode 1:  java ... PstScan 1 [subjects-file] [pst-root-dir]
 *          Default subjects file: subjects.txt. Default dir: current directory.
 * Mode 2:  java ... PstScan 2 [pst-root-dir]
 *          Default dir: current directory.
 * Mode 3:  java ... PstScan 3 \"Subject text\" /path/to/file.pst
 */
public class PstScan {

    public static void main(String[] args) {
        if (args.length < 1) {
            printUsage();
            System.exit(1);
        }
        String modeStr = args[0].trim();
        if ("1".equals(modeStr)) {
            String subjectFile = args.length > 1 ? args[1] : "subjects.txt";
            String pstRoot = args.length > 2 ? args[2] : ".";
            runMode1(subjectFile, pstRoot);
        } else if ("2".equals(modeStr)) {
            String pstRoot = args.length > 1 ? args[1] : ".";
            runMode2(pstRoot);
        } else if ("3".equals(modeStr)) {
            if (args.length < 3) {
                System.err.println("Mode 3 requires a subject string and a PST path.");
                printUsage();
                System.exit(1);
            }
            String subject = args[1];
            String pstPath = args[2];
            runMode3(subject, pstPath);
        } else {
            System.err.println("Invalid mode. Use 1, 2, or 3.");
            printUsage();
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.err.println("Usage: PstScan <1|2|3> [subjects-file] [pst-root-dir]");
        System.err.println("  1 = Find PST containing email with subject from list (default list: subjects.txt)");
        System.err.println("  2 = Find first email with sender name but no sender email; print details and PST name");
        System.err.println("  pst-root-dir = folder to search for .pst files (and subfolders); default: current directory");
        System.err.println("  3 = For a given PST and subject, print all details for matching emails");
        System.err.println("      Example: PstScan 3 \"FW: Duke Longboat Certs\" /path/to/file.pst");
    }

    private static void runMode1(String subjectFile, String pstRootPath) {
        List<String> subjects = loadSubjects(subjectFile);
        if (subjects.isEmpty()) {
            System.err.println("No subjects loaded from " + subjectFile + ". Put one subject (or substring) per line.");
            System.exit(1);
        }
        File pstRoot = resolvePstRoot(pstRootPath);
        if (pstRoot == null) {
            System.exit(1);
        }
        List<File> pstFiles = collectPstFiles(pstRoot, new ArrayList<File>());
        if (pstFiles.isEmpty()) {
            System.err.println("No .pst files found in " + pstRoot.getAbsolutePath() + " or subfolders.");
            System.exit(1);
        }
        System.err.println("Loaded " + subjects.size() + " subject(s). Scanning " + pstFiles.size() + " PST file(s)...");
        for (File pstFile : pstFiles) {
            String path = pstFile.getAbsolutePath();
            try {
                String found = findFirstMatchingSubject(path, subjects);
                if (found != null) {
                    System.out.println("PST_FILE=" + path);
                    System.out.println("MATCHING_SUBJECT=" + found);
                    return;
                }
            } catch (Exception e) {
                System.err.println("Error reading " + path + ": " + e.getMessage());
            }
        }
        System.out.println("No PST contained an email matching the subject list.");
    }

    private static void runMode2(String pstRootPath) {
        File pstRoot = resolvePstRoot(pstRootPath);
        if (pstRoot == null) {
            System.exit(1);
        }
        List<File> pstFiles = collectPstFiles(pstRoot, new ArrayList<File>());
        if (pstFiles.isEmpty()) {
            System.err.println("No .pst files found in " + pstRoot.getAbsolutePath() + " or subfolders.");
            System.exit(1);
        }
        System.err.println("Scanning " + pstFiles.size() + " PST file(s) for first email with sender name but no email...");
        for (File pstFile : pstFiles) {
            String path = pstFile.getAbsolutePath();
            try {
                MessageInfo found = findFirstSenderNameOnly(path);
                if (found != null) {
                    System.out.println("PST_FILE=" + path);
                    System.out.println("SUBJECT=" + nullToEmpty(found.subject));
                    System.out.println("SENDER_NAME=" + nullToEmpty(found.senderName));
                    System.out.println("SENDER_EMAIL=" + nullToEmpty(found.senderEmail));
                    System.out.println("DATE=" + nullToEmpty(found.date));
                    return;
                }
            } catch (Exception e) {
                System.err.println("Error reading " + path + ": " + e.getMessage());
            }
        }
        System.out.println("No email found with sender name only (no sender email).");
    }

    private static void runMode3(String subjectQuery, String pstPath) {
        File pstFile = new File(pstPath).getAbsoluteFile();
        if (!pstFile.isFile()) {
            System.err.println("PST file not found: " + pstFile.getAbsolutePath());
            System.exit(1);
        }
        System.err.println("Mode 3: scanning PST " + pstFile.getAbsolutePath());
        System.err.println("Subject filter (case-insensitive, contains): " + subjectQuery);
        try {
            int count = dumpMessagesBySubject(pstFile.getAbsolutePath(), subjectQuery);
            if (count == 0) {
                System.out.println("No messages found with subject matching: " + subjectQuery);
            } else {
                System.out.println("Total messages printed: " + count);
            }
        } catch (Exception e) {
            System.err.println("Error in mode 3: " + e.getMessage());
            e.printStackTrace(System.err);
            System.exit(1);
        }
    }

    private static List<String> loadSubjects(String path) {
        List<String> list = new ArrayList<String>();
        File f = new File(path);
        if (!f.isFile()) {
            return list;
        }
        BufferedReader br = null;
        try {
            br = new BufferedReader(new InputStreamReader(new FileInputStream(f), StandardCharsets.UTF_8));
            String line;
            while ((line = br.readLine()) != null) {
                line = line.trim();
                if (line.length() > 0 && !line.startsWith("#")) {
                    list.add(line);
                }
            }
        } catch (IOException e) {
            System.err.println("Could not read " + path + ": " + e.getMessage());
        } finally {
            if (br != null) {
                try {
                    br.close();
                } catch (IOException ignored) {
                }
            }
        }
        return list;
    }

    /** Returns the resolved, absolute directory to search for PSTs, or null if invalid (and prints error). */
    private static File resolvePstRoot(String pstRootPath) {
        File dir = new File(pstRootPath).getAbsoluteFile();
        if (!dir.exists()) {
            System.err.println("PST root directory does not exist: " + dir.getAbsolutePath());
            return null;
        }
        if (!dir.isDirectory()) {
            System.err.println("PST root is not a directory: " + dir.getAbsolutePath());
            return null;
        }
        return dir;
    }

    private static List<File> collectPstFiles(File dir, List<File> out) {
        File[] children = dir.listFiles();
        if (children == null) {
            return out;
        }
        for (File f : children) {
            if (f.isDirectory()) {
                collectPstFiles(f, out);
            } else if (f.isFile() && f.getName().toLowerCase(Locale.US).endsWith(".pst")) {
                out.add(f);
            }
        }
        return out;
    }

    private static int dumpMessagesBySubject(String pstPath, String subjectQuery) throws IOException, PSTException {
        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(pstPath);
            PSTFolder root = pstFile.getRootFolder();
            String normalized = subjectQuery.toLowerCase(Locale.US);
            int[] counter = new int[1];
            dumpFolderBySubject(root, "", normalized, counter);
            return counter[0];
        } finally {
            if (pstFile != null) {
                try {
                    pstFile.close();
                } catch (IOException ignored) {
                }
            }
        }
    }

    private static void dumpFolderBySubject(PSTFolder folder, String parentPath, String subjectQueryLower, int[] counter)
            throws PSTException, IOException {
        String currentPath = parentPath.isEmpty()
                ? folder.getDisplayName()
                : parentPath + "/" + folder.getDisplayName();

        PSTObject child = folder.getNextChild();
        while (child != null) {
            if (child instanceof PSTMessage) {
                PSTMessage msg = (PSTMessage) child;
                String subj = null;
                try {
                    subj = msg.getSubject();
                } catch (Exception ignored) {
                }
                if (subj != null && subj.toLowerCase(Locale.US).contains(subjectQueryLower)) {
                    counter[0]++;
                    printFullMessageInfo(counter[0], currentPath, msg);
                }
            }
            child = folder.getNextChild();
        }

        for (PSTFolder sub : folder.getSubFolders()) {
            dumpFolderBySubject(sub, currentPath, subjectQueryLower, counter);
        }
    }

    private static void printFullMessageInfo(int index, String folderPath, PSTMessage msg) {
        System.out.println("========== MESSAGE " + index + " ==========");
        System.out.println("FOLDER_PATH=" + nullToEmpty(folderPath));
        try {
            System.out.println("SUBJECT=" + nullToEmpty(msg.getSubject()));
        } catch (Exception e) {
            System.out.println("SUBJECT=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("SENDER_NAME=" + nullToEmpty(msg.getSenderName()));
        } catch (Exception e) {
            System.out.println("SENDER_NAME=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("SENDER_EMAIL=" + nullToEmpty(msg.getSenderEmailAddress()));
        } catch (Exception e) {
            System.out.println("SENDER_EMAIL=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("DISPLAY_TO=" + nullToEmpty(msg.getDisplayTo()));
        } catch (Exception e) {
            System.out.println("DISPLAY_TO=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("DISPLAY_CC=" + nullToEmpty(msg.getDisplayCC()));
        } catch (Exception e) {
            System.out.println("DISPLAY_CC=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("DISPLAY_BCC=" + nullToEmpty(msg.getDisplayBCC()));
        } catch (Exception e) {
            System.out.println("DISPLAY_BCC=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("DELIVERY_TIME=" + (msg.getMessageDeliveryTime() != null
                    ? msg.getMessageDeliveryTime().toString()
                    : ""));
        } catch (Exception e) {
            System.out.println("DELIVERY_TIME=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("CLIENT_SUBMIT_TIME=" + (msg.getClientSubmitTime() != null
                    ? msg.getClientSubmitTime().toString()
                    : ""));
        } catch (Exception e) {
            System.out.println("CLIENT_SUBMIT_TIME=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("INTERNET_MESSAGE_ID=" + nullToEmpty(msg.getInternetMessageId()));
        } catch (Exception e) {
            System.out.println("INTERNET_MESSAGE_ID=<error: " + e.getMessage() + ">");
        }
        try {
            System.out.println("TRANSPORT_HEADERS_START");
            String headers = msg.getTransportMessageHeaders();
            if (headers != null) {
                for (String line : headers.split("\\r?\\n")) {
                    System.out.println("  " + line);
                }
            }
            System.out.println("TRANSPORT_HEADERS_END");
        } catch (Exception e) {
            System.out.println("TRANSPORT_HEADERS_ERROR=" + e.getMessage());
        }
        try {
            String body = msg.getBody();
            if (body != null && !body.isEmpty()) {
                System.out.println("BODY_TEXT_START");
                System.out.println(body);
                System.out.println("BODY_TEXT_END");
            }
        } catch (Exception e) {
            System.out.println("BODY_TEXT_ERROR=" + e.getMessage());
        }
        try {
            String bodyHtml = msg.getBodyHTML();
            if (bodyHtml != null && !bodyHtml.isEmpty()) {
                System.out.println("BODY_HTML_START");
                System.out.println(bodyHtml);
                System.out.println("BODY_HTML_END");
            }
        } catch (Exception e) {
            System.out.println("BODY_HTML_ERROR=" + e.getMessage());
        }
    }

    private static String findFirstMatchingSubject(String pstPath, List<String> subjects) throws IOException, PSTException {
        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(pstPath);
            PSTFolder root = pstFile.getRootFolder();
            String[] result = new String[1];
            walkFoldersForSubject(root, subjects, result);
            return result[0];
        } finally {
            if (pstFile != null) {
                try {
                    pstFile.close();
                } catch (IOException ignored) {
                }
            }
        }
    }

    private static void walkFoldersForSubject(PSTFolder folder, List<String> subjects, String[] result) throws PSTException, IOException {
        if (result[0] != null) {
            return;
        }
        PSTObject child = folder.getNextChild();
        while (child != null && result[0] == null) {
            if (child instanceof PSTMessage) {
                String subj = ((PSTMessage) child).getSubject();
                if (subj != null) {
                    String subjLower = subj.toLowerCase(Locale.US);
                    for (String s : subjects) {
                        if (subjLower.contains(s.toLowerCase(Locale.US))) {
                            result[0] = subj;
                            return;
                        }
                    }
                }
            }
            child = folder.getNextChild();
        }
        for (PSTFolder sub : folder.getSubFolders()) {
            walkFoldersForSubject(sub, subjects, result);
            if (result[0] != null) {
                return;
            }
        }
    }

    private static class MessageInfo {
        String subject;
        String senderName;
        String senderEmail;
        String date;
    }

    private static MessageInfo findFirstSenderNameOnly(String pstPath) throws IOException, PSTException {
        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(pstPath);
            PSTFolder root = pstFile.getRootFolder();
            MessageInfo[] result = new MessageInfo[1];
            walkFoldersForSenderNameOnly(root, result);
            return result[0];
        } finally {
            if (pstFile != null) {
                try {
                    pstFile.close();
                } catch (IOException ignored) {
                }
            }
        }
    }

    private static void walkFoldersForSenderNameOnly(PSTFolder folder, MessageInfo[] result) throws PSTException, IOException {
        if (result[0] != null) {
            return;
        }
        PSTObject child = folder.getNextChild();
        while (child != null && result[0] == null) {
            if (child instanceof PSTMessage) {
                PSTMessage msg = (PSTMessage) child;
                String email = null;
                String name = null;
                try {
                    email = msg.getSenderEmailAddress();
                } catch (Exception ignored) {
                }
                try {
                    name = msg.getSenderName();
                } catch (Exception ignored) {
                }
                // "No email" = blank or not in email format (library sometimes puts display name in email field)
                boolean hasNoRealEmail = isBlank(email) || !looksLikeEmail(email);
                if (hasNoRealEmail && !isBlank(name)) {
                    MessageInfo info = new MessageInfo();
                    info.senderName = name;
                    info.senderEmail = "";
                    try {
                        info.subject = msg.getSubject();
                    } catch (Exception ignored) {
                    }
                    try {
                        if (msg.getMessageDeliveryTime() != null) {
                            info.date = msg.getMessageDeliveryTime().toString();
                        }
                    } catch (Exception ignored) {
                    }
                    result[0] = info;
                    return;
                }
                if (hasNoRealEmail && isBlank(name) && !isBlank(email)) {
                    // Library put display name in email field
                    MessageInfo info = new MessageInfo();
                    info.senderName = email.trim();
                    info.senderEmail = "";
                    try {
                        info.subject = msg.getSubject();
                    } catch (Exception ignored) {
                    }
                    try {
                        if (msg.getMessageDeliveryTime() != null) {
                            info.date = msg.getMessageDeliveryTime().toString();
                        }
                    } catch (Exception ignored) {
                    }
                    result[0] = info;
                    return;
                }
                if (hasNoRealEmail && isBlank(name) && isBlank(email)) {
                    String fromHeader = fromTransportHeader(msg);
                    if (!isBlank(fromHeader) && !looksLikeEmail(fromHeader)) {
                        MessageInfo info = new MessageInfo();
                        info.senderName = fromHeader.trim();
                        info.senderEmail = "";
                        try {
                            info.subject = msg.getSubject();
                        } catch (Exception ignored) {
                        }
                        try {
                            if (msg.getMessageDeliveryTime() != null) {
                                info.date = msg.getMessageDeliveryTime().toString();
                            }
                        } catch (Exception ignored) {
                        }
                        result[0] = info;
                        return;
                    }
                }
            }
            child = folder.getNextChild();
        }
        for (PSTFolder sub : folder.getSubFolders()) {
            walkFoldersForSenderNameOnly(sub, result);
            if (result[0] != null) {
                return;
            }
        }
    }

    private static String fromTransportHeader(PSTMessage msg) {
        try {
            String headers = msg.getTransportMessageHeaders();
            if (headers != null) {
                for (String line : headers.split("\\r?\\n")) {
                    if (line.toLowerCase(Locale.US).startsWith("from:")) {
                        int colon = line.indexOf(':');
                        if (colon >= 0 && colon + 1 < line.length()) {
                            return line.substring(colon + 1).trim();
                        }
                    }
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private static boolean looksLikeEmail(String s) {
        if (s == null) {
            return false;
        }
        return s.indexOf('@') >= 0 && s.indexOf('.') >= 0;
    }

    private static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static String nullToEmpty(String s) {
        return s == null ? "" : s;
    }
}
