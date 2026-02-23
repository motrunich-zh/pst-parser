package psttools;

import com.pff.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class PstScan {
    public static void main(String[] args) {
        if (args.length < 1) {
            printUsage();
            System.exit(1);
        }
        String modeStr = args[0].trim();
        if ("1".equals(modeStr)) {
            String subjectFile = args.length > 1 ? args[1] : "subjects.txt";
            runMode1(subjectFile);
        } else if ("2".equals(modeStr)) {
            runMode2();
        } else {
            System.err.println("Invalid mode. Use 1 or 2.");
            printUsage();
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.err.println("Usage: PstScan <1|2> [subjects-file]");
        System.err.println("  1 = Find PST containing email with subject from list (default list: subjects.txt)");
        System.err.println("  2 = Find first email with sender name but no sender email; print details and PST name");
        System.err.println("Run from the folder that contains .pst files (searches current dir and all subfolders).");
    }

    private static void runMode1(String subjectFile) {
        List<String> subjects = loadSubjects(subjectFile);
        if (subjects.isEmpty()) {
            System.err.println("No subjects loaded from " + subjectFile + ". Put one subject (or substring) per line.");
            System.exit(1);
        }
        File cwd = new File(".").getAbsoluteFile();
        List<File> pstFiles = collectPstFiles(cwd, new ArrayList<File>());
        if (pstFiles.isEmpty()) {
            System.err.println("No .pst files found in " + cwd + " or subfolders.");
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

    private static void runMode2() {
        File cwd = new File(".").getAbsoluteFile();
        List<File> pstFiles = collectPstFiles(cwd, new ArrayList<File>());
        if (pstFiles.isEmpty()) {
            System.err.println("No .pst files found in " + cwd + " or subfolders.");
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