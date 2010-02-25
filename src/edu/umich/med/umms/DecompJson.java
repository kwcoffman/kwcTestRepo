/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.FileReader;
import java.util.ArrayList;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import java.util.Collection;

/**
 *
 * @author kwc
 */

public class DecompJson {

    private int id;
    private int version;
    private DecompFile[] decompFiles;

    static public class DecompFile{
        private String inputFile;
        private DecompFileOp[] decompFileOps;

        static public class DecompFileOp {
            private String operation;  //private DecompOperation operation;
            private String outputFile;
            private String outputDir;
            private String repImageFile;
            private String citationImageFile;
            private String citationText;
            private int pageNum;
            private int imageNum;

            public DecompFileOp() {
            }
        }

        public DecompFile() {
        }

        public int getNumOps() {
            return this.decompFileOps.length;
        }
        public String getInputFile() {
            return this.inputFile;
        }
        public DecompOperation getOperation(int i) {
            if (this.decompFileOps[i].operation.equalsIgnoreCase("extract"))
                return DecompOperation.EXTRACT;
            if (this.decompFileOps[i].operation.equalsIgnoreCase("replace"))
                return DecompOperation.REPLACE;
            if (this.decompFileOps[i].operation.equalsIgnoreCase("cite"))
                return DecompOperation.CITE;
            if (this.decompFileOps[i].operation.equalsIgnoreCase("save"))
                return DecompOperation.SAVE;
            return null;
        }
        public String getOutputFile(int i) {
            return this.decompFileOps[i].outputFile;
        }
        public String getOutputDir(int i) {
            return this.decompFileOps[i].outputDir;
        }
        public String getRepImageFile(int i) {
            return this.decompFileOps[i].repImageFile;
        }
        public String getCitationImageFile(int i) {
            return this.decompFileOps[i].citationImageFile;
        }
        public String getCitationText(int i) {
            return this.decompFileOps[i].citationText;
        }
        public int getPageNum(int i) {
            return this.decompFileOps[i].pageNum;
        }
        public int getImageNum(int i) {
            return this.decompFileOps[i].imageNum;
        }
    }


    public DecompJson() {
    }

    public DecompJson(String filename) throws IOException, java.lang.Exception {
        Reader reader = new FileReader(filename);

        Gson gson = new GsonBuilder().create();

        DecompJson d = gson.fromJson(reader, DecompJson.class);

        this.id = d.id;
        this.version = d.version;
        this.decompFiles = d.decompFiles;

        reader.close();
    }
    public int getVersion() {
        return this.version;
    }
    public int getId() {
        return this.id;
    }

    public int getNumFiles()
    {
        return this.decompFiles.length;
    }

    public DecompFile getFile(int i)
    {
        return this.decompFiles[i];
    }
}

/*
public class DecompJson {

    private int id;
    private int version;
    private ArrayList<DecompFile> decompFiles = new ArrayList<DecompFile>();

    static private class DecompFile{
        private String inputFile;
        private ArrayList<DecompFileOp> decompFileOps = new ArrayList<DecompFileOp>();

        static private class DecompFileOp {
            private String optype;  //private DecompOperation optype;
            private String outputFile;
            private String outputDir;
            private String repImageFile;
            private String citationText;
            private int pageNum;
            private int imageNum;

            public DecompFileOp() {
                System.err.println("DecompFileOp constructor invoked!");
            }
        }

        public DecompFile() {
            System.err.println("DecompFile constructor invoked!");
        }
    }


    public DecompJson() {
        System.err.println("DecompJson constructor invoked!");
    }

    public DecompJson(String filename) throws IOException, java.lang.Exception {
        Reader reader = new FileReader(filename);

        Gson gson = new GsonBuilder().create();

        DecompJson d = gson.fromJson(reader, DecompJson.class);

        this.id = d.id;
        this.version = d.version;
        this.decompFiles = d.decompFiles;

        reader.close();
    }
    public int getVersion() {
        return this.version;
    }
    public int getId() {
        return this.id;
    }
}
*/