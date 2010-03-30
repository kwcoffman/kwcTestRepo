/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import java.io.IOException;
import java.io.Reader;
import java.io.FileReader;
import java.io.Writer;
import java.io.FileWriter;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

/**
 *
 * @author kwc
 */

public class DecompJson {

    private int id;
    private int version;
    private String resultJsonFileName;
    private DecompFile[] decompFiles;
    private int mainresultcode;
    private String mainresultdetails;

    static public class DecompFile{
        private String inputFile;
        private DecompFileOp[] decompFileOps;
        private int fileresultcode;
        private String fileresultdetails;

        static public class DecompFileOp {
            private String operation;  //private DecompOperation operation;
            private int opresultcode;
            private String opresultdetails;
            private String outputFile;
            private String outputDir;
            private String repImageFile;
            private String citationImageFile;
            private String citationText;
            private int pageNum;
            private int imageNum;
            private boolean includeCustomShapes;
            //private String logLevel;

            public DecompFileOp() {
            }

            public void setOperationResult(int result, String details) {
                this.opresultcode = result;
                this.opresultdetails = details;
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
            if (this.decompFileOps[i].operation == null)
                return null;

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
        public boolean getIncludeCustomShapes(int i) {
            return this.decompFileOps[i].includeCustomShapes;
        }
        public DecompFileOp getFileOp(int i) {
            return this.decompFileOps[i];
        }

        public void setFileResult(int result, String details) {
            this.fileresultcode = result;
            this.fileresultdetails = details;
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

    public int writeResults(String outfilename) {
        Writer writer;
        Gson gson = new GsonBuilder().create();
        String outputJson;

        try {
            writer = new FileWriter(outfilename);
            outputJson = gson.toJson(this, DecompJson.class);
            writer.write(outputJson);
            writer.close();
        } catch (Exception e) {
            return 1;
        }
        return 0;
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

    public String getResultJSONFileName() {
        return this.resultJsonFileName;
    }
    
    public DecompFile getFile(int i)
    {
        return this.decompFiles[i];
    }

    public void setMainResult(int result, String details) {
        this.mainresultcode = result;
        this.mainresultdetails = details;
    }
}