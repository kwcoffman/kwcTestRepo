package edu.umich.med.umms;

import edu.umich.med.umms.DecompJson.DecompFile;
/**
 *
 * @author kwc
 */
public class DecompParameters {
    public DecompOperation operation;
    public String inputFile;
    public String outputFile;
    public String outputDir;

    public String repImageFile;
    public String citationImageFile;
    public String citationText;
    public String jsonCommandFile;
    public int pageNum;
    public int imageNum;

    DecompParameters() {
        operation = null;
        inputFile = null;
        outputFile = null;
        outputDir = null;
        repImageFile = null;
        citationImageFile = null;
        citationText = null;
        jsonCommandFile = null;
        pageNum = -1;
        imageNum = -1;
    }

    DecompParameters(DecompFile df, int opnum)
    {
        operation = df.getOperation(opnum);
        inputFile = df.getInputFile();
        outputFile = df.getOutputFile(opnum);
        outputDir = df.getOutputDir(opnum);
        repImageFile = df.getRepImageFile(opnum);
        citationImageFile = df.getCitationImageFile(opnum);
        citationText = df.getCitationText(opnum);
        pageNum = df.getPageNum(opnum) - 1;
        imageNum = df.getImageNum(opnum) - 1;
    }

    public boolean ParametersAreValid()
    {
        return true;
    }
}