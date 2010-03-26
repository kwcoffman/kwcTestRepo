package edu.umich.med.umms;

import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;

/**
 *
 * @author kwc
 */
public class DecompFileProcessor {

    private com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();
    private org.apache.log4j.Level myLogLevel = org.apache.log4j.Level.WARN;

    private XComponentContext xContext = null;
    private XDesktop xDesktop = null;
    private XComponent xCompDoc = null;
    private XMultiComponentFactory xMCF = null;
    
    private String inputFileUrl = "";
    private String outputFileUrl = "";
    private String outputDir = "";
    private String fileType = "";

    private OOoFormat origFileFormat = null;
    private OOoFormat ooFileFormat = null;
    private OOoFormat processFileFormat = null;

    private DecompUtil myutil;

    private DecompCitationCollection citations;

    public DecompFileProcessor()
    {
        myutil = new DecompUtil();
        citations = new DecompCitationCollection();
    }

    public DecompFileProcessor(XComponentContext xCtx, XDesktop xDt, String infile) throws java.lang.Exception
    {
        xContext = xCtx;
        xDesktop = xDt;
        myutil = new DecompUtil();

        inputFileUrl = DecompUtil.fileNameToOOoURL(infile);

        xCompDoc = DecompUtil.openFileForProcessing(xDesktop, inputFileUrl);

        // get the remote office service manager
        xMCF = xContext.getServiceManager();

        // Verify it is a supported document type
        fileType = myutil.getDocumentType(xContext, xMCF, inputFileUrl);
        origFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
        if (!DecompUtil.isaSupportedFormat(origFileFormat)) {
            xCompDoc.dispose();
            java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + ": format " + fileType + ": unsupported file format");
            throw e;
        }
        citations = new DecompCitationCollection();
    }

    public void setXComponentContext(XComponentContext xCtx)
    {
        xContext = xCtx;
    }

    public void setXDesktop(XDesktop xDt)
    {
        xDesktop = xDt;
    }

    public void setInputFile(String inputFile)
    {
        inputFileUrl = DecompUtil.fileNameToOOoURL(inputFile);
    }

    public void setOutputFile(String newOutputFile)
    {
        outputFileUrl = DecompUtil.fileNameToOOoURL(newOutputFile);
    }

    public void setOutputDir(String newOutputDir)
    {
        outputDir = newOutputDir;
    }

    public void setLoggingLevel(org.apache.log4j.Level lvl)
    {
        myLogLevel = lvl;
        mylog.setLevel(myLogLevel);
    }

    public void printSupportedServices()
    {
        if (xCompDoc != null)
            myutil.printSupportedServices(xCompDoc);
    }

    private int replaceImage(String repImageFile, int pgnum, int imgnum) throws java.lang.Exception
    {
        int retcode = -1;
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;
        
        int handler = processFileFormat.getHandlerType();
        
        switch (handler) {
            case 0:
                DecompText dt = new DecompText();
                dt.setLoggingLevel(myLogLevel);
                retcode = dt.replaceImage(xContext, xMCF, xCompDoc, "xxxorigname",
                        DecompUtil.fileNameToOOoURL(repImageFile));
                break;
            case 2:
                DecompImpress di = new DecompImpress();
                di.setLoggingLevel(myLogLevel);
                retcode = di.replaceImage(xContext, xMCF, xCompDoc, "origname",
                        DecompUtil.fileNameToOOoURL(repImageFile), pgnum, imgnum);
                break;
            case 1: // Spreadsheets are not currently supported
            default:
                java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + " is an unsupported file format");
                throw e;
        }
        return retcode;
    }

    private int citeImage(String citationText, String citeImageFile, int pgnum, int imgnum) throws java.lang.Exception
    {
        int retcode = -1;
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;

        int handler = processFileFormat.getHandlerType();

        switch (handler) {
            case 0:
                DecompText dt = new DecompText();
                dt.setLoggingLevel(myLogLevel);
                retcode = dt.insertImageCitation(xContext, xMCF, xCompDoc, citationText,
                        DecompUtil.fileNameToOOoURL(citeImageFile), pgnum, imgnum);
                break;
            case 2:
                DecompImpress di = new DecompImpress();
                di.setLoggingLevel(myLogLevel);
                retcode = di.insertImageCitation(xContext, xMCF, xCompDoc, citationText,
                        /*DecompUtil.fileNameToOOoURL(citeImageFile),*/ pgnum, imgnum);
                citations.addCitationEntry(citationText, pgnum, imgnum);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + " is an unsupported file format");
                throw e;
        }
        return retcode;
    }

    private int extractImages(String outputDirectory, boolean excludeCustomShapes) throws java.lang.Exception
    {
        outputDir = outputDirectory;
        return extractImages(excludeCustomShapes);

    }

    private int extractImages(boolean excludeCustomShapes) throws java.lang.Exception
    {
        int retcode = -1;

        if (outputDir == null) {
            java.lang.Exception e = new java.lang.Exception("No output directory specified!");
            throw e;
        }
        // Possibly save original document as an OO document
        String newName = myutil.possiblyUseTemporaryDocument(xContext, xMCF,
                xCompDoc, inputFileUrl, origFileFormat);
        if (newName != null) {
            // Save document in another format and process that
            xCompDoc.dispose();
            inputFileUrl = DecompUtil.fileNameToOOoURL(newName);
            xCompDoc = DecompUtil.openFileForProcessing(xDesktop, inputFileUrl);
            if (xCompDoc == null) {
                java.lang.Exception e = new java.lang.Exception("Unable to open temporary file " + newName + " for processing");
                throw e;
            }
            fileType = myutil.getDocumentType(xContext, xMCF, inputFileUrl);
            ooFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
        }
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;

        int handler = processFileFormat.getHandlerType();

        switch (handler) {
            case 0:
                DecompText dt = new DecompText();
                dt.setLoggingLevel(myLogLevel);
                retcode = dt.extractImages(xContext, xMCF, xCompDoc, outputDir, excludeCustomShapes);
                break;
            case 2:
                DecompImpress di = new DecompImpress();
                di.setLoggingLevel(myLogLevel);
                retcode = di.extractImages(xContext, xMCF, xCompDoc, outputDir, excludeCustomShapes);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + newName != null ? newName : inputFileUrl  + " is an unsupported file format");
                throw e;
        }
        // If a temporary file was used, remove it now
        if (newName != null)
        {
            xCompDoc.dispose();
            DecompUtil.removeTemporaryDocument(newName);
        }
        return retcode;
    }

    /*
     * This is only used for PowerPoint files
     */
    private void addCitationPages(int pageOffset)
    {
        int i, j;
        DecompCitationCollection.DecompCitationCollectionEntry cpe;
        DecompImpress di;

        if (citations.numEntries() == 0)
            return;

        DecompCitationCollection.DecompCitationCollectionEntry[] cpeArray = citations.getDecompCitationCollectionEntryArray();

        di = new DecompImpress();
        di.setLoggingLevel(myLogLevel);
        di.addCitationPages(xCompDoc, cpeArray, pageOffset);

    }

    /*
     * This is only used for PowerPoint files
     */
    public int addFrontMatter(String originFile)
    {
        DecompImpress di = new DecompImpress();
        di.setLoggingLevel(myLogLevel);
        return di.insertFrontBoilerplate(xContext, xDesktop, xMCF, xCompDoc, originFile);
    }
    
    private int save() throws java.lang.Exception
    {
        // Do some extra work for PowerPoint files
        if (origFileFormat.getHandlerType() == 2) {
            int addedFrontPages;

            addedFrontPages = addFrontMatter("/Users/kwc/Downloads/RecompBoilerplate.ppt");  // XXX This needs to be a parameter!!
            if (addedFrontPages < 0)
                addedFrontPages = 0;
            addCitationPages(addedFrontPages);
        }

        if (outputFileUrl != null && outputFileUrl.compareTo("") != 0) {
            mylog.debug("Saving (possibly) modified document to new file, '%s'", outputFileUrl);
            myutil.storeDocument(xContext, xMCF, xCompDoc, outputFileUrl, origFileFormat.getFilterName());
        }
        return 0;
    }

    public void dispose() throws java.lang.Exception
    {
        if (this.xCompDoc != null)
            this.xCompDoc.dispose();
    }


    public int saveTo(String outputFile) throws java.lang.Exception
    {
        outputFileUrl = DecompUtil.fileNameToOOoURL(outputFile);
        return save();
    }


    public String doOperation(DecompParameters dp)
    {
        String ret = null;
        int retcode = -1;
        if (!dp.ValidSingleOp())
            return logOperationInformation(dp, "Invalid parameters", null);
        try {
            switch (dp.getOperation()) {
                case EXTRACT:
                    retcode = this.extractImages(dp.getOutputDir(), dp.getExcludeCustomShapes());
                    break;
                case REPLACE:
                    retcode = this.replaceImage(dp.getRepImageFile(), dp.getPageNum(), dp.getImageNum());
                    break;
                case CITE:
                    retcode = this.citeImage(dp.getCitationText(), dp.getCitationImageFile(), dp.getPageNum(), dp.getImageNum());
                    break;
                case SAVE:
                    retcode = this.saveTo(dp.getOutputFile());
                    break;
            }
        } catch (java.lang.Exception e) {
            return logOperationInformation(dp, "Caught Exception", e.getMessage());
        }
        if (retcode == 0) {
            return null;
        } else {
            return logOperationInformation(dp, "Operation Failed", null);
        }
    }

    private String logOperationInformation(DecompParameters dp, String errMsg, String exceptionMsg)
    {
        StringBuilder msg = new StringBuilder();

        msg.append(errMsg);
        if (exceptionMsg != null) {
            msg.append(" : ");
            msg.append(exceptionMsg);
        }
        msg.append("\n");
        msg.append(dp.toString());
        return msg.toString();
    }
}