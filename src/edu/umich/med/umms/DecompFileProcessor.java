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
    private String boilerplateUrl = "";

    private OOoFormat origFileFormat = null;
    private OOoFormat ooFileFormat = null;
    private OOoFormat processFileFormat = null;

    private DecompUtil myutil;

    private DecompCitationCollection citations;
    private DecompDelayedRemovalCollection imageRemovals;

    private WatchDog watchdogThread;
    private WatchDogAgain watchdogAgainThread;

    public DecompFileProcessor()
    {
        myutil = new DecompUtil();
        citations = new DecompCitationCollection();
        imageRemovals = new DecompDelayedRemovalCollection();

        xContext = myutil.bootstrap();
        xDesktop = myutil.getDesktop(xContext);
    }

    public DecompFileProcessor(/*XComponentContext xCtx, XDesktop xDt,*/ String infile) throws java.lang.Exception
    {
//        xContext = xCtx;
//        xDesktop = xDt;
        myutil = new DecompUtil();
        xContext = myutil.bootstrap();
        xDesktop = myutil.getDesktop(xContext);

        inputFileUrl = DecompUtil.fileNameToOOoURL(infile);

        xCompDoc = DecompUtil.openFileForProcessing(xDesktop, inputFileUrl);

        // get the remote office service manager
        xMCF = xContext.getServiceManager();

        // Verify it is a supported document type
        fileType = myutil.getDocumentType(xContext, xMCF, inputFileUrl);
        origFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
        if (!DecompUtil.isaSupportedFormat(origFileFormat)) {
            myutil.closeDocument(xContext, xCompDoc);
            java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + ": format " + fileType + ": unsupported file format");
            throw e;
        }
        citations = new DecompCitationCollection();
        imageRemovals = new DecompDelayedRemovalCollection();
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

    public void setBoilerPlateFile(String bpFileName)
    {
        boilerplateUrl = DecompUtil.fileNameToOOoURL(bpFileName);
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
                        DecompUtil.fileNameToOOoURL(repImageFile), pgnum, imgnum, imageRemovals);
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

    private int extractImages(String outputDirectory, boolean includeCustomShapes) throws java.lang.Exception
    {
        outputDir = outputDirectory;
        return extractImages(includeCustomShapes);

    }

    private int extractImages(boolean includeCustomShapes) throws java.lang.Exception
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
            close(xCompDoc);
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
                retcode = dt.extractImages(xContext, xMCF, xCompDoc, outputDir, includeCustomShapes);
                break;
            case 2:
                DecompImpress di = new DecompImpress();
                di.setLoggingLevel(myLogLevel);
                retcode = di.extractImages(xContext, xMCF, xCompDoc, outputDir, includeCustomShapes);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + newName != null ? newName : inputFileUrl  + " is an unsupported file format");
                throw e;
        }
        // If a temporary file was used, remove it now
        if (newName != null)
        {
            close(xCompDoc);
            DecompUtil.removeTemporaryDocument(newName);
        }
        return retcode;
    }

    /*
     * This is only used for PowerPoint files
     */
    private void addCitationPages(DecompImpress di, int pageOffset)
    {
        int i, j;
        DecompCitationCollection.DecompCitationCollectionEntry cpe;

        if (citations.numEntries() == 0)
            return;

        DecompCitationCollection.DecompCitationCollectionEntry[] cpeArray = citations.getDecompCitationCollectionEntryArray();

        di.addCitationPages(xCompDoc, cpeArray, pageOffset);

    }


    private int save() throws java.lang.Exception
    {
        // Do some extra work for PowerPoint files
        if (origFileFormat.getHandlerType() == 2) {
            int addedFrontPages;
            DecompImpress di = new DecompImpress();
            di.setLoggingLevel(myLogLevel);

            // Remove images before messing up page numbers!
            di.removeImages(xContext, xMCF, xCompDoc, imageRemovals);

            addedFrontPages = di.insertFrontBoilerplate(xContext, xDesktop, xMCF, xCompDoc, boilerplateUrl);
            if (addedFrontPages < 0) {
                addedFrontPages = 0;
                if (addedFrontPages == -2) {
                    // Retry the entire file?!?
                    return -2;
                }
            }
            addCitationPages(di, addedFrontPages);
        }

        if (outputFileUrl != null && outputFileUrl.compareTo("") != 0) {
            mylog.debug("Saving (possibly) modified document to new file, '%s'", outputFileUrl);
            myutil.storeDocument(xContext, xMCF, xCompDoc, outputFileUrl, origFileFormat.getFilterName());
        }
        return 0;
    }

    public void close(XComponent doc) throws java.lang.Exception
    {
        myutil.closeDocument(xContext, doc);
    }

    public void close() throws java.lang.Exception
    {
        close(xCompDoc);
    }

    public int saveTo(String outputFile, String boilerplateFile) throws java.lang.Exception
    {
        outputFileUrl = DecompUtil.fileNameToOOoURL(outputFile);
        boilerplateUrl = DecompUtil.fileNameToOOoURL(boilerplateFile);
        return save();
    }


    public String doOperation(DecompParameters dp) throws DecompException
    {
        String ret = null;
        int retcode = -1;
        if (!dp.ValidSingleOp())
            return logOperationInformation(dp, "Invalid parameters", null);
//        System.err.println("doOperation:  creating new watchdog thread...");
//        watchdogThread = new WatchDog(xContext);
//        watchdogAgainThread = new WatchDogAgain();

//        System.err.println("doOperation:  starting the watchdog monitor...");
//        watchdogThread.start();
//        watchdogAgainThread.start();
//        System.err.println("doOperation:  now continuing with the operation at hand!");
        try {
            switch (dp.getOperation()) {
                case EXTRACT:
                    retcode = this.extractImages(dp.getOutputDir(), dp.getIncludeCustomShapes());
                    break;
                case REPLACE:
                    retcode = this.replaceImage(dp.getRepImageFile(), dp.getPageNum(), dp.getImageNum());
                    break;
                case CITE:
                    retcode = this.citeImage(dp.getCitationText(), dp.getCitationImageFile(), dp.getPageNum(), dp.getImageNum());
                    break;
                case SAVE:
                    retcode = this.saveTo(dp.getOutputFile(), dp.getBoilerPlateFile());
                    break;
            }
        } catch (DecompException de) {
            mylog.error("doOperation: Caught DecompException (and re-throwing): " + de.getMessage());
            throw de;
        } catch (java.lang.Exception e) {
            ret = logOperationInformation(dp, "doOperation: CAUGHT EXCEPTION: ", e.getMessage());
            e.printStackTrace();
            throw new DecompException("doOperation: Exception while processing an operation -- should retry?");
	    //return ret;
        } finally {
//            System.err.println("doOperation: [in finally] telling watchdog thread to quit");
//            watchdogThread.signalHalt();
//            watchdogAgainThread.signalOpCompleted(); // If we make it here, we want to cancel killing OpenOffice
        }
        if (retcode == 0) {
            System.err.println("doOperation:  operation completed successfully!");
            // XXX Perhaps do a join here?
            // XXX That would just delay the real work while
            //     waiting for the watchdog thread to complete?
            return null;
        } else {
            return logOperationInformation(dp, "OPERATION FAILED", null);
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
