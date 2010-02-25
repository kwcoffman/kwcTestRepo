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

    private static final com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();

    private XComponentContext xContext = null;
    private XDesktop xDesktop = null;
    private XComponent xCompDoc = null;
    private XMultiComponentFactory xMCF = null;
    
    private String inputFileUrl = "";
    private String outputFileUrl = "";
    private String outputDir = "";
    private String fileType = "";

    private static OOoFormat origFileFormat = null;
    private static OOoFormat ooFileFormat = null;
    private static OOoFormat processFileFormat = null;

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


    public DecompFileProcessor(XComponentContext xCtx, XDesktop xDt, String infile) throws java.lang.Exception {
        xContext = xCtx;
        xDesktop = xDt;
        inputFileUrl = DecompUtil.fileNameToOOoURL(infile);

        xCompDoc = DecompUtil.openFileForProcessing(xDesktop, inputFileUrl);

        // get the remote office service manager
        xMCF = xContext.getServiceManager();

        // Verify it is a supported document type
        fileType = DecompUtil.getDocumentType(xContext, xMCF, inputFileUrl);
        origFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
        if (!DecompUtil.isaSupportedFormat(origFileFormat)) {
            xCompDoc.dispose();
            java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + ": format " + fileType + ": unsupported file format");
            throw e;
        }

    }

    public void printSupportedServices()
    {
        if (xCompDoc != null && DecompUtil.beingVerbose()) DecompUtil.printSupportedServices(xCompDoc);
    }

    public int replaceImage(String repImageFile, int pgnum, int imgnum) throws java.lang.Exception
    {
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;
        
        int handler = processFileFormat.getHandlerType();
        
        switch (handler) {
            case 0:
                DecompText.replaceImage(xContext, xMCF, xCompDoc, "xxxorigname",
                        DecompUtil.fileNameToOOoURL(repImageFile));
                break;
            case 2:
                DecompImpress.replaceImage(xContext, xMCF, xCompDoc, "origname",
                        DecompUtil.fileNameToOOoURL(repImageFile), pgnum, imgnum);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + " is an unsupported file format");
                throw e;
        }
        return 0;
    }

    public int citeImage(String citationText, String citeImageFile, int pgnum, int imgnum) throws java.lang.Exception
    {
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;

        int handler = processFileFormat.getHandlerType();

        switch (handler) {
            case 0:
                    DecompText.insertImageCitation(xContext, xMCF, xCompDoc,
                            citationText, DecompUtil.fileNameToOOoURL(citeImageFile), pgnum, imgnum);
                break;
            case 2:
                    DecompImpress.insertImageCitation(xContext, xMCF, xCompDoc,
                            citationText, DecompUtil.fileNameToOOoURL(citeImageFile), pgnum, imgnum);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + inputFileUrl + " is an unsupported file format");
                throw e;
        }
        return 0;
    }

    public int extractImages(String outputDirectory) throws java.lang.Exception
    {
        outputDir = outputDirectory;
        return extractImages();

    }

    public int extractImages() throws java.lang.Exception
    {
        // Possibly save original document as an OO document
        String newName = DecompUtil.possiblyUseTemporaryDocument(xContext, xMCF,
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
            fileType = DecompUtil.getDocumentType(xContext, xMCF, inputFileUrl);
            ooFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
        }
        processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;

        int handler = processFileFormat.getHandlerType();

        switch (handler) {
            case 0:
                DecompText.extractImages(xContext, xMCF, xCompDoc, outputDir);
                break;
            case 2:
                DecompImpress.extractImages(xContext, xMCF, xCompDoc, outputDir);
                break;
            case 1:
            default:
                java.lang.Exception e = new java.lang.Exception("File " + newName + " is an unsupported file format");
                throw e;
        }
        return 0;
    }

    public void dispose()
    {
        if (xCompDoc != null)
            xCompDoc.dispose();
    }


    public int saveTo(String outputFile) {
        outputFileUrl = DecompUtil.fileNameToOOoURL(outputFile);
        return save();
    }

    public int save()
    {
        if (outputFileUrl != null && outputFileUrl.compareTo("") != 0) {
            mylog.debug("Saving (possibly) modified document to new file, '%s'\n", outputFileUrl);
            DecompUtil.storeDocument(xContext, xMCF, xCompDoc, outputFileUrl, origFileFormat.getFilterName());
        }
        return 0;
    }

}