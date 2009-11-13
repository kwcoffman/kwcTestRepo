/*
 * OpenOfficeUNODecomposition.java
 *
 * Created on 2009.08.17 - 10:37:43
 *
 */

package edu.umich.med.umms;

import edu.umich.med.umms.OOoFormat;
import java.awt.Frame;
import java.lang.Integer;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.OptionGroup;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.ParseException;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleContext;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.beans.UnknownPropertyException;

import com.sun.star.comp.helper.Bootstrap;

import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.container.NoSuchElementException;

import com.sun.star.document.XExporter;
import com.sun.star.document.XFilter;
import com.sun.star.document.XTypeDetection;
import com.sun.star.document.XStorageBasedDocument;

import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XShapes;
import com.sun.star.drawing.XShape;
//import com.sun.star.drawing.GraphicObjectShape;

import com.sun.star.embed.XStorage;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
//import com.sun.star.gallery.GalleryItemType;
//import com.sun.star.gallery.XGalleryThemeProvider;
//import com.sun.star.gallery.XGalleryTheme;
//import com.sun.star.gallery.XGalleryItem;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;

import com.sun.star.io.XInputStream;
import com.sun.star.io.XStream;

import com.sun.star.lang.XServiceInfo;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.XOutputStreamToByteArrayAdapter;
//import com.sun.star.table.BorderLine;

import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.WrapTextMode;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextGraphicObjectsSupplier;
import com.sun.star.text.XTextEmbeddedObjectsSupplier;


import com.sun.star.ucb.XSimpleFileAccess2;

import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import com.sun.star.util.XSearchable;


/**
 *
 * @author kwc@umich.edu
 */
public class OpenOfficeUNODecomposition {

    private static final OOoFormat[] SUPPORTED_FORMATS = {
        OOoFormat.OpenDocument_Text,
        OOoFormat.Microsoft_Word_97_2000_XP,
        OOoFormat.Microsoft_Word_95,
        OOoFormat.Microsoft_Word_60,
        OOoFormat.Rich_Text_Format,
        //OOoFormat.Microsoft_Excel_97_2000_XP,
        //OOoFormat.Microsoft_Excel_95,
        //OOoFormat.Microsoft_Excel_50,
        //OOoFormat.OpenDocument_Spreadsheet,
        //OOoFormat.OpenOfficeorg_10_Spreadsheet,
        //OOoFormat.Text_CSV,
        OOoFormat.OpenDocument_Presentation,
        OOoFormat.OpenOfficeorg_10_Presentation,
        OOoFormat.Microsoft_PowerPoint_97_2000_XP,
        OOoFormat.Microsoft_PowerPoint_2007_XML
    };

    private static boolean verbose = false;
    private static boolean excludeCustomShapes = false;

    private static boolean functionImageReplacement = false;
    private static boolean functionExtractImages = false;
    private static boolean functionCopyFile = false;
    
    private static final String Usage =
            "[ --extract <options> | --replace <options> | --copy <options> ]\n\n" +
            "Global options: [--verbose]\n\n" +
            "--extract [ --exclude-custom-shapes ] --input <file> --output-dir <dir>\n\n" +
            "--replace <options> --input <file> --newimage <file> --pagenum <pagenum> --imagenum <imagenum>\n\n" +
            "--copy <options> --input <file> --output <file>\n";
    
    private static String inputName;
    private static String outputName;
    private static String outputDir;

    private static int repPageNum = -1;
    private static int repImageNum = -1;
    private static String repImageFile;

    private static OOoFormat origFileFormat = null;
    private static OOoFormat ooFileFormat = null;
    private static OOoFormat processFileFormat = null;


    /** Creates a new instance of OpenOfficeUNODecomposition */
    public OpenOfficeUNODecomposition() {
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)
    {
        XComponentContext xContext = null;
        String temporaryURL;
        Options opt = new Options();
        int err;


        defineOptions(opt);
        err = processArguments(opt, args);
        if (err != 0) {
            System.exit(3);
        }

        int exitCode = 0;
        try {
            // get the remote office component context
            xContext = Bootstrap.bootstrap();
            if (xContext == null) {
                System.err.println("ERROR: Couldn't connect to OpenOffice process.");
                System.exit(2);
            }
            if (verbose) System.out.println("Successfully connected to OpenOffice process!");

            // get the remote office service manager
            XMultiComponentFactory xMCF = xContext.getServiceManager();

            /* A desktop environment contains tasks with one or more
            frames in which components can be loaded. Desktop is the
            environment for components which can instantiate within
            frames. */
            XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class,
                    xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",
                    xContext));

//            String sInFileName = args[0];

            String sFileUrl = fileNameToOOoURL(inputName);

            XComponent xCompDoc = openFileForProcessing(xDesktop, inputName);
            if (xCompDoc == null) {
                System.out.printf("Unable to open original file '%s', aborting!\n", inputName);
                System.exit(3);
            }
            if (verbose) printSupportedServices(xCompDoc);  // original

            // Determine what kind of document it is and how to proceed
            String fileType = getDocumentType(xContext, xMCF, sFileUrl);
            origFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
            if (!isaSupportedFormat(origFileFormat)) {
                System.out.printf("File '%s', format '%s', is not supported!\n",
                        inputName, (origFileFormat != null) ? origFileFormat.getFormatName(): "<unknown>");
                xCompDoc.dispose();
                System.exit(5);
            }

            if (functionExtractImages) {
                // Possibly save original document as an OO document
                String newName = possiblyUseTemporaryDocument(xContext, xMCF,
                        xCompDoc, inputName, origFileFormat);
                if (newName != null) {
                    xCompDoc.dispose();
                    inputName = newName;
                    xCompDoc = openFileForProcessing(xDesktop, newName);
                    if (xCompDoc == null) {
                        System.out.printf("Unable to open temporary file '%s', aborting!\n", newName);
                        System.exit(4);
                    }
                    fileType = getDocumentType(xContext, xMCF, fileNameToOOoURL(inputName));
                    ooFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
                }
                if (verbose) printSupportedServices(xCompDoc);   // possibly OO version
            }

            processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;
            // Decide how to process the document
            int handler = processFileFormat.getHandlerType();
            switch (handler) {
                case 0:
                    if (functionImageReplacement) {
                        replaceTextDocImage(xContext, xMCF, xCompDoc, "xxxorigname", fileNameToOOoURL(repImageFile));

                    } else if (functionExtractImages) {
                        handleTextDocument(xContext, xMCF, xCompDoc, outputDir);
                    }
                    break;
                case 2:
                    if (functionImageReplacement) {
                        //replacePresentationDocImage(xContext, xMCF, xCompDoc, "origname", fileNameToOOoURL(repImageFile), repPageNum, repImageNum);
                    } else if (functionExtractImages) {
                        handlePresentationDocument(xContext, xMCF, xCompDoc, outputDir);
                    }
                    break;
                case 1:
                    //handleSpreadsheetDocument(xContext, xMCF, xCompDoc);
                    // Fall through for now
                default:
                    System.out.printf("File '%s', format '%s', is not supported!\n",
                        inputName, processFileFormat.getFormatName());
                    xCompDoc.dispose();
                    System.exit(6);
            }
//            System.out.printf("Original file '%s' has type '%s'\n",
//                    inputName, getDocumentType(xContext, xMCF, sFileUrl));
            // print document type
//            printDocumentType(xContext, xMCF, sFileUrl);



/*
            com.sun.star.drawing.XDrawView xDrawView =
                    (com.sun.star.drawing.XDrawView) UnoRuntime.queryInterface(
                    com.sun.star.drawing.XDrawView.class, xCompDoc);
            if (xDrawView == null) {
                System.out.println("Failed to get xDrawView");
            } else {
                com.sun.star.container.XComponentEnumeration xCompEnum =
                    (com.sun.star.container.XComponentEnumeration) UnoRuntime.queryInterface(
                    com.sun.star.container.XComponentEnumeration.class, xCompDoc);

                if (xCompEnum == null) {
                    System.out.println("Failed to get xComponentEnumeration");
                } else {
                    xCompEnum.
                }

            }
 */

//            XSearchable xSearchable = (XSearchable) UnoRuntime.queryInterface(XSearchable.class, xCompDoc);

//            if (xSearchable == null) {
//                System.out.println("No XSearchable here!");
//                //System.exit(22);
//            }

//                handleTextDocument(xContext, xMCF, xCompDoc, xTextDoc, outputDirectory);

//                XDrawPageSupplier xDrawPageSuppl =
//                    (XDrawPageSupplier) UnoRuntime.queryInterface(XDrawPageSupplier.class, xCompDoc);
//                if (xDrawPageSuppl != null) {
//                    handleTextDrawDocument(xContext, xMCF, xCompDoc, xDrawPageSuppl, outputDirectory);
//                } else {
//                    System.out.println("The text document does not have images?");
//                }
/*
                handleTextDocumentImages(xContext, xMCF, xCompDoc, xTextDoc, outputDirectory);
            } else {
                XDrawPagesSupplier xDrawPagesSuppl =
                    (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
                if (xDrawPagesSuppl != null) {
                    handleTextDocumentImages2(xContext, xMCF, xCompDoc, null, "/tmp/kwc-hack");
//                    handleDrawDocument(xContext, xMCF, xCompDoc, xDrawPagesSuppl, outputDirectory);
                } else {
                    System.out.println("We don't know how to process this document type!");
                    exitCode = 45;
                }
            }
*/

/*
            int secs = 10;
            int msecs = secs * 1000;

            System.out.println("About to sleep for " + secs + " seconds!");
            Thread.sleep(msecs);

            System.out.println("Back from our nap!  Disposing of the file now.");
 */
            if (outputName.compareTo("") != 0) {
                if (verbose)
                    System.out.printf("Saving (possibly) modified document to a new file, '%s'\n", outputName);
                storeDocument(xContext, xMCF, xCompDoc, outputName, origFileFormat.getFilterName());
            }
            if (verbose) System.out.println("We be done.");
            xCompDoc.dispose();
        } catch (java.lang.Exception e) {
            e.printStackTrace();
        } finally {
            System.exit(exitCode);
        }
    }

    private static void defineOptions(Options opt)
    {
        OptionGroup mainopts = new OptionGroup();
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("extract").create("e"));
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("replace").create("r"));
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("copy").create("c"));
        opt.addOptionGroup(mainopts);
/*
        opt.addOption("e", "extract", false, "Extract images from a document");
        opt.addOption("r", "replace", false, "Replace an image in a document");
        opt.addOption("c", "copy", false, "Copy a document (testing OO conversion)");
*/
        opt.addOption("h", "help", false, "Print this usage information");
        opt.addOption("v", "verbose", false, "Print verbose output information");
        opt.addOption("xc", "exclude-custom-shapes", false, "Do not include Custom shapes when extracting");

        opt.addOption("i", "input", true, "Input file name (full path)");
        opt.addOption("o", "output", true, "Output file name (full path)");
        opt.addOption("n", "newimage", true, "Replacement (new) image file (full path)");
        opt.addOption("d", "output-dir", true, "Name of directory to receive output");

        opt.addOption("p", "pagenum", true, "Page number");
        opt.addOption("g", "imagenum", true, "Image number");
    }

    private static int processArguments(Options opt, String[] args)
    {
        CommandLine cl;
        try {
            BasicParser parser = new BasicParser();
            cl = parser.parse(opt, args);
        } catch (ParseException ex) {
            System.out.printf("Error processing arguments: '%s'\n", ex.getMessage());
            return 1;
        }

        if (cl.hasOption("v")) {
            verbose = true;
        }

        if (cl.hasOption("h")) {
            printUsage(opt);
            return 8;
        }

        if (!cl.hasOption("e") && !cl.hasOption("r") && !cl.hasOption("c")) {
            System.out.printf("You must choose a main function\n");
            printUsage(opt);
            return 2;
        }

        /* extract */
        if (cl.hasOption("e")) {
            functionExtractImages = true;
            if (!cl.hasOption("i") || !cl.hasOption("d")) {
                System.out.printf("For extract, you must specify the input file and output directory.\n");
                printUsage(opt);
                return 3;
            }
            inputName = cl.getOptionValue("i");
            outputDir = cl.getOptionValue("d");

            if (cl.hasOption("xc")) {
                excludeCustomShapes = true;
            }
        }

        /* replace */
        if (cl.hasOption("r")) {
            functionImageReplacement = true;
            if (!cl.hasOption("i") || !cl.hasOption("o") || !cl.hasOption("n") || !cl.hasOption("p") || !cl.hasOption("g")) {
                System.out.printf("For replace, you must specify input, output, and new image files, as well as page and shape numbers.\n");
                printUsage(opt);
                return 4;
            }
            inputName = cl.getOptionValue("i");
            outputName = cl.getOptionValue("o");
            repImageFile = cl.getOptionValue("n");
            repPageNum = Integer.parseInt(cl.getOptionValue("p"));
            repImageNum = Integer.parseInt(cl.getOptionValue("g"));
            
            // Convert page and image numbers to index values
            if (repPageNum <= 0) {
                System.out.printf("Page number of '%d' is invalid", repPageNum);
                printUsage(opt);
                return 5;
            } else {
                repPageNum -= 1;
            }
            if (repImageNum <= 0) {
                System.out.printf("Image number of '%d' is invalid", repImageNum);
                printUsage(opt);
                return 6;
            } else {
                repImageNum -= 1;
            }
        }

        /* copy */
        if (cl.hasOption("c")) {
            functionCopyFile = true;
            if (!cl.hasOption("i") || !cl.hasOption("o")) {
                System.out.printf("For copy, you must specify input and output file names.\n");
                printUsage(opt);
                return 7;
            }
        }

        return 0;
    }

    private static void printUsage(Options opt)
    {
        HelpFormatter hf = new HelpFormatter();
        hf.setWidth(75);
        hf.printHelp(Usage, opt);
    }

    private static XComponent openFileForProcessing(XDesktop xDesktop, String inputFile)
    {
        // Set up to load the document
        PropertyValue propertyValues[] = new PropertyValue[3];
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "Hidden";
        propertyValues[0].Value = new Boolean(false);

        propertyValues[1] = new PropertyValue();
        propertyValues[1].Name = "ReadOnly";
        propertyValues[1].Value = new Boolean(true);

//        propertyValues[2] = new PropertyValue();
//        propertyValues[2].Name = "FilterName";
//        propertyValues[2].Value = new String("pdf_Portable_Document_Format");

//            String sFileName = "/Users/kwc/Downloads/OER/2009-civic-sedan-brochure-OO-modified.odg";
//            String sFileName = "/Users/kwc/Downloads/OER/2009-civic-sedan-brochure.pdf";
//            String sFileName = "/Users/kwc/Downloads/OER/2009-civic-sedan-brochure.reallyapdf";

        // Load the document
        //System.out.print("Opening file '" + sFileName + "' ... ");
        String sFileUrl = fileNameToOOoURL(inputFile);

        XComponentLoader xCompLoader = (XComponentLoader)
                UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);


        XComponent xCompDoc = null;
        try {
            xCompDoc = xCompLoader.loadComponentFromURL(
                sFileUrl, "_blank", 0, propertyValues);
        } catch (java.lang.Exception e) {
            System.out.printf("Failed to open file '%s', error was '%s'\n",
                    inputFile, e.getMessage());
            System.exit(2);
        }
        return xCompDoc;
    }

    private static void handleTextDocument(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String outputDir)
    {
        // Query for the XTextDocument interface
        XTextDocument xTextDoc =
                    (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            System.out.printf("Cannot get XTextDocument interface for Text Document???\n");
            System.exit(7);
        }

        handleTextDocumentImages(xContext, xMCF, xCompDoc, xTextDoc, outputDir);
        //handleTextDocumentImages2(xContext, xMCF, xCompDoc, xTextDoc, outputDir);
        return;     // XXX XXX XXX XXX XXX

/*
        // Querying for the interface XMultiServiceFactory on the xTextDoc
        XMultiServiceFactory xMSFDoc =
                (XMultiServiceFactory) UnoRuntime.queryInterface(
                XMultiServiceFactory.class, xTextDoc);

        if (xMSFDoc == null) {
            System.out.println("Failed to get xMSFDoc from xTextDoc");
        } else {
            Object oGraphic = null;
            try {
                // Create the service GraphicObject
                oGraphic =
                        xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
            } catch (Exception e) {
                e.printStackTrace();
            }

            // Get the text of the document
            com.sun.star.text.XText xText = xTextDoc.getText();

            String t = xText.getString();

            System.out.printf("The document text is:\n======================\n%s\n======================\n", t);

        }
 */
    }

    private static void handlePresentationDocument(XComponentContext xContext,
                                                   XMultiComponentFactory xMCF,
                                                   XComponent xCompDoc,
                                                   String outputDir)
    {
        // Query for the XDrawPagesSupplier interface
        XDrawPagesSupplier xDrawPagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
        if (xDrawPagesSuppl == null) {
            System.out.printf("Cannot get XDrawPagesSupplier interface for Presentation Document???\n");
            System.exit(8);
        }


        handleDrawDocument(xContext, xMCF, xCompDoc, xDrawPagesSuppl, outputDir);
        return;

    }

    private static void handleTextDocumentImages(XComponentContext xContext,
                                                XMultiComponentFactory xMCF,
                                                XComponent xCompDoc,
                                                XTextDocument xTextDoc,
                                                String outputDir)
    {
        // XTextFramesSupplier
//        try {
            XTextFramesSupplier xTextFramesSuppl = (XTextFramesSupplier)
                    UnoRuntime.queryInterface(XTextFramesSupplier.class, xCompDoc);
            XNameAccess xTFnames = xTextFramesSuppl.getTextFrames();
            if (xTFnames.hasElements()) {
                // yada yada yada
            }

//        } catch (Exception ex) {
//            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
//        }



        // XTextGraphicObjectsSupplier
        try {
//            XMultiServiceFactory xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(
//                XMultiServiceFactory.class, xCompDoc);

//            xGraphicProvider.storeGraphic(arg0, arg1);
            Object graphicProviderObject = xMCF.createInstanceWithContext(
                    "com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider)
                    UnoRuntime.queryInterface(XGraphicProvider.class, graphicProviderObject);

            XTextGraphicObjectsSupplier xTextGraphSuppl = (XTextGraphicObjectsSupplier)
                    UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xGOnames = xTextGraphSuppl.getGraphicObjects();
            if (xGOnames.hasElements()) {
                Class objclass = xGOnames.getClass();
                String[] graphicNames = xGOnames.getElementNames();
                int numImages = graphicNames.length;
                if (verbose) System.out.printf("There are %d GraphicObjects in this file\n", numImages);
                /*
                for (int i = 0; i < numImages; i++) {
                    try {
                        System.out.println("The name of the image is: " + graphicNames[i]);
                        //Object graphicsObject = xGOnames.getByName(graphicNames[i]);
                        XShape graphicShape = (XShape) UnoRuntime.queryInterface(XShape.class, xGOnames.getByName(graphicNames[i]));
                        printShapeProperties(graphicShape);
                        exportImage(xMCF, xContext, graphicShape, outputDir, 0, i + 1);
                    } catch (NoSuchElementException ex) {
                        Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (WrappedTargetException ex) {
                        Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
                */
                int i = 0;
                for (String name:graphicNames) {
                    try {
                        i++;
                        if (verbose) System.out.printf("The name of the image is: '%s'\n", name);
                        //Object graphicsObject = xGOnames.getByName(graphicNames[i]);
                        XShape graphicShape = (XShape)
                                UnoRuntime.queryInterface(XShape.class, xGOnames.getByName(name));
                        XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, graphicShape);
                        String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                        pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                        String outName = constructBaseImageName(outputDir, 1, i);
                        extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
//                        printShapeProperties(graphicShape);
//                        int replaceResult = replaceTextDocImage(xContext, xMCF, xCompDoc, xTextDoc, name, new String("file:///Users/kwc/Private/Pictures/Dogs/Haley-01.jpg"));
//                        exportImage(xContext, xMCF, graphicShape, outputDir, 0, i);
//                        storeImage(xContext, xMCF, xTextDoc, graphicShape, outputDir, 0, i);
                    } catch (NoSuchElementException ex) {
                        Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (WrappedTargetException ex) {
                        Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
                    }

                }
            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }

/*
        // XTextEmbeddedObjectsSupplier
        try {

            XTextEmbeddedObjectsSupplier xTextEmbeddedSuppl = (XTextEmbeddedObjectsSupplier) UnoRuntime.queryInterface(XTextEmbeddedObjectsSupplier.class, xCompDoc);
            XNameAccess xEOnames = xTextEmbeddedSuppl.getEmbeddedObjects();
        // Querying for the interface XMultiServiceFactory on the xTextDoc
            XMultiServiceFactory xMSFDoc =
            (XMultiServiceFactory) UnoRuntime.queryInterface(
            XMultiServiceFactory.class, xTextDoc);
            if (xMSFDoc == null) {
            System.out.println("Failed to get xMSFDoc from xTextDoc");
            } else {
            Object oGraphic = null;
            try {
            // Create the service GraphicObject
            oGraphic =
            xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
            } catch (Exception e) {
            e.printStackTrace();
            }
            // Getting the text
            com.sun.star.text.XText xText = xTextDoc.getText();
            String t = xText.getString();
            System.out.printf("The document text is:\n======================\n%s\n======================\n", t);
            }

        } catch (Exception ex) {
          Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
*/

/*
        // Querying for the interface XMultiServiceFactory on the xTextDoc
        XMultiServiceFactory xMSFDoc =
                (XMultiServiceFactory) UnoRuntime.queryInterface(
                XMultiServiceFactory.class, xTextDoc);

        if (xMSFDoc == null) {
            System.out.println("Failed to get xMSFDoc from xTextDoc");
        } else {
            Object oGraphic = null;
            try {
                // Create the service GraphicObject
                oGraphic =
                        xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
            } catch (Exception e) {
                e.printStackTrace();
            }

            // Getting the text
            com.sun.star.text.XText xText = xTextDoc.getText();

            String t = xText.getString();

            System.out.printf("The document text is:\n======================\n%s\n======================\n", t);
        }
*/
    }

/* (NOT USED!!!)
    //
    // For the two test files I tried, the hasByName("Pictures") returns nothing?
    //
    private static void handleTextDocumentImages2(XComponentContext xContext,
                                                  XMultiComponentFactory xMCF,
                                                  XComponent xCompDoc,
                                                  XTextDocument xTextDoc,
                                                  String outputDir)
    {
        try {
            Object oSFAcc = xMCF.createInstanceWithContext(
                    "com.sun.star.ucb.SimpleFileAccess", xContext);
            XSimpleFileAccess2 xFileWriter = (XSimpleFileAccess2)
                    UnoRuntime.queryInterface(XSimpleFileAccess2.class, oSFAcc);
            XStorageBasedDocument xStorageBasedDocument = (XStorageBasedDocument)
                    UnoRuntime.queryInterface(XStorageBasedDocument.class, xCompDoc);

            XStorage xDocStorage = xStorageBasedDocument.getDocumentStorage();
            XStorage xDocPictures = (XStorage) UnoRuntime.queryInterface(
                    XStorage.class, xDocStorage.getByName("Pictures"));

            XNameAccess xDocStorageNameAccess = (XNameAccess)
                    UnoRuntime.queryInterface(XNameAccess.class, xDocStorage);

            String[] allNames = xDocStorageNameAccess.getElementNames();
            if (verbose) {
                for (int i = 0; i < allNames.length; i++) {
                    System.out.printf("The big list has name '%s'\n", allNames[i]);
                }
            }

            if (!xDocStorageNameAccess.hasByName("Pictures")) {
                System.out.printf("Found no \"Pictures\" in the document!!!\n");
                return;
            }

            Object oPicturesStorage = xDocStorageNameAccess.getByName("Pictures");
            XNameAccess xPicturesNameAccess = (XNameAccess)
                    UnoRuntime.queryInterface(XNameAccess.class, oPicturesStorage);

            String[] aNames = xPicturesNameAccess.getElementNames();
            if (verbose) {
                System.out.printf("There were a total of %d pictures found via DocStorageAccess\n", aNames.length);
                for (int i = 0; i < aNames.length; i++) {
                    System.out.printf("Picture %d has name '%s'\n", i+1, aNames[i]);
                }
            }
            for (int i = 0; i < aNames.length; i++) {
                if (verbose) System.out.printf("Processing picture with name '%s'\n", aNames[i]);

                Object oElement = xPicturesNameAccess.getByName(aNames[i]);
                XStream xStream = (XStream) UnoRuntime.queryInterface(XStream.class, oElement);
                Object oInput = xStream.getInputStream();
                XInputStream xInputStream = (XInputStream)
                        UnoRuntime.queryInterface(XInputStream.class, oInput);

                xFileWriter.writeFile(outputDir + "/" + aNames[i], xInputStream);
            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            return;
        }
    }

*/


/* (Returns only "FrameShape" shapes...)

    private static void handleTextDocumentImages(XComponentContext xContext,
                                                XMultiComponentFactory xMCF,
                                                XComponent xCompDoc,
                                                String outputDir) {

        XDrawPageSupplier xDrawPageSuppl = (XDrawPageSupplier) UnoRuntime.queryInterface(XDrawPageSupplier.class, xCompDoc);

        XDrawPage docPage = xDrawPageSuppl.getDrawPage();

        int OpenOfficeUNODecomposition = docPage.getCount();

        for (int s = 1; s <= OpenOfficeUNODecomposition; s++) {
            XShape currShape = getPageShape(docPage, s-1);
            if (currShape == null) {
                System.out.printf("Failed to get currShape (%d) from page %d!\n", s, 1);
                xCompDoc.dispose();
                System.exit(33);
            }
            String currType = currShape.getShapeType();
            com.sun.star.awt.Size shapeSize = currShape.getSize();
            //com.sun.star.awt.Point shapePoint = currShape.getPosition();
            System.out.printf("--- Working with shape %d (size %dx%d)\ttype: %s---\n", s, shapeSize.Width, shapeSize.Height, currType);
//            try {
//                printShapeProperties(currShape);
//            } catch (WrappedTargetException ex) {
//                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
//            }

        }
    }
 */


/* NEVER MIND!   This uses XZipFileAccessor...  I don't think that's what I want...
 * This is from http://www.oooforum.org/forum/viewtopic.phtml?t=83756
 *
private static void exportEmbeddedGraphics(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String outputDir, String hrefPrefix){
      try{
         Object graphicProviderObject = xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", xContext);
         XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(XGraphicProvider.class, graphicProviderObject);

         XZipFileAccessor zipFileAccessor = new XZipFileAccessor(xDoc,ooConnector);

         //for all graphics
         XTextGraphicObjectsSupplier xTextGraphicObjectsSupplier = (XTextGraphicObjectsSupplier)UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
         XNameAccess xNameAccess = xTextGraphicObjectsSupplier.getGraphicObjects();
         String[] names = xNameAccess.getElementNames();
         for (int i=0;i<names.length;i++){
            XTextContent xContent = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, xNameAccess.getByName(names[i]));
            XPropertySet propSet = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xContent);
            String graphicURL = propSet.getPropertyValue("GraphicURL").toString();

            XGraphic xGraphicToExport = null;
            if (graphicURL.startsWith("vnd.sun.star.GraphicObject:")){//graphic id embedded
               graphicURL = graphicURL.substring("vnd.sun.star.GraphicObject:".length());
               //get graphic stream
               XInputStream graphicStream = (XInputStream) UnoRuntime.queryInterface(XInputStream.class,zipFileAccessor.getZipFileAccess().getStreamByPattern("*Pictures/"+ graphicURL + "*"));
               PropertyValue[] arParams = new PropertyValue[1];
               arParams[0] = OOUtils.createPropertyValue("InputStream",graphicStream);
               xGraphicToExport = xGraphicProvider.queryGraphic(arParams);
            }//if
            else if (graphicURL.startsWith("file:")){
               //if (!path.startsWith(destDir.replace("\\", "/"))){
                  PropertyValue[] arParams = new PropertyValue[1];
                  arParams[0] = OOUtils.createPropertyValue("URL",graphicURL);
                  xGraphicToExport = xGraphicProvider.queryGraphic(arParams);
               //}
            }
            if (xGraphicToExport != null){
               //save graphic to file
               PropertyValue[] params = new PropertyValue[2];
                 params[0] = OOUtils.createPropertyValue("MimeType","image/jpeg");
                 String fileName = names[i]+".jpg";
                 File file = new File(outputDir);
                 file.mkdirs();
                 params[1] = OOUtils.createPropertyValue("URL",ooConnector.createUNOFileURL(destDir+"/"+fileName));
               xGraphicProvider.storeGraphic(xGraphicToExport, params);

               //replace url to new location
               propSet.setPropertyValue("GraphicURL",hrefPrefix+"/"+fileName);
            }
         }//for
      }
      catch(java.lang.Exception e){
         e.printStackTrace();
      }
   }//exportEmbeddedGraphics
 */


/* (This is not used...)
 *
    private static void handleTextDrawDocument(XComponentContext xContext,
    XMultiComponentFactory xMCF,
    XComponent xCompDoc,
    XDrawPageSupplier xDrawPageSuppl,
    String outputDir)
    {
    try {
    XDrawPage xDrawPage = xDrawPageSuppl.getDrawPage();
    Object firstPage = xDrawPage.getByIndex(0);

    exportContextImage(xContext, xMCF, xDrawPage, outputDir, 1);
    int pageCount = xDrawPage.getCount();
    System.out.printf("xDrawPage.getCount returned a value of '%d' pages\n", pageCount);
    XDrawPage currPage = null;
    Class pageClass = null;
    XPropertySet pageProps = null;

    // Loop through all the pages of the document
    for (int p = 0; p < pageCount; p++) {
    currPage = getTextDrawPage(xDrawPage, p);
    if (currPage == null) {
    System.out.printf("Failed to get currPage at page %d!\n", p + 1);
    xCompDoc.dispose();
    System.exit(22);
    }
    System.out.printf("=== Working with page %d ===\n", p + 1);
    exportContextImage(xContext, xMCF, currPage, outputDir, p + 1);

    int OpenOfficeUNODecomposition = currPage.getCount();
    System.out.printf("Page %d has %d shapes\n", p + 1, OpenOfficeUNODecomposition);
    XShape currShape = null;

    // Loop through all the shapes within the page
    for (int s = 0; s < OpenOfficeUNODecomposition; s++) {
    currShape = getPageShape(currPage, s);
    if (currShape == null) {
    System.out.printf("Failed to get currShape (%d) from page %d!\n", s + 1, p + 1);
    xCompDoc.dispose();
    System.exit(33);
    }
    String currType = currShape.getShapeType();
    com.sun.star.awt.Size shapeSize = currShape.getSize();
    com.sun.star.awt.Point shapePoint = currShape.getPosition();
    System.out.printf("--- Working with shape %d (At %d:%d, size %dx%d)\ttype: %s---\n", s + 1, shapePoint.X, shapePoint.Y, shapeSize.Width, shapeSize.Height, currType);

    //printShapeProperties(currShape);

    if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
    System.out.printf("Handling GraphicObjectShape (%d) on page %d\n", s + 1, p + 1);
    //System.out.printf("The URL for this shape is '%s'", currPropSet.getPropertyValue("GraphicURL"));
    exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.TableShape")) {
    System.out.printf("Handling TableShape (%d) on page %d\n", s + 1, p + 1);
    //System.out.printf("The URL for this shape is '%s'", currPropSet.getPropertyValue("GraphicURL"));
    exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.GroupShape")) {
    System.out.printf("Handling GroupShape (%d) on page %d\n", s + 1, p + 1);
    //System.out.printf("The URL for this shape is '%s'", currPropSet.getPropertyValue("GraphicURL"));
    exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.CustomShape")) {
    System.out.printf("Handling CustomShape (%d) on page %d\n", s + 1, p + 1);
    //System.out.printf("The URL for this shape is '%s'", currPropSet.getPropertyValue("GraphicURL"));
    handleCustomShape(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
    } else {
    System.out.printf("SKIPPING %s (%d) on page %d\n", currType, s + 1, p + 1);
    }
    }
    }
    } catch (IndexOutOfBoundsException ex) {
    Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
    } catch (WrappedTargetException ex) {
    Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
    }
    }
*/

    private static byte[] readStream(InputStream inStream) throws IOException
    {
        byte[] bytes = new byte[1024];
        int numRead;
        ByteArrayOutputStream outStream = new ByteArrayOutputStream(1024);

        while ((numRead = inStream.read(bytes)) > 0) {
            outStream.write(bytes, 0, numRead);
        }

        byte[] outBytes = outStream.toByteArray();

        outStream.close();

        return outBytes;
    }

    private static byte[] GetImageByteStream(String imageURL)
    {
        byte[] returnBytes = null;
        InputStream inStream = null;
        URL theURL = null;

        try {
            theURL = new URL(imageURL);
            inStream = theURL.openStream();
            returnBytes = readStream(inStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (inStream != null) {
                    inStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return returnBytes;
    }

    private static XPropertySet duplicateObjectPropertySet(Object origObj)
    {
        XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, origObj);
//        XPropertySetInfo origPropSet = xPropSet.getPropertySetInfo();
//        Property[] origProps = origPropSet.getProperties();
        return xPropSet;
    }

    // XXX Revise this to use index value rather than a name!!!!! XXX
    /* From http://www.oooforum.org/forum/viewtopic.phtml?t=81870 */
    private static int replaceTextDocImage(XComponentContext xContext,
                                           XMultiComponentFactory xMCF,
                                           XComponent xCompDoc,
                                           String originalImageName,
                                           String replacementURL)
    {

        XTextDocument xTextDoc =
                    (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            System.out.printf("Cannot get XTextDocument interface for Text Document???\n");
            System.exit(7);
        }

        byte[] replacementByteArray = GetImageByteStream(replacementURL);
        ByteArrayToXInputStreamAdapter xSource = new ByteArrayToXInputStreamAdapter(replacementByteArray);

        // Querying for the interface XMultiServiceFactory on the xtextdocument
        XMultiServiceFactory xMSFDoc = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, xTextDoc);
        Object oGraphic = null;
        try {
            // Creating the service GraphicObject
            oGraphic = xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
            XNamed xGOName = (XNamed) UnoRuntime.queryInterface(XNamed.class, oGraphic);
            xGOName.setName(originalImageName);
        } catch (Exception exception) {
            System.out.println("Could not create TextGraphicObject instance");
            return 1;
        }

        // Get the original image

        Any xImageAny = null;
        Object xImageObject = null;
        try {
            XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xTextDoc);
            XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
            xImageAny = (Any) xNAGraphicObjects.getByName(originalImageName);
        } catch (NoSuchElementException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
        if (xImageAny == null) {
            return 1;
        }
        xImageObject = xImageAny.getObject();
        XTextContent xImage = (XTextContent) xImageObject;


        // Querying for the interface XTextContent on the GraphicObject
        XTextContent xTextContent = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, oGraphic);
        XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oGraphic);

  
        // Read the source
        PropertyValue[] sourceProps = new PropertyValue[1];
        sourceProps[0] = new PropertyValue();
        sourceProps[0].Name = "InputStream";
        sourceProps[0].Value = xSource;
        try {

            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(
                XGraphicProvider.class,
                xMCF.createInstanceWithContext(
                "com.sun.star.graphic.GraphicProvider", xContext));

            XGraphic xGraphic = xGraphicProvider.queryGraphic(sourceProps);

            XPropertySet origProps = duplicateObjectPropertySet(xImageObject);


            origProps.setPropertyValue("Graphic", xGraphic);

        } catch (Exception exception) {
            System.out.println("Couldn't set image properties");
            return 1;
        }

/*
        try {
            // Inserting the content
            // The controller gives us the TextViewCursor
            // query the viewcursor supplier interface
            XController xController = xTextDoc.getCurrentController();
            XTextViewCursorSupplier xViewCursorSupplier = (XTextViewCursorSupplier) UnoRuntime.queryInterface(XTextViewCursorSupplier.class, xController);

            XTextViewCursor xViewCursor = xViewCursorSupplier.getViewCursor();
            XText xText1 = xViewCursor.getText();
            XTextCursor xModelCursor = xText1.createTextCursorByRange(xViewCursor.getStart());
            XTextRange xTextRange1 = xModelCursor.getStart();
            xText1.insertTextContent(xTextRange1, xTextContent, true);
        } catch (Exception exception) {
            System.out.println("Could not insert Content");
            exception.printStackTrace(System.err);
            return 1;
        }
 */
        return 0;
    }

    /*
     * Original is from http://www.oooforum.org/forum/viewtopic.phtml?t=81870
     * Original code was dealing with Text Documents.  This is for Drawing
     * Documents.  See:
     *   http://wiki.services.openoffice.org/wiki/Documentation/DevGuide/Drawings/Navigating
     * which says,
     *
     *    "Initially, shapes in a document can only be accessed by their index.
     *    The only method to get more information about a shape on the page is
     *    to test for the shape type, so it is impossible to identify a
     *    particular shape. However, after a shape is inserted, you can name
     *    it in the user interface or through the shape interface
     *    com.sun.star.container.XNamed, and identify the shape by its name
     *    after retrieving it by index. Shapes cannot be accessed by their names."
     *    Arrgghh!!
     *
     * So for Drawing documents, we need to keep track of index values.  But
     * what if we insert a new image?  Doesn't that change the index values?
     * Need to understand the answer to that!
     *
     * For a test, assume that we're always changing image with index 1.  If
     * there is no image with index 1, then do nothing.
     */
    private static int replacePresentationDocImage(XComponentContext xContext,
                                                   XMultiComponentFactory xMCF,
                                                   XComponent xCompDoc,
                                                   String originalImageName,
                                                   String replacementURL,
                                                   int p,
                                                   int s)
    {

        XDrawPage currPage = null;
        XShape currShape = null;

        byte[] replacementByteArray = GetImageByteStream(replacementURL);
        ByteArrayToXInputStreamAdapter xSource = new ByteArrayToXInputStreamAdapter(replacementByteArray);

        // Query for the XDrawPagesSupplier interface
        XDrawPagesSupplier xDrawPagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
        if (xDrawPagesSuppl == null) {
            System.out.printf("Cannot get XDrawPagesSupplier interface for Presentation Document???\n");
            return(1);
        }

        XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
        int pageCount = xDrawPages.getCount();

        // XXX Avoid these loops, by supplying the page and shape index values???
        try {
            currPage = getDrawPage(xDrawPages, p);
            if (currPage == null) {
                System.out.printf("Failed to get page %d, with index number %d!\n", p+1, p);
                return(2);
            }
            System.out.printf("=== Working with page %d ===\n", p+1);
            int shapeCount = currPage.getCount();
            System.out.printf("Page %d has %d shapes\n", p+1, shapeCount);

            //XShape xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xDrawPages.getByIndex(p));
            currShape = getPageShape(currPage, s);

            String currType = currShape.getShapeType();
            System.out.printf("Working with shape of type: '%s'\n", currType);

            if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
                XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, currShape);
                String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                System.out.printf("The URL for the selected shape is '%s'\n", pictureURL);
            } else {
                System.out.printf("The selected shape is not a GraphicObjectShape!\n");
                return(3);
            }
        } catch (UnknownPropertyException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }

        // Read the source
        PropertyValue[] sourceProps = new PropertyValue[1];
        sourceProps[0] = new PropertyValue();
        sourceProps[0].Name = "InputStream";
        sourceProps[0].Value = xSource;
        try {

            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(
                XGraphicProvider.class,
                xMCF.createInstanceWithContext(
                "com.sun.star.graphic.GraphicProvider", xContext));

            XGraphic xGraphic = xGraphicProvider.queryGraphic(sourceProps);

            XPropertySet origProps = duplicateObjectPropertySet(currShape);

            origProps.setPropertyValue("Graphic", xGraphic);

        } catch (Exception exception) {
            System.out.println("Couldn't set image properties");
            return (4);
        }

        return 0;
    }

    private static void handleDrawDocument(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          XDrawPagesSupplier xDrawPagesSuppl,
                                          String outputDir)
    {
        try {
            XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
            Object firstPage = xDrawPages.getByIndex(0);
            int pageCount = xDrawPages.getCount();
            if (verbose) System.out.printf("xDrawPages.getCount returned a value of '%d' pages\n", pageCount);
            XDrawPage currPage = null;
            Class pageClass = null;
            XPropertySet pageProps = null;

            // Loop through all the pages of the document
            for (int p = 0; p < pageCount; p++) {
                currPage = getDrawPage(xDrawPages, p);
                if (currPage == null) {
                    System.out.printf("Failed to get currPage at page %d!\n", p+1);
                    xCompDoc.dispose();
                    System.exit(22);
                }
                System.out.printf("=== Working with page %d ===\n", p+1);
                exportContextImage(xContext, xMCF, currPage, outputDir, p+1);

                int shapeCount = currPage.getCount();
                if (verbose) System.out.printf("Page %d has %d shapes\n", p+1, shapeCount);
                XShape currShape = null;

                // Loop through all the shapes within the page
                for (int s = 0; s < shapeCount; s++) {
                    currShape = getPageShape(currPage, s);
                    if (currShape == null) {
                        System.out.printf("Failed to get currShape (%d) from page %d!\n", s+1, p+1);
                        xCompDoc.dispose();
                        System.exit(33);
                    }
                    String currType = currShape.getShapeType();
                    com.sun.star.awt.Size shapeSize = currShape.getSize();
                    com.sun.star.awt.Point shapePoint = currShape.getPosition();
                    if (verbose) System.out.printf("--- Working with shape %d (At %d:%d, size %dx%d)\ttype: %s---\n", s + 1, shapePoint.X, shapePoint.Y, shapeSize.Width, shapeSize.Height, currType);

                    //printShapeProperties(currShape);

                    /* Note that we specifically ignore TitleTextShape, OutlinerShape, and LineShape */
                    if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
                        if (verbose) System.out.printf("Handling GraphicObjectShape (%d) on page %d\n", s+1, p+1);
//                        exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                        XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, currShape);
                        String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                        pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                        String outName = constructBaseImageName(outputDir, p+1, s+1);
                        extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.TableShape")) {
                        if (verbose) System.out.printf("Handling TableShape (%d) on page %d\n", s+1, p+1);
                        exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.GroupShape")) {
                        if (verbose) System.out.printf("Handling GroupShape (%d) on page %d\n", s+1, p+1);
                        exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.CustomShape")) {
                        if (!excludeCustomShapes) {
                            if (verbose) System.out.printf("Handling CustomShape (%d) on page %d\n", s+1, p+1);
                            exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                        }
                    } else {
                        System.out.printf("SKIPPING unhandled shape type '%s' (%d) on page %d\n", currType, s+1, p+1);
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (UnknownPropertyException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static void printShapeProperties(XShape shape) throws WrappedTargetException {

        // Get and print all the shape's properties
        XPropertySet xShapeProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, shape);
        Property[] props = xShapeProperties.getPropertySetInfo().getProperties();
        for (int x = 0; x < props.length; x++) {
            try {
                System.out.println("    Property " + props[x].Name + " = " + xShapeProperties.getPropertyValue(props[x].Name));
            } catch (UnknownPropertyException ex) {
                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        System.out.println();

    }

    private static void printSupportedServices(XComponent xCompDoc)
    {
                XServiceInfo xServiceInfo = (XServiceInfo) UnoRuntime.queryInterface(
                    XServiceInfo.class, xCompDoc);
            if (xServiceInfo != null) {
                String[] svcNames = xServiceInfo.getSupportedServiceNames();
                if (svcNames.length > 0)
                    System.out.println("This document supports the following services:");
                for (int i = 0; i < svcNames.length; i++) {
                    System.out.printf("\t%s\n", svcNames[i]);
                }
            }
    }

    private static void exportContextImage(XComponentContext xContext,
                                           XMultiComponentFactory xMCF,
                                           XDrawPage page,
                                           String outputDir,
                                           int p)
    {

        XPropertySet shapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, page);
/*
        if (shapeProps != null) {
            try {
                Object shapeURLObj = shapeProps.getPropertyValue("GraphicURL");
                String shapeURL = shapeURLObj.toString();

            //System.out.printf("The URL for this shape is '%s'\n", shapeURL);
            } catch (Exception e) {
                System.out.printf("Unable to get GraphicURL property for context image: '%s'\n", e.getMessage());
            }
        }
 */
        String fname = String.format("%s/%s-%05d.%s", outputDir, "contextimage", p, "png");

        PropertyValue outProps[] = new PropertyValue[2];
        outProps[0] = new PropertyValue();
        outProps[0].Name = "MediaType";
        outProps[0].Value = "image/png";

        outProps[1] = new PropertyValue();
        outProps[1].Name = "URL";
        outProps[1].Value = "file://" + fname;

        if (verbose) System.out.printf("Exporting page %d to file '%s'\n", p, fname);

        try {

            Object GraphicExportFilter = xMCF.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", xContext);
            XExporter xExporter = (XExporter)UnoRuntime.queryInterface(XExporter.class, GraphicExportFilter);

            XComponent xCompPage = (XComponent)UnoRuntime.queryInterface(XComponent.class, page);
            xExporter.setSourceDocument(xCompPage);

            XFilter xFilter = (XFilter) UnoRuntime.queryInterface(XFilter.class, GraphicExportFilter);
            xFilter.filter(outProps);

        } catch (Exception e) {
            System.out.println("Caught Exception exporting image:" + e.getMessage());
        }
    }


    private static void exportImage(XComponentContext xContext,
                                    XMultiComponentFactory xMCF,
                                    XShape shape,
                                    String outputDir,
                                    int p,
                                    int s)
    {

        XPropertySet shapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, shape);
//        try {
//            Object shapeURLObj = shapeProps.getPropertyValue("GraphicURL");
//            String shapeURL = shapeURLObj.toString();
//
//            //System.out.printf("The URL for this shape is '%s'\n", shapeURL);
//        } catch (Exception e) {
//            System.out.printf("Unable to get GraphicURL property for object: '%s'\n", e.getMessage());
//        }

        String fname = String.format("%s/%s-%05d-%03d.%s", outputDir, "image", p, s, "png");

        PropertyValue outProps[] = new PropertyValue[2];
        outProps[0] = new PropertyValue();

        outProps[0].Name = "MediaType";
        outProps[0].Value = "image/png";
//        outProps[0].Name = "FilterName";
//        outProps[0].Value = "WMF - MS Windows Metafile";

        outProps[1] = new PropertyValue();
        outProps[1].Name = "URL";
        outProps[1].Value = fileNameToOOoURL(fname);

        if (verbose) System.out.printf("Exporting shape %d from page %d to file '%s'\n", s, p, fname);

        try {

            Object GraphicExportFilter = xMCF.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", xContext);
            XExporter xExporter = (XExporter)UnoRuntime.queryInterface(XExporter.class, GraphicExportFilter);

            XComponent xCompShape = (XComponent)UnoRuntime.queryInterface(XComponent.class, shape);
            xExporter.setSourceDocument(xCompShape);

            XFilter xFilter = (XFilter) UnoRuntime.queryInterface(XFilter.class, GraphicExportFilter);
            xFilter.filter(outProps);

        } catch (Exception e) {
            System.out.println("Caught Exception exporting image: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void storeImage(XComponentContext xContext,
                                   XMultiComponentFactory xMCF,
                                   XComponent xCompDoc,
                                   XShape xShape,
                                   String outputDir,
                                   int p,
                                   int s) {
        String newFName = String.format("%s/%s-%05d-%03d.%s", outputDir, "image", p, s, "jpg");
        String newFNameURL = fileNameToOOoURL(newFName);

        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, xShape);
        PropertyValue[] propertyValue = new PropertyValue[3];
        propertyValue[2] = new com.sun.star.beans.PropertyValue();
        propertyValue[2].Name = "Overwrite";
        propertyValue[2].Value = Boolean.valueOf(true);
        propertyValue[0] = new com.sun.star.beans.PropertyValue();
        propertyValue[0].Name = "FilterName";
        propertyValue[0].Value = "";    // "MS Word 97"; //destinationFormats.getFilterName();
//        propertyValue[2] = new PropertyValue();
//        propertyValue[2].Value = "AsTemplate";
//        propertyValue[2].Value = Boolean.valueOf(true);
        XOutputStreamToByteArrayAdapter outputStream = new XOutputStreamToByteArrayAdapter();
        propertyValue[1] = new com.sun.star.beans.PropertyValue();
        propertyValue[1].Name = "OutputStream";
        propertyValue[1].Value = outputStream;
        try {
//          xStorable.storeAsURL(newFNameURL, propertyValue);
            xStorable.storeToURL(newFNameURL, propertyValue);
        } catch (com.sun.star.io.IOException ex) {
            System.out.println(ex.getMessage());
            //ex.printStackTrace();
        } catch (Exception e) {
            System.out.println("Caught Exception storing image: " + e.getMessage());
            //e.printStackTrace();
        }
    }

/*
    private static void exportTextdocEmbeddedGraphics(XMultiComponentFactory xMCF,
            XComponentContext xContext,
            XComponent xCompDoc,
            String outputDir) {
        try {
            Object graphicProviderObject = xMCF.createInstanceWithContext(
                    "com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(XGraphicProvider.class, graphicProviderObject);
            XZipFileAccessor zipFileAccessor = new XZipFileAccessor(xCompDoc, );

        } catch (Exception e) {
            System.out.println("Caught Exception exporting text doc image: " + e.getMessage());
        //e.printStackTrace();
        }
    }
*/

/* (Should not be used!)
    private static void handleCustomShape(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XShape shape,
                                          String outputDir,
                                          int p,
                                          int s)
    {
        try {
            System.out.println("handleCustomShape: currently does nothing!!");
            printShapeProperties(shape);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

*/
    
    private static XDrawPage getTextDrawPage(XDrawPage xDrawPage, int nIndex)
    {
/*
        XDrawPage xDP = null;
        try {
            if ( nIndex < xDrawPage.getCount() )
                xDP = (XDrawPage) UnoRuntime.queryInterface(
                        XDrawPage.class, xDrawPage.getByIndex(nIndex));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return xDP;
        }
 */
        XDrawPage xDP = null;
        try {
            Object oDP = xDrawPage.getByIndex(nIndex);
            xDP = (XDrawPage) UnoRuntime.queryInterface(XDrawPage.class, oDP);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return xDP;
        }
    }

    private static XDrawPage getDrawPage(XDrawPages xDrawPages, int nIndex)
    {
        XDrawPage xDP = null;
        try {
            if ( nIndex < xDrawPages.getCount() )
                xDP = (XDrawPage) UnoRuntime.queryInterface(
                        XDrawPage.class, xDrawPages.getByIndex(nIndex));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return xDP;
        }
    }

    private static XShape getPageShape(XDrawPage xDrawPage, int nIndex)
    {
        XShape xShape = null;
        try {
            if ( nIndex < xDrawPage.getCount() )
                xShape = (XShape) UnoRuntime.queryInterface(
                        XShape.class, xDrawPage.getByIndex(nIndex));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return xShape;
        }
    }

/*
    private static void enumeratePageImages(com.sun.star.lang.XComponent xComp,
            com.sun.star.drawing.XDrawPage xDP) {
        XGalleryThemeProvider xGTP =
                (XGalleryThemeProvider) UnoRuntime.queryInterface(XGalleryThemeProvider.class, xComp);
        if (xGTP == null) {
            System.out.println("Failed to get XGalleryThemeProvider from xDrawPage");
            return;
        }

    }
*/

    private static void printDocumentType(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          String sURL)
    {
        Object oInterface = null;

        try {
            oInterface = xMCF.createInstanceWithContext(
                    "com.sun.star.document.TypeDetection", xContext);
        } catch (java.lang.Exception ex) {
            ex.printStackTrace();
        }

        if(oInterface == null) {
              System.err.println("__FUNCTION__: unable to create TypeDetection service");
        }

        XTypeDetection m_xDetection = (XTypeDetection)
                UnoRuntime.queryInterface(XTypeDetection.class, oInterface);

        // queryTypeByURL does a "flat" detection of filetype (looking at the suffix???)
        System.out.println("queryTypeByURL says '"
                + sURL + "' is of type: '" +
                m_xDetection.queryTypeByURL(sURL) +"'");

        // queryTypeByDescriptor does an optional "deep" detection of filetype
        PropertyValue testProps[][] = new PropertyValue[1][1];
        testProps[0][0] = new PropertyValue();
        testProps[0][0].Name = "URL";
        testProps[0][0].Value = sURL;

        System.out.println("queryTypeByDescriptor says '" +
                sURL + "' is of type: '" +
                m_xDetection.queryTypeByDescriptor(testProps, false) +"'");

    }

    private static String getDocumentType(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          String sURL)
    {
        Object oInterface = null;

        try {
            oInterface = xMCF.createInstanceWithContext(
                    "com.sun.star.document.TypeDetection", xContext);
        } catch (java.lang.Exception ex) {
            ex.printStackTrace();
        }

        if(oInterface == null) {
              System.err.println("__FUNCTION__: unable to create TypeDetection service");
              return null;
        }

        XTypeDetection m_xDetection = (XTypeDetection)
                UnoRuntime.queryInterface(XTypeDetection.class, oInterface);

        // do the optional "deep" detection of filetype
        PropertyValue testProps[][] = new PropertyValue[1][1];
        testProps[0][0] = new PropertyValue();
        testProps[0][0].Name = "URL";
        testProps[0][0].Value = sURL;

        return m_xDetection.queryTypeByDescriptor(testProps, true);
    }

    private static String fileNameToOOoURL(final String fName) {
        StringBuilder sLoadUrl = new StringBuilder("file://");
        sLoadUrl.append(fName.replace('\\', '/'));
        return sLoadUrl.toString();
    }

    private static void storeDocument(XComponentContext xContext,
                                      XMultiComponentFactory xMCF,
                                      XComponent xCompDoc,
                                      String newFName,
                                      String filterName)
    {
        String newFNameURL = fileNameToOOoURL(newFName);
        if (verbose) System.out.printf("Storing to '%s' using URL '%s'\n", newFName, newFNameURL);

        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(XStorable.class, xCompDoc);
        PropertyValue[] propertyValue = new PropertyValue[3];
        propertyValue[2] = new com.sun.star.beans.PropertyValue();
        propertyValue[2].Name = "Overwrite";
        propertyValue[2].Value = Boolean.valueOf(true);
        propertyValue[0] = new com.sun.star.beans.PropertyValue();
        propertyValue[0].Name = "FilterName";
        propertyValue[0].Value = filterName; // "MS Word 97"; // "MOOX"; //destinationFormats.getFilterName();
//        propertyValue[2] = new PropertyValue();
//        propertyValue[2].Value = "AsTemplate";
//        propertyValue[2].Value = Boolean.valueOf(true);
        XOutputStreamToByteArrayAdapter outputStream = new XOutputStreamToByteArrayAdapter();
        propertyValue[1] = new com.sun.star.beans.PropertyValue();
        propertyValue[1].Name = "OutputStream";
        propertyValue[1].Value = outputStream;
        try {
//          xStorable.storeAsURL(newFNameURL, propertyValue);
            xStorable.storeToURL(newFNameURL, propertyValue);
        } catch (com.sun.star.io.IOException ex) {
            System.out.println("Storing document: " + ex.getMessage());
            ex.printStackTrace();
        }

    }

    private static boolean isaSupportedFormat(OOoFormat format)
    {
        return (java.util.Arrays.asList(SUPPORTED_FORMATS).contains(format));
    }

    private static OOoFormat getNativeFormat(OOoFormat origFormat)
    {
        OOoFormat nativeFormat = origFormat;

        if (origFormat.getHandlerType() == 0) {
            nativeFormat = OOoFormat.OpenDocument_Text;
        } else if (origFormat.getHandlerType() == 2) {
            nativeFormat = OOoFormat.OpenDocument_Presentation;
        }
        return nativeFormat;
    }

    private static String possiblyUseTemporaryDocument(XComponentContext xContext,
                                               XMultiComponentFactory xMCF,
                                               XComponent xCompOrigDoc,
                                               String origName,
                                               OOoFormat origFormat)
    {
        String newName;
        OOoFormat nativeFormat;

        // If already in Native OO format, no need to continue
        if ((nativeFormat = getNativeFormat(origFormat)) == null) {
            return null;
        }
        if (nativeFormat == origFormat) {
            return null;
        }

        // XXX Need proper code to generate a random name!!!
        newName = new String("/tmp/foobar_thisshouldberandom_" + "xyz123" + "." + nativeFormat.getFileExtension());

        // Save in OO format and return the name of the temporary file
        storeDocument(xContext, xMCF, xCompOrigDoc, newName, nativeFormat.getFilterName());
        // Get current file type and determine if we need to save it in OO format

        return newName;
    }

    private static String getExtension(String fullName)
    {
        int dotPos = fullName.lastIndexOf(".");
        String strExtension = fullName.substring(dotPos + 1);
        return strExtension;
        //String strFilename = fullName.substring(0, dotPos);
    }

    private static String constructBaseImageName(String dir, int page, int num)
    {
        String outName = new String();

        outName = String.format("%s/image-%05d-%03d", dir, page, num);
        return outName;
    }

    private static void extractImageByURL(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String pictureURL,
                                          String outName)
    {
        try {
            Object oSFAcc = xMCF.createInstanceWithContext(
                    "com.sun.star.ucb.SimpleFileAccess", xContext);
            XSimpleFileAccess2 xFileWriter = (XSimpleFileAccess2)
                    UnoRuntime.queryInterface(XSimpleFileAccess2.class, oSFAcc);
            XStorageBasedDocument xStorageBasedDocument = (XStorageBasedDocument)
                    UnoRuntime.queryInterface(XStorageBasedDocument.class, xCompDoc);

            XStorage xDocStorage = xStorageBasedDocument.getDocumentStorage();
            XStorage xDocPictures = (XStorage) UnoRuntime.queryInterface(
                    XStorage.class, xDocStorage.getByName("Pictures"));

            XNameAccess xDocStorageNameAccess = (XNameAccess)
                    UnoRuntime.queryInterface(XNameAccess.class, xDocStorage);

            String[] allNames = xDocStorageNameAccess.getElementNames();
//            for (int i = 0; i < allNames.length; i++) {
//                System.out.printf("The big list has name '%s'\n", allNames[i]);
//            }

            if (!xDocStorageNameAccess.hasByName("Pictures")) {
                System.out.printf("Found no \"Pictures\" in the document!!!\n");
                return;
            }

            Object oPicturesStorage = xDocStorageNameAccess.getByName("Pictures");
            XNameAccess xPicturesNameAccess = (XNameAccess)
                    UnoRuntime.queryInterface(XNameAccess.class, oPicturesStorage);

            String[] aNames = xPicturesNameAccess.getElementNames();
            if (verbose) System.out.printf("There were a total of %d pictures found via DocStorageAccess\n", aNames.length);
            for (int i = 0; i < aNames.length; i++) {
                //System.out.printf("Picture %d has name '%s'\n", i+1, aNames[i]);
                //System.out.printf("Processing picture with name '%s'\n", aNames[i]);
                if (aNames[i].contains(pictureURL)) {
                    Object oElement = xPicturesNameAccess.getByName(aNames[i]);
                    XStream xStream = (XStream) UnoRuntime.queryInterface(XStream.class, oElement);
                    Object oInput = xStream.getInputStream();
                    XInputStream xInputStream = (XInputStream)
                            UnoRuntime.queryInterface(XInputStream.class, oInput);

                    xFileWriter.writeFile(outName + "." + getExtension(aNames[i]), xInputStream);
                    break;
                }

            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            return;
        }

    }
}
