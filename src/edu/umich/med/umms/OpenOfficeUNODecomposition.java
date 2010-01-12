/*
 * OpenOfficeUNODecomposition.java
 *
 * Created on 2009.08.17 - 10:37:43
 *
 */

package edu.umich.med.umms;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.OptionGroup;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.ParseException;

import com.sun.star.comp.helper.Bootstrap;

import com.sun.star.frame.XDesktop;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XComponent;

import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 *
 * @author kwc@umich.edu
 */
public class OpenOfficeUNODecomposition {

    public static final OOoFormat[] SUPPORTED_FORMATS = {
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

//    private static boolean verbose = false;
//    private static boolean excludeCustomShapes = false;

    private static boolean functionImageReplacement = false;
    private static boolean functionCitationInsertion = false;
    private static boolean functionExtractImages = false;
    private static boolean functionCopyFile = false;
    
    private static final String Usage =
            "[ --extract <options> | --replace <options> | --copy <options> ]\n\n" +
            "Global options: [--verbose]\n\n" +
            "--extract [ --exclude-custom-shapes ] --input <file> --output-dir <dir>\n\n" +
            "--replace <options> --input <file> --newimage <file> --pagenum <pagenum> --imagenum <imagenum>\n\n" +
            "--cite <options> --input <file> --citeimage <file> --pagenum <pagenum> --imagenum <imagenum>\n\n" +
            "--copy <options> --input <file> --output <file>\n";
    
    private static String inputName;
    private static String outputName;
    private static String outputDir;

    private static int repPageNum = -1;
    private static int repImageNum = -1;
    private static String repImageFile;
    private static String citeImageFile;

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
            if (DecompUtil.beingVerbose()) System.out.println("Successfully connected to OpenOffice process!");

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

            String sFileUrl = DecompUtil.fileNameToOOoURL(inputName);

            XComponent xCompDoc = DecompUtil.openFileForProcessing(xDesktop, inputName);
            if (xCompDoc == null) {
                System.out.printf("Unable to open original file '%s', aborting!\n", inputName);
                System.exit(3);
            }
            if (DecompUtil.beingVerbose()) DecompUtil.printSupportedServices(xCompDoc);  // original

            // Determine what kind of document it is and how to proceed
            String fileType = DecompUtil.getDocumentType(xContext, xMCF, sFileUrl);
            origFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
            if (!DecompUtil.isaSupportedFormat(origFileFormat)) {
                System.out.printf("File '%s', format '%s', is not supported!\n",
                        inputName, (origFileFormat != null) ? origFileFormat.getFormatName(): "<unknown>");
                xCompDoc.dispose();
                System.exit(5);
            }

            if (functionExtractImages) {
                // Possibly save original document as an OO document
                String newName = DecompUtil.possiblyUseTemporaryDocument(xContext, xMCF,
                        xCompDoc, inputName, origFileFormat);
                if (newName != null) {
                    xCompDoc.dispose();
                    inputName = newName;
                    xCompDoc = DecompUtil.openFileForProcessing(xDesktop, newName);
                    if (xCompDoc == null) {
                        System.out.printf("Unable to open temporary file '%s', aborting!\n", newName);
                        System.exit(4);
                    }
                    fileType = DecompUtil.getDocumentType(xContext, xMCF, DecompUtil.fileNameToOOoURL(inputName));
                    ooFileFormat = OOoFormat.findFormatWithDocumentType(fileType);
                }
                if (DecompUtil.beingVerbose()) DecompUtil.printSupportedServices(xCompDoc);   // possibly OO version
            }

            processFileFormat = (ooFileFormat == null) ? origFileFormat : ooFileFormat;
            // Decide how to process the document
            int handler = processFileFormat.getHandlerType();
            switch (handler) {
                case 0:
                    if (functionImageReplacement) {
                        DecompText.replaceTextDocImage(xContext, xMCF, xCompDoc, "xxxorigname", DecompUtil.fileNameToOOoURL(repImageFile));
                    } else if (functionCitationInsertion) {
                        DecompText.insertTextDocImageCitation(xContext, xMCF, xCompDoc, "xxxorigname", DecompUtil.fileNameToOOoURL(citeImageFile));
                    } else if (functionExtractImages) {
                        DecompText.handleDocument(xContext, xMCF, xCompDoc, outputDir);
                    }
                    break;
                case 2:
                    if (functionImageReplacement) {
                        DecompImpress.replacePresentationDocImage(xContext, xMCF, xCompDoc, "origname", DecompUtil.fileNameToOOoURL(repImageFile), repPageNum, repImageNum);
                    } else if (functionCitationInsertion) {
                        DecompImpress.insertPresentationDocImageCitation(xContext, xMCF, xCompDoc, "origname", DecompUtil.fileNameToOOoURL(citeImageFile), repPageNum, repImageNum);
                    } else if (functionExtractImages) {
                        DecompImpress.handleDocument(xContext, xMCF, xCompDoc, outputDir);
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
                if (DecompUtil.beingVerbose())
                    System.out.printf("Saving (possibly) modified document to a new file, '%s'\n", outputName);
                DecompUtil.storeDocument(xContext, xMCF, xCompDoc, outputName, origFileFormat.getFilterName());
            }
            if (DecompUtil.beingVerbose()) System.out.println("We be done.");
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
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("cite").create("t"));
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
        opt.addOption("ti", "citeimage", true, "Image file containing citation information (full path)");
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
            DecompUtil.setVerbosity(true);
        }

        if (cl.hasOption("h")) {
            printUsage(opt);
            return 8;
        }

        if (!cl.hasOption("e") && !cl.hasOption("r") && !cl.hasOption("c") && !cl.hasOption("t")) {
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
                DecompUtil.excludeCustomShapes();
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

        /* citation */
        if (cl.hasOption("t")) {
            functionCitationInsertion = true;
            if (!cl.hasOption("i") || !cl.hasOption("o") || !cl.hasOption("ti") || !cl.hasOption("p") || !cl.hasOption("g")) {
                System.out.printf("For replace, you must specify input, output, and citation image files, as well as page and shape numbers.\n");
                printUsage(opt);
                return 4;
            }
            inputName = cl.getOptionValue("i");
            outputName = cl.getOptionValue("o");
            citeImageFile = cl.getOptionValue("ti");
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

}
