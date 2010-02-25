/*
 * OpenOfficeUNODecomposition.java
 *
 * Created on 2009.08.17 - 10:37:43
 *
 */

package edu.umich.med.umms;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
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

    private static final String Usage =
            "[ --extract <options> | --replace <options> | --copy <options> ]\n\n" +
            "Global options: [--verbose]\n\n" +
            "--json <file>\n\n" +
            "--extract [ --exclude-custom-shapes ] --input <file> --output-dir <dir>\n\n" +
            "--replace <options> --input <file> --newimage <file> --pagenum <pagenum> --imagenum <imagenum>\n\n" +
            "--cite <options> --input <file> --citeimage <file> --pagenum <pagenum> --imagenum <imagenum>\n\n" +
            "--copy <options> --input <file> --output <file>\n";

    /*private static boolean functionImageReplacement = false;
    private static boolean functionCitationInsertion = false;
    private static boolean functionExtractImages = false;
    private static boolean functionCopyFile = false;
    private static boolean functionJsonCommands = false;*/

    private static com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();

    


    /** Creates a new instance of OpenOfficeUNODecomposition */
    public OpenOfficeUNODecomposition() {
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)
    {
        XComponentContext xContext = null;
        XMultiComponentFactory xMCF = null;
        XDesktop xDesktop = null;
        Options opt = new Options();
        DecompParameters dp = new DecompParameters();
        int err = 1;


        defineOptions(opt);
        err = processArguments(opt, args, dp);
        if (err != 0) {
            System.exit(1);
        }

        // get the remote office component context
        try {
            xContext = Bootstrap.bootstrap();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            mylog.error("Error connecting to OpenOffice process: " + e.getMessage());
            System.exit(3);
        }

        // get the remote office service manager
        xMCF = xContext.getServiceManager();


        try {
            /* A desktop environment contains tasks with one or more
            frames in which components can be loaded. Desktop is the
            environment for components which can instantiate within
            frames. */
            xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class,
                    xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",
                    xContext));
        } catch (com.sun.star.uno.Exception e) {
            mylog.error("Error getting OpenOffice desktop: " + e.getMessage());
            System.exit(4);
        }

        if (dp.operation == DecompOperation.JSON) {
            err = processJsonFile(xContext, xDesktop, dp);
        } else {
            DecompFileProcessor fp = getFileProcessor(xContext, xDesktop, dp.inputFile);
            if (fp != null) {
                err = processFileOperation(xContext, xDesktop, fp, dp);
            }
        }

        System.exit(err);
    }

    private static DecompFileProcessor getFileProcessor(XComponentContext xContext, XDesktop xDesktop, String fileName)
    {
        DecompFileProcessor fp = null;
        try {
            fp = new DecompFileProcessor(xContext, xDesktop, fileName);
        } catch (Exception e) {
            mylog.error("Error instantiating input file '%s' for processing: %s", fileName, e.getMessage());
            // System.exit(5);
        }
        return fp;
    }

    private static void saveFile(DecompFileProcessor fp, DecompParameters dp)
    {
        fp.saveTo(dp.outputFile);
    }


    private static void disposeFile(DecompFileProcessor fp)
    {
        fp.dispose();
    }

    private static int processFileOperation(XComponentContext xContext, XDesktop xDesktop, DecompFileProcessor fp, DecompParameters dp)
    {
        int ret = 1;
        try {
            switch (dp.operation) {
                case EXTRACT:
                    ret = fp.extractImages(dp.outputDir);
                    break;
                case REPLACE:
                    ret = fp.replaceImage(dp.repImageFile, dp.pageNum, dp.imageNum);
                    break;
                case CITE:
                    ret = fp.citeImage(dp.citationText, dp.citationImageFile, dp.pageNum, dp.imageNum);
                    break;
                case SAVE:
                    ret = fp.saveTo(dp.outputFile);
                    break;
            }
        } catch (java.lang.Exception e) {
            mylog.error("Error processing operation on file '%s': %s", dp.inputFile, e.getMessage());
        } finally {
            return ret;
        }
    }

    private static int processSingleOperation(XComponentContext xContext, XDesktop xDesktop, DecompParameters dp) {
        int ret = 1;
        DecompFileProcessor fp = getFileProcessor(xContext, xDesktop, dp.inputFile);
        if (dp != null) {
            ret = processFileOperation(xContext, xDesktop, fp, dp);
            saveFile(fp, dp);
            disposeFile(fp);
        }
        return ret;
    }

    private static int processJsonFile(XComponentContext xContext, XDesktop xDesktop, DecompParameters dp)
    {
        DecompJson dj = null;
        int ret = 1;
        try {
            dj = new DecompJson(dp.jsonCommandFile);
//            dj.jsonFromFile(dp.jsonCommandFile);
            mylog.debug("This is just a dummy message");
        } catch (IOException ex) {
            mylog.error("IOException while processing input JSON from file '%s': %s", dp.jsonCommandFile, ex.getMessage());
            return ret;
        } catch (java.lang.Exception e) {
            mylog.error("Exception while processing input JSON from file '%s': %s", dp.jsonCommandFile, e.getMessage());
            e.printStackTrace();
            return ret;
        }

        // Now process the commands...
        switch (dj.getVersion()) {
            case 1:
                for (int i = 0; i < dj.getNumFiles(); i++) {
                    DecompFileProcessor fp = getFileProcessor(xContext, xDesktop, dj.getFile(i).getInputFile());
                    if (fp != null) {
                        for (int op = 0; op < dj.getFile(i).getNumOps(); op++) {
                            DecompParameters ldp = new DecompParameters(dj.getFile(i), op);
                            ret = processFileOperation(xContext, xDesktop, fp, ldp);
                            if (ret != 0)
                                break; // XXX Really???
                        }
                        disposeFile(fp);
                    }
                }
                break;
            default:
                mylog.error("Recieved unsupported JSON version number: %d", dj.getVersion());
                ret = 2;
        }
        return ret;
    }

    private static void defineOptions(Options opt)
    {
        OptionGroup mainopts = new OptionGroup();
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("extract").create("e"));
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("replace").create("r"));
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("cite").create("t"));
        mainopts.addOption(OptionBuilder.hasArg(true).withLongOpt("json").create("j"));
        mainopts.addOption(OptionBuilder.hasArg(false).withLongOpt("copy").create("c"));
        opt.addOptionGroup(mainopts);

        opt.addOption("h", "help", false, "Print this usage information");
        opt.addOption("v", "verbose", false, "Print verbose output information");
        opt.addOption("xc", "exclude-custom-shapes", false, "Do not include Custom shapes when extracting");

        opt.addOption("i", "input", true, "Input file name (full path)");
        opt.addOption("o", "output", true, "Output file name (full path)");
        opt.addOption("n", "newimage", true, "Replacement (new) image file (full path)");
        opt.addOption("ti", "citeimage", true, "Image file containing citation information (full path)");
        opt.addOption("te", "citetext", true, "Full citation text");
        opt.addOption("d", "output-dir", true, "Name of directory to receive output");
        opt.addOption("m", "commandlist", true, "List of commands");

        opt.addOption("p", "pagenum", true, "Page number");
        opt.addOption("g", "imagenum", true, "Image number");
    }

    private static int processArguments(Options opt, String[] args, DecompParameters dp)
    {
        CommandLine cl;
        try {
            BasicParser parser = new BasicParser();
            cl = parser.parse(opt, args);
        } catch (ParseException ex) {
            mylog.error("Error processing arguments: '%s'\n", ex.getMessage());
            return 1;
        }

        if (cl.hasOption("v")) {
            DecompUtil.setVerbosity(true);
        }

        if (cl.hasOption("h")) {
            printUsage(opt);
            return 8;
        }

        if (!cl.hasOption("e") && !cl.hasOption("r") && !cl.hasOption("c")
                && !cl.hasOption("t") && !cl.hasOption("j")) {
            mylog.error("You must choose a main function\n");
            printUsage(opt);
            return 2;
        }

        /* JSON Commands */
        if (cl.hasOption("j")) {
            dp.operation = DecompOperation.JSON;
            dp.jsonCommandFile = cl.getOptionValue("j");
        }

        /* extract */
        if (cl.hasOption("e")) {
            if (!cl.hasOption("i") || !cl.hasOption("d")) {
                mylog.error("For extract, you must specify the input file and output directory.\n");
                printUsage(opt);
                return 3;
            }
            dp.operation = DecompOperation.EXTRACT;
            dp.inputFile = cl.getOptionValue("i");
            dp.outputDir = cl.getOptionValue("d");

            if (cl.hasOption("xc")) {
                DecompUtil.excludeCustomShapes();
            }
        }

        /* replace */
        if (cl.hasOption("r")) {
            if (!cl.hasOption("i") || !cl.hasOption("o") || !cl.hasOption("n") || !cl.hasOption("p") || !cl.hasOption("g")) {
                mylog.error("For replace, you must specify input, output, and new image files, as well as page and shape numbers.\n");
                printUsage(opt);
                return 4;
            }
            dp.operation = DecompOperation.REPLACE;
            dp.inputFile = cl.getOptionValue("i");
            dp.outputFile = cl.getOptionValue("o");
            dp.repImageFile = cl.getOptionValue("n");
            dp.pageNum = Integer.parseInt(cl.getOptionValue("p"));
            dp.imageNum = Integer.parseInt(cl.getOptionValue("g"));
            
            // Convert page and image numbers to index values
            if (dp.pageNum <= 0) {
                System.out.printf("Page number of '%d' is invalid", dp.pageNum);
                printUsage(opt);
                return 5;
            } else {
                dp.pageNum -= 1;
            }
            if (dp.imageNum <= 0) {
                System.out.printf("Image number of '%d' is invalid", dp.imageNum);
                printUsage(opt);
                return 6;
            } else {
                dp.imageNum -= 1;
            }
        }

        /* citation */
        if (cl.hasOption("t")) {
            if (!cl.hasOption("i") || !cl.hasOption("o") || !cl.hasOption("te") || !cl.hasOption("ti") || !cl.hasOption("p") || !cl.hasOption("g")) {
                System.out.printf("For replace, you must specify input, output, and citation image files, as well as the citation text and page and shape numbers.\n");
                printUsage(opt);
                return 4;
            }
            dp.operation = DecompOperation.CITE;
            dp.inputFile = cl.getOptionValue("i");
            dp.outputFile = cl.getOptionValue("o");
            dp.citationImageFile = cl.getOptionValue("ti");
            dp.citationText = cl.getOptionValue("te");
            dp.pageNum = Integer.parseInt(cl.getOptionValue("p"));
            dp.imageNum = Integer.parseInt(cl.getOptionValue("g"));

            // Convert page and image numbers to index values
            if (dp.pageNum <= 0) {
                System.out.printf("Page number of '%d' is invalid", dp.pageNum);
                printUsage(opt);
                return 5;
            } else {
                dp.pageNum -= 1;
            }
            if (dp.imageNum <= 0) {
                System.out.printf("Image number of '%d' is invalid", dp.imageNum);
                printUsage(opt);
                return 6;
            } else {
                dp.imageNum -= 1;
            }
        }

        /* copy */
        if (cl.hasOption("c")) {
            dp.operation = DecompOperation.SAVE;
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
