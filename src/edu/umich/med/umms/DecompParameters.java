package edu.umich.med.umms;

import edu.umich.med.umms.DecompJson.DecompFile;

/**
 *
 * @author kwc
 */
public class DecompParameters {
    private DecompOperation operation;
    private String inputFile;
    private String outputFile;
    private String outputDir;

    private String repImageFile;
    private String citationImageFile;
    private String citationText;
    private String jsonCommandFile;
    private String jsonResultFile;
    private int pageNum;
    private int imageNum;

    private boolean excludeCustomShapes = false;
    private String validationErrorMessage;
    private boolean indexesAdjusted = false;
    private org.apache.log4j.Level myLogLevel;

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

        validationErrorMessage = null;
        indexesAdjusted = false;
        myLogLevel = org.apache.log4j.Level.WARN;
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
        myLogLevel = org.apache.log4j.Level.WARN;
        excludeCustomShapes = df.getExcludeCustomShapes(opnum);

        validationErrorMessage = null;
        indexesAdjusted = true;
    }

    public void setOperation(DecompOperation op)
    {
        this.operation = op;
    }
    public void setInputFile(String f)
    {
        this.inputFile = f;
    }
    public void setOutputFile(String f)
    {
        this.outputFile = f;
    }
    public void setOutputDir(String f)
    {
        this.outputDir = f;
    }
    public void setRepImageFile(String f)
    {
        this.repImageFile = f;
    }
    public void setCitationImageFile(String f)
    {
        this.citationImageFile = f;
    }
    public void setCitationText(String f)
    {
        this.citationText = f;
    }
    public void setJsonCommandFile(String f)
    {
        this.jsonCommandFile = f;
    }
    public void setJsonResultFile(String f)
    {
        this.jsonResultFile = f;
    }
    public void setPageNum(int p)
    {
        this.pageNum = p;
    }
    public void setImageNum(int i)
    {
        this.imageNum = i;
    }
    public void setExcludeCustomShapes(boolean b)
    {
        this.excludeCustomShapes = b;
    }

    public boolean isValidLoggingLevel(String levelDesired)
    {
        if (levelDesired.equalsIgnoreCase("all")   ||
            levelDesired.equalsIgnoreCase("debug") ||
            levelDesired.equalsIgnoreCase("info")  ||
            levelDesired.equalsIgnoreCase("warn")  ||
            levelDesired.equalsIgnoreCase("error") ||
            levelDesired.equalsIgnoreCase("fatal") ||
            levelDesired.equalsIgnoreCase("all"))
            return true;
        else
            return false;
    }

    public boolean setLoggingLevel(String levelDesired)
    {
        if (levelDesired.equalsIgnoreCase("all")) {
            this.myLogLevel = org.apache.log4j.Level.ALL;
        } else if (levelDesired.equalsIgnoreCase("debug")) {
            this.myLogLevel = org.apache.log4j.Level.DEBUG;
        } else if (levelDesired.equalsIgnoreCase("info")) {
            this.myLogLevel = org.apache.log4j.Level.INFO;
        } else if (levelDesired.equalsIgnoreCase("warn")) {
            this.myLogLevel = org.apache.log4j.Level.WARN;
        } else if (levelDesired.equalsIgnoreCase("error")) {
            this.myLogLevel = org.apache.log4j.Level.ERROR;
        } else if (levelDesired.equalsIgnoreCase("fatal")) {
            this.myLogLevel = org.apache.log4j.Level.FATAL;
        } else if (levelDesired.equalsIgnoreCase("all")) {
            this.myLogLevel = org.apache.log4j.Level.ALL;
        } else {
            return false;
        }
        return true;
    }

    
    public DecompOperation getOperation()
    {
        return this.operation;
    }
    public String getInputFile()
    {
        return this.inputFile;
    }
    public String getOutputFile()
    {
        return this.outputFile;
    }
    public String getOutputDir()
    {
        return this.outputDir;
    }
    public String getRepImageFile()
    {
        return this.repImageFile;
    }
    public String getCitationImageFile()
    {
        return this.citationImageFile;
    }
    public String getCitationText()
    {
        return this.citationText;
    }
    public String getJsonCommandFile()
    {
        return this.jsonCommandFile;
    }
    public String getJsonResultFile()
    {
        return this.jsonResultFile;
    }
    public int getPageNum()
    {
        return this.pageNum;
    }
    public int getImageNum()
    {
        return this.imageNum;
    }
    public boolean getExcludeCustomShapes()
    {
        return this.excludeCustomShapes;
    }
    public String getValidationErrorMessage()
    {
        return this.validationErrorMessage;
    }
    public org.apache.log4j.Level getLoggingLevel()
    {
        return this.myLogLevel;
    }

    private boolean validIndexes()
    {
        int rc;

        if (!this.indexesAdjusted) {
            adjustIndexes();
        }
        if (this.pageNum < 0 || this.imageNum < 0) {
            this.validationErrorMessage = new String("Page or image value is invalid");
            return false;
        }

        return true;
    }

    public void adjustIndexes()
    {
        if (this.indexesAdjusted)
            return;
        this.pageNum--;
        this.imageNum--;
        this.indexesAdjusted = true;
    }

    public boolean ValidSingleOp()
    {
        if (this.operation == null) {
            this.validationErrorMessage = new String("No operation was specified.");
            return false;
        }
        switch (this.operation) {
            case EXTRACT:
                if (this.inputFile == null || this.outputDir == null) {
                    this.validationErrorMessage = new String("For extract, you must specify the input file and output directory.");
                    return false;
                }
                break;

            case REPLACE:
                if (this.inputFile == null || this.repImageFile == null || this.pageNum < 0 || this.imageNum < 0) {
                    this.validationErrorMessage = new String("For replace, you must specify input, output, and new image files, as well as page and shape numbers.");
                    return false;
                }
                if (!validIndexes())
                    return false;
                break;

            case CITE:
                if (this.inputFile == null || this.citationText == null || this.pageNum < 0 || this.imageNum < 0) {
                    this.validationErrorMessage = new String("For citation, you must specify input and output files, as well as the citation text and page and shape numbers.");
                    return false;
                }
                if (!validIndexes())
                    return false;
                break;

            case SAVE:
                if (this.inputFile == null || this.outputFile == null ) {
                    this.validationErrorMessage = new String("For copy(save), you must specify input and output file names");
                    return false;
                }
                break;

            case JSON:
                if (this.jsonCommandFile == null) {
                    this.validationErrorMessage = new String("For json, you must specify the json file name.");
                    return false;
                }
                break;

            default:
                this.validationErrorMessage = new String("Invalid operation value");
                return false;
        }
        return true;
    }

    public boolean Valid()
    {
        if (!ValidSingleOp())
            return false;
        switch (this.operation) {

            case REPLACE:
                if (this.outputFile == null) {
                    this.validationErrorMessage = new String("For replace, you must specify input, output, and new image files, as well as page and shape numbers.");
                    return false;
                }
                break;

            case CITE:
                if (this.outputFile == null) {
                    this.validationErrorMessage = new String("For citation, you must specify input and output files, as well as the citation text and page and shape numbers.");
                    return false;
                }
                break;
        }
        return true;
    }

    @Override
    public String toString()
    {
        StringBuffer sb = new StringBuffer();
        sb.append("\n----- DecompParameters -----\n");
        sb.append("operation:           " + this.operation.toString() + "\n");
        sb.append("inputFile:           " + this.inputFile + "\n");
        sb.append("outputFile:          " + this.outputFile + "\n");
        sb.append("outputDir:           " + this.outputDir + "\n");
        sb.append("repImageFile:        " + this.repImageFile + "\n");
        sb.append("citationImageFile:   " + this.citationImageFile + "\n");
        sb.append("citationText:        " + this.citationText + "\n");
        sb.append("jsonCommandFile:     " + this.jsonCommandFile + "\n");
        sb.append("jsonResultFile:      " + this.jsonResultFile + "\n");
        sb.append("pageNum:             " + this.pageNum + "\n");
        sb.append("imageNum:            " + this.imageNum + "\n");
        sb.append("excludeCustomShapes: " + this.excludeCustomShapes + "\n");
        return sb.toString();
    }
}