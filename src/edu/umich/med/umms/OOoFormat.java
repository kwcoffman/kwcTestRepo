package edu.umich.med.umms;

public enum OOoFormat {

    OpenDocument_Text("OpenDocument Text","writer8","odt","application/vnd.oasis.opendocument.text","writer8",0),
    OpenDocument_Text_Template("OpenDocument Text Template","writer8_template","ott","application/vnd.oasis.opendocument.text-template","writer8_template",0),
    OpenOfficeorg_10_Text_Document("OpenOffice.org 1.0 Text Document","StarOffice XML (Writer)","sxw","application/vnd.sun.xml.writer","writer_StarOffice_XML_Writer",0),
    OpenOfficeorg_10_Text_Document_Template("OpenOffice.org 1.0 Text Document Template","writer_StarOffice_XML_Writer_Template","stw","application/vnd.sun.xml.writer.template","writer_web_StarOffice_XML_Writer_Web_Template",0),
    Microsoft_Word_97_2000_XP("Microsoft Word 97/2000/XP","MS Word 97","doc","application/msword","writer_MS_Word_97",0),
    Microsoft_Word_95("Microsoft Word 95","MS Word 95","doc","application/msword","writer_MS_Word_95",0),
    Microsoft_Word_60("Microsoft Word 6.0","MS WinWord 6.0","doc","application/msword","writer_MS_WinWord60",0),
    Rich_Text_Format("Rich Text Format","Rich Text Format","rtf","application/rtf","writer_Rich_Text_Format",0),
    StarWriter_50("StarWriter 5.0","StarWriter 5.0","sdw","application/vnd.staroffice.writer","writer_StarWriter_50",0),
    StarWriter_50_Template("StarWriter 5.0 Template","StarWriter 5.0 Vorlage/Template","vor","application/vnd.stardivision.writer","writer_web_StarWriterWeb_50_VorlageTemplate",0),
    StarWriter_40("StarWriter 4.0","StarWriter 4.0","sdw","application/vnd.staroffice.writer","writer_StarWriter_40",0),
    StarWriter_40_Template("StarWriter 4.0 Template","StarWriter 4.0 Vorlage/Template","vor","application/vnd.stardivision.writer","writer_web_StarWriterWeb_40_VorlageTemplate",0),
    StarWriter_30("StarWriter 3.0","StarWriter 3.0","sdw","application/vnd.staroffice.writer","writer_StarWriter_30",0),
    StarWriter_30_Template("StarWriter 3.0 Template","StarWriter 3.0 Vorlage/Template","vor","application/vnd.stardivision.writer","writer_web_StarWriterWeb_30_VorlageTemplate",0),
    Text("Text","Text","txt","text/plain","writer_Text",0),
    Text_Encoded("Text Encoded","Text (encoded)","txt","text/plain","writer_Text_encoded",0),
    HTML_Document_OpenOfficeorg_Writer("HTML Document (OpenOffice.org Writer)","HTML (StarWriter)","html","text/html","writer_web_HTML",0),
    AportisDoc_Palm("AportisDoc (Palm)","AportisDoc Palm DB","pdb","application/octet-stream","writer_AportisDoc_PalmDB_File",0),
    DocBook("DocBook","DocBook File","xml","application/docbook+xml","writer_DocBook_File",0),
    Microsoft_Word_2003_XML("Microsoft Word 2003 XML","MS Word 2003 XML","xml","application/xml","writer_MS_Word_2003_XML",0),
    Pocket_Word("Pocket Word","PocketWord File","psw","application/octet-stream","writer_PocketWord_File",0),

//    OpenDocument_Spreadsheet("OpenDocument Spreadsheet","calc8","ods","application/x-vnd.oasis.opendocument.spreadsheet",1),
//    OpenDocument_Spreadsheet_Template("OpenDocument Spreadsheet Template","calc8_template","ots","application/vnd.oasis.opendocument.spreadsheet-template",1),
//    OpenOfficeorg_10_Spreadsheet("OpenOffice.org 1.0 Spreadsheet","StarOffice XML (Calc)","sxc","application/vnd.sun.xml.calc",1),
//    OpenOfficeorg_10_Spreadsheet_Template("OpenOffice.org 1.0 Spreadsheet Template","calc_StarOffice_XML_Calc_Template","stc","application/vnd.sun.xml.calc.template",1),
//    Data_Interchange_Format("Data Interchange Format","DIF","dif","application/octet-stream",1),
//    dBASE("dBASE","dBase","dbf","application/dbase",1),
//    Microsoft_Excel_97_2000_XP("Microsoft Excel 97/2000/XP","MS Excel 97","xls","application/vnd.ms-excel",1),
//    Microsoft_Excel_97_2000_XP_Template("Microsoft Excel 97/2000/XP Template","MS Excel 97 Vorlage/Template","xlt","aapplication/vnd.ms-excel",1),
//    Microsoft_Excel_95("Microsoft Excel 95","MS Excel 95","xls","application/vnd.ms-excel",1),
//    Microsoft_Excel_95_Template("Microsoft Excel 95 Template","MS Excel 95 Vorlage/Template","xlt","application/vnd.ms-excel",1),
//    Microsoft_Excel_50("Microsoft Excel 5.0","MS Excel 5.0/95","xls","application/vnd.ms-excel",1),
//    Microsoft_Excel_50_Template("Microsoft Excel 5.0 Template","MS Excel 5.0/95 Vorlage/Template","xlt","application/vnd.ms-excel",1),
//    StarCalc_50("StarCalc 5.0","StarCalc 5.0","sdc","application/vnd.stardivision.calc",1),
//    StarCalc_50_Template("StarCalc 5.0 Template","StarCalc 5.0 Vorlage/Template","vor","application/octet-stream",1),
//    StarCalc_40("StarCalc 4.0","StarCalc 4.0","sdc","application/vnd.stardivision.calc",1),
//    StarCalc_40_Template("StarCalc 4.0 Template","StarCalc 4.0 Vorlage/Template","vor","application/octet-stream",1),
//    StarCalc_30("StarCalc 3.0","StarCalc 3.0","sdc","application/vnd.stardivision.calc",1),
//    StarCalc_30_Template("StarCalc 3.0 Template","StarCalc 3.0 Vorlage/Template","vor","application/octet-stream",1),
//    SYLK("SYLK","SYLK","slk","application/excel",1),
//    Text_CSV("Text CSV","Text - txt - csv (StarCalc)","csv","text/csv",1),
//    HTML_Document_OpenOfficeorg_Calc("HTML Document (OpenOffice.org Calc)","HTML (StarCalc)","html","text/html",1),
//    Microsoft_Excel_2003_XML("Microsoft Excel 2003 XML","MS Excel 2003 XML","xml","application/xml",1),
//    Pocket_Excel("Pocket Excel","Pocket Excel","pxl","application/octet-stream",1),

    OpenDocument_Presentation("OpenDocument Presentation","impress8","odp","application/vnd.oasis.opendocument.presentation","impress8",2),
//    OpenDocument_Presentation_Template("OpenDocument Presentation Template","impress8_template","otp","application/vnd.oasis.opendocument.presentation-template",2),
    OpenOfficeorg_10_Presentation("OpenOffice.org 1.0 Presentation","StarOffice XML (Impress)","sxi","application/vnd.sun.xml.impress","impress_StarOffice_XML_Impress",2),
//    OpenOfficeorg_10_Presentation_Template("OpenOffice.org 1.0 Presentation Template","impress_StarOffice_XML_Impress_Template","sti","application/vnd.sun.xml.impress.template",2),
    Microsoft_PowerPoint_97_2000_XP("Microsoft PowerPoint 97/2000/XP","MS PowerPoint 97","ppt","application/vnd.ms-powerpoint","impress_MS_PowerPoint_97",2),
//    Microsoft_PowerPoint_97_2000_XP_Template("Microsoft PowerPoint 97/2000/XP Template","MS PowerPoint 97 Vorlage","pot","application/vnd.ms-powerpoint",2),
//    OpenOfficeorg_10_Drawing_OpenOfficeorg_Impress("OpenOffice.org 1.0 Drawing (OpenOffice.org Impress)","impress_StarOffice_XML_Draw","sxd","application/vnd.sun.xml.draw",2),
//    StarDraw_50_OpenOfficeorg_Impress("StarDraw 5.0 (OpenOffice.org Impress)","StarDraw 5.0 (StarImpress)","sda","application/vnd.stardivision.draw",2),
//    StarDraw_30_OpenOfficeorg_Impress("StarDraw 3.0 (OpenOffice.org Impress)","StarDraw 3.0 (StarImpress)","sdd","application/vnd.stardivision.draw",2),
//    StarImpress_50("StarImpress 5.0","StarImpress 5.0","sdd","application/vnd.stardivision.impress",2),
//    StarImpress_50_Template("StarImpress 5.0 Template","StarImpress 5.0 Vorlage","vor","application/octet-stream",2),
//    StarImpress_40("StarImpress 4.0","StarImpress 4.0","sdd","application/vnd.stardivision.impress",2),
//    StarImpress_40_Template("StarImpress 4.0 Template","StarImpress 4.0 Vorlage","vor","application/octet-stream",2),
//    OpenDocument_Drawing_Impress("OpenDocument Drawing (Impress)","impress8_draw","odg","application/vnd.oasis.opendocument.graphics",2)
    ;


    private String formatName;
    private String filterName;
    private String fileExtension;
    private String mimeType;
    private int handlerType;
    private String documentType;
    //private String ooEquivName;

    private OOoFormat(String formatName, String filterName, String fileExtension, String mimeType, String documentType, int handlerType) {
        this.formatName = formatName;
        this.filterName = filterName;
        this.fileExtension = fileExtension;
        this.mimeType = mimeType;
        this.documentType = documentType;
        this.handlerType = handlerType;
        //this.ooEquivName = ooEquivName;
    }

    public String getFileExtension() {
        return fileExtension;
    }

    public String getFilterName() {
        return filterName;
    }

    public String getFormatName() {
        return formatName;
    }

    public String getMimeType() {
        return mimeType;
    }

    public int getHandlerType() {
        return handlerType;
    }

    public String getDocumentType() {
        return documentType;
    }


    public static OOoFormat findFormatForFilterName(String filterName) {
        OOoFormat[] formats =  OOoFormat.values();
        for (OOoFormat format:formats) {
            if (format.getFilterName().equals(filterName)) {
                return format;
                }
        }
        return null;
    }

    public static OOoFormat findFormatWithDocumentType(String documentType) {
        OOoFormat[] formats =  OOoFormat.values();
        for (OOoFormat format:formats) {
            if (format.getDocumentType().equals(documentType)) {
                return format;
                }
        }
        return null;
    }

//    public String getOOEquivalent() {
//        if (ooEquivName.length() == 0)
//            return null;
//        return ooEquivName;
//    }

    @Override
    public String toString() {
        return fileExtension;
    }
}
