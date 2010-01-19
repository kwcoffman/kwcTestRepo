/*
 * DecompUtil.java
 *
 * Created on 2010.01.12
 *
 */

package edu.umich.med.umms;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.IOException;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;

import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.UnknownPropertyException;

import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;

import com.sun.star.document.XExporter;
import com.sun.star.document.XFilter;
import com.sun.star.document.XTypeDetection;
import com.sun.star.document.XStorageBasedDocument;

import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XShape;
//import com.sun.star.drawing.XTextShape; // Doesn't exist?
//import com.sun.star.drawing.GraphicObjectShape;

import com.sun.star.embed.XStorage;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
//import com.sun.star.gallery.GalleryItemType;
//import com.sun.star.gallery.XGalleryThemeProvider;
//import com.sun.star.gallery.XGalleryTheme;
//import com.sun.star.gallery.XGalleryItem;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.io.XInputStream;
import com.sun.star.io.XStream;

import com.sun.star.lang.XServiceInfo;

import com.sun.star.lib.uno.adapter.XOutputStreamToByteArrayAdapter;
//import com.sun.star.table.BorderLine;

import com.sun.star.ucb.XSimpleFileAccess2;

import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 *
 * @author kwc@umich.edu
 */
public class DecompUtil {

    private static boolean verbose = false;
    private static boolean excludeCustomShapes = false;

    public static boolean beingVerbose()
    {
        return verbose;
    }

    public static void setVerbosity(boolean value)
    {
        verbose = value;
    }

    public static boolean excludingCustomShapes()
    {
        return excludeCustomShapes;
    }

    public static void excludeCustomShapes()
    {
        excludeCustomShapes = true;
    }

    public static XComponent openFileForProcessing(XDesktop xDesktop, String inputFile)
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

    public static XPropertySet duplicateObjectPropertySet(Object origObj)
    {
        XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, origObj);
//        XPropertySetInfo origPropSet = xPropSet.getPropertySetInfo();
//        Property[] origProps = origPropSet.getProperties();
        return xPropSet;
    }

    /* http://www.oooforum.org/forum/viewtopic.phtml?t=45734 */
    public static String getInternalURL(XComponent xDrawDoc,
                                        String srcUrl,
                                        String imgName) throws Exception
    {
        XMultiServiceFactory xFactory = (XMultiServiceFactory)
                UnoRuntime.queryInterface(XMultiServiceFactory.class, xDrawDoc);
        Object xBitObj = xFactory.createInstance("com.sun.star.drawing.BitmapTable");
        XNameContainer bitMaps = (XNameContainer)
                UnoRuntime.queryInterface(XNameContainer.class, xBitObj);

        try {
            bitMaps.insertByName(imgName, srcUrl);
        } catch (com.sun.star.container.ElementExistException e) {
            // Cool, it is already there!
        }
        return (String) bitMaps.getByName(imgName);
    }

    public static XShape createShape(XComponent xDrawDoc,
                                     com.sun.star.awt.Point aPos,
                                     com.sun.star.awt.Size aSize,
                                     String sShapeType ) throws java.lang.Exception
    {
        XShape xShape = null;
        XMultiServiceFactory xFactory = (XMultiServiceFactory)
                UnoRuntime.queryInterface(XMultiServiceFactory.class, xDrawDoc);
        Object xObj = xFactory.createInstance(sShapeType);
        xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xObj);
        xShape.setPosition(aPos);
        xShape.setSize(aSize);
        return xShape;
    }

    public static void printShapeProperties(XShape shape) throws WrappedTargetException
    {

        // Get and print all the shape's properties
        XPropertySet xShapeProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, shape);
        Property[] props = xShapeProperties.getPropertySetInfo().getProperties();
        System.out.println("----- Printing Shape Properties -----");
        for (int x = 0; x < props.length; x++) {
            try {
                System.out.println("    Property " + props[x].Name + " = " + xShapeProperties.getPropertyValue(props[x].Name));
            } catch (UnknownPropertyException ex) {
                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        System.out.println();

    }

    public static void printObjectProperties(Object obj) throws WrappedTargetException
    {

        // Get and print all the shape's properties
        XPropertySet xShapeProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, obj);
        Property[] props = xShapeProperties.getPropertySetInfo().getProperties();
        System.out.println("----- Printing Object Properties -----");
        for (int x = 0; x < props.length; x++) {
            try {
                System.out.println("    Property " + props[x].Name + " = " + xShapeProperties.getPropertyValue(props[x].Name));
            } catch (UnknownPropertyException ex) {
                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        System.out.println();

    }

    // Assumes that the original image shape is supplied
    public static Point calculateCitationImagePosition(XShape xShape)
    {
        Point aPos = xShape.getPosition();
        Size aSize = xShape.getSize();
        Point citationPos = new Point();

        citationPos.X = aPos.X;
        citationPos.Y = aPos.Y + aSize.Height + 200;
        return citationPos;
    }

    // Assumes that the citation image shape is supplied
    public static Point calculateCitationTextPosition(XShape xShape)
    {
        Point aPos = xShape.getPosition();
        Size aSize = xShape.getSize();
        Point citeTextPos = new Point();

        citeTextPos.X = aPos.X + aSize.Width + 200;
        citeTextPos.Y = aPos.Y;
        return citeTextPos;
    }

    // Assumes that the original image shape is supplied
    public static Size calculateCitationImageSize(XShape xShape)
    {
        Point aPos = xShape.getPosition();
        Size aSize = xShape.getSize();
        Size citationSize = new Size();

        citationSize.Width = 88 * 20; // Image is 88x31 pixels -- show it 20 times that size
        citationSize.Height = 31 * 20;
        return citationSize;
    }

    // Assumes that the citation image shape is supplied
    public static Size calculateCitationTextSize(XShape xImage, XShape xCiteImage)
    {
        Size imageSize = xImage.getSize();
        Size citeSize = xCiteImage.getSize();
        Size citeTextSize = new Size();

        citeTextSize.Width = imageSize.Width - citeSize.Width - 200;
        citeTextSize.Height = citeSize.Height;
        return citeTextSize;
    }

    public static void printSupportedServices(XComponent xCompDoc)
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

    public static void exportContextImage(XComponentContext xContext,
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


    public static void exportImage(XComponentContext xContext,
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

    public static void extractImageByURL(XComponentContext xContext,
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
            if (DecompUtil.beingVerbose()) System.out.printf("There were a total of %d pictures found via DocStorageAccess\n", aNames.length);
            for (int i = 0; i < aNames.length; i++) {
                //System.out.printf("Picture %d has name '%s'\n", i+1, aNames[i]);
                //System.out.printf("Processing picture with name '%s'\n", aNames[i]);
                if (aNames[i].contains(pictureURL)) {
                    Object oElement = xPicturesNameAccess.getByName(aNames[i]);
                    XStream xStream = (XStream) UnoRuntime.queryInterface(XStream.class, oElement);
                    Object oInput = xStream.getInputStream();
                    XInputStream xInputStream = (XInputStream)
                            UnoRuntime.queryInterface(XInputStream.class, oInput);

                    xFileWriter.writeFile(outName + "." + DecompUtil.getExtension(aNames[i]), xInputStream);
                    break;
                }

            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            return;
        }

    }

    public static void printDocumentType(XComponentContext xContext,
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

    public static String getDocumentType(XComponentContext xContext,
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

    public static String fileNameToOOoURL(final String fName) {
        StringBuilder sLoadUrl = new StringBuilder("file://");
        sLoadUrl.append(fName.replace('\\', '/'));
        return sLoadUrl.toString();
    }

    public static void storeDocument(XComponentContext xContext,
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

    public static boolean isaSupportedFormat(OOoFormat format)
    {
        return (java.util.Arrays.asList(OpenOfficeUNODecomposition.SUPPORTED_FORMATS).contains(format));
    }

    public static OOoFormat getNativeFormat(OOoFormat origFormat)
    {
        OOoFormat nativeFormat = origFormat;

        if (origFormat.getHandlerType() == 0) {
            nativeFormat = OOoFormat.OpenDocument_Text;
        } else if (origFormat.getHandlerType() == 2) {
            nativeFormat = OOoFormat.OpenDocument_Presentation;
        }
        return nativeFormat;
    }

    public static String possiblyUseTemporaryDocument(XComponentContext xContext,
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

    public static String getExtension(String fullName)
    {
        int dotPos = fullName.lastIndexOf(".");
        String strExtension = fullName.substring(dotPos + 1);
        return strExtension;
        //String strFilename = fullName.substring(0, dotPos);
    }

    public static String constructBaseImageName(String dir, int page, int num)
    {
        String outName = new String();

        outName = String.format("%s/image-%05d-%03d", dir, page, num);
        return outName;
    }

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

    public static byte[] getImageByteStream(String imageURL)
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

}
