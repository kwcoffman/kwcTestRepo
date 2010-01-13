/*
 * DecompText.java
 *
 * Created on 2010.01.12
 *
 */

package edu.umich.med.umms;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;

import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.container.NoSuchElementException;

import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XShape;

import com.sun.star.frame.XStorable;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.XOutputStreamToByteArrayAdapter;

import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextEmbeddedObject;
import com.sun.star.text.XTextEmbeddedObjectsSupplier;
import com.sun.star.text.XTextGraphicObjectsSupplier;

import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;



/**
 *
 * @author kwc@umich.edu
 */
public class DecompText {

    /* This version uses names rather than indexes to gather the images... */
    public static void extractImagesUsingNames(XComponentContext xContext,
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
            Object oGraphicProvider = xMCF.createInstanceWithContext(
                    "com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider)
                    UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);

            XTextGraphicObjectsSupplier xTextGraphSuppl = (XTextGraphicObjectsSupplier)
                    UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xGOnames = xTextGraphSuppl.getGraphicObjects();
            if (xGOnames.hasElements()) {
                Class objclass = xGOnames.getClass();
                String[] graphicNames = xGOnames.getElementNames();
                int numImages = graphicNames.length;
                if (DecompUtil.beingVerbose()) System.out.printf("There are %d GraphicObjects in this file\n", numImages);
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
                        if (DecompUtil.beingVerbose()) System.out.printf("The name of the image is: '%s'\n", name);
                        //Object graphicsObject = xGOnames.getByName(graphicNames[i]);
                        XShape graphicShape = (XShape)
                                UnoRuntime.queryInterface(XShape.class, xGOnames.getByName(name));
                        XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, graphicShape);
                        String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                        pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                        String outName = DecompUtil.constructBaseImageName(outputDir, 1, i);
                        DecompUtil.extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
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

    }

    /* Using indexes rather than names */
    public static void extractImages(XComponentContext xContext,
                                     XMultiComponentFactory xMCF,
                                     XComponent xCompDoc,
                                     String outputDir)
    {
        XTextDocument xTextDoc =
                    (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            System.out.printf("Cannot get XTextDocument interface for Text Document???\n");
            System.exit(7);
        }

        try {
//            XMultiServiceFactory xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(
//                XMultiServiceFactory.class, xCompDoc);

//            xGraphicProvider.storeGraphic(arg0, arg1);
            Object oGraphicProvider = xMCF.createInstanceWithContext(
                    "com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider)
                    UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);

            XTextGraphicObjectsSupplier xTextGOS = (XTextGraphicObjectsSupplier)
                    UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xEONames = xTextGOS.getGraphicObjects();
            XIndexAccess xEOIndexes = (XIndexAccess)
                    UnoRuntime.queryInterface(XIndexAccess.class, xEONames);
            if (DecompUtil.beingVerbose()) System.out.printf("There are %d embedded objects in this document\n", xEOIndexes.getCount());

            for (int i = 0; i < xEOIndexes.getCount(); i++) {
                //Object oTextEO = xEOIndexes.getByIndex(i);
                //XTextEmbeddedObject xTextEO = (XTextEmbeddedObject) UnoRuntime.queryInterface(XTextEmbeddedObject.class, oTextEO);
                XShape graphicShape = (XShape)
                        UnoRuntime.queryInterface(XShape.class, xEOIndexes.getByIndex(i));
                XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, graphicShape);
                String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                String outName = DecompUtil.constructBaseImageName(outputDir, 1, i);
                DecompUtil.extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    // XXX Revise this to use index value rather than a name !?!?!?!?!? XXX
    /* From http://www.oooforum.org/forum/viewtopic.phtml?t=81870 */
    public static int replaceImage(XComponentContext xContext,
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

        byte[] replacementByteArray = DecompUtil.getImageByteStream(replacementURL);
        ByteArrayToXInputStreamAdapter xSource = new ByteArrayToXInputStreamAdapter(replacementByteArray);

        // Querying for the interface XMultiServiceFactory on the xtextdocument
        XMultiServiceFactory xMSFDoc = (XMultiServiceFactory)
                UnoRuntime.queryInterface(XMultiServiceFactory.class, xTextDoc);
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

        // XXX Should use getImageObjectByName(xCompDoc, originalImageName);
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

            XPropertySet origProps = DecompUtil.duplicateObjectPropertySet(xImageObject);


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

    /* XXX To be completed! */
    public static int insertImageCitation(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String originalImageName,
                                          String replacementURL)
    {
        return 0;
    }



    /* Original code is from:
     *   http://www.oooforum.org/forum/viewtopic.phtml?t=17139
     * Seems pretty complicated!!
     */
/*
    private static String convertLinkedImageToEmbeddedImage(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String origImageURL) {
        String jlsInternalUrl = null;
        XTextDocument xTextDoc = (XTextDocument)
                UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            return jlsInternalUrl;
        }

        try {
            XMultiServiceFactory xMSFDoc = (XMultiServiceFactory)
                    UnoRuntime.queryInterface(XMultiServiceFactory.class, xTextDoc);

            // Get document object from the class member m_xTextDocument
            XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier)
                    UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
            Object origGO = xNAGraphicObjects.getByName(origImageURL);
            XTextContent xTCGraphicObject = (XTextContent)
                    UnoRuntime.queryInterface(XTextContent.class, origGO);
            XPropertySet xPSTCGraphicObject = (XPropertySet)
                    UnoRuntime.queryInterface(XPropertySet.class, xTCGraphicObject);

            // Current internal URL of the image
            String jlsGOUrl = xPSTCGraphicObject.getPropertyValue("GraphicURL").toString();
            XTextRange xTRTCGraphicObject = xTCGraphicObject.getAnchor();

            // New internal URL
            jlsInternalUrl = new String("");

            // Temporary GraphicObject number 1.
            Object jloShape1 = xMSFDoc.createInstance("com.sun.star.drawing.GraphicObjectShape");
            XShape xGOShape1 = (XShape) UnoRuntime.queryInterface(XShape.class, jloShape1);
            XTextContent xTCShape1 = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, xGOShape1);
            XPropertySet xPSShape1 = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xGOShape1);

            // Temporary GraphicObject number 2.
            Object jloShape2 = xMSFDoc.createInstance("com.sun.star.drawing.GraphicObjectShape");
            XShape xGOShape2 = (XShape) UnoRuntime.queryInterface(XShape.class, jloShape2);
            XTextContent xTCShape2 = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, xGOShape2);
            XPropertySet xPSShape2 = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xGOShape2);

            // The temporary GraphicObject number 1 points to the original GraphicObject file
            xPSShape1.setPropertyValue("GraphicURL", jlsGOUrl);
            // By inserting the temporary GraphicObject number 1, OpenOffice creates its bitmap.
            xTRTCGraphicObject.getText().insertTextContent(xTRTCGraphicObject, xTCShape1, false);
            // Assign the Bitmap of temporary GraphicObject number 1 to temporary GraphicObject number 2.
            xPSShape2.setPropertyValue("GraphicObjectFillBitmap", xPSShape1.getPropertyValue("GraphicObjectFillBitmap"));
            // By inserting the temporary GraphicObject number 2, OpenOffice creates its internal URL (and its bitmap).
            xTRTCGraphicObject.getText().insertTextContent(xTRTCGraphicObject, xTCShape2, false);
            // The temporary GraphicObject number 1 is no longer needed.
            xTRTCGraphicObject.getText().removeTextContent(xTCShape1);
            // Get the internal URL of the temporary GraphicObject number 2.
            jlsInternalUrl = xPSShape2.getPropertyValue("GraphicURL").toString();
            // Assign to original GraphicObject the same internal URL as the Temporary GraphicObject number 2.
            xPSTCGraphicObject.setPropertyValue("GraphicURL", jlsInternalUrl);
            // The temporary GraphicObject number 2 is no longer needed.
            xTRTCGraphicObject.getText().removeTextContent(xTCShape2);
        } catch (UnknownPropertyException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (NoSuchElementException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
        return jlsInternalUrl;
    }
*/

    private static Object getImageObjectByName(XComponent xTextDoc, String name)
    {
        Any xImageAny = null;
        Object xImageObject = null;
        try {
            XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xTextDoc);
            XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
            xImageAny = (Any) xNAGraphicObjects.getByName(name);
        } catch (NoSuchElementException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
        if (xImageAny == null) {
            return 1;
        }
        return xImageAny.getObject();
    }

    private static void storeImage(XComponentContext xContext,
                                   XMultiComponentFactory xMCF,
                                   XComponent xCompDoc,
                                   XShape xShape,
                                   String outputDir,
                                   int p,
                                   int s) {
        String newFName = String.format("%s/%s-%05d-%03d.%s", outputDir, "image", p, s, "jpg");
        String newFNameURL = DecompUtil.fileNameToOOoURL(newFName);

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

    private static XDrawPage getDrawPage(XDrawPage xDrawPage, int nIndex)
    {
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

}
