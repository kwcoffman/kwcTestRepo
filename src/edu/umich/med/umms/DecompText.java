/*
 * DecompText.java
 *
 * Created on 2010.01.12
 *
 */

package edu.umich.med.umms;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;

import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.container.NoSuchElementException;

import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;

import com.sun.star.frame.XStorable;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.XOutputStreamToByteArrayAdapter;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.HoriOrientation;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.WrapTextMode;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;

import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextGraphicObjectsSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;

import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 *
 * @author kwc@umich.edu
 */
public class DecompText {

    private com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();
    private org.apache.log4j.Level myLogLevel = org.apache.log4j.Level.WARN;

    public DecompText()
    {
        mylog = com.spinn3r.log5j.Logger.getLogger();
    }

    public void setLoggingLevel(org.apache.log4j.Level lvl)
    {
        myLogLevel = lvl;
        mylog.setLevel(myLogLevel);
    }

    /* Extract Images using indexes rather than names */
    public int extractImages(XComponentContext xContext,
                                     XMultiComponentFactory xMCF,
                                     XComponent xCompDoc,
                                     String outputDir,
                                     boolean excludeCustomShapes)
    {
        DecompUtil du = new DecompUtil();
        du.setLoggingLevel(myLogLevel);

        XTextDocument xTextDoc =
                    (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            mylog.error("Cannot get XTextDocument interface for Text Document???");
            return 7;
        }

        try {
            Object oGraphicProvider = xMCF.createInstanceWithContext(
                    "com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider)
                    UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);

            XTextGraphicObjectsSupplier xTextGOS = (XTextGraphicObjectsSupplier)
                    UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xEONames = xTextGOS.getGraphicObjects();
            XIndexAccess xEOIndexes = (XIndexAccess)
                    UnoRuntime.queryInterface(XIndexAccess.class, xEONames);
            mylog.debug("There are %d GraphicsObjects in this document\n", xEOIndexes.getCount());

            for (int i = 0; i < xEOIndexes.getCount(); i++) {
                //Object oTextEO = xEOIndexes.getByIndex(i);
                //XTextEmbeddedObject xTextEO = (XTextEmbeddedObject) UnoRuntime.queryInterface(XTextEmbeddedObject.class, oTextEO);
                XShape graphicShape = (XShape)
                        UnoRuntime.queryInterface(XShape.class, xEOIndexes.getByIndex(i));
                XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, graphicShape);
                String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                String outName = DecompUtil.constructBaseImageName(outputDir, 1, i);
                du.extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
            }
        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            return 40;
        }
        return 0;
    }

    private Object getImageObject(XTextDocument xTextDoc, int s)
    {
        // Get reference to image object in the text document
        Any xImageAny = null;
        try {
            XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xTextDoc);
            XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
            XIndexAccess xEOIndexes = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, xNAGraphicObjects);
            //          xImageAny = (Any) xNAGraphicObjects.getByName(originalImageName);
            xImageAny = (Any) xEOIndexes.getByIndex(s);
        } catch (IndexOutOfBoundsException ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
        if (xImageAny == null) {
            return null;
        }
        return xImageAny.getObject();
    }


    // XXX Revise this to use index value rather than a name !?!?!?!?!? XXX
    /* From http://www.oooforum.org/forum/viewtopic.phtml?t=81870 */
    public int replaceImage(XComponentContext xContext,
                                           XMultiComponentFactory xMCF,
                                           XComponent xCompDoc,
                                           String originalImageName,
                                           String replacementURL)
    {

        XTextDocument xTextDoc =
                    (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            mylog.error("Cannot get XTextDocument interface for Text Document???");
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
            mylog.error("Could not create TextGraphicObject instance");
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

            XPropertySet origProps = DecompUtil.getObjectPropertySet(xImageObject);


            origProps.setPropertyValue("Graphic", xGraphic);

        } catch (Exception exception) {
            mylog.error("Couldn't set image properties");
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
            mylog.error("Could not insert Content");
            exception.printStackTrace(System.err);
            return 1;
        }
 */
        return 0;
    }

    // This places the button according to the Point and Size as specified.  I want it "inline" with the text
    private int insertLicenseButton(XComponentContext xContext,
                                           XMultiComponentFactory xMCF,
                                           XComponent xCompDoc,
                                           String citationURL,
                                           XText xText,
                                           XParagraphCursor xCursor)
    {
        try {
            String convertedURL = DecompUtil.getInternalURL(xCompDoc, citationURL, citationURL);

            // From http://www.oooforum.org/forum/viewtopic.phtml?t=74008
            XMultiServiceFactory xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, xCompDoc);
            Object xGraphic = xMSF.createInstance("com.sun.star.text.TextGraphicObject");
            XTextContent xTextContent = (XTextContent) UnoRuntime.queryInterface(
                    XTextContent.class, xGraphic);
            XPropertySet xImageProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextContent);

            xImageProps.setPropertyValue("AnchorType", TextContentAnchorType.AS_CHARACTER);
            xImageProps.setPropertyValue("GraphicURL", convertedURL);
            xImageProps.setPropertyValue("Width", new Integer(88*20));
            xImageProps.setPropertyValue("Height", new Integer(31*20));
            xImageProps.setPropertyValue("TextWrap", WrapTextMode.DYNAMIC);
            xImageProps.setPropertyValue("HoriOrient", HoriOrientation.LEFT);

            xText.insertTextContent(xCursor, xTextContent, false);

            return 0;
        } catch (Exception ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
        } catch (java.lang.Exception ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
        }
        return 1;
    }

    public int insertImageCitation(XComponentContext xContext,
                                   XMultiComponentFactory xMCF,
                                   XComponent xCompDoc,
                                   String citationText,
                                   String citationURL,
                                   int p,
                                   int s)
    {
        try {
            XTextDocument xTextDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
            if (xTextDoc == null) {
                mylog.error("Cannot get XTextDocument interface for Text Document???");
                System.exit(7);
            }
            XTextContent xImage = (XTextContent) getImageObject(xTextDoc, s);
            if (xImage == null) {
                return 1;
            }


            // This one successfully puts the text AFTER the image.  There must be a more straightforward way??
            XTextRange imageRange = xImage.getAnchor();
            XText xText = imageRange.getText();
            XTextCursor xCursor = xText.createTextCursorByRange(imageRange);

            XParagraphCursor xParaCursor = (XParagraphCursor) UnoRuntime.queryInterface(XParagraphCursor.class, xCursor);
            xParaCursor.gotoEndOfParagraph(false);

            // Insert "license button"

            xText.insertControlCharacter(xParaCursor, ControlCharacter.PARAGRAPH_BREAK, false);
            insertLicenseButton(xContext, xMCF, xCompDoc, citationURL, xText, xParaCursor);
            xText.insertString(xParaCursor, citationText, false);
//            xText.insertString(xParaCursor, "http://open.umich.edu ", false);
//            xText.insertString(xParaCursor, "This is a long string to see what will happen with a large amount of text if the text is ", false);
//            xText.insertString(xParaCursor, "too long to fit within the rectangle.  We'll add even more text to see what happens ", false);
//            xText.insertString(xParaCursor, "when the smaller text exceeds the defined rectangle that is supposed to contain the text.", false);
            xText.insertControlCharacter(xParaCursor, ControlCharacter.PARAGRAPH_BREAK, false);

        } catch (IllegalArgumentException ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        }
        return 0;
    }

    public int insertImageCitationsAsReferences(XComponentContext xContext,
                                                XMultiComponentFactory xMCF,
                                                XComponent xCompDoc,
                                                String citationString,
                                                String citationURL,
                                                int p,
                                                int s)
    {
        // XXX Somehow... need to find button image file name and citation text for each image...

        XTextDocument xTextDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
        if (xTextDoc == null) {
            mylog.error("Cannot get XTextDocument interface for Text Document???");
            System.exit(7);
        }
        XTextContent xImage = (XTextContent) getImageObject(xTextDoc, s);
        if (xImage == null) {
            return 1;
        }


        /* XXX FINISH THIS FOR TEXT DOCUMENTS */
        
        return 0;
    }

    public int insertImageCitationKindaWorks(XComponentContext xContext,
                                             XMultiComponentFactory xMCF,
                                             XComponent xCompDoc,
                                             String citationURL,
                                             int p,
                                             int s)
    {
        try {
            XTextDocument xTextDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
            if (xTextDoc == null) {
                mylog.error("Cannot get XTextDocument interface for Text Document???");
                System.exit(7);
            }
            Any xImageAny = null;
            Object xImageObject = null;
            try {
                XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xTextDoc);
                XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
                XIndexAccess xEOIndexes = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, xNAGraphicObjects);
                //          xImageAny = (Any) xNAGraphicObjects.getByName(originalImageName);
                xImageAny = (Any) xEOIndexes.getByIndex(s);
            } catch (IndexOutOfBoundsException ex) {
                Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
            } catch (WrappedTargetException ex) {
                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            }
            if (xImageAny == null) {
                return 1;
            }
            xImageObject = xImageAny.getObject();
            XTextContent xImage = (XTextContent) xImageObject;
            XTextRange origImageRange = xImage.getAnchor();
            XTextRange imageRange = origImageRange.getEnd();

            //XText xText = (XText) ((XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc)).getText();
            //XTextCursor xModelCursor = (XTextCursor) UnoRuntime.queryInterface(XTextCursor.class, xImage);
//            XTextCursor xModelCursor = xText.createTextCursor();

//            XParagraphCursor xParaCursor = (XParagraphCursor) UnoRuntime.queryInterface(XParagraphCursor.class, xModelCursor);
//            xParaCursor.gotoEndOfParagraph(false);
//            xParaCursor.gotoNextParagraph(false);
//            xText.insertControlCharacter(xParaCursor, ControlCharacter.PARAGRAPH_BREAK, false);
//            xText.insertString(xParaCursor, "http://open.umich.edu ", false);
//            xText.insertString(xParaCursor, "This is a long string to see what will happen with a large amount of text if the text is ", false);
//            xText.insertString(xParaCursor, "too long to fit within the rectangle.  We'll add even more text to see what happens ", false);
//            xText.insertString(xParaCursor, "when the smaller text exceeds the defined rectangle that is supposed to contain the text.", false);
//            xText.insertControlCharacter(xParaCursor, ControlCharacter.PARAGRAPH_BREAK, false);
            XText xText = imageRange.getText();
            xText.insertControlCharacter(imageRange, ControlCharacter.PARAGRAPH_BREAK, false);
            xText.insertString(imageRange, "http://open.umich.edu ", false);
            xText.insertString(imageRange, "This is a long string to see what will happen with a large amount of text if the text is ", false);
            xText.insertString(imageRange, "too long to fit within the rectangle.  We'll add even more text to see what happens ", false);
            xText.insertString(imageRange, "when the smaller text exceeds the defined rectangle that is supposed to contain the text.", false);
            xText.insertControlCharacter(imageRange, ControlCharacter.PARAGRAPH_BREAK, false);

        } catch (IllegalArgumentException ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        }
        return 0;
    }

/*
    // XXX To be completed! To be changed to use index rather than name???
    public int insertImageCitation_Take_1(XComponentContext xContext,
                                          XMultiComponentFactory xMCF,
                                          XComponent xCompDoc,
                                          String originalImageName,
                                          String citationURL,
                                          int p,
                                          int s)
    {
        try {
            XTextDocument xTextDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xCompDoc);
            if (xTextDoc == null) {
                mylog.error("Cannot get XTextDocument interface for Text Document???");
                System.exit(7);
            }
//  This is replaced by getImageObjectByName
//            Any aImage = null;
//            Object oImage = null;
//            try {
//                XTextGraphicObjectsSupplier xTGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xTextDoc);
//                XNameAccess xNAGraphicObjects = xTGOS.getGraphicObjects();
//                aImage = (Any) xNAGraphicObjects.getByName(originalImageName);
//            } catch (NoSuchElementException ex) {
//                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
//            } catch (WrappedTargetException ex) {
//                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
//            }
//            if (aImage == null) {
//                return 1;
//            }

            XDrawPage xDrawPage = getDrawPage(xCompDoc);
            if (xDrawPage == null)
                return 1;

            XShapes xShapes = getXShapes(xDrawPage);
            if (xShapes == null)
                return 1;


//            Object oImage = getImageObjectByIndex(xTextDoc, s);
////            Object oImage = getImageObjectByName(xTextDoc, originalImageName);
//            if (oImage == null)
//                return 1;
//            XTextContent xImage = (XTextContent) oImage;
//
//            // Use the original image location and size to determine the location
//            // and size (at least the width?) of the citation information
//            XPropertySet xOrigPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oImage);
//            //            DecompUtil.printObjectProperties(oOrigImage);
//            XShape xOrigImage = (XShape) UnoRuntime.queryInterface(XShape.class, oImage);
//            //            DecompUtil.printShapeProperties(xOrigImage);

            XShape xOrigImage = getXShapeByIndex(xContext, xMCF, xCompDoc, 1);
            DecompUtil.printShapeProperties(xOrigImage);

            XTextContent imageContent = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, xOrigImage);
            XTextRange imageRange = imageContent.getAnchor();
//            XTextCursor docCursor = ((XTextViewCursorSupplier)UnoRuntime.queryInterface(XTextViewCursorSupplier.class, xTextDoc.getCurrentController())).getViewCursor();


//            Point citeImagePos = DecompUtil.calculateCitationImagePosition(xOrigImage);
//            Size citeImageSize = DecompUtil.calculateCitationImageSize(xOrigImage);
//            String convertedURL = DecompUtil.getInternalURL(xCompDoc, citationURL, "image");

//            XShape xCIShape = DecompUtil.createShape(xCompDoc, citeImagePos, citeImageSize,
//                                    "com.sun.star.drawing.GraphicObjectShape");
//            XPropertySet xImageProps = (XPropertySet)
//                    UnoRuntime.queryInterface(XPropertySet.class, xCIShape);
//            xImageProps.setPropertyValue("GraphicURL", convertedURL);
//            xShapes.add(xCIShape);

//            // Caclulate citation text location using citation image location
//            Point citeTextPos = DecompUtil.calculateCitationTextPosition(xCIShape);
//            Size citeTextSize = DecompUtil.calculateCitationTextSize(xOrigImage, xCIShape);

//            // Add citation text
//            XShape xCTShape = DecompUtil.createShape(xCompDoc, citeTextPos, citeTextSize,
//                                    "com.sun.star.drawing.TextShape");  // There is also a TextShape?
//            xShapes.add(xCTShape);
//
//            XText xText = UnoRuntime.queryInterface(XText.class, xCTShape);
//            XTextCursor xTextCursor = xText.createTextCursor();
//            XPropertySet xTxtProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
//
//            xTxtProps.setPropertyValue("CharHeight", 8);
//            xTxtProps.setPropertyValue("CharColor", new Integer(0xffffff));
//            xText.setString("http://open.umich.edu This is a long string to see what will happen with a large amount of text if the text is too long to fit within the rectangle.  We'll add even more text to see what happens when the smaller text exceeds the defined rectangle that is supposed to contain the text.");

            //docCursor.gotoStart(false);
            //docCursor.gotoRange(xOrigImage., false);
            // Get View Cursor
            XText xText = xTextDoc.getText();
            // Get Model Cursor
            XTextCursor xModelCursor = xText.createTextCursorByRange(imageRange);

            xText.insertControlCharacter(xModelCursor, ControlCharacter.PARAGRAPH_BREAK, false);
            xText.insertString(xModelCursor, "http://open.umich.edu ", false);
            xText.insertString(xModelCursor, "This is a long string to see what will happen with a large amount of text if the text is ", false);
            xText.insertString(xModelCursor, "too long to fit within the rectangle.  We'll add even more text to see what happens ", false);
            xText.insertString(xModelCursor, "when the smaller text exceeds the defined rectangle that is supposed to contain the text.", false);
            xText.insertControlCharacter(xModelCursor, ControlCharacter.PARAGRAPH_BREAK, false);


            return 0;
        } catch (Exception ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        } catch (java.lang.Exception ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        }
    }
*/


    /* Original code is from:
     *   http://www.oooforum.org/forum/viewtopic.phtml?t=17139
     * Seems pretty complicated!!
     */
/*
    private String convertLinkedImageToEmbeddedImage(XComponentContext xContext,
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

    private Object getImageObjectByName(XComponent xTextDoc, String name)
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
            return null;
        }
        return xImageAny.getObject();
    }

/*
    private Object getImageObjectByIndex(XComponent xTextDoc, int x)
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
            return null;
        }
        return xImageAny.getObject();
    }
*/
    private XShape getXShapeByIndex(XComponentContext xContext,
                                    XMultiComponentFactory xMCF,
                                    XComponent xCompDoc,
                                    int x)
    {
        try {
            Object oGraphicProvider = xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);
            XTextGraphicObjectsSupplier xTextGOS = (XTextGraphicObjectsSupplier) UnoRuntime.queryInterface(XTextGraphicObjectsSupplier.class, xCompDoc);
            XNameAccess xEONames = xTextGOS.getGraphicObjects();
            XIndexAccess xEOIndexes = (XIndexAccess) UnoRuntime.queryInterface(XIndexAccess.class, xEONames);
            mylog.debug("There are %d GraphicsObjects in this document\n", xEOIndexes.getCount());
            XShape graphicShape = (XShape) UnoRuntime.queryInterface(XShape.class, xEOIndexes.getByIndex(x));
            return graphicShape;
        } catch (Exception ex) {
            Logger.getLogger(DecompText.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    private void storeImage(XComponentContext xContext,
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
            mylog.error("Caught I/O Exception: " + ex.getMessage());
            //ex.printStackTrace();
        } catch (Exception e) {
            mylog.error("Caught Exception storing image: " + e.getMessage());
            //e.printStackTrace();
        }
    }

    private XDrawPage getDrawPage(XDrawPage xDrawPage, int nIndex)
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

    private XDrawPage getDrawPage(XComponent xCompDoc)
    {
            XDrawPageSupplier xDrawPageSuppl =
                    (XDrawPageSupplier) UnoRuntime.queryInterface(XDrawPageSupplier.class, xCompDoc);
            if (xDrawPageSuppl == null) {
                mylog.error("Failed to get xDrawPageSuppl from xComp");
                return null;
            }

            XDrawPage xDrawPage = xDrawPageSuppl.getDrawPage();
            return xDrawPage;

    }

    private XShapes getXShapes(XDrawPage xDrawPage)
    {
        XShapes xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, xDrawPage);
        return xShapes;
    }

}
