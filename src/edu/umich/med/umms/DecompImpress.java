/*
 * DecompImpress.java
 *
 * Created on 2010.01.12
 *
 */

package edu.umich.med.umms;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.UnknownPropertyException;

import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XShape;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
//import com.sun.star.table.BorderLine;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;

import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 *
 * @author kwc@umich.edu
 */
public class DecompImpress {

    public static void handleDocument(XComponentContext xContext,
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
            if (DecompUtil.beingVerbose()) System.out.printf("xDrawPages.getCount returned a value of '%d' pages\n", pageCount);
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
                DecompUtil.exportContextImage(xContext, xMCF, currPage, outputDir, p+1);

                int shapeCount = currPage.getCount();
                if (DecompUtil.beingVerbose()) System.out.printf("Page %d has %d shapes\n", p+1, shapeCount);
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
                    if (DecompUtil.beingVerbose()) System.out.printf("--- Working with shape %d (At %d:%d, size %dx%d)\ttype: %s---\n", s + 1, shapePoint.X, shapePoint.Y, shapeSize.Width, shapeSize.Height, currType);

                    //printShapeProperties(currShape);

                    /* Note that we specifically ignore TitleTextShape, OutlinerShape, and LineShape */
                    if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
                        if (DecompUtil.beingVerbose()) System.out.printf("Handling GraphicObjectShape (%d) on page %d\n", s+1, p+1);
//                        exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                        XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, currShape);
                        String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                        pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                        String outName = DecompUtil.constructBaseImageName(outputDir, p+1, s+1);
                        DecompUtil.extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.TableShape")) {
                        if (DecompUtil.beingVerbose()) System.out.printf("Handling TableShape (%d) on page %d\n", s+1, p+1);
                        DecompUtil.exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.GroupShape")) {
                        if (DecompUtil.beingVerbose()) System.out.printf("Handling GroupShape (%d) on page %d\n", s+1, p+1);
                        DecompUtil.exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.CustomShape")) {
                        if (!DecompUtil.excludingCustomShapes()) {
                            if (DecompUtil.beingVerbose()) System.out.printf("Handling CustomShape (%d) on page %d\n", s+1, p+1);
                            DecompUtil.exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
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
   public static int replacePresentationDocImage(XComponentContext xContext,
                                                 XMultiComponentFactory xMCF,
                                                 XComponent xCompDoc,
                                                 String originalImageName,
                                                 String replacementURL,
                                                 int p,
                                                 int s)
    {

        XDrawPage currPage = null;
        XShape currShape = null;

        byte[] replacementByteArray = DecompUtil.getImageByteStream(replacementURL);
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

            XPropertySet origProps = DecompUtil.duplicateObjectPropertySet(currShape);

            origProps.setPropertyValue("Graphic", xGraphic);

        } catch (Exception exception) {
            System.out.println("Couldn't set image properties");
            return (4);
        }

        return 0;
    }

    /* XXX  To be completed
       http://www.oooforum.org/forum/viewtopic.phtml?t=45734
     */
    public static int insertPresentationDocImageCitation(XComponentContext xContext,
                                                   XMultiComponentFactory xMCF,
                                                   XComponent xCompDoc,
                                                   String originalImageName,
                                                   String citationURL,
                                                   int p,
                                                   int s)
    {
        try {
            XDrawPagesSupplier xDrawPagesSuppl =
                    (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
            if (xDrawPagesSuppl == null) {
                System.out.println("Failed to get xDrawPagesSuppl from xComp");
                return 1;
            }

            XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
//            Object pageObject = xDrawPages.getByIndex(p);

            System.out.printf("insertPresentationDocImageCitation: document has '%d' pages\n", xDrawPages.getCount());

            XDrawPage drawPage =
                    (XDrawPage) UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex(p));


            //put something on the drawpage

            System.out.printf("drawPage.getCount says there are %d objects\n", drawPage.getCount());

            com.sun.star.drawing.XShapes xShapes = (com.sun.star.drawing.XShapes)
                    UnoRuntime.queryInterface(com.sun.star.drawing.XShapes.class, drawPage);

            // Use the original image location and size to determine the location
            // and size (at least the width?) of the citation information

            Object oOrigImage = xShapes.getByIndex(s);
            XPropertySet xOrigPropSet = (XPropertySet)
                    UnoRuntime.queryInterface(XPropertySet.class, oOrigImage);
//            printObjectProperties(oOrigImage);
            XShape xOrigImage = (XShape) UnoRuntime.queryInterface(XShape.class, oOrigImage);
//            printShapeProperties(xOrigImage);

            Point citeImagePos = DecompUtil.calculateCitationImagePosition(xOrigImage);
            Size citeImageSize = DecompUtil.calculateCitationImageSize(xOrigImage);

            String convertedURL = DecompUtil.getInternalURL(xCompDoc, citationURL, "image");

            try {
                // Add citation image
                XShape xCIShape = DecompUtil.createShape(xCompDoc, citeImagePos, citeImageSize,
                                        "com.sun.star.drawing.GraphicObjectShape");
                XPropertySet xImageProps = (XPropertySet)
                UnoRuntime.queryInterface(XPropertySet.class, xCIShape);
                xImageProps.setPropertyValue("GraphicURL", convertedURL);
                xShapes.add(xCIShape);
                System.out.printf("drawPage.getCount nows says there are %d objects\n", drawPage.getCount());

                Point citeTextPos = DecompUtil.calculateCitationTextPosition(xCIShape);
                Size citeTextSize = DecompUtil.calculateCitationTextSize(xOrigImage, xCIShape);

                // Add citation text
                XShape xCTShape = DecompUtil.createShape(xCompDoc, citeTextPos, citeTextSize,
                                        "com.sun.star.drawing.TextShape");  // There is also a TextShape?
                xShapes.add(xCTShape);
/***
                XPropertySet xCTProps = UnoRuntime.queryInterface(XPropertySet.class, xCTShape);
                // blue fill color
                xCTProps.setPropertyValue("FillColor", new Integer(0x00c000));
                // black line color
                xCTProps.setPropertyValue("LineColor", new Integer(0xffffff));

                xCTProps.setPropertyValue("TextFitToSize", com.sun.star.drawing.TextFitToSizeType.PROPORTIONAL);
                xCTProps.setPropertyValue("TextAutoGrowHeight", new Boolean(true));
                xCTProps.setPropertyValue("TextAutoGrowWidth", new Boolean(true));
***/
/*
                // blue fill color
                xCTProps.setPropertyValue("FillColor", new Integer(0x0000C0));
                // black line color
                xCTProps.setPropertyValue("LineColor", new Integer(0x000000));
                xCTProps.setPropertyValue("Name", "Rounded Gray Rectangle");

                //xCTProps.setPropertyValue("TextFitToSize", new Boolean(true));
                xCTProps.setPropertyValue("TextFitToSize", com.sun.star.drawing.TextFitToSizeType.PROPORTIONAL);
                //xCTProps.setPropertyValue("TextAutoGrowHeight", new Boolean(true));
                xCTProps.setPropertyValue("TextAutoGrowHeight", true);
                xCTProps.setPropertyValue("TextAutoGrowWidth", true);

                xCTProps.setPropertyValue("CharColor", new Integer(0xCC0000));  //  XXX Doesn't do anything?
                xCTProps.setPropertyValue("CharHeight", 4);  //  XXX Doesn't do anything?
                //xCTProps.setPropertyValue("LineStyle", com.sun.star.drawing.LineStyle.NONE);
                //xCTProps.setPropertyValue("FillStyle", com.sun.star.drawing.FillStyle.NONE);
*/
                XText xText = UnoRuntime.queryInterface(XText.class, xCTShape);
                XTextCursor xTextCursor = xText.createTextCursor();
                XPropertySet xTxtProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
                
                xTxtProps.setPropertyValue("CharHeight", 8);
                xTxtProps.setPropertyValue("CharColor", new Integer(0xffffff));
                xText.setString("http://open.umich.edu This is a long string to see what will happen with a large amount of text if the text is too long to fit within the rectangle.  We'll add even more text to see what happens when the smaller text exceeds the defined rectangle that is supposed to contain the text.");
                //xTxtProps.setPropertyValue("HyperLinkURL", "http://open.umich.edu"); // XXX Doesn't work for impress documents.
                //xText.setString("CC-BY");

            } catch (java.lang.Exception ex) {
                Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            }

        } catch (Exception ex) {
            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        }
        return 0;
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

}
