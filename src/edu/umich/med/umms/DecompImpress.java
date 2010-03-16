/*
 * DecompImpress.java
 *
 * Created on 2010.01.12
 *
 */

package edu.umich.med.umms;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.datatransfer.XTransferable;

import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageDuplicator;
import com.sun.star.drawing.XShapes;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;

import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextRange;

import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 *
 * @author kwc@umich.edu
 */
public class DecompImpress {

    private com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();
    private org.apache.log4j.Level myLogLevel = org.apache.log4j.Level.WARN;

    DecompImpress(org.apache.log4j.Level lvl)
    {
        mylog = com.spinn3r.log5j.Logger.getLogger();
        mylog.setLevel(lvl);
    }
    
    DecompImpress()
    {
        mylog = com.spinn3r.log5j.Logger.getLogger();
    }

    public void setLoggingLevel(org.apache.log4j.Level lvl)
    {
        myLogLevel = lvl;
        mylog.setLevel(myLogLevel);
    }

    public void extractImages(XComponentContext xContext,
                              XMultiComponentFactory xMCF,
                              XComponent xCompDoc,
                              String outputDir,
                              boolean excludeCustomShapes)
    {
        // Query for the XDrawPagesSupplier interface
        XDrawPagesSupplier xDrawPagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
        if (xDrawPagesSuppl == null) {
            mylog.error("Cannot get XDrawPagesSupplier interface for Presentation Document???");
            System.exit(8);
        }
        DecompUtil du = new DecompUtil();
        du.setLoggingLevel(myLogLevel);

        try {
            XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
            Object firstPage = xDrawPages.getByIndex(0);
            int pageCount = xDrawPages.getCount();
            mylog.debug("xDrawPages.getCount returned a value of '%d' pages", pageCount);
            XDrawPage currPage = null;
            Class pageClass = null;
            XPropertySet pageProps = null;

            // Loop through all the pages of the document
            for (int p = 0; p < pageCount; p++) {
                currPage = getDrawPage(xDrawPages, p);
                if (currPage == null) {
                    mylog.error("Failed to get currPage at page %d!", p+1);
                    xCompDoc.dispose();
                    System.exit(22);
                }
                mylog.debug("=== Working with page %d ===", p+1);
                du.exportContextImage(xContext, xMCF, currPage, outputDir, p+1);

                int shapeCount = currPage.getCount();
                mylog.debug("Page %d has %d shapes", p+1, shapeCount);
                XShape currShape = null;

                // Loop through all the shapes within the page
                for (int s = 0; s < shapeCount; s++) {
                    currShape = getPageShape(currPage, s);
                    if (currShape == null) {
                        mylog.error("Failed to get currShape (%d) from page %d!", s+1, p+1);
                        xCompDoc.dispose();
                        System.exit(33);
                    }
                    String currType = currShape.getShapeType();
                    com.sun.star.awt.Size shapeSize = currShape.getSize();
                    com.sun.star.awt.Point shapePoint = currShape.getPosition();
                    mylog.debug("--- Working with shape %d (At %d:%d, size %dx%d)\ttype: %s---", s + 1, shapePoint.X, shapePoint.Y, shapeSize.Width, shapeSize.Height, currType);

                    //printShapeProperties(currShape);

                    /* Note that we specifically ignore TitleTextShape, OutlinerShape, and LineShape */
                    if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
                        mylog.debug("Handling GraphicObjectShape (%d) on page %d\n", s+1, p+1);
//                        exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                        XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, currShape);
                        String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                        pictureURL = pictureURL.substring(27);  // Chop off the leading "vnd.sun.star.GraphicObject:"
                        String outName = DecompUtil.constructBaseImageName(outputDir, p+1, s+1);
                        du.extractImageByURL(xContext, xMCF, xCompDoc, pictureURL, outName);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.TableShape")) {
                        mylog.debug("Handling TableShape (%d) on page %d\n", s+1, p+1);
                        du.exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.GroupShape")) {
                        mylog.debug("Handling GroupShape (%d) on page %d\n", s+1, p+1);
                        du.exportImage(xContext, xMCF, currShape, outputDir, p + 1, s + 1);
                    } else if (currType.equalsIgnoreCase("com.sun.star.drawing.CustomShape")) {
                        if (!excludeCustomShapes) {
                            mylog.debug("Handling CustomShape (%d) on page %d\n", s+1, p+1);
                            du.exportImage(xContext, xMCF, currShape, outputDir, p+1, s+1);
                        }
                    } else {
                        mylog.debug("SKIPPING unhandled shape type '%s' (%d) on page %d\n", currType, s+1, p+1);
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            mylog.error("extractImages: Caught IndexOutOfBoundsException!");
            //Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            mylog.error("extractImages: Caught WrappedTargetException!");
            //Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (UnknownPropertyException ex) {
            mylog.error("extractImages: Caught UnknownPropertyException!");
            //Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
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
     * XXX So for Drawing documents, we need to keep track of index values.  But
     * what if we insert a new image?  Doesn't that change the index values?
     * Need to understand the answer to that!
     *
     */
   public int replaceImage(XComponentContext xContext,
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
        if (replacementByteArray == null)
            return 4;

        ByteArrayToXInputStreamAdapter xSource = new ByteArrayToXInputStreamAdapter(replacementByteArray);

        // Query for the XDrawPagesSupplier interface
        XDrawPagesSupplier xDrawPagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
        if (xDrawPagesSuppl == null) {
            mylog.error("Cannot get XDrawPagesSupplier interface for Presentation Document???");
            return(1);
        }

        XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
        int pageCount = xDrawPages.getCount();

        try {
            currPage = getDrawPage(xDrawPages, p);
            if (currPage == null) {
                mylog.error("Failed to get page %d, with index number %d!", p+1, p);
                return(2);
            }
            mylog.debug("=== Working with page %d ===", p+1);
            int shapeCount = currPage.getCount();
            mylog.debug("Page %d has %d shapes\n", p+1, shapeCount);

            //XShape xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xDrawPages.getByIndex(p));
            currShape = getPageShape(currPage, s);

            String currType = currShape.getShapeType();
            mylog.debug("Working with shape of type: '%s'", currType);

            if (currType.equalsIgnoreCase("com.sun.star.drawing.GraphicObjectShape")) {
                XPropertySet textProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, currShape);
                String pictureURL = textProps.getPropertyValue("GraphicURL").toString();
                mylog.debug("The URL for the selected shape is '%s'", pictureURL);
            } else {
                mylog.debug("The selected shape is not a GraphicObjectShape!");
                return(3);
            }
        } catch (UnknownPropertyException ex) {
            mylog.error("ReplaceImage: Caught UnknownPropertyException!");
            return 5;
            //Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            mylog.error("ReplaceImage: Caught WrappedTargetException!");
            return 6;
            //Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
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
            mylog.error("Couldn't set image properties");
            return (4);
        }

        return 0;
    }

    private int insertCitationPageText(XShape shape, String text, String fontname, int fontsize,
            boolean prependParagraph, boolean centerParagraph, boolean doubleSpaced)
    {
        try {
            XText xText = (XText) UnoRuntime.queryInterface(XText.class, shape);
            XTextCursor xTextCursor = xText.createTextCursor();
            XTextRange cursorRange = (XTextRange) UnoRuntime.queryInterface(XTextRange.class, xTextCursor);
            XPropertySet xTxtProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);

            xTextCursor.gotoEnd(false);
            if (prependParagraph)
                xText.insertControlCharacter(cursorRange, ControlCharacter.PARAGRAPH_BREAK, false);
            if (centerParagraph)
                xTxtProps.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
            if (doubleSpaced) {
                xTxtProps.setPropertyValue("ParaTopMargin", 200);
            }
            xTxtProps.setPropertyValue("CharFontName", fontname);
            xTxtProps.setPropertyValue("CharHeight", fontsize);
            xText.insertString(cursorRange, text, false);
        } catch (java.lang.Exception ex) {
            mylog.error("insertCitationPageText: Caught Exception: " + ex.getMessage());
            return 1;
        }
        return 0;
    }

    private int insertFullCitation(XComponent xCompDoc,
                                   XShape cShape,
                                   int p,     // Cited image page number
                                   int s,     // Cited image shape number
                                   int i,     // index of citation for this page
                                   String citationString,
                                   String citationLicense,
                                   String citationLicenseBadgeURL,
                                   boolean newParagraph)
    {
        try {

            // Note that the order of operations here is important!!
            StringBuffer fullCite = new StringBuffer();

            fullCite.append("Slide " + p + ", Image " + s + ": ");
            fullCite.append(citationString);
            if (citationLicense != null)
                fullCite.append(", " + citationLicense);
            if (citationLicenseBadgeURL != null)
                fullCite.append(", " + citationLicenseBadgeURL);
            insertCitationPageText(cShape, fullCite.toString(), "Arial", 10, newParagraph, false, true);

        } catch (java.lang.Exception ex) {
            mylog.error("insertFullCitation: Caught Exception!");
            //Logger.getLogger(DecompImpress.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        }
        return 0;
    }


   private XShape addCitationPage(XComponent xCompDoc)
   {

        XDrawPage newPage = null;
        XShape citeShape = null;
        try {
            XDrawPagesSupplier xDrawPagesSuppl = (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
            if (xDrawPagesSuppl == null) {
                //mylog.debug("Failed to get xDrawPagesSuppl from xComp");
                return null;
            }
            XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();
            int currentCount = xDrawPages.getCount();
            newPage = xDrawPages.insertNewByIndex(currentCount);

            XShape titleShape = DecompUtil.createShape(xCompDoc, new Point(1000, 1000), new Size(23000, 500), "com.sun.star.drawing.TextShape");
            newPage.add(titleShape);
            insertCitationPageText(titleShape, "Additional Source Information", "Arial", 24, false, true, false);
            insertCitationPageText(titleShape, "for more information see: http://open.umich.edu/wiki/CitationPolicy", "Arial", 12, true, true, false);

            citeShape = DecompUtil.createShape(xCompDoc, new Point(1000, 3000), new Size(23000, 14000), "com.sun.star.drawing.TextShape");
            newPage.add(citeShape);
        } finally {
            return citeShape;
        }
   }

   public int addCitationPages(XComponent xCompDoc, DecompCitationCollection.DecompCitationCollectionEntry[] entries)
   {
       int i;
       int perPage = 15;
       XShape cShape = null;
       XDrawPage cPage = null;
       DecompCitationCollection.DecompCitationCollectionEntry cpe;

       for (i = 0; i < entries.length; i++)
       {
            if (i % perPage == 0)
                cShape = addCitationPage(xCompDoc);

            cpe = entries[i];
            boolean newParagraph = (i % perPage != 0);
            insertFullCitation(xCompDoc, cShape, cpe.pageNum, cpe.imageNum, (i % perPage), cpe.fullCitation, null, null, newParagraph);
       }
       return 0;
   }
   
   /* XXX  To be completed
      http://www.oooforum.org/forum/viewtopic.phtml?t=45734
    */
    public int insertImageCitation(XComponentContext xContext,
                                   XMultiComponentFactory xMCF,
                                   XComponent xCompDoc,
                                   String citationText,
                                   String citationURL,
                                   int p,
                                   int s)
    {
        try {

            XDrawPage drawPage = getDrawPageByIndex(xCompDoc, p);
            if (drawPage == null)
                return 1;

            //mylog.debug("drawPage.getCount says there are %d objects\n", drawPage.getCount());

            XShapes xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, drawPage);

            // Use the original image location and size to determine the location
            // and size (at least the width?) of the citation information

            Object oOrigImage = xShapes.getByIndex(s);
            XPropertySet xOrigPropSet = (XPropertySet)
                    UnoRuntime.queryInterface(XPropertySet.class, oOrigImage);
//            DecompUtil.printObjectProperties(oOrigImage);
            XShape xOrigImage = (XShape) UnoRuntime.queryInterface(XShape.class, oOrigImage);
//            DecompUtil.printShapeProperties(xOrigImage);

            // Add citation image
            String convertedURL = DecompUtil.getInternalURL(xCompDoc, citationURL, citationURL);

            // Calculate citation image location using original image properties
            Point citeImagePos = DecompUtil.calculateCitationImagePosition(xOrigImage);
            Size citeImageSize = DecompUtil.calculateCitationImageSize(xOrigImage);

            XShape xCIShape = DecompUtil.createShape(xCompDoc, citeImagePos, citeImageSize,
                                    "com.sun.star.drawing.GraphicObjectShape");
            XPropertySet xImageProps = (XPropertySet)
                    UnoRuntime.queryInterface(XPropertySet.class, xCIShape);
            xImageProps.setPropertyValue("GraphicURL", convertedURL);
            xShapes.add(xCIShape);
            mylog.debug("drawPage.getCount nows says there are %d objects\n", drawPage.getCount());

            // Caclulate citation text location using citation image location
            Point citeTextPos = DecompUtil.calculateCitationTextPosition(xCIShape);
            Size citeTextSize = DecompUtil.calculateCitationTextSize(xOrigImage, xCIShape);

            // Add citation text
            XShape xCTShape = DecompUtil.createShape(xCompDoc, citeTextPos, citeTextSize,
                                    "com.sun.star.drawing.TextShape");  // There is also a TextShape?
            xShapes.add(xCTShape);

            XText xText = (XText) UnoRuntime.queryInterface(XText.class, xCTShape);
            XTextCursor xTextCursor = xText.createTextCursor();
            XPropertySet xTxtProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);

            xTxtProps.setPropertyValue("CharFontName", "Arial");
            xTxtProps.setPropertyValue("CharHeight", 8);
//            xTxtProps.setPropertyValue("CharColor", new Integer(0xffffff));
            xText.setString(citationText);
            //xTxtProps.setPropertyValue("HyperLinkURL", "http://open.umich.edu"); // XXX Doesn't work for impress documents.
            //xText.setString("CC-BY");

        } catch (java.lang.Exception ex) {
            mylog.error("insertImageCitation: Caught exception: " + ex.getMessage());
//            Logger.getLogger(OpenOfficeUNODecomposition.class.getName()).log(Level.SEVERE, null, ex);
            return 1;
        }
        return 0;
    }

    public int insertFrontBoilerplate(XComponentContext xContext,
                                      XDesktop xDesktop,
                                      XMultiComponentFactory xMCF,
                                      XComponent destDoc,
                                      String fileName,
                                      String filterName)
    {
        XComponent sourceDoc = null;
        String sourceFileUrl = DecompUtil.fileNameToOOoURL(fileName);
        XDrawPagesSupplier destPagesSuppl;
        XDrawPagesSupplier sourcePagesSuppl;
        PropertyValue props[] = new PropertyValue[0];

        try {
            sourceDoc = DecompUtil.openFileForProcessing(xDesktop, sourceFileUrl);
        } catch (java.lang.Exception ex) {
            mylog.error("insertFrontBoilerplate: Exception while opening source file: " + fileName);
            return 1;
        }

        // Query for the XDrawPagesSupplier interfaces
        sourcePagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, sourceDoc);
        if (sourcePagesSuppl == null) {
            mylog.error("Cannot get XDrawPagesSupplier interface for source Presentation Document???");
            sourceDoc.dispose();
            return 1;
        }
        XDrawPages sourceDrawPages = sourcePagesSuppl.getDrawPages();
        int sourceCount = sourceDrawPages.getCount();

        for (int i = 0; i < sourceCount; i++) {
            XDrawPage sourcePage = getDrawPageByIndex(sourceDoc, i);

//            executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:DuplicatePage", props);

            //executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:GrabControlFocus", props);
            executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:DrawingMode", props);
            executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:SelectAll", props);
            executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:Copy", props);
            //DecompUtil.sleepFor(2);

            //executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:GrabControlFocus", props);
            executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:InsertPage", props);
            executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:Paste", props);
            executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:Paste", props);

            executeDispatch(xContext, xDesktop, xMCF, sourceDoc, ".uno:DeletePage", props);

            //DecompUtil.sleepFor(2);
        }
        sourceDoc.dispose();
/*
        //executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:SelectAll", props);
        //executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:Copy", props);
        //executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:Paste", props);
        //executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:Paste", props);

//        props[0] = new PropertyValue();
//        props[0].Name = "URL";
//        props[0].Value = DecompUtil.fileNameToOOoURL(fileName);

        props[0] = new PropertyValue();
        props[0].Name = "FileName";
        props[0].Value = new String(fileName);

        props[1] = new PropertyValue();
        props[1].Name = "FilterName";
        props[1].Value = new String(filterName);

//        props[2] = new PropertyValue();
//        props[2].Name = "AsLink";
//        props[2].Value = new Boolean(false);

//        executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:ImportFromFile", props);

        executeDispatch(xContext, xDesktop, xMCF, destDoc, ".uno:SelectAll", props);
*/
        return 0;
    }

    // Based on http://www.oooforum.org/forum/viewtopic.phtml?t=48271
    private void executeDispatch(XComponentContext xContext,
                                 XDesktop xDesktop,
                                 XMultiComponentFactory xMCF,
                                 Object pobjDoc,
                                 String cmd,
                                 PropertyValue[] props)
    {
        try {
            XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, pobjDoc);
            XController xController = xModel.getCurrentController();
            XFrame xFrame = xController.getFrame();
            mylog.error("executeDispatch: The name of the frame is " + xFrame.getName());
            if (!xFrame.isActive())
                xFrame.activate();
            XDispatchProvider impressDispatchProvider = (XDispatchProvider) UnoRuntime.queryInterface(XDispatchProvider.class, xFrame);
            Object oDispatchHelper = xMCF.createInstanceWithContext("com.sun.star.frame.DispatchHelper", xContext);
            XDispatchHelper dispatchHelper = (XDispatchHelper) UnoRuntime.queryInterface(XDispatchHelper.class, oDispatchHelper);

//            Object oListener = xMCF.createInstanceWithContext("com.sun.star.awt.XTopWindowListener", xContext);
//            XFrameActionListener xListener = (XFrameActionListener) UnoRuntime.queryInterface(XFrameActionListener.class, oListener);
//            xFrame.addFrameActionListener(xListener);

            printDispatchInfo(cmd, props);
            dispatchHelper.executeDispatch(impressDispatchProvider, cmd, "", 0, props);
            xFrame.contextChanged();
        } catch (Exception e) {
            throw new RuntimeException(e);
        } catch (java.lang.Exception ex) {
            mylog.error("executeDispatch: Yikes!: " + ex.getMessage());
        }
    }

    private void printDispatchInfo(String cmd, PropertyValue[] props)
    {
        mylog.error("Executing dispatch command '%s' with parameters:", cmd);
        if (props == null)
            return;
        for (int i = 0; i < props.length; i++) {
            if (props[i] != null)
                mylog.error("   '%s': '%s'", props[i].Name, props[i].Value);
        }
    }
    
    private XDrawPage getDrawPageByIndex(XComponent xCompDoc, int nIndex)
    {
        XDrawPagesSupplier xDrawPagesSuppl =
                (XDrawPagesSupplier) UnoRuntime.queryInterface(XDrawPagesSupplier.class, xCompDoc);
        if (xDrawPagesSuppl == null) {
            //mylog.debug("Failed to get xDrawPagesSuppl from xComp");
            return null;
        }

        XDrawPages xDrawPages = xDrawPagesSuppl.getDrawPages();

        return getDrawPage(xDrawPages, nIndex);
    }

    private XDrawPage getDrawPage(XDrawPages xDrawPages, int nIndex)
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

    private XShape getPageShape(XDrawPage xDrawPage, int nIndex)
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

    // ASSumes there is a single group shape on the page,
    // therefore returns the first one found...
    private XShape getPageGroupShape(XDrawPage xDrawPage)
    {
        XShape xShape = null;
        int shapeCount = xDrawPage.getCount();
        try {
            for (int i = 0; i < shapeCount; i++) {
                xShape = getPageShape(xDrawPage, i);
                String currType = xShape.getShapeType();
                if (currType.equalsIgnoreCase("com.sun.star.drawing.GroupShape")) {
                    return xShape;
                }
            }
            return null;
        } catch (java.lang.Exception e) {
            e.printStackTrace();
            return null;
        }
    }
    
    private int insertLicenseButton()
    {
        return 0;
    }

}
