
package edu.umich.med.umms;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 *
 * The reason for this class is that when we need to remove images from a page,
 * the indexing gets messed up, and subsequent operations (citations, removals)
 * try to operate on the wrong image.  To fix this, we delay _removing_ any
 * image until all the other operations on the file have been performed.
 *
 * Since each operation is independent, we keep track of the needed removals
 * here and then use this collection to removal them when finishing up the
 * operations on the file.
 *
 * Note that the list is returned in reverse order so that we don't mess
 * up the shape order on a page when removing.  For example, page 5 has
 * six shapes and three (marked with "*") need to be replaced:
 *
 *  0 TitleTextShape
 *  1 GraphicObjectShape *
 *  2 OutlinerShape
 *  3 GraphicObjectShape
 *  4 GroupShape *
 *  5 GraphicObjectShape *
 *
 * To replace number 4 GroupShape, we have to remove it and insert a new
 * GraphicObjectShape.  If we remove 4 before doing the replacement of 5,
 * we'll get the wrong image when replacing 5 because it would have beeen
 * reindexed to number 4!
 *
 * @author kwc
 */
public class DecompDelayedRemovalCollection {
    public class DecompDelayedRemovalCollectionEntry {
        public int pageNum;
        public int imageNum;

        DecompDelayedRemovalCollectionEntry(int page, int image)
        {
            this.pageNum = page;
            this.imageNum = image;
        }
    }

    private List <DecompDelayedRemovalCollectionEntry> entrylist;


    DecompDelayedRemovalCollection()
    {
        entrylist = new ArrayList<DecompDelayedRemovalCollectionEntry>();
    }

    public int numEntries()
    {
        return entrylist.size();
    }

    public int addDelayedRemovalEntry(int pageNum, int imageNum)
    {
        DecompDelayedRemovalCollectionEntry pe = new DecompDelayedRemovalCollectionEntry(pageNum, imageNum);
        entrylist.add(pe);
        return 0;
    }

    public DecompDelayedRemovalCollectionEntry[] getDecompDelayedRemovalCollectionEntryArray() {
        if (entrylist.isEmpty())
            return null;

        // Sort the list (in reverse order!)
        Collections.sort(entrylist, new Comparator<DecompDelayedRemovalCollectionEntry>() {
            public int compare(DecompDelayedRemovalCollectionEntry a, DecompDelayedRemovalCollectionEntry b) {
                if (a.pageNum == b.pageNum) {
                    return (b.imageNum - a.imageNum);
                } else {
                    return (b.pageNum - a.pageNum);
                }
            }
        });

        return entrylist.toArray(new DecompDelayedRemovalCollectionEntry[0]);
    }

}
