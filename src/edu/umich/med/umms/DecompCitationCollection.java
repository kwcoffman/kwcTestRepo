/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 *
 * @author kwc
 */
public class DecompCitationCollection {

    public class DecompCitationCollectionEntry {
        public String fullCitation;
        public int pageNum;
        public int imageNum;

        DecompCitationCollectionEntry(String citation, int page, int image)
        {
            this.fullCitation = citation;
            this.pageNum = page;
            this.imageNum = image;
        }
    }

    private List <DecompCitationCollectionEntry> entrylist;


    DecompCitationCollection()
    {
        entrylist = new ArrayList<DecompCitationCollectionEntry>();
    }

    public int numEntries()
    {
        return entrylist.size();
    }
    
    public int addCitationEntry(String citation, int pageNum, int imageNum)
    {
        DecompCitationCollectionEntry pe = new DecompCitationCollectionEntry(citation, pageNum, imageNum);
        entrylist.add(pe);
        return 0;
    }

    public DecompCitationCollectionEntry[] getDecompCitationCollectionEntryArray() {
        if (entrylist.isEmpty())
            return null;

        // Sort the list
        Collections.sort(entrylist, new Comparator<DecompCitationCollectionEntry>() {
            public int compare(DecompCitationCollectionEntry a, DecompCitationCollectionEntry b) {
                if (a.pageNum == b.pageNum) {
                    return (a.imageNum - b.imageNum);
                } else {
                    return (a.pageNum - b.pageNum);
                }
            }
        });

        return entrylist.toArray(new DecompCitationCollectionEntry[0]);
    }

}
