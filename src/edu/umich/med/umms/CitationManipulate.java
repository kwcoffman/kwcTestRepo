/*
 * CitationManipulate.java
 *
 * This class manipulates citation strings.
 */

package edu.umich.med.umms;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.ArrayList;


/**
 *
 * @author lotia
 */
public class CitationManipulate {
    private static final String URL_EXP = "\\(?\\b(http[s]?://|www[.])[-A-Za-z0-9+&@#/%?=~_()|!:,.;]*[-A-Za-z0-9+&@#/%=~_()|]";
    private static final Pattern URL_PATTERN = Pattern.compile(URL_EXP);

    private static final String LIC_EXP = "http://.*\\.?creativecommons.org/licenses/([^/]*)/([^/]*)/?";
    private static final Pattern LIC_PATTERN = Pattern.compile(LIC_EXP);

    private static final String CC_LIC_LOC = "http://i.creativecommons.org/l/";
    private static final String CC_BADGE_IMG_NAME = "88x31.png";

    private ArrayList<String> foundURLs = new ArrayList<String>();
    private Matcher urlMatcher;
    private Matcher licMatcher;


    /**
     * The parameterless constructor. To enable use when multiple citations are
     * going to be tested in a single object instantiation.
     */
    public CitationManipulate() {
        return;
    }


    /**
     * The url matched against the regexp on object instantiation. This can be
     * used in cases where a single string is matched per object instantiation.
     * 
     * @param citationString
     */
    public CitationManipulate(String citationString) {
        detectURLs(citationString);
    }
    

    /**
     * Detect the URLs in the passed citation.
     *
     * This method matches the URL_EXP against the passed string and stores any
     * found URLs for later processing.
     *
     * @param citationString
     */
    public void detectURLs(String citationString) {
        foundURLs.clear();
        urlMatcher = URL_PATTERN.matcher(citationString);
        while (urlMatcher.find()) {
            foundURLs.add(urlMatcher.group());
        }
    }


    /**
     * Return the citation text with URLs removed.
     *
     * @return String citation text with URLs removed.
     */
    public String getShortCitation() {
        return urlMatcher.replaceAll("").replace(" ,", " ").replace("  ", "");
    }


    /**
     * Return the URL to the badge license based on the specified license URL.
     * This method currently only works for CC license badges.
     *
     * @return String URL of CC license badge.
     */
    /* TODO: use boolean variables and simplify conditional.
     * TODO: provide functionality for licenses other than CC.
     */
    public String getBadgeURL() {
        String badgeURL = null;
        if (!foundURLs.isEmpty()) {
            for (String detectedURL : foundURLs) {
                licMatcher = LIC_PATTERN.matcher(detectedURL);
                if (licMatcher.find() && (licMatcher.groupCount() == 2)) {
                    badgeURL = CC_LIC_LOC + licMatcher.group(1) + "/" +
                            licMatcher.group(2) + "/" + CC_BADGE_IMG_NAME;
                }
            }
        }
        return badgeURL;
    }

//    /**
//     * A test method for figuring out if the code is working.
//     * @param args the command line arguments
//     */
//    public static void main(String[] args) {
//        CitationManipulate stuff = new CitationManipulate();
//
//        String[] TEST_STRINGS = {
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by-nd/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by-nc-nd/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by-nc/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by-nc-sa/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by-sa/2.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by/3.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by-nd/3.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by-nc-nd/3.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by-nc/3.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by-nc-sa/3.0/",
//        "Jot Powers, \"Bounty Hunter\", Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 3.0 http://creativecommons.org/licenses/by-sa/3.0/",
//        "Jot Powers, Bounty Hunter, Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, CC: BY-SA 2.0 http://creativecommons.org/licenses/by/2.0/",
//        "Jot Powers, Bounty Hunter, Wikimedia Commons, http://commons.wikimedia.org/wiki/File:Bounty_hunter_2.JPG, http://creativecommons.org/licenses/by/2.0/ CC: BY-SA 2.0" ,
//    };
//
//        int numstrings = TEST_STRINGS.length;
//        for (int x = 0; x < numstrings; x++) {
//            stuff.detectURLs(TEST_STRINGS[x]);
//            System.out.println("The full citiation is: " + TEST_STRINGS[x]);
//            System.out.println("The short citation is: " + stuff.getShortCitation());
//            System.out.println("The license URL is: " + stuff.getBadgeURL());
//        }
//    }

}
