import java.io.File;
import java.io.IOException;
import java.util.Collection;
import java.util.Iterator;
import java.util.Locale;

import org.tmatesoft.svn.core.SVNException;
import org.tmatesoft.svn.core.SVNLogEntry;
import org.tmatesoft.svn.core.SVNURL;
import org.tmatesoft.svn.core.auth.ISVNAuthenticationManager;
import org.tmatesoft.svn.core.internal.io.dav.DAVRepositoryFactory;
import org.tmatesoft.svn.core.io.SVNRepository;
import org.tmatesoft.svn.core.io.SVNRepositoryFactory;
import org.tmatesoft.svn.core.wc.SVNWCUtil;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


/**
 * This class is used to 
 *
 * @author skumar
 * Version 1.0
 */
public class TestNotes {

    private String filename;
    /**
	 * @return the filename
	 */
	public String getFilename() {
		return filename;
	}
	/**
	 * @param filename the filename to set
	 */
	public void setFilename(String filename) {
		this.filename = filename;
	}
	private WritableWorkbook workbook;
    private Workbook workbook2;
    private String IT_SHEET = "IT";
    private String ST_SHEET = "ST";
    private String UAT_SHEET = "UAT";
    private String PROD_SHEET = "PRODUCTION";
    private String CR_SHEET = "CR";
    private String SRS_SHEET = "SRS";
    private String CONFIG_SHEET = "CONFIG";
    private String OTHERS_SHEET = "OTHERS";
    private String TAG_RELEASE = "TAGRELEASE";
	private long startRevision ;
	private long endRevision ;
	private String svnURLstr = null;
	private String svnUser = null;
	private String svnPwd = null;
	/**
	 * 
	 */
	public TestNotes(String filename) {
		this.filename = filename;
	}
	/**
	 * 
	 */
	public TestNotes() {
	}
	/**
	 * @param args
	 * @throws IOException 
	 * @throws BiffException 
	 * @throws SVNException 
	 */
	public static void main(String[] args) throws BiffException, IOException, SVNException {
		TestNotes jdemo = new TestNotes();
		if(args.length<5) {
			System.out.println("Usage : java -jar TagReleaseNotes.jar [SVN URL] [SVN User name] [SVN password] [start revision] [end revision] [filename]");
			return;
			//Usage : java -jar TagReleaseNotes.jar BIGWv1_LaybySelfServe 472151 472151 24695 24980 tag.xls
			//Usage : java -jar TagReleaseNotes.jar DSAUv2-CR-Floodlight-Jun2012-02.00.00.00 472151 472151 19719 24804 tag.xls
		}
		String defaultSvnUrl = "http://172.20.191.147/MCRCommerce/Tag/";
		
		if(args[0]!=null && !"".equals(args[0].trim())) {
			if(args[0].startsWith("http://")||args[0].startsWith("svn://")) {
				jdemo.svnURLstr = args[0];
			}
			else {
				System.out.println("SVN server is not mentioned using default server : "+defaultSvnUrl);
				jdemo.svnURLstr = defaultSvnUrl+args[0];
			}
		}
		else {
			System.out.println("Tag name not mentioned, Please refer usage");
		}
		jdemo.svnUser = args[1];
		jdemo.svnPwd = args[2];
		jdemo.startRevision = (args[3]==null?0:Long.parseLong(args[3]));
		jdemo.endRevision = (args[4]==null?0:Long.parseLong(args[4]));
		if(args[5]==null || "".equals(args[5].trim())) {
			jdemo.filename = "tag.xls";
		}
		else {
			jdemo.filename = args[5];
		}
		try {
			jdemo.write();
		} catch (WriteException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public void write() throws IOException, WriteException, BiffException, SVNException {

        WorkbookSettings ws = new WorkbookSettings();
        WritableSheet s1;
        WritableSheet s2;
        WritableSheet s3;
        WritableSheet s4;
        WritableSheet s5;
        WritableSheet s6;
        WritableSheet s7;
        WritableSheet s8;
        WritableSheet s9;
        ws.setLocale(new Locale("en", "EN"));
        try {
        	File fil = new File(getFilename());
        	if(!fil.exists()) {
            	workbook = Workbook.createWorkbook(new File(filename), ws);
            	s1 = workbook.createSheet(TAG_RELEASE, 0);
                s2 = workbook.createSheet(IT_SHEET, 1);
                s3 = workbook.createSheet(ST_SHEET, 2);
                s4 = workbook.createSheet(UAT_SHEET, 3);
                s5 = workbook.createSheet(PROD_SHEET, 4);
                s6 = workbook.createSheet(CR_SHEET, 5);
                s7 = workbook.createSheet(SRS_SHEET, 6);
                s8 = workbook.createSheet(OTHERS_SHEET, 7);
                s9 = workbook.createSheet(CONFIG_SHEET, 8);
                createTegReleaseSheet(s1);
                createITSheet(s2);
                createSTUATOTHSheet(s3);
                createSTUATOTHSheet(s4);
                createSTUATOTHSheet(s5);
                createCRSheet(s6);
                createCRSheet(s7);
                createSTUATOTHSheet(s8);
                workbook.write();
                workbook.close();
            }
        }
        catch (IOException ioe) {
        	ioe.printStackTrace();
        }
        getSVNLog();
    
	}
	public void getSVNLog() throws SVNException, WriteException, IOException, BiffException {

		System.out.println("Starting...");
		DAVRepositoryFactory.setup();
		SVNURL svnURL = SVNURL.parseURIEncoded(svnURLstr);
		SVNRepository repository = SVNRepositoryFactory.create(svnURL);
		ISVNAuthenticationManager authManager = SVNWCUtil.createDefaultAuthenticationManager(svnUser, svnPwd);
		repository.setAuthenticationManager(authManager);
		Collection logEntries = repository.log(new String[] { "" }, null, startRevision, endRevision, false, false);
		Iterator itr = logEntries.iterator();
		WritableSheet wrs;
		while(itr.hasNext()) {
			workbook2 = Workbook.getWorkbook(new File(getFilename()));
	        //workbook = Workbook.createWorkbook(new File(filename), ws);
	        workbook = Workbook.createWorkbook(new File(getFilename()), workbook2);
			SVNLogEntry svnlog = (SVNLogEntry)itr.next();
			String msg = svnlog.getMessage();
			String[] msgStr = msg.split(":");
			if(msgStr!=null && IT_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(IT_SHEET);
			}
			else if(msgStr!=null && ST_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(ST_SHEET);
			}
			else if(msgStr!=null && UAT_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(UAT_SHEET);
			}
			else if(msgStr!=null && CR_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(CR_SHEET);
			}
			else if(msgStr!=null && SRS_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(SRS_SHEET);
			}
			else if(msgStr!=null && PROD_SHEET.equalsIgnoreCase(msgStr[0])) {
				wrs = workbook.getSheet(PROD_SHEET);
			}
			else {
				wrs = workbook.getSheet(OTHERS_SHEET);
			}
			int row = wrs.getRows();
	        System.out.println("Row : "+row+" Sheet : "+wrs.getName()+" # :"+msgStr[0]+ " Message : "+msgStr[2]);
			WritableCellFormat wrappedText = new WritableCellFormat(WritableWorkbook.ARIAL_10_PT);
	        wrappedText.setWrap(true);
	        Label l = new Label(0, row, msgStr[1], wrappedText);
	        wrs.addCell(l);
	        l = new Label(2, row, msgStr[2], wrappedText);
	        wrs.addCell(l);
	        workbook.write();
	        workbook.close();
			//System.out.println("Revision : "+svnlog.getRevision()+" Date : "+svnlog.getDate()+" Message : "+svnlog.getMessage()+" Path : "+svnlog.getChangedPaths());
		}
		System.out.println("Done....");
	}
    private void createTegReleaseSheet(WritableSheet s) throws WriteException {
    	s.setColumnView(2, 16);
    	s.setColumnView(3, 20);
    	s.setColumnView(4, 18);
    	s.setColumnView(5, 15);
    	s.setColumnView(6, 13);
	    WritableFont arial14ptBold = null;
        arial14ptBold = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
        WritableCellFormat arial14BoldFormat = new WritableCellFormat(arial14ptBold);
        arial14BoldFormat.setAlignment(jxl.format.Alignment.CENTRE);
        Label l = null;
        l = new Label(3, 3, "TAG RELEASE NOTES", arial14BoldFormat);
        s.addCell(l);
        WritableImage wi = new WritableImage(3.0D, 6D, 0.6D, 3D, new File("resources/tcsicon.PNG"));
        s.addImage(wi);
        wi = new WritableImage(3.0D, 10D, 1.8D, 1D, new File("resources/tcslabel.PNG"));
        s.addImage(wi);
	    WritableFont arial10pt = null;
        arial10pt = new WritableFont(WritableFont.ARIAL, 10);
        WritableCellFormat arial10Format = new WritableCellFormat(arial10pt);
        arial10Format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        l = new Label(2, 13, "Project:", arial10Format);
        s.addCell(l);
        s.mergeCells(3, 13, 6, 13);
        l = new Label(3, 13, "", arial10Format);
        s.addCell(l);
        l = new Label(2, 14, "Project Number:", arial10Format);
        s.addCell(l);
        s.mergeCells(3, 14, 6, 14);
        l = new Label(3, 14, "", arial10Format);
        s.addCell(l);
        l = new Label(2, 15, "Enhancement:", arial10Format);
        s.addCell(l);
        s.mergeCells(3, 15, 6, 15);
        l = new Label(3, 15, "", arial10Format);
        s.addCell(l);
        l = new Label(2, 16, "Author:", arial10Format);
        s.addCell(l);
        s.mergeCells(3, 16, 6, 16);
        l = new Label(3, 16, "", arial10Format);
        s.addCell(l);
        l = new Label(2, 17, "Version:", arial10Format);
        s.addCell(l);
        l = new Label(3, 17, "", arial10Format);
        s.addCell(l);
        l = new Label(4, 17, "Date: ", arial10Format);
        s.addCell(l);
        l = new Label(5, 17, "", arial10Format);
        s.addCell(l);
        s.mergeCells(5, 17, 6, 17);
        l = new Label(5, 17, "", arial10Format);
        s.addCell(l);
        l = new Label(2, 18, "Reviewer:", arial10Format);
        s.addCell(l);
        l = new Label(3, 18, "", arial10Format);
        s.addCell(l);
        l = new Label(4, 18, "Review Date", arial10Format);
        s.addCell(l);
        s.mergeCells(5, 18, 6, 18);
        l = new Label(5, 18, "", arial10Format);
        s.addCell(l);
	    
	    WritableFont arial14ptBoldLeft = null;
	    arial14ptBoldLeft = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
        WritableCellFormat arial14ptBoldLeftFormat = new WritableCellFormat(arial14ptBoldLeft);
        l = new Label(2, 21, "Approvals & distribution", arial14ptBoldLeftFormat);
        s.addCell(l);
        
        l = new Label(2, 23, "Name", arial10Format);
        s.addCell(l);
        l = new Label(3, 23, "Position & Department", arial10Format);
        s.addCell(l);
        l = new Label(4, 23, "Approver/ Reviewer/ Info only", arial10Format);
        s.addCell(l);
        l = new Label(5, 23, "Signature / Record of email sign-off", arial10Format);
        s.addCell(l);
        l = new Label(6, 23, "Date", arial10Format);
        s.addCell(l);
        
        l = new Label(2, 35, "Tag Details", arial14ptBoldLeftFormat);
        s.addCell(l);

	    WritableFont arial10ptBold = null;
	    arial10ptBold = new WritableFont(WritableFont.ARIAL, 10);
        WritableCellFormat arial10ptBoldFormat = new WritableCellFormat(arial10ptBold);
        arial10ptBoldFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        s.mergeCells(2, 36, 3, 36);
        l = new Label(2, 36, "Tag Name:", arial10ptBoldFormat);
        s.addCell(l);
        s.mergeCells(4, 36, 6, 36);
        l = new Label(4, 36, "", arial10Format);
        s.addCell(l);
        s.mergeCells(2, 37, 3, 37);
        l = new Label(2, 37, "Deployed Application:", arial10ptBoldFormat);
        s.addCell(l);
        s.mergeCells(4, 37, 6, 37);
        l = new Label(4, 37, "", arial10Format);
        s.addCell(l);
        s.mergeCells(2, 38, 3, 38);
        l = new Label(2, 38, "Revision:", arial10ptBoldFormat);
        s.addCell(l);
        s.mergeCells(4, 38, 6, 38);
        l = new Label(4, 38, "", arial10Format);
        s.addCell(l);
        s.mergeCells(2, 39, 3, 39);
        l = new Label(2, 39, "Last Deployed Production Tag:", arial10ptBoldFormat);
        s.addCell(l);
        s.mergeCells(4, 39, 6, 39);
        l = new Label(4, 39, "", arial10Format);
        s.addCell(l);
        //arial10BoldFormat.setBackground(jxl.format.Colour.GRAY_25);
        //arial10BoldFormat.setWrap(true);
        //arial10BoldFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
    }
    private void createITSheet(WritableSheet s) throws WriteException {
	    WritableCellFormat wrappedText = new WritableCellFormat(WritableWorkbook.ARIAL_10_PT);
	    wrappedText.setWrap(true);
	    s.setColumnView(0, 14);
	    s.setColumnView(1, 14);
	    s.setColumnView(2, 90);
	    s.setColumnView(3, 21);
	    s.setColumnView(4, 21);
	    Label l = null;
	    WritableFont arial10ptBold = null;
        arial10ptBold = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
        WritableCellFormat arial10BoldFormat = new WritableCellFormat(arial10ptBold);
        arial10BoldFormat.setBackground(jxl.format.Colour.GRAY_25);
        arial10BoldFormat.setWrap(true);
        arial10BoldFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        l = new Label(0, 3, "Defect No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(1, 3, "Ref No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(2, 3, "Defect Description", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(3, 3, "WEB/CC/Batch", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(4, 3, "Released in Tag Version", arial10BoldFormat);
	    s.addCell(l);
    }
    private void createSTUATOTHSheet(WritableSheet s) throws WriteException {
	    WritableCellFormat wrappedText = new WritableCellFormat(WritableWorkbook.ARIAL_10_PT);
	    wrappedText.setWrap(true);
	    s.setColumnView(0, 14);
	    s.setColumnView(1, 14);
	    s.setColumnView(2, 90);
	    s.setColumnView(3, 21);
	    s.setColumnView(4, 21);
	    s.setColumnView(5, 21);
	    Label l = null;
	    WritableFont arial10ptBold = null;
        arial10ptBold = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
        WritableCellFormat arial10BoldFormat = new WritableCellFormat(arial10ptBold);
        arial10BoldFormat.setBackground(jxl.format.Colour.GRAY_25);
        arial10BoldFormat.setWrap(true);
        arial10BoldFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        l = new Label(0, 3, "Defect No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(1, 3, "Ref No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(2, 3, "Defect Description", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(3, 3, "WEB/CC/Batch", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(4, 3, "Test Owner", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(5, 3, "Released in Tag Version", arial10BoldFormat);
	    s.addCell(l);
    }
    private void createCRSheet(WritableSheet s) throws WriteException {
	    WritableCellFormat wrappedText = new WritableCellFormat(WritableWorkbook.ARIAL_10_PT);
	    wrappedText.setWrap(true);
	    s.setColumnView(0, 14);
	    s.setColumnView(1, 14);
	    s.setColumnView(2, 90);
	    s.setColumnView(3, 21);
	    if(CR_SHEET.equals(s.getName())) {
	    	s.setColumnView(4, 50);
	    }
	    else {
	    	s.setColumnView(4, 25);
	    }
	    s.setColumnView(5, 21);
	    s.setColumnView(6, 21);
	    Label l = null;
	    WritableFont arial10ptBold = null;
        arial10ptBold = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
        WritableCellFormat arial10BoldFormat = new WritableCellFormat(arial10ptBold);
        arial10BoldFormat.setBackground(jxl.format.Colour.GRAY_25);
        arial10BoldFormat.setWrap(true);
        arial10BoldFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        l = new Label(0, 3, "Defect No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(1, 3, "Ref No", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(2, 3, "Defect Description", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(3, 3, "WEB/CC/Batch", arial10BoldFormat);
	    s.addCell(l);
	    if(CR_SHEET.equals(s.getName())) {
	        l = new Label(4, 3, "Reference Doc Location", arial10BoldFormat);
		    s.addCell(l);
	    }
	    else {
	        l = new Label(4, 3, "In PROD", arial10BoldFormat);
		    s.addCell(l);
	    }
        l = new Label(5, 3, "Test Owner", arial10BoldFormat);
	    s.addCell(l);
        l = new Label(6, 3, "Released in Tag Version", arial10BoldFormat);
	    s.addCell(l);
    }
}
