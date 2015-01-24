package example;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Vector;

import javax.mail.Header;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import com.pff.PSTAttachment;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

/**
 * The Objective of this class to provide an example  to Extract messages from the PST files and its attachments.
 * This class will print the result of extracting messages, TO, FROM, CC, subject and body from messages to console,
 * the objective of this is to simplify the indexing process, so if you want to index the content of a single pst file
 * should pipe the console to a file and index it.
 * Some limitations of this code:
 * - Doesn't set up the permissions of the Attachments as the pst file has, it creates with the permissions of the process.
 * <b>But in linux you can use setfacl/set-acl in order to do this change.</b>
 * - Doesn't prepare a specific format for further indexing with search engines like GSA, but you can use the <b>console ouput 
 * to generate a valid format.</b>
 * - Doesn't extract the content of the attachments, meaning the information that is inside the file attached to the email message.
 * - This is not a proper crawler, so won't get a list of PST inside a folder and transform all of them, but here you have 
 * an example that will work on many of the UNIX like systems: 
 * <b>$find /Users/iolalla/workspace/ -name "*.pst" -exec java -cp java-libpst-0.8.1.jar:javax.mail-1.5.2.jar:. example.PSTExtractor {} \;</b>
 * 
 * @Author: Israel Olalla aka: iolalla
*/

public class PSTExtractor {
	
	// This is here to keep the original name without PST so we can use it at all the methods.
	public static String ORIGINAL = "";
		
	public static void main(String[] args) throws Exception {
		PSTExtractor test = new PSTExtractor();
		if (args.length != 1) {
			throw new Exception("Should have an argument with the pst file.");
		}
		if  (args[0].toLowerCase().endsWith(".pst"))  {
			test.processPST(args[0]);
		} else {
			throw new Exception("Should have an argument with the pst file, yep those should end with a .pst extension.");
		}
	}
	
	/**
	 * We need this in order to be able to make the processFolder Recursive... and yes a PST is more than a folder.
	 * 
	 * @param filename of the PST file to process.
	 * @throws Exception
	 */
	
	public void processPST(String filename) throws Exception {
		//let's keep the name without PST so we can create a folder aside the original pst file
		ORIGINAL = filename.substring(0, filename.lastIndexOf(".pst"));
		PSTFile pstFile = new PSTFile(filename);
		processFolder(pstFile.getRootFolder());
	}
	
	/**
	 * This method will extract the .msg files and its attachments
	 * 
	 * @param PSTFolder folders inside the PST
	 * @throws Exception
	 */
	
	public void processFolder(PSTFolder folder) throws Exception {
		//
		ArrayList<PSTAttachment> lista = null;
		// go through the folders...
		if (folder.hasSubfolders()) {
			Vector<PSTFolder> childFolders = folder.getSubFolders();
			for (PSTFolder childFolder : childFolders) {
				processFolder(childFolder);
			}
		}
		// and now the emails for this folder
		if (folder.getContentCount() > 0) {
			PSTMessage email = (PSTMessage)folder.getNextChild();
			while (email != null) {
				//
				if (email.hasAttachments()) {
					lista = new ArrayList<PSTAttachment>();
					for (int i= 0; i < email.getNumberOfAttachments();i++) {
						lista.add(email.getAttachment(i));
					}
				}
				// here is where we send the information we have to the MSG Builder
				createMessage(folder.getDisplayName().toString(), 
						      email.getSubject(),
						      email.getBodyHTML(), 
						      email.getTransportMessageHeaders(),
						      lista);
				// TODO: maybe should replace this block with a Log ??
				System.out.println("Email: "+ email.getDescriptorNodeId());
				System.out.println("Headers: "+ email.getTransportMessageHeaders());
				System.out.println("Subject: "+ email.getSubject());
				System.out.println("Body:"+ email.getBody());
				if (lista != null && lista.size()>0) {
					for(int i = 0; i < lista.size();i++) {
						System.out.println("Attachment FileName: " + lista.get(i).getFilename());
					}
				}
				email = (PSTMessage)folder.getNextChild();
			}
		}
	}
	
	/*
	 * This method will create the message, saves the file and all the attachments included in the email
	 * 
	 * 	 @param String folder where the emails will be stored, mainly the name ofthe pst file without the extension
	 *   @param String name name of the 
	 *   @param String subject
	 *   @param String body
	 *   @param List<PSTAttachment> attachments
	 *   @throws Exception
	 * */

	public void createMessage(String folder, 
							   String subject, 
							   String body, 
							   String headers, 
							   List<PSTAttachment> attachments) throws Exception {
		    //TODO: Review this code, could be shorter
		    //First we create the folder for to save all the files with the same name as the pst file
		    String NEWORIGINAL= ORIGINAL;
		    File folderA = new File(NEWORIGINAL);
		    if (!folderA.exists()) { 
		    	 folderA.mkdir();
		    }
		    // Then we create the path for the folder with all the info from the email's folders 
		    NEWORIGINAL += "/" +URLEncoder.encode(folder,"UTF-8");
			File folderC = new File(NEWORIGINAL);
			if (!folderC.exists()) { 
				folderC.mkdir();
			}
			// Then we create the path for email's folders, based in the subject 
			NEWORIGINAL += "/" +URLEncoder.encode(subject, "UTF-8");
		    // Checks the existence of Files and Attachments
			File folderF = new File(NEWORIGINAL);
			if (!folderF.exists()) { 
				folderF.mkdir();
			}
			// Finally we create the path for the attachments
			if (attachments != null && attachments.size() >0) {
				NEWORIGINAL += "/files";
				File filesF = new File(folderF.getAbsolutePath()+"/files");
				if (!filesF.exists()) {
					filesF.mkdir();
				}
			}
			
	        Message message = new MimeMessage(Session.getInstance(System.getProperties()));
	        // The objective of this block is to get the message headers from the PST to the msg
	        // Those headers will include the to, from, CC and important information to track the email
	        InternetHeaders headersList = new InternetHeaders(new ByteArrayInputStream(headers.getBytes(StandardCharsets.UTF_8)));
	        Enumeration<Header> listOfHeaders = headersList.getAllHeaders();
	        while (listOfHeaders.hasMoreElements()) { 
	        	Header test = listOfHeaders.nextElement();
	        	if (test.getName() != null)
	          	message.setHeader(test.getName(), test.getValue());
	        }
	        //Then we set the subject
	        message.setSubject(subject);
	        // Now we set the content and attachments
	        // create the message part 
	        MimeBodyPart content = new MimeBodyPart();
	        // fill message
	        content.setText(body);
	        Multipart multipart = new MimeMultipart();
	        multipart.addBodyPart(content);
	        if (attachments != null && attachments.size() >0) {
		        for (PSTAttachment attachment : attachments) {
		        	InputStream fin= (InputStream)attachment.getFileInputStream();
		        	String FILEOUT= NEWORIGINAL + "/" + attachment.getFilename();
		        	File file = new File(FILEOUT);
		        	FileOutputStream fos = new FileOutputStream(file);
		        		int i=0;  
		        		while((i=fin.read())!=-1){  
		        			fos.write((byte)i);  
		        		}  
		        	fin.close();
		        	fos.close();
		        	// TODO: should replace this with Log call
		        	if (file.exists()) System.out.print(" Created File: " + file.getName());
		        }
	        }
	        // integration of the attachments with the new Message
	        message.setContent(multipart);
	        // store the file with the message in the msg format
	        // Could be done with a format html/txt.. but would loose attachments
	        message.writeTo(new FileOutputStream(new File(folderF, URLEncoder.encode(subject, "UTF-8")+".msg")));
	}
}