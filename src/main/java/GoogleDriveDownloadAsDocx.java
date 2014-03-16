import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpResponse;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.Drive.Files;
import com.google.api.services.drive.Drive.Files.Copy;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;

public class GoogleDriveDownloadAsDocx {

	private static String DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

	/**
	 * Download a file from Google Drive as a docx4j WordprocessingMLPackage 
	 * 
	 * @param service
	 * @param title
	 * @return
	 * @throws IOException
	 * @throws Docx4JException
	 */
	public static WordprocessingMLPackage downloadByTitle(Drive service, String title)
			throws IOException, Docx4JException {
		
		
		File file = getFileByTitle( service,  title);
		
		if (file==null) {
			System.out.println("No file at Google Drive with title '" + title + "'");	
			return null;
		}


		System.out.println("File ID: " + file.getId());

		System.out.println(file.getMimeType());
		
		String url = null;
		File googleDoc = null;
		if (file.getMimeType().equals(DOCX_MIME)) {
			
			// Its a docx; just download it
			System.out.println("Fetching the docx");
			url = file.getDownloadUrl();
			
		} else if (file.getExportLinks()!=null
				&& file.getExportLinks().get(DOCX_MIME)!=null) {
			
			System.out.println("Fetching the Google Doc as a docx");
			Map<String, String> exportLinks = file.getExportLinks();
			url = exportLinks.get(DOCX_MIME);
			
		} else {
			
			// Not a docx or a Google Doc; try to convert
			// First convert to Google Doc
			System.out.println("Converting to docx via Google Doc");
			File copiedFile = new File();
		    copiedFile.setTitle(file.getTitle());
		      Copy copy =  service.files().copy(file.getId(), copiedFile);
		      copy.setConvert(true);
		     googleDoc = copy.execute();
			
			Map<String, String> exportLinks = googleDoc.getExportLinks();
			url = exportLinks.get(DOCX_MIME);
		      
		}

		if (url == null) {
			throw new IOException( "Now download url for '" + title + "' of type " + file.getMimeType());
		} else {

			HttpRequest req = service.getRequestFactory().buildGetRequest(
					new GenericUrl(url));
			HttpResponse res = req.execute();

			if (res.getStatusCode() == 200) {
				InputStream contentStream = res.getContent();

				// delete
				if (googleDoc!=null) {
					service.files().trash(googleDoc.getId()).execute();
				}

				System.out.println("The docx4j bit..");							
				return Docx4J.load(contentStream);

			} else {
				throw new IOException("response code: " + res.getStatusCode());
			}

		}

	}
	
	  /**
	   * Retrieve a list of File resources.
	   *
	   * @param service Drive API service instance.
	   * @return List of File resources.
	   */
	  private static File getFileByTitle(Drive service, String title) throws IOException {
		  
	    Files.List request = service.files().list();
	    request.setQ("title = '" + title +"'");

	        FileList files = request.execute();
	        
	        if (files.getItems().size()==0) {
				System.out.println("No file at Google Drive with title '" + title + "'");	
				return null;
	        }
	        if (files.getItems().size()>1) {
				System.out.println("Multiple files with title '" + title + "'; using first.");	
	        }

	       return files.getItems().get(0);
	  }	


	public static void main(String[] args) throws IOException, Docx4JException {

		// Create a new authorized API client
		// Do this first so things fail fast until you have it configured
		// correctly
		Drive service = CredentialsViaCommandLine.getDrive();

		WordprocessingMLPackage wmlPkg = downloadByTitle(service, "html_test.htm");
		
		if (wmlPkg==null) {
			System.out.println("FAILED!");				
		} else {
			System.out.println("Success");							
		}
	}

}
