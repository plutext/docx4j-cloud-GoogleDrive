import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import com.google.api.client.http.ByteArrayContent;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpResponse;
import com.google.api.client.util.IOUtils;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.Drive.Files.Insert;
import com.google.api.services.drive.model.File;

public class GoogleDriveConvertFromDocxTo {

	private static String DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

	/**
	 * Convert a WordprocessingMLPackage to the specified output format, using Google Drive
	 * 
	 * @param di
	 * @param wordMLPackage
	 * @param outType
	 * @return
	 * @throws IOException
	 * @throws Docx4JException
	 */
	public static InputStream convert(Drive di,
			WordprocessingMLPackage wordMLPackage, OutputMime outType)
			throws IOException, Docx4JException {

		ByteArrayContent mediaContent = new ByteArrayContent(DOCX_MIME,
				getDocxAsByteArray(wordMLPackage));
		Insert insert = di.files().insert(new File(), mediaContent);
		insert.setConvert(true); // convert this file to the corresponding Google Docs format.
		File file = insert.execute();

		System.out.println("File ID: " + file.getId());

		System.out.println(file.getMimeType());

		String url = null;
		if (outType == OutputMime.GDOC) {
			// Just download it
			url = file.getDownloadUrl();
		} else {
			Map<String, String> exportLinks = file.getExportLinks();
			url = exportLinks.get(outType.getMime());
		}
		if (url == null) {
			throw new IOException(outType.getMime() + " not supported");
		} else {

			HttpRequest req = di.getRequestFactory().buildGetRequest(
					new GenericUrl(url));
			HttpResponse res = req.execute();

			if (res.getStatusCode() == 200) {
				InputStream contentStream = res.getContent();

				// delete
				di.files().trash(file.getId()).execute();

				return contentStream;

			} else {
				throw new IOException("response code: " + res.getStatusCode());
			}

		}

	}

	enum OutputMime {

		// see https://developers.google.com/drive/web/integrate-open#open_files_using_the_open_with_contextual_menu
		GDOC("application/vnd.google-apps.document"), 
		HTML("text/html"), 
		TXT("text/plain"), 
		PDF("application/pdf"), 
		RTF("application/rtf"), 
		DOCX("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), 
		ODT("application/vnd.oasis.opendocument.text");

		private String mime;

		public String getMime() {
			return mime;
		}

		private OutputMime(String mime) {
			this.mime = mime;
		}
	}

	public static byte[] getDocxAsByteArray(OpcPackage opcPackage)
			throws Docx4JException {

		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		Save saver = new Save(opcPackage);
		saver.save(baos);

		return baos.toByteArray();
	}

	public static void main(String[] args) throws IOException, Docx4JException {

		// Create a new authorized API client
		// Do this first so things fail fast until you have it configured
		// correctly
		Drive service = CredentialsViaCommandLine.getDrive();

		// For this sample, create a docx from scratch
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
				.createPackage();
		MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
		mdp.addStyledParagraphOfText("Title", "Example 1");

		OutputStream outputStream = new FileOutputStream(new java.io.File(
				System.getProperty("user.dir") + "/out2.PDF"));

		IOUtils.copy(convert(service, wordMLPackage, OutputMime.PDF),
				outputStream);

		outputStream.close();

	}

}
