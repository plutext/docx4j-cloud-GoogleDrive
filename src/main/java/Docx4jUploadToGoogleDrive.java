import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import com.google.api.client.http.ByteArrayContent;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.Permission;

public class Docx4jUploadToGoogleDrive {

	private static String DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

	/**
	 * Upload a wordMLPackage (or presentationML or spreadsheetML pkg) to Google
	 * Drive as a docx
	 * 
	 * @param service
	 * @param wordMLPackage
	 * @param title
	 * @param description
	 * @param shareReadableWithEveryone
	 * @throws IOException
	 * @throws Docx4JException
	 */
	public static void upload(Drive service, OpcPackage wordMLPackage,
			String title, String description, boolean shareReadableWithEveryone)
			throws IOException, Docx4JException {

		// Insert a file
		File body = new File();
		body.setTitle(title);
		body.setDescription(description);
		body.setMimeType(DOCX_MIME);

		ByteArrayContent mediaContent = new ByteArrayContent(DOCX_MIME,
				getDocxAsByteArray(wordMLPackage));

		File file = service.files().insert(body, mediaContent).execute();

		System.out.println("File ID: " + file.getId());

		// https://developers.google.com/drive/v2/reference/permissions/insert
		if (shareReadableWithEveryone) {
			insertPermission(service, file.getId(), null, "anyone", "reader");
			System.out.println(".. shared");
		}

	}

	/**
	 * Insert a new permission.
	 * 
	 * @param service
	 *            Drive API service instance.
	 * @param fileId
	 *            ID of the file to insert permission for.
	 * @return The inserted permission if successful, {@code null} otherwise.
	 */
	private static Permission insertPermission(Drive service, String fileId,
			String value, String type, String role) {
		Permission newPermission = new Permission();

		newPermission.setValue(value);
		newPermission.setType(type);
		newPermission.setRole(role);
		try {
			return service.permissions().insert(fileId, newPermission)
					.execute();
		} catch (IOException e) {
			System.out.println("An error occurred: " + e);
		}
		return null;
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

		// Upload to Google Drive!
		upload(service, wordMLPackage, "my title GP", "my desc", false);

	}

}
