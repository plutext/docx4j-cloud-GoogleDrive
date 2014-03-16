docx4j-cloud-GoogleDrive
========================

docx4j integration with Google Drive

This project demonstrates:

* Upload a wordMLPackage (or presentationML or spreadsheetML pkg) to Google Drive as a docx
* Download a file from Google Drive as a docx4j WordprocessingMLPackage
* Convert a WordprocessingMLPackage to the specified output format, using Google Drive

Getting started: Clone the project, and set it up using Maven in your IDE.  

Enabling the Drive API: set up a project and application in the Developers Console:
* press the red "CREATE NEW CLIENT ID" button, then choose application type "Installed Application"; I then chose subtype "Other" 
* hit the "Download JSON" button; save it as client_secret.json in your project dir

Now you can run the classes in src/main/java, for example: Docx4jUploadToGoogleDrive

It ought to say something like:

   Please open the following URL in your browser then type the authorization code:
      https://accounts.google.com/o/oauth2/auth?access_type=online&client_id=622239...
      
Paste the auth code into your IDE's console (System.in, probably the same place which displayed the above message) 
then press enter.  If you aren't logged into your Google account in your browser, its at this point that you'll be
asked to log in.

