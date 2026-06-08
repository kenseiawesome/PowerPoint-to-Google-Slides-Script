function pptxToSlides() {
  const SOURCE_FOLDER_ID = ''; //Add the drive ID of the source folder here
  const DESTINATION_FOLDER_ID   = ''; // Add the Drive ID of the target folder here
  // These CAN be the same if you want them in the same folder

  const sourceFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
  const destinationFolder   = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

  const pptxFiles = sourceFolder.getFilesByType(MimeType.MICROSOFT_POWERPOINT);

  while (pptxFiles.hasNext()) {
    const pptx = pptxFiles.next();
    const blob = pptx.getBlob();
    
    const newFileName = pptx.getName().replace(/\.pptx?$/i, '');

    const resource = {
      name: newFileName,
      mimeType: 'application/vnd.google-apps.presentation',
      parents: [DESTINATION_FOLDER_ID]
    };

    const converted = Drive.Files.create(resource, blob, {
      supportsAllDrives: true
    });

    Logger.log(`Converted: ${pptx.getName()} → ${converted.name}`);
  }

  Logger.log('All PPTX files processed.');
}