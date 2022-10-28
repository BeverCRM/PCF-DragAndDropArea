import { IInputs } from '../generated/ManifestTypes';
import FileHelper from '../helpers/FileHelper';

let _context: ComponentFramework.Context<IInputs>;

const notificationOptions = {
  errorsCount: 0,
  importedSucsessCount: 0,
  filesCount: 0,
  details: '',
  message: '',
};

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
  },

  async getEntitySetName(entityTypeName: string) {
    const entityMetadata = await _context.utils.getEntityMetadata(entityTypeName);
    return entityMetadata.EntitySetName;
  },

  async getEntityDisplayName() {
    // @ts-ignore
    const { entityTypeName } = _context.page;
    const entityMetadata = await _context.utils.getEntityMetadata(entityTypeName);
    return entityMetadata._displayName;
  },

  async hasNotes() {
    // @ts-ignore
    const contextPage = _context.page;
    const entityMetadataResponse =
     await fetch(`${contextPage.getClientUrl()}/api/data/v9.0/EntityDefinitions(LogicalName='` +
     `${contextPage.entityTypeName}')`);
    const entityMetadata = await entityMetadataResponse.json();
    return entityMetadata.HasNotes;
  },

  async uploadFile(file: File, filesCount: number) {
    try {
      notificationOptions.filesCount = filesCount;
      const buffer: ArrayBuffer = await FileHelper.readFileAsArrayBufferAsync(file);
      const body: string = FileHelper.arrayBufferToBase64(buffer);
      // @ts-ignore
      const { entityTypeName, entityId } = _context.page;
      const entitySetName = await this.getEntitySetName(entityTypeName);

      const data: any = {
        'subject': _context.parameters.noteTitle.raw,
        'notetext': _context.parameters.noteDescription.raw,
        'filename': file.name,
        'documentbody': body,
        'objecttypecode': entityTypeName,
      };

      data[`objectid_${entityTypeName}@odata.bind`] = `/${entitySetName}(${entityId})`;

      await _context.webAPI.createRecord('annotation', data);
      notificationOptions.importedSucsessCount += 1;
    }
    catch (ex: any) {
      console.error(ex.message);
      notificationOptions.details += `File Name ${file.name} Error message ${ex.message}`;
      notificationOptions.errorsCount += 1;
    }
  },

  refreshTimeline() {
    // @ts-ignore
    parent.Xrm.Page.getControl('Timeline')?.refresh();
    this.showNotificationPopup();
  },

  showNotificationPopup() {
    if (notificationOptions.errorsCount === 0) {
      const message = notificationOptions.filesCount === 1
        ? `${notificationOptions.importedSucsessCount} of ` +
        `${notificationOptions.importedSucsessCount} file imported successfully`
        : `${notificationOptions.importedSucsessCount} of ` +
        `${notificationOptions.importedSucsessCount} files imported successfully`;

      _context.navigation.openConfirmDialog({ text: message });
      notificationOptions.importedSucsessCount = 0;
    }
    else {
      notificationOptions.message = notificationOptions.filesCount === 1
        ? `${notificationOptions.errorsCount} 
        of ${notificationOptions.filesCount} file errored during import`
        : `${notificationOptions.errorsCount} 
        of ${notificationOptions.filesCount} files errored during import`;

      _context.navigation.openErrorDialog(notificationOptions);
      notificationOptions.errorsCount = 0, notificationOptions.importedSucsessCount = 0;
      notificationOptions.details = '';
    }
  },
};
