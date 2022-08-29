import { IInputs } from '../generated/ManifestTypes';
import FileHelper from '../helpers/FileHelper';

let _context: ComponentFramework.Context<IInputs>;

const notificationOption = {
  errosCount: 0,
  importedSucsessCount: 0,
  details: '',
  message: '',
};

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
  },

  async uploadFile(file: File) {
    try {
      const buffer: ArrayBuffer = await FileHelper.readFileAsArrayBufferAsync(file);
      const body: string = FileHelper.arrayBufferToBase64(buffer);
      // @ts-ignore
      const { entityTypeName, entityId } = _context.page;
      const entityMetadata =
      await _context.utils.getEntityMetadata(entityTypeName, entityId);

      const data: any = {
        'subject': '',
        'filename': file.name,
        'documentbody': body,
        'objecttypecode': entityTypeName,
      };

      data[`objectid_${entityTypeName}@odata.bind`] =
      `/${entityMetadata.EntitySetName}(${entityId})`;

      await _context.webAPI.createRecord('annotation', data);
      notificationOption.importedSucsessCount += 1;
    }
    catch (ex: any) {
      console.error(ex.message);
      notificationOption.message = ex.message;
      notificationOption.details += `
      File Name -${file.name}
      Error message ${ex.message}`;
      notificationOption.errosCount += 1;
    }
  },

  refreshTimeline() {
    // @ts-ignore
    parent.Xrm.Page.getControl('Timeline')?.refresh();
  },

  showNotificationPopup() {
    if (notificationOption.errosCount === 0) {
      const message = notificationOption.importedSucsessCount === 1
        ? `${notificationOption.importedSucsessCount} file imported successfully`
        : `${notificationOption.importedSucsessCount} files imported successfully`;

      _context.navigation.openAlertDialog({ text: message });
      notificationOption.importedSucsessCount = 0;
    }
    else {
      notificationOption.errosCount > 1 ? notificationOption.message += ` 
       ${notificationOption.errosCount} errors`
        : notificationOption.message += ` ${notificationOption.errosCount} error`;

      _context.navigation.openErrorDialog(notificationOption);
      notificationOption.errosCount = 0;
    }
  },
};
