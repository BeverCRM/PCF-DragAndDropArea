import { IInputs } from '../generated/ManifestTypes';
import FileHelper from '../helpers/FileHelper';

let _context: ComponentFramework.Context<IInputs>;

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

    }
    catch (ex: any) {
      console.error(ex.message);
    }
  },

  refreshTimeline() {
    // @ts-ignore
    parent.Xrm.Page.getControl('Timeline')?.refresh();
  },
};
