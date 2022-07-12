import * as React from 'react';
import { IInputs } from '../generated/ManifestTypes';

import FileHelper from '../helpers/FileHelper';

export interface IFileUploadAreaProps {
  context: ComponentFramework.Context<IInputs>;
}

interface IFileUploadAreaState {
  dragCounter: number;
}

export class FileUploadArea extends React.Component<IFileUploadAreaProps, IFileUploadAreaState> {
  constructor(props: IFileUploadAreaProps) {
    super(props);
    this.state = {
      dragCounter: 0,
    };
  }

  dragOver = (event: React.DragEvent) => {
    event.preventDefault();
  }

  dragEnter() {
    this.setState({ dragCounter: this.state.dragCounter + 1 });
  }

  dragLeave() {
    this.setState({ dragCounter: this.state.dragCounter - 1 });
  }

  async drop(event: React.DragEvent) {
    event.preventDefault();
    const files = event.dataTransfer?.files;
    await this.handleFiles(files);

    this.setState({ dragCounter: 0 });
    const _parent: any = parent;
    _parent.Xrm.Page.getControl('Timeline').refresh();
  }

  async handleFiles(files: FileList) {
    for (const file of Array.from(files)) {
      await this.uploadFile(file);
    }
  }

  async uploadFile(file: File) {
    try {
      const buffer: ArrayBuffer = await FileHelper.readFileAsArrayBufferAsync(file);
      const body: string = FileHelper.arrayBufferToBase64(buffer);

      const _context = this.props.context as any;
      const { entityTypeName, entityId } = _context.page;

      const data: any = {
        'subject': '',
        'filename': file.name,
        'documentbody': body,
        'objecttypecode': entityTypeName,
      };

      data[`objectid_${entityTypeName}@odata.bind`] = `/${entityTypeName}s(${entityId})`;

      await this.props.context.webAPI.createRecord('annotation', data);
    }
    catch (ex: any) {
      console.error(ex.message);
    }
  }

  public render(): React.ReactNode {
    return (
      <div className="draganddroparea">
        <div className="container">
          <div
            onDragOver={this.dragOver.bind(this)}
            onDragEnter={this.dragEnter.bind(this)}
            onDragLeave={this.dragLeave.bind(this)}
            onDrop={this.drop.bind(this)}
            className={`dropArea 
              ${this.state.dragCounter > 0 ? 'dropAreaHover' : ''}`}
          >
            <div className="textIcon">
              <div className="text"><p>Drag and drop files here to upload</p></div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
