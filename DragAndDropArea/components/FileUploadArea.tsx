import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import CrmService from '../services/CrmService';

export interface IFileUploadAreaProps { }

interface IFileUploadAreaState {
  dragCounter: number;
  spinner: boolean;
  importedFilesCount: number
  filesCount: number;
}

export class FileUploadArea extends React.Component<IFileUploadAreaProps, IFileUploadAreaState> {
  constructor(props: IFileUploadAreaProps) {
    super(props);

    this.state = {
      dragCounter: 0,
      spinner: false,
      importedFilesCount: 0,
      filesCount: 0,
    };
  }

  dragOver(event: React.DragEvent) {
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
    this.setState({ spinner: true, filesCount: Array.from(files).length });
    for (const file of Array.from(files)) {
      await CrmService.uploadFile(file);
      this.setState({ importedFilesCount: this.state.importedFilesCount + 1 });
    }

    this.setState({ dragCounter: 0, spinner: false, importedFilesCount: 0 });
    CrmService.refreshTimeline();
  }

  public render(): React.ReactNode {
    const { filesCount, importedFilesCount, spinner } = this.state;

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
          >{spinner
              ? <>
                <div className="textIcon">
                  <p className='spinnerText'>
                    { `importing ${importedFilesCount} of ${filesCount} files` }</p>
                  <Spinner size={SpinnerSize.large} />
                </div>
              </>
              : <div className={'text'}>
                <p> Drag and drop files here to upload</p></div>}
          </div>
        </div>
      </div>
    );
  }
}
