import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import CrmService from '../services/CrmService';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IFileUploadAreaProps { }

interface IFileUploadAreaState {
  dragCounter: number;
  isLoading: boolean;
  importedFilesCount: number
  filesCount: number;
  isDisabled: boolean;
  isRenderedOneTime: boolean;
}

export class FileUploadArea extends React.Component<IFileUploadAreaProps, IFileUploadAreaState> {
  constructor(props: IFileUploadAreaProps) {
    super(props);

    this.state = {
      dragCounter: 0,
      isLoading: false,
      importedFilesCount: 0,
      filesCount: 0,
      isDisabled: true,
      isRenderedOneTime: false,
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
    this.setState({ isLoading: true, filesCount: Array.from(files).length });
    for (const file of Array.from(files)) {
      await CrmService.uploadFile(file, Array.from(files).length);
      this.setState({ importedFilesCount: this.state.importedFilesCount + 1 });
    }
    this.setState({ dragCounter: 0, isLoading: false, importedFilesCount: 0 });
    CrmService.refreshTimeline();
  }

  async checkNotes() {
    const hasNotes = await CrmService.hasNotes();
    this.setState({ isDisabled: hasNotes, isRenderedOneTime: true });
  }

  public render(): React.ReactNode {
    const { filesCount, importedFilesCount, isLoading, isDisabled } = this.state;
    if (!this.state.isRenderedOneTime) this.checkNotes();

    return (
      <div className="draganddroparea">
        { isDisabled
          ? <div className="container">
            <div
              onDragOver={this.dragOver.bind(this)}
              onDragEnter={this.dragEnter.bind(this)}
              onDragLeave={this.dragLeave.bind(this)}
              onDrop={this.drop.bind(this)}
              className={`dropArea 
              ${this.state.dragCounter > 0 ? 'dropAreaHover' : ''}`}
            >{isLoading
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
          </div> : <div className='errorContainer'>
            <Icon className='errorIcone' iconName="error"></Icon>
            <p className='errorMessage'>Notes (including attachments) are disabled</p>
          </div>
        }
      </div>
    );
  }
}
