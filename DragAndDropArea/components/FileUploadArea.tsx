import * as React from 'react';

import CrmService from '../services/CrmService';

export interface IFileUploadAreaProps { }

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
    await CrmService.uploadFiles(files);

    this.setState({ dragCounter: 0 });
    CrmService.refreshTimeline();
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
