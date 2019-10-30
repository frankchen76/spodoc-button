import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  autobind,
  ColorPicker,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent,
  IColor,
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react';

interface IProgressDialogProps {
  title: string;
  message: string;
  close: () => void;
  submit: (color: IColor) => void;
  defaultColor?: IColor;
}

class ProgressDialogContent extends React.Component<IProgressDialogProps, {}> {
  private _pickedColor: IColor;

  constructor(props) {
      super(props);
      // Default Color
      this._pickedColor = props.defaultColor || { hex: 'FFFFFF', str: '', r: null, g: null, b: null, h: null, s: null, v: null };
  }

  public render(): JSX.Element {
    return <DialogContent
      title={this.props.title}
      onDismiss={this.props.close}
      showCloseButton={false}
    >
      <Spinner label={this.props.message} />
      {/* <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._pickedColor); }} />
      </DialogFooter> */}
    </DialogContent>;
  }

  //@autobind
  private _onColorChange = (ev: React.SyntheticEvent<HTMLElement, Event>, color: IColor) => {
      this._pickedColor = color;
  }
}

export default class ProgressDialog extends BaseDialog {
  public message: string;
  public title: string;
  public colorCode: IColor;

  public render(): void {
      ReactDOM.render(<ProgressDialogContent
      title={this.title}
      close={ this.close }
      message={ this.message }
      defaultColor={ this.colorCode }
      submit={ this._submit }
      />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  // protected onAfterClose(): void {
  //     super.onAfterClose();

  //     // Clean up the element for the next dialog
  //     ReactDOM.unmountComponentAtNode(this.domElement);
  // }

  //@autobind
  private _submit = (color: IColor) => {
      this.colorCode = color;
      this.close();
  }
}
