import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  autobind,
  ColorPicker,
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";

interface IColorPickerDialogContentProps {
  width: string;
  height: string;
  url: string;
  close: () => void;
  submit: (color: string) => void;
  defaultColor?: string;
}

class ColorPickerDialogContent extends React.Component<IColorPickerDialogContentProps, {hidden: boolean, url: string}> {
  private _pickedColor: string;

  constructor(props) {
    super(props);
    this.state = {
      hidden: false,
      url: this.props.url,
    }
    // Default Color
    this._pickedColor = props.defaultColor || '#FFFFFF';
  }
  private _onDialogDismiss() {
    this.props.close();
    this.setState ({hidden: !this.state.hidden, url: this.props.url});
    window.location.reload();
  }
  public render(): JSX.Element {

    return (
      <IFrameDialog
        url={this.state.url}
        hidden={this.state.hidden}
        onDismiss={this._onDialogDismiss.bind(this)}
        //width={"1350px"}
        //height={"648px"}
        width={this.props.width}
        height={this.props.height}
      />
    );
    /*
    return <DialogContent
      title='Color Picker'
      subText={this.props.message}
      onDismiss={this.props.close}
      showCloseButton={true}
    >
      <ColorPicker color={this._pickedColor} onColorChanged={this._onColorChange} />
      <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._pickedColor); }} />
      </DialogFooter>
    </DialogContent>;*/
  }

  @autobind
  private _onColorChange(color: string): void {
    this._pickedColor = color;
  }
}

export default class ColorPickerDialog extends BaseDialog {
  public url: string;
  public colorCode: string;
  public hidden: boolean;
  public width: string;
  public height: string;

  public render(): void {
    
    ReactDOM.render(<ColorPickerDialogContent
      width={ this.width}
      height={this.height}
      close={ this.close }
      url={ this.url }
      defaultColor={ this.colorCode }
      submit={ this._submit }
    />, this.domElement);
    /*
    ReactDOM.render(
      <IFrameDialog
        url="https://blueboxsolutionsdev.sharepoint.com/teams/binh_spfx/Lists/list1/DispForm.aspx?ID=1&e=LxA9Zt"
        hidden={this.hidden}
        onDismiss={this._onDialogDismiss.bind(this)}
        width={"500px"}
        height={"500px"}
      />
      ,this.domElement);*/
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  @autobind
  private _submit(color: string): void {
    this.colorCode = color;
    this.close();
  }
}