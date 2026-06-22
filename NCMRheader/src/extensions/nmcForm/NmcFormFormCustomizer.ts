import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import NmcForm, { INmcFormProps } from './components/NmcForm';
import TransmittalService from './service/TransmittalService';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INmcFormFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'NmcFormFormCustomizer';

export default class NmcFormFormCustomizer
  extends BaseFormCustomizer<INmcFormFormCustomizerProperties> {

  public async onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated NmcFormFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));

    try {
      //initializeFileTypeIcons();
      await TransmittalService.init(this.context)
      return Promise.resolve();
    } catch (error) {
      Log.error(LOG_SOURCE, new Error(`Error initializing TransmittalformFormCustomizer: ${error}`));
      console.error("Error initializing TransmittalformFormCustomizer:", error);
      return Promise.reject(error);
    }
    return Promise.resolve();
  }

  public render(): void {
    // Use this method to perform your custom rendering.

    const nmcForm: React.ReactElement<{}> =
      React.createElement(NmcForm, {
        context: this.context,
        itemID: this.context.itemId,
        item: this.context.item,
        //transmittalsService: this._transmittalsService,
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose
       } as INmcFormProps);

    ReactDOM.render(nmcForm, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
