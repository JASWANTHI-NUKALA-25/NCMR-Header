import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext, IListItem } from '@microsoft/sp-listview-extensibility';
import styles from './NmcForm.module.scss';
import NcmrFormContainer from './NcmrFormContainer';
import { ITransmittalItem } from '../../..';

export interface INmcFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
  itemID: number;
  item: (IListItem & ITransmittalItem) | undefined;
}

const LOG_SOURCE: string = 'NmcForm';

export default class NmcForm extends React.Component<INmcFormProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: NmcForm mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: NmcForm unmounted');
  }

  public render(): React.ReactElement<{}> {
    const { displayMode, item } = this.props;
    const headerTitle =
      displayMode === FormDisplayMode.New
        ? 'NCMR - New Item'
        : `NCMR - ${item?.Title || ''}`;

    return (
      <div className={styles.nmcForm}>
        <div className={styles.modalHeader}>
          <span className={styles.modalTitle}>{headerTitle}</span>
        </div>
        <NcmrFormContainer {...this.props} />
      </div>
    );
  }
}
