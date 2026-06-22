import * as React from "react";
import { CommandBar, ICommandBarItemProps } from "@fluentui/react/lib/CommandBar";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { IListItem } from "@microsoft/sp-listview-extensibility";
import { ITransmittalItem } from "../../..";
//import { IListItem } from "@microsoft/sp-listview-extensibility";

export interface ITransmittalCommanProps {
    displayMode: FormDisplayMode;
    onSave: () => void;
    onCancel: () => void;
    item: (IListItem & ITransmittalItem) | undefined;
    itemCount: number | undefined;
}

export interface ITransmittalCommanState {
   // openTransmittalInfo?: IOpenTransmittalInfo;
}

export class TransmittalCommanBar extends React.Component<ITransmittalCommanProps, ITransmittalCommanState> {

    constructor(props: ITransmittalCommanProps) {
        super(props);
        this.state = {};
    }

    // public componentDidMount(): void {
    //     const { transmittalsService } = this.props;
    //     this.setState({
    //         openTransmittalInfo: transmittalsService.getOpenTransmittalInfo()
    //     });
    //     document.addEventListener(OPEN_TRANSMITTAL_CHANGE_EVENT, this._onOpenTransmittalChange.bind(this));
    // }
    public render(): React.ReactElement<ITransmittalCommanProps> {
        const { onSave, onCancel, } = this.props;
        //const { openTransmittalInfo } = this.state;
        const commandBarItems: ICommandBarItemProps[] = [
            {
                key: "title",
                text: "Title: Notice",
                iconProps: { iconName: "TextField" },
                disabled: true
            },
            {
                key: "createdBy",
                text: "createdBy:Kondi",
                iconProps: { iconName: "Contact" },
                disabled: true
            },
            {
                key: "createdOn",
                text: "createdOn: 2026-06-05",
                iconProps: { iconName: "Calendar" },
                disabled: true
            },
            {
                key: "modifiedBy",
                text: "modifiedBy: Eswarudu",
                iconProps: { iconName: "Contact" },
                disabled: true
            },
            {
                key: "modifiedOn",
                text: "modifiedOn: 2026-06-10",
                iconProps: { iconName: "Calendar" },
                disabled: true
            },
            {
                key: "attachment",
                text: "Attachment",
                iconProps: { iconName: "Attach" },
                onClick: () => console.log("Attachment clicked")
            },
            {
                key: "save",
                text: "Save",
                iconProps: { iconName: "Save" },
                onClick: onSave
            },
            {
                key: "cancel",
                text: "Cancel",
                iconProps: { iconName: "Cancel" },
                onClick: onCancel
            }
        ];


        return (
            <>
                <CommandBar items={commandBarItems} />
            </>
        );
    }
}



