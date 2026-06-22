/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  TextField,
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react";
import { FormCustomizerContext, IListItem } from "@microsoft/sp-listview-extensibility";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { ITransmittalItem } from "../../..";
import styles from "./NmcForm.module.scss";
import TransmittalService from "../service/TransmittalService";

const reportTypeOptions: IDropdownOption[] = [
  { key: "RUNNING_BUSINESS", text: "RUNNING BUSINESS" },
  { key: "PROJECT_PHASE", text: "Project phase" },
  { key: "SERVICE_PART", text: "Service part" },
];

const escalationLevelOptions: IDropdownOption[] = [
  { key: "1_ALERT", text: "1.Alert" },
  { key: "2_8D_STANDARD", text: "2.8D Standard" },
  { key: "3_ADVANCED", text: "3.Advanced" },
];

const complaintTypeOptions: IDropdownOption[] = [
  { key: "PART_QUALITY", text: "Part Quality" },
  { key: "LOGISTICAL_MISTAKE", text: "Logistical Mistake" },
];

const buOptions: IDropdownOption[] = [
  { key: "ERGUN", text: "ERGUN" },
  { key: "BMW", text: "BMW" },
  { key: "AUDI", text: "AUDI" },
];

export interface IFormContainerProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  itemID: number;
  item: (IListItem & ITransmittalItem) | undefined;
  onClose?: () => void;
  onSave: () => void;
}

export interface IFormContainerState {
  title: string;
  createdBy: string;
  modifiedBy: string;
  errorControls: { stateName: string }[];
  d0StartDate: string;
  d0ReportType: string;
  bu: string;
  d0EscalationLevel: string;
  d0ComplaintType: string;
}

export default class NcmrFormContainer extends React.Component<
  IFormContainerProps,
  IFormContainerState
> {
  constructor(props: IFormContainerProps) {
    super(props);
    const itemAny = props.item as any;
    this.state = {
      title: props.item?.Title || "",
      errorControls: [],
      createdBy: "",
      modifiedBy: "",
      d0StartDate: itemAny?.D0StartDateNCMR
        ? new Date(itemAny.D0StartDateNCMR).toISOString().split("T")[0]
        : new Date().toISOString().split("T")[0],
      d0ReportType: itemAny?.D0ReportType || "RUNNING_BUSINESS",
      bu: itemAny?.BU || "",
      d0EscalationLevel: itemAny?.D0EscalationLevel || "",
      d0ComplaintType: itemAny?.D0ComplaintType || "PART_QUALITY",
    };
  }

  public async componentDidMount(): Promise<void> {
    if (this.props.item?.AuthorId) {
      const name = await TransmittalService.getUserNameById(this.props.item.AuthorId);
      this.setState({ createdBy: name });
    } else {
      this.setState({ createdBy: this.props.context.pageContext.user.displayName });
    }
    if (this.props.item?.EditorId) {
      const name = await TransmittalService.getUserNameById(this.props.item.EditorId);
      this.setState({ modifiedBy: name });
    }
  }

  private async submitForm(): Promise<void> {
    try {
      const payload = {
        Title: this.state.title,
        D0StartDateNCMR: this.state.d0StartDate,
        D0ReportType: this.state.d0ReportType,
        BU: this.state.bu,
        D0EscalationLevel: this.state.d0EscalationLevel,
        D0ComplaintType: this.state.d0ComplaintType,
      };
      if (this.props.itemID) {
        await TransmittalService.updateListItem(
          this.props.context.list.title,
          payload,
          this.props.itemID
        );
      } else {
        await TransmittalService.createListItem(this.props.context.list.title, payload);
      }
      this.props.onSave();
    } catch (err) {
      console.log(err);
    }
  }

  private formatDate(dateStr: string | undefined): string {
    if (!dateStr) return "";
    try {
      return new Date(dateStr).toISOString().split("T")[0];
    } catch {
      return "";
    }
  }

  private renderNotes(isNew: boolean): React.ReactNode {
    const line1 = isNew
      ? "1) Fields identified with star (*) are compulsorily required to be filled before submitting a NCMR."
      : "1) Fields identified with star (*) are compulsorily required to be filled before submitting or completing any NCMR/Task.";
    return (
      <div className={styles.notesSection}>
        <p className={styles.notesText}>
          <strong>Notes:</strong>{line1}
          <br />
          <span className={styles.notesIndent}>
            2) Fields with red highlight are KPI calculation related and any update is restricted.
          </span>
        </p>
      </div>
    );
  }

  render(): React.ReactNode {
    const { displayMode, item } = this.props;
    const isNew = displayMode === FormDisplayMode.New;
    const isReadOnly = displayMode === FormDisplayMode.Display;
    const itemAny = item as any;

    const status = isNew ? "Notice" : (itemAny?.Status || "Active");
    const step = isNew ? "D0" : (itemAny?.Step || "D0");
    const createdOn = item?.Created
      ? this.formatDate(item.Created)
      : new Date().toISOString().split("T")[0];
    const modifiedOn = item?.Modified ? this.formatDate(item.Modified) : "";
    const ownedOn = item?.Created ? this.formatDate(item.Created) : createdOn;

    const metaCols = [
      { label: "Status", value: status },
      { label: "Step", value: step },
      { label: "Created On", value: createdOn },
      { label: "Created By", value: this.state.createdBy },
      { label: "Modified On", value: modifiedOn },
      { label: "Modified By", value: this.state.modifiedBy },
      { label: isNew ? "Owner On" : "Owned On", value: ownedOn },
      { label: isNew ? "Owner By" : "Owned By", value: this.state.createdBy },
    ];

    const escalationLabel = itemAny?.D0EscalationLevelDisplay || "";

    return (
      <div className={styles.ncmrContainer}>
        {/* ── Metadata bar ── */}
        <div className={styles.metadataBar}>

          {/* Row 1: Title + toolbar */}
          <div className={styles.metaRow1}>
            <div className={styles.titleSection}>
              <span className={styles.metaLabelBold}>Title</span>
              {!isNew && item?.Title && (
                <span className={styles.itemTitleOrange}>
                  {item.Title}{escalationLabel ? ` (${escalationLabel})` : ""}
                </span>
              )}
            </div>

            <div className={styles.toolbarButtons}>
              {isNew ? (
                <>
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Attach" }}
                    text="Attachment"
                  />
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Save" }}
                    text="Save"
                    onClick={() => this.submitForm()}
                  />
                  <PrimaryButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Trophy2" }}
                    text="Submit"
                    onClick={() => this.submitForm()}
                  />
                </>
              ) : (
                <>
                  <DefaultButton className={styles.toolbarBtn} iconProps={{ iconName: "Export" }} text="Export" />
                  <DefaultButton className={styles.toolbarBtn} iconProps={{ iconName: "Ringer" }} text="Subscribe" />
                  <span className={styles.subscribeOffBadge}>Off</span>
                  <DefaultButton className={styles.toolbarBtn} text="Show N.C. Disposition" />
                  <DefaultButton className={styles.toolbarBtn} iconProps={{ iconName: "Comment" }} text="Comments" />
                  <DefaultButton className={styles.toolbarBtn} iconProps={{ iconName: "Attach" }} text="Attachment" />
                  <DefaultButton className={styles.toolbarBtn} iconProps={{ iconName: "History" }} text="Flow History" />
                </>
              )}
            </div>
          </div>

          {/* Meta columns */}
          <div className={styles.metaColumnsScroll}>
            {metaCols.map((col) => (
              <div key={col.label} className={styles.metaColumn}>
                <span className={styles.metaColLabel}>{col.label}</span>
                <span className={styles.metaColValue}>{col.value}</span>
              </div>
            ))}
          </div>

          {/* Edit/view action buttons */}
          {!isNew && (
            <div className={styles.editActionsRow}>
              <DefaultButton
                className={styles.toolbarBtn}
                iconProps={{ iconName: "CirclePause" }}
                text="Suspend"
              />
              <DefaultButton
                className={styles.toolbarBtn}
                iconProps={{ iconName: "Delete" }}
                text="Delete"
              />
              <DefaultButton
                className={`${styles.toolbarBtn} ${styles.editActionBtn}`}
                iconProps={{ iconName: "Edit" }}
                text="Edit"
              />
            </div>
          )}
        </div>

        {/* ── Tabs ── */}
        <div className={styles.tabsContainer}>
          <Pivot>
            {/* D0 */}
            <PivotItem headerText="D0">
              <div className={styles.tabContent}>
                {this.renderNotes(isNew)}
                <div className={styles.sectionDivider}>--Non-Conforming Material Report--</div>
                <div className={styles.d0FormGrid}>
                  {/* Left column */}
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <label className={styles.fieldLabel}>D0-Start-Date NCMR</label>
                      <TextField
                        type="date"
                        value={this.state.d0StartDate}
                        onChange={(_, val) => this.setState({ d0StartDate: val || "" })}
                        disabled={isReadOnly}
                      />
                      <span className={styles.fieldCaption}>
                        (The date you received the request for a supplier 8D)
                      </span>
                    </div>

                    <div className={styles.fieldBlock}>
                      <div className={styles.fieldLabelRow}>
                        <label className={`${styles.fieldLabel} ${styles.requiredLabel}`}>
                          D0-Report type
                        </label>
                        <span className={styles.asterisk}>*</span>
                      </div>
                      <Dropdown
                        options={reportTypeOptions}
                        selectedKey={this.state.d0ReportType || undefined}
                        onChange={(_, opt) =>
                          this.setState({ d0ReportType: (opt?.key as string) || "" })
                        }
                        disabled={isReadOnly}
                      />
                      <span className={styles.fieldCaption}>
                        (Default setting-change into &quot;Project phase&quot; or &quot;Service part&quot; if applicable)
                      </span>
                    </div>
                  </div>

                  {/* Right column */}
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <div className={styles.fieldLabelRow}>
                        <label className={`${styles.fieldLabel} ${styles.requiredLabel}`}>BU</label>
                        <span className={styles.asterisk}>*</span>
                      </div>
                      <div className={styles.dropdownWithAsterisk}>
                        <Dropdown
                          options={buOptions}
                          selectedKey={this.state.bu || undefined}
                          onChange={(_, opt) =>
                            this.setState({ bu: (opt?.key as string) || "" })
                          }
                          disabled={isReadOnly}
                          className={styles.dropdownFlex}
                        />
                        {!this.state.bu && <span className={styles.asteriskRight}>*</span>}
                      </div>
                      <span className={styles.fieldCaption}>(Where NCMR originated)</span>
                    </div>

                    <div className={styles.fieldBlock}>
                      <div className={styles.fieldLabelRow}>
                        <label className={`${styles.fieldLabel} ${styles.requiredLabel}`}>
                          D0-Escalation-level
                        </label>
                        <span className={styles.asterisk}>*</span>
                      </div>
                      <div className={styles.dropdownWithAsterisk}>
                        <Dropdown
                          options={escalationLevelOptions}
                          selectedKey={this.state.d0EscalationLevel || undefined}
                          onChange={(_, opt) =>
                            this.setState({ d0EscalationLevel: (opt?.key as string) || "" })
                          }
                          disabled={isReadOnly}
                          className={styles.dropdownFlex}
                        />
                        {!this.state.d0EscalationLevel && (
                          <span className={styles.asteriskRight}>*</span>
                        )}
                      </div>
                      <span className={styles.fieldCaption}>
                        (Only need to confirm of D4_Claim Warranted, D4_Quantity for PPM and D9
                        handling if &quot;1.Alert&quot; selected.)
                      </span>
                    </div>

                    <div className={styles.fieldBlock}>
                      <div className={styles.fieldLabelRow}>
                        <label className={`${styles.fieldLabel} ${styles.requiredLabel}`}>
                          D0-Complaint type
                        </label>
                        <span className={styles.asterisk}>*</span>
                      </div>
                      <Dropdown
                        options={complaintTypeOptions}
                        selectedKey={this.state.d0ComplaintType || undefined}
                        onChange={(_, opt) =>
                          this.setState({ d0ComplaintType: (opt?.key as string) || "" })
                        }
                        disabled={isReadOnly}
                      />
                      <span className={styles.fieldCaption}>
                        (Does the problem/complaint involve the quality of the part, or a logistical
                        mistake regarding the part?)
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            </PivotItem>

            {/* D1 */}
            <PivotItem headerText="D1">
              <div className={styles.tabContent}>
                {this.renderNotes(isNew)}
                <div className={styles.d0FormGrid}>
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <TextField
                        required
                        label="Title"
                        value={this.state.title}
                        onChange={(_, val) => this.setState({ title: val || "" })}
                        disabled={isReadOnly}
                      />
                    </div>
                  </div>
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <TextField required label="Bu1" placeholder="BU" disabled={isReadOnly} />
                    </div>
                  </div>
                </div>
              </div>
            </PivotItem>

            {/* D2–D8 */}
            {(["D2", "D3", "D4", "D5", "D6", "D7", "D8"] as const).map((tab) => (
              <PivotItem key={tab} headerText={tab}>
                <div className={styles.tabContent}>
                  {this.renderNotes(isNew)}
                  <div className={styles.d0FormGrid}>
                    <div className={styles.formColumn}>
                      <div className={styles.fieldBlock}>
                        <TextField required label="Title" placeholder="Title" disabled={isReadOnly} />
                      </div>
                    </div>
                  </div>
                </div>
              </PivotItem>
            ))}

            {/* D9 */}
            <PivotItem headerText="D9">
              <div className={styles.tabContent}>
                {this.renderNotes(isNew)}
                <div className={styles.d0FormGrid}>
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <TextField
                        label="Customer Response Reference"
                        placeholder="Customer Response Reference"
                        disabled={isReadOnly}
                      />
                    </div>
                  </div>
                  <div className={styles.formColumn}>
                    <div className={styles.fieldBlock}>
                      <TextField
                        type={isReadOnly ? undefined : "date"}
                        label="Customer Response Date"
                        disabled={isReadOnly}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </PivotItem>
          </Pivot>
        </div>
      </div>
    );
  }
}
