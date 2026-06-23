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
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import {
  FormCustomizerContext,
  IListItem,
} from "@microsoft/sp-listview-extensibility";
import { FormDisplayMode } from "@microsoft/sp-core-library";
import { ITransmittalItem } from "../../..";
import styles from "./NmcForm.module.scss";
import TransmittalService from "../service/TransmittalService";

// ── Dropdown option lists ──────────────────────────────────────────────────
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
const failureCodeOptions: IDropdownOption[] = [
  { key: "FC01", text: "FC01 - Dimensional" },
  { key: "FC02", text: "FC02 - Surface defect" },
  { key: "FC03", text: "FC03 - Material" },
  { key: "FC04", text: "FC04 - Wrong part" },
];
const partGroupOptions: IDropdownOption[] = [
  { key: "PG01", text: "Mechanical" },
  { key: "PG02", text: "Electrical" },
  { key: "PG03", text: "Plastic" },
  { key: "PG04", text: "Rubber" },
];
const productGroupOptions: IDropdownOption[] = [
  { key: "PRG01", text: "Engine" },
  { key: "PRG02", text: "Transmission" },
  { key: "PRG03", text: "Chassis" },
  { key: "PRG04", text: "Body" },
];
const whereOptions: IDropdownOption[] = [
  { key: "W01", text: "Incoming" },
  { key: "W02", text: "Assembly" },
  { key: "W03", text: "Final Inspection" },
  { key: "W04", text: "Customer" },
];

// ── Interfaces ─────────────────────────────────────────────────────────────
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
  // D0
  d0StartDate: string;
  d0ReportType: string;
  bu: string;
  d0EscalationLevel: string;
  d0ComplaintType: string;
  // D1
  d1Name8DChampion: string;
  d0SupplierContactTier1: string;
  d0SupplierContactTier2: string;
  d0SupplierContactTier3: string;
  d1ClosingDate: string;
  // D2
  d2PartNumber: string;
  d2FailureCode: string;
  d2PartGroup: string;
  d2ProductGroup: string;
  d2Where: string;
  d2What: string;
  d2Why: string;
}

// ── Component ──────────────────────────────────────────────────────────────
export default class NcmrFormContainer extends React.Component<
  IFormContainerProps,
  IFormContainerState
> {
  constructor(props: IFormContainerProps) {
    super(props);
    const a = props.item as any;
    this.state = {
      title: props.item?.Title || "",
      errorControls: [],
      createdBy: "",
      modifiedBy: "",
      // D0
      d0StartDate: a?.D0StartDateNCMR
        ? new Date(a.D0StartDateNCMR).toISOString().split("T")[0]
        : new Date().toISOString().split("T")[0],
      d0ReportType: a?.D0ReportType || "RUNNING_BUSINESS",
      bu: a?.BU || "",
      d0EscalationLevel: a?.D0EscalationLevel || "",
      d0ComplaintType: a?.D0ComplaintType || "PART_QUALITY",
      // D1
      d1Name8DChampion: a?.D1Name8DChampion || "",
      d0SupplierContactTier1: a?.D0SupplierContactTier1 || "",
      d0SupplierContactTier2: a?.D0SupplierContactTier2 || "",
      d0SupplierContactTier3: a?.D0SupplierContactTier3 || "",
      d1ClosingDate: a?.D1ClosingDate
        ? new Date(a.D1ClosingDate).toISOString().split("T")[0]
        : "",
      // D2
      d2PartNumber: a?.D2PartNumber || "",
      d2FailureCode: a?.D2FailureCode || "",
      d2PartGroup: a?.D2PartGroup || "",
      d2ProductGroup: a?.D2ProductGroup || "",
      d2Where: a?.D2Where || "",
      d2What: a?.D2What || "",
      d2Why: a?.D2Why || "",
    };
  }

  public async componentDidMount(): Promise<void> {
    if (this.props.item?.AuthorId) {
      const name = await TransmittalService.getUserNameById(
        this.props.item.AuthorId,
      );
      this.setState({ createdBy: name });
    } else {
      this.setState({
        createdBy: this.props.context.pageContext.user.displayName,
      });
    }
    if (this.props.item?.EditorId) {
      const name = await TransmittalService.getUserNameById(
        this.props.item.EditorId,
      );
      this.setState({ modifiedBy: name });
    }
  }

  /**
   * isDraft = true  → Save   → Status "Notice"  (saved as draft)
   * isDraft = false → Submit → Status "Active"   (formally submitted)
   * Edit-mode Save  → Status is omitted so SharePoint keeps its current value
   */
  private async submitForm(isDraft: boolean = false): Promise<void> {
    try {
      const isNewForm = this.props.displayMode === FormDisplayMode.New;

      // Only set Status on new items; edit-mode save preserves existing status
      const statusField: { Status?: string } = isNewForm
        ? { Status: isDraft ? "Notice" : "Active" }
        : {};

      const payload = {
        Title: this.state.title,
        ...statusField,
        // D0
        D0StartDateNCMR: this.state.d0StartDate,
        D0ReportType: this.state.d0ReportType,
        BU: this.state.bu,
        D0EscalationLevel: this.state.d0EscalationLevel,
        D0ComplaintType: this.state.d0ComplaintType,
        // D1
        D1Name8DChampion: this.state.d1Name8DChampion,
        D0SupplierContactTier1: this.state.d0SupplierContactTier1,
        D0SupplierContactTier2: this.state.d0SupplierContactTier2,
        D0SupplierContactTier3: this.state.d0SupplierContactTier3,
        D1ClosingDate: this.state.d1ClosingDate,
        // D2
        D2PartNumber: this.state.d2PartNumber,
        D2FailureCode: this.state.d2FailureCode,
        D2PartGroup: this.state.d2PartGroup,
        D2ProductGroup: this.state.d2ProductGroup,
        D2Where: this.state.d2Where,
        D2What: this.state.d2What,
        D2Why: this.state.d2Why,
      };

      if (this.props.itemID) {
        await TransmittalService.updateListItem(
          this.props.context.list.title,
          payload,
          this.props.itemID,
        );
      } else {
        await TransmittalService.createListItem(
          this.props.context.list.title,
          payload,
        );
      }
      this.props.onSave();
    } catch (err) {
      console.log(err);
    }
  }

  private fmt(d: string | undefined): string {
    if (!d) return "";
    try {
      return new Date(d).toISOString().split("T")[0];
    } catch {
      return "";
    }
  }

  // ── Shared notes banner ────────────────────────────────────────────────
  private renderNotes(isNew: boolean): React.ReactNode {
    const line1 = isNew
      ? "1) Fields identified with star (*) are compulsorily required to be filled before submitting a NCMR."
      : "1) Fields identified with star (*) are compulsorily required to be filled before submitting or completing any NCMR/Task.";
    return (
      <div className={styles.notesSection}>
        <p className={styles.notesText}>
          <strong>Notes:</strong>
          {line1}
          <br />
          <span className={styles.notesIndent}>
            2) Fields with red highlight are KPI calculation related and any
            update is restricted.
          </span>
        </p>
      </div>
    );
  }

  // ── Required label helper ──────────────────────────────────────────────
  private reqLabel(text: string): React.ReactNode {
    return (
      <div className={styles.fieldLabelRow}>
        <label className={`${styles.fieldLabel} ${styles.requiredLabel}`}>
          {text}
        </label>
        <span className={styles.asterisk}>*</span>
      </div>
    );
  }

  // ── D0 tab ─────────────────────────────────────────────────────────────
  private renderD0(isNew: boolean, isReadOnly: boolean): React.ReactNode {
    return (
      <div className={styles.tabContent}>
        {this.renderNotes(isNew)}
        <div className={styles.sectionDivider}>
          --Non-Conforming Material Report--
        </div>
        <div className={styles.d0FormGrid}>
          {/* Left */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>D0-Start-Date NCMR</label>
              <TextField
                type="date"
                value={this.state.d0StartDate}
                onChange={(_, v) => this.setState({ d0StartDate: v || "" })}
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (The date you received the request for a supplier 8D)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D0-Report type")}
              <Dropdown
                options={reportTypeOptions}
                selectedKey={this.state.d0ReportType || undefined}
                onChange={(_, o) =>
                  this.setState({ d0ReportType: (o?.key as string) || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Default setting-change into &quot;Project phase&quot; or
                &quot;Service part&quot; if applicable)
              </span>
            </div>
          </div>
          {/* Right */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              {this.reqLabel("BU")}
              <div className={styles.dropdownWithAsterisk}>
                <Dropdown
                  options={buOptions}
                  selectedKey={this.state.bu || undefined}
                  onChange={(_, o) =>
                    this.setState({ bu: (o?.key as string) || "" })
                  }
                  disabled={isReadOnly}
                  className={styles.dropdownFlex}
                />
                {!this.state.bu && (
                  <span className={styles.asteriskRight}>*</span>
                )}
              </div>
              <span className={styles.fieldCaption}>
                (Where NCMR originated)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D0-Escalation-level")}
              <div className={styles.dropdownWithAsterisk}>
                <Dropdown
                  options={escalationLevelOptions}
                  selectedKey={this.state.d0EscalationLevel || undefined}
                  onChange={(_, o) =>
                    this.setState({
                      d0EscalationLevel: (o?.key as string) || "",
                    })
                  }
                  disabled={isReadOnly}
                  className={styles.dropdownFlex}
                />
                {!this.state.d0EscalationLevel && (
                  <span className={styles.asteriskRight}>*</span>
                )}
              </div>
              <span className={styles.fieldCaption}>
                (Only need to confirm of D4_Claim Warranted, D4_Quantity for PPM
                and D9 handling if &quot;1.Alert&quot; selected.)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D0-Complaint type")}
              <Dropdown
                options={complaintTypeOptions}
                selectedKey={this.state.d0ComplaintType || undefined}
                onChange={(_, o) =>
                  this.setState({ d0ComplaintType: (o?.key as string) || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Does the problem/complaint involve the quality of the part, or
                a logistical mistake regarding the part?)
              </span>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── D1 tab ─────────────────────────────────────────────────────────────
  private renderD1(isNew: boolean, isReadOnly: boolean): React.ReactNode {
    return (
      <div className={styles.tabContent}>
        {this.renderNotes(isNew)}
        <div className={styles.sectionDivider}>--Team Approach--</div>
        <p className={styles.sectionDescription}>
          The Problem Solving Investigator, uses a multi-person approach to
          collect information from those directly associated with and/or
          impacted by the problem. The Team Member names and contact info should
          be listed here.
        </p>
        <div className={styles.d0FormGrid}>
          {/* Left */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>D1-Name_8D_Champion</label>
              <TextField
                value={this.state.d1Name8DChampion}
                onChange={(_, v) =>
                  this.setState({ d1Name8DChampion: v || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Person at BU responsible for this 8D report)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>
                D0-Supplier contact (Tier.1)
              </label>
              <TextField
                value={this.state.d0SupplierContactTier1}
                onChange={(_, v) =>
                  this.setState({ d0SupplierContactTier1: v || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (The contact name at supplier responsible for this 8D)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>
                D0-Supplier contact (Tier.2)
              </label>
              <TextField
                value={this.state.d0SupplierContactTier2}
                onChange={(_, v) =>
                  this.setState({ d0SupplierContactTier2: v || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (The contact name at tier.2 supplier responsible for this 8D, if
                tier.1 supplier is internal supplier)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>
                D0-Supplier contact (Tier.3)
              </label>
              <TextField
                value={this.state.d0SupplierContactTier3}
                onChange={(_, v) =>
                  this.setState({ d0SupplierContactTier3: v || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (The contact name at tier.3 supplier responsible for this 8D, if
                tier.2 supplier is internal supplier, e.g. :ACI)
              </span>
            </div>
          </div>
          {/* Right */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D1-Closing-Date")}
              <TextField
                type="date"
                value={this.state.d1ClosingDate}
                onChange={(_, v) => this.setState({ d1ClosingDate: v || "" })}
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Please complete the team members list of supplier it possible
                before marking the closing data)
              </span>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── D2 tab ─────────────────────────────────────────────────────────────
  private renderD2(isNew: boolean, isReadOnly: boolean): React.ReactNode {
    return (
      <div className={styles.tabContent}>
        {this.renderNotes(isNew)}
        <div className={styles.sectionDivider}>--Problem description--</div>
        <p className={styles.sectionDescription}>
          Collect the data from those involved with the problem and produce a
          Well Formed Description.
        </p>
        <div className={styles.d0FormGrid}>
          {/* Left */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D2-Part number")}
              <TextField
                value={this.state.d2PartNumber}
                onChange={(_, v) => this.setState({ d2PartNumber: v || "" })}
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Fill in the part number relevant for this specific problem)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D2-Failure code")}
              <Dropdown
                options={failureCodeOptions}
                selectedKey={this.state.d2FailureCode || undefined}
                onChange={(_, o) =>
                  this.setState({ d2FailureCode: (o?.key as string) || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Please select the failure code)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D2-What (Short Description)")}
              <span className={styles.fieldCaption}>
                (What is the problem in comparison to a conforming part?)
              </span>
              <div className={styles.richTextWrapper}>
                <RichText
                  value={this.state.d2What}
                  onChange={(v) => {
                    this.setState({ d2What: v || "" });
                    return v;
                  }}
                  isEditMode={!isReadOnly}
                />
              </div>
            </div>
          </div>
          {/* Right */}
          <div className={styles.formColumn}>
            <div className={styles.fieldBlock}>
              {this.reqLabel("D2-Part group")}
              <Dropdown
                options={partGroupOptions}
                selectedKey={this.state.d2PartGroup || undefined}
                onChange={(_, o) =>
                  this.setState({ d2PartGroup: (o?.key as string) || "" })
                }
                disabled={isReadOnly}
              />
              <span className={styles.fieldCaption}>
                (Contains group description of specific part concerned)
              </span>
            </div>
            <div className={styles.fieldBlock}>
              <div className={styles.twoColDropdown}>
                <div className={styles.twoColDropdownLeft}>
                  {this.reqLabel("D2-Product group")}
                  <Dropdown
                    options={productGroupOptions}
                    selectedKey={this.state.d2ProductGroup || undefined}
                    onChange={(_, o) =>
                      this.setState({
                        d2ProductGroup: (o?.key as string) || "",
                      })
                    }
                    disabled={isReadOnly}
                  />
                  <span className={styles.fieldCaption}>
                    (Line of doing business)
                  </span>
                </div>
                <div className={styles.twoColDropdownRight}>
                  {this.reqLabel("D2-Where")}
                  <Dropdown
                    options={whereOptions}
                    selectedKey={this.state.d2Where || undefined}
                    onChange={(_, o) =>
                      this.setState({ d2Where: (o?.key as string) || "" })
                    }
                    disabled={isReadOnly}
                  />
                  <span className={styles.fieldCaption}>(Area of reject)</span>
                </div>
              </div>
            </div>
            <div className={styles.fieldBlock}>
              <label className={styles.fieldLabel}>
                D2-Why (Short description)
              </label>
              <span className={styles.fieldCaption}>
                (Why is it a problem? Explain the effect.)
              </span>
              <div className={styles.richTextWrapper}>
                <RichText
                  value={this.state.d2Why}
                  onChange={(v) => {
                    this.setState({ d2Why: v || "" });
                    return v;
                  }}
                  isEditMode={!isReadOnly}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Render ─────────────────────────────────────────────────────────────
  render(): React.ReactNode {
    const { displayMode, item } = this.props;
    const isNew = displayMode === FormDisplayMode.New;
    const isReadOnly = displayMode === FormDisplayMode.Display;
    const a = item as any;

    const status = isNew ? "Notice" : a?.Status || "Active";
    const step = isNew ? "D0" : a?.Step || "D0";
    const createdOn = item?.Created
      ? this.fmt(item.Created)
      : new Date().toISOString().split("T")[0];
    const modifiedOn = item?.Modified ? this.fmt(item.Modified) : "";
    const ownedOn = item?.Created ? this.fmt(item.Created) : createdOn;

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

    return (
      <div className={styles.ncmrContainer}>
        <div className={styles.metadataBar}>
          {/* Row 1: Title section + toolbar */}
          <div className={styles.metaRow1}>
            <div className={styles.titleSection}>
              <span className={styles.metaLabelBold}>Title</span>
              {!isNew && item?.Title && (
                <span className={styles.itemTitleOrange}>
                  {item.Title}
                  {a?.D0EscalationLevelDisplay
                    ? ` (${a.D0EscalationLevelDisplay})`
                    : ""}
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
                    title="Save as draft (Status: Notice)"
                    onClick={() => this.submitForm(true)}
                  />
                  <PrimaryButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Trophy2" }}
                    text="Submit"
                    title="Submit to SharePoint (Status: Active)"
                    onClick={() => this.submitForm(false)}
                  />
                </>
              ) : (
                <>
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Export" }}
                    text="Export"
                  />
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Ringer" }}
                    text="Subscribe"
                  />
                  <span className={styles.subscribeOffBadge}>Off</span>
                  <DefaultButton
                    className={styles.toolbarBtn}
                    text="Show N.C. Disposition"
                  />
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Comment" }}
                    text="Comments"
                  />
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "Attach" }}
                    text="Attachment"
                  />
                  <DefaultButton
                    className={styles.toolbarBtn}
                    iconProps={{ iconName: "History" }}
                    text="Flow History"
                  />
                </>
              )}
            </div>
          </div>

          {/* Meta columns */}
          <div className={styles.metaColumnsScroll}>
            {metaCols.map((c) => (
              <div key={c.label} className={styles.metaColumn}>
                <span className={styles.metaColLabel}>{c.label}</span>
                <span className={styles.metaColValue}>{c.value}</span>
              </div>
            ))}
          </div>

          {/* Edit / view action row */}
          {!isNew && (
            <div className={styles.editActionsRow}>
              <PrimaryButton
                className={`${styles.toolbarBtn} ${styles.saveBtn}`}
                iconProps={{ iconName: "Save" }}
                text="Save"
                title="Save changes to SharePoint list"
                onClick={() => this.submitForm(false)}
              />
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

        {/* ── Tabs ───────────────────────────────────────────────── */}
        <div className={styles.tabsContainer}>
          <Pivot>
            <PivotItem headerText="D0">
              {this.renderD0(isNew, isReadOnly)}
            </PivotItem>
            <PivotItem headerText="D1">
              {this.renderD1(isNew, isReadOnly)}
            </PivotItem>
            <PivotItem headerText="D2">
              {this.renderD2(isNew, isReadOnly)}
            </PivotItem>

            {/* D3–D8 placeholder tabs */}
            {["D3", "D4", "D5", "D6", "D7", "D8"].map((tab) => (
              <PivotItem key={tab} headerText={tab}>
                <div className={styles.tabContent}>
                  {this.renderNotes(isNew)}
                  <div className={styles.d0FormGrid}>
                    <div className={styles.formColumn}>
                      <div className={styles.fieldBlock}>
                        <TextField
                          required
                          label="Title"
                          placeholder="Title"
                          disabled={isReadOnly}
                        />
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

        {/* ── Footer Save / Submit (below all tab content) ───────── */}
        {!isReadOnly && (
          <div className={styles.formFooterArea}>
            <DefaultButton
              iconProps={{ iconName: "Attach" }}
              text="Attachment"
              className={styles.toolbarBtn}
            />
            <DefaultButton
              iconProps={{ iconName: "Save" }}
              text="Save"
              className={styles.toolbarBtn}
              title={isNew ? "Save as draft (Status: Notice)" : "Save changes to SharePoint list"}
              onClick={() => this.submitForm(isNew ? true : false)}
            />
            {isNew && (
              <PrimaryButton
                iconProps={{ iconName: "Trophy2" }}
                text="Submit"
                className={styles.toolbarBtn}
                title="Submit to SharePoint (Status: Active)"
                onClick={() => this.submitForm(false)}
              />
            )}
          </div>
        )}
      </div>
    );
  }
}
