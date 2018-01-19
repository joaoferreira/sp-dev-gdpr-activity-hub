import * as React from "react";
import styles from "./GdprInsertRequest.module.scss";
import { IGdprInsertRequestProps } from "./IGdprInsertRequestProps";

import * as strings from "gdprInsertRequestStrings";

import pnp from "sp-pnp-js";

import { SPPeoplePicker } from "../../../components/SPPeoplePicker";
import { SPTaxonomyPicker } from "../../../components/SPTaxonomyPicker";
import { ISPTermObject } from "../../../components/SPTermStoreService";
import { SPDateTimePicker } from "../../../components/SPDateTimePicker";

import { GDPRDataManager } from "../../../components/GDPRDataManager";

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from "office-ui-fabric-react/lib/Utilities";

/**
 * Dialog
 */
import { Dialog, DialogType, DialogFooter } from "office-ui-fabric-react/lib/Dialog";

/**
 * Choice Group
 */
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react/lib/ChoiceGroup";

/**
 * Text Field
 */
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Icon } from "office-ui-fabric-react/lib/Icon";

/**
 * Toggle
 */
import { Toggle } from "office-ui-fabric-react/lib/Toggle";

/**
 * Button
 */
import { PrimaryButton, DefaultButton, Button, IButtonProps } from "office-ui-fabric-react/lib/Button";

import { IGdprInsertRequestState } from "./IGdprInsertRequestState";

export default class GdprInsertRequest extends React.Component<IGdprInsertRequestProps, IGdprInsertRequestState> {

  /**
   * Main constructor for the component
   */
  constructor() {
    super();

    this.state = {
      currentRequestType: "Access",
      isValid: false,
      showDialogResult: false,
    };
  }

  public render(): React.ReactElement<IGdprInsertRequestProps> {
    console.log(this.state.currentRequestType);
    return (
      <div className={styles.gdprRequest}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <ChoiceGroup
                label={strings.RequestTypeFieldLabel}
                onChange={this._onChangedRequestType}
                selectedKey={this.state.currentRequestType}
                options={[
                  {
                    key: "Access",
                    iconProps: { iconName: "QuickNote" },
                    text: strings.RequestTypeAccessLabel,
                    checked: true,
                  },
                  {
                    key: "Correct",
                    iconProps: { iconName: "EditNote" },
                    text: strings.RequestTypeCorrectLabel,
                  },
                  {
                    key: "Export",
                    iconProps: { iconName: "NoteForward" },
                    text: strings.RequestTypeExportLabel,
                  },
                  {
                    key: "Objection",
                    iconProps: { iconName: "NoteReply" },
                    text: strings.RequestTypeObjectionLabel,
                  },
                  {
                    key: "Erase",
                    iconProps: { iconName: "EraseTool" },
                    text: strings.RequestTypeEraseLabel,
                  }
                ]}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.TitleFieldLabel}
                placeholder={strings.TitleFieldPlaceholder}
                required={true}
                value={this.state.title}
                onChanged={this._onChangedTitle}
                onGetErrorMessage={this._getErrorMessageTitle}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.DataSubjectFieldLabel}
                placeholder={strings.DataSubjectFieldPlaceholder}
                required={true}
                value={this.state.dataSubject}
                onChanged={this._onChangedDataSubject}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.DataSubjectEmailFieldLabel}
                placeholder={strings.DataSubjectEmailFieldPlaceholder}
                required={false}
                value={this.state.dataSubjectEmail}
                onChanged={this._onChangedDataSubjectEmail}
                onGetErrorMessage={this._getErrorMessageDataSubjectEmail}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <Toggle
                defaultChecked={false}
                label={strings.VerifiedDataSubjectFieldLabel}
                onText={strings.YesText}
                offText={strings.NoText}
                checked={this.state.verifiedDataSubject}
                onChanged={this._onChangedVerifiedDataSubject}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPPeoplePicker
                context={this.props.context}
                label={strings.RequestAssignedToFieldLabel}
                required={true}
                placeholder={strings.RequestAssignedToFieldPlaceholder}
                onChanged={this._onChangedRequestAssignedTo}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker
                showTime={false}
                includeSeconds={false}
                dateLabel={strings.RequestInsertionDateFieldLabel}
                datePlaceholder={strings.RequestInsertionDateFieldPlaceholder}
                isRequired={true}
                initialDateTime={this.state.requestInsertionDate}
                onChanged={this._onChangedRequestInsertionDate}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker
                showTime={false}
                includeSeconds={false}
                dateLabel={strings.RequestDueDateFieldLabel}
                datePlaceholder={strings.RequestDueDateFieldPlaceholder}
                isRequired={true}
                initialDateTime={this.state.requestDueDate}
                onChanged={this._onChangedRequestDueDate}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.AdditionalNotesFieldLabel}
                multiline
                autoAdjustHeight
                value={this.state.additionalNotes}
                onChanged={this._onChangedAdditionalNotes}
              />
            </div>
          </div>
          {
            (this.state.currentRequestType === "Access" || this.state.currentRequestType === "Export") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Delivery Methods"
                    label={strings.DeliveryMethodFieldLabel}
                    placeholder={strings.DeliveryMethodFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedDeliveryMethod}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Correct") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.CorrectionDefinitionFieldLabel}
                    multiline
                    autoAdjustHeight
                    required={true}
                    value={this.state.correctionDefinition}
                    onChanged={this._onChangedCorrectionDefinition}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Export") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Delivery Format"
                    label={strings.DeliveryFormatFieldLabel}
                    placeholder={strings.DeliveryFormatFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedDeliveryFormat}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Objection") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.PersonalDataFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.personalData}
                    onChanged={this._onChangedPersonalData}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Objection") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Processing Type"
                    label={strings.ProcessingTypeFieldLabel}
                    placeholder={strings.ProcessingTypeFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedProcessingType}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Erase") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.NotifyApplicableFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.notifyApplicable}
                    onChanged={this._onChangedNotifyApplicable}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentRequestType === "Objection" || this.state.currentRequestType === "Erase") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ReasonFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.reason}
                    onChanged={this._onChangedReason}
                  />
                </div>
              </div>
              : null
          }
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <PrimaryButton
                data-automation-id="saveRequest"
                text={strings.SaveButtonText}
                disabled={!this.state.isValid}
                onClick={this._saveClick}
              />
              &nbsp;&nbsp;
              <Button
                data-automation-id="cancel"
                text={strings.CancelButtonText}
                onClick={this._cancelClick}
              />
            </div>
          </div>
        </div>
        <Dialog
          isOpen={this.state.showDialogResult}
          type={DialogType.normal}
          onDismiss={this._closeInsertDialogResult}
          title={strings.ItemInsertedDialogTitle}
          subText={strings.ItemInsertedDialogSubText}
          isBlocking={true}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={this._insertNextClick}
              text={strings.InsertNextLabel} />
            <DefaultButton
              onClick={this._goHomeClick}
              text={strings.GoHomeLabel} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  @autobind
  private _onChangedRequestType(ev: React.FormEvent<HTMLInputElement>, option: any): any {
    let state: any = this.state;
    state.currentRequestType = option.key;
    state.deliveryMethod = null;
    state.deliveryFormat = null;
    state.processingType = [];
    this.setState({ ...this.state, ...state });
  }

  @autobind
  private _getErrorMessageTitle(value: string): string {

    return (value == null || value.length == 0 || value.length >= 10)
      ? ''
      : `${strings.TitleFieldValidationErrorMessage} ${value.length}.`;
  }

  private _getErrorMessageDataSubjectEmail(value: string): string {

    let emailRegEx: RegExp = new RegExp(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/);

    if (value != null && value.length > 0 && !emailRegEx.test(value)) {
      return (strings.DataSubjectEmailFieldValidationErrorMessage);
    } else {
      return ("");
    }
  }

  @autobind
  private _onChangedTitle(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, title: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, title: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedDataSubject(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dataSubject: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedDataSubjectEmail(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dataSubjectEmail: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedVerifiedDataSubject(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, verifiedDataSubject: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedRequestAssignedTo(items: string[]): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: items });
    this.setState((prevState: any, props: any): any => ({ ...this.state, requestAssignedTo: items[0], isValid: isValid }));
  }

  @autobind
  private _onChangedRequestInsertionDate(newValue: Date): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, requestInsertionDate: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedRequestDueDate(newValue: Date): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, requestDueDate: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedAdditionalNotes(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, additionalNotes: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedDeliveryMethod(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, deliveryMethod: terms[0], isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, deliveryMethod: null, isValid: isValid }));
    }
  }

  @autobind
  private _onChangedCorrectionDefinition(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, correctionDefinition: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedDeliveryFormat(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, deliveryFormat: terms[0], isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, deliveryFormat: null, isValid: isValid }));
    }
  }

  @autobind
  private _onChangedPersonalData(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, personalData: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedProcessingType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms });
      this.setState((prevState: any, props: any): any => ({ ...this.state, processingType: terms, isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: [] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, processingType: [], isValid: isValid }));
    }
  }

  @autobind
  private _onChangedNotifyApplicable(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, notifyApplicable: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedReason(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, reason: newValue, isValid: isValid }));
  }

  @autobind
  private _saveClick(event: any): any {
    event.preventDefault();
    if (this._formIsValid(this.state)) {
      let dataManager: any = new GDPRDataManager();
      dataManager.setup({
        requestsListId: this.props.targetList,
      });

      let request: any = {
        kind: this.state.currentRequestType,
        title: this.state.title,
        dataSubject: this.state.dataSubject,
        dataSubjectEmail: this.state.dataSubjectEmail,
        verifiedDataSubject: this.state.verifiedDataSubject,
        requestAssignedTo: this.state.requestAssignedTo,
        requestInsertionDate: this.state.requestInsertionDate,
        requestDueDate: this.state.requestDueDate,
        additionalNotes: this.state.additionalNotes,
      };

      switch (request.kind) {
        case "Access":
          request.deliveryMethod = {
            Label: this.state.deliveryMethod.name,
            TermGuid: this.state.deliveryMethod.guid,
            WssId: -1,
          };
          break;
        case "Correct":
          request.correctionDefinition = this.state.correctionDefinition;
          break;
        case "Erase":
          request.notifyThirdParties = this.state.notifyApplicable;
          request.reason = this.state.reason;
          break;
        case "Export":
          request.deliveryMethod = {
            Label: this.state.deliveryMethod.name,
            TermGuid: this.state.deliveryMethod.guid,
            WssId: -1,
          };
          request.deliveryFormat = {
            Label: this.state.deliveryFormat.name,
            TermGuid: this.state.deliveryFormat.guid,
            WssId: -1,
          };
          break;
        case "Objection":
          request.personalData = this.state.personalData;
          request.processingType = this.state.processingType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          request.reason = this.state.reason;
          break;
      }

      dataManager.insertNewRequest(request).then((itemId: number) => {
        this.setState((prevState: any, props: any): any => ({ ...this.state, showDialogResult: true }));
      });
    }
  }

  @autobind
  private _cancelClick(event: any): any {
    event.preventDefault();
    window.history.back();
  }

  private _formIsValid(state: any): boolean {
    let isValid: boolean =
      (state.title != null && state.title.length > 0) &&
      (state.dataSubject != null && state.dataSubject.length > 0) &&
      (state.requestAssignedTo != null && state.requestAssignedTo.length > 0) &&
      (state.requestInsertionDate != null) &&
      (state.requestDueDate != null);

    if (state.currentRequestType == "Access" || state.currentRequestType == "Export") {
      isValid = isValid && state.deliveryMethod != null;
    }
    if (state.currentRequestType == "Export") {
      isValid = isValid && state.deliveryFormat != null;
    }
    if (state.currentRequestType == "Correct") {
      isValid = isValid && state.correctionDefinition != null && state.correctionDefinition.length > 0;
    }
    if (state.currentRequestType == "Objection") {
      isValid = isValid && state.processingType != null && state.processingType.length > 0;
    }
    return (isValid);
  }

  @autobind
  private _closeInsertDialogResult(): void {
    this.setState((prevState: any, props: any): any => ({ ...this.state, showDialogResult: false}));
  }

  @autobind
  private _insertNextClick(event: any): void {
    event.preventDefault();
    this._closeInsertDialogResult();
  }

  @autobind
  private _goHomeClick(event: any): void {
    event.preventDefault();
    pnp.sp.web.select("Url").get().then((web) => {
      window.location.replace(web.Url);
    });
  }
}
