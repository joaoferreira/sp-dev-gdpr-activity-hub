// tslint:disable-next-line:max-line-length
import * as React from 'react';
import styles from './GdprInsertEvent.module.scss';
import { IGdprInsertEventProps } from './IGdprInsertEventProps';

import * as strings from 'gdprInsertEventStrings';

import pnp from "sp-pnp-js";

import { SPPeoplePicker } from '../../../components/SPPeoplePicker';
import { SPTaxonomyPicker } from '../../../components/SPTaxonomyPicker';
import { ISPTermObject } from '../../../components/SPTermStoreService';
import { SPLookupItemsPicker } from '../../../components/SPLookupItemsPicker';
import { SPDateTimePicker } from '../../../components/SPDateTimePicker';

import { GDPRDataManager } from '../../../components/GDPRDataManager';

/**
 * Common Infrastructure
 */
import {
  BaseComponent,
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';

/**
 * Dialog
 */
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

/**
 * Choice Group
 */
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

/**
 * Text Field
 */
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

/**
 * Toggle
 */
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

/**
 * Button
 */
import { PrimaryButton, DefaultButton, Button, IButtonProps } from 'office-ui-fabric-react/lib/Button';

import { IGdprInsertEventState } from './IGdprInsertEventState';

export default class GdprInsertEvent extends React.Component<IGdprInsertEventProps, IGdprInsertEventState> {

  /**
  * Main constructor for the component
  */
  constructor() {
    super();

    let now: Date = new Date();

    this.state = {
      currentEventType: "DataBreach",
      isValid: false,
      showDialogResult: false,
      includesChildrenInProgress: false,
      toBeDetermined: false,
      indirectDataProvider: false,
      eventStartDate: now,
    };
  }

  public render(): React.ReactElement<IGdprInsertEventProps> {
    return (
      <div className={styles.gdprEvent}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <ChoiceGroup
                selectedKey={this.state.currentEventType}
                label={strings.EventTypeFieldLabel}
                onChange={this._onChangedEventType}
                options={[
                  {
                    key: 'DataBreach',
                    iconProps: { iconName: 'PeopleAlert' },
                    text: strings.EventTypeDataBreachLabel,
                    checked: true,
                  },
                  {
                    key: 'IdentityRisk',
                    iconProps: { iconName: 'SecurityGroup' },
                    text: strings.EventTypeIdentityRiskLabel,
                  },
                  {
                    key: 'DataConsent',
                    iconProps: { iconName: 'ReminderGroup' },
                    text: strings.EventTypeDataConsentLabel,
                  },
                  {
                    key: 'DataConsentWithdrawal',
                    iconProps: { iconName: 'PeopleBlock' },
                    text: strings.EventTypeDataConsentWithdrawalLabel,
                  },
                  {
                    key: 'DataProcessing',
                    iconProps: { iconName: 'PeopleRepeat' },
                    text: strings.EventTypeDataProcessingLabel,
                  },
                  {
                    key: 'DataArchived',
                    iconProps: { iconName: 'Package' },
                    text: strings.EventTypeDataArchivedLabel,
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
                onChanged={this._onChangedTitle}
                value={this.state.title}
                onGetErrorMessage={this._getErrorMessageTitle}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.NotifiedByFieldLabel}
                placeholder={strings.NotifiedByFieldPlaceholder}
                required={true}
                value={this.state.notifiedBy}
                onChanged={this._onChangedNotifiedBy}
                onGetErrorMessage={this._getErrorMessageNotifiedBy}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPPeoplePicker
                context={this.props.context}
                label={strings.EventAssignedToFieldLabel}
                required={true}
                onChanged={this._onChangedEventAssignedTo}
                placeholder={strings.EventAssignedToFieldPlaceholder}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker
                showTime={true}
                includeSeconds={false}
                isRequired={true}
                dateLabel={strings.EventStartDateFieldLabel}
                datePlaceholder={strings.EventStartDateFieldPlaceholder}
                hoursLabel={strings.EventStartTimeHoursFieldLabel}
                hoursValidationError={strings.HoursValidationError}
                minutesLabel={strings.EventStartTimeMinutesFieldLabel}
                minutesValidationError={strings.MinutesValidationError}
                initialDateTime={this.state.eventStartDate}
                onChanged={this._onChangedEventStartDate}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <SPDateTimePicker
                showTime={true}
                includeSeconds={false}
                isRequired={true}
                dateLabel={strings.EventEndDateFieldLabel}
                datePlaceholder={strings.EventEndDateFieldPlaceholder}
                hoursLabel={strings.EventEndTimeHoursFieldLabel}
                hoursValidationError={strings.HoursValidationError}
                minutesLabel={strings.EventEndTimeMinutesFieldLabel}
                minutesValidationError={strings.MinutesValidationError}
                initialDateTime={this.state.eventEndDate}
                onChanged={this._onChangedEventEndDate}
              />
            </div>
          </div>
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField
                label={strings.PostEventReportFieldLabel}
                multiline
                autoAdjustHeight
                required={true}
                value={this.state.postEventReport}
                onChanged={this._onChangedPostEventReport}
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
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Breach Type"
                    label={strings.BreachTypeFieldLabel}
                    placeholder={strings.BreachTypeFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedBreachType}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "IdentityRisk") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Risk Type"
                    label={strings.RiskTypeFieldLabel}
                    placeholder={strings.RiskTypeFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedRiskType}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach" || this.state.currentEventType === "IdentityRisk") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Severity"
                    label={strings.SeverityFieldLabel}
                    placeholder={strings.SeverityFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedSeverity}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.DPANotifiedFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.dpaNotified}
                    onChanged={this._onChangedDPANotified}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  {this.state.dpaNotified ?
                    <SPDateTimePicker
                      showTime={true}
                      includeSeconds={false}
                      isRequired={this.state.dpaNotified}
                      dateLabel={strings.DPANotificationDateFieldLabel}
                      datePlaceholder={strings.DPANotificationDateFieldPlaceholder}
                      hoursLabel={strings.DPANotificationTimeHoursFieldLabel}
                      minutesLabel={strings.DPANotificationTimeMinutesFieldLabel}
                      initialDateTime={this.state.dpaNotificationDate}
                      onChanged={this._onChangedDPANotificationDate}
                    />
                    : null}
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach" && !this.state.toBeDetermined) ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.EstimatedAffectedSubjectsFieldLabel}
                    autoAdjustHeight
                    value={this.state.estimatedAffectedSubjects && this.state.estimatedAffectedSubjects.toString()}
                    onChanged={this._onChangedEstimatedAffectedSubjects}
                    onGetErrorMessage={this._getErrorMessageEstimatedAffectedSubjects}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.ToBeDeterminedFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.toBeDetermined}
                    onChanged={this._onChangedEstimatedAffectedSubjectsToBeDetermined}
                  />
                </div>
              </div>
              : null
          }
          {
            ((this.state.currentEventType === "DataBreach" && !this.state.includesChildrenInProgress) || this.state.currentEventType === "DataArchived") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.IncludesChildrenFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.includesChildren}
                    onChanged={this._onChangedIncludesChildren}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.IncludesChildrenInProgressFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.includesChildrenInProgress}
                    onChanged={this._onChangedIncludesChildrenInProgress}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ActionPlanFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.actionPlan}
                    onChanged={this._onChangedActionPlan}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.BreachResolvedFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.breachResolved}
                    onChanged={this._onChangedBreachResolved}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataBreach") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ActionsTakenFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.actionsTaken}
                    onChanged={this._onChangedActionsTaken}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent" || this.state.currentEventType === "DataArchived") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Sensitive Data Type"
                    label={strings.IncludesSensitiveDataFieldLabel}
                    placeholder={strings.IncludesSensitiveDataFieldPlaceholder}
                    required={false}
                    onChanged={this._onChangedIncludesSensitiveData}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.ConsentIsInternalFieldLabel}
                    onText={strings.InternalConsentText}
                    offText={strings.ExternalConsentText}
                    checked={this.state.consentIsInternal}
                    onChanged={this._onChangedConsentIsInternal}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.DataSubjectIsChildFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.dataSubjectIsChild}
                    onChanged={this._onChangedDataSubjectIsChild}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.IndirectDataProviderFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.indirectDataProvider}
                    onChanged={this._onChangedIndirectDataProvider}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent" && this.state.indirectDataProvider) ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.DataProviderFieldLabel}
                    autoAdjustHeight
                    value={this.state.dataProvider}
                    onChanged={this._onChangedDataProvider}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ConsentNotesFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.consentNotes}
                    onChanged={this._onChangedConsentNotes}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsent") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Consent Type"
                    label={strings.ConsentTypeFieldLabel}
                    placeholder={strings.ConsentTypeFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedConsentType}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsentWithdrawal") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPTaxonomyPicker
                    context={this.props.context}
                    termSetName="Consent Type"
                    label={strings.ConsentWithdrawalTypeFieldLabel}
                    placeholder={strings.ConsentWithdrawalTypeFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedConsentWithdrawalType}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsentWithdrawal") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.OriginalConsentAvailableFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.originalConsentAvailable}
                    onChanged={this._onChangedOriginalConsentAvailable}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsentWithdrawal" && this.state.originalConsentAvailable) ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPLookupItemsPicker
                    sourceListId={this.props.targetList}
                    context={this.props.context}
                    label={strings.OriginalConsentFieldLabel}
                    placeholder={strings.OriginalConsentFieldPlaceholder}
                    required={this.state.originalConsentAvailable}
                    onChanged={this._onChangedOriginalConsent}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataConsentWithdrawal") ?
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
            (this.state.currentEventType === "DataProcessing") ?
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
            (this.state.currentEventType === "DataProcessing") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <SPPeoplePicker
                    context={this.props.context}
                    label={strings.ProcessorsFieldLabel}
                    placeholder={strings.ProcessorsFieldPlaceholder}
                    required={true}
                    onChanged={this._onChangedProcessors}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataArchived") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ArchivedDataFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.archivedData}
                    onChanged={this._onChangedArchivedData}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataArchived") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <Toggle
                    defaultChecked={false}
                    label={strings.AnonymizeFieldLabel}
                    onText={strings.YesText}
                    offText={strings.NoText}
                    checked={this.state.anonymize}
                    onChanged={this._onChangedAnonymize}
                  />
                </div>
              </div>
              : null
          }
          {
            (this.state.currentEventType === "DataArchived") ?
              <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
                <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <TextField
                    label={strings.ArchivingNotesFieldLabel}
                    multiline
                    autoAdjustHeight
                    value={this.state.archivingNotes}
                    onChanged={this._onChangedArchivingNotes}
                  />
                </div>
              </div>
              : null
          }
          <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-black ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <PrimaryButton
                data-automation-id='saveRequest'
                text={strings.SaveButtonText}
                disabled={!this.state.isValid}
                onClick={this._saveClick}
              />
              &nbsp;&nbsp;
              <Button
                data-automation-id='cancel'
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
          isBlocking={true}>
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

  private _getErrorMessageTitle(value: string): string {
    return (value == null || value.length == 0 || value.length >= 10)
      ? ''
      : `${strings.TitleFieldValidationErrorMessage} ${value.length}.`;
  }

  private _getErrorMessageNotifiedBy(value: string): string {
    return (value == null || value.length == 0 || value.length >= 5)
      ? ''
      : `${strings.NotifiedByFieldValidationErrorMessage} ${value.length}.`;
  }



  @autobind
  private _onChangedEventType(ev: React.FormEvent<HTMLInputElement>, option: any): void {
    // let isValid: boolean = this._formIsValid();
    this.setState((prevState: any, props: any): any => ({
      ...this.state,
      currentEventType: option.key,
      breachType: null,
      riskType: null,
      severity: null,
      includesSensitiveData: null,
      consentType: [],
      consentWithdrawalType: [],
      originalConsent: 0,
      processingType: null,
      processors: [],
      // isValid: isValid
    }));
  }

  @autobind
  private _onChangedEstimatedAffectedSubjectsToBeDetermined(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, toBeDetermined: checked, isValid: isValid }));
  }

  private _getErrorMessageEstimatedAffectedSubjects(value: string): string {

    if (value != null && value.length > 0 && isNaN(Number(value))) {
      return (strings.EstimatedAffectedSubjectsFieldValidationErrorMessage);
    } else {
      return ("");
    }
  }

  @autobind
  private _onChangedTitle(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, title: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedNotifiedBy(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, notifiedBy: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedEventAssignedTo(items: string[]): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: items[0] });
    this.setState((prevState: any, props: any): any => ({ ...this.state, eventAssignedTo: items[0], isValid: isValid }));
  }

  @autobind
  private _onChangedEventStartDate(newValue: Date): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, eventStartDate: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedEventEndDate(newValue: Date): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, eventEndDate: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedPostEventReport(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, postEventReport: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedAdditionalNotes(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, additionalNotes: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedBreachType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, breachType: terms[0], isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, breachType: null, isValid: isValid }));
    }
  }

  @autobind
  private _onChangedRiskType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, riskType: terms, isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, riskType: [], isValid: isValid }));
    }
  }

  @autobind
  private _onChangedSeverity(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, severity: terms[0], isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, severity: null, isValid: isValid }));
    }
  }

  @autobind
  private _onChangedDPANotified(newValue: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dpaNotified: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedDPANotificationDate(newValue: Date): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dpaNotificationDate: newValue, isValid: isValid }));

  }

  @autobind
  private _onChangedIncludesChildrenInProgress(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, includesChildrenInProgress: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedEstimatedAffectedSubjects(newValue: number): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, estimatedAffectedSubjects: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedIncludesChildren(newValue: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, includesChildren: newValue, isValid: isValid }));

  }

  @autobind
  private _onChangedActionPlan(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, actionPlan: newValue, isValid: isValid }));

  }

  @autobind
  private _onChangedBreachResolved(newValue: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, breachResolved: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedActionsTaken(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, actionsTaken: newValue, isValid: isValid }));

  }

  @autobind
  private _onChangedIncludesSensitiveData(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, includesSensitiveData: terms[0], isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, includesSensitiveData: null, isValid: isValid }));
    }
  }

  @autobind
  private _onChangedConsentIsInternal(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, consentIsInternal: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedDataSubjectIsChild(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dataSubjectIsChild: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedIndirectDataProvider(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, indirectDataProvider: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedDataProvider(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, dataProvider: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedConsentNotes(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, consentNotes: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedConsentType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, consentType: terms, isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, consentType: [], isValid: isValid }));
    }
  }

  @autobind
  private _onChangedConsentWithdrawalType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, consentWithdrawalType: terms, isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, consentWithdrawalType: [], isValid: isValid }));
    }
  }

  @autobind
  private _onChangedConsentWithdrawalNotes(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, consentWithdrawalNotes: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedOriginalConsentAvailable(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, originalConsentAvailable: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedOriginalConsent(selectedItemIds: number[]): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: selectedItemIds[0] });
    this.setState((prevState: any, props: any): any => ({ ...this.state, originalConsent: selectedItemIds[0], isValid: isValid }));
  }

  @autobind
  private _onChangedNotifyApplicable(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, notifyApplicable: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedProcessingType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: terms[0] });
      this.setState((prevState: any, props: any): any => ({ ...this.state, processingType: terms, isValid: isValid }));
    } else {
      let isValid: any = this._formIsValid({ ...this.state, dataSubject: null });
      this.setState((prevState: any, props: any): any => ({ ...this.state, processingType: [], isValid: isValid }));
    }
  }

  @autobind
  private _onChangedProcessors(items: string[]): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: items[0] });
    this.setState((prevState: any, props: any): any => ({ ...this.state, processors: items, isValid: isValid }));
  }

  @autobind
  private _onChangedArchivedData(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, archivedData: newValue, isValid: isValid }));
  }

  @autobind
  private _onChangedAnonymize(checked: boolean): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: checked });
    this.setState((prevState: any, props: any): any => ({ ...this.state, anonymize: checked, isValid: isValid }));
  }

  @autobind
  private _onChangedArchivingNotes(newValue: string): void {
    let isValid: any = this._formIsValid({ ...this.state, dataSubject: newValue });
    this.setState((prevState: any, props: any): any => ({ ...this.state, archivingNotes: newValue, isValid: isValid }));
  }

  @autobind
  private _saveClick(event: any): void {
    event.preventDefault();
    if (this._formIsValid(this.state)) {
      let dataManager: any = new GDPRDataManager();
      dataManager.setup({
        eventsListId: this.props.targetList,
      });

      let eventItem: any = {
        kind: this.state.currentEventType,
        title: this.state.title,
        notifiedBy: this.state.notifiedBy,
        eventAssignedTo: this.state.eventAssignedTo,
        eventStartDate: this.state.eventStartDate,
        eventEndDate: this.state.eventEndDate,
        postReport: this.state.postEventReport,
        additionalNotes: this.state.additionalNotes,
      };

      switch (eventItem.kind) {
        case "DataBreach":
          eventItem.breachType = {
            Label: this.state.breachType.name,
            TermGuid: this.state.breachType.guid,
            WssId: -1,
          };
          eventItem.severity = {
            Label: this.state.severity.name,
            TermGuid: this.state.severity.guid,
            WssId: -1,
          };
          eventItem.dpaNotified = this.state.dpaNotified;
          eventItem.dpaNotificationDate = this.state.dpaNotificationDate;
          eventItem.estimatedNumberOfAffectedDataSubjects = this.state.estimatedAffectedSubjects;
          eventItem.toBeDetermined = this.state.toBeDetermined;
          eventItem.includesChildrenData = this.state.includesChildren;
          eventItem.inProgress = this.state.includesChildrenInProgress;
          eventItem.actionPlan = this.state.actionPlan;
          eventItem.breachResolved = this.state.breachResolved;
          eventItem.actionsTaken = this.state.actionsTaken;
          break;
        case "IdentityRisk":
          eventItem.riskType = this.state.riskType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          eventItem.severity = {
            Label: this.state.severity.name,
            TermGuid: this.state.severity.guid,
            WssId: -1,
          };
          break;
        case "DataConsent":
          eventItem.consentIsInternal = this.state.consentIsInternal;
          if (this.state.includesSensitiveData) {
            eventItem.includesSensitiveData = {
              Label: this.state.includesSensitiveData.name,
              TermGuid: this.state.includesSensitiveData.guid,
              WssId: -1,
            };
          }
          eventItem.dataSubjectIsChild = this.state.dataSubjectIsChild;
          eventItem.indirectDataProvider = this.state.indirectDataProvider;
          eventItem.dataProvider = this.state.dataProvider;
          eventItem.consentNotes = this.state.consentNotes;
          eventItem.consentType = this.state.consentType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          break;
        case "DataConsentWithdrawal":
          eventItem.withdrawalType = this.state.consentWithdrawalType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          eventItem.withdrawalNotes = this.state.consentWithdrawalNotes;
          eventItem.originalConsentId = this.state.originalConsent;
          eventItem.notifyThirdParties = this.state.notifyApplicable;
          eventItem.originalConsentAvailable = this.state.originalConsentAvailable;
          break;
        case "DataProcessing":
          eventItem.processingType = this.state.processingType.map(i => {
            return {
              Label: i.name,
              TermGuid: i.guid,
              WssId: -1,
            };
          });
          eventItem.processors = this.state.processors;
          break;
        case "DataArchived":
          eventItem.archivedData = this.state.archivedData;
          if (this.state.includesSensitiveData) {
            eventItem.includesSensitiveData = {
              Label: this.state.includesSensitiveData.name,
              TermGuid: this.state.includesSensitiveData.guid,
              WssId: -1,
            };
          }
          eventItem.includesChildrenData = this.state.includesChildren;
          eventItem.anonymize = this.state.anonymize;
          eventItem.archivingNotes = this.state.archivingNotes;
          break;
      }
      dataManager.insertNewEvent(eventItem).then((itemId: number) => {
        this.setState((prevState: any, props: any): any => ({ ...this.state, showDialogResult: true}));
      });
    }
  }

  @autobind
  private _cancelClick(event: any): void {
    event.preventDefault();
    window.history.back();
  }

  private _formIsValid(state: any): boolean {
    let isValid: boolean =
      (state.title != null && state.title.length > 0) &&
      (state.notifiedBy != null && state.notifiedBy.length > 0) &&
      (state.eventAssignedTo != null && state.eventAssignedTo.length > 0) &&
      (state.eventStartDate != null) &&
      (state.postEventReport != null && state.postEventReport.length > 0);

    if (state.currentEventType == "DataBreach") {
      isValid = isValid && state.breachType != null;
      isValid = isValid && state.severity != null;
      isValid = isValid && ((state.dpaNotified && state.dpaNotificationDate != null) || (!state.dpaNotified));
    }
    if (state.currentEventType == "IdentityRisk") {
      isValid = isValid && state.riskType != null;
      isValid = isValid && state.severity != null;
    }
    if (state.currentEventType == "DataConsent") {
      isValid = isValid && state.consentType != null && state.consentType.length > 0;
    }
    if (state.currentEventType == "DataConsentWithdrawal") {
      isValid = isValid && state.consentWithdrawalType != null && state.consentWithdrawalType.length > 0;
      isValid = isValid && ((state.originalConsentAvailable && state.originalConsent > 0) || (!state.originalConsentAvailable));
    }
    if (state.currentEventType == "DataProcessing") {
      isValid = isValid && state.processingType != null && state.processingType.length > 0;
      isValid = isValid && state.processors != null && state.processors.length > 0;
    }
    if (state.currentEventType == "DataArchived") {
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
