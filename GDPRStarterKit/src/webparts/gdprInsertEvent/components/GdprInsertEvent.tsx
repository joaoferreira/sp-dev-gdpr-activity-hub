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
                label={strings.SaveButtonText}
                disabled={!this.state.isValid}
                onClick={this._saveClick}
              />
              &nbsp;&nbsp;
              <Button
                data-automation-id='cancel'
                label={strings.CancelButtonText}
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
              label={strings.InsertNextLabel} />
            <DefaultButton
              onClick={this._goHomeClick}
              label={strings.GoHomeLabel} />
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
  private _updateState(state: IGdprInsertEventState): void {
    state.isValid = this._formIsValid();
    this.setState(state);
  }

  @autobind
  private _onChangedEventType(ev: React.FormEvent<HTMLInputElement>, option: any) {
    let state: any = this.state;
    state.currentEventType = option.key;
    state.breachType = null;
    state.riskType = null;
    state.severity = null;
    state.includesSensitiveData = null;
    state.consentType = [];
    state.consentWithdrawalType = [];
    state.originalConsent = 0;
    state.processingType = null;
    state.processors = [];

    this._updateState(state);
  }

  @autobind
  private _onChangedEstimatedAffectedSubjectsToBeDetermined(checked: boolean): void {
    this._updateState({ ...this.state, toBeDetermined: checked });
  }

  private _getErrorMessageEstimatedAffectedSubjects(value: string): string {

    if (value != null && value.length > 0 && isNaN(Number(value))) {
      return (strings.EstimatedAffectedSubjectsFieldValidationErrorMessage);
    }
    else {
      return ("");
    }
  }

  @autobind
  private _onChangedTitle(newValue: string): void {
    this._updateState({ ...this.state, title: newValue });
  }

  @autobind
  private _onChangedNotifiedBy(newValue: string): void {
    this._updateState({ ...this.state, notifiedBy: newValue });
  }

  @autobind
  private _onChangedEventAssignedTo(items: string[]): void {
    this._updateState({ ...this.state, eventAssignedTo: items[0] });
  }

  @autobind
  private _onChangedEventStartDate(newValue: Date): void {
    this._updateState({ ...this.state, eventStartDate: newValue });
  }

  @autobind
  private _onChangedEventEndDate(newValue: Date): void {
    this._updateState({ ...this.state, eventEndDate: newValue });
  }

  @autobind
  private _onChangedPostEventReport(newValue: string): void {
    this._updateState({ ...this.state, postEventReport: newValue });
  }

  @autobind
  private _onChangedAdditionalNotes(newValue: string): void {
    this._updateState({ ...this.state, additionalNotes: newValue });
  }

  @autobind
  private _onChangedBreachType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, breachType: terms[0] });
    }
    else {
      this._updateState({ ...this.state, breachType: null });
    }
  }

  @autobind
  private _onChangedRiskType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, riskType: terms });
    }
    else {
      this._updateState({ ...this.state, riskType: [] });
    }
  }

  @autobind
  private _onChangedSeverity(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, severity: terms[0] });
    }
    else {
      this._updateState({ ...this.state, severity: null });
    }
  }

  @autobind
  private _onChangedDPANotified(newValue: boolean): void {
    this._updateState({ ...this.state, dpaNotified: newValue });
  }

  @autobind
  private _onChangedDPANotificationDate(newValue: Date): void {
    this._updateState({ ...this.state, dpaNotificationDate: newValue });

  }

  @autobind
  private _onChangedIncludesChildrenInProgress(checked: boolean): void {
    this._updateState({ ...this.state, includesChildrenInProgress: checked });
  }

  @autobind
  private _onChangedEstimatedAffectedSubjects(newValue: number): void {
    this._updateState({ ...this.state, estimatedAffectedSubjects: newValue });
  }

  @autobind
  private _onChangedIncludesChildren(newValue: boolean): void {
    this._updateState({ ...this.state, includesChildren: newValue });

  }

  @autobind
  private _onChangedActionPlan(newValue: string): void {
    this._updateState({ ...this.state, actionPlan: newValue });

  }

  @autobind
  private _onChangedBreachResolved(newValue: boolean): void {
    this._updateState({ ...this.state, breachResolved: newValue });
  }

  @autobind
  private _onChangedActionsTaken(newValue: string): void {
    this._updateState({ ...this.state, actionsTaken: newValue });

  }

  @autobind
  private _onChangedIncludesSensitiveData(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, includesSensitiveData: terms[0] });
    }
    else {
      this._updateState({ ...this.state, includesSensitiveData: null });
    }
  }

  @autobind
  private _onChangedConsentIsInternal(checked: boolean): void {
    this._updateState({ ...this.state, consentIsInternal: checked });
  }

  @autobind
  private _onChangedDataSubjectIsChild(checked: boolean): void {
    this._updateState({ ...this.state, dataSubjectIsChild: checked });
  }

  @autobind
  private _onChangedIndirectDataProvider(checked: boolean): void {
    this._updateState({ ...this.state, indirectDataProvider: checked });
  }

  @autobind
  private _onChangedDataProvider(newValue: string): void {
    this._updateState({ ...this.state, dataProvider: newValue });
  }

  @autobind
  private _onChangedConsentNotes(newValue: string): void {
    this._updateState({ ...this.state, consentNotes: newValue });
  }

  @autobind
  private _onChangedConsentType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, consentType: terms });
    }
    else {
      this._updateState({ ...this.state, consentType: [] });
    }
  }

  @autobind
  private _onChangedConsentWithdrawalType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, consentWithdrawalType: terms });
    }
    else {
      this._updateState({ ...this.state, consentWithdrawalType: [] });
    }
  }

  @autobind
  private _onChangedConsentWithdrawalNotes(newValue: string): void {
    this._updateState({ ...this.state, consentWithdrawalNotes: newValue });
  }

  @autobind
  private _onChangedOriginalConsentAvailable(checked: boolean): void {
    this._updateState({ ...this.state, originalConsentAvailable: checked });
  }

  @autobind
  private _onChangedOriginalConsent(selectedItemIds: number[]): void {
    this._updateState({ ...this.state, originalConsent: selectedItemIds[0] });
  }

  @autobind
  private _onChangedNotifyApplicable(checked: boolean): void {
    this._updateState({ ...this.state, notifyApplicable: checked });
  }

  @autobind
  private _onChangedProcessingType(terms: ISPTermObject[]): void {
    if (terms != null && terms.length > 0) {
      this._updateState({ ...this.state, processingType: terms });
    }
    else {
      this._updateState({ ...this.state, processingType: [] });
    }
  }

  @autobind
  private _onChangedProcessors(items: string[]): void {
    this._updateState({ ...this.state, processors: items });
  }

  @autobind
  private _onChangedArchivedData(newValue: string): void {
    this._updateState({ ...this.state, archivedData: newValue });
  }

  @autobind
  private _onChangedAnonymize(checked: boolean): void {
    this._updateState({ ...this.state, anonymize: checked });
  }

  @autobind
  private _onChangedArchivingNotes(newValue: string): void {
    this._updateState({ ...this.state, archivingNotes: newValue });
  }

  @autobind
  private _saveClick(event) {
    event.preventDefault();
    if (this._formIsValid()) {
      let dataManager = new GDPRDataManager();
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
        this._updateState({ ...this.state, showDialogResult: true });
      });
    }
  }

  @autobind
  private _cancelClick(event) {
    event.preventDefault();
    window.history.back();
  }

  private _formIsValid(): boolean {
    let isValid: boolean =
      (this.state.title != null && this.state.title.length > 0) &&
      (this.state.notifiedBy != null && this.state.notifiedBy.length > 0) &&
      (this.state.eventAssignedTo != null && this.state.eventAssignedTo.length > 0) &&
      (this.state.eventStartDate != null) &&
      (this.state.postEventReport != null && this.state.postEventReport.length > 0);

    if (this.state.currentEventType == "DataBreach") {
      isValid = isValid && this.state.breachType != null;
      isValid = isValid && this.state.severity != null;
      isValid = isValid && ((this.state.dpaNotified && this.state.dpaNotificationDate != null) || (!this.state.dpaNotified));
    }
    if (this.state.currentEventType == "IdentityRisk") {
      isValid = isValid && this.state.riskType != null;
      isValid = isValid && this.state.severity != null;
    }
    if (this.state.currentEventType == "DataConsent") {
      isValid = isValid && this.state.consentType != null && this.state.consentType.length > 0;
    }
    if (this.state.currentEventType == "DataConsentWithdrawal") {
      isValid = isValid && this.state.consentWithdrawalType != null && this.state.consentWithdrawalType.length > 0;
      isValid = isValid && ((this.state.originalConsentAvailable && this.state.originalConsent > 0) || (!this.state.originalConsentAvailable));
    }
    if (this.state.currentEventType == "DataProcessing") {
      isValid = isValid && this.state.processingType != null && this.state.processingType.length > 0;
      isValid = isValid && this.state.processors != null && this.state.processors.length > 0;
    }
    if (this.state.currentEventType == "DataArchived") {
    }

    return (isValid);
  }

  @autobind
  private _closeInsertDialogResult() {
    this._updateState({ ...this.state, showDialogResult: false });
  }

  @autobind
  private _insertNextClick(event) {
    event.preventDefault();
    this._closeInsertDialogResult();
  }

  @autobind
  private _goHomeClick(event) {
    event.preventDefault();
    pnp.sp.web.select("Url").get().then((web) => {
      window.location.replace(web.Url);
    });
  }
}
