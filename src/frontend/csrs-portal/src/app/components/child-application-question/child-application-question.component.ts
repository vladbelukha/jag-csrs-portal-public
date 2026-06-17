import { HttpClient, HttpResponse } from '@angular/common/http';
import { Component, Inject, OnInit } from '@angular/core';
import { FormArray, FormBuilder, FormGroup, Validators } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { LoggerService } from '@core/services/logger.service';
import { ConfirmDialogComponent } from '@shared/dialogs/confirm-dialog/confirm-dialog.component';
import { AccountService } from 'app/api/api/account.service';
import { LookupService } from 'app/api/api/lookup.service';

import { MatDatepickerInputEvent } from '@angular/material/datepicker';
import { MatSelectChange } from '@angular/material/select';
import { Router } from '@angular/router';
import { DialogOptions } from '@shared/dialogs/dialog-options.model';
import {
  OidcSecurityService,
  PublicEventsService,
} from 'angular-auth-oidc-client';
import { ModalDialogComponent } from 'app/components/modal-dialog/modal-dialog.component';

// -- import data structure
import {
  Child,
  FileStatus,
  LookupValue,
  NewFileRequest,
  Party,
  PartyRole,
} from 'app/api/model/models';

import { DatePipe } from '@angular/common';
import { ModalDialogHtmlComponent } from '../modal-dialog-htmlcontent/modal-dialog-htmlcontent.component';

@Component({
  selector: 'app-child-application-question',
  templateUrl: './child-application-question.component.html',
  styleUrls: ['./child-application-question.component.scss'],
})
export class ChildApplicationQuestionComponent implements OnInit {
  firstFormGroup: FormGroup;
  secondFormGroup: FormGroup;
  thirdFormGroup: FormGroup;
  fourthFormGroup;
  fourthFormGroup1: FormGroup;
  fifthFormGroup: FormGroup;
  sixFormGroup: FormGroup;
  seventhFormGroup: FormGroup;
  eFormGroup: FormGroup;
  nineFormGroup: FormGroup;

  provinces: any = [];
  genders: any = [];
  identities: any = [];
  referrals: any = [];
  preferredContactMethods: any = [];
  courtLevels: any = [];
  courtLocations: any = [];
  PartyRole = PartyRole;

  today = new Date();
  isEditable: boolean = false;
  isDisabledSubmit: boolean = true;
  isChildDelete: boolean = true;
  child: Child;
  _reponse: HttpResponse<any>;

  _yes: number = 867670000;
  _no: number = 867670001;
  _iDontKnow: number = 867670002;

  _courtOrder: number = 867670000;
  _writtenAgreement: number = 867670001;

  data: any = null;
  result: any = [];

  errorMessage: any = '';
  errorMailMessage: any = '';
  errorIncomeMessage: any = '';
  errorDateMessage: any = '';
  errorMaxMessage: any = '';
  errorMaxOtherIdentityMessage: any = '';
  tooltips: any = [];
  isHiddens: any = [];

  dateOfOrder: Date;
  birthOfDateOtherParty: Date;
  isVisibleOtherIdentity: boolean = false;

  constructor(
    public oidc: OidcSecurityService,
    private eventService: PublicEventsService,
    private _formBuilder: FormBuilder,
    private http: HttpClient,
    @Inject(AccountService) private accountService,
    @Inject(LookupService) private lookupService,
    @Inject(LoggerService) private logger,
    @Inject(Router) private router,
    public dialog: MatDialog,
    private datePipe: DatePipe,
  ) {}

  ngOnInit() {
    this.provinces = [{ id: '123', value: 'British Columbia' }];
    this.identities = [{ id: '123', value: 'Native' }];
    this.genders = [{ id: '123', value: 'Male' }];
    this.courtLocations = [{ id: '123', value: 'Victoria Court' }];
    this.referrals = [{ id: '123', value: 'BCFMA' }];

    this.errorMessage = 'Error: Field is required. ';
    this.errorMailMessage = 'Email address without @ or domain name. ';
    this.errorIncomeMessage = 'Field should have numerical values. ';
    this.errorDateMessage = 'Date cannot be in future.';
    this.errorMaxMessage = 'You can only enter up to 3000 characters.';

    this.errorMaxOtherIdentityMessage =
      'You can only enter up to 100 characters.';

    this.tooltips = [
      'A child over the age of majority (19 in B.C.) who is still dependent on their parents. For example, due to illness, disability or pursuit of post secondary education.',
      'On a Court Order: look for the date the order was made or granted. Often this can be found on the first page below the names of the parties. On a Written Agreement: look for the date the agreement was stamped when it was filed by the court. Often this can be found at the top of the first page.',
      'The person paying child support.',
      'These are additional amounts to be paid over and above the base child support amount. Some examples of these expenses (often referred to as Section 7 expenses) include childcare, medical or dental premiums, healthcare costs.',
      'Income Assistance (IA) is the welfare program in BC. It provides financial support for low income or no income individuals. ' +
        'The Ministry of Social Development and Poverty Reduction has three income assistance programs: ' +
        'Income Assistance (IA), ' +
        'Persons with Persistent Multiple Barriers (PPMB), ' +
        'Persons with Disabilities (PWD).',
    ];

    this.getReferrals();
    this.getIdentities();
    this.getProvinces();
    this.getGenders();
    this.getPreferredcontactmethods();
    this.getCourtLevels();
    this.getCourtLocations();

    this.firstFormGroup = this._formBuilder.group({
      firstControl: ['', Validators.required],
      secondControl: ['', Validators.required],
    });
    this.secondFormGroup = this._formBuilder.group({
      firstName: ['', Validators.required],
      givenNames: [''],
      lastName: ['', Validators.required],
      birthdate: ['', Validators.required],
      address1: ['', Validators.required],
      city: ['', Validators.required],
      province: ['', Validators.required],
      postalCode: ['', Validators.required],
      phoneNumber: [''],
      email: ['', [Validators.required, Validators.email]], //Validators.pattern('^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+.[a-zA-Z0-9-.]+$')
      PreferredName: [''],
      saddress: [''],
      cellNumber: [''],
      workNumber: ['', Validators.required],
      gender: [''],
      identity: [''],
      otherIdentity: ['', Validators.maxLength(100)],
    });

    this.thirdFormGroup = this._formBuilder.group({
      firstName: ['', Validators.required],
      givenNames: [''],
      lastName: ['', Validators.required],
      pname: [''],
      birthdate: [''],
      saddress1: [''],
      saddress2: [''],
      city: [''],
      province: [''],
      postalCode: [''],
      homePhoneNumber: [''],
      cellPhoneNumber: [''],
      workPhoneNumber: [''],
      email: ['', Validators.email],
      gender: [''],
    });

    this.fourthFormGroup1 = this._formBuilder.group({
      users: this._formBuilder.array([
        this._formBuilder.group({
          firstName: ['', Validators.required],
          lastName: ['', Validators.required],
          birthdate: ['', Validators.required],
          childDependency: [],
          middleName: [],
        }),
      ]),
    });

    this.fifthFormGroup = this._formBuilder.group({
      orderDate: [''],
      courtLocation: [''],
      payorIncome: ['', Validators.pattern('^-?[0-9]\\d*(\\.\\d{1,2})?$')],
      recalculationOrdered: [],
      isSpecifiedIncome: [],
    });

    // setup default values
    this.fifthFormGroup.controls['recalculationOrdered'].patchValue('Yes');
    this.fifthFormGroup.controls['isSpecifiedIncome'].patchValue('Yes');

    this.sixFormGroup = this._formBuilder.group({
      childSafety: [''],
      childSafetyDescription: ['', Validators.maxLength(3000)],
      contactMethod: [''],
      enrollFMEP: [''],
      FMEPinput: [''],
      incomeAssistance: [''],
      referral: [''],
    });

    // setup default values
    this.sixFormGroup.controls['childSafety'].patchValue('No');
    this.sixFormGroup.controls['contactMethod'].patchValue('Email');
    this.sixFormGroup.controls['enrollFMEP'].patchValue('No');
    this.sixFormGroup.controls['incomeAssistance'].patchValue('Yes');

    this.seventhFormGroup = this._formBuilder.group({
      secondCtrl: [''],
    });
    this.eFormGroup = this._formBuilder.group({
      secondCtrl: [false, Validators.requiredTrue],
    });
    this.nineFormGroup = this._formBuilder.group({
      secondCtrl: ['', Validators.required],
    });
    //this.setFormDataFromLocal();
  }

  onIdentityChange(event: MatSelectChange) {
    let identity: LookupValue =
      this.identities.find((x) => x.id == event.value) ?? null;
    //this.logger.info('seleceted identity: ', identity.value);
    if (identity.value === 'Other') {
      this.isVisibleOtherIdentity = true;
    } else {
      this.isVisibleOtherIdentity = false;
    }
    //this.logger.info('isVisibleOtherIdentity: ', this.isVisibleOtherIdentity);
  }

  clearDate(event) {
    event.stopPropagation();
    this.dateOfOrder = null;
  }

  clearDateOtherParty(event) {
    event.stopPropagation();
    this.birthOfDateOtherParty = null;
  }

  setFormDataFromLocal() {
    if (localStorage.getItem('formData')) {
      let data = localStorage.getItem('formData');
      data = JSON.parse(data);
      if (data['firstStep']) {
        this.firstFormGroup.patchValue(data['firstStep']);
      }
      if (data['secondFormGroup']) {
        this.secondFormGroup.patchValue(data['secondFormGroup']);
      }
      if (data['thirdFormGroup']) {
        this.thirdFormGroup.patchValue(data['thirdFormGroup']);
      }
      if (data['fourthFormGroup']) {
        this.fourthFormGroup.patchValue(data['fourthFormGroup']);
      }
      if (data['fifthFormGroup']) {
        this.fifthFormGroup.patchValue(data['fifthFormGroup']);
      }
      if (data['sixFormGroup']) {
        this.sixFormGroup.patchValue(data['sixFormGroup']);
      }
      if (data['seventhFormGroup']) {
        this.seventhFormGroup.patchValue(data['seventhFormGroup']);
      }
      if (data['eFormGroup']) {
        this.eFormGroup.patchValue(data['eFormGroup']);
      }
      if (data['nineFormGroup']) {
        this.nineFormGroup.patchValue(data['nineFormGroup']);
      }
    }
  }

  onDateChange(event: MatDatepickerInputEvent<Date>, i: number): void {
    //var childYears = this.diff_years(event.value, new Date());
    var childYears = this.ageFromDateOfBirthday(event.value);
    //this.logger.info(`childYears = ${childYears}`);
    this.isHiddens[i] = childYears >= 19 ? true : false;
    //this.logger.warn(`childYears = ${childYears}, isHiddens[i] = ${this.isHiddens[i]}`);
  }

  diff_years(dt2, dt1) {
    var diff = (dt2.getTime() - dt1.getTime()) / 1000;
    diff /= 60 * 60 * 24;
    return Math.abs(Math.round(diff / 365.25));
  }

  ageFromDateOfBirthday(dateOfBirth: any): number {
    const today = new Date();
    //this.logger.warn(`today = ${today}`);
    const birthDate = new Date(dateOfBirth);
    //this.logger.warn(`birthDate = ${birthDate}`);
    let age = today.getFullYear() - birthDate.getFullYear();
    //this.logger.warn(`age = ${age}`);
    const m = today.getMonth() - birthDate.getMonth();
    //this.logger.warn(`m = ${m}`);

    if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
      age--;
    }
    //this.logger.warn(`age = ${age}`);

    return age;
  }

  forSubmitBtn(event) {
    //this.logger.info(`event: ${event}`);
    //this.logger.info(`event.checked: ${event.checked}`);
    this.isDisabledSubmit = !event.checked;
  }

  openDialog(inData) {
    const dialogRef = this.dialog.open(ConfirmDialogComponent, {
      width: '550px',
      data: inData,
    });

    dialogRef.afterClosed().subscribe((result) => {
      console.log(`Dialog result: ${result}`);
    });
  }

  editPage(stepper, index) {
    this.isEditable = true;
    stepper.selectedIndex = index;
  }
  getIdentities() {
    this.accountService.apiAccountIdentitiesGet().subscribe({
      next: (data) => {
        this.identities = data;
        //this.logger.info('this.identities',this.identities);
      },
      error: (e) => {
        //this.logger.error('error is getIdentities', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getProvinces() {
    this.accountService.apiAccountProvincesGet().subscribe({
      next: (data) => {
        this.provinces = data;
        //this.logger.info('this.provinces',this.provinces);
      },
      error: (e) => {
        //this.logger.error('error in getProvinces', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getGenders() {
    this.accountService.apiAccountGendersGet().subscribe({
      next: (data) => {
        this.genders = data;
        //this.logger.info('this.genders',this.genders);
      },
      error: (e) => {
        //this.logger.error('error is getGenders', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getCourtLocations() {
    this.lookupService.apiLookupCourtlocationsGet().subscribe({
      next: (data) => {
        this.courtLocations = data;
        //this.logger.info('this.courtLocations',this.courtLocations);
      },
      error: (e) => {
        //this.logger.error('error in getCourtLocations', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getCourtLevels() {
    this.lookupService.apiLookupCourtlevelsGet().subscribe({
      next: (data) => {
        this.courtLevels = data;
        //this.logger.info('this.courtLevels',this.courtLevels);
      },
      error: (e) => {
        //this.logger.error('error in getCourtLevels', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getReferrals() {
    this.accountService.apiAccountReferralsGet().subscribe({
      next: (data) => {
        this.referrals = data;
        //this.logger.info('this.referals',this.referrals);
      },
      error: (e) => {
        //this.logger.error('error in getReferrals', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  getPreferredcontactmethods() {
    this.accountService.apiAccountPreferredcontactmethodsGet().subscribe({
      next: (data) => {
        this.preferredContactMethods = data;
        //this.logger.info('this.preferredContactMethods',this.preferredContactMethods);
      },
      error: (e) => {
        //this.logger.error('error in getPreferredcontactmethods', e);
        this.data = {
          title: 'Error',
          content: e.message,
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }

  callAntherchild() {
    const usersArray = this.fourthFormGroup1.controls.users as FormArray;
    const arraylen = usersArray.length;

    const newUsergroup: FormGroup = this._formBuilder.group({
      firstName: ['', Validators.required],
      lastName: ['', Validators.required],
      birthdate: ['', Validators.required],
      //givenNames: [],
      childDependency: [],
      middleName: [],
    });

    usersArray.insert(arraylen, newUsergroup);
    this.isHiddens.push(false);

    this.isChildDelete = false;
  }

  deletechild1() {
    const arraylen = this.fourthFormGroup1.get('users')['controls'].length;
    (<FormArray>this.fourthFormGroup1.controls.users).removeAt(arraylen - 1);
    if (arraylen === 1) {
      this.isChildDelete = true;
    } else {
      this.isChildDelete = false;
    }
  }

  saveLater() {
    this.isDisabledSubmit = true;
    const formData = {
      firstStep: this.firstFormGroup.value,
      secondFormGroup: this.secondFormGroup.value,
      thirdFormGroup: this.thirdFormGroup.value,
      fourthFormGroup1: this.fourthFormGroup1.value,
      fifthFormGroup: this.fifthFormGroup.value,
      sixFormGroup: this.sixFormGroup.value,
      seventhFormGroup: this.seventhFormGroup.value,
      eFormGroup: this.eFormGroup.value,
      nineFormGroup: this.nineFormGroup.value,
    };

    this.prepareData();
    this.isDisabledSubmit = false;
  }
  save() {
    this.prepareData();
    //localStorage.getsetItemItem('formData', '');
  }

  //To edit the content of SERVICE PORTAL TERMS OF USE, Change the word document and convert to html and paste the content here
  showInfoCollectionDisclosure(): void {
    const dialogRef = this.dialog.open(ModalDialogHtmlComponent, {
      width: '850px',
      data: '',
    });

    dialogRef.afterClosed().subscribe((result) => {
      //this.logger.info(`Dialog result: ${result}`);
    });
  }

  //To edit the content of “INFORMATION COLLECTION, USE AND DISCLOSURE”, Change the word document and convert to html and paste the content here
  showTermsOfUse(): void {
    const dialogRef = this.dialog.open(ModalDialogHtmlComponent, {
      width: '750px',
      data: '',
    });

    dialogRef.afterClosed().subscribe((result) => {
      //this.logger.info(`Dialog result: ${result}`);
    });
  }

  openModalDialog(): void {
    const dialogRef = this.dialog.open(ModalDialogComponent, {
      width: '450px',
      data: this.data,
    });

    dialogRef.afterClosed().subscribe((result) => {
      //this.logger.info(`Dialog result: ${result}`);
    });
  }

  getProvinceById(id) {
    let province: LookupValue = this.provinces.find((x) => x.id == id) ?? null;
    return province != null ? province.value : '-';
  }

  getGenderById(id) {
    let gender: LookupValue = this.genders.find((x) => x.id == id) ?? null;
    return gender != null ? gender.value : '-';
  }

  getIdentityById(id) {
    let identity: LookupValue = this.identities.find((x) => x.id == id) ?? null;

    if (identity.value === 'Other') {
      return 'Other: ' + this.secondFormGroup.value.otherIdentity;
    }

    return identity != null ? identity.value : '-';
  }

  getReferralById(id) {
    let referral: LookupValue = this.referrals.find((x) => x.id == id) ?? null;
    return referral != null ? referral.value : '-';
  }

  getCourtLocationById(id) {
    let courtLocation: LookupValue =
      this.courtLocations.find((x) => x.id == id) ?? null;
    return courtLocation != null ? courtLocation.value : '-';
  }

  transformDate(date) {
    return this.datePipe.transform(date, 'yyyy-MM-dd');
  }

  findId(value) {
    if (value == 'Yes') {
      return this._yes;
    } else if (value == 'No') {
      return this._no;
    }
    return this._iDontKnow;
  }

  getCourtTyleFile(value) {
    const courtTypeFileId =
      value == 'Order' ? this._courtOrder : this._writtenAgreement;
    const inCourtFileType: LookupValue = { id: courtTypeFileId, value: value };
    //this.logger.warn(`inCourtFileType: ${inCourtFileType}`);
    return inCourtFileType;
  }

  prepareData() {
    //---------
    /*const users = this.fourthFormGroup1.value.users;
    let childs: Array<Child> = new Array<Child>();

    for(let user in users) {
      let child: Child = {
        firstName: user.firstName,
        middleName: user.middleName,
        lastName: user.lastName,
        dateOfBirth: this.transformDate(user.birthdate),
        childIsDependent: user.childDependency
      };
      childs.push(child)
    }*/

    //--------
    // --- populate children
    const users = this.fourthFormGroup1.value.users;
    this.logger.info('users: ', users);
    let childs: Array<Child> = new Array<Child>();
    for (var i = 0; i < users.length; i++) {
      let child: Child = {
        firstName: users[i].firstName,
        middleName: users[i].middleName,
        lastName: users[i].lastName,
        dateOfBirth: this.transformDate(users[i].birthdate),
        childIsDependent: users[i].childDependency,
      };
      childs.push(child);
    }

    // --- populate partyRole
    const roleData = this.firstFormGroup.value;
    let partyRole: PartyRole = PartyRole.Unknown;
    let partyEnrolled = '';

    if (roleData.firstControl == PartyRole.Recipient) {
      partyRole = PartyRole.Recipient;
      partyEnrolled = 'Recipient';
    } else {
      partyRole = PartyRole.Payor;
      partyEnrolled = 'Payor';
    }

    // --- populate party
    const partyData = this.secondFormGroup.value;
    const file1Data = this.fifthFormGroup.value;
    const file2Data = this.sixFormGroup.value;

    //let LookupValue
    let inGender: LookupValue =
      this.genders.find((x) => x.id == partyData.gender) ?? null;
    let inProvince: LookupValue =
      this.provinces.find((x) => x.id == partyData.province) ?? null;
    let inIdentityParty: LookupValue =
      this.identities.find((x) => x.id == partyData.identity) ?? null;
    let inReferral: LookupValue =
      this.referrals.find((x) => x.id == file2Data.referral) ?? null;
    let inPreferredContactMethod: LookupValue =
      this.preferredContactMethods.find(
        (x) => x.value == file2Data.contactMethod,
      ) ?? null;

    let inParty: Party = {
      partyId: '00000000-0000-0000-0000-000000000000',
      firstName: partyData.firstName,
      middleName: partyData.givenNames,
      lastName: partyData.lastName,
      preferredName: partyData.PreferredName,
      dateOfBirth: this.transformDate(partyData.birthdate),
      gender: inGender,
      addressStreet1: partyData.address1,
      addressStreet2: partyData.saddress,
      city: partyData.city,
      province: inProvince,
      postalCode: partyData.postalCode,
      homePhone: partyData.phoneNumber,
      workPhone: partyData.workNumber,
      cellPhone: partyData.cellNumber,
      email: partyData.email,
      optOutElectronicDocuments: null, // ??? may need to remove?
      identity: inIdentityParty,
      referral: inReferral,
      preferredContactMethod: inPreferredContactMethod,
      incomeAssistance: this.findId(file2Data.incomeAssistance),
      otherIdentity: partyData.otherIdentity,
    };

    // --- populate other party
    const otherPartyData = this.thirdFormGroup.value;
    let inOtherGender: LookupValue =
      this.genders.find((x) => x.id == otherPartyData.gender) ?? null;
    let inOtherProvince: LookupValue =
      this.provinces.find((x) => x.id == otherPartyData.province) ?? null;

    let inOtherParty: Party = {
      partyId: '00000000-0000-0000-0000-000000000000',
      firstName: otherPartyData.firstName,
      middleName: otherPartyData.givenNames,
      lastName: otherPartyData.lastName,
      preferredName: otherPartyData.pname,
      dateOfBirth: this.transformDate(otherPartyData.birthdate),
      gender: inOtherGender,
      addressStreet1: otherPartyData.saddress1,
      addressStreet2: otherPartyData.saddress2,
      city: otherPartyData.city,
      province: inOtherProvince,
      postalCode: otherPartyData.postalCode,
      homePhone: otherPartyData.homePhoneNumber,
      workPhone: otherPartyData.workPhoneNumber,
      cellPhone: otherPartyData.cellPhoneNumber,
      email: otherPartyData.email,
      optOutElectronicDocuments: null,
      identity: null,
      referral: null,
      preferredContactMethod: null,
      referenceNumber: null,
    };

    // --- populate file
    let inCourtFileType: LookupValue = this.getCourtTyleFile(
      roleData.secondControl,
    );

    let inBcCourtLocation: LookupValue =
      this.courtLocations.find((x) => x.id == file1Data.courtLocation) ?? null;

    //let inBcCourtLevel: CourtLookupValue = this.courtLevels[0];
    let inBcCourtLevel: LookupValue =
      this.courtLevels.find((x) => x.value == 'Provincial') ?? null;

    let inFile: any = {
      status: FileStatus.Unknown,
      usersRole: partyRole,
      fileId: '0',
      fileNumber: null,
      partyEnrolled: partyEnrolled,
      courtFileType: inCourtFileType,
      bcCourtLevel: inBcCourtLevel,
      bcCourtLocation: inBcCourtLocation,
      dateOfOrderOrWA: this.transformDate(file1Data.orderDate),
      incomeOnOrder: file1Data.payorIncome,
      section7Expenses: file1Data.isSpecifiedIncome,
      safetyAlertRecipient: null,
      recipientSafetyConcernDescription: null,
      safetyAlertPayor: null,
      payorSafetyConcernDescription: null,
      isFMEPFileActive: file2Data.enrollFMEP,
      fmepFileNumber: file2Data.FMEPinput,
      recalculationOrderByCourt: file1Data.recalculationOrdered,
      otherParty: inOtherParty,
      children: childs,
    };

    // --- populate
    var newFileRequest: NewFileRequest = {
      user: inParty,
      file: inFile,
    };

    if (partyEnrolled == 'Recipient') {
      newFileRequest.file.safetyAlertRecipient = file2Data.childSafety;
      newFileRequest.file.recipientSafetyConcernDescription =
        file2Data.childSafetyDescription;
    } else {
      newFileRequest.file.safetyAlertPayor = file2Data.childSafety;
      newFileRequest.file.payorSafetyConcernDescription =
        file2Data.childSafetyDescription;
    }

    //this.logger.info("newFileRequest:", newFileRequest);

    this.accountService.apiAccountCreatePost(newFileRequest).subscribe({
      next: (outData: any) => {
        var partyId = outData.partyId;
        var fileId = outData.fileId;
        var fileNumber = outData.fileNumber;

        //this.logger.info("partyId", partyId);
        //this.logger.info("fileId", fileId);
        //this.logger.info("fileNumber", fileNumber);

        let customOptions: DialogOptions = { data: { fileNumber: fileNumber } };
        this.openDialog(customOptions);
        this.router.navigate(['/communication'], {
          queryParams: { index: 1, fileNumber: fileNumber },
        });
      },
      error: (e) => {
        //this.logger.error('error in prepareData', e);
        this.data = {
          title: 'Error',
          content:
            'The information you entered is not valid. Please enter the information given to you by the Child Support Recalculation Service.',
          content_normal: 'If you continue to have problems, contact us at ',
          content_link: '1-866-660-2684',
          weight: 'normal',
          color: 'red',
        };
        this.openModalDialog();
      },
    });
  }
}
