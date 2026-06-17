import { HttpClient } from '@angular/common/http';
import { Component, Inject, OnInit } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { LoggerService } from '@core/services/logger.service';
import { ConfirmDialogComponent } from '@shared/dialogs/confirm-dialog/confirm-dialog.component';
import { AccountService } from 'app/api/api/account.service';
import { LookupService } from 'app/api/api/lookup.service';

import { MatSelectChange } from '@angular/material/select';
import { ActivatedRoute, Router } from '@angular/router';
import { ModalDialogComponent } from 'app/components/modal-dialog/modal-dialog.component';

// -- import data structure
import {
  CSRSAccountFile,
  CSRSAccountRequest,
  LookupValue,
  Party,
} from 'app/api/model/models';

import { DatePipe } from '@angular/common';
import { ModalDialogHtmlComponent } from '@components/modal-dialog-htmlcontent/modal-dialog-htmlcontent.component';
@Component({
  selector: 'app-application-form-stepper',
  templateUrl: './application-form-stepper.component.html',
  styleUrls: ['./application-form-stepper.component.scss'],
})
export class ApplicationFormStepperComponent implements OnInit {
  secondFormGroup: FormGroup;
  sixFormGroup: FormGroup;
  eFormGroup: FormGroup;
  nineFormGroup: FormGroup;

  provinces: any = [];
  genders: any = [];
  identities: any = [];
  preferredContactMethods: any = [];

  today = new Date();
  isEditable = false;
  isDisabledSubmit: boolean = true;

  _yes: number = 867670000;
  _no: number = 867670001;
  _iDontKnow: number = 867670002;

  data: any = null;

  partyId: any = '';
  fileId: any = '';

  errorMessage: any = '';
  errorMailMessage: any = '';
  errorIncomeMessage: any = '';
  errorDateMessage: any = '';
  errorMaxMessage: any = '';
  errorMaxOtherIdentityMessage: any = '';
  isVisibleOtherIdentity: boolean = false;

  constructor(
    private _formBuilder: FormBuilder,
    private http: HttpClient,
    @Inject(AccountService) private accountService,
    @Inject(LookupService) private lookupService,
    @Inject(LoggerService) private logger,
    @Inject(Router) private router,
    public dialog: MatDialog,
    private datePipe: DatePipe,
    private route: ActivatedRoute,
  ) {}

  ngOnInit() {
    this.route.queryParams.subscribe((params) => {
      this.partyId = params.partyId;
      this.fileId = params.fileId;
    });

    this.provinces = [{ id: '123', value: 'British Columbia' }];
    this.identities = [{ id: '123', value: 'Native' }];
    this.genders = [{ id: '123', value: 'Male' }];
    this.preferredContactMethods = [{ id: '123', value: 'Email' }];

    this.errorMessage = 'Error: Field is required. ';
    this.errorMailMessage = 'Email address without @ or domain name. ';
    this.errorIncomeMessage = 'Field should have numerical values. ';
    this.errorDateMessage = 'Date cannot be in future.';
    this.errorMaxMessage = 'You can only enter up to 3000 characters.';
    this.errorMaxOtherIdentityMessage =
      'You can only enter up to 100 characters.';

    this.getIdentities();
    this.getProvinces();
    this.getGenders();
    this.getPreferredcontactmethods();

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
      email: ['', [Validators.required, Validators.email]],
      PreferredName: [''],
      saddress: [''],
      cellNumber: [''],
      workNumber: ['', Validators.required],
      gender: [''],
      identity: [''],
      otherIdentity: ['', Validators.maxLength(100)],
    });

    this.sixFormGroup = this._formBuilder.group({
      childSafety: [''],
      childSafetyDescription: ['', Validators.maxLength(3000)],
      contactMethod: [''],
      incomeAssistance: [''],
    });

    // setup default values
    this.sixFormGroup.controls['childSafety'].patchValue('No');
    this.sixFormGroup.controls['contactMethod'].patchValue('Email');
    this.sixFormGroup.controls['incomeAssistance'].patchValue('Yes');

    this.eFormGroup = this._formBuilder.group({
      secondCtrl: [false, Validators.requiredTrue],
    });
    this.nineFormGroup = this._formBuilder.group({});

    //this.setFormDataFromLocal();
    this.data = {
      //type: 'error',
      title: 'Technical error',
      weight: 'normal',
      color: 'red',
    };
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

  forSubmitBtn(event) {
    this.isDisabledSubmit = !event.checked;
  }

  setFormDataFromLocal() {
    if (localStorage.getItem('formData')) {
      let data = localStorage.getItem('formData');
      data = JSON.parse(data);
      if (data['secondFormGroup']) {
        this.secondFormGroup.patchValue(data['secondFormGroup']);
      }
      if (data['sixFormGroup']) {
        this.sixFormGroup.patchValue(data['sixFormGroup']);
      }
    }
  }

  openDialog(inData) {
    const dialogRef = this.dialog.open(ConfirmDialogComponent, {
      width: '550px',
      data: inData,
    });

    dialogRef.afterClosed().subscribe((result) => {
      //this.logger.log(`Dialog result: ${result}`);
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

  getPreferredcontactmethods() {
    this.accountService.apiAccountPreferredcontactmethodsGet().subscribe({
      next: (data) => {
        this.preferredContactMethods = data;
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

  saveLater() {
    const formData = {
      secondFormGroup: this.secondFormGroup.value,
      sixFormGroup: this.sixFormGroup.value,
    };

    //this.logger.info("formData", formData);
    this.prepareData();
    //localStorage.setItem('formData', JSON.stringify(formData));
  }

  save() {
    localStorage.getsetItemItem('formData', '');
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

  showInfoCollectionDisclosure(): void {
    const dialogRef = this.dialog.open(ModalDialogHtmlComponent, {
      width: '850px',
      data: '',
    });
  }

  showTermsOfUse(): void {
    const dialogRef = this.dialog.open(ModalDialogHtmlComponent, {
      width: '750px',
      data: '',
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

  prepareData() {
    // --- populate party
    const partyData = this.secondFormGroup.value;
    const file2Data = this.sixFormGroup.value;

    //let LookupValue
    let inGender: LookupValue =
      this.genders.find((x) => x.id == partyData.gender) ?? null;
    let inProvince: LookupValue =
      this.provinces.find((x) => x.id == partyData.province) ?? null;
    let inIdentityParty: LookupValue =
      this.identities.find((x) => x.id == partyData.identity) ?? null;
    let inPreferredContactMethod: LookupValue =
      this.preferredContactMethods.find(
        (x) => x.value == file2Data.contactMethod,
      ) ?? null;

    let inParty: Party = {
      partyId: this.partyId,
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
      referral: null,
      preferredContactMethod: inPreferredContactMethod,
      incomeAssistance: this.findId(file2Data.incomeAssistance),
      otherIdentity: partyData.otherIdentity,
    };

    // --- populate file

    let inFile: CSRSAccountFile = {
      fileId: this.fileId,
      fileNumber: null,
      safetyAlertRecipient: file2Data.childSafety,
      recipientSafetyConcernDescription: file2Data.childSafetyDescription,
      safetyAlertPayor: file2Data.childSafety,
      payorSafetyConcernDescription: file2Data.childSafetyDescription,
    };

    // --- populate csrsAccountRequest
    let csrsAccountRequest: CSRSAccountRequest = {
      user: inParty,
      csrsAccountFile: inFile,
    };

    //this.logger.info("csrsAccountRequest:", csrsAccountRequest);

    this.accountService
      .apiAccountUpdatecsrsaccountPost(csrsAccountRequest)
      .subscribe({
        next: (outData: any) => {
          var partyId = outData.partyId;
          var fileId = outData.fileId;

          //this.logger.info("partyId", partyId);
          //this.logger.info("fileId", fileId);

          if (partyId == this.partyId && fileId == this.fileId) {
            //this.logger.info("inside partyId == this.partyId &&  fileId == this.fileId");

            this.data = {
              type: 'check',
              title: ' Account setup complete',
              content: 'Your account setup request has been submitted',
              content_normal: null,
              content_link: null,
              weight: 'normal',
              color: 'green',
            };
            this.openModalDialog();

            //this.logger.info("redirect to Communication");
            this.router.routeReuseStrategy.shouldReuseRoute = () => false;
            this.router.navigate(['/communication']);
          }
        },
        error: (e) => {
          //this.logger.error('error in prepareData', e);
          this.data = {
            //type: 'error',
            title: 'Error.',
            content:
              'The information you entered is not valid. Please enter the information given to you by the Child Support Recalculation Service.',
            content_normal: 'If you continue to have problems, contact us at ',
            content_link: '1-866-660-2684.',
            weight: 'normal',
            color: 'red',
          };
          this.openModalDialog();
        },
      });
  }
}
