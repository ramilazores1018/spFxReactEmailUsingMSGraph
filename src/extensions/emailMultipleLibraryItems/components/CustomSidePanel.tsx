import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';    
import {IMailMessage,IRecipient,IAttachment}  from './IMailMessage'; 
import {
  Label,
  TextField,
  DefaultButton,
  PrimaryButton,
  Button,
  autobind,
  Panel,
  PanelType,
  Spinner,
  SpinnerType,
  DialogFooter
} from 'office-ui-fabric-react';

import { sp } from '@pnp/sp';

import {
  RowAccessor
} from '@microsoft/sp-listview-extensibility';

import { SPListItem } from '@microsoft/sp-page-context';
import { IListItem } from './IListItem';

export interface ICustomSidePanelState {
  sending: boolean;
  listItem:any;
  appStatus:string;
  emailTemplateTo:string,
  emailTemplateBody:string;
  emailTemplateSubject:string;
  emailTemplateCC:string;
  emailTemplateBcc:string;
  emailTemplateSupplierName:string;
  msgColor:any;
  buttonDisabled:boolean;
  arrayOfAttachments:any;
}

export interface ICustomSidePanelProps {
  onClose: () => void;
  isOpen: boolean;
  emailTo: string;
  emailAttachments: any[];
  emailToName;
  libraryItems: any[];
  msGraphClientFactory: any;
  listName:  any;
  spHttpClient:  any;
  siteUrl: any;
}

export default class CustomSidePanel extends React.Component<
ICustomSidePanelProps, 
ICustomSidePanelState>  {

  constructor(props: ICustomSidePanelProps) {
    super(props);

    this.state = { 
      sending: false,
      listItem:[],
      appStatus:'',
      emailTemplateTo:this.props.emailTo,
      emailTemplateSubject:'',
      emailTemplateCC:'',
      emailTemplateBcc:'',
      emailTemplateBody:'',
      msgColor:{color:'black'},
      emailTemplateSupplierName:'',
      buttonDisabled:false,
      arrayOfAttachments:this.props.libraryItems != null ? this.props.libraryItems.map(x => x["fileLeafRef"]) : []

     };

  }

  @autobind
  private _onEmailToChanged(uiTextEmailTo: string) {
    this.setState({
      emailTemplateTo:uiTextEmailTo,
    });
  }

  @autobind
  private _onEmailCcChanged(uiTextEmailCc: string) {

    this.setState({
      emailTemplateCC:uiTextEmailCc,
    });
  }


  @autobind
  private _onEmailSubjectChanged(uiTextEmailSubject: string) {
    this.setState({
     emailTemplateSubject:uiTextEmailSubject
    });
  }

  @autobind
  private _onEmailBodyChanged(uiTextEmailBody: string) {
    this.setState({
      emailTemplateBody:uiTextEmailBody
    });
  }

  @autobind
  private _onCancel() {
    
    this.setState({
      appStatus:'',
      msgColor:{color:'black'},
      buttonDisabled:false,
      emailTemplateBody:''
    });

    this.props.onClose();
  
  }

  @autobind
  private async _onSend() {

      this.setState({
           sending: true,
           appStatus:'Start sending email...',
           msgColor:{color:'black'},
           buttonDisabled:true
      });


      let emailProperties = this.props;
      let arrayOfEmailAttachments = emailProperties.emailAttachments;
      let arryOfLibraryItems = emailProperties.libraryItems;

      let toRecipientList: Array<IRecipient> = new Array<IRecipient>();
      let ccRecipientsList: Array<IRecipient> = new Array<IRecipient>();
      let fileAttachmentsList = [];
      let totalBytes:number = 0;
    
      arryOfLibraryItems.map((item: any) => {
 
        totalBytes += parseInt(item.fileSize);

      });

      totalBytes = totalBytes/1048576;

      if (totalBytes > 2.9){
        
        // Reserve if incase needed in the prompt
        // Math.round(totalBytes) +" MB.");
        //this.setState({ sending: false });
        //this.props.onClose();

        this.setState({
          appStatus:'Sorry, Microsoft Online does not allow attaching files \nlarger than 2.9Mb to emails being sent on your behalf.',
          msgColor:{color:'red'},
          buttonDisabled:false
        });
    

        return;
      }


    this.setState({
      appStatus:'Attaching file(s) to email,please wait..',
      msgColor:{color:'black'}
    });

    this.loadFiles(arryOfLibraryItems).then(result => {

      let arrayOfFilesToConvert: any = [];

      for (let i = 0; i < result.length; i++) {
        arrayOfFilesToConvert.push(this.convertFile(result[i], arryOfLibraryItems));
      }

      this.convertAllFiles(arrayOfFilesToConvert).then(convertedFiles => {

        for (let x = 0; x < convertedFiles.length; x++) {
          arrayOfEmailAttachments.push(this.assignFileAttachment(convertedFiles[x]));
        }

        this.setState({
          appStatus:'Sending email...',
          msgColor:{color:'black'}
        });
    

        if (this.state.emailTemplateTo .indexOf(',') > -1){
          this.state.emailTemplateTo.split(',').map((item: any) => {
            toRecipientList.push({
              emailAddress: {address:item.replace(/\s/g, "")}
            });
          });
        }else {
          toRecipientList.push({
            emailAddress: {address:this.state.emailTemplateTo.replace(/\s/g, "")}
          });
        }

        if (this.state.emailTemplateCC.indexOf(',') > -1){
          this.state.emailTemplateCC.split(',').map((item: any) => {
            ccRecipientsList.push({
              emailAddress: {address:item.replace(/\s/g, "")}
            });
          });
        }else {
          ccRecipientsList.push({
            emailAddress: {address:this.state.emailTemplateCC.replace(/\s/g, "")}
          });
        }
      
        emailProperties.emailAttachments.map((item: any) => {
          fileAttachmentsList.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": item.Name,
            "contentBytes": item.Content.ContentData,
            "isInline": "false"
          });
        });


  
            const mailData :IMailMessage = {
              message: {
                  toRecipients: toRecipientList,
                  ccRecipients: ccRecipientsList,
                  subject: this.state.emailTemplateSubject
                  ,
                  body: {
                      contentType:"Text",
                      content:this.state.emailTemplateBody==='' ? this.buildEmailBodyContent(this.state.listItem,this.state.arrayOfAttachments.join(', ')):this.state.emailTemplateBody,
                  },
                  attachments:fileAttachmentsList,

              },
              saveToSentItems: true

            };
 
              this.props.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
                client
                .api('/me/sendMail')
                .post(mailData, (err, res) => {
                
                  if (err){

                    console.log(err); 

                    this.setState({
                      appStatus:'Error in sending email. Please contact the site admin.',
                      msgColor:{color:'red'},
                      buttonDisabled:false
                    });

                   return;
                  }

                    this.setState({ 
                      sending: false,
                      appStatus:'',
                      msgColor:{color:'black'},
                      buttonDisabled:false
                     });

                     alert("Email Sent.");

                    this.props.onClose();

                });
              
              });

      });

    }).catch((errorDetails) => {

      console.log(errorDetails);

      this.setState({
        appStatus:'Error in sending email. Please contact the site admin.',
        msgColor:{color:'red'},
        buttonDisabled:false
      });

    });

  }

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&
      $select=Id,`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }

  private assignFileAttachment(fileObj: any) {
    return { "Content": { "ContentData": fileObj.base64File, "ContentTransferEncoding": "Base64" }, "Name": fileObj.fileName };
  }


  private convertFile(refFile: any, arryOfLibraryItems_: any): Promise<any> {
    return new Promise<any>(resolve => {

      var reader = new FileReader();
      reader.onload = () => {
        let result: any = reader.result;
        let data64result = btoa(result);

        resolve({ "fileName": refFile.fileRef.substring(refFile.fileRef.lastIndexOf('/') + 1), "base64File": data64result });

        arryOfLibraryItems_
          .filter(obj => obj.fileRef === refFile.fileRef)[0]["base64String"] = data64result;

      };
      reader.readAsBinaryString(refFile.fileBlob);
    });

  }

  private getFileLocation(value_FileRef_: string): Promise<any> {
    return new Promise<any>(resolve => {

      sp.web.getFileByServerRelativeUrl(value_FileRef_)
        .getBlob()
        .then((blob: Blob) => {
          resolve({ "fileRef": value_FileRef_, "fileBlob": blob });
        }).catch((errorDetails) => {
          resolve("ERROR : " + errorDetails);
        });

    });

  }

  private async convertAllFiles(arrayOfFilesToConvert_: any) {

    let resolvedFinalArray = await Promise.all(arrayOfFilesToConvert_);
    return resolvedFinalArray;

  }

  private async loadFiles(arryOfLibraryItems) {

    let arrayOfAttachments = arryOfLibraryItems.map(x => x["fileRef"]);
    let arrayOfPromises: any = [];

    for (let i = 0; i < arrayOfAttachments.length; i++) {
      arrayOfPromises.push(this.getFileLocation(arrayOfAttachments[i]));
    }

    let resolvedFinalArray = await Promise.all(arrayOfPromises);
    return resolvedFinalArray;

  }


  private async sendEmail(emailProperties: any) {
    return;
  }

  public listData: Array<Object>;

  private buildEmailBodyContent(emailTemplateItem:any,listofAttachments:string) {

    var results: string;
    var supplierName:string;
    var emailSignatureAddress:string;
    var emailSignatureWebsite:string;

    var defaFaultBody:string =  `Supplier @@SupplierName,
  
    Please find attached the following remittance advice/s:
    
    @@Attachments
    
    Should you have any questions, please do not hesitate to contact us on the numbers below to discuss.
    
    @@EmailSignatureAddress
    @@EmailSignatureWebsite`;

    let propsItems = this.props.libraryItems;
      
    if (propsItems.length){
      for (let i = 0; i < propsItems.length; i++) {
        if (propsItems[i].supplierName !==''){
         supplierName = propsItems[i].supplierName;
         break;
        }
 
     }
    }
   

    var supplierName_:string = (supplierName) ? supplierName : emailTemplateItem.SupplierName;
    var emailSignatureAddress_:string = (emailTemplateItem.EmailSignatureAddress) ? emailTemplateItem.EmailSignatureAddress:'Rinnai Australia - 100 Atlantic Drive Keysborough VIC 3173';
    var emailSignatureWebsite_:string = (emailTemplateItem.EmailSignatureWebsite) ? emailTemplateItem.EmailSignatureWebsite:'rinnai.com.au';
        
    
    if (emailTemplateItem.EmailBodyTemplate){

        var settingsTemplate:string = emailTemplateItem.EmailBodyTemplate;
       
        results = emailTemplateItem.EmailBodyTemplate
                  .replace(new RegExp('@@SupplierName', 'g'),supplierName_)
                  .replace(new RegExp('@@EmailSignatureAddress', 'g'),emailSignatureAddress_)
                  .replace(new RegExp('@@EmailSignatureWebsite', 'g'),emailSignatureWebsite_)
                  .replace(new RegExp('@@Attachments', 'g'),listofAttachments);
     
    }else{

      results = defaFaultBody
      .replace(new RegExp('@@SupplierName', 'g'),supplierName_)
      .replace(new RegExp('@@EmailSignatureAddress', 'g'),emailSignatureAddress_)
      .replace(new RegExp('@@EmailSignatureWebsite', 'g'),emailSignatureWebsite_)
      .replace(new RegExp('@@Attachments', 'g'),listofAttachments);

    }

    return results;

  }

  componentDidMount(){

    this.getLatestItemId()
    .then((itemId: number): Promise<SPHttpClientResponse> => {
      if (itemId === -1) {
        throw new Error('No items found in the list');
      }


      return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemId})?$select=
      Title,
      EmailBodyTemplate,
      SupplierName,
      EmailCC,
      EmailSubject,
      EmailSignatureAddress,
      EmailSignatureWebsite`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
    })
    .then((response: SPHttpClientResponse): Promise<IListItem> => {
      return response.json();
    })
    .then((item: IListItem): void => {

      this.setState({
        listItem:item,
        emailTemplateSubject:item.EmailSubject,
        emailTemplateCC:item.EmailCC,
        emailTemplateSupplierName :''
      });

    }, (error: any): void => {
      

    });


  }



  public render(): React.ReactElement<ICustomSidePanelProps> {

    let { isOpen} = this.props;
      
      const{
        emailTemplateTo,
        emailTemplateSubject,
        emailTemplateCC,
        emailTemplateBcc,
        emailTemplateBody,
        listItem,
        arrayOfAttachments
      } = this.state;


    const attachmentItemsAsHtml = arrayOfAttachments.map(function (item) {
      return <li>{item}</li>;
    });

  
    const bodyEmail = this.state.emailTemplateBody==='' ? this.buildEmailBodyContent(this.state.listItem,this.state.arrayOfAttachments.join(', ')):this.state.emailTemplateBody;

    return <Panel
          isLightDismiss={false}
          isOpen={isOpen}
          type={PanelType.smallFixedFar}
       >
     
          <TextField label="To:" underlined required placeholder="Enter an email address" value={emailTemplateTo} onChanged={this._onEmailToChanged} ></TextField>
          <TextField label="Cc:" underlined placeholder="Enter an email address" value={emailTemplateCC} onChanged={this._onEmailCcChanged}></TextField>
          <TextField label="Subject:" underlined required placeholder="Remittance Advice - Rinnai Australia" value={emailTemplateSubject} onChanged={this._onEmailSubjectChanged}></TextField>
          <TextField label="" multiline rows={14} value={bodyEmail} onChanged={this._onEmailBodyChanged}></TextField>
      <div>
        <ul>
          {attachmentItemsAsHtml}
        </ul>
      </div>
      <div>
        <span style={this.state.msgColor}>
          {this.state.appStatus}
        </span>
       </div>

      <DialogFooter>
        <DefaultButton disabled={this.state.buttonDisabled} text="Cancel" onClick={this._onCancel}></DefaultButton>
        <PrimaryButton disabled={this.state.buttonDisabled} text="Send" onClick={this._onSend}></PrimaryButton>
      </DialogFooter>

    </Panel>;

  }

}