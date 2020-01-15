import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  RowAccessor,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import CustomSidePanel, {ICustomSidePanelProps} from './components/CustomSidePanel';
import { sp } from '@pnp/sp';
import {autobind, assign} from '@uifabric/utilities';
import * as React from 'react';	
import * as ReactDom from 'react-dom';

import * as strings from 'EmailMultipleLibraryItemsCommandSetStrings';
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';  

/**
 * Ramil: Notes to self
 * Convert office UI icons into bit 64 icons  https://codepen.io/joshmcrty/pen/GOBWeV
 */

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEmailMultipleLibraryItemsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'EmailMultipleLibraryItemsCommandSet';

export default class EmailMultipleLibraryItemsCommandSet extends BaseListViewCommandSet<IEmailMultipleLibraryItemsCommandSetProperties> {

  private panelPlaceHolder: HTMLDivElement = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized EmailMultipleLibraryItemsCommandSet');
   
    //Setup Pnp JS with SPFx context
    sp.setup({
    spfxContext:this.context

    });

      //Create the container for our React component
      this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
   
    return Promise.resolve();
  }

  private _showPanel(emailTo: string
    , emailAttachments: any[]
    , emailToName: string
    , libraryItems: any[]) {

    this._renderPanelComponent({
    isOpen: true,
    emailTo,
    emailAttachments,
    emailToName,
    libraryItems,
    onClose: this._dismissPanel
    });

}

@autobind
private _dismissPanel() {
this._renderPanelComponent({isOpen:false});
}

private _renderPanelComponent(props: any){
const element: React.ReactElement<ICustomSidePanelProps> = React.createElement(CustomSidePanel, assign({
onClose: null,
emailTo: null,
emailAttachments: null,
libraryItems: null,
msGraphClientFactory: this.context.msGraphClientFactory,
listName: 'Settings',
spHttpClient: this.context.spHttpClient,
siteUrl: this.context.pageContext.web.absoluteUrl
},props));
ReactDom.render(element, this.panelPlaceHolder);

}

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      //  This command should be hidden unless one or more rows are selected; use === for exactly 1 row.
      compareOneCommand.visible = event.selectedRows.length >= 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':

           let libraryItems = [];

          if (event.selectedRows.length > 0 || event.selectedRows.length === 1) {

            event.selectedRows.forEach((row: RowAccessor, index: number) => {
              libraryItems.push({"id":row.getValueByName('ID')
                            , "fileLeafRef": row.getValueByName('FileLeafRef')
                            , "fileRef": row.getValueByName('FileRef')
                            , "fileSize": row.getValueByName('File_x0020_Size')
                            , "fileType": row.getValueByName('File_x0020_Type')
                            , "title": row.getValueByName('Title')
                            , "supplierId": row.getValueByName('SupplierId')
                            , "supplierName":row.getValueByName('SupplierName')
                            , "supplierAddress":row.getValueByName('SupplierAddress')
                            , "supplierPhone":row.getValueByName('SupplierPhone')
                            , "supplierEmail":row.getValueByName('SupplierEmail')
                            , "supplierFax":row.getValueByName('SupplierFax')
                            , "supplierAddressNo":row.getValueByName('SupplierAddressNo')
                            , "adviceDate":row.getValueByName('AdviceDate')
                            , "adviceNo": row.getValueByName('AdviceNo')
                            , "totalAmmount": row.getValueByName('TotalAmmount')
                            , "base64String": ""
                          });
        
            });

            let emailTo =  libraryItems.map( x => x["supplierEmail"])
                            .filter(function(el){
                                return ((el != null) && (el != ""));
                            })
                            .join(', ') ;

      
            let emailAttachments = [];

            let emailToName = libraryItems.map( x => x["supplierName"])
                              .filter(function(el){
                                return  ((el != null) && (el != ""));
                              })
                              .join(', ') ;
            
            this._showPanel(
                  emailTo
                , emailAttachments
                , emailToName
                , libraryItems
            );

        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }


}
