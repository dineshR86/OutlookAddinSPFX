import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse,MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPUser } from "@microsoft/sp-page-context";

export class AddinService {
  private _spclient: SPHttpClient;
  private _graphclient:MSGraphClient;
  public _currentuser: SPUser;
  public _weburl: string = 'https://cloudmission.sharepoint.com/sites/xRMLite/';
  public _casefolderrelativepath: string = '/sites/xRMLite/CaseFiles';
  public _caseid: string = 'b5fd8cf2-1277-4daa-9196-98a1b6d32401';
  public _statusid: string = '17216dcb-35ca-4675-8838-818ac63fdc30';
  public _caselibraryid: string = '1915c903-0f25-4435-962a-3014eedfe2ef';
  public _outlookemailsid: string = 'de692daf-26ef-4489-b83b-19f4fc83af27';
  public _mail: any;
  private _mailmessage:string;
  public _mailsubject:string;
  //private _graphclient:any;

  constructor(context: WebPartContext, mail: any,graphFactory:MSGraphClientFactory) {
    this._spclient = context.spHttpClient;
    this._currentuser = context.pageContext.user;
    this._mail = mail;
    graphFactory.getClient().then((client:MSGraphClient)=>{
      this._graphclient=client;
    });
    //console.log("Email: ", this._mail);
    mail.body.getAsync('text', (result)=> {
      if (result.status === 'succeeded') {
        this._mailmessage=result.value;
      }
    });

    

  }

  public getRootSite(): Promise<any> {
    console.log(this._currentuser);
    const openticketsurl = `https://oaktondidata.sharepoint.com/sites/Test3/_api/site/RootWeb`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        console.log("Service ", data);
        return data;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

  public getCategories(): Promise<any> {
    const openticketsurl = `${this._weburl}_api/Web/Lists(guid'873475f3-0aeb-4ae9-b900-c27f5f8bfd0f')/Items?$select=ID,Title`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        let cats: any[] = [];
        const def = {
          key: "-1",
          text: "-Vælg-"
        };
        cats.push(def);
        //cats.push({})
        data.value.forEach(x => {
          const cat = {
            key: x.ID,
            text: x.Title
          };
          cats.push(cat);
        });
        return cats;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

  public getCaseStatus(): Promise<any> {
    const openticketsurl = `${this._weburl}_api/Web/Lists(guid'${this._caseid}')/Fields(guid'${this._statusid}')?$select=Choices`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        let stats: any[] = [];
        const def = {
          key: "-1",
          text: "-Vælg-"
        };
        stats.push(def);
        data.Choices.forEach(x => {
          const stat = {
            key: x,
            text: x
          };
          stats.push(stat);
        });
        return stats;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

  public getCases(status?: string): Promise<any> {
    let openticketsurl = '';
    if (status != null && status.length > 0) {
      openticketsurl = `${this._weburl}_api/Web/Lists(guid'${this._caseid}')/Items?$select=ID,Title,Status&$filter=Status eq '${status}'`;
    }
    else {
      openticketsurl = `${this._weburl}_api/Web/Lists(guid'${this._caseid}')/Items?$select=ID,Title,Status`;
    }

    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        //console.log("Cases ", data.value);
        let cass: any[] = [];
        const def = {
          key: "-1",
          text: "-Vælg-"
        };
        cass.push(def);
        data.value.forEach(x => {
          const cas = {
            key: x.ID,
            text: x.Title
          };
          cass.push(cas);
        });
        return cass;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

  public getCaseFolderTitle(caseid?: string): Promise<any> {
    const casefoldersurl = `${this._weburl}_api/Web/Lists(guid'${this._caselibraryid}')/Items?$filter=RelatedItemId eq '${caseid}' &$select=ID,FileLeafRef,Title`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(casefoldersurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        //console.log("Service ", data.value);
        let casetitle: string = '';
        //cats.push({})
        data.value.forEach(x => {
          casetitle = x.Title;
        });
        return casetitle;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }
  ///sites/xRMLite/_api/Web/GetFolderByServerRelativePath(decodedurl='/sites/xRMLite/CaseFiles/1')/Folders?&$select=Name
  public getCaseSubFolders(folderrelativepath?: string): Promise<any> {
    debugger;
    const casefoldersurl = `${this._weburl}_api/Web/GetFolderByServerRelativePath(decodedurl='${this._casefolderrelativepath}/${folderrelativepath}')/Folders?&$select=Name`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(casefoldersurl, SPHttpClient.configurations.v1, options).then(
      (response: any) => {
        if (response.status >= 200 && response.status < 300) {
          return response.json();
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      })
      .then((data: any) => {
        //console.log("SubFolders ", data);
        let folders: any[] = [];
        const def = {
          key: "-1",
          text: "-Vælg-"
        };
        folders.push(def);
        data.value.forEach(x => {
          const folder = {
            key: x.Name,
            text: x.Name
          };
          folders.push(folder);
        });
        return folders;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

  public saveemail(addinobj): Promise<any> {
    const lmail = this._mail;
    const emailobj = {
      Title: lmail.subject,
      Message: this._mailmessage,
      To: this.buildAddressString(lmail.to),
      From: `${lmail.from.displayName}:${lmail.from.emailAddress}`,
      CategoryId:addinobj.catid,
      RelatedItemListId: "Lists/Cases",
      RelatedItemId:addinobj.caseid,
      ConversationId:lmail.conversationId,
      ConversationTopic:lmail.subject,
      InOut:"In"
    };
    //console.log(emailobj);
    const addemailurl: string = `${this._weburl}_api/web/lists(guid'${this._outlookemailsid}')/items`;
    const httpclientoptions: ISPHttpClientOptions = {
      body: JSON.stringify(emailobj)
    };

    return this._spclient.post(addemailurl, SPHttpClient.configurations.v1, httpclientoptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status >= 200 && response.status < 300) {
          return response.status;
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      });
  }

  public saveAttachment(relativepath:string,filename:string,filecontent:any):Promise<any>{
    const addemailurl: string = `${this._weburl}_api/Web/GetFolderByServerRelativePath('/sites/xRMLite/CaseFiles/1')/Files/add(overwrite=true,url='${filename}')`;
    const httpclientoptions: ISPHttpClientOptions = {
      body: atob(filecontent)
    };

    return this._spclient.post(addemailurl, SPHttpClient.configurations.v1, httpclientoptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status >= 200 && response.status < 300) {
          return response.status;
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      });
  }

  public getAttachments():void{

      // get information about the current user from the Microsoft Graph
      this._graphclient.api(`/me/messages/${this._mail.itemId}/attachments`).get((error, response: any, rawResponse?: any) => {
          debugger;
          console.log("Graph Response ",response);
          response.value.forEach((file)=>{
            this.saveAttachment("",file.name,file.contentBytes);
          });
      });
  }

  private buildAddressString(addresses): string {
    let y: string="";
    addresses.forEach((x) => {
      y = `${y}${x.displayName}:${x.emailAddress};`;
    });
    return y;
  }

  public composemail(addinid:string):void{
    let mailRecepients = [{
      "displayName": "",
      "emailAddress": "ankerh@emails.itsm360cloud.net"
  }];
      this._mail.subject.setAsync(`${this._mailsubject} ${addinid}`, (asyncResult) =>{
        if (asyncResult.status === "failed") {
          console.log("Action failed with error: " + asyncResult.error.message);
      } else {
          console.log("Action Subject appended");
          this._mail.bcc.setAsync(mailRecepients, (result) =>{
              if (result.error) {
                  console.log(result.error);
              } else {
                  console.log("Recipients added to the bcc");
                  Office.context.ui.closeContainer();
              }
          });
      }
    });
    
    
  }

}
