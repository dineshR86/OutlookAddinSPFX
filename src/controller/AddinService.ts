import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse,MSGraphClient, MSGraphClientFactory, ISPHttpClientBatchOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPUser } from "@microsoft/sp-page-context";

export class AddinService {
  private _spclient: SPHttpClient;
  private _graphclient:MSGraphClient;
  public _currentuser: SPUser;
  public _defConfigData:any;
  // public _weburl: string = 'https://ankerhan.sharepoint.com/sites/Sager/';
  // public _casefolderrelativepath: string = '/sites/Sager/CaseFiles';
  // public _caseid: string = '32a64dc2-f7ef-4be0-b555-5d7c7c1a57be';
  // public _statusid: string = '17216dcb-35ca-4675-8838-818ac63fdc30';
  // public _caselibraryid: string = '2f5a0de2-55d3-4f05-a1fc-7e35d6fad5da';
  // public _outlookemailsid: string = '6f4a8255-ac2e-4acd-ad0b-5886a8a35865';
  // public _casedriveid:string='b!PDBGl16kuUSfD7ci8Pd9dMO5nQPkcrpBhY_Bobl8PRHiDVov01UFT6H8fjXW-tXa';
  // public _catid:string='fb1e18ad-29c3-4889-8b28-0c28e960de30';
  
  public _weburl: string = 'https://cloudmission.sharepoint.com/sites/xRMLite/';
  public _casefolderrelativepath: string = '/sites/xRMLite/CaseFiles';
  public _caseid: string = 'b5fd8cf2-1277-4daa-9196-98a1b6d32401';
  public _statusid: string = '17216dcb-35ca-4675-8838-818ac63fdc30';
  public _caselibraryid: string = '1915c903-0f25-4435-962a-3014eedfe2ef';
  public _outlookemailsid: string = 'de692daf-26ef-4489-b83b-19f4fc83af27';
  public _casedriveid:string='b!HQciseR9TEyvXyK7-eL2DEEg8eS76CtEmIrZT4djMw8DyRUZJQ81RJYqMBTu3-Lv';
  public _catid:string='873475f3-0aeb-4ae9-b900-c27f5f8bfd0f';
  public _addinconfig:string='ae886127-283d-4e43-92af-387766613759';

  public _mail: any;
  private _mailmessage:string;
  public _mailsubject:string;
  public _mailbody:string;
  //private _graphclient:any;

  constructor(context: WebPartContext, mail: any,graphFactory:MSGraphClientFactory) {
    this._spclient = context.spHttpClient;
    this._currentuser = context.pageContext.user;
    this._mail = mail;
    graphFactory.getClient().then((client:MSGraphClient)=>{
      this._graphclient=client;
    });
    // console.log("Current User: ", this._currentuser);
    mail.body.getAsync('text', (result)=> {
      if (result.status === 'succeeded') {
        this._mailmessage=result.value;
      }
    });
    // this.getConfigData().then((dat)=>{
    //   this._defConfigData=dat;
    //   console.log("Config ",this._defConfigData);
    // });
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
    const catsurl = `${this._weburl}_api/Web/Lists(guid'${this._catid}')/Items?$select=ID,Title`;
    const options: ISPHttpClientOptions = {
      headers: {
        "odata-version": "3.0",
        "accept": "application/json;odata=nometadata"
      },
      method: "GET"
    };
    return this._spclient.get(catsurl, SPHttpClient.configurations.v1, options).then(
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
          key: -1,
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
        console.log("Service ", data.value);
        let casetitle: string = '';
        //cats.push({})
        data.value.forEach(x => {
          casetitle = x.FileLeafRef;
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

  // public saveAttachment(relativepath:string,filename:string,filecontent:any):Promise<any>{
  //   // const addemailurl: string = `${this._weburl}_api/Web/GetFolderByServerRelativePath('/sites/xRMLite/CaseFiles/1')/Files/add(overwrite=true,url='${filename}')`;
  //   // const httpclientoptions: ISPHttpClientOptions = {
  //   //   body: atob(filecontent)
  //   // };

  //   // return this._spclient.post(addemailurl, SPHttpClient.configurations.v1, httpclientoptions)
  //   //   .then((response: SPHttpClientResponse) => {
  //   //     if (response.status >= 200 && response.status < 300) {
  //   //       return response.status;
  //   //     }
  //   //     else { return Promise.reject(new Error(JSON.stringify(response))); }
  //   //   });
  //   // let web=new Web(this._weburl);
  //   // return web.getFolderByServerRelativeUrl("CaseFiles/1").files.add(filename,atob(filecontent),true).then((result)=>{
  //   //   return result;
  //   // });
  // }

  public saveAttachments(folderpath:string):void{
    let attachments:any[]=[];
    //console.log("attachments",attachments);
    const mailid=this._mail.itemId;
      this._graphclient.api(`/me/messages/${mailid}/attachments`).get((error, response: any, rawResponse?: any) => {
        attachments=response.value;
        console.log("attachments",attachments);
        attachments.forEach((x,index)=>{
          const data = atob(x.contentBytes);
          const array = Uint8Array.from(data, b => b.charCodeAt(0));
          this._graphclient.api(`drives/${this._casedriveid}/root:/${folderpath}/${x.name}:/content`).put(array).then((result)=>{
            console.log("Sucess",result);
            if(attachments.length==(index+1)){
              Office.context.ui.closeContainer();
            }
          }).catch((ex)=>{
            console.log("Error",ex);
          });
          //console.log("Index ",index);
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

    this._mail.body.getTypeAsync((result)=>{
      if (result.status == Office.AsyncResultStatus.Failed){
        console.log(result.error.message);
    }else{
      console.log("email type: ",result.value);
      if (result.value == 'html'){
        //this._mail.body.setSelectedDataAsync
        this._mail.body.prependAsync(
          `<div style="margin-left:90%;font-size:8px"><span hidden>###AHC-REF-${addinid}###</span></div>`,
          { coercionType: Office.CoercionType.Html, 
          asyncContext: { var3: 1, var4: 2 } },
          (asyncResult)=> {
              if (asyncResult.status == 
                  Office.AsyncResultStatus.Failed){
                  console.log(asyncResult.error.message);
              }
              else {
                this._mail.bcc.setAsync(mailRecepients, (bccresult) =>{
                  if (bccresult.error) {
                      console.log(bccresult.error);
                  } else {
                      console.log("Recipients added to the bcc");
                      Office.context.ui.closeContainer();
                  }
      });
              }
          });
      }else{
       this._mail.body.prependAsync(
        `###AHC REF ${addinid}###`,
          { coercionType: Office.CoercionType.Text, 
              asyncContext: { var3: 1, var4: 2 } },
          (asyncResult) =>{
              if (asyncResult.status == 
                  Office.AsyncResultStatus.Failed){
                  console.log(asyncResult.error.message);
              }
              else {
                this._mail.bcc.setAsync(mailRecepients, (bccresult) =>{
                  if (bccresult.error) {
                      console.log(bccresult.error);
                  } else {
                      console.log("Recipients added to the bcc");
                      Office.context.ui.closeContainer();
                  }
      });
              }
           });
      }
    }
    });
    
    
  }

  public saveConfigData(configdat:any):Promise<any>{
    const addinconfigobj = {
      Title: configdat.case,
      StatusID:configdat.status,
      UsersMail:this._currentuser.email
    };
    console.log(addinconfigobj);
    const addemailurl: string = `${this._weburl}_api/web/lists(guid'${this._addinconfig}')/items`;
    const httpclientoptions: ISPHttpClientOptions = {
      body: JSON.stringify(addinconfigobj)
    };

    return this._spclient.post(addemailurl, SPHttpClient.configurations.v1, httpclientoptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status >= 200 && response.status < 300) {
          return response.status;
        }
        else { return Promise.reject(new Error(JSON.stringify(response))); }
      });
  }

  public updateConfigData(configdat:any,configid:string):Promise<any>{
    const updateurl=`${this._weburl}_api/web/lists(guid'${this._addinconfig}')/items(${configid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._addinconfig}')/items(${configid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify({
                  Title: configdat.case,
                  StatusID:configdat.status
                  });
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json",
                        "Content-Type":"application/json",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
            
          }).catch((ex) => {
                console.log("Error while updating status: ", ex);
                throw ex;
            });
  }

  public getConfigData():Promise<any>{
    const _curemail=this._currentuser.email;
    const openticketsurl = `${this._weburl}_api/Web/Lists(guid'${this._addinconfig}')/Items?$select=ID,Title,StatusID,UsersMail&$filter=UsersMail eq '${_curemail}'`;
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
        let configdata:any;
        data.value.forEach(x => {
          configdata={
            Case:x.Title,
            Status:x.StatusID,
            ID:x.ID
          };
        });
        return configdata;
      }).catch((ex) => {
        console.log("Error while fetching My tickets count: ", ex);
        throw ex;
      });
  }

}
