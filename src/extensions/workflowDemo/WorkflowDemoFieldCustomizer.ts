import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration} from '@microsoft/sp-http';  
import * as strings from 'WorkflowDemoFieldCustomizerStrings';
import styles from './WorkflowDemoFieldCustomizer.module.scss';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import ColorPickerDialog from './ColorPickerDialog';
//import { Dialog, BaseDialog } from '@microsoft/sp-dialog';
/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IWorkflowDemoFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
  CDNUrl:string;
}


declare var jQuery: any;
declare var Bluebox: any;
declare var BlueboxSPFX: any;




const LOG_SOURCE: string = 'WorkflowDemoFieldCustomizer';

export default class WorkflowDemoFieldCustomizer
  extends BaseFieldCustomizer<IWorkflowDemoFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    var app = this;
    var appProp = app.properties;
    var pageCtx = app.context.pageContext;
    var isSameSession = false;
    var isSameSessionAndSiteCollection = false;
    var bbx:any;
    var bbxSPFxVersion = "1.0.0.4";

    //DEBUG OVERRIDE
    appProp.CDNUrl = "https://bbxclientsdevstoragecdn.blob.core.windows.net";

    //Define Global Namespace and variables
    if(window["Bluebox"]){
        bbx = window["Bluebox"];
        isSameSession = true;
    } else{
        bbx = {
            _Name: "Bluebox",
            _Version: bbxSPFxVersion,
            Constant: {}
        };
        //window["Bluebox"] = bbx;
        //window["BlueboxSPFX"] = true; //For Bluebox Codes to differentiate between classic and modern.
    }
    
    //For Legacy Purposes.
    //Will be refreshed every page load.
    window["_spPageContextInfo"] = pageCtx.legacyPageContext;

    var tenantName:string = location.host.split(".")[0];
    var tenantSite:string = pageCtx.site.serverRelativeUrl;
    var tenantWeb:string = pageCtx.web.serverRelativeUrl;
    var cdnCoreUrl:string = null; //app.UrlCombine(appProp.CDNUrl, "sp-common", appProp.CDNVersion);
    var cdnRootUrl:string = app.UrlCombine(appProp.CDNUrl, tenantName);
    var cdnSiteUrl:string;
    var buildNumber;

    ///////////////////////////////////////////////////////////////////////
    //Build CDN Tenant Url

    //Get Site Name if not at "root" site collection
    //Root: tenantSite == '/'
    //Site: tenantSite == '/sites/sitename' => 'sitename'

    if(tenantSite.length > 1) {
        tenantSite = tenantSite.split("/")[2];
    }

    if(tenantWeb.length > 1) {
        tenantWeb = tenantWeb.split(tenantSite).pop();
        if(tenantWeb.substring(0, 1) == "/")
            tenantWeb = tenantWeb.substring(1);
    }

    cdnSiteUrl = app.UrlCombine(cdnRootUrl, tenantSite);

    var initialPromise = null;

    if(isSameSessionAndSiteCollection) {
      return Promise.resolve();
    }

    //return promise for oninit, inside callback wait for blueboxcore stuff to finishloading then resolve() for onRenderCell to execute
    return new Promise((resolve, reject) => {
        //Prepare site setting to build Constant object.
        bbx._SiteSetting.cdnCoreUrl = cdnCoreUrl;
        bbx._SiteSetting.cdnRootUrl = cdnRootUrl;
        bbx._SiteSetting.cdnSiteUrl = cdnSiteUrl;

        //Build Constant Object Immediately if Utility is already loaded
        //Otherwise, wait for it to be loaded.
        if(bbx.Utility && jQuery.fn.popr){
            bbx.Utility.Constant.Build(bbx._SiteSetting);
                //load popr css if not loaded
            if(jQuery("link[href*='popr.css']").length == 0) {
              window["Bluebox"].Utility.Loader.IncludeCss("popr.css","https://bbxclientsdevstoragecdn.blob.core.windows.net/sp-common/3.18/ext/popr/popr.css", 0);
              console.log('popr css loaded');
            }
            resolve();
        } else {
          console.log('Waiting on Utility and popr');
            var waitMax = 550;
            var waitCount = 0;
            var intervalId = setInterval(() => {
                if(bbx.Utility){
                    clearInterval(intervalId);
                    bbx.Utility.Constant.Build(bbx._SiteSetting);
                    resolve();
                } else {
                    waitCount++;
                    if(waitCount >= waitMax) {
                        clearInterval(intervalId);
                        console.error("Fail to wait for Bluebox.Utility");
                        reject();
                    }
                }
            }, 100);
        }
    });
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    console.log(jQuery.fn.popr);
    console.log("Rendering");
    let itemId = event.listItem.getValueByName('ID').toString();
    let listId = this.context.pageContext.list.id.toString();
    let fieldHtml: HTMLDivElement = event.domElement.firstChild as HTMLDivElement;
    let listName = event.listItem.getValueByName('_ModerationStatus');
    let contentTypeId =  event.listItem.getValueByName('ContentTypeId');

    (<any>window).ShowAuditHistory = (listId, itemId) =>{
      //Dialog.alert(`this is a dialog`)    
      this._makePOSTRequest(listId, itemId);
    }

    let poprId = `popr-action-${itemId}`;

    //parent element overflow set to visible so popr items can be seen
    fieldHtml.parentElement.parentElement.style.overflow = 'visible';
    this._render(itemId,poprId,listId,contentTypeId,listName,fieldHtml);
    jQuery('.popr').popr();
    jQuery('.popr').show();
    //event.domElement.innerHTML = "test";
    //ReactDOM.render(actionCog, event.domElement);
  }

  private _render(itemId: string, poprId: string, listId: string, contentTypeId: String, listName: String, fieldHtml: HTMLDivElement): void {
    var html = [];
    
    
    html.push('<div class="popr li-icon-cog" style="font-size: 24px" data-id="' + poprId + '" ></div>');
    html.push('<div class="popr-box" data-box-id="' + poprId + '" style="display:none;">');
    html.push(`<div class="disp-audit-history"><a onclick="ShowAuditHistory('${listId}','${itemId}')" href="javascript:;" ><div class="popr-item">ACTION</div></a></div>`);
    html.push('<div class="disp-doc-suggestions"></div>');


    console.log("finished");
    jQuery('.popr').popr();
    jQuery('.popr').show();
    fieldHtml.innerHTML = html.join("");
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }


  private IncludeJsAsync(id:string, url:string, build:string) : Promise<any> {
    var result:any = {
        isPreviouslyExist: false,
        id: id,
        url: url + "?v=" + build
    };

    var node:any = document.getElementById(id);

    return new Promise((resolve, reject) => {
        if (!node) {
            node = document.createElement('script');
            node.type = 'text/javascript';
            node.src = result.url;
            node.id = result.id;
            node.setAttribute("async", "async");
            node.onload = () => { resolve(result); };
            node.onerror = () => { reject(result); };
            document.getElementsByTagName('head')[0].appendChild(node);
        } else {
            result.isPreviouslyExist = true;
            resolve(result);
        }
    });
}

private UrlNoTrailingSlash(url:string) : string {
  var size = url.length;
  if(size > 0 && url.substring(size - 1) == "/") {
      url = url.substring(0, size - 1);
  }
  return url;
}

private UrlCombine(...args:any[]) : string {
  var tokens: string[] = [];

  //Remove trailing slash (/) from first argument
  if(args.length > 0) {
      tokens.push(this.UrlNoTrailingSlash(args[0].trim()));
  }

  //Process remaining arguments
  for(var i: number = 1; i < args.length; i++) {
      var url:string = args[i].trim();

      //Strip Leading Slash
      if(url.substring(0, 1) == "/")
          url = url.substring(1);

      tokens = tokens.concat(url.split("/"));
  }

  return this.UrlNoTrailingSlash(tokens.join("/"));
}


private _makePOSTRequest(listId: string, itemId: string): void {
  console.log('item', itemId);
  const spOpts: ISPHttpClientOptions = {
    //body: `{ Title: 'Developer Workbench', BaseTemplate: 100 }`
  };
//https://blueboxsolutionsdev.sharepoint.com/teams/binh_spfx/wfsvc/2abfa33f4c004f258e9f948f4c1981d1/WFInitForm.aspx?List={e2eb4b8a-e186-4be1-b004-062267f22d37}&ID=14&ItemGuid={572BFBE8-241A-443C-AD6E-FE949B07B499}&TemplateID={B638F160-28BB-41CE-8FB3-D6AD270C2124}&WF4=1&Source=https%3A%2F%2Fblueboxsolutionsdev%2Esharepoint%2Ecom%2Fteams%2Fbinh%5Fspfx%2FLists%2Flist1%2FAllItems%2Easpx%3FloadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests%2Ejs%26fieldCustomizers%3D%7B%2522Action%2522%3A%7B%2522id%2522%3A%2520%252297d4a9b7%2D61b5%2D457a%2Daea0%2Df22c587916b1%2522%7D%7D
     var url = "https://blueboxsolutionsdev.sharepoint.com/teams/binh_spfx/wfsvc/2abfa33f4c004f258e9f948f4c1981d1/WFInitForm.aspx?List={e2eb4b8a-e186-4be1-b004-062267f22d37}&ID=14&ItemGuid={572BFBE8-241A-443C-AD6E-FE949B07B499}&TemplateID={B638F160-28BB-41CE-8FB3-D6AD270C2124}&WF4=1&Source=https%3A%2F%2Fblueboxsolutionsdev%2Esharepoint%2Ecom%2Fteams%2Fbinh%5Fspfx%2FLists%2Flist1%2FAllItems%2Easpx%3FloadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests%2Ejs%26fieldCustomizers%3D%7B%2522Action%2522%3A%7B%2522id%2522%3A%2520%252297d4a9b7%2D61b5%2D457a%2Daea0%2Df22c587916b1%2522%7D%7D";
  //this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists`, SPHttpClient.configurations.v1, spOpts)
     this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/SP.WorkflowServices.WorkflowSubscriptionService.Current/EnumerateSubscriptionsByList(guid'e2eb4b8a-e186-4be1-b004-062267f22d37')?select=ID,Name`, SPHttpClient.configurations.v1, spOpts)
    .then((response: SPHttpClientResponse) => {
      // Access properties of the response object. 
      console.log(`Status code: ${response.status}`);
      console.log(`Status text: ${response.statusText}`);
      //var wfURL = "https://blueboxsolutionsdev.sharepoint.com/teams/binh_spfx/wfsvc/2abfa33f4c004f258e9f948f4c1981d1/WFInitForm.aspx?List={e2eb4b8a-e186-4be1-b004-062267f22d37}&ID=14&ItemGuid={572BFBE8-241A-443C-AD6E-FE949B07B499}&TemplateID={B638F160-28BB-41CE-8FB3-D6AD270C2124}&WF4=1&Source=https%3A%2F%2Fblueboxsolutionsdev%2Esharepoint%2Ecom%2Fteams%2Fbinh%5Fspfx%2FLists%2Flist1%2FAllItems%2Easpx%3FloadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests%2Ejs%26fieldCustomizers%3D%7B%2522Action%2522%3A%7B%2522id%2522%3A%2520%252297d4a9b7%2D61b5%2D457a%2Daea0%2Df22c587916b1%2522%7D%7D&isdlg=1";
      var wfURL = `${this.context.pageContext.web.absoluteUrl}/wfsvc/2abfa33f4c004f258e9f948f4c1981d1/WFInitForm.aspx?List=${listId}&ID=${itemId}&ItemGuid={572BFBE8-241A-443C-AD6E-FE949B07B499}&TemplateID={B638F160-28BB-41CE-8FB3-D6AD270C2124}&WF4=1&Source=https%3A%2F%2Fblueboxsolutionsdev%2Esharepoint%2Ecom%2Fteams%2Fbinh%5Fspfx%2FLists%2Flist1%2FAllItems%2Easpx%3FloadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests%2Ejs%26fieldCustomizers%3D%7B%2522Action%2522%3A%7B%2522id%2522%3A%2520%252297d4a9b7%2D61b5%2D457a%2Daea0%2Df22c587916b1%2522%7D%7D&isdlg=1`;
      //window.open("https://blueboxsolutionsdev.sharepoint.com/teams/binh_spfx/wfsvc/2abfa33f4c004f258e9f948f4c1981d1/WFInitForm.aspx?List={e2eb4b8a-e186-4be1-b004-062267f22d37}&ID=14&ItemGuid={572BFBE8-241A-443C-AD6E-FE949B07B499}&TemplateID={B638F160-28BB-41CE-8FB3-D6AD270C2124}&WF4=1&Source=https%3A%2F%2Fblueboxsolutionsdev%2Esharepoint%2Ecom%2Fteams%2Fbinh%5Fspfx%2FLists%2Flist1%2FAllItems%2Easpx%3FloadSPFX%3Dtrue%26debugManifestsFile%3Dhttps%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests%2Ejs%26fieldCustomizers%3D%7B%2522Action%2522%3A%7B%2522id%2522%3A%2520%252297d4a9b7%2D61b5%2D457a%2Daea0%2Df22c587916b1%2522%7D%7D");
      const dialog: ColorPickerDialog = new ColorPickerDialog();

      console.log(wfURL);
      console.log(url);
      dialog.hidden = false;
      dialog.url = wfURL;
    
      //do width and height calculation here
      //dialog.width = "1350px";
      //dialog.height = "650px";
      dialog.width = "500px";
      dialog.height = "650px";
      dialog.show().then(()=>{
        dialog.close();
      });


      //response.json() returns a promise so you get access to the json in the resolve callback.
      response.json().then((responseJSON: JSON) => {
        console.log(responseJSON);
      });
    });
}











}
