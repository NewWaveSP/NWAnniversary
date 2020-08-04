import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { MSGraphClient } from "@microsoft/sp-http";
import * as moment from 'moment';

export class SPService {
  private graphClient: MSGraphClient = null;
  constructor(private _context: WebPartContext | ApplicationCustomizerContext) { }
  // Get Profiles
  public async getPAnniversarys(upcommingDays: number): Promise<any[]> {
    let _results, _today: string, _month: string, _day: number;
    let _filter: string, _orderby: string;
    let _FinalDate: string; 
    var nextLink = "@odata.nextLink";

    var _Userresults = [];

    try {
      // If we are in Dezember we have to look if there are Anniversary in January
      // we have to build a condition to select Anniversary in January based on number of upcommingDays
      // we can not use the year for teste , the year is always 2000.
      // console.log(_month);
      // _results = await this.graphClient.api(`users?$select=mail,displayName,Department,extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber ,accountEnabled&$top=999`).version('beta').get();
     // _results = await this.graphClient.api('/users').version('beta').get();
      // _filter = "accountEnabled eq true";
      // _orderby = "fields/Anniversary";
     // _results = null;
      _today = '2000-' + moment().format('MM-DD');
      _month = moment().format('MM');
      _month = _month.replace('0','')
      _day = parseInt(moment().format('DD'));
      _filter = "(startswith(extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber,'" + _month + "')) and (accountEnabled eq true)";
      let userProperties = ["mail","displayName","jobTitle","extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber","accountEnabled"];
      this.graphClient = await this._context.msGraphClientFactory.getClient();
      _results = await this.graphClient.api('/users').version('beta').top(999).select(userProperties).filter(_filter).get();
    //_results = await this.graphClient.api('/users').version('beta').select(userProperties).filter(_filter).get();
    // _results = await this.graphClient.api('/users').version('beta').select(userProperties).get();
      _Userresults = _Userresults.concat(_results.value);
      
 /*
      while(_results[nextLink]) {
        _results[nextLink] = _results[nextLink].replace('https://graph.microsoft.com/beta/', '');
        _results = await this.graphClient.api('/users').version('beta').select(userProperties).get();
        _Userresults = _Userresults.concat(_results.value);
        if( !(_results[nextLink]) ) {  break;  }
      }
      

      if (_results[nextLink]) {
        _results[nextLink] = _results[nextLink].replace('https://graph.microsoft.com/beta/', '');
        _results = await this.graphClient.api('/users').version('beta').select(userProperties).get();
        _Userresults = _Userresults.concat(_results.value);
      }

      if (_results[nextLink]) {
        _results[nextLink] = _results[nextLink].replace('https://graph.microsoft.com/beta/', '');
        _results = await this.graphClient.api('/users').version('beta').select(userProperties).get();
        _Userresults = _Userresults.concat(_results.value);
      }

*/
      

      return _Userresults;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
export default SPService;


