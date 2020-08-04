var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as moment from 'moment';
var SPService = (function () {
    function SPService(_context) {
        this._context = _context;
        this.graphClient = null;
    }
    // Get Profiles
    SPService.prototype.getPAnniversarys = function (upcommingDays) {
        return __awaiter(this, void 0, void 0, function () {
            var _results, _today, _month, _day, _filter, _orderby, _FinalDate, nextLink, _Userresults, userProperties, _a, error_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        nextLink = "@odata.nextLink";
                        _Userresults = [];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
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
                        _month = _month.replace('0', '');
                        _day = parseInt(moment().format('DD'));
                        _filter = "(startswith(extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber,'" + _month + "')) and (accountEnabled eq true)";
                        userProperties = ["mail", "displayName", "jobTitle", "extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber", "accountEnabled"];
                        _a = this;
                        return [4 /*yield*/, this._context.msGraphClientFactory.getClient()];
                    case 2:
                        _a.graphClient = _b.sent();
                        return [4 /*yield*/, this.graphClient.api('/users').version('beta').top(999).select(userProperties).filter(_filter).get()];
                    case 3:
                        _results = _b.sent();
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
                        return [2 /*return*/, _Userresults];
                    case 4:
                        error_1 = _b.sent();
                        console.dir(error_1);
                        return [2 /*return*/, Promise.reject(error_1)];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    return SPService;
}());
export { SPService };
export default SPService;
//# sourceMappingURL=SPService.js.map