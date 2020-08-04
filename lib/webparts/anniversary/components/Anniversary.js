var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
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
import * as React from 'react';
import styles from './Anniversary.module.scss';
import { HappyAnniversary } from '../../../controls/happyanniversary';
import * as moment from 'moment';
import SPService from '../../../services/SPService';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
var imgBackgroundBallons = require('../../../../assets/ballonsBackgroud.png');
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'AnniversaryWebPartStrings';
var listItems = [];
var Anniversarys = (function (_super) {
    __extends(Anniversarys, _super);
    function Anniversarys(props) {
        var _this = _super.call(this, props) || this;
        _this._users = [];
        _this._tempusers = [];
        _this._spServices = new SPService(_this.props.context);
        _this.state = {
            Users: [],
            showAnniversarys: true
        };
        return _this;
    }
    Anniversarys.prototype.componentDidMount = function () {
        this.GetUsers();
    };
    // Render
    Anniversarys.prototype.render = function () {
        var _center = !this.state.showAnniversarys ? "center" : "";
        {
            return (React.createElement("div", { className: styles.anniversary, style: { textAlign: _center } },
                React.createElement("div", { className: styles.container },
                    React.createElement(WebPartTitle, { displayMode: this.props.displayMode, title: this.props.title, updateProperty: this.props.updateProperty }),
                    !this.state.showAnniversarys ?
                        React.createElement("div", { className: styles.backgroundImgBallons },
                            React.createElement(Image, { imageFit: ImageFit.cover, src: imgBackgroundBallons, width: 150, height: 150 }),
                            React.createElement(Label, { className: styles.subTitle }, strings.MessageNoAnniversarys))
                        :
                            React.createElement(HappyAnniversary, { users: this.state.Users }))));
        }
    };
    // Sort Array of Anniversarys
    Anniversarys.prototype.SortAnniversarys = function (users) {
        return users.sort(function (a, b) {
            if (a.Fakeanniversary > b.Fakeanniversary) {
                return 1;
            }
            if (a.Fakeanniversary < b.Fakeanniversary) {
                return -1;
            }
            return 0;
        });
    };
    Anniversarys.prototype.ordinal_suffix_of = function (Ann_number) {
        var j = Ann_number % 10, k = Ann_number % 100;
        if (j == 1 && k != 11) {
            return Ann_number + "st";
        }
        if (j == 2 && k != 12) {
            return Ann_number + "nd";
        }
        if (j == 3 && k != 13) {
            return Ann_number + "rd";
        }
        return Ann_number + "th";
    };
    // Load List Of Users
    Anniversarys.prototype.GetUsers = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _otherMonthsAnniversarys, _dezemberAnniversarys, CurrentYear, AllowedItem, _i, listItems_1, item, AnniversaryDate, TodayDate, YearDiff, Yeardiff, FakeAnniversary, Fake_Anniversary, DateDiff, Isexpired;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._spServices.getPAnniversarys(this.props.NumberAnniversary)];
                    case 1:
                        listItems = _a.sent();
                        AllowedItem = this.props.NumberAnniversary;
                        if (listItems && listItems.length > 0) {
                            _otherMonthsAnniversarys = [];
                            _dezemberAnniversarys = [];
                            for (_i = 0, listItems_1 = listItems; _i < listItems_1.length; _i++) {
                                item = listItems_1[_i];
                                if (item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber) {
                                    AnniversaryDate = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber);
                                    TodayDate = moment();
                                    YearDiff = TodayDate.diff(AnniversaryDate, 'year');
                                    YearDiff = YearDiff + 1;
                                    Yeardiff = this.ordinal_suffix_of(YearDiff);
                                    CurrentYear = moment().format('YYYY');
                                    FakeAnniversary = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber, 'MM/DD/YYYY').year(CurrentYear).format();
                                    Fake_Anniversary = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber, 'MM/DD/YYYY').year(CurrentYear);
                                    DateDiff = TodayDate.diff(Fake_Anniversary, 'days');
                                    Isexpired = Fake_Anniversary.diff(TodayDate, 'days');
                                    if (Isexpired < 0) {
                                        FakeAnniversary = moment(FakeAnniversary).add(1, 'years').format();
                                        YearDiff = YearDiff + 1;
                                        Yeardiff = this.ordinal_suffix_of(YearDiff);
                                    }
                                    this._users.push({ key: item.mail, userName: item.displayName, userEmail: item.mail, jobDescription: item.jobTitle, anniversary: moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber).local().format(), Ann_days: Yeardiff, Fakeanniversary: FakeAnniversary });
                                }
                            }
                            //  this._users=[];
                            this._tempusers = this.SortAnniversarys(this._users);
                            this._users = [];
                            this._users = this._tempusers.filter(function (itm, i) {
                                if (AllowedItem > i)
                                    return itm;
                            });
                            this.setState({
                                Users: this._users,
                                showAnniversarys: this._users.length === 0 ? false : true
                            });
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    return Anniversarys;
}(React.Component));
export default Anniversarys;
//# sourceMappingURL=Anniversary.js.map