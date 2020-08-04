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
import * as React from 'react';
import styles from './HappyAnniversary.module.scss';
import HappyBirdthayCard from '../happyAnniversaryCard/HappyAnniversaryCard';
import * as moment from 'moment';
var HappyAnniversary = (function (_super) {
    __extends(HappyAnniversary, _super);
    // private _showAnniversarys: boolean = true;
    function HappyAnniversary(props) {
        return _super.call(this, props) || this;
    }
    //
    HappyAnniversary.prototype.render = function () {
        return (React.createElement("div", { className: styles.happyAnniversary }, this.props.users.map(function (user) {
            return (React.createElement("div", { className: styles.container },
                React.createElement(HappyBirdthayCard, { userName: user.userName, jobDescription: user.jobDescription, anniversary: moment(user.anniversary, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MMM DD'), userEmail: user.userEmail, Ann_days: user.Ann_days })));
        })));
    };
    return HappyAnniversary;
}(React.Component));
export { HappyAnniversary };
export default HappyAnniversary;
//# sourceMappingURL=HappyAnniversary.js.map