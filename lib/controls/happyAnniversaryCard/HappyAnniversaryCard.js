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
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
import * as React from 'react';
import styles from './HappyAnniversaryCard.module.scss';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'AnniversaryWebPartStrings';
import * as moment from 'moment';
import { DocumentCardActions, } from 'office-ui-fabric-react/lib/DocumentCard';
var img = require('../../../assets/baloons.png');
var IMG_WIDTH = 200;
var IMG_HEIGTH = 190;
var HappyAnniversaryCard = (function (_super) {
    __extends(HappyAnniversaryCard, _super);
    function HappyAnniversaryCard(props) {
        var _this = _super.call(this, props) || this;
        _this._AnniversaryMsg = '';
        // Render tertiary text
        _this._onRenderTertiaryText = function (props) {
            return (React.createElement("div", null,
                React.createElement("span", { className: 'ms-fontWeight-semibold', style: { color: '#71afe5' } }, props.tertiaryText)));
        };
        var photo = "/_layouts/15/userphoto.aspx?size=L&username=" + _this.props.userEmail;
        _this._Persona = {
            imageUrl: photo ? photo : '',
            imageInitials: _this._getInitial(_this.props.userName),
            text: _this.props.userName,
            secondaryText: _this.props.jobDescription,
            tertiaryText: _this.props.anniversary,
        };
        _this.state = {
            isAnniversaryToday: _this._AnniversaryIsToday(_this.props.anniversary)
        };
        _this._onRenderTertiaryText = _this._onRenderTertiaryText.bind(_this);
        _this._getInitial = _this._getInitial.bind(_this);
        _this._AnniversaryIsToday = _this._AnniversaryIsToday.bind(_this);
        return _this;
    }
    // Render
    HappyAnniversaryCard.prototype.render = function () {
        var _this = this;
        this._AnniversaryMsg = this.state.isAnniversaryToday ? strings.HappyAnniversaryMsg : strings.NextAnniversaryMsg;
        return (React.createElement("div", { className: styles.happyAnniversaryCard },
            React.createElement("div", { className: styles.documentCardWrapper },
                React.createElement("div", { className: styles.documentCard },
                    React.createElement(Image, { imageFit: ImageFit.cover, src: img, width: IMG_WIDTH, height: IMG_HEIGTH }),
                    React.createElement(Label, { className: styles.centered },
                        this.props.Ann_days,
                        " ",
                        this._AnniversaryMsg),
                    this.state.isAnniversaryToday ? React.createElement(Label, { className: styles.displayBirthdayToday },
                        " ",
                        this.props.anniversary) : React.createElement(Label, { className: styles.displayBirthday },
                        " ",
                        this.props.anniversary),
                    React.createElement("div", { className: styles.personaContainer },
                        React.createElement(Persona, __assign({}, this._Persona, { size: PersonaSize.regular, className: styles.persona, onRenderTertiaryText: this._onRenderTertiaryText }))),
                    React.createElement("div", { className: styles.actions },
                        React.createElement(DocumentCardActions, { actions: [
                                {
                                    iconProps: { iconName: 'Mail' },
                                    onClick: function (ev) {
                                        ev.preventDefault();
                                        ev.stopPropagation();
                                        window.location.href = "mailto:" + _this.props.userEmail + "?subject=" + _this._AnniversaryMsg + "!";
                                    },
                                    ariaLabel: 'email',
                                    title: 'Say Congrats'
                                }
                            ] }))))));
    };
    // Today is Anniversary ?
    HappyAnniversaryCard.prototype._AnniversaryIsToday = function (anniversary) {
        var _todayDay = moment().date();
        var _todayMonth = moment().month() + 1;
        var _AnniversaryDay = moment(anniversary, 'Do MMM').date();
        var _AnniversaryMonth = moment(anniversary, 'Do MMM').month() + 1;
        var _retvalue = (_todayDay === _AnniversaryDay && _todayMonth === _AnniversaryMonth) ? true : false;
        return _retvalue;
    };
    // Get Initials
    HappyAnniversaryCard.prototype._getInitial = function (userName) {
        var _arr = userName.split(' ');
        var _initial = _arr[0].charAt(0).toUpperCase() + (_arr[1] ? _arr[1].charAt(0).toLocaleUpperCase() : "");
        return _initial;
    };
    return HappyAnniversaryCard;
}(React.Component));
export { HappyAnniversaryCard };
export default HappyAnniversaryCard;
//# sourceMappingURL=HappyAnniversaryCard.js.map