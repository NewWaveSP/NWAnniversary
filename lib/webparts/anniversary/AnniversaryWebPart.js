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
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'AnniversaryWebPartStrings';
import Anniversary from './components/Anniversary';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
var AnniversaryWebPart = (function (_super) {
    __extends(AnniversaryWebPart, _super);
    function AnniversaryWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    AnniversaryWebPart.prototype.onInit = function () {
        return _super.prototype.onInit.call(this).then(function (_) {
            // other init code may be present
        });
    };
    AnniversaryWebPart.prototype.render = function () {
        var _this = this;
        var element = React.createElement(Anniversary, {
            title: this.properties.title,
            NumberAnniversary: this.properties.NumberAnniversary,
            context: this.context,
            displayMode: this.displayMode,
            updateProperty: function (value) {
                _this.properties.title = value;
            }
        });
        ReactDom.render(element, this.domElement);
    };
    AnniversaryWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(AnniversaryWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    AnniversaryWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('title', {
                                    label: strings.DescriptionFieldLabel
                                }),
                                PropertyFieldNumber("NumberAnniversary", {
                                    key: "NumberAnniversary",
                                    label: strings.NumberUpComingDaysLabel,
                                    description: strings.NumberUpComingDaysLabel,
                                    value: this.properties.NumberAnniversary,
                                    maxValue: 100,
                                    disabled: false
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return AnniversaryWebPart;
}(BaseClientSideWebPart));
export default AnniversaryWebPart;
//# sourceMappingURL=AnniversaryWebPart.js.map