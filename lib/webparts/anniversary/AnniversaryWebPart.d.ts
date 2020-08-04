import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IAnniversaryWebPartProps {
    title: string;
    NumberAnniversary: number;
}
export default class AnniversaryWebPart extends BaseClientSideWebPart<IAnniversaryWebPartProps> {
    private graphCLient;
    onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
