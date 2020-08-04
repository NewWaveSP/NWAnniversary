/// <reference types="react" />
import * as React from 'react';
import { IHappyAnniversaryCardProps } from './IHappyAnniversaryCardProps';
import { IHappyAnniversaryCardPState } from './IHappyAnniversaryCardState';
export declare class HappyAnniversaryCard extends React.Component<IHappyAnniversaryCardProps, IHappyAnniversaryCardPState> {
    private _Persona;
    private _AnniversaryMsg;
    constructor(props: IHappyAnniversaryCardProps);
    render(): React.ReactElement<IHappyAnniversaryCardProps>;
    private _AnniversaryIsToday(anniversary);
    private _getInitial(userName);
    private _onRenderTertiaryText;
}
export default HappyAnniversaryCard;
