/// <reference types="react" />
import * as React from 'react';
import { IAnniversaryProps } from './IAnniversaryProps';
import { IAnniversaryState } from './IAnniversarysState';
export default class Anniversarys extends React.Component<IAnniversaryProps, IAnniversaryState> {
    private _users;
    private _tempusers;
    private _spServices;
    constructor(props: IAnniversaryProps);
    componentDidMount(): void;
    render(): React.ReactElement<IAnniversaryProps>;
    private SortAnniversarys(users);
    private ordinal_suffix_of(Ann_number);
    private GetUsers();
}
