import * as React from 'react';
import styles from './HappyAnniversaryCard.module.scss';
import { IHappyAnniversaryCardProps } from './IHappyAnniversaryCardProps';
import { IHappyAnniversaryCardPState } from './IHappyAnniversaryCardState';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPersonaSharedProps, Persona, PersonaSize, IPersonaProps, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Image, IImageProps, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'AnniversaryWebPartStrings';
import * as moment from 'moment';
import {
  DocumentCardActions,
} from 'office-ui-fabric-react/lib/DocumentCard';
const img: string = require('../../../assets/baloons.png');
const IMG_WIDTH: number = 200;
const IMG_HEIGTH: number = 190;

export class HappyAnniversaryCard extends React.Component<IHappyAnniversaryCardProps, IHappyAnniversaryCardPState> {
  private _Persona: IPersonaSharedProps;
  private _AnniversaryMsg: string = '';

  constructor(props: IHappyAnniversaryCardProps) {
    super(props);
    const photo: string = `/_layouts/15/userphoto.aspx?size=L&username=${this.props.userEmail}`;

    this._Persona = {
      imageUrl: photo ? photo : '',
      imageInitials: this._getInitial(this.props.userName),
      text: this.props.userName,
      secondaryText: this.props.jobDescription,
      tertiaryText: this.props.anniversary,
    };

    this.state = {
      isAnniversaryToday: this._AnniversaryIsToday(this.props.anniversary)
    };

    this._onRenderTertiaryText = this._onRenderTertiaryText.bind(this);
    this._getInitial = this._getInitial.bind(this);
    this._AnniversaryIsToday = this._AnniversaryIsToday.bind(this);
  }
  // Render
  public render(): React.ReactElement<IHappyAnniversaryCardProps> {
    this._AnniversaryMsg = this.state.isAnniversaryToday ? strings.HappyAnniversaryMsg : strings.NextAnniversaryMsg;

    return (
      <div className={styles.happyAnniversaryCard}>
        <div className={styles.documentCardWrapper}>
          <div className={styles.documentCard}>
            <Image
              imageFit={ImageFit.cover}
              src={img}
              width={IMG_WIDTH}
              height={IMG_HEIGTH}
            />
            <Label className={styles.centered}>{this.props.Ann_days} {this._AnniversaryMsg}</Label>
            { this.state.isAnniversaryToday ? <Label className={styles.displayBirthdayToday}> {this.props.anniversary}</Label> : <Label className={styles.displayBirthday}> {this.props.anniversary}</Label> }

            <div className={styles.personaContainer}>
              <Persona
                {...this._Persona}
                size={PersonaSize.regular}
                className={styles.persona}
                onRenderTertiaryText={this._onRenderTertiaryText}
              />
            </div>
            <div className={styles.actions}>
              <DocumentCardActions
                actions={[
                  {
                    iconProps: { iconName: 'Mail' },
                    onClick: (ev: any) => {
                      ev.preventDefault();
                      ev.stopPropagation();
                      window.location.href = `mailto:${this.props.userEmail}?subject=${this._AnniversaryMsg}!`;
                    },
                    ariaLabel: 'email',
                    title: 'Say Congrats'
                  }
                ]}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }


  // Today is Anniversary ?
  private _AnniversaryIsToday(anniversary: string): boolean {
    const _todayDay = moment().date();
    const _todayMonth = moment().month() + 1;
    const _AnniversaryDay = moment(anniversary, 'Do MMM').date();
    const _AnniversaryMonth = moment(anniversary, 'Do MMM').month() + 1;

    const _retvalue = (_todayDay === _AnniversaryDay && _todayMonth === _AnniversaryMonth) ? true : false;

    return _retvalue;
  }
  // Get Initials
  private _getInitial(userName: string): string {
    const _arr = userName.split(' ');
    const _initial = _arr[0].charAt(0).toUpperCase() + (_arr[1] ? _arr[1].charAt(0).toLocaleUpperCase() : "");
    return _initial;
  }
  // Render tertiary text
  private _onRenderTertiaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <span className='ms-fontWeight-semibold' style={{ color: '#71afe5' }}>
          {props.tertiaryText}</span>
      </div>
    );
  }
}
export default HappyAnniversaryCard;
