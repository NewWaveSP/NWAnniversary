import * as React from 'react';
import styles from './Anniversary.module.scss';
import { IAnniversaryProps } from './IAnniversaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HappyAnniversary, IUser } from '../../../controls/happyanniversary';
import * as moment from 'moment';
import { IAnniversaryState } from './IAnniversarysState';
import SPService from '../../../services/SPService';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
const imgBackgroundBallons: string = require('../../../../assets/ballonsBackgroud.png');
import { Image, IImageProps, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as strings from 'AnniversaryWebPartStrings';

let listItems = [];

export default class Anniversarys extends React.Component<IAnniversaryProps, IAnniversaryState> {
  private _users: IUser[] = [];
  private _tempusers: IUser[] = [];
  private _spServices: SPService;
  constructor(props: IAnniversaryProps) {
    super(props);
    this._spServices = new SPService(this.props.context);
    this.state = {
      Users: [],
      showAnniversarys: true
    };
  }

  public componentDidMount(): void {
    this.GetUsers();
  }

  // Render
  public render(): React.ReactElement<IAnniversaryProps> {
    let _center: any = !this.state.showAnniversarys ? "center" : "";
   {
      return (
        <div className={styles.anniversary}
          style={{ textAlign: _center }} >
          <div className={styles.container}>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.title}
              updateProperty={this.props.updateProperty} />
            {
              !this.state.showAnniversarys ?
                <div className={styles.backgroundImgBallons}>
                  <Image imageFit={ImageFit.cover}
                    src={imgBackgroundBallons}
                    width={150}
                    height={150}
                  />
                  <Label className={styles.subTitle}>{strings.MessageNoAnniversarys}</Label>
                </div>
                :
                <HappyAnniversary users={this.state.Users}
                />
            }
          </div>
        </div>
      );
    }

  }

  // Sort Array of Anniversarys
  private SortAnniversarys(users: IUser[]) {
    return users.sort((a, b) => {
      if (a.Fakeanniversary > b.Fakeanniversary) {
        return 1;
      }
      if (a.Fakeanniversary < b.Fakeanniversary) {
        return -1;
      }
      return 0;
    });
  }
  private ordinal_suffix_of(Ann_number) {
    var j = Ann_number % 10,
      k = Ann_number % 100;
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
  }
  // Load List Of Users
  private async GetUsers() {
    let _otherMonthsAnniversarys: IUser[], _dezemberAnniversarys: IUser[];
    let CurrentYear: any;
    listItems = await this._spServices.getPAnniversarys(this.props.NumberAnniversary);
    var AllowedItem = this.props.NumberAnniversary;
    if (listItems && listItems.length > 0) {
      _otherMonthsAnniversarys = [];
      _dezemberAnniversarys = [];
      for (const item of listItems) {
        if (item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber) {
          var AnniversaryDate = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber);
          var TodayDate = moment();
          var YearDiff = TodayDate.diff(AnniversaryDate, 'year');

          YearDiff = YearDiff + 1;
          var Yeardiff = this.ordinal_suffix_of(YearDiff);
          CurrentYear = moment().format('YYYY');
          var FakeAnniversary = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber, 'MM/DD/YYYY').year(CurrentYear).format();
          var Fake_Anniversary = moment(item.extension_bd2923d14f8d4d96a1f0280682a13e4c_employeeNumber, 'MM/DD/YYYY').year(CurrentYear);
           var DateDiff = TodayDate.diff(Fake_Anniversary, 'days');
          var Isexpired = Fake_Anniversary.diff(TodayDate, 'days');
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
      this._users = this._tempusers.filter((itm, i) => {
        if (AllowedItem > i)
          return itm;
      }); this.setState(
        {
          Users: this._users,
          showAnniversarys: this._users.length === 0 ? false : true
        });
    }
  }
}