import * as React from 'react';
import styles from './HappyAnniversary.module.scss';
import { IHappyAnniversaryProps } from './IHappyAnniversaryProps';
import { IHappyAnniversarystate } from './IHappyAnniversarystate';
import { escape } from '@microsoft/sp-lodash-subset';
import { IUser } from './IUser';
import HappyBirdthayCard from '../happyAnniversaryCard/HappyAnniversaryCard';
import * as moment from 'moment';

export class HappyAnniversary extends React.Component<IHappyAnniversaryProps, IHappyAnniversarystate> {

  // private _showAnniversarys: boolean = true;
  constructor(props: IHappyAnniversaryProps) {
    super(props);
  }

  //
  public render(): React.ReactElement<IHappyAnniversaryProps> {

    return (
      <div className={styles.happyAnniversary}>
        {
          this.props.users.map((user: IUser) => {

            return (
              <div className={styles.container}>
                <HappyBirdthayCard userName={user.userName}
                  jobDescription={user.jobDescription}
                  anniversary={moment(user.anniversary, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MMM DD')}
                  userEmail={user.userEmail}
                  Ann_days={user.Ann_days}
                />
              </div>
            );
          })
        }
      </div>
    );
  }
}
export default HappyAnniversary;
