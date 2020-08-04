declare interface IAnniversaryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  NumberUpComingDaysLabel: string;
  AnniversaryControlDefaultDay: string;
  HappyAnniversaryMsg: string;
  NextAnniversaryMsg: string;
  MessageNoAnniversarys: strings;
}

declare module 'AnniversaryWebPartStrings' {
  const strings: IAnniversaryWebPartStrings;
  export = strings;
}

