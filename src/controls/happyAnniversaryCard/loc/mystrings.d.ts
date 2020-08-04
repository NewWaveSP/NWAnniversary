declare interface IControlStrings {
  AnniversaryControlDefaultDay: string,
  HappyAnniversaryMsg: string,
  NextAnniversaryMsg: string,
  MessageNoAnniversarys: string
}

declare module 'ControlStrings' {
  const strings: IControlStrings;
  export = strings;
}
