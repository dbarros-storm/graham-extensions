declare interface IShareMultiDocumentsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ShareMultiDocumentsCommandSetStrings' {
  const strings: IShareMultiDocumentsCommandSetStrings;
  export = strings;
}
