declare interface IReviseDocumentCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ReviseDocumentCommandSetStrings' {
  const strings: IReviseDocumentCommandSetStrings;
  export = strings;
}
