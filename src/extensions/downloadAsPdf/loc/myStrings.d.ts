declare interface IDownloadAsPdfCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'DownloadAsPdfCommandSetStrings' {
  const strings: IDownloadAsPdfCommandSetStrings;
  export = strings;
}
