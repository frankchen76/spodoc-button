declare interface ICallAzureFunctionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CallAzureFunctionCommandSetStrings' {
  const strings: ICallAzureFunctionCommandSetStrings;
  export = strings;
}
