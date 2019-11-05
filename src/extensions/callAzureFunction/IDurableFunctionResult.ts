export interface IDurableFunctionResult {
    id: number;
    statusQueryGetUri: string;
    sendEventPostUri: string;
    terminatePostUri: string;
    rewindPostUri: string;
    purgeHistoryDeleteUri: string;
}
