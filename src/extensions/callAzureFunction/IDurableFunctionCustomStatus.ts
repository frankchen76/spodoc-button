export interface IDurableFunctionCustomStatus {
    name: string;
    instanceId: string;
    runtimeStatus: string;
    input: any;
    customStatus: {
        Process: number;
        Message: string;
    };
    output: any;
    createTime: string;
    lastUpdateTime: string;
}
