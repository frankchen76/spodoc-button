import { autobind } from "@uifabric/utilities";

export class SettingValue {
    public title: string;
    public iconImage: string;
    public apiUrl: string;
    public displayLists: string[];
    public refreshPage: boolean;

    @autobind
    public IsShownForList(listTitle: string): boolean {
        return this.displayLists.indexOf(listTitle) != -1;
    }
}
