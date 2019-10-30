import { autobind } from "@uifabric/utilities";

export class SettingValue {
    public title: string;
    public iconImage: string;
    public apiUrl: string;
    public displayLists: string[];

    @autobind
    public IsShownForList(listTitle: string): boolean {
        return this.displayLists.indexOf(listTitle) != -1;
    }
}
