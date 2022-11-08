import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { IJsDisplayListWebPartProps } from './IJsDisplayListWebPartProps';
import { Version } from '@microsoft/sp-core-library';
export default class JsDisplayListWebPart extends BaseClientSideWebPart<IJsDisplayListWebPartProps> {
    render(): void;
    protected get disableReactivePropertyChanges(): boolean;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    protected onPropertyPaneConfigurationStart(): void;
    private _dropdownOptions;
    private _getListTitles;
    private _getListData;
    private _getListDocuments;
    private _getListComments;
    private _renderList;
    private _renderListDocuments;
    private _renderListComments;
    private _createComment;
    private _renderWebPart;
    private _renderListAsync;
    private _renderListTitles;
}
//# sourceMappingURL=JsDisplayListWebPart.d.ts.map