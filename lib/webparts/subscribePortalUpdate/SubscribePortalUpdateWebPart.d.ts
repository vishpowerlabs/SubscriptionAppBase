import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import { ISubscribePortalUpdateWebPartProps } from './ISubscribePortalUpdateWebPartProps';
export interface ISubscriptionItem {
    Id: number;
    SubscribedBy: string;
    SubscribedFor: string[];
    Subscribe: string;
    SubscribeOn: string;
    Title: string;
}
export default class SubscribePortalUpdateWebPart extends BaseClientSideWebPart<ISubscribePortalUpdateWebPartProps> {
    private _isSubscribed;
    private _isLoading;
    private _currentUserEmail;
    private _alreadyInitialized;
    private _productNameOptions;
    private _productNameOptionsLoaded;
    private _ensureDefaultProperties;
    private _logDebug;
    onInit(): Promise<void>;
    render(): void;
    private _loadInitialStatus;
    private _ensureUserEmail;
    private _checkSubscription;
    private _subscribe;
    private _unsubscribe;
    private _getSubscriptionItem;
    private _createSubscriptionItem;
    private _updateSubscriptionItem;
    private _loadProductNameChoices;
    private _setLoading;
    private _getHtml;
    private _wireEvents;
    protected get dataVersion(): Version;
    protected get disableReactivePropertyChanges(): boolean;
    protected get supportsThemeVariants(): boolean;
    protected onPropertyPaneConfigurationStart(): Promise<void>;
    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SubscribePortalUpdateWebPart.d.ts.map