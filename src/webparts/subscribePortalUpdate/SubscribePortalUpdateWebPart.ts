import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import styles from './SubscribePortalUpdateWebPart.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISubscribePortalUpdateWebPartProps } from './ISubscribePortalUpdateWebPartProps';

export interface ISubscriptionItem {
  Id: number;
  SubscribedBy: string;
  SubscribedFor: string[];
  Subscribe: string;
  SubscribeOn: string;
  Title: string;
}

export default class SubscribePortalUpdateWebPart
  extends BaseClientSideWebPart<ISubscribePortalUpdateWebPartProps> {

  private _isSubscribed: boolean = false;
  private _isLoading: boolean = false;
  private _currentUserEmail: string = '';
  private _alreadyInitialized: boolean = false;
  private _productNameOptions: IPropertyPaneDropdownOption[] = [];
  private _productNameOptionsLoaded: boolean = false;

  // ----------------- Helpers / defaults -----------------

  private _ensureDefaultProperties(): void {
    if (!this.properties.ListName || this.properties.ListName.trim().length === 0) {
      this.properties.ListName = 'subscriptionlist';
    }

    if (!this.properties.PrefixText || this.properties.PrefixText.trim().length === 0) {
      this.properties.PrefixText = 'Subscribe to';
    }

    if (!this.properties.SuffixText || this.properties.SuffixText.trim().length === 0) {
      this.properties.SuffixText = 'updates';
    }

    if (!this.properties.ProductNameColumnName || this.properties.ProductNameColumnName.trim().length === 0) {
      this.properties.ProductNameColumnName = 'SubscribedFor';
    }

    if (!this.properties.HeaderFontSize || this.properties.HeaderFontSize <= 0) {
      this.properties.HeaderFontSize = 22;
    }

    if (!this.properties.DescriptionFontSize || this.properties.DescriptionFontSize <= 0) {
      this.properties.DescriptionFontSize = 14;
    }

    if (!this.properties.ButtonFontSize || this.properties.ButtonFontSize <= 0) {
      this.properties.ButtonFontSize = 15;
    }
  }

  private _logDebug(message: string, extra?: any): void {
    // eslint-disable-next-line no-console
    console.log(`[SubscribePortalUpdate] ${message}`, extra ?? '');
  }

  // ----------------- Lifecycle -----------------

  public async onInit(): Promise<void> {
    this._ensureDefaultProperties();
    this._logDebug('onInit', {
      listName: this.properties.ListName,
      productName: this.properties.ProductName
    });
    return super.onInit();
  }

  public render(): void {
    this._ensureDefaultProperties();

    this.domElement.innerHTML = this._getHtml();
    this._wireEvents();

    if (!this._alreadyInitialized) {
      this._alreadyInitialized = true;
      void this._loadInitialStatus();
    }
  }

  private async _loadInitialStatus(): Promise<void> {
    try {
      this._setLoading(true);
      await this._ensureUserEmail();
      await this._checkSubscription();
    } catch (err) {
      this._logDebug('Error checking subscription status', err);
    } finally {
      this._setLoading(false);
    }
  }

  // ----------------- User / context -----------------

  private async _ensureUserEmail(): Promise<void> {
    if (this._currentUserEmail) {
      return;
    }

    if (this.context.pageContext.user?.email) {
      this._currentUserEmail = this.context.pageContext.user.email;
      this._logDebug('Using pageContext user email', this._currentUserEmail);
      return;
    }

    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser`;
    this._logDebug('GET current user', url);

    const resp: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=none'
        }
      }
    );

    if (!resp.ok) {
      const text = await resp.text();
      this._logDebug('Error response from currentuser', { status: resp.status, body: text });
      throw new Error(`Unable to get current user. Status: ${resp.status}. Body: ${text}`);
    }

    const json: any = await resp.json();
    this._currentUserEmail = json.Email;
    this._logDebug('Resolved current user email', this._currentUserEmail);
  }

  // ----------------- Subscription status -----------------

  private async _checkSubscription(): Promise<void> {
    const item = await this._getSubscriptionItem();
    this._logDebug('Loaded subscription item', item);

    if (!item || !this.properties.ProductName) {
      this._isSubscribed = false;
    } else {
      const arr = item.SubscribedFor || [];
      this._isSubscribed = arr.indexOf(this.properties.ProductName) > -1;
    }
    this.render();
  }

  // ----------------- Subscribe / Unsubscribe -----------------

  private async _subscribe(): Promise<void> {
    this._setLoading(true);

    try {
      const product = this.properties.ProductName;
      if (!product) {
        throw new Error('ProductName is not configured.');
      }

      const item = await this._getSubscriptionItem();

      if (!item) {
        this._logDebug('No existing subscription item, creating new');
        await this._createSubscriptionItem(product);
      } else {
        const current = item.SubscribedFor || [];
        if (current.indexOf(product) === -1) {
          const updated = current.slice();
          updated.push(product);
          this._logDebug('Updating existing item with new product', { id: item.Id, updated });
          await this._updateSubscriptionItem(item.Id, updated);
        } else {
          this._isSubscribed = true;
          this.render();
          this._setLoading(false);
          return;
        }
      }

      this._isSubscribed = true;
      this.render();
    } catch (err) {
      this._logDebug('Error subscribing', err);
    }

    this._setLoading(false);
  }

  private async _unsubscribe(): Promise<void> {
    this._setLoading(true);

    try {
      const product = this.properties.ProductName;
      if (!product) {
        throw new Error('ProductName is not configured.');
      }

      const item = await this._getSubscriptionItem();
      if (!item) {
        this._isSubscribed = false;
        this.render();
        this._setLoading(false);
        return;
      }

      const current = item.SubscribedFor || [];
      const remaining = current.filter(x => x !== product);

      this._logDebug('Updating item for unsubscribe', {
        id: item.Id,
        current,
        remaining
      });

      await this._updateSubscriptionItem(item.Id, remaining);

      this._isSubscribed = false;
      this.render();
    } catch (err) {
      this._logDebug('Error unsubscribing', err);
    }

    this._setLoading(false);
  }

  // ----------------- REST helpers -----------------

  private async _getSubscriptionItem(): Promise<ISubscriptionItem | null> {
    const list = this.properties.ListName;
    if (!list) {
      throw new Error('ListName is not configured.');
    }

    const email = this._currentUserEmail;
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${list}')/items` +
      `?$filter=SubscribedBy eq '${email}'` +
      `&$select=Id,Title,SubscribedBy,SubscribedFor,Subscribe,SubscribeOn`;

    this._logDebug('GET subscription item', url);

    const resp: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=none'
        }
      }
    );

    if (!resp.ok) {
      const text = await resp.text();
      this._logDebug('Error from GET subscription item', { status: resp.status, body: text });

      if (resp.status === 404) {
        throw new Error('List not found. Body: ' + text);
      }
      throw new Error(`Failed to get subscription item. Status: ${resp.status}. Body: ${text}`);
    }

    const json: any = await resp.json();
    const row = (json.value && json.value[0]) || null;

    if (!row) {
      return null;
    }

    const raw = row.SubscribedFor;
    let subscribedFor: string[] = [];

    if (Array.isArray(raw)) {
      subscribedFor = raw;
    } else if (raw && Array.isArray(raw.results)) {
      subscribedFor = raw.results;
    }

    const result: ISubscriptionItem = {
      Id: row.Id,
      Title: row.Title,
      SubscribedBy: row.SubscribedBy,
      Subscribe: row.Subscribe,
      SubscribeOn: row.SubscribeOn,
      SubscribedFor: subscribedFor
    };

    this._logDebug('Parsed subscription item', result);
    return result;
  }

  private async _createSubscriptionItem(product: string): Promise<void> {
    const list = this.properties.ListName;
    if (!list) {
      throw new Error('ListName is not configured.');
    }

    // Multi-choice with odata.metadata=none → plain array
    const body: any = {
      Title: this._currentUserEmail,
      SubscribedBy: this._currentUserEmail,
      Subscribe: 'Yes',
      SubscribeOn: new Date().toISOString(),
      SubscribedFor: [product]
    };

    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${list}')/items`;

    this._logDebug('POST create subscription item', { url, body });

    const resp: SPHttpClientResponse = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json;odata.metadata=none'
        },
        body: JSON.stringify(body)
      }
    );

    if (!resp.ok) {
      const text = await resp.text();
      this._logDebug('Error from POST create item', { status: resp.status, body: text });
      throw new Error(`Failed to create subscription item. Status: ${resp.status}. Body: ${text}`);
    }

    const json = await resp.json();
    this._logDebug('Created subscription item', json);
  }

  private async _updateSubscriptionItem(id: number, products: string[]): Promise<void> {
    const list = this.properties.ListName;
    if (!list) {
      throw new Error('ListName is not configured.');
    }

    // Multi-choice with odata.metadata=none → plain array
    const body: any = {
      SubscribeOn: new Date().toISOString(),
      Subscribe: products.length > 0 ? 'Yes' : 'No',
      SubscribedFor: products
    };

    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${list}')/items(${id})`;

    this._logDebug('POST update subscription item', { url, body });

    const resp: SPHttpClientResponse = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json;odata.metadata=none',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(body)
      }
    );

    if (!resp.ok) {
      const text = await resp.text();
      this._logDebug('Error from POST update item', { status: resp.status, body: text });
      throw new Error(`Failed to update subscription item. Status: ${resp.status}. Body: ${text}`);
    }

    this._logDebug('Update subscription item OK');
  }

  // ----------------- Load Product Name Choices -----------------

  private async _loadProductNameChoices(): Promise<IPropertyPaneDropdownOption[]> {
    const list = this.properties.ListName;
    const columnName = this.properties.ProductNameColumnName;

    if (!list || !columnName) {
      this._logDebug('List name or column name not configured', { list, columnName });
      return [];
    }

    try {
      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${list}')/fields` +
        `?$filter=EntityPropertyName eq '${columnName}'` +
        `&$select=Choices,TypeAsString`;

      this._logDebug('GET field choices', url);

      const resp: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata.metadata=minimal'
          }
        }
      );

      if (!resp.ok) {
        const text = await resp.text();
        this._logDebug('Error from GET field choices', { status: resp.status, body: text });
        throw new Error(`Failed to get field choices. Status: ${resp.status}`);
      }

      const json: any = await resp.json();
      
      if (!json.value || json.value.length === 0) {
        this._logDebug('Field not found', columnName);
        return [];
      }

      const field = json.value[0];
      const choices = field.Choices || [];

      this._logDebug('Loaded choices', choices);

      const options: IPropertyPaneDropdownOption[] = choices.map((choice: string) => ({
        key: choice,
        text: choice
      }));

      return options;
    } catch (err) {
      this._logDebug('Error loading product name choices', err);
      return [];
    }
  }

  // ----------------- UI helpers -----------------

  private _setLoading(flag: boolean): void {
    this._isLoading = flag;
    this.render();
  }

  private _getHtml(): string {
    const product = this.properties.ProductName || '';
    const prefix = this.properties.PrefixText || 'Subscribe to';
    const suffix = this.properties.SuffixText || 'updates';

    const buttonText = this._isSubscribed
      ? `Unsubscribe from ${product}`
      : `${prefix} ${product} ${suffix}`.replace(/\s+/g, ' ').trim();

    const btnClass = this._isSubscribed ? styles.buttonSecondary : styles.buttonPrimary;

    return `
      <div class="${styles.subscribePortalUpdate}">
        <div class="${styles.card}">
          <div class="${styles.cardHeader}">
            <div class="${styles.headerContent}">
              <h3 class="${styles.title}" style="font-size: ${this.properties.HeaderFontSize}px;">
                ${escape(product || 'Product Updates')}
              </h3>
              ${this.properties.Description ? `
                <p class="${styles.description}" style="font-size: ${this.properties.DescriptionFontSize}px;">
                  ${escape(this.properties.Description)}
                </p>
              ` : ''}
            </div>
          </div>

          <div class="${styles.cardBody}">
            <button 
              id="toggleBtn"
              type="button"
              class="${styles.button} ${btnClass}"
              ${this._isLoading ? 'disabled' : ''}
              aria-busy="${this._isLoading}"
              style="font-size: ${this.properties.ButtonFontSize}px;"
            >
              ${this._isLoading ? `
                <span class="${styles.spinner}"></span>
                <span>Processing...</span>
              ` : `
                <span class="${styles.buttonText}">${escape(buttonText)}</span>
              `}
            </button>
          </div>
        </div>
      </div>
    `;
  }

  private _wireEvents(): void {
    const btn = this.domElement.querySelector('#toggleBtn') as HTMLButtonElement | null;
    if (btn) {
      btn.addEventListener('click', () => {
        if (this._isLoading) {
          return;
        }
        if (this._isSubscribed) {
          void this._unsubscribe();
        } else {
          void this._subscribe();
        }
      });
    }
  }

  // ----------------- SPFx plumbing -----------------

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  // Support section backgrounds
  protected get supportsThemeVariants(): boolean {
    return true;
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if (!this._productNameOptionsLoaded) {
      this.context.propertyPane.refresh();
      this._productNameOptions = await this._loadProductNameChoices();
      this._productNameOptionsLoaded = true;
      this.context.propertyPane.refresh();
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // If ListName or ProductNameColumnName changes, reload the choices
    if (propertyPath === 'ListName' || propertyPath === 'ProductNameColumnName') {
      this._productNameOptionsLoaded = false;
      this._productNameOptions = [];
      
      // Reload choices
      void this._loadProductNameChoices().then((options) => {
        this._productNameOptions = options;
        this._productNameOptionsLoaded = true;
        this.context.propertyPane.refresh();
      });
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Subscribe Portal Update Settings' },
          groups: [
            {
              groupName: 'List Configuration',
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: 'List Name',
                  description: 'Default: "subscriptionlist"'
                }),
                PropertyPaneTextField('ProductNameColumnName', {
                  label: 'Product Name Column',
                  description: 'The internal name of the multi-choice column (Default: "SubscribedFor")'
                })
              ]
            },
            {
              groupName: 'Product Selection',
              groupFields: [
                PropertyPaneDropdown('ProductName', {
                  label: 'Product Name',
                  options: this._productNameOptions,
                  disabled: !this._productNameOptionsLoaded || this._productNameOptions.length === 0
                })
              ]
            },
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('Description', {
                  label: 'Description',
                  multiline: true
                }),
                PropertyPaneTextField('PrefixText', {
                  label: 'Button Prefix Text',
                  description: 'Default: "Subscribe to"'
                }),
                PropertyPaneTextField('SuffixText', {
                  label: 'Button Suffix Text',
                  description: 'Default: "updates"'
                })
              ]
            },
            {
              groupName: 'Font Sizes',
              groupFields: [
                PropertyPaneSlider('HeaderFontSize', {
                  label: 'Header Font Size (px)',
                  min: 12,
                  max: 48,
                  step: 1,
                  value: this.properties.HeaderFontSize || 22,
                  showValue: true
                }),
                PropertyPaneSlider('DescriptionFontSize', {
                  label: 'Description Font Size (px)',
                  min: 10,
                  max: 24,
                  step: 1,
                  value: this.properties.DescriptionFontSize || 14,
                  showValue: true
                }),
                PropertyPaneSlider('ButtonFontSize', {
                  label: 'Button Font Size (px)',
                  min: 12,
                  max: 24,
                  step: 1,
                  value: this.properties.ButtonFontSize || 15,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
//vishnu